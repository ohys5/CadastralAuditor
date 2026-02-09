import os
from PyQt5.QtWidgets import QAction, QMessageBox
from qgis.core import (
    QgsProject, QgsVectorLayer, QgsField, 
    QgsEditorWidgetSetup, QgsDefaultValue,
    QgsVectorFileWriter, QgsCoordinateReferenceSystem,
    QgsCoordinateTransformContext
)
from PyQt5.QtCore import QVariant
from .dialog import QFieldAutoSetupDialog

class QFieldAutoSetupPlugin:
    def __init__(self, iface):
        self.iface = iface
        self.action = None

    def initGui(self):
        self.action = QAction("QField 자동 설정", self.iface.mainWindow())
        self.action.triggered.connect(self.run)
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToMenu("&QField Auto Setup", self.action)

    def unload(self):
        self.iface.removeToolBarIcon(self.action)
        self.iface.removePluginMenu("&QField Auto Setup", self.action)

    def run(self):
        dlg = QFieldAutoSetupDialog()
        if dlg.exec_():
            data = dlg.get_data()
            self.setup_project(data)

    def setup_project(self, data):
        try:
            project_name = data['name']
            base_path = data['path']
            template_path = data['template']
            crs_authid = data['crs']
            use_photo = data['photo']

            # 1. 템플릿 로드 로직 (기존 프로젝트 위에 얹기)
            if template_path and os.path.exists(template_path):
                # 템플릿이 프로젝트 파일(.qgs/.qgz)인 경우
                if template_path.endswith('.qgs') or template_path.endswith('.qgz'):
                    QgsProject.instance().read(template_path)
                # 템플릿이 일반 벡터 레이어(.shp 등)인 경우
                else:
                    QgsProject.instance().clear()
                    base_layer = QgsVectorLayer(template_path, "기초_데이터", "ogr")
                    if base_layer.isValid():
                        QgsProject.instance().addMapLayer(base_layer)
            else:
                # 템플릿 없으면 초기화
                QgsProject.instance().clear()

            # 2. 좌표계 설정 (템플릿이 있어도 입력된 좌표계로 강제 통일 권장)
            crs = QgsCoordinateReferenceSystem(crs_authid)
            QgsProject.instance().setCrs(crs)

            # 3. 새 프로젝트 폴더 생성
            project_dir = os.path.join(base_path, project_name)
            if not os.path.exists(project_dir):
                os.makedirs(project_dir)
            
            # 4. 현장 조사용 GeoPackage 레이어 생성
            gpkg_path = os.path.join(project_dir, "survey_data.gpkg")
            layer_name = "현장조사_포인트"
            
            # GeoPackage 생성 옵션
            options = QgsVectorFileWriter.SaveVectorOptions()
            options.driverName = "GPKG"
            options.layerName = layer_name
            
            # 임시 레이어 생성 후 필드 정의
            temp_layer = QgsVectorLayer(f"Point?crs={crs_authid}", layer_name, "memory")
            pr = temp_layer.dataProvider()
            
            # 표준 필드 추가
            fields = [
                QgsField("survey_id", QVariant.Int),
                QgsField("survey_date", QVariant.DateTime),
                QgsField("memo", QVariant.String),
                QgsField("category", QVariant.String) # 조사 유형(예: 양호/불량)
            ]
            if use_photo:
                fields.append(QgsField("photo_path", QVariant.String))
                
            pr.addAttributes(fields)
            temp_layer.updateFields()
            
            # GeoPackage로 저장
            write_result, error_msg = QgsVectorFileWriter.writeAsVectorFormat(
                temp_layer, gpkg_path, options
            )

            if write_result == QgsVectorFileWriter.NoError:
                # 5. 생성된 레이어를 프로젝트에 로드
                final_layer = QgsVectorLayer(f"{gpkg_path}|layername={layer_name}", layer_name, "ogr")
                QgsProject.instance().addMapLayer(final_layer)
                
                # --- [중요] 레이어 순서 조정 ---
                # 조사 포인트가 배경지도(템플릿) 위에 보이도록 최상단으로 이동
                root = QgsProject.instance().layerTreeRoot()
                my_node = root.findLayer(final_layer.id())
                clone = my_node.clone()
                root.insertChildNode(0, clone) # 0번 인덱스 = 최상단
                root.removeChildNode(my_node)
                
                # --- [핵심] 위젯 자동 설정 ---
                
                # A. 날짜 필드: 현재 시간 자동 입력 (now())
                idx_date = final_layer.fields().indexOf("survey_date")
                if idx_date != -1:
                    final_layer.setDefaultValueDefinition(idx_date, QgsDefaultValue("now()"))
                
                # B. 사진 필드: 첨부(Attachment) 위젯 + 상대 경로 설정
                if use_photo:
                    idx_photo = final_layer.fields().indexOf("photo_path")
                    if idx_photo != -1:
                        # QField에서 사진을 찍으려면 'Attachment' 위젯이 필수
                        config = {
                            'DocumentViewer': 1,  # 1: 이미지 뷰어
                            'FileStorage': 0,     # 0: 일반 파일 저장
                            'RelativeStorage': 1, # 1: 프로젝트 기준 상대 경로 (중요!)
                            'StorageMode': 0      # 0: 파일 경로 저장
                        }
                        # 위젯 설정 적용
                        widget_setup = QgsEditorWidgetSetup('Attachment', config)
                        final_layer.setEditorWidgetSetup(idx_photo, widget_setup)
                
                # C. 카테고리 필드: 값 맵(Value Map) 예시 추가
                idx_cat = final_layer.fields().indexOf("category")
                if idx_cat != -1:
                    kv_config = {
                        'map': [
                            {'조사 필요': 'check'},
                            {'양호': 'good'},
                            {'불량': 'bad'}
                        ]
                    }
                    final_layer.setEditorWidgetSetup(idx_cat, QgsEditorWidgetSetup('ValueMap', kv_config))

                # 6. 최종 프로젝트 저장 (.qgs)
                # 템플릿 원본은 건드리지 않고, 새 경로에 다른 이름으로 저장
                new_project_path = os.path.join(project_dir, f"{project_name}.qgs")
                QgsProject.instance().write(new_project_path)
                
                QMessageBox.information(None, "완료", f"프로젝트가 생성되었습니다!\n\n저장 위치: {new_project_path}\n\n이제 QFieldSync를 통해 패키징하세요.")
                
            else:
                QMessageBox.critical(None, "오류", f"레이어 생성 실패: {error_msg}")

        except Exception as e:
            QMessageBox.critical(None, "오류", f"작업 중 문제가 발생했습니다:\n{str(e)}")
            import traceback
            traceback.print_exc()