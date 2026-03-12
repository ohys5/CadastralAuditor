import os
import csv
import math
from PyQt5.QtWidgets import QAction, QMessageBox, QFileDialog, QTableWidgetItem, QApplication, QInputDialog, QDialog, QVBoxLayout, QLabel, QListWidget, QDialogButtonBox, QListWidgetItem, QPushButton, QHBoxLayout, QCheckBox, QAbstractItemView
from PyQt5.QtCore import Qt, QTimer, QSizeF
from PyQt5.QtGui import QColor, QTextDocument
from qgis.core import (
    QgsProject, QgsMapLayer, QgsWkbTypes, QgsSpatialIndex, 
    QgsFeatureRequest, QgsGeometry, Qgis, QgsPointXY, QgsCoordinateTransform, QgsRectangle, QgsTextAnnotation
)
from qgis.gui import QgsRubberBand, QgsVertexMarker, QgsMapToolEmitPoint
from .cadastral_auditor_dialog import CadastralAuditorDialog
from .analyzer import GeometricMatcher, BestFitSolver, SegmentAuditor, SmartGeometryComparator, PointToLineAuditor, ConfidenceStringMatcher
from qgis.core import QgsFeature

def get_angle_diff(geom1, geom2):
    """
    두 형상(Line) 간의 각도 차이를 반환 (0~90도 범위로 정규화)
    """
    # 1. 라인 타입 체크
    if geom1.type() != QgsWkbTypes.LineGeometry or geom2.type() != QgsWkbTypes.LineGeometry:
        return 0.0
        
    # 2. 형상에서 첫점과 끝점 추출 (단순화된 각도 계산)
    line1 = geom1.asPolyline() if not geom1.isMultipart() else geom1.asMultiPolyline()[0]
    line2 = geom2.asPolyline() if not geom2.isMultipart() else geom2.asMultiPolyline()[0]
    
    if len(line1) < 2 or len(line2) < 2: return 0.0
    
    # 3. 각도(Azimuth) 계산
    dx1, dy1 = line1[-1].x() - line1[0].x(), line1[-1].y() - line1[0].y()
    dx2, dy2 = line2[-1].x() - line2[0].x(), line2[-1].y() - line2[0].y()
    
    angle1 = math.degrees(math.atan2(dy1, dx1))
    angle2 = math.degrees(math.atan2(dy2, dx2))
    
    # 4. 차이값 정규화 (평행=0, 수직=90)
    diff = abs(angle1 - angle2)
    while diff > 90:
        diff = abs(diff - 180)
        
    return diff

def get_advanced_similarity(geom_curr, cad_geom):
    # 1. 평균 거리 계산 (정점들을 샘플링하여 거리 평균 산출)
    dist_sum = 0
    length = geom_curr.length()
    if length == 0: return float('inf'), 0.0
    
    points = [geom_curr.interpolate(length * i/4) for i in range(5)]
    for p in points:
        dist_sum += p.distance(cad_geom)
    avg_dist = dist_sum / 5
    
    # 2. 선분 길이 유사도 (너무 길거나 짧은 선 제외)
    l_curr = geom_curr.length()
    l_cad = cad_geom.length()
    len_ratio = min(l_curr, l_cad) / max(l_curr, l_cad) if max(l_curr, l_cad) > 0 else 0.0
    
    return avg_dist, len_ratio

def check_projection_overlap(geom_curr, cad_geom, width=2.0):
    """
    현황선이 지적선 버퍼 내에 얼마나 포함되는지 비율 계산 (0.0 ~ 1.0)
    Proximity Trap(거리는 가깝지만 서로 엇갈린 경우) 방지용
    """
    if cad_geom.isEmpty() or geom_curr.isEmpty():
        return 0.0
        
    buff = cad_geom.buffer(width, 4)
    intersection = geom_curr.intersection(buff)
    return intersection.length() / geom_curr.length() if geom_curr.length() > 0 else 0.0

def select_target_candidate(survey_geom, cad_layer, candidate_ids, mode='distance', exclusion_limit=2.0):
    best_candidate = None
    best_info = {
        'score': float('inf'),
        'dist': float('inf'),
        'angle': 0.0
    }

    for fid in candidate_ids:
        cad_feat = cad_layer.getFeature(fid)
        cad_geom = cad_feat.geometry()
        
        angle = get_angle_diff(survey_geom, cad_geom)
        score = float('inf')
        dist_metric = float('inf')

        if mode == 'distance':
            # [Mode A] Distance Based: '가장 가까운 점' 기준
            if angle > 45: continue
            
            # 단순 최단 거리 (Nearest Point)
            dist = survey_geom.distance(cad_geom)
            if dist > exclusion_limit: continue
            
            score = dist
            dist_metric = dist
            
        elif mode == 'original': # Area Based
            # [Mode B] Area/Shape Based: '전체적인 형상' 기준
            if angle > 20: continue # 평행성 우선 (엄격)
            
            # 하우스도르프 거리 (형상 유사도)
            hausdorff = survey_geom.hausdorffDistance(cad_geom)
            
            # 투영 중첩률 (Shift된 경우를 고려하여 버퍼 폭 확장)
            check_width = max(hausdorff, exclusion_limit) * 1.2
            overlap = check_projection_overlap(survey_geom, cad_geom, width=check_width)
            
            if overlap < 0.5: continue
            
            # 점수 산정: 하우스도르프 거리가 작을수록, 중첩이 길수록 좋음
            score = hausdorff * (1.0 - overlap + 0.1)
            dist_metric = hausdorff

        if score < best_info['score']:
            best_info['score'] = score
            best_info['dist'] = dist_metric
            best_info['angle'] = angle
            best_candidate = cad_feat

    return best_candidate, best_info

class NumericTableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        try:
            return float(self.text()) < float(other.text())
        except ValueError:
            return super().__lt__(other)

class GroupPreviewDialog(QDialog):
    """분석 전 그룹 확인 및 선택(제외) 다이얼로그"""
    def __init__(self, parent, groups, layer, canvas):
        super().__init__(parent)
        self.setWindowTitle("분석 그룹 조합")
        self.resize(450, 400)
        self.original_groups = groups
        self.layer = layer
        self.canvas = canvas
        self.selection_rubber_band = None
        
        # Data structure for managed groups: list of {'indices': [int], 'features': [QgsFeature]}
        self.analysis_groups = []
        
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("분석할 그룹을 조합하세요. 목록의 각 항목이 개별 분석 대상이 됩니다."))
        
        # --- 조합 버튼 ---
        merge_btn_layout = QHBoxLayout()
        self.btn_merge = QPushButton("선택 그룹 통합")
        self.btn_unmerge = QPushButton("선택 그룹 분리")
        self.btn_reset = QPushButton("초기화")
        merge_btn_layout.addWidget(self.btn_merge)
        merge_btn_layout.addWidget(self.btn_unmerge)
        merge_btn_layout.addWidget(self.btn_reset)
        layout.addLayout(merge_btn_layout)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        layout.addWidget(self.list_widget)
        
        self.reset_groups() # Initialize and populate

        # --- 시그널 연결 ---
        self.list_widget.itemSelectionChanged.connect(self.on_selection_changed)
        self.btn_merge.clicked.connect(self.merge_selected)
        self.btn_unmerge.clicked.connect(self.unmerge_selected)
        self.btn_reset.clicked.connect(self.reset_groups)
        
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def refresh_list_widget(self):
        self.list_widget.clear()
        for group_data in self.analysis_groups:
            name = "그룹 " + "+".join(str(i + 1) for i in sorted(group_data['indices']))
            count = len(group_data['features'])
            list_item = QListWidgetItem(f"{name} ({count}개 객체)")
            self.list_widget.addItem(list_item)

    def merge_selected(self):
        selected_rows = sorted([self.list_widget.row(item) for item in self.list_widget.selectedItems()], reverse=True)
        if len(selected_rows) < 2:
            QMessageBox.information(self, "알림", "통합할 그룹을 2개 이상 선택하세요.")
            return

        new_indices = set()
        new_features = []
        
        for row in selected_rows:
            group_to_merge = self.analysis_groups.pop(row)
            new_indices.update(group_to_merge['indices'])
            new_features.extend(group_to_merge['features'])

        self.analysis_groups.append({
            'indices': sorted(list(new_indices)),
            'features': new_features
        })
        self.refresh_list_widget()

    def unmerge_selected(self):
        selected_rows = sorted([self.list_widget.row(item) for item in self.list_widget.selectedItems()], reverse=True)
        if not selected_rows: return

        for row in selected_rows:
            group_to_unmerge = self.analysis_groups.pop(row)
            # Only unmerge if it's a merged group
            if len(group_to_unmerge['indices']) > 1:
                for original_idx in group_to_unmerge['indices']:
                    self.analysis_groups.append({
                        'indices': [original_idx],
                        'features': self.original_groups[original_idx]
                    })
        self.refresh_list_widget()

    def reset_groups(self):
        self.analysis_groups.clear()
        for i, group in enumerate(self.original_groups):
            self.analysis_groups.append({'indices': [i], 'features': group})
        if hasattr(self, 'list_widget'):
            self.refresh_list_widget()

    def get_selected_groups(self):
        # Returns a list of feature lists, corresponding to the final user-defined groups
        return [group['features'] for group in self.analysis_groups]

    def _get_group_geometry(self, features):
        if not features: return None
        geoms = [f.geometry() for f in features if f.geometry() and not f.geometry().isEmpty()]
        if not geoms: return None
        combined = QgsGeometry.collectGeometry(geoms)
        return combined.convexHull() if not combined.isEmpty() else None

    def on_selection_changed(self):
        if self.selection_rubber_band:
            self.canvas.scene().removeItem(self.selection_rubber_band)
            self.selection_rubber_band = None
        
        selected_items = self.list_widget.selectedItems()
        if not selected_items: return
        
        # Highlight the first selected item
        row = self.list_widget.row(selected_items[0])
        group_geom = self._get_group_geometry(self.analysis_groups[row]['features'])
        if not group_geom: return
        
        self.selection_rubber_band = QgsRubberBand(self.canvas, QgsWkbTypes.PolygonGeometry)
        self.selection_rubber_band.setToGeometry(group_geom, self.layer)
        self.selection_rubber_band.setFillColor(Qt.transparent)
        self.selection_rubber_band.setStrokeColor(QColor(255, 0, 0, 255))
        self.selection_rubber_band.setWidth(4)
        self.selection_rubber_band.show()
        
        ct = QgsCoordinateTransform(self.layer.crs(), self.canvas.mapSettings().destinationCrs(), QgsProject.instance())
        bbox_tr = ct.transform(group_geom.boundingBox())
        bbox_tr.grow(bbox_tr.width() * 0.2)
        self.canvas.setExtent(bbox_tr)
        self.canvas.refresh()

    def closeEvent(self, event):
        if self.selection_rubber_band:
            self.canvas.scene().removeItem(self.selection_rubber_band)
            self.selection_rubber_band = None
        super().closeEvent(event)

class GroupSelectionDialog(QDialog):
    """분석 결과 선택 및 지도 미리보기 다이얼로그"""
    def __init__(self, parent, display_items, groups, group_results, all_features, canvas, layer):
        super().__init__(parent)
        self.setModal(False) # [수정] 비모달(Modeless) 설정 -> 지도 조작 가능
        self.setWindowTitle("분석 결과 선택 (미리보기)")
        self.resize(450, 300)
        self.canvas = canvas
        self.layer = layer
        self.groups = groups
        self.group_results = group_results # [추가] 결과 데이터 저장
        self.all_features = all_features
        self.group_visuals = [] # (rubber_band, annotation)
        self.selection_rubber_band = None
        
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("목록에서 항목을 선택하면 지도에 해당 영역이 표시됩니다.\n적용할 결과를 선택하고 '확인'을 누르세요."))
        
        self.list_widget = QListWidget()
        self.list_widget.addItems(display_items)
        layout.addWidget(self.list_widget)
        
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)
        
        self.list_widget.currentRowChanged.connect(self.highlight_group)
        
        # [추가] 모든 그룹 시각화 (히트맵/라벨)
        self.visualize_all_groups()
        
    def _get_group_geometry(self, features):
        """선택된 객체들의 Convex Hull을 계산하여 대표 영역으로 반환"""
        if not features:
            return None

        # 1. 모든 유효한 지오메트리를 수집
        geoms = []
        for f in features:
            g = f.geometry()
            if g and not g.isEmpty():
                geoms.append(g)
        
        if not geoms:
            return None

        # 2. 지오메트리들을 하나로 합침
        combined_geom = QgsGeometry.collectGeometry(geoms)
        
        if combined_geom.isEmpty():
            return None
            
        # 3. Convex Hull을 계산하여 반환
        return combined_geom.convexHull()

    def visualize_all_groups(self):
        colors = [
            QColor(255, 0, 0, 80),    # Red
            QColor(0, 0, 255, 80),    # Blue
            QColor(0, 255, 0, 80),    # Green
            QColor(255, 165, 0, 80),  # Orange
            QColor(128, 0, 128, 80),  # Purple
            QColor(0, 255, 255, 80)   # Cyan
        ]
        
        for i, group_feats in enumerate(self.groups):
            if not group_feats: continue
            
            # [수정] BBox 대신 Convex Hull을 사용하여 그룹 영역 계산
            group_geom = self._get_group_geometry(group_feats)
            
            if not group_geom or group_geom.isEmpty():
                continue

            # RubberBand 생성 (영역 표시)
            color = colors[i % len(colors)]
            rb = QgsRubberBand(self.canvas, QgsWkbTypes.PolygonGeometry)
            rb.setToGeometry(group_geom, self.layer)
            rb.setColor(color)
            rb.setWidth(2)
            rb.setStrokeColor(QColor(color.red(), color.green(), color.blue(), 255))
            rb.show()
            
            # Annotation 생성 (라벨 표시)
            ann = None
            try:
                ann = QgsTextAnnotation()
                # [수정] 좌표 변환 (Layer CRS -> Canvas CRS)
                ct = QgsCoordinateTransform(self.layer.crs(), self.canvas.mapSettings().destinationCrs(), QgsProject.instance())
                center_pt = ct.transform(group_geom.centroid().asPoint())
                ann.setMapPosition(center_pt)
                doc = QTextDocument()
                doc.setHtml(f"<div style='color:black; font-weight:bold; font-size:15px; background-color:rgba(255,255,255,0.7); padding:2px;'>그룹 {i+1}</div>")
                # [수정] DeprecationWarning 수정: setFrameSize -> doc.setPageSize
                doc.setPageSize(QSizeF(80, 30))
                ann.setDocument(doc)
                ann.setFrameBackgroundColor(QColor(0, 0, 0, 0))
                QgsProject.instance().annotationManager().addAnnotation(ann)
            except Exception:
                pass
                
            self.group_visuals.append((rb, ann))

    def highlight_group(self, row):
        if self.selection_rubber_band:
            self.canvas.scene().removeItem(self.selection_rubber_band)
            self.selection_rubber_band = None
            
        if row < 0: return
        
        target_features = []
        # groups 리스트는 개별 그룹들만 포함. 마지막 '전체 통합' 항목은 all_features 사용
        if row < len(self.groups):
            target_features = self.groups[row]
        else:
            target_features = self.all_features
            
        # [수정] BBox 대신 Convex Hull을 사용하여 그룹 영역 계산
        group_geom = self._get_group_geometry(target_features)
        
        if not group_geom or group_geom.isEmpty():
            return
        
        # 선택된 영역 강조 (진한 테두리)
        self.selection_rubber_band = QgsRubberBand(self.canvas, QgsWkbTypes.PolygonGeometry)
        self.selection_rubber_band.setToGeometry(group_geom, self.layer)
        self.selection_rubber_band.setFillColor(Qt.transparent)
        self.selection_rubber_band.setStrokeColor(QColor(255, 0, 0, 255))
        self.selection_rubber_band.setWidth(4)
        self.selection_rubber_band.show()
        
        # 해당 영역이 잘 보이도록 지도 중심 이동
        # [수정] Convex Hull의 BBox를 기준으로 확대 (좌표 변환 적용)
        ct = QgsCoordinateTransform(self.layer.crs(), self.canvas.mapSettings().destinationCrs(), QgsProject.instance())
        bbox_tr = ct.transform(group_geom.boundingBox())
        # [수정] 사용자가 줌 조절을 할 수 있도록 20%의 여유 공간을 두고 확대
        bbox_tr.grow(bbox_tr.width() * 0.2)
        self.canvas.setExtent(bbox_tr)

    def closeEvent(self, event):
        if self.selection_rubber_band:
            self.canvas.scene().removeItem(self.selection_rubber_band)
            
        for rb, ann in self.group_visuals:
            if rb:
                self.canvas.scene().removeItem(rb)
            if ann:
                QgsProject.instance().annotationManager().removeAnnotation(ann)
                
        super().closeEvent(event)

class CadastralAuditor:
    def __init__(self, iface):
        self.iface = iface
        self.dlg = None
        self.action = None
        self.rubber_bands = []
        self.sel_dlg = None # [추가] 결과 선택 다이얼로그 참조 유지
        self.preview_visuals = [] # [추가] 그룹 미리보기 시각화 요소
        
        # Flash effect variables
        self.flash_timer = QTimer()
        self.flash_rubber_band = None
        self.flash_counter = 0
        self.flash_timer.timeout.connect(self.flash_tick)
        
        # Track accumulated shift
        self.accumulated_dx = 0.0
        self.accumulated_dy = 0.0
        
        # Error visualization
        self.error_lines = {}
        self.error_rubber_band = None
        self.error_markers = []
        
        # [Added] Debug visualization
        self.debug_data = {}
        self.debug_rubber_bands = []
        
        # [Added] Control Points Tool
        self.cp_tool = None
        self.cp_capture_step = 0 # 0: Source, 1: Target
        self.current_cp_source = None

    def initGui(self):
        self.action = QAction("Cadastral Auditor", self.iface.mainWindow())
        self.action.triggered.connect(self.run)
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToMenu("&Cadastral Auditor", self.action)

    def unload(self):
        self.iface.removeToolBarIcon(self.action)
        self.iface.removePluginMenu("&Cadastral Auditor", self.action)
        self.clear_highlights()
        if self.flash_timer.isActive():
            self.flash_timer.stop()
        if self.flash_rubber_band:
            self.iface.mapCanvas().scene().removeItem(self.flash_rubber_band)
        if self.error_rubber_band:
            self.iface.mapCanvas().scene().removeItem(self.error_rubber_band)
            self.error_rubber_band = None
        for m in self.error_markers:
            self.iface.mapCanvas().scene().removeItem(m)
        self.error_markers.clear()
        self.clear_preview_visuals()
        self.clear_debug_visuals()

    def run(self):
        if not self.dlg:
            self.dlg = CadastralAuditorDialog()
            self.dlg.btn_analyze.clicked.connect(self.run_analysis)
            self.dlg.btn_apply_shift.clicked.connect(self.apply_feature_shift)
            self.dlg.btn_export.clicked.connect(self.export_to_excel)
            self.dlg.btn_auto_calc.clicked.connect(self.calculate_optimal_shift)
            self.dlg.btn_select_area.clicked.connect(self.enable_polygon_selection)
            self.dlg.btn_select_multi.clicked.connect(self.enable_multi_selection)
            self.dlg.table_results.cellClicked.connect(self.zoom_to_feature)
            
            # Control Points Connections
            self.dlg.btn_add_cp.clicked.connect(self.start_cp_capture)
            self.dlg.btn_clear_cp.clicked.connect(self.clear_control_points)
            
            # Connect Nudge Buttons
            # Fix: lambda must accept the 'checked' argument from clicked signal
            self.dlg.btn_up.clicked.connect(lambda _: self.nudge_layer(0, 1))
            self.dlg.btn_down.clicked.connect(lambda _: self.nudge_layer(0, -1))
            self.dlg.btn_left.clicked.connect(lambda _: self.nudge_layer(-1, 0))
            self.dlg.btn_right.clicked.connect(lambda _: self.nudge_layer(1, 0))
            self.dlg.btn_ul.clicked.connect(lambda _: self.nudge_layer(-1, 1))
            self.dlg.btn_ur.clicked.connect(lambda _: self.nudge_layer(1, 1))
            self.dlg.btn_dl.clicked.connect(lambda _: self.nudge_layer(-1, -1))
            self.dlg.btn_dr.clicked.connect(lambda _: self.nudge_layer(1, -1))
        
        self.refresh_layers()
        self.dlg.show()
        # self.dlg.exec_() # [수정] 모달 실행 제거 -> 비모달 실행 (지도 조작 가능)

    def enable_polygon_selection(self):
        """다각형으로 객체 선택 도구 활성화"""
        self.iface.actionSelectPolygon().trigger()
        self.iface.messageBar().pushMessage("영역 선택", "지도에 다각형을 그려서 분석할 객체들을 선택하세요. (새로운 선택)", level=Qgis.Info)

    def enable_multi_selection(self):
        """다중 영역 선택 (추가 선택) 도구 활성화"""
        self.iface.actionSelectPolygon().trigger()
        self.iface.messageBar().pushMessage("다중 영역 선택", "Shift 키를 누른 상태로 다각형을 그려서 영역을 추가하세요.", level=Qgis.Info)

    def start_cp_capture(self):
        """가중치 지점 입력 도구 시작"""
        cad_layer = self.dlg.cb_layer_cadastral.currentData()
        if not cad_layer:
            QMessageBox.warning(self.dlg, "오류", "지적도 레이어를 먼저 선택해주세요.")
            return
            
        self.cp_tool = QgsMapToolEmitPoint(self.iface.mapCanvas())
        self.cp_tool.canvasClicked.connect(self.on_cp_clicked)
        self.iface.mapCanvas().setMapTool(self.cp_tool)
        self.cp_capture_step = 0
        self.current_cp_source = None
        self.iface.messageBar().pushMessage("지점 추가", "이동할 지점(Source - 현황선)을 클릭하세요.", level=Qgis.Info)

    def on_cp_clicked(self, point, button):
        """지도 클릭 이벤트 처리"""
        # 좌표 변환 준비 (Canvas -> Cadastral CRS)
        # 최적화 로직이 Cadastral CRS 기준이므로, 입력 좌표도 이에 맞춤
        canvas = self.iface.mapCanvas()
        cad_layer = self.dlg.cb_layer_cadastral.currentData()
        ct = QgsCoordinateTransform(canvas.mapSettings().destinationCrs(), cad_layer.crs(), QgsProject.instance())
        
        transformed_pt = ct.transform(point)
        
        if self.cp_capture_step == 0:
            self.current_cp_source = transformed_pt
            self.cp_capture_step = 1
            self.iface.messageBar().pushMessage("지점 추가", "목표 지점(Target - 지적선)을 클릭하세요.", level=Qgis.Info)
        elif self.cp_capture_step == 1:
            # 테이블에 추가
            row = self.dlg.tbl_control_points.rowCount()
            self.dlg.tbl_control_points.insertRow(row)
            
            # 소수점 3자리로 표시
            self.dlg.tbl_control_points.setItem(row, 0, QTableWidgetItem(f"{self.current_cp_source.x():.3f}"))
            self.dlg.tbl_control_points.setItem(row, 1, QTableWidgetItem(f"{self.current_cp_source.y():.3f}"))
            self.dlg.tbl_control_points.setItem(row, 2, QTableWidgetItem(f"{transformed_pt.x():.3f}"))
            self.dlg.tbl_control_points.setItem(row, 3, QTableWidgetItem(f"{transformed_pt.y():.3f}"))
            
            weight = self.dlg.sb_cp_weight.value()
            self.dlg.tbl_control_points.setItem(row, 4, QTableWidgetItem(str(weight)))
            
            self.iface.mapCanvas().unsetMapTool(self.cp_tool)
            self.iface.messageBar().pushMessage("완료", "가중치 지점이 추가되었습니다.", level=Qgis.Success)
            
            # 시각적 피드백 (선 그리기 등)은 생략하거나 추후 추가
            self.dlg.activateWindow() # 다이얼로그 다시 활성화

    def clear_control_points(self):
        self.dlg.tbl_control_points.setRowCount(0)
        self.iface.messageBar().pushMessage("초기화", "가중치 지점 목록을 비웠습니다.", level=Qgis.Info)

    def refresh_layers(self):
        self.dlg.cb_layer_cadastral.clear()
        self.dlg.cb_layer_current.clear()
        self.dlg.cb_layer_target.clear()
        
        layers = QgsProject.instance().mapLayers().values()
        for layer in layers:
            if layer.type() == QgsMapLayer.VectorLayer and \
               layer.geometryType() in [QgsWkbTypes.PolygonGeometry, QgsWkbTypes.LineGeometry]:
                self.dlg.cb_layer_cadastral.addItem(layer.name(), layer)
                self.dlg.cb_layer_current.addItem(layer.name(), layer)
                self.dlg.cb_layer_target.addItem(layer.name(), layer)

    def clear_highlights(self):
        for rb in self.rubber_bands:
            self.iface.mapCanvas().scene().removeItem(rb)
        self.rubber_bands.clear()
        self.clear_preview_visuals()
        self.remove_error_visuals()

    def run_analysis(self):
        # 1. Setup
        print("DEBUG: === 분석 시작 ===")
        self.dlg.clear_table()
        self.dlg.table_results.setSortingEnabled(False) # Disable sorting during insertion
        self.dlg.lbl_summary.setText("분석 중...")
        self.clear_highlights()
        self.clear_preview_visuals()
        self.accumulated_dx = 0.0
        self.accumulated_dy = 0.0
        self.dlg.sb_shift_x.setValue(0.0)
        self.dlg.sb_shift_y.setValue(0.0)
        self.error_lines = {}
        self.clear_preview_visuals()
        self.debug_data = {}
        
        # Use currentLayer() instead of currentData() for consistency with request
        # Assuming populate_layers adds layer object as data, currentData() is safer but request uses currentLayer() logic
        # We will stick to currentData() as it was in original code and is robust for QComboBox with UserRole
        cad_layer = self.dlg.cb_layer_cadastral.currentData()
        cur_layer = self.dlg.cb_layer_current.currentData()
        
        if not cad_layer or not cur_layer:
            QMessageBox.warning(self.dlg, "오류", "레이어를 선택해주세요.")
            return
            
        # UI 설정값 로드
        tol_min = self.dlg.sb_tol_min.value()       # 양호 범위 (위치 오차 기준)
        tol_max = self.dlg.sb_tol_max.value()       # 한계 범위 (형상 오차 기준)
        exclusion_limit = self.dlg.sb_exclusion_limit.value() # 검색 반경

        # [추가] 모드 확인
        selected_mode = 'distance'
        if self.dlg.rb_mode_original.isChecked():
            selected_mode = 'original'

        # [Task 1] 좌표계 변환기 생성 (Current -> Cadastral)
        # 분석 루프 시작 전, 두 레이어의 CRS를 비교하여 QgsCoordinateTransform 객체를 생성
        crs_transform = QgsCoordinateTransform(
            cur_layer.crs(), 
            cad_layer.crs(), 
            QgsProject.instance()
        )

        # 2. Build Spatial Index for Cadastral Layer
        index = QgsSpatialIndex(cad_layer.getFeatures())
        matcher = ConfidenceStringMatcher(sigma=0.1)
        
        cur_features = list(cur_layer.getFeatures())
        total = len(cur_features)
        self.dlg.progress_bar.setMaximum(total)
        
        row = 0
        
        # Statistics Counters
        stats = {
            "pass": 0,
            "shift": 0,
            "shape": 0,
            "critical": 0
        }
        
        # 3. Iterate Current Features
        for i, cur_feat in enumerate(cur_features):
            self.dlg.progress_bar.setValue(i + 1)
            # print(f"DEBUG: [Feature {i+1}] ID: {cur_feat.id()} 처리 중...")
            
            # [Task 1] 좌표 변환 (Clone 후 변환하여 원본 보존)
            # 현황선(Current)의 형상을 지적도 좌표계로 변환(transform)한 후 검색 및 거리 계산을 수행
            geom_curr = QgsGeometry(cur_feat.geometry()) # 복제
            if not geom_curr: continue
            
            try:
                geom_curr.transform(crs_transform) # 지적도 좌표계로 변환
            except Exception as e:
                print(f"DEBUG: Geometry transform failed for feature {cur_feat.id()}: {e}")
                continue
            
            # Create a temporary feature with transformed geometry
            # This is needed because process_pair expects a feature, or we can modify process_pair to accept geometry
            # The provided guide passes geom_curr to process_pair, but process_pair in analyzer.py expects feature.
            # We will create a temp feature to maintain compatibility with analyzer.py
            cur_feat_tr = QgsFeature(cur_feat)
            cur_feat_tr.setGeometry(geom_curr)
            
            # 공간 검색
            # [Mode Check] 면적 모드일 경우 탐색 반경 확장 (Shift 감지용)
            search_radius = exclusion_limit
            if selected_mode == 'original':
                search_radius = exclusion_limit * 3.0
                
            search_rect = geom_curr.boundingBox()
            search_rect.grow(search_radius)
            candidate_ids = index.intersects(search_rect)
            
            # [Task 2] 최적 후보 선정 (Best Fit Selection)
            best_match_feat, best_info = select_target_candidate(
                geom_curr, cad_layer, candidate_ids, 
                mode=selected_mode, exclusion_limit=exclusion_limit
            )
            
            min_dist = best_info['dist']
            best_angle_diff = best_info['angle']
            
            # [Task 3] 매칭 결과 처리
            if best_match_feat:
                # 정밀 분석 수행 (변환된 geom 사용)
                result = matcher.process_pair(
                    cur_feat_tr, 
                    best_match_feat, 
                    th_shape=tol_max, 
                    th_pos=tol_min,
                    mode=selected_mode
                )
                
                topology = result['topology']
                score = result['score']
                status = result['status']
                nd_cost = result['nd_cost']
                shift_x = result.get('shift_x', 0.0)
                shift_y = result.get('shift_y', 0.0)
                
                print(f"DEBUG:   -> 위상: {topology}, 점수: {score:.3f}, 판정: {status}")
                
                # Determine precise distance for display (Consistency with Status)
                display_dist = min_dist
                if selected_mode == 'distance':
                    display_dist = score
                elif 'position_error' in result:
                    display_dist = result['position_error']
                
                # [수정] 상대 거리가 제외 범위를 초과하면 평가에서 제외 (테이블에 추가하지 않음)
                if display_dist > exclusion_limit:
                    continue
                
                # Update Statistics (Use 'status' to match Table display)
                # [Fix] "불부합" contains "부합", so check "불부합" first or use startswith
                if "불부합" in status:
                    stats["critical"] += 1
                elif "부합" in status:  # Now safe to check "부합"
                    stats["pass"] += 1
                elif "위치정정" in status:
                    stats["shift"] += 1
                elif "형상" in status:
                    stats["shape"] += 1
                else:
                    stats["critical"] += 1
                
                # 시각화: 불부합일 경우 하이라이트
                # if status == "불부합":
                #     self.highlight_feature(best_match_feat, cad_layer) # Highlight cadastral feature instead

                # [시각화] 비교 점 및 오차 벡터 생성 (PointToLineAuditor 활용)
                # 사용자가 지적선과 현형선의 비교 대상 점을 시각적으로 확인할 수 있도록 함
                p2l = PointToLineAuditor(best_match_feat, cur_feat_tr, densify_distance=1.0)
                p2l_res = p2l.process()
                error_geom = p2l_res["error_vectors"]
                if error_geom and not error_geom.isEmpty():
                    self.error_lines[cur_feat.id()] = (error_geom, cad_layer.crs())
                
                # [추가] 시각화 데이터 저장 (Visual Debugging Data)
                debug_info = {}
                
                # 1. 타겟 경계선 (Target Boundary)
                tgt_geom = QgsGeometry(best_match_feat.geometry())
                if tgt_geom.type() == QgsWkbTypes.PolygonGeometry:
                    tgt_geom.convertToType(QgsWkbTypes.LineGeometry)
                debug_info['target'] = tgt_geom
                
                # 2. 버퍼 영역 (Search/Overlap Buffer)
                if selected_mode == 'original':
                    h_dist = geom_curr.hausdorffDistance(tgt_geom)
                    check_width = max(h_dist, 0.5) * 1.5
                    debug_info['buffer'] = tgt_geom.buffer(check_width, 5)
                else:
                    debug_info['buffer'] = geom_curr.buffer(exclusion_limit, 5)

                # 3. 오차 벡터
                if error_geom and not error_geom.isEmpty():
                    debug_info['vectors'] = error_geom
                
                self.debug_data[cur_feat.id()] = debug_info

                # Add to Table (Inline implementation of add_result_row logic)
                self.dlg.table_results.insertRow(row)
                item_serial = NumericTableWidgetItem(str(i+1))
                item_serial.setData(Qt.UserRole, cur_feat.id())
                self.dlg.table_results.setItem(row, 0, item_serial)
                
                # Try to get PNU or JIBUN from cadastral layer (assuming field exists, else ID)
                match_label = str(best_match_feat.id())
                for field in ['jibun', 'pnu', 'JIBUN', 'PNU']:
                    if best_match_feat.fieldNameIndex(field) != -1:
                        match_label = str(best_match_feat[field])
                        break
                
                self.dlg.table_results.setItem(row, 1, QTableWidgetItem(match_label))
                self.dlg.table_results.setItem(row, 2, NumericTableWidgetItem(f"{display_dist:.3f}"))
                self.dlg.table_results.setItem(row, 3, NumericTableWidgetItem(f"{best_angle_diff:.1f}"))
                self.dlg.table_results.setItem(row, 4, QTableWidgetItem(topology))
                self.dlg.table_results.setItem(row, 5, NumericTableWidgetItem(f"{score:.3f}"))
                self.dlg.table_results.setItem(row, 6, QTableWidgetItem(status))
                self.dlg.table_results.setItem(row, 7, NumericTableWidgetItem(f"{nd_cost:.3f}"))
                self.dlg.table_results.setItem(row, 8, NumericTableWidgetItem(f"{shift_x:.3f}"))
                self.dlg.table_results.setItem(row, 9, NumericTableWidgetItem(f"{shift_y:.3f}"))
                
                item_status = self.dlg.table_results.item(row, 6)
                if "불부합" in status:
                    item_status.setForeground(QColor("red"))
                elif "부합" in status:
                    item_status.setForeground(QColor("green"))
                elif "위치정정" in status:
                    item_status.setForeground(QColor("blue"))
                elif "회전" in status:
                    item_status.setForeground(QColor(204, 204, 0)) # Dark Yellow/Gold
                elif "형상 불일치" in status:
                    item_status.setForeground(QColor(255, 140, 0)) # Dark Orange
                
                row += 1

        # Calculate and Display Summary
        total_processed = row
        pass_rate = (stats["pass"] / total_processed * 100) if total_processed > 0 else 0.0
        
        summary_text = f"총 {total_processed}건 중 부합 {stats['pass']}건 (부합율: {pass_rate:.1f}%) | 위치정정: {stats['shift']} | 형상불일치: {stats['shape']} | 불부합: {stats['critical']}"
        self.dlg.lbl_summary.setText(summary_text)

        self.dlg.table_results.setSortingEnabled(True) # Re-enable sorting
        self.dlg.btn_export.setEnabled(True)
        self.dlg.btn_apply_shift.setEnabled(True)
        QMessageBox.information(self.dlg, "완료", "분석이 완료되었습니다.")

    def classify_result(self, error, min_val, max_val):
        if error <= min_val:
            return "Good"
        elif error <= max_val:
            return "Warning"
        else:
            return "Critical"

    def clear_debug_visuals(self):
        for rb in self.debug_rubber_bands:
            try:
                self.iface.mapCanvas().scene().removeItem(rb)
            except:
                pass
        self.debug_rubber_bands.clear()

    def draw_debug_geometry(self, geometry, color, width=2, line_style=Qt.SolidLine, is_polygon=False):
        if not geometry: return
        
        canvas = self.iface.mapCanvas()
        dest_crs = canvas.mapSettings().destinationCrs()
        cad_layer = self.dlg.cb_layer_cadastral.currentData()
        if not cad_layer: return
        
        ct = QgsCoordinateTransform(cad_layer.crs(), dest_crs, QgsProject.instance())
        
        geom_canvas = QgsGeometry(geometry)
        try:
            geom_canvas.transform(ct)
        except:
            return

        rb_type = QgsWkbTypes.PolygonGeometry if is_polygon else QgsWkbTypes.LineGeometry
        rb = QgsRubberBand(canvas, rb_type)
        rb.setToGeometry(geom_canvas, None)
        rb.setWidth(width)
        rb.setLineStyle(line_style)
        
        if is_polygon:
            rb.setFillColor(color)
            rb.setStrokeColor(Qt.transparent)
        else:
            rb.setStrokeColor(color)
            rb.setFillColor(Qt.transparent)
        
        rb.show()
        self.debug_rubber_bands.append(rb)

    def zoom_to_feature(self, row, column):
        item = self.dlg.table_results.item(row, 0)
        if not item:
            return
            
        fid = item.data(Qt.UserRole)
        cur_layer = self.dlg.cb_layer_current.currentData()
        
        if cur_layer and fid is not None:
            feat = cur_layer.getFeature(fid)
            if feat.isValid():
                # Transform coordinates to Map Canvas CRS for Zoom
                canvas = self.iface.mapCanvas()
                dest_crs = canvas.mapSettings().destinationCrs()
                ct = QgsCoordinateTransform(cur_layer.crs(), dest_crs, QgsProject.instance())
                
                geom = feat.geometry()
                
                if self.dlg.chk_fixed_scale.isChecked():
                    center = geom.centroid().asPoint()
                    center_tr = ct.transform(center)
                    canvas.setCenter(center_tr)
                else:
                    bbox = geom.boundingBox()
                    bbox.grow(5.0) # Increased buffer
                    bbox_geom = QgsGeometry.fromRect(bbox)
                    bbox_geom.transform(ct)
                    canvas.setExtent(bbox_geom.boundingBox())
                canvas.refresh()
                self.flash_feature(feat, cur_layer)
                
                # Clear previous error visuals
                self.clear_debug_visuals()
                self.remove_error_visuals()
                
                if fid in self.debug_data:
                    data = self.debug_data[fid]
                    
                    # 1. Buffer (연두색 반투명)
                    if 'buffer' in data:
                        self.draw_debug_geometry(data['buffer'], QColor(0, 255, 0, 40), is_polygon=True)
                    
                    # 2. Target (노란색 실선) - 비교 대상 지적선
                    if 'target' in data:
                        self.draw_debug_geometry(data['target'], QColor(255, 255, 0), width=1)
                        
                    # 3. Vectors (마젠타 점선)
                    if 'vectors' in data:
                        self.draw_debug_geometry(data['vectors'], QColor(255, 0, 255), width=2, line_style=Qt.DashLine)

    def apply_feature_shift(self):
        row = self.dlg.table_results.currentRow()
        if row < 0:
            QMessageBox.warning(self.dlg, "경고", "기준이 될 객체를 목록에서 선택해주세요.")
            return
        
        # Get shift values from table directly (Delta), not from spinboxes (Accumulated)
        try:
            dx = float(self.dlg.table_results.item(row, 8).text())
            dy = float(self.dlg.table_results.item(row, 9).text())
        except (ValueError, AttributeError):
            return
            
        if dx == 0 and dy == 0:
            QMessageBox.information(self.dlg, "알림", "이동할 거리가 없습니다 (0, 0).")
            return

        layer = self.dlg.cb_layer_target.currentData()
        if layer:
            if not layer.isEditable():
                layer.startEditing()
            
            # Apply shift to ALL features in the layer
            for feat in layer.getFeatures():
                geom = feat.geometry()
                geom.translate(dx, dy)
                layer.changeGeometry(feat.id(), geom)
            
            self.accumulated_dx += dx
            self.accumulated_dy += dy
            self.dlg.sb_shift_x.setValue(self.accumulated_dx)
            self.dlg.sb_shift_y.setValue(self.accumulated_dy)
            
            self.iface.mapCanvas().refresh()
            QMessageBox.information(self.dlg, "이동 완료", f"현형선 전체를 이동했습니다.\n\n이번 이동: {dx:.3f}m, {dy:.3f}m\n누적 이동: {self.accumulated_dx:.3f}m, {self.accumulated_dy:.3f}m")

    def _cluster_features(self, features, threshold=5.0):
        """선택된 객체들을 거리 기반으로 그룹화 (Connected Components)"""
        if not features: return []
        
        # [수정 1] 거리가 0일 때는 완전히 맞닿은(거리가 0인) 객체들을 묶어주도록 변경
        # 음수(-1 등)를 입력했을 때만 강제로 모든 객체를 개별 그룹화합니다.
        if threshold < 0:
            return [[f] for f in features]
        
        id_map = {f.id(): f for f in features}
        pool = set(id_map.keys())
        index = QgsSpatialIndex()
        index.addFeatures(features)
        
        groups = []
        
        while pool:
            seed_id = next(iter(pool))
            pool.remove(seed_id)
            
            current_group = [id_map[seed_id]]
            queue = [seed_id]
            
            while queue:
                current_id = queue.pop(0)
                curr_feat = id_map[current_id]
                geom = curr_feat.geometry()
                
                # Search neighbors
                bbox = geom.boundingBox()
                if threshold > 0:
                    bbox.grow(threshold) # 임계값만큼 검색 범위 확장
                
                candidate_ids = index.intersects(bbox)
                
                for cid in candidate_ids:
                    if cid in pool:
                        neighbor_geom = id_map[cid].geometry()
                        # [수정 2] Bounding Box 교차 여부뿐만 아니라, 
                        # '실제 형상 간의 최단 거리'를 한 번 더 정밀하게 검사하여 False Positive 방지
                        if geom.distance(neighbor_geom) <= threshold:
                            pool.remove(cid)
                            queue.append(cid)
                            current_group.append(id_map[cid])
            
            groups.append(current_group)
            
        return groups

    def _prepare_test_data(self, features, cad_layer, crs_transform, exclusion_limit, selected_mode):
        """최적화에 사용할 매칭 쌍(Test Data) 추출"""
        test_data = []
        cad_index = QgsSpatialIndex(cad_layer.getFeatures())
        
        self.dlg.progress_bar.setMaximum(len(features))
        self.dlg.lbl_summary.setText("매칭 대상 추출 중...")
        
        for i, f in enumerate(features):
            if i % 50 == 0: QApplication.processEvents()
            self.dlg.progress_bar.setValue(i + 1)
            
            g = QgsGeometry(f.geometry())
            try:
                g.transform(crs_transform)
                if g.isEmpty(): continue
                
                if g.type() == QgsWkbTypes.PolygonGeometry:
                    g = g.boundary()
                
                search_radius = exclusion_limit
                if selected_mode == 'original':
                    search_radius = exclusion_limit * 3.0
                
                bbox = g.boundingBox()
                bbox.grow(search_radius) 
                cids = cad_index.intersects(bbox)
                
                best_match_feat, best_info = select_target_candidate(
                    g, cad_layer, cids, 
                    mode=selected_mode, exclusion_limit=exclusion_limit
                )
                
                if best_match_feat:
                    best_cad = QgsGeometry(best_match_feat.geometry())
                    sample_pts = []
                    densified = g.densifyByDistance(1.0)
                    for v in densified.vertices():
                        sample_pts.append(QgsPointXY(v.x(), v.y()))
                    
                    if not sample_pts:
                        sample_pts = [QgsPointXY(v.x(), v.y()) for v in g.vertices()]
                    
                    # [수정됨] 매칭된 두 기하학적 형상의 중심점을 함께 저장 (미끄러짐 방지용)
                    cur_centroid = g.centroid().asPoint()
                    cad_centroid = best_cad.centroid().asPoint()
                    
                    test_data.append((sample_pts, best_cad, cur_centroid, cad_centroid))
                    
            except Exception:
                continue
                
        return test_data

    def calculate_optimal_shift(self):
        """
        지정된 범위 내에서 그리드 탐색을 수행하여 평균 거리 오차(MAE)가 최소가 되는 이동량을 찾습니다.
        선택된 객체 수와 그룹에 따라 시나리오별(전체/단일표본/비교분석) 로직을 수행합니다.
        """
        cur_layer = self.dlg.cb_layer_current.currentData()
        cad_layer = self.dlg.cb_layer_cadastral.currentData()
        
        if not cur_layer or not cad_layer:
            QMessageBox.warning(self.dlg, "오류", "레이어를 선택해주세요.")
            return

        # [수정] 그룹 미리보기 및 사용자 확인 단계
        selected_count = cur_layer.selectedFeatureCount()
        groups = []
        if selected_count > 0:
            self.clear_preview_visuals()
            
            selected_features = list(cur_layer.selectedFeatures())
            cluster_threshold = self.dlg.sb_cluster_dist.value()
            groups = self._cluster_features(selected_features, threshold=cluster_threshold)
            
            self.preview_visuals = self.visualize_groups(groups, cur_layer)
            
            # [수정] 그룹 선택 다이얼로그 표시 (체크박스로 제외 가능)
            preview_dlg = GroupPreviewDialog(self.dlg, groups, cur_layer, self.iface.mapCanvas())
            if preview_dlg.exec_() == QDialog.Accepted:
                groups = preview_dlg.get_selected_groups()
                if not groups:
                    self.iface.messageBar().pushMessage("알림", "선택된 그룹이 없습니다. 분석을 취소합니다.", level=Qgis.Warning)
                    self.clear_preview_visuals()
                    return
                
                # [수정] 사용자가 선택한 그룹의 객체들로 selected_features 갱신 (제외된 그룹은 전체 통합 분석에서도 빠지도록 함)
                selected_features = [f for group in groups for f in group]
            else:
                self.iface.messageBar().pushMessage("알림", "분석이 취소되었습니다.", level=Qgis.Info)
                self.clear_preview_visuals()
                return
            self.clear_preview_visuals()

        # 1. 시나리오 판단 (Selection Logic)
        selected_count = cur_layer.selectedFeatureCount()
        target_features = []
        mode_desc = ""
        is_comparative = False

        if selected_count == 0:
            # [Scenario 1] 선택 없음 -> 전체 분석 모드
            target_features = list(cur_layer.getFeatures())
            mode_desc = "전체 분석 모드"
        else:
            # 그룹화 결과는 미리보기 단계에서 이미 계산됨
            if len(groups) == 1:
                mode_desc = "단일 표본 분석 모드 (우량 표본)"
                target_features = groups[0]
            else:
                mode_desc = "비교 분석 모드"
                is_comparative = True

        # 2. 공통 파라미터 준비
        search_range = self.dlg.sb_search_range.value()
        step = self.dlg.sb_search_step.value()
        if step <= 0: return

        exclusion_limit = self.dlg.sb_exclusion_limit.value()
        selected_mode = 'distance'
        if self.dlg.rb_mode_original.isChecked():
            selected_mode = 'original'
        
        crs_transform = QgsCoordinateTransform(cur_layer.crs(), cad_layer.crs(), QgsProject.instance())
        is_iterative = self.dlg.chk_iterative.isChecked()

        # 가중치 지점 파싱 (Global Control Points)
        control_points = []
        if self.dlg.tbl_control_points.rowCount() > 0:
            for r in range(self.dlg.tbl_control_points.rowCount()):
                try:
                    sx = float(self.dlg.tbl_control_points.item(r, 0).text())
                    sy = float(self.dlg.tbl_control_points.item(r, 1).text())
                    tx = float(self.dlg.tbl_control_points.item(r, 2).text())
                    ty = float(self.dlg.tbl_control_points.item(r, 3).text())
                    w  = float(self.dlg.tbl_control_points.item(r, 4).text())
                    control_points.append((QgsPointXY(sx, sy), QgsPointXY(tx, ty), w))
                except ValueError:
                    continue
            
        # 3. 실행 로직 분기
        if is_comparative:
            # [비교 분석 모드]
            group_results = []
            display_items = []
            
            for idx, group_feats in enumerate(groups):
                self.dlg.lbl_summary.setText(f"그룹 {idx+1} 분석 중... ({len(group_feats)}개)")
                test_data = self._prepare_test_data(group_feats, cad_layer, crs_transform, exclusion_limit, selected_mode)
                
                if not test_data:
                    # [수정] 매칭 실패 시에도 객체 수를 표시하여 누락 오해 방지
                    display_items.append(f"그룹 {idx+1} ({len(group_feats)}개): 매칭 실패 (매칭 대상 없음)")
                    group_results.append(None)
                    continue
                
                # 해당 그룹 범위 내의 Control Point만 필터링하여 적용 (Optional)
                # 여기서는 단순화를 위해 전체 CP 적용 (사용자가 알아서 입력했다고 가정)
                dx, dy, init_mae, final_mae = self._run_optimization_core(test_data, control_points, search_range, step, is_iterative)
                group_results.append((dx, dy))
                display_items.append(f"그룹 {idx+1} ({len(group_feats)}개): X={dx:+.3f}m, Y={dy:+.3f}m (오차 {init_mae:.3f} -> {final_mae:.3f})")
            
            # [추가] 전체 통합 분석 결과 추가 (선택된 모든 객체 대상)
            self.dlg.lbl_summary.setText(f"전체 통합 분석 중...")
            global_test_data = self._prepare_test_data(selected_features, cad_layer, crs_transform, exclusion_limit, selected_mode)
            
            if global_test_data:
                g_dx, g_dy, g_init, g_final = self._run_optimization_core(global_test_data, control_points, search_range, step, is_iterative)
                group_results.append((g_dx, g_dy))
                display_items.append(f"[전체 통합] ({len(selected_features)}개): X={g_dx:+.3f}m, Y={g_dy:+.3f}m (오차 {g_init:.3f} -> {g_final:.3f})")
            
            self.dlg.lbl_summary.setText("비교 분석 완료")
            
            # [수정] 사용자 정의 다이얼로그 (지도 하이라이트 기능 포함)
            if self.sel_dlg:
                self.sel_dlg.close()
                
            self.sel_dlg = GroupSelectionDialog(self.dlg, display_items, groups, group_results, selected_features, self.iface.mapCanvas(), cur_layer)
            self.sel_dlg.accepted.connect(self.on_group_selection_accepted)
            self.sel_dlg.show()
            
            return

        else:
            # [전체 분석] 또는 [단일 표본 분석]
            test_data = self._prepare_test_data(target_features, cad_layer, crs_transform, exclusion_limit, selected_mode)
            
            if not test_data:
                QMessageBox.warning(self.dlg, "실패", "매칭 가능한 객체가 없어 최적화를 수행할 수 없습니다.")
                return

            # 사용자 확인용 요약
            self.dlg.lbl_summary.setText(f"최적화 대상: {len(test_data)}개 (제외범위 {exclusion_limit}m 이내)")
            QApplication.processEvents()

            # 최적화 실행
            best_dx, best_dy, initial_avg_mae, final_avg_mae = self._run_optimization_core(
                test_data, control_points, search_range, step, is_iterative
            )
            
            # 결과 적용 확인
            self._show_apply_dialog(best_dx, best_dy, initial_avg_mae, final_avg_mae, mode_desc, is_iterative)

    def on_group_selection_accepted(self):
        """결과 선택 다이얼로그에서 확인 버튼을 눌렀을 때 호출"""
        if not self.sel_dlg: return
        
        row = self.sel_dlg.list_widget.currentRow()
        if row >= 0:
            item_text = self.sel_dlg.list_widget.item(row).text()
            res = self.sel_dlg.group_results[row]
        
            if res:
                dx, dy = res
                # 적용 확인
                target_name = item_text.split(':')[0]
                msg = (f"선택된 [{target_name}]의 결과:\n"
                       f"이동량 X: {dx:.3f} m\n"
                       f"이동량 Y: {dy:.3f} m\n\n"
                       "이 이동량을 전체 레이어에 적용하시겠습니까?")
                       
                if QMessageBox.question(self.dlg, "적용 확인", msg, QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                    self.apply_shift_vector(dx, dy)

    def _run_optimization_core(self, test_data, control_points, search_range, step, is_iterative):
        """실제 그리드 탐색 로직 (Core) - 거리 기반 판정과 일치하도록 최적화"""
        if not test_data:
            return 0.0, 0.0, 0.0, 0.0
            
        max_iter = 10 if is_iterative else 1
        
        accum_dx = 0.0
        accum_dy = 0.0
        
        # [수정됨] 초기 상태의 오차 계산 (Baseline)
        total_error_sum_initial = 0.0
        for pts, cad_geom, cur_centroid, cad_centroid in test_data:
            pair_dist_sum = 0.0
            max_dist = 0.0
            for pt in pts:
                dist = cad_geom.distance(QgsGeometry.fromPointXY(pt))
                pair_dist_sum += dist
                if dist > max_dist:
                    max_dist = dist
            avg_dist = pair_dist_sum / len(pts)
            centroid_dist = math.sqrt(cur_centroid.sqrDist(cad_centroid))
            
            # 단순 평균 거리가 아닌, 직관적인 최대 오차를 무겁게 반영 (평균 30% + 최대거리 60% + 미끄러짐 방지 10%)
            total_error_sum_initial += (avg_dist * 0.3) + (max_dist * 0.6) + (centroid_dist * 0.1)
            
        # 가중치 지점 초기 오차 추가
        cp_weight_sum = 0.0
        for src, dst, w in control_points:
            total_error_sum_initial += math.sqrt(src.sqrDist(dst)) * w
            cp_weight_sum += w
        
        total_weight = len(test_data) + cp_weight_sum
        initial_avg_mae = total_error_sum_initial / total_weight if total_weight > 0 else 0.0
        current_avg_mae = initial_avg_mae
        
        steps = int(search_range / step)
        total_steps_per_iter = (2 * steps + 1) ** 2
        
        for iteration in range(max_iter):
            self.dlg.progress_bar.setMaximum(total_steps_per_iter)
            iter_msg = f"반복 {iteration+1}/{max_iter}" if is_iterative else "단일 실행"
            self.dlg.lbl_summary.setText(f"최적 위치 탐색 중... ({iter_msg})")
            
            best_dx_iter, best_dy_iter = 0.0, 0.0
            min_avg_mae_iter = current_avg_mae
            
            current_step = 0
            
            for x in range(-steps, steps + 1):
                dx = x * step
                for y in range(-steps, steps + 1):
                    dy = y * step
                    
                    if x == 0 and y == 0:
                        current_step += 1
                        continue
                    
                    current_step += 1
                    if current_step % 50 == 0:
                        self.dlg.progress_bar.setValue(current_step)
                        QApplication.processEvents()
                    
                    # [수정됨] 전체 쌍에 대해 이동 후 오차 계산
                    total_cost = 0.0
                    for pts, cad_geom, cur_centroid, cad_centroid in test_data:
                        pair_dist_sum = 0.0
                        max_dist = 0.0
                        for pt in pts:
                            t_pt = QgsPointXY(pt.x() + dx, pt.y() + dy)
                            dist = cad_geom.distance(QgsGeometry.fromPointXY(t_pt))
                            pair_dist_sum += dist
                            if dist > max_dist:
                                max_dist = dist
                                
                        avg_dist = pair_dist_sum / len(pts)
                        
                        # 중심점 이동 반영 (미끄러짐 페널티)
                        t_centroid = QgsPointXY(cur_centroid.x() + dx, cur_centroid.y() + dy)
                        centroid_dist = math.sqrt(t_centroid.sqrDist(cad_centroid))
                        
                        # 판정 알고리즘에서 가장 민감하게 반응하는 '가장 멀리 튄 점(max_dist)'의 비중을 높임
                        total_cost += (avg_dist * 0.3) + (max_dist * 0.6) + (centroid_dist * 0.1)
                    
                    # 가중치 지점 오차 합산
                    cp_error_sum = 0.0
                    for src, dst, w in control_points:
                        t_pt = QgsPointXY(src.x() + dx, src.y() + dy)
                        cp_error_sum += math.sqrt(t_pt.sqrDist(dst)) * w
                    
                    avg_mae = (total_cost + cp_error_sum) / total_weight if total_weight > 0 else float('inf')
                    
                    if avg_mae < min_avg_mae_iter:
                        min_avg_mae_iter = avg_mae
                        best_dx_iter = dx
                        best_dy_iter = dy
            
            if best_dx_iter == 0 and best_dy_iter == 0:
                break 
            
            accum_dx += best_dx_iter
            accum_dy += best_dy_iter
            current_avg_mae = min_avg_mae_iter
            
            if is_iterative:
                for i in range(len(test_data)):
                    pts, cad_geom, cur_centroid, cad_centroid = test_data[i]
                    new_pts = [QgsPointXY(p.x() + best_dx_iter, p.y() + best_dy_iter) for p in pts]
                    new_centroid = QgsPointXY(cur_centroid.x() + best_dx_iter, cur_centroid.y() + best_dy_iter)
                    test_data[i] = (new_pts, cad_geom, new_centroid, cad_centroid)
                
                control_points = [(QgsPointXY(src.x() + best_dx_iter, src.y() + best_dy_iter), dst, w) for src, dst, w in control_points]
            else:
                break
        
        best_dx = round(accum_dx, 3)
        best_dy = round(accum_dy, 3)
        final_avg_mae = current_avg_mae

        return best_dx, best_dy, initial_avg_mae, final_avg_mae

    def _show_apply_dialog(self, best_dx, best_dy, initial_avg_mae, final_avg_mae, mode_desc, is_iterative):
        """최적화 결과 표시 및 적용 여부 확인 다이얼로그"""
        self.dlg.progress_bar.setValue(self.dlg.progress_bar.maximum())
        self.dlg.lbl_summary.setText(f"탐색 완료. 최소 평균오차: {final_avg_mae:.3f}m")
        
        mode_str = "반복 최적화" if is_iterative else "단일 최적화"
        
        if best_dx == 0 and best_dy == 0:
            QMessageBox.information(self.dlg, "알림", f"[{mode_str}] 이미 최적 위치입니다.")
            return

        msg = (f"[{mode_str} 결과]\n\n"
               f"모드: {mode_desc}\n"
               f"초기 평균 오차: {initial_avg_mae:.3f} m\n"
               f"최적 평균 오차: {final_avg_mae:.3f} m\n"
               f"개선량: {initial_avg_mae - final_avg_mae:.3f} m\n\n"
               f"총 이동량 X: {best_dx:.3f} m\n"
               f"총 이동량 Y: {best_dy:.3f} m\n\n"
               "이 이동량을 전체 레이어에 적용하시겠습니까?")
               
        ret = QMessageBox.question(self.dlg, "최적화 결과", msg, QMessageBox.Yes | QMessageBox.No)
        
        if ret == QMessageBox.Yes:
            self.apply_shift_vector(best_dx, best_dy)

    def apply_shift_vector(self, dx, dy):
        """벡터(dx, dy)만큼 현형선(Current) 레이어 전체 이동"""
        # [수정] 최적화 계산은 현형선 기준으로 수행되므로, 이동 대상도 현형선이어야 함
        layer = self.dlg.cb_layer_current.currentData()
        if not layer: return
        
        if not layer.isEditable():
            layer.startEditing()
            
        for feat in layer.getFeatures():
            geom = feat.geometry()
            geom.translate(dx, dy)
            layer.changeGeometry(feat.id(), geom)
            
        self.accumulated_dx += dx
        self.accumulated_dy += dy
        self.dlg.sb_shift_x.setValue(self.accumulated_dx)
        self.dlg.sb_shift_y.setValue(self.accumulated_dy)
        
        self.iface.mapCanvas().refresh()
        QMessageBox.information(self.dlg, "적용 완료", "현형선 레이어 이동이 완료되었습니다.")

    def nudge_layer(self, dx_sign, dy_sign):
        # 1. Get Distance & Unit
        dist = self.dlg.sb_nudge_dist.value()
        if dist <= 0:
            return
            
        unit = self.dlg.cb_nudge_unit.currentText()
        factor = 0.01 if unit == "cm" else 1.0
        
        # 2. Calculate Delta
        delta = dist * factor
        dx = delta * dx_sign
        dy = delta * dy_sign
        
        # 3. Apply to Layer
        layer = self.dlg.cb_layer_target.currentData()
        if not layer:
            self.iface.messageBar().pushMessage("경고", "이동 대상 레이어가 선택되지 않았습니다.", level=Qgis.Warning)
            return
            
        if not layer.isEditable():
            if not layer.startEditing():
                self.iface.messageBar().pushMessage("오류", "레이어를 편집 모드로 전환할 수 없습니다.", level=Qgis.Critical)
                return
            
        # Apply to selection if exists, else all
        features = layer.selectedFeatures() if layer.selectedFeatureCount() > 0 else layer.getFeatures()
            
        layer.beginEditCommand("Manual Nudge")
        for feat in features:
            geom = feat.geometry()
            geom.translate(dx, dy)
            layer.changeGeometry(feat.id(), geom)
        layer.endEditCommand()
        
        self.accumulated_dx += dx
        self.accumulated_dy += dy
        
        self.dlg.sb_shift_x.setValue(self.accumulated_dx)
        self.dlg.sb_shift_y.setValue(self.accumulated_dy)
        
        layer.triggerRepaint()
        self.iface.mapCanvas().refresh()
        self.iface.messageBar().pushMessage("이동 완료", f"누적 이동량 -> X: {self.accumulated_dx:.3f}m, Y: {self.accumulated_dy:.3f}m", level=Qgis.Info, duration=2)

    def remove_error_visuals(self):
        if self.error_rubber_band:
            try:
                self.iface.mapCanvas().scene().removeItem(self.error_rubber_band)
            except:
                pass
            self.error_rubber_band = None
            
        for m in self.error_markers:
            try:
                self.iface.mapCanvas().scene().removeItem(m)
            except:
                pass
        self.error_markers.clear()

    def clear_preview_visuals(self):
        """미리보기용으로 생성된 모든 시각적 요소를 제거합니다."""
        for rb, ann in self.preview_visuals:
            if rb:
                try:
                    self.iface.mapCanvas().scene().removeItem(rb)
                except:
                    pass
            if ann:
                try:
                    QgsProject.instance().annotationManager().removeAnnotation(ann)
                except:
                    pass
        self.preview_visuals.clear()
        self.iface.mapCanvas().refresh()

    def visualize_groups(self, groups, layer):
        """지도 위에 그룹들을 다른 색상으로 시각화합니다."""
        visuals = []
        canvas = self.iface.mapCanvas()
        
        colors = [
            QColor(255, 0, 0, 80),    # Red
            QColor(0, 0, 255, 80),    # Blue
            QColor(0, 255, 0, 80),    # Green
            QColor(255, 165, 0, 80),  # Orange
            QColor(128, 0, 128, 80),  # Purple
            QColor(0, 255, 255, 80)   # Cyan
        ]
        
        for i, group_feats in enumerate(groups):
            if not group_feats: continue
            
            group_geom = self._get_group_geometry(group_feats)
            
            if not group_geom or group_geom.isEmpty():
                continue

            # RubberBand
            color = colors[i % len(colors)]
            rb = QgsRubberBand(canvas, QgsWkbTypes.PolygonGeometry)
            rb.setToGeometry(group_geom, layer)
            rb.setColor(color)
            rb.setWidth(2)
            rb.setStrokeColor(QColor(color.red(), color.green(), color.blue(), 255))
            rb.show()
            
            # Annotation
            ann = None
            try:
                ann = QgsTextAnnotation()
                ct = QgsCoordinateTransform(layer.crs(), canvas.mapSettings().destinationCrs(), QgsProject.instance())
                center_pt = ct.transform(group_geom.centroid().asPoint())
                ann.setMapPosition(center_pt)
                doc = QTextDocument()
                doc.setHtml(f"<div style='color:black; font-weight:bold; font-size:15px; background-color:rgba(255,255,255,0.7); padding:2px;'>그룹 {i+1}</div>")
                doc.setPageSize(QSizeF(80, 30))
                ann.setDocument(doc)
                ann.setFrameBackgroundColor(QColor(0, 0, 0, 0))
                QgsProject.instance().annotationManager().addAnnotation(ann)
            except Exception as e:
                print(f"Error creating annotation for preview: {e}")
                pass
                
            visuals.append((rb, ann))
        
        canvas.refresh()
        return visuals

    def _get_group_geometry(self, features):
        """선택된 객체들의 Convex Hull을 계산하여 대표 영역으로 반환"""
        if not features:
            return None

        geoms = []
        for f in features:
            g = f.geometry()
            if g and not g.isEmpty():
                geoms.append(g)
        
        if not geoms:
            return None

        combined_geom = QgsGeometry.collectGeometry(geoms)
        
        if combined_geom.isEmpty():
            return None
            
        return combined_geom.convexHull()

    def show_error_line(self, geometry, crs):
        if self.error_rubber_band:
            try:
                self.iface.mapCanvas().scene().removeItem(self.error_rubber_band)
            except:
                pass
            self.error_rubber_band = None
            
        for m in self.error_markers:
            try:
                self.iface.mapCanvas().scene().removeItem(m)
            except:
                pass
        self.error_markers.clear()

    def show_error_line(self, geometry, crs):
        print(f"DEBUG: show_error_line 호출됨")
        print(f"DEBUG:   -> 입력 Geometry: {geometry.asWkt()}")
        print(f"DEBUG:   -> 입력 CRS: {crs.authid()} (Valid: {crs.isValid()})")
        
        self.remove_error_visuals()
        
        # Transform geometry to Map Canvas CRS for visualization
        dest_crs = self.iface.mapCanvas().mapSettings().destinationCrs()
        print(f"DEBUG:   -> 캔버스 CRS: {dest_crs.authid()}")
        
        geom_canvas = QgsGeometry(geometry)
        
        if crs.isValid() and dest_crs.isValid():
            try:
                ct = QgsCoordinateTransform(crs, dest_crs, QgsProject.instance())
                geom_canvas.transform(ct)
            except Exception as e:
                print(f"DEBUG: 좌표 변환 실패: {e}")
        else:
            print("DEBUG: CRS 유효하지 않음, 변환 건너뜀")
            
        print(f"DEBUG:   -> 변환된 Geometry: {geom_canvas.asWkt()}")
        
        self.error_rubber_band = QgsRubberBand(self.iface.mapCanvas(), QgsWkbTypes.LineGeometry)
        self.error_rubber_band.setToGeometry(geom_canvas, None) # Already transformed
        self.error_rubber_band.setColor(QColor(255, 0, 255)) # Magenta
        self.error_rubber_band.setWidth(2)
        self.error_rubber_band.setLineStyle(Qt.DashLine)
        self.error_rubber_band.show()
        
        # Draw Points for all error vectors
        # Handle MultiLineString
        if geom_canvas.isMultipart():
            lines = geom_canvas.asMultiPolyline()
        else:
            lines = [geom_canvas.asPolyline()]
            
        for line_pts in lines:
            if len(line_pts) >= 2:
                start_pt = QgsPointXY(line_pts[0])
                end_pt = QgsPointXY(line_pts[-1])
                
                # Start Marker (Green X)
                m_start = QgsVertexMarker(self.iface.mapCanvas())
                m_start.setCenter(start_pt)
                m_start.setColor(QColor(0, 255, 0))
                m_start.setIconType(QgsVertexMarker.ICON_X)
                m_start.setIconSize(8)
                m_start.setPenWidth(2)
                self.error_markers.append(m_start)
                
                # End Marker (Red Circle)
                m_end = QgsVertexMarker(self.iface.mapCanvas())
                m_end.setCenter(end_pt)
                m_end.setColor(QColor(255, 0, 0))
                m_end.setIconType(QgsVertexMarker.ICON_CIRCLE)
                m_end.setIconSize(8)
                m_end.setPenWidth(2)
                self.error_markers.append(m_end)

    def highlight_feature(self, feature, layer=None):
        rb = QgsRubberBand(self.iface.mapCanvas(), feature.geometry().type())
        rb.setToGeometry(feature.geometry(), layer)
        rb.setStrokeColor(QColor(255, 0, 0, 255)) # Red outline
        rb.setFillColor(Qt.transparent) # Transparent fill
        rb.setWidth(2)
        self.rubber_bands.append(rb)

    def flash_feature(self, feature, layer=None):
        # Reset previous flash
        self.flash_timer.stop()
        if self.flash_rubber_band:
            self.iface.mapCanvas().scene().removeItem(self.flash_rubber_band)
            self.flash_rubber_band = None
        
        # Create new rubber band for flashing
        self.flash_rubber_band = QgsRubberBand(self.iface.mapCanvas(), feature.geometry().type())
        self.flash_rubber_band.setToGeometry(feature.geometry(), layer)
        self.flash_rubber_band.setColor(QColor(0, 255, 255, 180)) # Cyan with transparency
        self.flash_rubber_band.setWidth(4)
        
        self.flash_counter = 0
        self.flash_timer.start(250) # 250ms interval

    def flash_tick(self):
        if not self.flash_rubber_band:
            self.flash_timer.stop()
            return

        # Blink 3 times (On/Off * 3 = 6 ticks)
        if self.flash_counter >= 6:
            self.flash_timer.stop()
            self.iface.mapCanvas().scene().removeItem(self.flash_rubber_band)
            self.flash_rubber_band = None
            return
            
        if self.flash_counter % 2 == 0:
            self.flash_rubber_band.hide()
        else:
            self.flash_rubber_band.show()
            
        self.flash_counter += 1

    def export_to_excel(self):
        path, _ = QFileDialog.getSaveFileName(self.dlg, "결과 저장", "", "CSV Files (*.csv)")
        if not path:
            return
            
        try:
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                
                # Header
                headers = []
                for col in range(self.dlg.table_results.columnCount()):
                    headers.append(self.dlg.table_results.horizontalHeaderItem(col).text())
                writer.writerow(headers)
                
                # Data
                for row in range(self.dlg.table_results.rowCount()):
                    row_data = []
                    for col in range(self.dlg.table_results.columnCount()):
                        item = self.dlg.table_results.item(row, col)
                        row_data.append(item.text() if item else "")
                    writer.writerow(row_data)

            QMessageBox.information(self.dlg, "성공", f"파일이 저장되었습니다:\n{path}")
        except Exception as e:
            QMessageBox.critical(self.dlg, "오류", f"저장 중 오류 발생:\n{str(e)}")

    def remove_error_visuals(self):
        if self.error_rubber_band:
            try:
                self.iface.mapCanvas().scene().removeItem(self.error_rubber_band)
            except:
                pass
            self.error_rubber_band = None
            
        for m in self.error_markers:
            try:
                self.iface.mapCanvas().scene().removeItem(m)
            except:
                pass
        self.error_markers.clear()

    def show_error_line(self, geometry, crs):
        print(f"DEBUG: show_error_line 호출됨")
        print(f"DEBUG:   -> 입력 Geometry: {geometry.asWkt()}")
        print(f"DEBUG:   -> 입력 CRS: {crs.authid()} (Valid: {crs.isValid()})")
        
        self.remove_error_visuals()
        
        # Transform geometry to Map Canvas CRS for visualization
        dest_crs = self.iface.mapCanvas().mapSettings().destinationCrs()
        print(f"DEBUG:   -> 캔버스 CRS: {dest_crs.authid()}")
        
        geom_canvas = QgsGeometry(geometry)
        
        if crs.isValid() and dest_crs.isValid():
            try:
                ct = QgsCoordinateTransform(crs, dest_crs, QgsProject.instance())
                geom_canvas.transform(ct)
            except Exception as e:
                print(f"DEBUG: 좌표 변환 실패: {e}")
        else:
            print("DEBUG: CRS 유효하지 않음, 변환 건너뜀")
            
        print(f"DEBUG:   -> 변환된 Geometry: {geom_canvas.asWkt()}")
        
        self.error_rubber_band = QgsRubberBand(self.iface.mapCanvas(), QgsWkbTypes.LineGeometry)
        self.error_rubber_band.setToGeometry(geom_canvas, None) # Already transformed
        self.error_rubber_band.setColor(QColor(255, 0, 255)) # Magenta
        self.error_rubber_band.setWidth(2)
        self.error_rubber_band.setLineStyle(Qt.DashLine)
        self.error_rubber_band.show()
        
        # Draw Points for all error vectors
        # Handle MultiLineString
        if geom_canvas.isMultipart():
            lines = geom_canvas.asMultiPolyline()
        else:
            lines = [geom_canvas.asPolyline()]
            
        for line_pts in lines:
            if len(line_pts) >= 2:
                start_pt = QgsPointXY(line_pts[0])
                end_pt = QgsPointXY(line_pts[-1])
                
                # Start Marker (Green X)
                m_start = QgsVertexMarker(self.iface.mapCanvas())
                m_start.setCenter(start_pt)
                m_start.setColor(QColor(0, 255, 0))
                m_start.setIconType(QgsVertexMarker.ICON_X)
                m_start.setIconSize(8)
                m_start.setPenWidth(2)
                self.error_markers.append(m_start)
                
                # End Marker (Red Circle)
                m_end = QgsVertexMarker(self.iface.mapCanvas())
                m_end.setCenter(end_pt)
                m_end.setColor(QColor(255, 0, 0))
                m_end.setIconType(QgsVertexMarker.ICON_CIRCLE)
                m_end.setIconSize(8)
                m_end.setPenWidth(2)
                self.error_markers.append(m_end)

    def highlight_feature(self, feature, layer=None):
        rb = QgsRubberBand(self.iface.mapCanvas(), feature.geometry().type())
        rb.setToGeometry(feature.geometry(), layer)
        rb.setStrokeColor(QColor(255, 0, 0, 255)) # Red outline
        rb.setFillColor(Qt.transparent) # Transparent fill
        rb.setWidth(2)
        self.rubber_bands.append(rb)

    def flash_feature(self, feature, layer=None):
        # Reset previous flash
        self.flash_timer.stop()
        if self.flash_rubber_band:
            self.iface.mapCanvas().scene().removeItem(self.flash_rubber_band)
            self.flash_rubber_band = None
        
        # Create new rubber band for flashing
        self.flash_rubber_band = QgsRubberBand(self.iface.mapCanvas(), feature.geometry().type())
        self.flash_rubber_band.setToGeometry(feature.geometry(), layer)
        self.flash_rubber_band.setColor(QColor(0, 255, 255, 180)) # Cyan with transparency
        self.flash_rubber_band.setWidth(4)
        
        self.flash_counter = 0
        self.flash_timer.start(250) # 250ms interval

    def flash_tick(self):
        if not self.flash_rubber_band:
            self.flash_timer.stop()
            return

        # Blink 3 times (On/Off * 3 = 6 ticks)
        if self.flash_counter >= 6:
            self.flash_timer.stop()
            self.iface.mapCanvas().scene().removeItem(self.flash_rubber_band)
            self.flash_rubber_band = None
            return
            
        if self.flash_counter % 2 == 0:
            self.flash_rubber_band.hide()
        else:
            self.flash_rubber_band.show()
            
        self.flash_counter += 1

    def export_to_excel(self):
        path, _ = QFileDialog.getSaveFileName(self.dlg, "결과 저장", "", "CSV Files (*.csv)")
        if not path:
            return
            
        try:
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                
                # Header
                headers = []
                for col in range(self.dlg.table_results.columnCount()):
                    headers.append(self.dlg.table_results.horizontalHeaderItem(col).text())
                writer.writerow(headers)
                
                # Data
                for row in range(self.dlg.table_results.rowCount()):
                    row_data = []
                    for col in range(self.dlg.table_results.columnCount()):
                        item = self.dlg.table_results.item(row, col)
                        row_data.append(item.text() if item else "")
                    writer.writerow(row_data)

            QMessageBox.information(self.dlg, "성공", f"파일이 저장되었습니다:\n{path}")
        except Exception as e:
            QMessageBox.critical(self.dlg, "오류", f"저장 중 오류 발생:\n{str(e)}")