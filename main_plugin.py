import os
import csv
import math
from PyQt5.QtWidgets import QAction, QMessageBox, QFileDialog, QTableWidgetItem
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QColor
from qgis.core import (
    QgsProject, QgsMapLayer, QgsWkbTypes, QgsSpatialIndex, 
    QgsFeatureRequest, QgsGeometry, Qgis, QgsPointXY, QgsCoordinateTransform
)
from qgis.gui import QgsRubberBand, QgsVertexMarker
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

class NumericTableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        try:
            return float(self.text()) < float(other.text())
        except ValueError:
            return super().__lt__(other)

class CadastralAuditor:
    def __init__(self, iface):
        self.iface = iface
        self.dlg = None
        self.action = None
        self.rubber_bands = []
        
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

    def run(self):
        if not self.dlg:
            self.dlg = CadastralAuditorDialog()
            self.dlg.btn_analyze.clicked.connect(self.run_analysis)
            self.dlg.btn_apply_shift.clicked.connect(self.apply_feature_shift)
            self.dlg.btn_export.clicked.connect(self.export_to_excel)
            self.dlg.table_results.cellClicked.connect(self.zoom_to_feature)
            
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
        self.dlg.exec_()

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
        self.remove_error_visuals()

    def run_analysis(self):
        # 1. Setup
        print("DEBUG: === 분석 시작 ===")
        self.dlg.clear_table()
        self.dlg.table_results.setSortingEnabled(False) # Disable sorting during insertion
        self.dlg.lbl_summary.setText("분석 중...")
        self.clear_highlights()
        self.accumulated_dx = 0.0
        self.accumulated_dy = 0.0
        self.dlg.sb_shift_x.setValue(0.0)
        self.dlg.sb_shift_y.setValue(0.0)
        self.error_lines = {}
        
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
            search_rect = geom_curr.boundingBox()
            search_rect.grow(exclusion_limit)
            candidate_ids = index.intersects(search_rect)
            # print(f"DEBUG:   -> 검색 반경 내 후보 객체 수: {len(candidate_ids)}")
            
            # [Task 2] 최적 후보 선정 (Best Fit Selection)
            best_match_feat = None
            min_dist = float('inf')
            
            # 후보군(candidate_ids)에 대해 필터링 수행
            for fid in candidate_ids:
                cad_feat = cad_layer.getFeature(fid)
                cad_geom = cad_feat.geometry()
                
                # 거리 계산 (이전 방식 참조: Hausdorff Distance 사용)
                # 1. Min Distance for exclusion filter (Fast check)
                d_min = geom_curr.distance(cad_geom)
                
                # 필터 A: 거리 제한 (Step 1) - 최단 거리 기준
                if d_min > exclusion_limit:
                    continue
                
                # 필터 B: 각도 제한 (Step 2) - 45도 이상이면 탈락
                angle_diff = get_angle_diff(geom_curr, cad_geom)
                if angle_diff > 45:
                    continue
                
                # 3. Average Distance (MAE) for ranking
                # Use PointToLineAuditor to calculate average distance of analysis points
                p2l = PointToLineAuditor(cad_feat, cur_feat_tr, densify_distance=1.0)
                p2l_res = p2l.process()
                dist = p2l_res["mae"]
                
                # 최솟값 갱신 (Step 3)
                if dist < min_dist:
                    min_dist = dist
                    best_match_feat = cad_feat
            
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
                
                # Update Statistics
                if min_dist <= tol_min:
                    stats["pass"] += 1
                elif "위치정정" in status:
                    stats["shift"] += 1
                elif "형상" in status:
                    stats["shape"] += 1
                else:
                    stats["critical"] += 1
                
                # 시각화: 불부합일 경우 하이라이트
                if status == "불부합":
                    self.highlight_feature(cur_feat, cur_layer) # Highlight original feature on map

                # [시각화] 비교 점 및 오차 벡터 생성 (PointToLineAuditor 활용)
                # 사용자가 지적선과 현형선의 비교 대상 점을 시각적으로 확인할 수 있도록 함
                p2l = PointToLineAuditor(best_match_feat, cur_feat_tr)
                p2l_res = p2l.process()
                error_geom = p2l_res["error_vectors"]
                if error_geom and not error_geom.isEmpty():
                    self.error_lines[cur_feat.id()] = (error_geom, cad_layer.crs())

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
                self.dlg.table_results.setItem(row, 2, NumericTableWidgetItem(f"{min_dist:.3f}"))
                self.dlg.table_results.setItem(row, 3, QTableWidgetItem(topology))
                self.dlg.table_results.setItem(row, 4, NumericTableWidgetItem(f"{score:.3f}"))
                self.dlg.table_results.setItem(row, 5, QTableWidgetItem(status))
                self.dlg.table_results.setItem(row, 6, NumericTableWidgetItem(f"{nd_cost:.3f}"))
                self.dlg.table_results.setItem(row, 7, NumericTableWidgetItem(f"{shift_x:.3f}"))
                self.dlg.table_results.setItem(row, 8, NumericTableWidgetItem(f"{shift_y:.3f}"))
                
                item_status = self.dlg.table_results.item(row, 5)
                if "불부합" in status:
                    item_status.setForeground(QColor("red"))
                elif "부합" in status:
                    item_status.setForeground(QColor("green"))
                elif "위치정정" in status:
                    item_status.setForeground(QColor("blue"))
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

    def zoom_to_feature(self, row, column):
        item = self.dlg.table_results.item(row, 0)
        if not item:
            return
            
        fid = item.data(Qt.UserRole)
        layer = self.dlg.cb_layer_current.currentData()
        
        if layer and fid is not None:
            feat = layer.getFeature(fid)
            if feat.isValid():
                # Transform coordinates to Map Canvas CRS for Zoom
                canvas = self.iface.mapCanvas()
                dest_crs = canvas.mapSettings().destinationCrs()
                ct = QgsCoordinateTransform(layer.crs(), dest_crs, QgsProject.instance())
                
                geom = feat.geometry()
                
                if self.dlg.chk_fixed_scale.isChecked():
                    center = geom.centroid().asPoint()
                    center_tr = ct.transform(center)
                    canvas.setCenter(center_tr)
                else:
                    bbox = geom.boundingBox()
                    bbox.grow(2.0) # Buffer in layer units (e.g., meters)
                    bbox_geom = QgsGeometry.fromRect(bbox)
                    bbox_geom.transform(ct)
                    canvas.setExtent(bbox_geom.boundingBox())
                canvas.refresh()
                self.flash_feature(feat, layer)
                
                # Clear previous error visuals
                self.remove_error_visuals()
                
                if fid in self.error_lines:
                    geom, crs = self.error_lines[fid]
                    self.show_error_line(geom, crs)

    def apply_feature_shift(self):
        row = self.dlg.table_results.currentRow()
        if row < 0:
            QMessageBox.warning(self.dlg, "경고", "기준이 될 객체를 목록에서 선택해주세요.")
            return
        
        # Get shift values from table directly (Delta), not from spinboxes (Accumulated)
        try:
            dx = float(self.dlg.table_results.item(row, 7).text())
            dy = float(self.dlg.table_results.item(row, 8).text())
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
        rb = QgsRubberBand(self.iface.mapCanvas(), QgsWkbTypes.PolygonGeometry)
        rb.setToGeometry(feature.geometry(), layer)
        rb.setColor(QColor(255, 0, 0, 100)) # Red with transparency
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
        self.flash_rubber_band.setColor(QColor(255, 255, 0, 180)) # Yellow with transparency
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