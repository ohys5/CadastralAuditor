from PyQt5 import QtWidgets, QtCore
from qgis.core import QgsProject, QgsMapLayer

class CadastralAuditorDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(CadastralAuditorDialog, self).__init__(parent)
        self.setWindowTitle("Cadastral Auditor")
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
        self.resize(560, 350)
        
        layout = QtWidgets.QVBoxLayout(self)
        
        # --- Input Selection ---
        input_group = QtWidgets.QGroupBox("레이어 선택")
        input_layout = QtWidgets.QGridLayout()
        
        input_layout.addWidget(QtWidgets.QLabel("지적도 레이어 (Cadastral):"), 0, 0)
        self.cb_layer_cadastral = QtWidgets.QComboBox()
        input_layout.addWidget(self.cb_layer_cadastral, 0, 1)
        
        input_layout.addWidget(QtWidgets.QLabel("현형선 레이어 (Current):"), 1, 0)
        self.cb_layer_current = QtWidgets.QComboBox()
        input_layout.addWidget(self.cb_layer_current, 1, 1)
        
        input_layout.addWidget(QtWidgets.QLabel("이동 대상 레이어 (Target):"), 2, 0)
        self.cb_layer_target = QtWidgets.QComboBox()
        input_layout.addWidget(self.cb_layer_target, 2, 1)
        
        input_layout.addWidget(QtWidgets.QLabel("양호 범위 (Min, m):"), 3, 0)
        self.sb_tol_min = QtWidgets.QDoubleSpinBox()
        self.sb_tol_min.setDecimals(3)
        self.sb_tol_min.setRange(0.001, 10.0)
        self.sb_tol_min.setValue(0.10)
        self.sb_tol_min.setSingleStep(0.01)
        input_layout.addWidget(self.sb_tol_min, 3, 1)
        
        input_layout.addWidget(QtWidgets.QLabel("한계 범위 (Max, m):"), 4, 0)
        self.sb_tol_max = QtWidgets.QDoubleSpinBox()
        self.sb_tol_max.setDecimals(3)
        self.sb_tol_max.setRange(0.001, 10.0)
        self.sb_tol_max.setValue(0.30)
        self.sb_tol_max.setSingleStep(0.01)
        input_layout.addWidget(self.sb_tol_max, 4, 1)
        
        input_layout.addWidget(QtWidgets.QLabel("상대거리 제외범위 (Exclusion, m):"), 5, 0)
        self.sb_exclusion_limit = QtWidgets.QDoubleSpinBox()
        self.sb_exclusion_limit.setDecimals(3)
        self.sb_exclusion_limit.setRange(0.001, 100.0)
        self.sb_exclusion_limit.setValue(2.00)
        self.sb_exclusion_limit.setSingleStep(0.1)
        input_layout.addWidget(self.sb_exclusion_limit, 5, 1)
        
        self.sb_tol_min.valueChanged.connect(self.validate_tolerance)
        self.sb_tol_max.valueChanged.connect(self.validate_tolerance)
        
        self.chk_fixed_scale = QtWidgets.QCheckBox("이동 시 현재 화면 배율 유지")
        input_layout.addWidget(self.chk_fixed_scale, 6, 0, 1, 2)
        
        input_group.setLayout(input_layout)
        layout.addWidget(input_group)
        
        # --- Manual Shift Adjustment ---
        shift_group = QtWidgets.QGroupBox("이동량 확인 (Shift Amount)")
        shift_layout = QtWidgets.QHBoxLayout()
        
        shift_layout.addWidget(QtWidgets.QLabel("X (m):"))
        self.sb_shift_x = QtWidgets.QDoubleSpinBox()
        self.sb_shift_x.setRange(-1000.0, 1000.0)
        self.sb_shift_x.setDecimals(3)
        self.sb_shift_x.setReadOnly(True)
        self.sb_shift_x.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
        shift_layout.addWidget(self.sb_shift_x)
        
        shift_layout.addWidget(QtWidgets.QLabel("Y (m):"))
        self.sb_shift_y = QtWidgets.QDoubleSpinBox()
        self.sb_shift_y.setRange(-1000.0, 1000.0)
        self.sb_shift_y.setDecimals(3)
        self.sb_shift_y.setReadOnly(True)
        self.sb_shift_y.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
        shift_layout.addWidget(self.sb_shift_y)
        
        shift_group.setLayout(shift_layout)
        layout.addWidget(shift_group)

        # --- Manual Nudge (수동 미세 조정) ---
        nudge_group = QtWidgets.QGroupBox("수동 미세 조정 (Manual Nudge)")
        nudge_layout = QtWidgets.QGridLayout()
        
        nudge_layout.addWidget(QtWidgets.QLabel("이동 거리:"), 0, 0)
        self.sb_nudge_dist = QtWidgets.QDoubleSpinBox()
        self.sb_nudge_dist.setRange(0.0, 1000.0)
        self.sb_nudge_dist.setValue(1.0)
        nudge_layout.addWidget(self.sb_nudge_dist, 0, 1)
        
        self.cb_nudge_unit = QtWidgets.QComboBox()
        self.cb_nudge_unit.addItems(["m", "cm"])
        nudge_layout.addWidget(self.cb_nudge_unit, 0, 2)
        
        # Direction Buttons (3x3 Grid)
        self.btn_ul = QtWidgets.QPushButton("◤")
        self.btn_up = QtWidgets.QPushButton("▲")
        self.btn_ur = QtWidgets.QPushButton("◥")
        self.btn_left = QtWidgets.QPushButton("◀")
        self.btn_right = QtWidgets.QPushButton("▶")
        self.btn_dl = QtWidgets.QPushButton("◣")
        self.btn_down = QtWidgets.QPushButton("▼")
        self.btn_dr = QtWidgets.QPushButton("◢")
        
        for i, btn in enumerate([self.btn_ul, self.btn_up, self.btn_ur, self.btn_left, self.btn_right, self.btn_dl, self.btn_down, self.btn_dr]):
            r, c = [(1,0), (1,1), (1,2), (2,0), (2,2), (3,0), (3,1), (3,2)][i]
            nudge_layout.addWidget(btn, r, c)
            
        nudge_group.setLayout(nudge_layout)
        layout.addWidget(nudge_group)

        # --- [추가] 자동 위치 보정 (Auto-Correction) ---
        auto_group = QtWidgets.QGroupBox("자동 위치 보정 (Auto-Correction)")
        auto_layout = QtWidgets.QGridLayout()
        
        auto_layout.addWidget(QtWidgets.QLabel("탐색 반경(m):"), 0, 0)
        self.sb_search_range = QtWidgets.QDoubleSpinBox()
        self.sb_search_range.setValue(2.0)
        self.sb_search_range.setSingleStep(0.5)
        auto_layout.addWidget(self.sb_search_range, 0, 1)
        
        auto_layout.addWidget(QtWidgets.QLabel("간격(m):"), 0, 2)
        self.sb_search_step = QtWidgets.QDoubleSpinBox()
        self.sb_search_step.setValue(0.1)
        self.sb_search_step.setSingleStep(0.1)
        auto_layout.addWidget(self.sb_search_step, 0, 3)

        auto_layout.addWidget(QtWidgets.QLabel("그룹 거리(m):"), 0, 4)
        self.sb_cluster_dist = QtWidgets.QDoubleSpinBox()
        self.sb_cluster_dist.setRange(0.0, 500.0)
        self.sb_cluster_dist.setValue(15.0)
        self.sb_cluster_dist.setSingleStep(1.0)
        self.sb_cluster_dist.setToolTip("선택된 객체들이 이 거리 이상 떨어져 있으면 다른 그룹으로 분리합니다.")
        auto_layout.addWidget(self.sb_cluster_dist, 0, 5)

        self.chk_iterative = QtWidgets.QCheckBox("반복 최적화 (수렴할 때까지 실행)")
        auto_layout.addWidget(self.chk_iterative, 1, 0, 1, 6)
        
        btn_select_layout = QtWidgets.QHBoxLayout()
        self.btn_select_area = QtWidgets.QPushButton("영역 선택 (다각형)")
        self.btn_select_multi = QtWidgets.QPushButton("다중 영역 선택 (추가)")
        btn_select_layout.addWidget(self.btn_select_area)
        btn_select_layout.addWidget(self.btn_select_multi)
        auto_layout.addLayout(btn_select_layout, 2, 0, 1, 4)
        
        self.btn_auto_calc = QtWidgets.QPushButton("최적 이동량 계산")
        auto_layout.addWidget(self.btn_auto_calc, 2, 4, 1, 2)
        
        auto_group.setLayout(auto_layout)
        layout.addWidget(auto_group)

        # --- [추가] 가중치 지점 (Control Points) ---
        self.cp_group = QtWidgets.QGroupBox("가중치 지점 (Control Points)")
        cp_layout = QtWidgets.QVBoxLayout()
        
        self.tbl_control_points = QtWidgets.QTableWidget()
        self.tbl_control_points.setColumnCount(5)
        self.tbl_control_points.setHorizontalHeaderLabels(["Cur X", "Cur Y", "Tar X", "Tar Y", "Weight"])
        self.tbl_control_points.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tbl_control_points.setFixedHeight(70)
        cp_layout.addWidget(self.tbl_control_points)
        
        cp_btn_layout = QtWidgets.QHBoxLayout()
        
        cp_btn_layout.addWidget(QtWidgets.QLabel("가중치:"))
        self.sb_cp_weight = QtWidgets.QDoubleSpinBox()
        self.sb_cp_weight.setRange(0.1, 1000.0)
        self.sb_cp_weight.setValue(50.0)
        self.sb_cp_weight.setSingleStep(10.0)
        cp_btn_layout.addWidget(self.sb_cp_weight)
        
        self.btn_add_cp = QtWidgets.QPushButton("지점 추가 (지도 선택)")
        self.btn_clear_cp = QtWidgets.QPushButton("목록 초기화")
        cp_btn_layout.addWidget(self.btn_add_cp)
        cp_btn_layout.addWidget(self.btn_clear_cp)
        cp_layout.addLayout(cp_btn_layout)
        
        self.cp_group.setLayout(cp_layout)
        layout.addWidget(self.cp_group)

        # --- [추가] 분석 알고리즘 선택 ---
        self.gb_mode = QtWidgets.QGroupBox("판정 알고리즘 선택")
        self.mode_layout = QtWidgets.QHBoxLayout()
        
        self.rb_mode_distance = QtWidgets.QRadioButton("거리 기반 (단순/직관적)")
        self.rb_mode_original = QtWidgets.QRadioButton("면적/점수 기반 (기존 방식)")
        
        self.rb_mode_distance.setChecked(True)  # 기본값 설정
        
        self.mode_layout.addWidget(self.rb_mode_distance)
        self.mode_layout.addWidget(self.rb_mode_original)
        self.gb_mode.setLayout(self.mode_layout)
        layout.addWidget(self.gb_mode)

        # --- Action Buttons ---
        btn_layout = QtWidgets.QHBoxLayout()
        self.btn_analyze = QtWidgets.QPushButton("분석 실행 (Analyze)")
        self.btn_apply_shift = QtWidgets.QPushButton("전체 레이어 이동 (Shift Layer)")
        self.btn_apply_shift.setEnabled(False)
        self.btn_export = QtWidgets.QPushButton("CSV 저장 (Export)")
        self.btn_export.setEnabled(False)
        
        btn_layout.addWidget(self.btn_analyze)
        btn_layout.addWidget(self.btn_apply_shift)
        btn_layout.addWidget(self.btn_export)
        layout.addLayout(btn_layout)
        
        # --- Results Table ---
        self.table_results = QtWidgets.QTableWidget()
        self.table_results.setColumnCount(11)
        self.table_results.setHorizontalHeaderLabels([
            "연번", "매칭ID", "상대 거리(m)", "각도(deg)", "위상", "점수", "판정", "ND Cost", "dx(m)", "dy(m)", "비고"
        ])
        self.table_results.setSortingEnabled(True)
        layout.addWidget(self.table_results)
        
        # --- Progress Bar ---
        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        
        # --- Summary Label ---
        self.lbl_summary = QtWidgets.QLabel("분석 결과 요약: 대기 중")
        self.lbl_summary.setAlignment(QtCore.Qt.AlignCenter)
        layout.addWidget(self.lbl_summary)
        
        # --- Close Button ---
        self.button_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)
        
        # Initialize layer combo boxes
        self.populate_layers()

    def populate_layers(self):
        layers = QgsProject.instance().mapLayers().values()
        for layer in layers:
            if layer.type() == QgsMapLayer.VectorLayer:
                self.cb_layer_cadastral.addItem(layer.name(), layer)
                self.cb_layer_current.addItem(layer.name(), layer)
                self.cb_layer_target.addItem(layer.name(), layer)

    def validate_tolerance(self):
        min_val = self.sb_tol_min.value()
        max_val = self.sb_tol_max.value()
        if min_val >= max_val:
            if self.sender() == self.sb_tol_min:
                self.sb_tol_max.setValue(min_val + 0.01)
            else:
                self.sb_tol_min.setValue(max_val - 0.01)

    def clear_table(self):
        self.table_results.setRowCount(0)
        self.btn_export.setEnabled(False)
        self.btn_apply_shift.setEnabled(False)