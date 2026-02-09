from PyQt5 import QtWidgets, QtCore
from qgis.core import QgsProject, QgsMapLayer

class CadastralAuditorDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(CadastralAuditorDialog, self).__init__(parent)
        self.setWindowTitle("Cadastral Auditor")
        self.resize(800, 500)
        
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
        
        input_layout.addWidget(QtWidgets.QLabel("제외 범위 (Exclusion, m):"), 5, 0)
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
        self.table_results.setColumnCount(10)
        self.table_results.setHorizontalHeaderLabels([
            "연번", "매칭ID", "상대 거리(m)", "위상", "점수", "판정", "ND Cost", "dx(m)", "dy(m)", "비고"
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