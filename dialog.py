import os
from PyQt5 import QtWidgets, QtCore
from qgis.core import QgsProject

class QFieldAutoSetupDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(QFieldAutoSetupDialog, self).__init__(parent)
        self.setWindowTitle("QField 자동 설정 (v2.0)")
        self.setMinimumWidth(450)
        
        layout = QtWidgets.QVBoxLayout(self)
        
        # --- 모드 선택 ---
        self.mode_group = QtWidgets.QGroupBox("작업 모드")
        mode_layout = QtWidgets.QVBoxLayout()
        self.radio_new = QtWidgets.QRadioButton("새 프로젝트 생성")
        self.radio_existing = QtWidgets.QRadioButton("기존 프로젝트에 추가")
        self.radio_new.setChecked(True)
        mode_layout.addWidget(self.radio_new)
        mode_layout.addWidget(self.radio_existing)
        self.mode_group.setLayout(mode_layout)
        layout.addWidget(self.mode_group)

        # --- 기존 프로젝트 선택 (기본 비활성) ---
        self.existing_group = QtWidgets.QGroupBox("기존 프로젝트 파일 (.qgs / .qgz)")
        self.existing_group.setEnabled(False)
        existing_layout = QtWidgets.QHBoxLayout()
        self.existing_path = QtWidgets.QLineEdit()
        self.btn_browse_qgs = QtWidgets.QPushButton("찾기")
        self.btn_browse_qgs.clicked.connect(self.browse_qgs)
        existing_layout.addWidget(self.existing_path)
        existing_layout.addWidget(self.btn_browse_qgs)
        self.existing_group.setLayout(existing_layout)
        layout.addWidget(self.existing_group)
        
        # --- 공통 설정 ---
        self.common_group = QtWidgets.QGroupBox("기본 설정")
        common_layout = QtWidgets.QGridLayout()
        
        common_layout.addWidget(QtWidgets.QLabel("조사 프로젝트 이름:"), 0, 0)
        self.project_name = QtWidgets.QLineEdit()
        self.project_name.setText("현장조사_프로젝트")
        common_layout.addWidget(self.project_name, 0, 1)
        
        common_layout.addWidget(QtWidgets.QLabel("데이터 저장 폴더:"), 1, 0)
        self.save_path = QtWidgets.QLineEdit()
        self.save_path.setText(os.path.expanduser("~"))
        self.btn_browse_folder = QtWidgets.QPushButton("...")
        self.btn_browse_folder.clicked.connect(self.browse_folder)
        common_layout.addWidget(self.save_path, 1, 1)
        common_layout.addWidget(self.btn_browse_folder, 1, 2)
        
        common_layout.addWidget(QtWidgets.QLabel("좌표계 (CRS):"), 2, 0)
        self.crs_input = QtWidgets.QLineEdit()
        self.crs_input.setText("EPSG:5186")
        common_layout.addWidget(self.crs_input, 2, 1)
        
        self.common_group.setLayout(common_layout)
        layout.addWidget(self.common_group)
        
        # 옵션
        self.photo_checkbox = QtWidgets.QCheckBox("사진 촬영 필드 포함")
        self.photo_checkbox.setChecked(True)
        layout.addWidget(self.photo_checkbox)
        
        # 시그널 연결
        self.radio_existing.toggled.connect(self.toggle_mode)
        
        # 버튼 박스
        self.button_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def toggle_mode(self, checked):
        self.existing_group.setEnabled(checked)
        # 기존 프로젝트 모드일 때는 프로젝트 이름 자동 설정 시도
        if checked and self.existing_path.text():
            self.update_name_from_path(self.existing_path.text())

    def browse_qgs(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "QGIS 프로젝트 선택", "", "QGIS 프로젝트 (*.qgs *.qgz)"
        )
        if file_path:
            self.existing_path.setText(file_path)
            self.update_name_from_path(file_path)
            # 저장 폴더도 프로젝트 위치로 기본 제안
            self.save_path.setText(os.path.dirname(file_path))

    def update_name_from_path(self, path):
        base = os.path.basename(path)
        name = os.path.splitext(base)[0]
        self.project_name.setText(name)

    def browse_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "저장 폴더 선택")
        if folder:
            self.save_path.setText(folder)
            
    def get_data(self):
        return {
            "is_new": self.radio_new.isChecked(),
            "existing_project": self.existing_path.text(),
            "name": self.project_name.text(),
            "path": self.save_path.text(),
            "crs": self.crs_input.text(),
            "photo": self.photo_checkbox.isChecked()
        }
