from PyQt5.QtWidgets import QWidget, QPushButton, QLabel, QVBoxLayout, QFileDialog, QMessageBox
from app.logic.excel_converter import convert_excel


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("엑셀 변환기")
        self.setFixedSize(400, 200)

        self.label = QLabel("엑셀 파일을 선택해주세요")
        self.button_select = QPushButton("파일 선택")
        self.button_convert = QPushButton("변환 시작")

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.button_select)
        layout.addWidget(self.button_convert)
        self.setLayout(layout)

        self.input_file = ""

        self.button_select.clicked.connect(self.select_file)
        self.button_convert.clicked.connect(self.convert_file)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "엑셀 파일 선택", "", "Excel Files (*.xlsx)")
        if file_path:
            self.input_file = file_path
            self.label.setText(f"선택된 파일: {file_path}")

    def convert_file(self):
        if not self.input_file:
            QMessageBox.warning(self, "경고", "먼저 파일을 선택하세요.")
            return

        try:
            output_path = convert_excel(self.input_file)
            QMessageBox.information(self, "완료", f"변환 완료!\n저장 위치: {output_path}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"변환 중 오류 발생:\n{str(e)}")
