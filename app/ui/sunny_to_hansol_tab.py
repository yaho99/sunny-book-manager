from PyQt5.QtCore import QDate
from PyQt5.QtWidgets import *

from app.logic.sunny_to_hansol import convert_sunny_to_hansol


class OutputTab(QWidget):
    def __init__(self):
        super().__init__()

        self.input_file = ""
        self.output_file = ""
        self.start_date = QDate.currentDate()

        self.setup_ui()

    def setup_ui(self):
        # Input file selection
        self.input_file_guide_label = QLabel("최고의 복사지 장부 엑셀 파일을 선택해주세요")
        self.input_file_label = QLabel("선택된 파일:")
        self.select_input_button = QPushButton("최고의 복사지 장부 파일 선택")

        # Output file selection
        self.output_file_guide_label = QLabel("한솔 장부 엑셀 파일을 선택해주세요")
        self.output_file_label = QLabel("선택된 파일:")
        self.select_output_button = QPushButton("한솔 장부 파일 선택")

        # Start date selection
        self.start_date_guide_label = QLabel("원하는 장부 변환 시작 일자를 선택해주세요")
        self.start_date_label = QLabel(f"선택된 날짜: {self.start_date.toString('yyyy-MM-dd')}")
        self.select_start_date_calendar = QCalendarWidget()

        # Convert button
        self.button_convert = QPushButton("변환 시작")

        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.input_file_guide_label)
        layout.addWidget(self.input_file_label)
        layout.addWidget(self.select_input_button)
        layout.addWidget(self.output_file_guide_label)
        layout.addWidget(self.output_file_label)
        layout.addWidget(self.select_output_button)
        layout.addWidget(self.start_date_guide_label)
        layout.addWidget(self.start_date_label)
        layout.addWidget(self.select_start_date_calendar)
        layout.addWidget(self.button_convert)
        self.setLayout(layout)

        # Connect signals
        self.select_input_button.clicked.connect(self.select_input_file)
        self.select_output_button.clicked.connect(self.select_output_file)
        self.select_start_date_calendar.selectionChanged.connect(self.select_start_date)
        self.button_convert.clicked.connect(self.convert_file)

    def select_input_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "엑셀 파일 선택", "", "Excel Files (*.xlsx)")
        if file_path:
            if not file_path.endswith('.xlsx'):
                QMessageBox.warning(self, "경고", "엑셀(.xlsx) 파일만 선택 가능합니다.")
                return
            self.input_file = file_path
            self.input_file_label.setText(f"선택된 파일: {file_path}")

    def select_output_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "엑셀 파일 선택", "", "Excel Files (*.xlsx)")
        if file_path:
            if not file_path.endswith('.xlsx'):
                QMessageBox.warning(self, "경고", "엑셀(.xlsx) 파일만 선택 가능합니다.")
                return
            self.output_file = file_path
            self.output_file_label.setText(f"선택된 파일: {file_path}")

    def select_start_date(self):
        selected_date = self.select_start_date_calendar.selectedDate()
        if selected_date.isValid():
            self.start_date = selected_date
            self.start_date_label.setText(f"선택된 날짜: {self.start_date.toString('yyyy-MM-dd')}")

    def convert_file(self):
        if not self.input_file:
            QMessageBox.warning(self, "경고", "최고의 복사지 장부 파일을 선택하세요.")
            return

        if not self.output_file:
            QMessageBox.warning(self, "경고", "한솔 장부 파일을 선택하세요.")
            return

        try:
            output_path = convert_sunny_to_hansol(self.input_file, self.output_file, self.start_date)
            QMessageBox.information(self, "완료", f"변환 완료!\n저장 위치: {output_path}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"변환 중 오류 발생:\n{str(e)}")
