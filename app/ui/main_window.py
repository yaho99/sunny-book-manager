from PyQt5.QtWidgets import *

from app.ui.smart_store_to_sunny_tab import InputTab
from app.ui.sunny_to_hansol_tab import OutputTab


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("최고의 장부")
        self.setFixedSize(800, 600)

        self.setup_ui()

    def setup_ui(self):
        # Create tab widget
        self.tab_widget = QTabWidget()

        # Create tabs
        self.input_tab = InputTab()
        self.output_tab = OutputTab()

        # Add tabs to widget
        self.tab_widget.addTab(self.input_tab, "스마트스토어 → 최고의 복사지")
        self.tab_widget.addTab(self.output_tab, "최고의 복사지 → 한솔")

        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.tab_widget)
        self.setLayout(layout)
