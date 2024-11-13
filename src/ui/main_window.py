from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QTabWidget
from src.ui.tabs.consignment_tab import ConsignmentTab
from src.ui.tabs.wholesale_tab import WholesaleTab

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('거래처별 발주서 추출 프로그램')
        self.setGeometry(300, 300, 500, 300)
        
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        
        tab_widget = QTabWidget()
        tab_widget.addTab(ConsignmentTab(), "위탁 발주서")
        tab_widget.addTab(WholesaleTab(), "사입 발주서")
        
        main_layout.addWidget(tab_widget)
        main_widget.setLayout(main_layout)