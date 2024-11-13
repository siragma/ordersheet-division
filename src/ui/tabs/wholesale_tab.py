import os
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QLabel, QProgressBar, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt
from src.processors.wholesale_order import WholesaleProcessor

class WholesaleTab(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.processor = None
        self.excel_file = None
        
    def initUI(self):
        layout = QVBoxLayout()
        
        self.status_label = QLabel('작성한 사입발주 양식 파일을 첨부해주세요')
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)
        
        self.file_btn = QPushButton('엑셀 파일 선택')
        self.file_btn.clicked.connect(self.select_file)
        layout.addWidget(self.file_btn)
        
        self.file_path = QLabel()
        self.file_path.setWordWrap(True)
        layout.addWidget(self.file_path)
        
        self.process_btn = QPushButton('거래처별 발주서 생성')
        self.process_btn.clicked.connect(self.process_file)
        self.process_btn.setEnabled(False)
        layout.addWidget(self.process_btn)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        self.setLayout(layout)
    
    def select_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "파일 선택",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        
        if file_name:
            self.file_path.setText(file_name)
            self.excel_file = file_name
            self.process_btn.setEnabled(True)
            self.status_label.setText('파일이 선택되었습니다')
    
    def process_file(self):
        try:
            save_folder = QFileDialog.getExistingDirectory(
                self,
                "저장할 폴더 선택"
            )
            
            if not save_folder:
                return
            
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            self.file_btn.setEnabled(False)
            self.process_btn.setEnabled(False)
            
            self.processor = WholesaleProcessor(self.excel_file, save_folder)
            self.processor.progress.connect(self.update_progress)
            self.processor.finished.connect(self.process_completed)
            self.processor.error.connect(self.process_error)
            self.processor.start()
            
        except Exception as e:
            self.process_error(str(e))
    
    def update_progress(self, value, message):
        self.progress_bar.setValue(value)
        self.status_label.setText(message)
    
    def process_completed(self):
        QMessageBox.information(self, "완료", "모든 발주서가 생성되었습니다!")
        self.reset_ui()
    
    def process_error(self, error_message):
        QMessageBox.critical(self, "오류", f"처리 중 오류가 발생했습니다: {error_message}")
        self.reset_ui()
    
    def reset_ui(self):
        self.status_label.setText('사입용 엑셀 파일을 선택해주세요')
        self.file_btn.setEnabled(True)
        self.process_btn.setEnabled(False)
        self.progress_bar.setVisible(False)
        self.excel_file = None 