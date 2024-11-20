import os
from datetime import datetime
import pandas as pd
from PyQt5.QtCore import QThread, pyqtSignal

class ConsignmentProcessor(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    
    def __init__(self, excel_file, save_folder):
        super().__init__()
        self.excel_file = excel_file
        self.save_folder = save_folder
        
    def run(self):
        try:
            # 엑셀 파일 읽기
            self.progress.emit(10, '파일 읽는 중...')
            df = pd.read_excel(self.excel_file, sheet_name='발주양식')
            
            # 필요한 열만 선택
            self.progress.emit(20, '데이터 처리 중...')
            columns_needed = ['거래처명', '자사코드', '상품코드', '상품명', '칼라명', '사이즈', '발주수량']
            df_filtered = df[columns_needed]
            
            # 발주수량이 있는 행만 필터링
            df_filtered = df_filtered[df_filtered['발주수량'].notna() & (df_filtered['발주수량'] != 0) & (df_filtered['발주수량'] != '')]
            
            # 현재 날짜
            current_date = datetime.now().strftime('%y%m%d')
            
            # 거래처별로 파일 생성
            vendors = df_filtered['거래처명'].unique()
            total_vendors = len(vendors)
            
            for i, vendor in enumerate(vendors, 1):
                self.progress.emit(
                    20 + int((i / total_vendors) * 60), 
                    f'처리 중... ({i}/{total_vendors} 거래처)'
                )
                
                vendor_df = df_filtered[df_filtered['거래처명'] == vendor]
                total_qty = vendor_df['발주수량'].sum()
                
                # 파일명 생성
                file_name = f'(홀라인){vendor}_발주서_{current_date}.xlsx'
                full_path = os.path.join(self.save_folder, file_name)
                
                # 엑셀 파일로 저장
                with pd.ExcelWriter(full_path, engine='xlsxwriter') as writer:
                    vendor_df.to_excel(writer, sheet_name='발주내역', index=False)
                    workbook = writer.book
                    worksheet = writer.sheets['발주내역']
                    
                    # 스타일 포맷 정의
                    header_format = workbook.add_format({
                        'bg_color': '#D9D9D9',  # 회색 배경
                        'border': 1,            # 테두리
                        'valign': 'vcenter',    # 수직 가운데 정렬
                        'align': 'center',      # 가운데 정렬 (헤더만)
                        'bold': True            # 굵은 글씨 (헤더만)
                    })
                    
                    cell_format = workbook.add_format({
                        'border': 1,            # 테두리
                        'valign': 'vcenter'     # 수직 가운데 정렬만 유지
                    })
                    
                    total_format = workbook.add_format({
                        'bold': True,           # 굵은 글씨
                        'valign': 'vcenter'     # 수직 가운데 정렬만 유지
                    })
                    
                    empty_format = workbook.add_format({
                        'valign': 'vcenter'     # 수직 가운데 정렬만 유지
                    })
                    
                    # 헤더 스타일 적용
                    for col_num, value in enumerate(columns_needed):
                        worksheet.write(0, col_num, value, header_format)
                    
                    # 데이터 셀 스타일 적용
                    for row in range(len(vendor_df)):
                        for col in range(len(columns_needed)):
                            worksheet.write(row + 1, col, vendor_df.iloc[row, col], cell_format)
                    
                    # 합계 행 추가
                    row_count = len(vendor_df) + 1
                    for col in range(len(columns_needed)):
                        if col == 6:  # 발주수량 열
                            worksheet.write(row_count, col, total_qty, total_format)
                        else:
                            worksheet.write(row_count, col, '', empty_format)
                    
                    # 열 너비 자동 조정
                    for col_num, column in enumerate(columns_needed):
                        max_length = max(
                            vendor_df[column].astype(str).apply(len).max(),
                            len(str(column)),
                            3
                        )
                        
                        if column == '상품명':  # 상품명 열은 더 넓게 설정
                            worksheet.set_column(col_num, col_num, max_length + 7)  # 여유 공간 더 추가
                        else:
                            worksheet.set_column(col_num, col_num, max_length + 1)  # 다른 열은 기존대로
                
                # 잠시 대기하여 프로그레스바가 너무 빨리 진행되지 않도록 함
                self.msleep(100)
            
            # 마지막 진행률 업데이트를 천천히 수행
            for i in range(80, 101, 2):
                self.progress.emit(i, '마무리 중...')
                self.msleep(50)
            
            self.finished.emit()
            
        except Exception as e:
            self.error.emit(str(e))