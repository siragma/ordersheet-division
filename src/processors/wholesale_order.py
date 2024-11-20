import os
from datetime import datetime
import pandas as pd
from PyQt5.QtCore import QThread, pyqtSignal

class WholesaleProcessor(QThread):
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
            try:
                df = pd.read_excel(self.excel_file, sheet_name='발주양식', header=1)
            except ValueError as e:
                if "No sheet named" in str(e):
                    self.error.emit("'발주양식' 시트를 찾을 수 없습니다. 엑셀 파일의 시트 이름을 확인해주세요.")
                    return
                raise e
            
            # 필요한 열만 선택 (순서 변경)
            columns_needed = ['거래처명', '자사코드', '상품코드', '상품명', '칼라명', '사이즈', '발주수량', '공급가']  # 순서 변경
            try:
                df_filtered = df[columns_needed]
            except KeyError as e:
                missing_columns = [col for col in columns_needed if col not in df.columns]
                self.error.emit(f"필요한 열을 찾을 수 없습니다: {', '.join(missing_columns)}")
                return
            
            # NaN 값을 0으로 변경
            df_filtered['공급가'] = df_filtered['공급가'].fillna(0)
            df_filtered['발주수량'] = df_filtered['발주수량'].fillna(0)
            
            # 발주수량이 있는 행만 필터링 (0이나 빈 값 제외)
            df_filtered = df_filtered[df_filtered['발주수량'] != 0]
            
            # 공급가합 계산
            df_filtered['공급가합'] = df_filtered['공급가'].astype(float) * df_filtered['발주수량'].astype(float)
            
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
                total_amount = vendor_df['공급가합'].sum()
                
                # 파일명 생성
                file_name = f'(홀라인){vendor}_발주서_{current_date}.xlsx'
                full_path = os.path.join(self.save_folder, file_name)
                
                # NaN 값을 모두 0으로 변환
                vendor_df = vendor_df.fillna(0)
                
                # 엑셀 파일로 저장
                with pd.ExcelWriter(full_path, engine='xlsxwriter') as writer:
                    vendor_df.to_excel(writer, sheet_name='발주내역', index=False, na_rep='0')
                    workbook = writer.book
                    worksheet = writer.sheets['발주내역']
                    
                    # 스타일 포맷 정의
                    header_format = workbook.add_format({
                        'bg_color': '#D9D9D9',
                        'border': 1,
                        'valign': 'vcenter',
                        'align': 'center',
                        'bold': True,
                        'font_size': 9
                    })
                    
                    cell_format = workbook.add_format({
                        'border': 1,
                        'valign': 'vcenter',
                        'font_size': 9
                    })
                    
                    total_format = workbook.add_format({
                        'bold': True,
                        'valign': 'vcenter',
                        'num_format': '#,##0',
                        'font_size': 9
                    })
                    
                    empty_format = workbook.add_format({
                        'valign': 'vcenter',
                        'font_size': 9
                    })
                    
                    number_format = workbook.add_format({
                        'border': 1,
                        'valign': 'vcenter',
                        'num_format': '#,##0',
                        'font_size': 9
                    })
                    
                    # 헤더 스타일 적용
                    for col_num, value in enumerate(vendor_df.columns):
                        worksheet.write(0, col_num, value, header_format)
                    
                    # 데이터 셀 스타일 적용
                    for row in range(len(vendor_df)):
                        for col in range(len(vendor_df.columns)):
                            value = vendor_df.iloc[row, col]
                            if col == 8:  # 공급가합 열
                                # 엑셀 수식 사용 (발주수량 * 공급가)
                                formula = f'=G{row+2}*H{row+2}'  # row+2는 헤더행과 1부터 시작하는 엑셀 행 번호 때문
                                worksheet.write_formula(row + 1, col, formula, number_format)
                            elif col in [6, 7]:  # 발주수량, 공급가 열
                                worksheet.write(row + 1, col, value, number_format)
                            else:
                                worksheet.write(row + 1, col, value, cell_format)
                    
                    # 합계 행 추가
                    row_count = len(vendor_df) + 1
                    for col in range(len(vendor_df.columns)):
                        if col == 6:  # 발주수량 열
                            worksheet.write(row_count, col, total_qty, total_format)
                        elif col == 8:  # 공급가합 열
                            # 공급가합 열의 합계도 수식으로 변경
                            last_data_row = row_count
                            formula = f'=SUM(I2:I{last_data_row})'
                            worksheet.write_formula(row_count, col, formula, total_format)
                        else:
                            worksheet.write(row_count, col, '', empty_format)
                    
                    # 열 너비 자동 조정
                    for col_num, column in enumerate(vendor_df.columns):
                        max_length = max(
                            vendor_df[column].astype(str).apply(len).max(),
                            len(str(column)),
                            3
                        )
                        
                        if column == '상품명':
                            worksheet.set_column(col_num, col_num, max_length + 8)  # 여유 공간 증가
                        elif column == '거래처명':
                            worksheet.set_column(col_num, col_num, max_length + 5)   # 거래처명 여유 공간
                        elif column == '발주수량':
                            worksheet.set_column(col_num, col_num, max_length + 4)   # 발주수량 여유 공간
                        else:
                            worksheet.set_column(col_num, col_num, max_length + 1)
                
                self.msleep(100)
            
            # 마지막 진행률 업데이트
            for i in range(80, 101, 2):
                self.progress.emit(i, '마무리 중...')
                self.msleep(50)
            
            self.finished.emit()
            
        except Exception as e:
            self.error.emit(str(e))