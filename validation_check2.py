import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
from datetime import datetime

WATCH_FOLDER = "C:\XBRL"
def validate_and_highlight_excel(file_path, output_file):
    # 엑셀 파일의 첫 번째 시트만 읽기
    df = pd.read_excel(file_path, sheet_name=0, engine='openpyxl', dtype=str)
    
    # 문자열 변환 (혹시 dtype=str이 적용되지 않은 경우 대비)
    df = df.astype(str).applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # 1번열(시트)과 7번열(맵핑표) 데이터 추출
    sheet_col = df['시트'].tolist()  # 1번열 (시트)
    mapping_col = df['맵핑표 번호'].tolist()  # 7번열 (맵핑표)
    xbrl_col = df['XBRL표 번호'].tolist()
    content_col = df['content'].tolist()
    depth_col = df['depth'].tolist()
    label_col = df['label(국문)'].tolist()

    error_cells = []  # 오류 발생한 셀 목록

    # ✅ **[1] 시트가 변경될 때 맵핑표 번호가 1로 리셋되는지 확인**
    prev_sheet = sheet_col[0]

    for i in range(1, len(sheet_col)):  
        current_sheet = sheet_col[i]
        current_mapping = mapping_col[i]

        if current_sheet != prev_sheet:
            if current_mapping.strip() != "1":
                error_cells.append((i + 2, df.columns.get_loc('맵핑표 번호') + 1))  # 엑셀 행, 열 좌표 저장
            prev_sheet = current_sheet  

    # ✅ **[2] 맵핑표 번호가 변경될 때 XBRL 번호 검증**
    for i in range(1, len(mapping_col)):  
        mapping_no = mapping_col[i].strip()
        xbrl_no = xbrl_col[i].strip()
        content = content_col[i].strip()

        if mapping_no == "1":  
            if content == "text":
                if xbrl_no != "0":
                    error_cells.append((i + 2, df.columns.get_loc('XBRL표 번호') + 1))
            else:  
                xbrl_split = xbrl_no.split(",")[0].strip() if "," in xbrl_no else xbrl_no
                if xbrl_split != "1":
                    error_cells.append((i + 2, df.columns.get_loc('XBRL표 번호') + 1))

    # ✅ **[3] XBRL 번호가 변경될 때 depth가 1로 초기화되는지 확인**
    prev_xbrl = xbrl_col[0].strip()

    for i in range(1, len(xbrl_col)):  
        current_xbrl = xbrl_col[i].strip()
        current_depth = depth_col[i].strip()

        if current_xbrl != prev_xbrl:  
            if current_depth != "1":
                error_cells.append((i + 2, df.columns.get_loc('depth') + 1))  
            prev_xbrl = current_xbrl  

    # 📌 **오류가 발생한 셀을 노란색으로 하이라이트**
    wb = load_workbook(file_path)
    ws = wb.active  # 첫 번째 시트 선택
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for row, col in error_cells:
        ws.cell(row=row, column=col).fill = yellow_fill  # 해당 셀을 노란색으로 변경
        ws.cell(row=1, column=col).fill = yellow_fill

    # 변경된 엑셀 저장
    wb.save(output_file)

    print(error_cells)
    for cell in error_cells:
        print(cell, label_col[cell[0]])
    return len(error_cells) == 0  # 오류가 없으면 True, 있으면 False

# ✅ **사용 예시**
file_name = "1.1.0-mptg_master_별도_20250219102913_(BS,PL,CE,CF,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40).xlsx"
folder_path = r"C:\XBRL"  # 백슬래시 문제 방지
file_path = os.path.join(folder_path, file_name)



class FileHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory:
            file_path = event.src_path
            print(f"📂 새 파일 감지됨: {file_path}")
            
            # 파일 검증 실행
            if self.validate_file(file_path):
                print("✅ 파일 검증 성공!")
            else:
                print("❌ 파일 검증 실패 - 파일 삭제")
                # 잘못된 부분 색칠해서 반환

    def validate_file(self, df):
        """
        Pandas DataFrame을 이용한 데이터 검증
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"검증결과_{timestamp}.xlsx"

        result = validate_and_highlight_excel(file_path, file_name)
        print("✅ 검증 결과:", result)  # 모든 행이 규칙을 만족하면 True, 아니면 False

        
        return True

if __name__ == "__main__":
    event_handler = FileHandler()
    observer = Observer()
    observer.schedule(event_handler, WATCH_FOLDER, recursive=False)
    observer.start()
    
    print(f"📡 폴더 모니터링 시작: {WATCH_FOLDER}")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("🛑 폴더 모니터링 중지됨")

    observer.join()
