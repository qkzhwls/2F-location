import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import xlwt
from datetime import datetime

# ==========================================
# 1. 공통 헬퍼 함수
# ==========================================
def normalize_token(text):
    """문자열 앞뒤 공백 제거 및 문자열 변환"""
    if pd.isna(text):
        return ""
    return str(text).strip()

def is_excluded_opt(text):
    """'2층' 텍스트가 포함되면 제외 (VBA IsExcludedOpt 구현)"""
    if "2층" in text:
        return True
    return False

def save_as_xls(dataframe, output_path, sheet_name="Sheet1"):
    """DataFrame을 .xls (Excel 97-2003) 형식으로 저장"""
    try:
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet(sheet_name)

        # 헤더 쓰기
        columns = list(dataframe.columns)
        for col_idx, col_name in enumerate(columns):
            ws.write(0, col_idx, str(col_name))

        # 데이터 쓰기
        for row_idx, row in enumerate(dataframe.itertuples(index=False), start=1):
            for col_idx, val in enumerate(row):
                if pd.isna(val):
                    val = ""
                ws.write(row_idx, col_idx, str(val))

        wb.save(output_path)
        return True
    except Exception as e:
        raise Exception(f"xls 저장 실패: {str(e)}")

def get_file_path():
    """파일 선택 대화상자"""
    file_path = filedialog.askopenfilename(
        title="작업할 통합 엑셀 파일(.xlsm, .xlsx)을 선택하세요",
        filetypes=[("Excel Files", "*.xlsm *.xlsx *.xls")]
    )
    return file_path

# ==========================================
# 2. 기능 1: 당일입고 변환
# ==========================================
def run_daily_inbound():
    file_path = get_file_path()
    if not file_path: return

    try:
        # 데이터 로드
        df_in = pd.read_excel(file_path, sheet_name="입고전표", engine='openpyxl')
        # 실사전표는 헤더가 없거나 복잡할 수 있어 header=None으로 읽고 인덱스로 접근
        df_audit = pd.read_excel(file_path, sheet_name="당일입고실사전표", header=None, engine='openpyxl')

        # Dictionary: 상품코드 -> 추가될 옵션 문자열
        dict_code_opt = {}
        
        # 1) 블록 스캔 로직 (VBA Do While i <= lastRow실사 구현)
        i = 0
        max_row = len(df_audit)
        
        while i < max_row:
            row_val = normalize_token(df_audit.iloc[i, 0]) # A열
            
            if row_val == "전표번호":
                # 구간 설정 (VBA: i + 2 -> Python Index: i + 2)
                # VBA는 헤더 포함 행번호, Python iloc은 0부터 시작. 
                # VBA: i(전표번호) -> 데이터 시작 i+2. 
                audit_start = i + 2 
                audit_end = audit_start
                
                # 구간 끝 찾기
                while audit_end < max_row and normalize_token(df_audit.iloc[audit_end, 0]) != "전표번호":
                    audit_end += 1
                
                # 해당 구간(audit_start ~ audit_end) 처리
                option_val = ""
                
                # 2층 옵션 찾기
                for j in range(audit_start, audit_end):
                    if j >= max_row: break
                    col_e = normalize_token(df_audit.iloc[j, 4]) # E열
                    if col_e == "2층":
                        temp_opt = normalize_token(df_audit.iloc[j, 5]) # F열
                        if not is_excluded_opt(temp_opt):
                            option_val = temp_opt
                        break
                
                # 옵션 적용
                if option_val:
                    for j in range(audit_start, audit_end):
                        if j >= max_row: break
                        col_e = normalize_token(df_audit.iloc[j, 4]) # E열
                        if col_e != "2층":
                            p_code = normalize_token(df_audit.iloc[j, 2]) # C열
                            if p_code:
                                if p_code in dict_code_opt:
                                    # 중복 체크 후 추가
                                    if option_val not in dict_code_opt[p_code].split(','):
                                        dict_code_opt[p_code] += "," + option_val
                                else:
                                    dict_code_opt[p_code] = option_val
                
                i = audit_end # 다음 검색 위치 이동
            else:
                i += 1

        # 2) 결과 생성 (변경분만)
        results = []
        
        for idx, row in df_in.iterrows():
            p_code = normalize_token(row[0]) # A열
            original_b = normalize_token(row[1]) # B열
            new_val = original_b
            
            if p_code in dict_code_opt:
                added_opts = dict_code_opt[p_code].split(',')
                for opt in added_opts:
                    opt = opt.strip()
                    if opt and not is_excluded_opt(opt):
                        # 기존 값에 없는 경우에만 추가 (VBA InStr 로직 대체)
                        current_tokens = [x.strip() for x in new_val.split(',')]
                        if opt not in current_tokens:
                            if new_val:
                                new_val += "," + opt
                            else:
                                new_val = opt
            
            # 값이 변경된 경우만 결과에 추가
            if new_val != original_b:
                results.append({
                    "상품코드": p_code,
                    "옵션추가항목1": new_val.upper()
                })

        if not results:
            messagebox.showinfo("알림", "옵션이 변경된 상품이 없습니다.")
            return

        # 저장
        df_result = pd.DataFrame(results)
        today_str = datetime.now().strftime("%Y%m%d")
        save_path = os.path.join(os.path.dirname(file_path), f"입고전표_당일입고변환_{today_str}.xls")
        
        save_as_xls(df_result, save_path, sheet_name="입고전표_당일입고변환")
        messagebox.showinfo("성공", f"당일입고변환 완료!\n{save_path}")

    except Exception as e:
        messagebox.showerror("오류", str(e))

# ==========================================
# 3. 기능 2: 자리 변경
# ==========================================
def run_location_change():
    file_path = get_file_path()
    if not file_path: return

    try:
        df_in = pd.read_excel(file_path, sheet_name="입고전표", engine='openpyxl')
        df_z = pd.read_excel(file_path, sheet_name="자리변경실사전표", header=None, engine='openpyxl')

        # 1) 입고전표 데이터를 딕셔너리로 로드 (검색 속도 향상)
        # {상품코드: 옵션값}
        opt_dict = {}
        for idx, row in df_in.iterrows():
            p_code = normalize_token(row[0])
            if p_code:
                opt_dict[p_code] = normalize_token(row[1])

        # 2) 2층 위치 수집 (VBA Logic)
        floor2_indices = []
        for i in range(len(df_z)):
            # VBA는 row 2부터 시작, 여기선 header=None이므로 1행(index 0)은 제목일 수 있음.
            # 하지만 VBA 로직상 E열 값만 체크하므로 전체 스캔해도 무방
            col_e = normalize_token(df_z.iloc[i, 4]) # E열
            if col_e == "2층":
                floor2_indices.append(i)

        if len(floor2_indices) % 2 != 0:
            messagebox.showwarning("경고", "2층 항목 개수가 짝수가 아닙니다.")
            return

        # 3) 페어 치환 로직
        modified_codes = set() # 변경된 상품코드 추적용

        for i in range(0, len(floor2_indices), 2):
            start_idx = floor2_indices[i]
            end_idx = floor2_indices[i+1]
            
            delete_opt = normalize_token(df_z.iloc[start_idx, 5]) # F열
            add_opt = normalize_token(df_z.iloc[end_idx, 5])   # F열
            
            if is_excluded_opt(delete_opt): delete_opt = ""
            if is_excluded_opt(add_opt): add_opt = ""

            # 사이 구간 순회
            for j in range(start_idx + 1, end_idx):
                col_e = normalize_token(df_z.iloc[j, 4])
                if col_e != "2층":
                    p_code = normalize_token(df_z.iloc[j, 2]) # C열
                    
                    if p_code in opt_dict:
                        current_opt = opt_dict[p_code]
                        found_replaced = False
                        
                        if current_opt:
                            # 쉼표로 분리하여 정확히 일치하는 토큰만 교체
                            opt_list = [x.strip() for x in current_opt.split(',')]
                            new_opt_list = []
                            for token in opt_list:
                                if delete_opt and token == delete_opt:
                                    new_opt_list.append(add_opt)
                                    found_replaced = True
                                else:
                                    new_opt_list.append(token)
                            
                            if found_replaced:
                                # 재조립 (빈값 제거, 중복 쉼표 제거는 split/join으로 자연스럽게 해결됨)
                                # add_opt가 빈값이면 삭제 효과
                                clean_list = [x for x in new_opt_list if x]
                                opt_dict[p_code] = ",".join(clean_list)
                                modified_codes.add(p_code)

        # 4) 실사에 등장했던 코드만 필터링 (VBA logic Step 4)
        # 하지만 VBA Step 5를 보면 'optDict.keys'를 순회하며 'filterDict'에 있는 것만 저장함.
        # 여기서 filterDict는 실사전표 C열에 있는 모든 코드.
        filter_dict = set()
        for i in range(len(df_z)):
            # 헤더(행0) 제외하고 데이터만
            if i > 0: 
                c_val = normalize_token(df_z.iloc[i, 2])
                if c_val: filter_dict.add(c_val)

        # 5) 저장 결과 생성
        final_rows = []
        for p_code, opt_val in opt_dict.items():
            if p_code in filter_dict and p_code in modified_codes: # 변경된 것만 저장하려면 modified_codes 체크 필요
                # VBA 코드는 modified 여부와 상관없이 filterDict에 있고 값이 있으면 저장하는 듯 보이나
                # 논리상 변경된 것을 저장하는 것이 맞음 (VBA 원본: If Trim(optDict(productCode)) <> "" Then 저장)
                # 안전하게 VBA 로직 그대로: 필터에 있고 값이 비어있지 않으면 저장
                if opt_val:
                    final_rows.append({
                        "상품코드": p_code,
                        "옵션추가항목1": opt_val.upper()
                    })

        if not final_rows:
            messagebox.showinfo("알림", "조건에 맞는 데이터가 없습니다.")
            return

        df_result = pd.DataFrame(final_rows)
        today_str = datetime.now().strftime("%Y%m%d")
        save_path = os.path.join(os.path.dirname(file_path), f"입고전표_자리변경결과_{today_str}.xls")
        
        save_as_xls(df_result, save_path, sheet_name="자리변경")
        messagebox.showinfo("성공", f"자리변경 결과 저장 완료:\n{save_path}")

    except Exception as e:
        messagebox.showerror("오류", str(e))


# ==========================================
# 4. 기능 3: 재입고
# ==========================================
def run_restock():
    file_path = get_file_path()
    if not file_path: return

    try:
        df_in = pd.read_excel(file_path, sheet_name="입고전표", engine='openpyxl')
        df_audit = pd.read_excel(file_path, sheet_name="재입고변경실사전표", header=None, engine='openpyxl')

        rep_dict = {}
        filter_dict = set()
        is_in_block = False
        current_rep = ""

        # 1) 실사 누적
        for i in range(len(df_audit)):
            col_a = normalize_token(df_audit.iloc[i, 0])
            
            if col_a == "전표번호":
                is_in_block = False
                current_rep = ""
            elif col_a.isdigit() or (col_a.replace('.','',1).isdigit()):
                p_code = normalize_token(df_audit.iloc[i, 2]) # C열
                
                if not is_in_block:
                    current_rep = normalize_token(df_audit.iloc[i, 5]) # F열
                    is_in_block = True
                
                if not is_excluded_opt(current_rep):
                    if p_code and current_rep:
                        if p_code in rep_dict:
                            rep_dict[p_code] += "," + current_rep
                        else:
                            rep_dict[p_code] = current_rep
                        filter_dict.add(p_code)

        # 2) 병합 및 중복 제거
        final_rows = []
        for idx, row in df_in.iterrows():
            key = normalize_token(row[0]) # A열
            val_b = normalize_token(row[1]) # B열
            
            if key and key in filter_dict:
                all_opts = rep_dict.get(key, "") + "," + val_b
                
                # 중복제거 (Dedup)
                dedup_set = set()
                for v in all_opts.split(','):
                    v = normalize_token(v)
                    if v and not is_excluded_opt(v):
                        dedup_set.add(v)
                
                final_val = ",".join(sorted(list(dedup_set))).upper()
                
                final_rows.append({
                    "상품코드": key,
                    "옵션추가항목1": final_val
                })

        if not final_rows:
            messagebox.showinfo("알림", "저장할 데이터가 없습니다.")
            return

        # 3) 저장
        df_result = pd.DataFrame(final_rows)
        today_str = datetime.now().strftime("%Y%m%d")
        save_path = os.path.join(os.path.dirname(file_path), f"입고전표_재입고변경결과_{today_str}.xls")
        
        save_as_xls(df_result, save_path, sheet_name="재입고")
        messagebox.showinfo("성공", f"재입고변경 결과 저장 완료:\n{save_path}")

    except Exception as e:
        messagebox.showerror("오류", str(e))

# ==========================================
# 5. 메인 GUI
# ==========================================
if __name__ == "__main__":
    root = tk.Tk()
    root.title("전표 변환 통합 프로그램")
    root.geometry("400x250")

    lbl = tk.Label(root, text="원하는 작업을 선택하세요\n(입고전표, 실사전표가 포함된 엑셀 파일 필요)", pady=10)
    lbl.pack()

    btn1 = tk.Button(root, text="1. 당일입고 변환", command=run_daily_inbound, height=2, width=30, bg="#e1f5fe")
    btn1.pack(pady=5)

    btn2 = tk.Button(root, text="2. 자리 변경", command=run_location_change, height=2, width=30, bg="#fff9c4")
    btn2.pack(pady=5)

    btn3 = tk.Button(root, text="3. 재입고 (옵션추가)", command=run_restock, height=2, width=30, bg="#ffebee")
    btn3.pack(pady=5)

    root.mainloop()
