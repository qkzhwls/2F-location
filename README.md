import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import xlwt 
from datetime import datetime
import ctypes
import warnings

# 윈도우 배율 호환
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

warnings.filterwarnings("ignore")

# ==========================================
# [설정] 정확한 헤더명 지정
# ==========================================
CONFIG = {
    "INBOUND": {
        "CODE_HEADERS": ["상품코드"],           
        "OPTION_HEADERS": ["옵션추가항목1"]      
    },
    "AUDIT": {
        "ANCHOR_TEXT": "전표번호",
        "TARGET_TEXT": "2층",
        "CODE_OFFSET": 2,    
        "OPTION_OFFSET": 1   
    }
}

path_inbound_file = ""
path_audit_files = []

# ==========================================
# 1. 헬퍼 함수
# ==========================================
def normalize_token(text):
    if pd.isna(text): return ""
    return str(text).strip()

def clean_text(text):
    if pd.isna(text): return ""
    return str(text).replace(" ", "").strip()

def is_excluded_opt(text):
    if "2층" in str(text): return True
    return False

def log(msg):
    timestamp = datetime.now().strftime("[%H:%M:%S] ")
    try:
        txt_log.insert(tk.END, timestamp + str(msg) + "\n")
        txt_log.see(tk.END)
    except:
        print(timestamp + str(msg))

def save_as_xls_direct(dataframe, output_path, sheet_name="Sheet1"):
    try:
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet(sheet_name)
        headers = dataframe.columns.tolist()
        for col_idx, header in enumerate(headers):
            ws.write(0, col_idx, str(header))
        for row_idx, row in enumerate(dataframe.values):
            for col_idx, val in enumerate(row):
                val_str = str(val) if not pd.isna(val) else ""
                ws.write(row_idx + 1, col_idx, val_str)
        wb.save(output_path)
        return True
    except Exception as e:
        log(f"[저장 오류] {e}")
        return False

def load_excel_or_csv(file_path, sheet_name=0, header=0):
    ext = os.path.splitext(file_path)[1].lower()
    filename = os.path.basename(file_path)
    log(f"파일 읽기: {filename}")
    
    df = None
    if ext == '.csv':
        encodings = ['utf-8-sig', 'cp949', 'euc-kr', 'utf-8']
        for enc in encodings:
            try:
                try: df = pd.read_csv(file_path, encoding=enc, header=header)
                except: df = pd.read_csv(file_path, encoding=enc, header=header, skiprows=1)
                df_str = df.head(5).to_string()
                if "전표" in df_str or "상품" in df_str or "Code" in df_str:
                    log(f"-> CSV 읽기 성공 ({enc})")
                    break
            except: continue
    if df is None:
        try:
            if ext == '.xls':
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=header, engine='xlrd')
            elif ext in ['.xlsx', '.xlsm']:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=header, engine='openpyxl')
            else:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=header)
        except Exception as e:
            try:
                try: dfs = pd.read_html(file_path, header=header, encoding='utf-8')
                except: dfs = pd.read_html(file_path, header=header, encoding='cp949')
                if dfs: df = dfs[0]; log("-> HTML 형식 읽기 성공")
            except:
                log(f"-> 읽기 실패: {e}")
                raise e
    return df

# ==========================================
# 2. 로직 (실사전표 기준 + 표지판 제외 / v28.0)
# ==========================================

def find_header_row_index(df, keywords):
    for r in range(min(5, len(df))):
        row_str = df.iloc[r].astype(str).tolist()
        for val in row_str:
            for k in keywords:
                if k in clean_text(val): return r
    return 0

def find_col_index_by_name(df, keywords):
    for idx, col_name in enumerate(df.columns):
        col_str = clean_text(col_name)
        for k in keywords:
            if k in col_str: return idx
    return -1

def find_value_in_row(row, keyword):
    for idx, val in enumerate(row):
        if keyword in clean_text(val): return idx
    return -1

def get_safe_value(df, row_idx, col_idx):
    try: return normalize_token(df.iloc[row_idx, col_idx])
    except: return ""

def prepare_inbound_df(file_path):
    df_raw = load_excel_or_csv(file_path, header=None)
    header_idx = find_header_row_index(df_raw, CONFIG["INBOUND"]["CODE_HEADERS"])
    
    final_df = None
    if header_idx > 0:
        log(f"-> 입고전표 헤더 발견: {header_idx+1}행")
        new_header = df_raw.iloc[header_idx]
        df_data = df_raw[header_idx+1:].reset_index(drop=True)
        df_data.columns = new_header
        final_df = df_data
    else:
        final_df = load_excel_or_csv(file_path, header=0)
        
    if final_df is not None:
        count = len(final_df)
        log(f"-> [확인] 입고전표 데이터 로드 완료 ({count}행)")
        
    return final_df

def get_inbound_mapping(df_in):
    cfg = CONFIG["INBOUND"]
    idx_code = find_col_index_by_name(df_in, cfg["CODE_HEADERS"])
    idx_opt = find_col_index_by_name(df_in, cfg["OPTION_HEADERS"])
    if idx_code == -1: idx_code = 0
    if idx_opt == -1: idx_opt = 1 if df_in.shape[1] < 3 else 2
    return idx_code, idx_opt

def build_inbound_dict(df_in):
    """입고전표를 딕셔너리로 변환 {상품코드: 기존옵션}"""
    idx_code, idx_opt = get_inbound_mapping(df_in)
    inbound_dict = {}
    for i in range(len(df_in)):
        k = normalize_token(df_in.iloc[i, idx_code])
        v = normalize_token(df_in.iloc[i, idx_opt])
        if k: inbound_dict[k] = v
    return inbound_dict

def scan_audit_files_robust(p_audit_list):
    """실사전표를 스캔하여 {상품코드: 스캔된_위치값(콤마로누적)} 형태의 딕셔너리 반환"""
    dict_scanned = {}
    audit_cfg = CONFIG["AUDIT"]

    for p_audit in p_audit_list:
        try: df_audit = load_excel_or_csv(p_audit, sheet_name=0, header=None)
        except: continue
        
        col_idx_anchor = 0
        for r in range(min(20, len(df_audit))):
            idx = find_value_in_row(df_audit.iloc[r], audit_cfg["ANCHOR_TEXT"])
            if idx != -1: col_idx_anchor = idx; break
        
        i = 0; max_row = len(df_audit)
        while i < max_row:
            cell_val = clean_text(get_safe_value(df_audit, i, col_idx_anchor))
            
            if audit_cfg["ANCHOR_TEXT"] in cell_val:
                s = i + 2; e = s
                while e < max_row:
                    next_val = clean_text(get_safe_value(df_audit, e, col_idx_anchor))
                    if audit_cfg["ANCHOR_TEXT"] in next_val: break
                    e += 1
                
                # 블록 내 순차 스캔
                current_opt = ""
                for j in range(s, e):
                    if j >= max_row: break
                    row_data = df_audit.iloc[j]
                    
                    # 1. 위치 갱신 확인 (이 줄은 표지판인가?)
                    is_location_header = False
                    idx_target = find_value_in_row(row_data, audit_cfg["TARGET_TEXT"])
                    if idx_target != -1:
                        is_location_header = True # 표지판 행으로 표시
                        found_opt = get_safe_value(df_audit, j, idx_target + audit_cfg["OPTION_OFFSET"])
                        if not is_excluded_opt(found_opt):
                            current_opt = found_opt
                    
                    # 2. 상품 적용 (표지판 행은 절대 상품으로 추가하지 않음)
                    if current_opt and not is_location_header:
                        p = get_safe_value(df_audit, j, col_idx_anchor + audit_cfg["CODE_OFFSET"])
                        if p:
                            if p in dict_scanned:
                                current_opts = dict_scanned[p].split(',')
                                if current_opt not in current_opts:
                                    dict_scanned[p] += "," + current_opt
                            else:
                                dict_scanned[p] = current_opt
                i = e
            else: i += 1
    return dict_scanned

# [1. 당일입고] (실사전표 기준)
def logic_daily_inbound(p_in, p_audit_list):
    log("=== [1. 당일입고] 시작 (v28.0: 표지판제외) ===")
    try: df_in = prepare_inbound_df(p_in)
    except: return pd.DataFrame()
    
    inbound_dict = build_inbound_dict(df_in)
    scanned_data = scan_audit_files_robust(p_audit_list)
    log(f"-> 실사전표에서 {len(scanned_data)}개의 상품을 확인했습니다.")

    results = []
    for code, scanned_loc in scanned_data.items():
        orig = inbound_dict.get(code, "")
        
        new_v = orig
        adds = scanned_loc.split(',')
        for a in adds:
            a = a.strip()
            if a:
                toks = [x.strip() for x in new_v.split(',')]
                if a not in toks:
                    new_v = (new_v + "," + a) if new_v else a
        
        results.append({
            CONFIG["INBOUND"]["CODE_HEADERS"][0]: code,
            CONFIG["INBOUND"]["OPTION_HEADERS"][0]: new_v.upper()
        })
            
    log(f"완료: 총 {len(results)}개 상품 저장")
    return pd.DataFrame(results)

# [2. 자리변경] (실사전표 기준)
def logic_location_change(p_in, p_audit_list):
    log("=== [2. 자리변경] 시작 (v28.0: 표지판제외) ===")
    try: df_in = prepare_inbound_df(p_in)
    except: return pd.DataFrame()

    inbound_dict = build_inbound_dict(df_in)
    
    mod_results = {}
    audit_cfg = CONFIG["AUDIT"]

    for p_audit in p_audit_list:
        try: df_z = load_excel_or_csv(p_audit, sheet_name=0, header=None)
        except: continue

        f2_info = [] 
        for i in range(len(df_z)):
            idx = find_value_in_row(df_z.iloc[i], audit_cfg["TARGET_TEXT"])
            if idx != -1: f2_info.append((i, idx))
        
        if len(f2_info) % 2 != 0: continue
        code_col_idx = 2 
        if f2_info: code_col_idx = max(0, f2_info[0][1] - 2)

        for k in range(0, len(f2_info), 2):
            s_row, s_col = f2_info[k]
            e_row, e_col = f2_info[k+1]
            d_opt = get_safe_value(df_z, s_row, s_col + 1)
            a_opt = get_safe_value(df_z, e_row, e_col + 1)
            
            if is_excluded_opt(d_opt): d_opt = ""
            if is_excluded_opt(a_opt): a_opt = ""
            
            for j in range(s_row + 1, e_row):
                # 2층(표지판)이 아닌 행만 상품으로 취급
                if find_value_in_row(df_z.iloc[j], audit_cfg["TARGET_TEXT"]) == -1:
                    p = get_safe_value(df_z, j, code_col_idx)
                    if p:
                        curr = mod_results.get(p, inbound_dict.get(p, ""))
                        if curr:
                            lst = [x.strip() for x in curr.split(',')]
                            new_l = []
                            rep = False
                            for t in lst:
                                if d_opt and t == d_opt: 
                                    new_l.append(a_opt)
                                    rep = True
                                else: 
                                    new_l.append(t)
                            
                            if rep:
                                clean_l = [x for x in new_l if x]
                                mod_results[p] = ",".join(clean_l)

    results = []
    for code, val in mod_results.items():
        results.append({
            CONFIG["INBOUND"]["CODE_HEADERS"][0]: code,
            CONFIG["INBOUND"]["OPTION_HEADERS"][0]: val.upper()
        })
        
    log(f"완료: 총 {len(results)}개 상품 변경됨")
    return pd.DataFrame(results)

# [3. 재입고] (실사전표 기준)
def logic_restock(p_in, p_audit_list):
    log("=== [3. 재입고] 시작 (v28.0: 표지판제외) ===")
    try: df_in = prepare_inbound_df(p_in)
    except: return pd.DataFrame()

    inbound_dict = build_inbound_dict(df_in)
    scanned_data = {}
    audit_cfg = CONFIG["AUDIT"]

    for p_audit in p_audit_list:
        try: df_audit = load_excel_or_csv(p_audit, sheet_name=0, header=None)
        except: continue

        col_idx_anchor = 0
        for r in range(min(20, len(df_audit))):
            idx = find_value_in_row(df_audit.iloc[r], audit_cfg["ANCHOR_TEXT"])
            if idx != -1: col_idx_anchor = idx; break

        blk = False; c_rep = ""
        for i in range(len(df_audit)):
            cell_a = clean_text(get_safe_value(df_audit, i, col_idx_anchor))
            
            # 2층(표지판) 행은 재입고 로직에서 무시하거나 초기화 용도로만 씀
            if audit_cfg["ANCHOR_TEXT"] in cell_a:
                blk = False; c_rep = ""
            elif get_safe_value(df_audit, i, col_idx_anchor).replace('.','',1).isdigit():
                # 여기서도 2층 키워드가 있는 행인지 체크
                if find_value_in_row(df_audit.iloc[i], audit_cfg["TARGET_TEXT"]) != -1:
                    continue # 표지판 행이면 스킵

                p = get_safe_value(df_audit, i, col_idx_anchor + audit_cfg["CODE_OFFSET"]) 
                if not blk:
                    c_rep = get_safe_value(df_audit, i, col_idx_anchor + 5)
                    blk = True
                
                if not is_excluded_opt(c_rep) and p and c_rep:
                    if p in scanned_data:
                        if c_rep not in scanned_data[p]: scanned_data[p].append(c_rep)
                    else:
                        scanned_data[p] = [c_rep]

    results = []
    for code, new_locs in scanned_data.items():
        orig_val = inbound_dict.get(code, "")
        old_locs = [x.strip() for x in orig_val.split(',') if x.strip()]
        
        merged = []
        seen = set()
        for loc in new_locs:
            if loc not in seen and not is_excluded_opt(loc):
                merged.append(loc); seen.add(loc)
        for loc in old_locs:
            if loc not in seen and not is_excluded_opt(loc):
                merged.append(loc); seen.add(loc)
                
        results.append({
            CONFIG["INBOUND"]["CODE_HEADERS"][0]: code,
            CONFIG["INBOUND"]["OPTION_HEADERS"][0]: ",".join(merged).upper()
        })

    log(f"완료: 총 {len(results)}개 상품 저장")
    return pd.DataFrame(results)

# [4. 덮어쓰기] (실사전표 기준)
def logic_overwrite(p_in, p_audit_list):
    log("=== [4. 덮어쓰기] 시작 (v28.0: 표지판제외) ===")
    try: df_in = prepare_inbound_df(p_in)
    except: return pd.DataFrame()
    
    # 여기서도 표지판 제외 로직이 적용된 함수 사용
    scanned_data = scan_audit_files_robust(p_audit_list)
    log(f"-> 실사전표에서 {len(scanned_data)}개의 상품을 확인했습니다.")

    results = []
    for code, scanned_loc in scanned_data.items():
        results.append({
            CONFIG["INBOUND"]["CODE_HEADERS"][0]: code,
            CONFIG["INBOUND"]["OPTION_HEADERS"][0]: scanned_loc.upper()
        })
            
    log(f"완료: 총 {len(results)}개 상품 저장")
    return pd.DataFrame(results)

# ==========================================
# 3. 버튼/UI 핸들러
# ==========================================
def reset_all_files():
    global path_inbound_file, path_audit_files
    path_inbound_file = ""
    path_audit_files = []
    lbl_inbound_path.config(text="선택 안됨")
    lbl_audit_path.config(text="없음")
    log("모든 파일 선택이 초기화되었습니다.")

def check_ready():
    if not path_inbound_file:
        log("오류: 입고전표가 선택되지 않았습니다.")
        messagebox.showwarning("경고", "입고전표 파일을 선택하세요.")
        return False
    if not path_audit_files:
        log("오류: 실사전표가 선택되지 않았습니다.")
        messagebox.showwarning("경고", "실사전표 파일을 추가하세요.")
        return False
    return True

def run_btn1():
    if not check_ready(): return
    try:
        df = logic_daily_inbound(path_inbound_file, path_audit_files)
        if df.empty: 
            log("결과: 처리된 데이터가 없습니다.")
            messagebox.showinfo("알림", "처리된 데이터가 없습니다."); return
        
        save_name = f"입고전표_당일입고변환_{datetime.now().strftime('%Y%m%d')}.xls"
        save_path = os.path.join(os.path.dirname(path_inbound_file), save_name)
        if save_as_xls_direct(df, save_path, "당일입고변환"):
            log(f"저장 완료: {save_name}")
            messagebox.showinfo("완료", "저장되었습니다.")
            reset_all_files()
    except Exception as e: 
        log(f"오류 발생: {e}")
        messagebox.showerror("오류", str(e))

def run_btn2():
    if not check_ready(): return
    try:
        df = logic_location_change(path_inbound_file, path_audit_files)
        if df.empty: 
            log("결과: 변경 대상이 없습니다.")
            messagebox.showinfo("알림", "변경 대상 없음"); return
        
        save_name = f"입고전표_자리변경결과_{datetime.now().strftime('%Y%m%d')}.xls"
        save_path = os.path.join(os.path.dirname(path_inbound_file), save_name)
        if save_as_xls_direct(df, save_path, "자리변경"):
            log(f"저장 완료: {save_name}")
            messagebox.showinfo("완료", "저장되었습니다.")
            reset_all_files()
    except Exception as e: 
        log(f"오류 발생: {e}")
        messagebox.showerror("오류", str(e))

def run_btn3():
    if not check_ready(): return
    try:
        df = logic_restock(path_inbound_file, path_audit_files)
        if df.empty: 
            log("결과: 처리된 데이터가 없습니다.")
            messagebox.showinfo("알림", "처리된 데이터 없음"); return
        
        save_name = f"입고전표_재입고변경결과_{datetime.now().strftime('%Y%m%d')}.xls"
        save_path = os.path.join(os.path.dirname(path_inbound_file), save_name)
        if save_as_xls_direct(df, save_path, "재입고"):
            log(f"저장 완료: {save_name}")
            messagebox.showinfo("완료", "저장되었습니다.")
            reset_all_files()
    except Exception as e: 
        log(f"오류 발생: {e}")
        messagebox.showerror("오류", str(e))

def run_btn4_overwrite():
    if not check_ready(): return
    try:
        df = logic_overwrite(path_inbound_file, path_audit_files)
        if df.empty: 
            log("결과: 처리된 데이터가 없습니다.")
            messagebox.showinfo("알림", "처리된 데이터 없음"); return
        
        save_name = f"입고전표_삭제변환_{datetime.now().strftime('%Y%m%d')}.xls"
        save_path = os.path.join(os.path.dirname(path_inbound_file), save_name)
        if save_as_xls_direct(df, save_path, "삭제변환"):
            log(f"저장 완료: {save_name}")
            messagebox.showinfo("완료", "저장되었습니다.")
            reset_all_files()
    except Exception as e: 
        log(f"오류 발생: {e}")
        messagebox.showerror("오류", str(e))

# ==========================================
# 4. UI 설정
# ==========================================
def select_inbound_file():
    global path_inbound_file
    path = filedialog.askopenfilename(title="입고전표", filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All", "*.*")])
    if path:
        path_inbound_file = path
        lbl_inbound_path.config(text=f"입고: {os.path.basename(path)}")
        log(f"입고파일 선택됨: {os.path.basename(path)}")

def add_audit_files():
    global path_audit_files
    paths = filedialog.askopenfilenames(title="실사전표(다중)", filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All", "*.*")])
    if paths:
        path_audit_files.extend(list(paths))
        lbl_audit_path.config(text=f"총 {len(path_audit_files)}개 대기")
        log(f"실사파일 {len(paths)}개 추가됨 (총 {len(path_audit_files)}개)")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("전표 변환 v28.0 (표지판 제외)")
    root.geometry("600x650")

    frame_files = tk.LabelFrame(root, text="1. 파일 선택", padx=10, pady=10)
    frame_files.pack(fill='x', padx=10, pady=5)

    tk.Button(frame_files, text="입고전표 선택", command=select_inbound_file, bg="#f0f0f0").grid(row=0, column=0, padx=5)
    lbl_inbound_path = tk.Label(frame_files, text="선택 안됨", fg="gray")
    lbl_inbound_path.grid(row=0, column=1, sticky="w")

    tk.Button(frame_files, text="실사전표 추가", command=add_audit_files, bg="#e0f7fa").grid(row=1, column=0, padx=5, pady=5)
    lbl_audit_path = tk.Label(frame_files, text="선택 안됨", fg="gray")
    lbl_audit_path.grid(row=1, column=1, sticky="w")
    
    tk.Button(frame_files, text="전체 초기화", command=reset_all_files, bg="#ffcdd2").grid(row=0, column=2, rowspan=2, padx=15, sticky="ns")

    frame_btns = tk.Frame(root, padx=10, pady=10)
    frame_btns.pack(fill='x', padx=10)
    frame_btns.columnconfigure((0,1), weight=1)

    tk.Button(frame_btns, text="1. 당일입고 (추가)", command=run_btn1, bg="#e1f5fe", font=("맑은 고딕", 11), height=2).grid(row=0, column=0, sticky="ew", padx=2, pady=2)
    tk.Button(frame_btns, text="2. 자리변경 (교체)", command=run_btn2, bg="#fff9c4", font=("맑은 고딕", 11), height=2).grid(row=0, column=1, sticky="ew", padx=2, pady=2)
    tk.Button(frame_btns, text="3. 재입고 (병합)", command=run_btn3, bg="#ffebee", font=("맑은 고딕", 11), height=2).grid(row=1, column=0, sticky="ew", padx=2, pady=2)
    tk.Button(frame_btns, text="4. 삭제/덮어쓰기", command=run_btn4_overwrite, bg="#e0e0e0", font=("맑은 고딕", 11), height=2).grid(row=1, column=1, sticky="ew", padx=2, pady=2)

    frame_log = tk.LabelFrame(root, text="진단 로그", padx=5, pady=5)
    frame_log.pack(fill='both', expand=True, padx=10, pady=10)
    txt_log = scrolledtext.ScrolledText(frame_log, height=10)
    txt_log.pack(fill='both', expand=True)

    log("프로그램 준비 완료. v28.0 (표지판 제외)")
    root.mainloop()
