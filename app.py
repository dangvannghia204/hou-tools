import os
import tempfile
import shutil
import re
import zipfile
from datetime import datetime
from copy import copy
import pandas as pd
import numpy as np
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
import streamlit as st

try:
    from pyxlsb import open_workbook as open_xlsb
except ImportError:
    pass

# ==========================================
# CẤU HÌNH TRANG & CSS COMPACT (SINGLE-SCREEN)
# ==========================================
st.set_page_config(page_title="Excel Data Workspace", page_icon="📊", layout="wide")

st.markdown("""
    <style>
        .block-container { 
            padding-top: 1rem !important; 
            padding-bottom: 0rem !important; 
            max-width: 1200px;
        }
        .app-header {
            background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
            padding: 1rem 1.5rem;
            border-radius: 8px;
            color: white;
            margin-bottom: 1rem;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
        }
        .app-header h2 { color: white; margin: 0; font-size: 1.6rem; font-weight: 700; padding-bottom: 0.2rem;}
        .app-header p { margin: 0; font-size: 0.95rem; opacity: 0.9; }

        .instruction-card {
            background-color: #F8FAFC;
            border-left: 4px solid #3B82F6;
            padding: 1rem;
            border-radius: 6px;
            color: #334155;
            font-size: 0.95rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            margin-bottom: 1rem;
            min-height: 160px;
        }
        
        [data-testid="stFileUploadDropzone"] {
            min-height: 100px !important;
            padding: 1rem !important;
        }
        [data-testid="stFileUploadDropzone"] div { gap: 0.2rem; }
        
        .stButton>button {
            width: 100%;
            border-radius: 6px;
            height: 45px;
            font-size: 15px;
            font-weight: 600;
            background-color: #2563EB;
            color: white;
            border: none;
            transition: all 0.2s ease-in-out;
        }
        .stButton>button:hover { background-color: #1D4ED8; transform: translateY(-1px); }
        
        .stDownloadButton>button {
            width: 100%;
            border-radius: 6px;
            height: 45px;
            font-size: 15px;
            font-weight: 600;
            background-color: #10B981; 
            color: white;
            border: none;
            transition: all 0.2s ease-in-out;
        }
        .stDownloadButton>button:hover { background-color: #059669; transform: translateY(-1px); }
        
        css-1v0mbdj { margin-top: 0rem; }
        .st-emotion-cache-1kyxreq { gap: 0.5rem; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# CLASS XỬ LÝ LÕI MỚI CHO CHỨC NĂNG KHLM
# ==========================================
class ExcelDataProcessor:
    def __init__(self, input_file="Data_fill.xlsx", keywords_str="DNP, HCM", log_callback=print):
        self.input_file = input_file
        self.keywords_str = keywords_str
        self.log = log_callback
        self.base_dir = os.path.dirname(os.path.abspath(input_file)) if os.path.dirname(input_file) else os.getcwd()

    def run_all_steps(self):
        self.log("BẮT ĐẦU QUÁ TRÌNH XỬ LÝ DỮ LIỆU...")
        if not os.path.exists(self.input_file):
            self.log(f"LỖI: Không tìm thấy file {self.input_file}")
            return False

        self._preprocess_khlm()
        out_khl = os.path.join(self.base_dir, "Ketqua_KHLM.xlsx")
        self._run_source_1(is_hoc_lai=False, output_path=out_khl)
        
        out_hl = os.path.join(self.base_dir, "Ketquahoclai_KHLM.xlsx")
        self._run_source_1(is_hoc_lai=True, output_path=out_hl)
        
        out_merge = os.path.join(self.base_dir, "Data_Merge.xlsx")
        self._merge_results(out_khl, out_hl, out_merge)
        
        out_lopmon = os.path.join(self.base_dir, "LopMon_Ketqua.xlsx")
        self._run_source_2_and_map(out_merge, out_lopmon)
        
        self.log("\nHOÀN THÀNH TOÀN BỘ QUÁ TRÌNH!")
        return True

    def _preprocess_khlm(self):
        wb = load_workbook(self.input_file)
        actual_khlm = next((s for s in wb.sheetnames if s.lower() == 'khlm'), None)
        if not actual_khlm: return

        df_khlm = pd.read_excel(self.input_file, sheet_name=actual_khlm)
        cols = df_khlm.columns.tolist()
        col_map = {str(c).strip().upper(): c for c in cols}
        keywords = [k.strip().upper() for k in self.keywords_str.split(',') if k.strip()]

        for index, row in df_khlm.iterrows():
            ten_lop = str(row.get('TenLop', '')).upper()
            dia_phuong = str(row.get('DiaPhuong', ''))
            dia_phuong_upper = dia_phuong.upper()

            if 'HL' in ten_lop or 'HL' in dia_phuong_upper or 'HỌC LẠI' in dia_phuong_upper:
                df_khlm.at[index, 'DiaPhuongHL'] = 'HL'

            current_dp_khl = str(row.get('DiaPhuongKHL', ''))
            if current_dp_khl.lower() == 'nan': current_dp_khl = ''
            
            added_codes = []
            for kw in keywords:
                if kw in dia_phuong_upper:
                    if kw in col_map: 
                        original_col_name = col_map[kw]
                        val = row[original_col_name]
                        if pd.notna(val) and isinstance(val, (int, float)) and val > 0:
                            added_codes.append(kw)
            
            if added_codes:
                str_to_add = ",".join(added_codes)
                if current_dp_khl.strip():
                    df_khlm.at[index, 'DiaPhuongKHL'] = f"{current_dp_khl},{str_to_add}"
                else:
                    df_khlm.at[index, 'DiaPhuongKHL'] = str_to_add

        with pd.ExcelWriter(self.input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
             df_khlm.to_excel(writer, sheet_name=actual_khlm, index=False)

    def _get_col_idx(self, df, col_name, default_name):
        if col_name in df.columns: return df.columns.get_loc(col_name)
        else: return -1

    def _run_source_1(self, is_hoc_lai, output_path):
        xls = pd.ExcelFile(self.input_file)
        actual_khlm = next((s for s in xls.sheet_names if s.lower() == 'khlm'), None)
        actual_data = next((s for s in xls.sheet_names if s.lower() == 'data'), None)
        if not actual_khlm or not actual_data: return

        khlm = pd.read_excel(self.input_file, sheet_name=actual_khlm, header=0)
        data = pd.read_excel(self.input_file, sheet_name=actual_data)

        idx_tenlop = self._get_col_idx(khlm, 'TenLop', 'TenLop')
        idx_mamon_khlm = self._get_col_idx(khlm, 'MaMon', 'MaMon')
        idx_dp_khl = self._get_col_idx(khlm, 'DiaPhuongKHL', 'DiaPhuongKHL')
        idx_dp_hl = self._get_col_idx(khlm, 'DiaPhuongHL', 'DiaPhuongHL')

        idx_loplt_data = self._get_col_idx(data, 'LopLT', 'LopLT')
        idx_mamon_data = self._get_col_idx(data, 'MaMon', 'MaMon')
        idx_matram_data = self._get_col_idx(data, 'MaTram', 'MaTram')
        idx_msv_data = self._get_col_idx(data, 'MSV', 'MSV')

        assigned = {}
        total_rows = len(khlm)
        filter_keywords = [k.strip().upper() for k in self.keywords_str.split(',') if k.strip()]
        n = idx_dp_hl if is_hoc_lai else idx_dp_khl

        for i in range(total_rows):
            row = khlm.iloc[i]
            kc = str(row.iloc[idx_mamon_khlm]).strip().upper()
            kh = str(row.iloc[idx_tenlop]).strip().upper()
            kj = str(row.iloc[n]).strip().upper()
            
            if kc == 'NAN' or not kc or kc == "": continue

            kc_parts = [p.strip() for p in kc.split('/') if p.strip()]
            mask_mamon = data.iloc[:, idx_mamon_data].astype(str).str.upper().apply(
                lambda x: any((x.strip() in p) or (p in x.strip()) for p in kc_parts)
            )
            mask_matram = data.iloc[:, idx_matram_data].astype(str).str.upper().apply(
                lambda x: (x.strip() in kj) or (kj in x.strip())
            )

            if not is_hoc_lai:
                mask_loplt = data.iloc[:, idx_loplt_data].astype(str).str.upper().apply(
                    lambda x: (x.strip() in kh) or (kh in x.strip())
                )
                mask = mask_mamon & mask_loplt & mask_matram
            else:
                mask = mask_mamon & mask_matram

            msv_col_name = data.columns[idx_msv_data]
            matched_msvs = data.loc[mask, msv_col_name].unique().tolist()

            if any(kw in kj for kw in filter_keywords):
                if kc not in assigned: assigned[kc] = set()
                final_msvs = [m for m in matched_msvs if m not in assigned[kc]]
                assigned[kc].update(final_msvs)
                matched_msvs = final_msvs

            khlm.iloc[i, 4] = ", ".join(map(str, matched_msvs))
            khlm.iloc[i, 5] = len(matched_msvs)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            khlm.to_excel(writer, sheet_name='Result', index=False)

    def _merge_results(self, file_khl, file_hl, file_merge_out):
        xls1 = pd.ExcelFile(file_khl)
        actual_res1 = next((s for s in xls1.sheet_names if s.lower() == 'result'), None)
        xls2 = pd.ExcelFile(file_hl)
        actual_res2 = next((s for s in xls2.sheet_names if s.lower() == 'result'), None)
        
        if not actual_res1 or not actual_res2: return

        df1 = pd.read_excel(file_khl, sheet_name=actual_res1, usecols=[0, 1, 2, 3, 4])
        df2 = pd.read_excel(file_hl, sheet_name=actual_res2, usecols=[0, 1, 2, 3, 4])

        df1_valid = df1[df1[df1.columns[4]].astype(str).str.strip().replace('nan', '') != '']
        df2_valid = df2[df2[df2.columns[4]].astype(str).str.strip().replace('nan', '') != '']

        df_merged = pd.concat([df1_valid, df2_valid], ignore_index=True)
        df_merged.to_excel(file_merge_out, sheet_name='data', index=False)

    def _run_source_2_and_map(self, file_merge_in, file_lopmon_out):
        wb = load_workbook(file_merge_in)
        data_sheet_name = next((s for s in wb.sheetnames if s.lower() == 'data'), None)
        if not data_sheet_name: return
        ws_data = wb[data_sheet_name]
        
        actual_res = next((s for s in wb.sheetnames if s.lower() == 'result'), None)
        if actual_res: del wb[actual_res]
        ws_result = wb.create_sheet('Result')
        ws_result.append(["Cột A (Mã tách)", "Cột B (Từ A)", "Cột C (Từ B)", "Cột D (Từ C)", "Cột E (Từ D)", "Cột F (Trống)", "Cột G (A+D)", "Cột H (Copy B)"])
        result_data_for_mapping = {}

        for row in range(2, ws_data.max_row + 1):
            val_a = ws_data.cell(row=row, column=1).value
            val_b = ws_data.cell(row=row, column=2).value
            val_c = ws_data.cell(row=row, column=3).value
            val_d = ws_data.cell(row=row, column=4).value
            val_e = ws_data.cell(row=row, column=5).value

            if val_a is None and val_b is None and val_e is None: continue

            if val_e is not None:
                for item in str(val_e).split(','):
                    clean_item = item.strip()
                    if clean_item:
                        col_g = str(clean_item) + str(val_c if val_c is not None else "")
                        col_h = val_a
                        ws_result.append([clean_item, val_a, val_b, val_c, val_d, "", col_g, col_h])
                        result_data_for_mapping[col_g] = col_h
                        if val_c and '/' in str(val_c):
                            for part in str(val_c).split('/'):
                                part_clean = part.strip()
                                if part_clean: result_data_for_mapping[str(clean_item) + part_clean] = col_h
            else:
                 col_g = "" + str(val_c if val_c is not None else "")
                 col_h = val_a
                 ws_result.append(["", val_a, val_b, val_c, val_d, "", col_g, col_h])
                 result_data_for_mapping[col_g] = col_h
                 if val_c and '/' in str(val_c):
                     for part in str(val_c).split('/'):
                         part_clean = part.strip()
                         if part_clean: result_data_for_mapping["" + part_clean] = col_h

        wb.save(file_lopmon_out)

        wb_fill = load_workbook(self.input_file)
        actual_fill_data = next((s for s in wb_fill.sheetnames if s.lower() == 'data'), None)
        if not actual_fill_data: return
        ws_fill_data = wb_fill[actual_fill_data]
        
        header_row = 1
        col_idx_msv, col_idx_mamon, col_idx_macoursett = -1, -1, -1
        for cell in ws_fill_data[header_row]:
            val = str(cell.value).strip().upper() if cell.value else ""
            if val == 'MSV': col_idx_msv = cell.column
            elif val == 'MAMON': col_idx_mamon = cell.column
            elif val == 'MACOURSETT': col_idx_macoursett = cell.column

        if col_idx_msv == -1 or col_idx_mamon == -1: return
        if col_idx_macoursett == -1:
            col_idx_macoursett = ws_fill_data.max_column + 1
            ws_fill_data.cell(row=header_row, column=col_idx_macoursett, value="MaCourseTT")

        for row in range(2, ws_fill_data.max_row + 1):
            msv = str(ws_fill_data.cell(row=row, column=col_idx_msv).value or "").strip()
            mamon = str(ws_fill_data.cell(row=row, column=col_idx_mamon).value or "").strip()
            lookup_key = msv + mamon 
            if lookup_key in result_data_for_mapping:
                ws_fill_data.cell(row=row, column=col_idx_macoursett, value=result_data_for_mapping[lookup_key])
            elif '/' in mamon:
                for part in mamon.split('/'):
                    part_clean = part.strip()
                    alt_key = msv + part_clean
                    if alt_key in result_data_for_mapping:
                        ws_fill_data.cell(row=row, column=col_idx_macoursett, value=result_data_for_mapping[alt_key])
                        break 
        
        wb_fill.save(os.path.join(self.base_dir, "Data_fill_Finish.xlsx"))

# ==========================================
# CÁC HÀM TIỆN ÍCH CHUNG
# ==========================================
def apply_full_border(ws):
    thin = Side(border_style="thin", color="000000")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row: cell.border = border

def auto_fit_columns(ws, padding=2):
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    length = len(str(cell.value))
                    if length > max_length: max_length = length
            except: pass
        ws.column_dimensions[column].width = max(10, min(max_length + padding, 50))

def roman_to_int(r):
    m = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10}
    return m.get(str(r).upper(), r)

def format_hk(a6):
    if not a6: return ""
    try:
        rp = re.search(r'Học kỳ\s+([I|V|X]+)', str(a6), re.IGNORECASE)
        hn = roman_to_int(rp.group(1)) if rp else ""
        yp = re.search(r'(\d{4})-(\d{4})', str(a6))
        if yp: return f"HK{hn}_{yp.group(1)[2:]}.{yp.group(2)[2:]}"
        return f"HK{hn}"
    except: return str(a6)

def format_excel_date(cell):
    v = cell.value
    return v.strftime("%d/%m/%Y") if isinstance(v, datetime) else (str(v) if v else "")

# ==========================================
# CÁC HÀM LOGIC XỬ LÝ (Streamlit Wrapper)
# ==========================================
def fill_khlm_logic(folder_path, keywords_str):
    target_file = ""
    for f in os.listdir(folder_path):
        if f.endswith(".xlsx") and not f.startswith("~$"):
            try:
                temp_wb = load_workbook(os.path.join(folder_path, f), read_only=True)
                lower_sheets = [s.lower() for s in temp_wb.sheetnames]
                if 'khlm' in lower_sheets and 'data' in lower_sheets:
                    target_file = os.path.join(folder_path, f)
                    break
            except: continue
            
    if not target_file: 
        raise ValueError("Không tìm thấy file Excel nào chứa đủ 2 sheet 'KHLM' và 'data'!")

    # Thực thi Core Logic cập nhật
    processor = ExcelDataProcessor(
        input_file=target_file, 
        keywords_str=keywords_str, 
        log_callback=lambda msg: None 
    )
    success = processor.run_all_steps()
    
    if not success:
        raise ValueError("Xử lý thất bại. Vui lòng kiểm tra lại định dạng file.")
        
    # Tạo file Zip nén tất cả file Output
    zip_path = os.path.join(folder_path, "KetQua_KHLM_TongHop.zip")
    output_files = [
        "Ketqua_KHLM.xlsx", "Ketquahoclai_KHLM.xlsx", 
        "Data_Merge.xlsx", "LopMon_Ketqua.xlsx", "Data_fill_Finish.xlsx"
    ]
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for fname in output_files:
            fpath = os.path.join(folder_path, fname)
            if os.path.exists(fpath):
                zipf.write(fpath, fname)
                
    return zip_path

def extract_courses_logic(folder_path):
    files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and not f.startswith('~$') and f != "Courses_List.xlsx"]
    if not files: raise ValueError("Không có file Excel nào để xử lý!")
    output_file = os.path.join(folder_path, "Courses_List.xlsx")
    wb_out = Workbook(); wb_out.remove(wb_out.active)
    for idx, file_name in enumerate(files):
        try:
            parts = file_name.split('- ')
            major_name = parts[1].split('.')[0].strip() if len(parts) > 1 else file_name.split('.')[0].strip()
        except: major_name = file_name[:25] 
        ws_major = wb_out.create_sheet(title=major_name[:30])
        unique_courses = set()
        wb_src = load_workbook(os.path.join(folder_path, file_name), data_only=True)
        for sheet_name in wb_src.sheetnames:
            ws_src = wb_src[sheet_name]
            for r in range(16, ws_src.max_row + 1):
                val_a = str(ws_src.cell(row=r, column=1).value or "").strip().lower()
                if "lưu ý:" in val_a: break
                col_b, col_c, col_d = ws_src.cell(row=r, column=2).value, ws_src.cell(row=r, column=3).value, ws_src.cell(row=r, column=4).value 
                if col_b and col_c: unique_courses.add(f"{str(col_b).strip()}_{str(col_c).strip()}_{str(col_d or 0).strip()} Tín chỉ")
        sorted_list = sorted(list(unique_courses))
        for i, course in enumerate(sorted_list, 1): ws_major.cell(row=i, column=1, value=course)
        apply_full_border(ws_major); auto_fit_columns(ws_major)
    wb_out.save(output_file)
    return output_file

def fill_sllm_logic(folder_path):
    file_path = os.path.join(folder_path, "Data_SLLM.xlsx")
    if not os.path.exists(file_path): raise ValueError("Không tìm thấy tệp Data_SLLM.xlsx!")
    wb = load_workbook(file_path)
    if "Data" not in wb.sheetnames or "ThongKe" not in wb.sheetnames: raise ValueError("Tệp Data_SLLM.xlsx phải chứa sheet 'Data' và 'ThongKe'!")
    ws_data = wb["Data"]
    data_map = {}
    for r in range(2, ws_data.max_row + 1):
        cell_k = ws_data.cell(row=r, column=11).value
        val_k = str(cell_k or "").strip()
        region = "HCM" if len(val_k) >= 4 and val_k[1:4].upper() == "DNV" else ("DNP" if len(val_k) >= 4 and val_k[1:4].upper() == "DNP" else "")
        ws_data.cell(row=r, column=12, value=region)
        ma_lm, ma_hp = str(ws_data.cell(row=r, column=4).value or "").strip().lower(), str(ws_data.cell(row=r, column=8).value or "").strip().lower()
        if ma_lm and ma_hp:
            key = (ma_hp, ma_lm)
            if key not in data_map: data_map[key] = {"DNP": 0, "HCM": 0}
            if region in ["DNP", "HCM"]: data_map[key][region] += 1
    ws_tk = wb["ThongKe"]
    for r in range(2, ws_tk.max_row + 1):
        ma_hp_tk, ma_lm_str = str(ws_tk.cell(row=r, column=2).value or "").strip().lower(), str(ws_tk.cell(row=r, column=1).value or "").strip()
        dnp_sum, hcm_sum = 0, 0
        if ma_hp_tk and ma_lm_str:
            codes = [c.strip().lower() for c in ma_lm_str.split(',')]
            for code in codes:
                key = (ma_hp_tk, code)
                if key in data_map: dnp_sum += data_map[key]["DNP"]; hcm_sum += data_map[key]["HCM"]
        ws_tk.cell(row=r, column=3, value=dnp_sum); ws_tk.cell(row=r, column=4, value=hcm_sum)
    for sheet in [ws_data, ws_tk]: apply_full_border(sheet); auto_fit_columns(sheet)
    wb.save(file_path)
    return file_path

def compare_data_logic(folder_path):
    file_xlsb, file_src, output_file = os.path.join(folder_path, "Data.xlsb"), os.path.join(folder_path, "Data_Source.xlsx"), os.path.join(folder_path, "Compared_Result.xlsx")
    if not os.path.exists(file_xlsb) or not os.path.exists(file_src): raise ValueError("Không tìm thấy tệp Data.xlsb hoặc Data_Source.xlsx!")
    wb_src = load_workbook(file_src, data_only=True)
    if "DSSV" not in wb_src.sheetnames: raise ValueError("Thiếu sheet DSSV trong Data_Source.xlsx")
    ws_src = wb_src["DSSV"]
    source_codes = {str(ws_src.cell(row=r, column=10).value).strip().lower() for r in range(2, ws_src.max_row + 1) if ws_src.cell(row=r, column=10).value}
    wb_out = Workbook(); ws_out = wb_out.active; ws_out.title = "Result"; headers_added = False
    with open_xlsb(file_xlsb) as wb_bin:
        with wb_bin.get_sheet(1) as sheet:
            for row in sheet.rows():
                val_a = str(row[0].v).strip().lower() if row[0].v is not None else ""
                if not headers_added:
                    ws_out.append([c.v for c in row]); headers_added = True; continue
                if val_a in source_codes: ws_out.append([c.v for c in row])
    apply_full_border(ws_out); auto_fit_columns(ws_out); wb_out.save(output_file)
    return output_file

def filter_sv_logic(folder_path):
    f_dk, f_src = os.path.join(folder_path, "DanhSachDangKy.xlsx"), os.path.join(folder_path, "Data_Source.xlsx")
    if not os.path.exists(f_dk) or not os.path.exists(f_src): raise ValueError("Thiếu file DanhSachDangKy.xlsx hoặc Data_Source.xlsx")
    wb_src = load_workbook(f_src, data_only=True); ws_src = wb_src["DSSV"]
    sv_set = {str(ws_src.cell(row=r, column=4).value).strip().lower() for r in range(2, ws_src.max_row + 1) if ws_src.cell(row=r, column=4).value}
    wb_dk = load_workbook(f_dk, data_only=True)
    reg_data, learn_keys, h_reg, h_learn = [], set(), [], []
    if "Dangky" in wb_dk.sheetnames:
        ws = wb_dk["Dangky"]; h_reg = [c.value for c in ws[1]]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[1] and str(r[1]).strip().lower() in sv_set: reg_data.append(r)
    if "Danghoc" in wb_dk.sheetnames:
        ws = wb_dk["Danghoc"]; h_learn = [c.value for c in ws[1]]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[1] and str(r[1]).strip().lower() in sv_set: learn_keys.add((str(r[1] or "").strip().lower(), str(r[6] or "").strip().lower(), str(r[14] or "").strip().lower()))
    m_rows, n_rows = [], []
    for r in reg_data:
        rk = (str(r[1] or "").strip().lower(), str(r[2] or "").strip().lower(), str(r[19] or "").strip().lower())
        if rk in learn_keys: m_rows.append(r)
        else: n_rows.append(r)
    output_file = os.path.join(folder_path, "Filter_Result.xlsx"); wb_res = Workbook()
    configs = [("Register_List", reg_data, h_reg), ("matching", m_rows, h_reg), ("not_matching", n_rows, h_reg)]
    for i, (t, d, h) in enumerate(configs):
        ws = wb_res.active if i == 0 else wb_res.create_sheet(t); ws.title = t; ws.append(h)
        for row in d: ws.append(row)
        apply_full_border(ws); auto_fit_columns(ws)
    wb_res.save(output_file)
    return output_file

def export_khhtct_logic(folder_path):
    tf = os.path.join(folder_path, "Merged_GK300.xlsx")
    if not os.path.exists(tf): raise ValueError("Cần file Merged_GK300.xlsx")
    wb = load_workbook(tf); ws_d = wb["DSSV_GK300"]; ws_k = wb["KHHT_GK300"]; kd = {}
    for r in range(2, ws_k.max_row + 1):
        k, m = ws_k.cell(row=r, column=1).value, ws_k.cell(row=r, column=4).value 
        if k and m:
            sk = str(k).strip().lower()
            if sk not in kd: kd[sk] = []
            kd[sk].append(m)
    if "KHHTCT_GK300" in wb.sheetnames: del wb["KHHTCT_GK300"]
    ws_n = wb.create_sheet("KHHTCT_GK300")
    hd = ["STT", "Lớp", "Tài khoản SV", "Họ và đệm", "Tên", "Ngày sinh", "Mã sinh viên", "Mã lớp môn"]
    for c, h in enumerate(hd, 1): ws_n.cell(row=1, column=c, value=h).font = Font(bold=True)
    cr, stt = 2, 1
    for r in range(2, ws_d.max_row + 1):
        l = ws_d.cell(row=r, column=17).value
        if l:
            sl = str(l).strip()
            if len(sl) >= 4:
                mk = (sl[0] + sl[-3:]).lower()
                if mk in kd:
                    for ml in kd[mk]:
                        dr = [stt, sl, ws_d.cell(row=r, column=4).value, ws_d.cell(row=r, column=12).value, ws_d.cell(row=r, column=13).value, ws_d.cell(row=r, column=14).value, ws_d.cell(row=r, column=10).value, ml]
                        for c, v in enumerate(dr, 1): ws_n.cell(row=cr, column=c, value=v)
                        cr += 1; stt += 1
    apply_full_border(ws_n); auto_fit_columns(ws_n); wb.save(tf)
    return tf

def export_dssv_logic(folder_path):
    sf, df = os.path.join(folder_path, "Data_Source.xlsx"), os.path.join(folder_path, "Merged_GK300.xlsx")
    if not os.path.exists(sf) or not os.path.exists(df): raise ValueError("Cần file Data_Source.xlsx và Merged_GK300.xlsx")
    wb_d = load_workbook(df); ws_k = wb_d["KHHT_GK300"]; cs = {str(ws_k.cell(row=r, column=12).value).strip().lower() for r in range(2, ws_k.max_row + 1) if ws_k.cell(row=r, column=12).value} 
    wb_s = load_workbook(sf, data_only=True); ws_s = wb_s["DSSV"]
    if "DSSV_GK300" in wb_d.sheetnames: del wb_d["DSSV_GK300"]
    ws_n = wb_d.create_sheet("DSSV_GK300")
    for c in range(1, ws_s.max_column + 1): ws_n.cell(row=1, column=c, value=ws_s.cell(row=1, column=c).value).font = Font(bold=True)
    cr = 2
    for r in range(2, ws_s.max_row + 1):
        vu = ws_s.cell(row=r, column=17).value
        if vu:
            su = str(vu).strip()
            if len(su) >= 4:
                mk = (su[0] + su[-3:]).lower()
                if mk in cs:
                    for c in range(1, ws_s.max_column + 1): ws_n.cell(row=cr, column=c, value=ws_s.cell(row=r, column=c).value)
                    cr += 1
    apply_full_border(ws_n); auto_fit_columns(ws_n); wb_d.save(df)
    return df

def export_khht_logic(folder_path):
    fs = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and not f.startswith('~$') and f not in ["Merged_GK300.xlsx", "Data_Source.xlsx", "ClassName.xlsx", "Filter_Result.xlsx", "DanhSachDangKy.xlsx"]]
    if not fs: raise ValueError("Không có file dữ liệu GK300!")
    op = os.path.join(folder_path, "Merged_GK300.xlsx"); wb = Workbook(); ws = wb.active; ws.title = "KHHT_GK300"
    hd = ["Mã LT", "HK", "Tên HP", "Mã HP", "Số TC", "Ngày bắt đầu (Dự kiến)", "Ngày kết thúc (Dự kiến)", "Mã ngành", "Khóa", "Ghi chú", "", "Mã LT (Duy nhất)"]
    for c, h in enumerate(hd, 1): ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    mlt = []
    for f in fs:
        swb = load_workbook(os.path.join(folder_path, f), data_only=True)
        for sn in swb.sheetnames:
            s = swb[sn]; mlt.append(sn); hk, bd, kt = format_hk(s["A6"].value), format_excel_date(s["J11"]), format_excel_date(s["X11"])
            for r in range(13, s.max_row + 1):
                t, m, tc = s.cell(row=r, column=2).value, s.cell(row=r, column=3).value, s.cell(row=r, column=4).value
                if t and m: ws.append([sn, hk, t, m, tc, bd, kt, str(sn)[0], str(sn)[-2:], ""])
    unq = sorted(list(set(mlt)))
    for i, v in enumerate(unq, 2): ws.cell(row=i, column=12, value=v)
    apply_full_border(ws); auto_fit_columns(ws); wb.save(op)
    return op

def process_files_logic(folder_path):
    fs = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and not f.startswith('~$') and f not in ["File_Merged.xlsx", "ClassName.xlsx", "Merged_GK300.xlsx", "Data_Source.xlsx", "Filter_Result.xlsx", "DanhSachDangKy.xlsx"]]
    if not fs: raise ValueError("Không có file để gộp!")
    op = os.path.join(folder_path, "File_Merged.xlsx"); mwb = Workbook(); dws = mwb.active; dws.title = "Sheet_Merged"
    yl = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid"); cdr, mcw = 1, {}
    for f in fs:
        swb = load_workbook(os.path.join(folder_path, f), data_only=False)
        for sn in swb.sheetnames:
            sws = swb[sn]
            for cd in sws.column_dimensions.values():
                for ci in range(cd.min, cd.max + 1):
                    if ci <= 36: l = get_column_letter(ci); mcw[l] = max(mcw.get(l, 0), cd.width or 0)
            hrm = {}
            for hr in [8,9,10,11,12]:
                hrm[hr] = cdr
                if sws.row_dimensions[hr].height: dws.row_dimensions[cdr].height = sws.row_dimensions[hr].height
                for c in range(1, 37):
                    sc, dc = sws.cell(row=hr, column=c), dws.cell(row=cdr, column=c, value=sws.cell(row=hr, column=c).value)
                    if sc.has_style: dc.font, dc.border, dc.fill, dc.number_format, dc.alignment = copy(sc.font), copy(sc.border), copy(sc.fill), copy(sc.number_format), copy(sc.alignment)
                    if hr == 8: dc.fill = yl
                cdr += 1
            vdr = [r for r in range(13, sws.max_row + 1) if sws.cell(row=r, column=10).value and "-" in str(sws.cell(row=r, column=10).value).lower() and not sws.row_dimensions[r].hidden]
            if not vdr: continue
            brm = {}
            for sr in vdr:
                brm[sr] = cdr
                if sws.row_dimensions[sr].height: dws.row_dimensions[cdr].height = sws.row_dimensions[sr].height
                for c in range(1, 37):
                    sc, dc = sws.cell(row=sr, column=c), dws.cell(row=cdr, column=c, value=sws.cell(row=sr, column=c).value)
                    if sc.has_style: dc.font, dc.border, dc.fill, dc.number_format, dc.alignment = copy(sc.font), copy(sc.border), copy(sc.fill), copy(sc.number_format), copy(sc.alignment)
                dws.cell(row=cdr, column=37, value=f"{f} -> {sn}"); cdr += 1
            for mr in sws.merged_cells.ranges:
                rir = [r for r in vdr if mr.min_row <= r <= mr.max_row]
                if len(rir) >= 1:
                    mnr, mxr = brm[min(rir)], brm[max(rir)]; mnc, mxc = max(mr.min_col, 1), min(mr.max_col, 36)
                    if mnc <= mxc and (mxr > mnr or mxc > mnc):
                        try: dws.merge_cells(start_row=mnr, start_column=mnc, end_row=mxr, end_column=mxc)
                        except: pass
    for l, w in mcw.items(): dws.column_dimensions[l].width = w
    dws.column_dimensions[get_column_letter(37)].width = 30
    apply_full_border(dws); mwb.save(op)
    return op

def extract_class_names_logic(folder_path):
    fs = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and not f.startswith('~$') and f not in ["Merged_GK300.xlsx", "ClassName.xlsx", "Filter_Result.xlsx"]]
    if not fs: raise ValueError("Không có file để xử lý!")
    op = os.path.join(folder_path, "ClassName.xlsx"); wb = Workbook(); ws = wb.active; cr = 1
    for f in fs:
        swb = load_workbook(os.path.join(folder_path, f), read_only=True)
        for sn in swb.sheetnames: ws.cell(row=cr, column=1, value=sn); cr += 1
    apply_full_border(ws); auto_fit_columns(ws); wb.save(op)
    return op

# ==========================================
# MENU SIDEBAR (CẬP NHẬT TÊN CHỨC NĂNG)
# ==========================================
with st.sidebar:
    st.markdown("<h3 style='color: #1E3A8A; font-weight: 700; margin-top: -15px;'>DATA WORKSPACE</h3>", unsafe_allow_html=True)
    st.markdown("---")
    
    menu_options = {
        "Gộp File Nguồn": "Gộp File Nguồn",
        "Điền KHLM (Updated) (*)": "Điền KHLM (Updated) (*)",
        "Lọc KQHT Sinh viên": "Lọc kết quả học tập của sinh viên",
        "Lọc SV Học lại/Cải thiện": "Lọc sinh viên học lại & học cải thiện",
        "Xuất Mã lớp (GK300)": "Xuất Mã lớp theo GK300",
        "Xuất KHHT (1*)": "Xuất KHHT theo GK300 (1*)",
        "Xuất DSSV (2*)": "Xuất DSSV theo GK300 (2*)",
        "Xuất KHHTCT (3*)": "Xuất KHHTCT theo GK300 (3*)",
        "Thống kê Lớp & Môn (*)": "Thống kê số lượng theo lớp/nhóm lớp & môn (*)",
        "Xuất Môn theo Ngành (*)": "Xuất môn theo ngành học (*)"
    }
    
    selected_label = st.radio("CHỌN CHỨC NĂNG BÊN DƯỚI", list(menu_options.keys()))
    choice = menu_options[selected_label]
    
    st.markdown("---")
    st.caption("Ver 5.0 | Advanced Processor Edition")

# ==========================================
# GIAO DIỆN CHÍNH (SINGLE-SCREEN GRID LAYOUT)
# ==========================================
st.markdown(f"""
    <div class="app-header">
        <h2>{choice}</h2>
        <p>Thực thi tự động tác vụ Excel một cách nhanh chóng và chính xác.</p>
    </div>
""", unsafe_allow_html=True)

# Khai báo từ điển hướng dẫn
instructions = {
    "Gộp File Nguồn": "Đầu vào là file <b>GK300</b> của 1 hoặc nhiều khóa (Mỗi sheet chứa bảng đăng ký).",
    "Điền KHLM (Updated) (*)": "<b>Yêu cầu file:</b> <code>Data_fill.xlsx</code><br><br><b>Sheet KHLM:</b> Cần có các cột <code>TenLop</code>, <code>MaMon</code>, <code>DiaPhuongKHL</code>, <code>DiaPhuongHL</code>.<br><b>Sheet Data:</b> Cần có các cột <code>LopLT</code>, <code>MaMon</code>, <code>MaTram</code>, <code>MSV</code>.",
    "Lọc kết quả học tập của sinh viên": "<b>Yêu cầu:</b> <code>Data_Source.xlsx</code> & <code>Data.xlsb</code><br><br><b>Data_Source.xlsx:</b> Sheet 'DSSV', Cột Q (Lớp), Cột J (Mã SV), Cột D (TK SV).<br><b>Data.xlsb:</b> File nhị phân, Cột A (Mã SV).",
    "Lọc sinh viên học lại & học cải thiện": "<b>Yêu cầu:</b> <code>Data_Source.xlsx</code> & <code>DanhSachDangKy.xlsx</code><br><br><b>DanhSachDangKy.xlsx:</b> Sheet 'Dangky: B,C,T' và 'Danghoc: B,G,O', Cột B (TK SV), Cột C&G (Mã môn), Cột T&O (Số TC).",
    "Xuất Mã lớp theo GK300": "Đầu vào là file <b>GK300</b> của 1 hoặc nhiều khóa.",
    "Xuất KHHT theo GK300 (1*)": "Đầu vào là file <b>GK300</b> của 1 hoặc nhiều khóa.",
    "Xuất DSSV theo GK300 (2*)": "<b>Yêu cầu:</b> <code>Merged_GK300.xlsx</code> tạo từ (1*) & <code>Data_Source.xlsx</code><br><b>Merged_GK300.xlsx:</b> Sheet 'KHHT_GK300', Cột L(12) là Mã LT.",
    "Xuất KHHTCT theo GK300 (3*)": "<b>Yêu cầu:</b> Cần cung cấp file <code>Merged_GK300.xlsx</code>.",
    "Thống kê số lượng theo lớp/nhóm lớp & môn (*)": "<b>Yêu cầu:</b> File <code>Data_SLLM.xlsx</code><br><b>Sheet Data:</b> Cột L (Mã trạm).<br><b>Sheet ThongKe:</b> Cột A (Tên lớp), B (Mã môn), Tiêu đề Cột C (Mã Trạm).",
    "Xuất môn theo ngành học (*)": "Đầu vào là file dữ liệu môn học phân bổ theo ngành và theo khóa."
}

col_info, col_action = st.columns([1.1, 1], gap="medium")

# --- Cột Trái: Thông tin Hướng dẫn & Cấu hình ---
with col_info:
    st.markdown(f"""
        <div class="instruction-card">
            <strong>📌 YÊU CẦU DỮ LIỆU ĐẦU VÀO:</strong><br><br>
            {instructions.get(choice, "Vui lòng upload các file Excel theo đúng định dạng.")}
        </div>
    """, unsafe_allow_html=True)
    
    keywords_str = ""
    if choice == "Điền KHLM (Updated) (*)":
        keywords_str = st.text_input("Thiết lập điều kiện (Từ khóa địa phương):", value="DNP, ĐN ( học tại HCM)")

# --- Cột Phải: Uploader & Button Xử lý ---
with col_action:
    uploaded_files = st.file_uploader("Kéo thả các file Excel vào đây", accept_multiple_files=True, type=['xlsx', 'xlsb'])
    
    if st.button("🚀 XỬ LÝ DỮ LIỆU"):
        if not uploaded_files:
            st.error("⚠️ Vui lòng tải dữ liệu lên trước!")
        else:
            temp_dir = tempfile.mkdtemp()
            try:
                for uploaded_file in uploaded_files:
                    with open(os.path.join(temp_dir, uploaded_file.name), "wb") as f:
                        f.write(uploaded_file.getbuffer())
                
                with st.spinner('⚙️ Đang xử lý...'):
                    result_file = None
                    if choice == "Gộp File Nguồn": result_file = process_files_logic(temp_dir)
                    elif choice == "Điền KHLM (Updated) (*)": result_file = fill_khlm_logic(temp_dir, keywords_str)
                    elif choice == "Lọc kết quả học tập của sinh viên": result_file = compare_data_logic(temp_dir)
                    elif choice == "Lọc sinh viên học lại & học cải thiện": result_file = filter_sv_logic(temp_dir)
                    elif choice == "Xuất Mã lớp theo GK300": result_file = extract_class_names_logic(temp_dir)
                    elif choice == "Xuất KHHT theo GK300 (1*)": result_file = export_khht_logic(temp_dir)
                    elif choice == "Xuất DSSV theo GK300 (2*)": result_file = export_dssv_logic(temp_dir)
                    elif choice == "Xuất KHHTCT theo GK300 (3*)": result_file = export_khhtct_logic(temp_dir)
                    elif choice == "Thống kê số lượng theo lớp/nhóm lớp & môn (*)": result_file = fill_sllm_logic(temp_dir)
                    elif choice == "Xuất môn theo ngành học (*)": result_file = extract_courses_logic(temp_dir)

                    if result_file and os.path.exists(result_file):
                        st.toast('Xử lý thành công!', icon='✅')
                        with open(result_file, "rb") as f:
                            file_data = f.read()
                        
                        file_ext = os.path.splitext(result_file)[1].lower()
                        mime_type = "application/zip" if file_ext == '.zip' else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        
                        st.download_button(
                            label="⬇️ TẢI FILE KẾT QUẢ VỀ MÁY",
                            data=file_data,
                            file_name=os.path.basename(result_file),
                            mime=mime_type
                        )
                    else:
                        st.error("❌ Xử lý thất bại. Kiểm tra cấu trúc file.")
                        
            except Exception as e:
                st.error(f"❌ Lỗi hệ thống: `{str(e)}`")
            finally:
                shutil.rmtree(temp_dir)
