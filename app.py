import os
import tempfile
import shutil
import re
from datetime import datetime
from copy import copy
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
import streamlit as st

try:
    from pyxlsb import open_workbook as open_xlsb
except ImportError:
    pass

# ==========================================
# CẤU HÌNH TRANG & CSS NÂNG CAO (UI/UX)
# ==========================================
st.set_page_config(page_title="Excel Data Workspace", page_icon="📊", layout="wide")

st.markdown("""
    <style>
        /* Tối ưu không gian hiển thị chính */
        .block-container { 
            padding-top: 2rem; 
            padding-bottom: 2rem; 
            max-width: 1000px; /* Căn giữa nội dung, không bị tràn viền quá rộng */
        }
        
        /* Banner Header phong cách SaaS */
        .app-header {
            background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
            padding: 1.5rem 2rem;
            border-radius: 12px;
            color: white;
            margin-bottom: 2rem;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
        }
        .app-header h1 { color: white; margin: 0; font-size: 2.2rem; font-weight: 700; padding-bottom: 0.5rem;}
        .app-header p { margin: 0; font-size: 1.1rem; opacity: 0.9; }

        /* Khối Hướng dẫn (Instruction Card) */
        .instruction-card {
            background-color: #F8FAFC;
            border-left: 5px solid #3B82F6;
            padding: 1.2rem 1.5rem;
            border-radius: 6px;
            margin-bottom: 2rem;
            color: #334155;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }
        
        /* Tùy chỉnh Nút bấm xử lý (Màu xanh dương) */
        .stButton>button {
            width: 100%;
            border-radius: 8px;
            height: 50px;
            font-size: 16px;
            font-weight: 600;
            background-color: #2563EB;
            color: white;
            border: none;
            box-shadow: 0 4px 6px rgba(37, 99, 235, 0.2);
            transition: all 0.2s ease-in-out;
        }
        .stButton>button:hover {
            background-color: #1D4ED8;
            box-shadow: 0 6px 8px rgba(37, 99, 235, 0.3);
            transform: translateY(-2px);
        }
        
        /* Tùy chỉnh Nút Tải xuống (Màu xanh lá) */
        .stDownloadButton>button {
            width: 100%;
            border-radius: 8px;
            height: 50px;
            font-size: 16px;
            font-weight: 600;
            background-color: #10B981; 
            color: white;
            border: none;
            box-shadow: 0 4px 6px rgba(16, 185, 129, 0.2);
            transition: all 0.2s ease-in-out;
            margin-top: 1rem;
        }
        .stDownloadButton>button:hover {
            background-color: #059669;
            box-shadow: 0 6px 8px rgba(16, 185, 129, 0.3);
            transform: translateY(-2px);
        }
        
        /* Ẩn bớt các viền không cần thiết của Streamlit */
        css-1v0mbdj { margin-top: 1rem; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# CÁC HÀM TIỆN ÍCH
# ==========================================
def apply_full_border(ws):
    thin = Side(border_style="thin", color="000000")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = border

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
# CÁC HÀM LOGIC XỬ LÝ LÕI
# ==========================================
def fill_khlm_logic(folder_path, keywords_str):
    target_file = ""
    for f in os.listdir(folder_path):
        if f.endswith(".xlsx") and not f.startswith("~$"):
            try:
                temp_wb = load_workbook(os.path.join(folder_path, f), read_only=True)
                if "KHLM" in temp_wb.sheetnames and "data" in temp_wb.sheetnames:
                    target_file = os.path.join(folder_path, f)
                    break
            except: continue
    if not target_file:
        raise ValueError("Không tìm thấy file Excel nào chứa đủ 2 sheet 'KHLM' và 'data'!")

    filter_keywords = [k.strip().upper() for k in keywords_str.split(',') if k.strip()]
    khlm = pd.read_excel(target_file, sheet_name='KHLM', header=0)
    data = pd.read_excel(target_file, sheet_name='data')
    assigned = {}
    total_rows = len(khlm)
    
    for i in range(total_rows):
        row = khlm.iloc[i]
        kc = str(row.iloc[2]).strip().upper() 
        ki = str(row.iloc[8]).strip().upper() 
        if kc == 'NAN' or not kc or kc == "": continue
        
        mask = (data.iloc[:, 10].astype(str).str.upper().apply(lambda x: (x in kc) or (kc in x))) & \
               (data.iloc[:, 11].astype(str).str.upper().apply(lambda x: (x in ki) or (ki in x)))
        msv_col_name = data.columns[2]
        matched_msvs = data.loc[mask, msv_col_name].unique().tolist()
        
        if any(kw in ki for kw in filter_keywords):
            if kc not in assigned: assigned[kc] = set()
            final_msvs = [m for m in matched_msvs if m not in assigned[kc]]
            assigned[kc].update(final_msvs)
            matched_msvs = final_msvs
        
        khlm.iloc[i, 4] = ", ".join(map(str, matched_msvs)) 
        khlm.iloc[i, 5] = len(matched_msvs) 

    col_k_data = data.iloc[:, 10].astype(str).str.strip().str.upper()
    mapping_ab = col_k_data.value_counts().reset_index(); mapping_ab.columns = ['A', 'B']; mapping_ab = mapping_ab.sort_values(by='A')
    khlm.iloc[:, 5] = pd.to_numeric(khlm.iloc[:, 5], errors='coerce').fillna(0)
    df_cd = khlm.iloc[:, [2, 5]].copy(); df_cd.columns = ['C', 'D']; df_cd['C'] = df_cd['C'].astype(str).str.strip().str.upper()
    df_cd = df_cd[(df_cd['C'] != 'NAN') & (df_cd['D'] > 0)].sort_values(by='C')
    mapping_ef = df_cd.groupby('C')['D'].sum().reset_index(); mapping_ef.columns = ['E', 'F']; mapping_ef = mapping_ef.sort_values(by='E')

    max_len = max(len(mapping_ab), len(df_cd), len(mapping_ef), len(khlm))
    mapping_final = pd.DataFrame(index=range(max_len))
    mapping_final['Mã môn (Data)'] = pd.Series(mapping_ab['A'].values)
    mapping_final['SL mã (Data)'] = pd.Series(mapping_ab['B'].values)
    mapping_final['Mã môn (Result)'] = pd.Series(df_cd['C'].values)
    mapping_final['SL dòng (Result)'] = pd.Series(df_cd['D'].values)
    mapping_final['Mã duy nhất (E)'] = pd.Series(mapping_ef['E'].values)
    mapping_final['Tổng thống kê (F)'] = pd.Series(mapping_ef['F'].values)
    
    def check_g(row):
        a, b, e, f = row['Mã môn (Data)'], row['SL mã (Data)'], row['Mã duy nhất (E)'], row['Tổng thống kê (F)']
        return "YES" if (pd.notna(a) and pd.notna(e) and str(a)==str(e) and b==f) else "NO"
    mapping_final['Kiểm tra (G)'] = mapping_final.apply(check_g, axis=1)
    mapping_final[' '] = pd.Series([None] * max_len)
    mapping_final['MÃ COURSE'] = pd.Series(khlm.iloc[:, 0].values)
    mapping_final['MÃ SV'] = pd.Series(khlm.iloc[:, 4].values)

    output_path = os.path.join(folder_path, 'Ket_Qua_KHLM_Mapping_Integrated.xlsx')
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        khlm.to_excel(writer, sheet_name='Result', index=False)
        data.to_excel(writer, sheet_name='data', index=False)
        mapping_final.to_excel(writer, sheet_name='Mapping', index=False)
        
        ws_res = writer.sheets['Result']
        for r in range(2, ws_res.max_row + 1):
            if ws_res.cell(row=r, column=6).value in [None, 0]: ws_res.row_dimensions[r].hidden = True

        ws_map = writer.sheets['Mapping']
        fill_ae = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        fill_bf = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
        fill_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        for r_idx, row in enumerate(ws_map.iter_rows(min_row=1, max_row=max_len+1, min_col=1, max_col=10)):
            for cell in row:
                cell.border = border; cell.alignment = Alignment(horizontal='center')
                if r_idx == 0: cell.font = Font(bold=True); continue
                col = cell.column_letter
                if col in ['A', 'E']: cell.fill = fill_ae
                elif col in ['B', 'F']: cell.fill = fill_bf
                if col == 'G' and cell.value == "YES": cell.fill = fill_green
                if col in ['I', 'J'] and ws_map.cell(row=cell.row, column=10).value: cell.fill = fill_green
    return output_path

def extract_courses_logic(folder_path):
    files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and not f.startswith('~$') and f != "Courses_List.xlsx"]
    if not files: raise ValueError("Không có file Excel nào để xử lý!")

    output_file = os.path.join(folder_path, "Courses_List.xlsx")
    wb_out = Workbook()
    default_sheet = wb_out.active
    wb_out.remove(default_sheet)

    for idx, file_name in enumerate(files):
        try:
            parts = file_name.split('- ')
            if len(parts) > 1: major_name = parts[1].split('.')[0].strip()
            else: major_name = file_name.split('.')[0].strip()
        except:
            major_name = file_name[:25] 

        ws_major = wb_out.create_sheet(title=major_name[:30])
        unique_courses = set()
        file_path = os.path.join(folder_path, file_name)
        wb_src = load_workbook(file_path, data_only=True)

        for sheet_name in wb_src.sheetnames:
            ws_src = wb_src[sheet_name]
            for r in range(16, ws_src.max_row + 1):
                val_a = str(ws_src.cell(row=r, column=1).value or "").strip().lower()
                if "lưu ý:" in val_a: break
                
                col_b = ws_src.cell(row=r, column=2).value 
                col_c = ws_src.cell(row=r, column=3).value 
                col_d = ws_src.cell(row=r, column=4).value 
                
                if col_b and col_c:
                    course_line = f"{str(col_b).strip()}_{str(col_c).strip()}_{str(col_d or 0).strip()} Tín chỉ"
                    unique_courses.add(course_line)

        sorted_list = sorted(list(unique_courses))
        for i, course in enumerate(sorted_list, 1): ws_major.cell(row=i, column=1, value=course)
        apply_full_border(ws_major)
        auto_fit_columns(ws_major)
        
    wb_out.save(output_file)
    return output_file

def fill_sllm_logic(folder_path):
    file_path = os.path.join(folder_path, "Data_SLLM.xlsx")
    if not os.path.exists(file_path): raise ValueError("Không tìm thấy tệp Data_SLLM.xlsx!")

    wb = load_workbook(file_path)
    if "Data" not in wb.sheetnames or "ThongKe" not in wb.sheetnames:
        raise ValueError("Tệp Data_SLLM.xlsx phải chứa sheet 'Data' và 'ThongKe'!")

    ws_data = wb["Data"]
    data_map = {}
    
    for r in range(2, ws_data.max_row + 1):
        cell_k = ws_data.cell(row=r, column=11).value
        val_k = str(cell_k or "").strip()
        region = ""
        if len(val_k) >= 4:
            sub_k = val_k[1:4].upper()
            if sub_k == "DNV": region = "HCM"
            elif sub_k == "DNP": region = "DNP"
        
        ws_data.cell(row=r, column=12, value=region)
        ma_lm = str(ws_data.cell(row=r, column=4).value or "").strip().lower()
        ma_hp = str(ws_data.cell(row=r, column=8).value or "").strip().lower()
        
        if ma_lm and ma_hp:
            key = (ma_hp, ma_lm)
            if key not in data_map: data_map[key] = {"DNP": 0, "HCM": 0}
            if region in ["DNP", "HCM"]: data_map[key][region] += 1

    ws_tk = wb["ThongKe"]
    for r in range(2, ws_tk.max_row + 1):
        ma_hp_tk = str(ws_tk.cell(row=r, column=2).value or "").strip().lower()
        ma_lm_str = str(ws_tk.cell(row=r, column=1).value or "").strip()
        dnp_sum = 0
        hcm_sum = 0
        
        if ma_hp_tk and ma_lm_str:
            codes = [c.strip().lower() for c in ma_lm_str.split(',')]
            for code in codes:
                key = (ma_hp_tk, code)
                if key in data_map:
                    dnp_sum += data_map[key]["DNP"]
                    hcm_sum += data_map[key]["HCM"]
        
        ws_tk.cell(row=r, column=3, value=dnp_sum)
        ws_tk.cell(row=r, column=4, value=hcm_sum)

    for sheet in [ws_data, ws_tk]:
        apply_full_border(sheet)
        auto_fit_columns(sheet)

    wb.save(file_path)
    return file_path

def compare_data_logic(folder_path):
    file_xlsb = os.path.join(folder_path, "Data.xlsb")
    file_src = os.path.join(folder_path, "Data_Source.xlsx")
    output_file = os.path.join(folder_path, "Compared_Result.xlsx")

    if not os.path.exists(file_xlsb) or not os.path.exists(file_src):
        raise ValueError("Không tìm thấy tệp Data.xlsb hoặc Data_Source.xlsx!")

    wb_src = load_workbook(file_src, data_only=True)
    if "DSSV" not in wb_src.sheetnames: raise ValueError("Thiếu sheet DSSV trong Data_Source.xlsx")
    
    ws_src = wb_src["DSSV"]
    source_codes = {str(ws_src.cell(row=r, column=10).value).strip().lower() 
                    for r in range(2, ws_src.max_row + 1) if ws_src.cell(row=r, column=10).value}

    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "Result"
    headers_added = False
    
    with open_xlsb(file_xlsb) as wb_bin:
        with wb_bin.get_sheet(1) as sheet:
            for row in sheet.rows():
                val_a = str(row[0].v).strip().lower() if row[0].v is not None else ""
                if not headers_added:
                    ws_out.append([c.v for c in row])
                    headers_added = True
                    continue
                if val_a in source_codes: ws_out.append([c.v for c in row])

    apply_full_border(ws_out)
    auto_fit_columns(ws_out)
    wb_out.save(output_file)
    return output_file

def filter_sv_logic(folder_path):
    f_dk, f_src = os.path.join(folder_path, "DanhSachDangKy.xlsx"), os.path.join(folder_path, "Data_Source.xlsx")
    if not os.path.exists(f_dk) or not os.path.exists(f_src): raise ValueError("Thiếu file DanhSachDangKy.xlsx hoặc Data_Source.xlsx")
    
    wb_src = load_workbook(f_src, data_only=True); ws_src = wb_src["DSSV"]
    sv_set = {str(ws_src.cell(row=r, column=4).value).strip().lower() for r in range(2, ws_src.max_row + 1) if ws_src.cell(row=r, column=4).value}
    wb_dk = load_workbook(f_dk, data_only=True)
    reg_data, learn_keys = [], set()
    h_reg, h_learn = [], []
    
    if "Dangky" in wb_dk.sheetnames:
        ws = wb_dk["Dangky"]; h_reg = [c.value for c in ws[1]]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[1] and str(r[1]).strip().lower() in sv_set: reg_data.append(r)
    if "Danghoc" in wb_dk.sheetnames:
        ws = wb_dk["Danghoc"]; h_learn = [c.value for c in ws[1]]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[1] and str(r[1]).strip().lower() in sv_set:
                learn_keys.add((str(r[1] or "").strip().lower(), str(r[6] or "").strip().lower(), str(r[14] or "").strip().lower()))
                
    m_rows, n_rows = [], []
    for r in reg_data:
        rk = (str(r[1] or "").strip().lower(), str(r[2] or "").strip().lower(), str(r[19] or "").strip().lower())
        if rk in learn_keys: m_rows.append(r)
        else: n_rows.append(r)
        
    output_file = os.path.join(folder_path, "Filter_Result.xlsx")
    wb_res = Workbook()
    configs = [("Register_List", reg_data, h_reg), ("matching", m_rows, h_reg), ("not_matching", n_rows, h_reg)]
    for i, (t, d, h) in enumerate(configs):
        ws = wb_res.active if i == 0 else wb_res.create_sheet(t)
        ws.title = t; ws.append(h)
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
# CẤU TRÚC MENU (SIDEBAR)
# ==========================================
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/732/732220.png", width=60) # Icon Excel trang trí nhẹ nhàng
    st.markdown("<h3 style='color: #1E3A8A; font-weight: 700; margin-top: -10px;'>CÔNG CỤ XỬ LÝ DỮ LIỆU</h3>", unsafe_allow_html=True)
    st.markdown("---")
    
    menu_options = {
        "📂 Gộp File Nguồn": "Gộp File Nguồn",
        "✏️ Điền kế hoạch lớp môn (KHLM)(*)": "Điền kế hoạch lớp môn (KHLM)(*)",
        "🔍 Lọc kết quả học tập của SV": "Lọc kết quả học tập của sinh viên",
        "🔄 Lọc SV học lại & cải thiện": "Lọc sinh viên học lại & học cải thiện",
        "🏷️ Xuất Mã lớp theo GK300": "Xuất Mã lớp theo GK300",
        "📄 Xuất KHHT theo GK300 (1*)": "Xuất KHHT theo GK300 (1*)",
        "👥 Xuất DSSV theo GK300 (2*)": "Xuất DSSV theo GK300 (2*)",
        "📝 Xuất KHHTCT theo GK300 (3*)": "Xuất KHHTCT theo GK300 (3*)",
        "📊 Thống kê SL theo lớp & môn (*)": "Thống kê số lượng theo lớp/nhóm lớp & môn (*)",
        "🎓 Xuất môn theo ngành học (*)": "Xuất môn theo ngành học (*)"
    }
    
    selected_label = st.radio("CHỌN CHỨC NĂNG BÊN DƯỚI", list(menu_options.keys()))
    choice = menu_options[selected_label]
    
    st.markdown("---")
    st.caption("✨ Web Version 3.1 | Phát triển bởi Đặng Văn Nghĩa")

# ==========================================
# GIAO DIỆN CHÍNH (MAIN CONTENT)
# ==========================================

# 1. Header Web App (Thay thế st.title để trông gọn và xịn hơn)
st.markdown(f"""
    <div class="app-header">
        <h1>{selected_label}</h1>
        <p>Hệ thống tự động hóa xử lý và phân tích số liệu Excel Professional</p>
    </div>
""", unsafe_allow_html=True)

# 2. Khối Hướng dẫn Dữ liệu (Dùng thẻ div Custom HTML để hiển thị dạng Card)
instructions = {
    "Gộp File Nguồn": "Đầu vào là file <b>GK300</b> của 1 hoặc nhiều khóa (Mỗi sheet chứa bảng đăng ký kế hoạch học tập).",
    "Điền kế hoạch lớp môn (KHLM)(*)": "<b>Yêu cầu:</b> File <code>Data_fill.xlsx</code><br><br><b>Sheet KHLM:</b> Cột E(Mã SV), F(SL), C(Mã môn), H(Tên lớp), I(Địa phương/Mã trạm).<br><b>Sheet data:</b> Cột A(Mã lớp), C(Mã SV), K(Mã môn), L(Mã trạm).",
    "Lọc kết quả học tập của sinh viên": "<b>Yêu cầu:</b> <code>Data_Source.xlsx</code> và <code>Data.xlsb</code><br><br><b>Data_Source.xlsx:</b> Sheet 'DSSV', Cột Q (Tên lớp), Cột J (Mã SV), Cột D (Tài khoản SV).<br><b>Data.xlsb:</b> File nhị phân, Cột A là 'Mã SV'.",
    "Lọc sinh viên học lại & học cải thiện": "<b>Yêu cầu:</b> <code>Data_Source.xlsx</code> và <code>DanhSachDangKy.xlsx</code><br><br><b>DanhSachDangKy.xlsx:</b> Sheet 'Dangky: B,C,T' và 'Danghoc: B,G,O', Cột B (Tài khoản SV), Cột C&G (Mã môn), Cột T&O (Số TC).",
    "Xuất Mã lớp theo GK300": "Đầu vào là file <b>GK300</b> của 1 hoặc nhiều khóa (Mỗi sheet chứa bảng đăng ký kế hoạch học tập).",
    "Xuất KHHT theo GK300 (1*)": "Đầu vào là file <b>GK300</b> của 1 hoặc nhiều khóa (Mỗi sheet chứa bảng đăng ký kế hoạch học tập).",
    "Xuất DSSV theo GK300 (2*)": "<b>Yêu cầu:</b> <code>Merged_GK300.xlsx</code> tạo từ (1*) và <code>Data_Source.xlsx</code><br><br><b>Merged_GK300.xlsx:</b> Sheet 'KHHT_GK300', Cột L (12) là 'Mã LT'.",
    "Xuất KHHTCT theo GK300 (3*)": "<b>Yêu cầu:</b> Cần cung cấp file <code>Merged_GK300.xlsx</code>.",
    "Thống kê số lượng theo lớp/nhóm lớp & môn (*)": "<b>Yêu cầu:</b> File <code>Data_SLLM.xlsx</code><br><br><b>Sheet Data:</b> Cột L (Mã trạm).<br><b>Sheet ThongKe:</b> Cột A (Tên lớp), Cột B (Mã môn), Tiêu đề Cột C (Mã Trạm).",
    "Xuất môn theo ngành học (*)": "Đầu vào là file dữ liệu môn học phân bổ theo ngành và theo khóa (Mỗi sheet chứa bảng đăng ký kế hoạch học tập)."
}

st.markdown(f"""
    <div class="instruction-card">
        <strong>📌 Hướng dẫn chuẩn bị dữ liệu:</strong><br><br>
        {instructions.get(choice, "Vui lòng upload các file Excel theo đúng định dạng.")}
    </div>
""", unsafe_allow_html=True)

# 3. Input Cấu hình Phụ (Chỉ hiện khi chọn chức năng Điền KHLM)
keywords_str = ""
if choice == "Điền kế hoạch lớp môn (KHLM)(*)":
    st.markdown("##### ⚙️ Thiết lập điều kiện lọc nâng cao")
    keywords_str = st.text_input("Nhập các từ khóa địa phương (cách nhau bằng dấu phẩy):", value="DNP, HCM(Oanh)")
    st.caption("Gợi ý: DNP, HCM(Oanh), BP, SG...")
    st.write("") # Dòng trống tạo khoảng cách

# 4. Khu vực Upload và Thực thi (Sắp xếp cột gọn gàng)
col_space_left, col_main, col_space_right = st.columns([1, 8, 1])

with col_main:
    uploaded_files = st.file_uploader("Kéo thả các file Excel vào khu vực này", accept_multiple_files=True, type=['xlsx', 'xlsb'])
    
    st.write("") # Khoảng trống trước nút bấm
    
    if st.button("🚀 BẮT ĐẦU XỬ LÝ DỮ LIỆU"):
        if not uploaded_files:
            st.error("⚠️ Hệ thống chưa nhận được file. Vui lòng tải dữ liệu lên trước khi chạy!")
        else:
            temp_dir = tempfile.mkdtemp()
            try:
                for uploaded_file in uploaded_files:
                    file_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                
                with st.spinner('⚙️ Đang phân tích và xử lý dữ liệu...'):
                    result_file = None
                    
                    if choice == "Gộp File Nguồn": result_file = process_files_logic(temp_dir)
                    elif choice == "Điền kế hoạch lớp môn (KHLM)(*)": result_file = fill_khlm_logic(temp_dir, keywords_str)
                    elif choice == "Lọc kết quả học tập của sinh viên": result_file = compare_data_logic(temp_dir)
                    elif choice == "Lọc sinh viên học lại & học cải thiện": result_file = filter_sv_logic(temp_dir)
                    elif choice == "Xuất Mã lớp theo GK300": result_file = extract_class_names_logic(temp_dir)
                    elif choice == "Xuất KHHT theo GK300 (1*)": result_file = export_khht_logic(temp_dir)
                    elif choice == "Xuất DSSV theo GK300 (2*)": result_file = export_dssv_logic(temp_dir)
                    elif choice == "Xuất KHHTCT theo GK300 (3*)": result_file = export_khhtct_logic(temp_dir)
                    elif choice == "Thống kê số lượng theo lớp/nhóm lớp & môn (*)": result_file = fill_sllm_logic(temp_dir)
                    elif choice == "Xuất môn theo ngành học (*)": result_file = extract_courses_logic(temp_dir)

                    if result_file and os.path.exists(result_file):
                        st.balloons()
                        st.success("✅ Quá trình xử lý đã hoàn tất. File kết quả của bạn đã sẵn sàng!")
                        
                        with open(result_file, "rb") as f:
                            file_data = f.read()
                        
                        st.download_button(
                            label="⬇️ TẢI XUỐNG FILE KẾT QUẢ",
                            data=file_data,
                            file_name=os.path.basename(result_file),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("❌ Xử lý thất bại. Vui lòng kiểm tra lại tính hợp lệ của dữ liệu đầu vào.")
                        
            except Exception as e:
                st.error(f"❌ Phát hiện lỗi hệ thống: `{str(e)}`")
                st.info("💡 Lời khuyên: Hãy kiểm tra lại tên file và đối chiếu với Hướng dẫn chuẩn bị dữ liệu bên trên.")
            
            finally:
                shutil.rmtree(temp_dir)
