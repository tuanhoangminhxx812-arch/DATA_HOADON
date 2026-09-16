import os
import sys
import io
import re
import datetime
import xml.etree.ElementTree as ET
from copy import copy

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import pypdf
except ImportError:
    pypdf = None

# Set wide page layout & modern title
st.set_page_config(
    page_title="Hóa Đơn MTMN - Tách & Xuất DataLoad",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for polished interface
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #4B5563;
        margin-bottom: 1.2rem;
    }
    .metric-box {
        background-color: #F8FAFC;
        padding: 0.9rem;
        border-radius: 8px;
        border: 1px solid #E2E8F0;
        text-align: center;
    }
    .error-card {
        background-color: #FEF2F2;
        padding: 1rem;
        border-radius: 8px;
        border-left: 5px solid #DC2626;
        margin-bottom: 1rem;
    }
    .info-banner {
        background-color: #EFF6FF;
        border-left: 5px solid #3B82F6;
        padding: 0.8rem 1rem;
        border-radius: 6px;
        margin-bottom: 1rem;
        font-size: 0.95rem;
        color: #1E40AF;
    }
</style>
""", unsafe_allow_html=True)

# ----------------- PATH & TEMPLATE SETUP ----------------- #

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE_NAME = "DataLoad_MTMN_T09 (MAU).xlsx"
DEFAULT_TEMPLATE_PATH = os.path.join(CURRENT_DIR, DEFAULT_TEMPLATE_NAME)
if not os.path.exists(DEFAULT_TEMPLATE_PATH):
    DEFAULT_TEMPLATE_PATH = r'd:\DATA\DATA_HOADON\DataLoad_MTMN_T09 (MAU).xlsx'

# ----------------- PARSER FUNCTIONS ----------------- #

def extract_period(text, nlap):
    """
    Trích xuất kỳ tiền điện (vd: T08-2026, T07-2026...) từ nội dung hàng hóa dịch vụ THHDVu.
    Nếu không có, dùng ngày lập hóa đơn NLap.
    """
    cleaned = re.sub(r'\s+', ' ', text or '')
    dates = re.findall(r'(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})', cleaned)
    if dates:
        for d, m, y in dates:
            if int(y) >= 2020:
                return f'T{int(m):02d}-{y}'
    m = re.search(r'(?:tháng|T)\s*(\d{1,2})\s*[-/]\s*(\d{4})', cleaned, re.IGNORECASE)
    if m:
        return f'T{int(m.group(1)):02d}-{m.group(2)}'
    if nlap:
        parts = str(nlap).split('-')
        if len(parts) >= 2:
            return f'T{int(parts[1]):02d}-{parts[0]}'
    return 'T08-2026'

def clean_tax_rate(val):
    """
    Xử lý thuế suất.
    Trả về số nguyên (8, 10, 5, 0) hoặc chuỗi nếu là KCT (Không chịu thuế), KKKNT.
    """
    if val is None or str(val).strip() == '':
        return 0
    s = str(val).strip().upper()
    if 'KCT' in s or 'KHÔNG CHỊU THUẾ' in s:
        return 'KCT'
    if 'KKKNT' in s:
        return 'KKKNT'
    m = re.search(r'(\d+)', s)
    if m:
        return int(m.group(1))
    return 0

def parse_xml_invoice(content_bytes, filename=""):
    """
    Phân tích cú pháp file hóa đơn điện tử XML.
    """
    try:
        root = ET.fromstring(content_bytes)
    except Exception as e:
        raise ValueError(f"Lỗi cú pháp XML ({e})")
        
    khms = (root.find('.//KHMSHDon').text or '').strip() if root.find('.//KHMSHDon') is not None else ''
    kh = (root.find('.//KHHDon').text or '').strip() if root.find('.//KHHDon') is not None else ''
    sh = (root.find('.//SHDon').text or '').strip() if root.find('.//SHDon') is not None else ''
    full_kh = khms + kh
    sh_num = int(sh) if sh.isdigit() else sh
    
    nlap_str = (root.find('.//NLap').text or '').strip() if root.find('.//NLap') is not None else ''
    try:
        nlap_date = datetime.datetime.strptime(nlap_str, '%Y-%m-%d')
    except:
        nlap_date = nlap_str
        
    thhdvu = (root.find('.//THHDVu').text or '').strip() if root.find('.//THHDVu') is not None else ''
    period = extract_period(thhdvu, nlap_str)
    
    ten_nban = (root.find('.//NBan/Ten').text or '').strip() if root.find('.//NBan/Ten') is not None else ''
    mst_nban = (root.find('.//NBan/MST').text or '').strip() if root.find('.//NBan/MST') is not None else ''
    if mst_nban.isdigit() and not mst_nban.startswith('0'):
        mst_val = int(mst_nban)
    else:
        mst_val = mst_nban
        
    tgtcthue = root.find('.//TgTCThue')
    if tgtcthue is not None and tgtcthue.text:
        tien_truoc_thue = round(float(tgtcthue.text))
    else:
        thtien = root.find('.//ThTien')
        tien_truoc_thue = round(float(thtien.text)) if thtien is not None and thtien.text else 0
        
    tgtthue = root.find('.//TgTThue')
    if tgtthue is not None and tgtthue.text:
        tien_thue = round(float(tgtthue.text))
    else:
        tthue = root.find('.//TThue')
        tien_thue = round(float(tthue.text)) if tthue is not None and tthue.text else 0
        
    tsuat = root.find('.//TSuat')
    tsuat_val = clean_tax_rate(tsuat.text if tsuat is not None else '8')
    
    # Có thuế khi tiền thuế > 0 và mức thuế không phải 0 / KCT
    is_taxable = (tien_thue > 0) and (tsuat_val not in [0, '0', '0%', 'KCT', 'KKKNT'])
    
    noi_dung = f'Thuế GTGT điện MTMN {period}'
    nd = f'ĐIỆN MTMN {period}'
    
    return {
        'filename': filename,
        'format': 'XML',
        'period': period,
        'tien_thue': tien_thue,
        'noi_dung': noi_dung,
        'ky_hieu_hd': full_kh,
        'so_hd': sh_num,
        'ngay_hd': nlap_date,
        'ten_kh': ten_nban,
        'mst': mst_val,
        'nd': nd,
        'tien_truoc_thue': tien_truoc_thue,
        'thue_suat': tsuat_val,
        'is_taxable': is_taxable
    }

def parse_pdf_invoice(content_bytes, filename=""):
    """
    Phân tích cú pháp file hóa đơn điện tử PDF (TT78 / NĐ123).
    """
    if pypdf is None:
        raise ValueError("Thư viện pypdf chưa được cài đặt!")
        
    try:
        reader = pypdf.PdfReader(io.BytesIO(content_bytes))
    except Exception as e:
        raise ValueError(f"Không thể đọc file PDF ({e})")
        
    full_text = ""
    for page in reader.pages:
        txt = page.extract_text() or ""
        full_text += txt + "\n"
        
    if not full_text.strip():
        raise ValueError("File PDF không chứa văn bản (có thể là file scan/ảnh dạng PDF).")
        
    lines = [line.strip() for line in full_text.split('\n') if line.strip()]
    
    # 1. Ký hiệu
    full_kh = ""
    m_full_kh = re.search(r'\b([12][A-Z]\d{2}[A-Z]{3})\b', full_text)
    if m_full_kh:
        full_kh = m_full_kh.group(1)
    else:
        m_ms = re.search(r'(?:Mẫu\s*số|Ký\s*hiệu\s*mẫu\s*số)[:\s]*([12][A-Z]?)', full_text, re.IGNORECASE)
        khms = m_ms.group(1).strip() if m_ms else "1"
        m_kh = re.search(r'(?:Ký\s*hiệu|Ký\s*hiệu\s*HĐ|Serial)[:\s]*([A-Z0-9]{5,8})', full_text, re.IGNORECASE)
        kh = m_kh.group(1).strip() if m_kh else "C26MHC"
        full_kh = khms + kh if not kh.startswith(khms) else kh

    # 2. Số hóa đơn
    sh_num = 1
    m_sh = re.search(r'(?:Số\s*hóa\s*đơn|Số\s*HĐ|Số\s*\(No\.\)|Số/No\.|Số|No\.)[:\s]*0*(\d+)', full_text, re.IGNORECASE)
    if m_sh:
        sh_num = int(m_sh.group(1))
    else:
        m_file_sh = re.search(r'-\s*0*(\d+)(?:\.pdf)?$', filename, re.IGNORECASE)
        if m_file_sh:
            sh_num = int(m_file_sh.group(1))

    # 3. Ngày hóa đơn
    nlap_date = datetime.datetime.now()
    m_date1 = re.search(r'Ngày\s*0?(\d{1,2})\s*tháng\s*0?(\d{1,2})\s*năm\s*(\d{4})', full_text, re.IGNORECASE)
    if m_date1:
        d, m, y = int(m_date1.group(1)), int(m_date1.group(2)), int(m_date1.group(3))
        nlap_date = datetime.datetime(y, m, d)
    else:
        m_date2 = re.search(r'(?:Ngày\s*lập|Ngày\s*HĐ|Ngày)[:\s]*0?(\d{1,2})[/.-]0?(\d{1,2})[/.-](\d{4})', full_text, re.IGNORECASE)
        if m_date2:
            d, m, y = int(m_date2.group(1)), int(m_date2.group(2)), int(m_date2.group(3))
            nlap_date = datetime.datetime(y, m, d)

    # 4. Tên người bán
    ten_nban = ""
    m_ten = re.search(r'(?:Đơn\s*vị\s*bán(?:\s*hàng)?|Tên\s*người\s*bán|Người\s*bán)[:\s]*([^\n\r]+)', full_text, re.IGNORECASE)
    if m_ten:
        ten_nban = m_ten.group(1).strip()
    else:
        for line in lines[:15]:
            if line.upper().startswith(('CÔNG TY', 'CTY', 'DOANH NGHIỆP', 'DNTN', 'CHI NHÁNH')):
                ten_nban = line
                break

    # 5. MST người bán
    mst_val = ""
    m_mst_list = re.findall(r'(?:Mã\s*số\s*thuế|MST)(?:\s*\(Tax\s*code\))?[:\s]*([0-9]{10}(?:-[0-9]{3})?)', full_text, re.IGNORECASE)
    if m_mst_list:
        for mst_cand in m_mst_list:
            if not mst_cand.startswith('0300951119'):
                mst_val = mst_cand
                break
        if not mst_val:
            mst_val = m_mst_list[0]
    if mst_val.isdigit() and not mst_val.startswith('0'):
        mst_val = int(mst_val)

    # 6. Kỳ phát điện & Nội dung
    period = extract_period(full_text, nlap_date.strftime('%Y-%m-%d'))
    noi_dung = f'Thuế GTGT điện MTMN {period}'
    nd = f'ĐIỆN MTMN {period}'

    # 7. Tiền trước thuế, Thuế, Thuế suất
    def parse_money(s):
        if not s: return 0
        cleaned = re.sub(r'[^\d]', '', str(s))
        return int(cleaned) if cleaned else 0

    tien_truoc_thue = 0
    tien_thue = 0
    thue_suat = 8
    
    m_ttt = re.search(r'(?:Cộng\s*tiền\s*hàng|Tổng\s*tiền\s*chưa\s*thuế|Tiền\s*trước\s*thuế|Tổng\s*cộng\s*tiền\s*hàng)[:\s]*([\d.,]+)', full_text, re.IGNORECASE)
    if m_ttt:
        tien_truoc_thue = parse_money(m_ttt.group(1))

    m_ts = re.search(r'(?:Thuế\s*suất\s*GTGT|Thuế\s*suất|VAT\s*rate)[:\s]*([0-9%]+|KCT|KKKNT)', full_text, re.IGNORECASE)
    if m_ts:
        ts_raw = m_ts.group(1).upper()
        if 'KCT' in ts_raw or 'KKKNT' in ts_raw:
            thue_suat = 'KCT'
        else:
            m_num = re.search(r'(\d+)', ts_raw)
            thue_suat = int(m_num.group(1)) if m_num else 8

    m_tt = re.search(r'(?:Tiền\s*thuế\s*GTGT|Tiền\s*thuế|VAT\s*amount)[:\s]*([\d.,]+)', full_text, re.IGNORECASE)
    if m_tt:
        tien_thue = parse_money(m_tt.group(1))
    elif isinstance(thue_suat, (int, float)) and thue_suat > 0 and tien_truoc_thue > 0:
        tien_thue = round(tien_truoc_thue * thue_suat / 100)

    is_taxable = (tien_thue > 0) and (thue_suat not in [0, '0', '0%', 'KCT', 'KKKNT'])

    return {
        'filename': filename,
        'format': 'PDF',
        'period': period,
        'tien_thue': tien_thue,
        'noi_dung': noi_dung,
        'ky_hieu_hd': full_kh,
        'so_hd': sh_num,
        'ngay_hd': nlap_date,
        'ten_kh': ten_nban,
        'mst': mst_val,
        'nd': nd,
        'tien_truoc_thue': tien_truoc_thue,
        'thue_suat': thue_suat,
        'is_taxable': is_taxable
    }

def parse_invoice_file(filename, file_bytes):
    """
    Tự động xác định định dạng XML hoặc PDF và trích xuất dữ liệu.
    """
    fname_lower = filename.lower()
    if fname_lower.endswith('.xml'):
        return parse_xml_invoice(file_bytes, filename)
    elif fname_lower.endswith('.pdf'):
        return parse_pdf_invoice(file_bytes, filename)
    else:
        # Thử XML trước, nếu không được thử PDF
        try:
            return parse_xml_invoice(file_bytes, filename)
        except:
            return parse_pdf_invoice(file_bytes, filename)

# ----------------- EXCEL GENERATOR ----------------- #

def generate_excel_bytes(valid_items, template_bytes_or_path):
    """
    Sinh file Excel chuẩn theo mẫu:
    - Sheet 'DATALOAD TỔNG' và từng sheet theo từng tháng.
    - Sắp xếp: Hóa đơn có thuế lên trên, hóa đơn không thuế xuống dưới.
    - Hóa đơn không thuế được TÔ VÀNG TOÀN BỘ DÒNG để nhận biết rõ ràng.
    - Cột ngày tháng định dạng: dd/mm/yyyy
    - Cột số tiền định dạng có dấu phân cách: #,##0
    """
    if isinstance(template_bytes_or_path, str):
        wb_template = openpyxl.load_workbook(template_bytes_or_path)
    else:
        wb_template = openpyxl.load_workbook(io.BytesIO(template_bytes_or_path))
        
    ws_tmpl = wb_template.active
    
    wb_new = openpyxl.Workbook()
    wb_new.remove(wb_new.active)
    
    periods = sorted(list(set(item['period'] for item in valid_items)), reverse=True)
    
    sheets_to_create = [('DATALOAD TỔNG', valid_items)]
    for p in periods:
        items_p = [item for item in valid_items if item['period'] == p]
        sheets_to_create.append((p, items_p))
        
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    font_tnr = Font(name='Times New Roman', size=10)
    
    for sheet_name, raw_items in sheets_to_create:
        ws = wb_new.create_sheet(title=sheet_name)
        
        # Sao chép độ rộng cột từ template
        for col_letter, cd in ws_tmpl.column_dimensions.items():
            ws.column_dimensions[col_letter].width = cd.width
            
        # Sao chép dòng 1 (Header)
        for col_idx in range(1, ws_tmpl.max_column + 1):
            src_cell = ws_tmpl.cell(1, col_idx)
            dest_cell = ws.cell(1, col_idx, value=src_cell.value)
            if src_cell.fill and src_cell.fill.start_color and 'FFFF00' in str(src_cell.fill.start_color.rgb):
                dest_cell.fill = yellow_fill
            if src_cell.font:
                dest_cell.font = copy(src_cell.font)
            if src_cell.alignment:
                dest_cell.alignment = copy(src_cell.alignment)
            if src_cell.border:
                dest_cell.border = copy(src_cell.border)

        # Sắp xếp: Có thuế lên trên, Không thuế xuống dưới
        taxable_items = [it for it in raw_items if it['is_taxable']]
        nontaxable_items = [it for it in raw_items if not it['is_taxable']]
        sorted_items = taxable_items + nontaxable_items

        for idx, item in enumerate(sorted_items, start=1):
            r = idx + 1
            is_taxable = item['is_taxable']
            
            # Col A: STT
            if idx == 1:
                ws.cell(r, 1, value=1)
            else:
                ws.cell(r, 1, value=f'=A{r-1}+1')
            ws.cell(r, 1).font = font_tnr
            
            # Col B: Mã hạch toán
            ws.cell(r, 2, value='082900.000000.000.13311000000.1182.0000.000.000000.0000.0000').font = font_tnr
            
            # Col C: TAB
            ws.cell(r, 3, value='TAB').font = font_tnr
            
            # Col D: Thuế (định dạng số có dấu phân cách hàng nghìn)
            cD = ws.cell(r, 4, value=item['tien_thue'])
            cD.font = font_tnr
            cD.number_format = '#,##0'
            
            # Col E: TAB
            ws.cell(r, 5, value='TAB').font = font_tnr
            
            # Col F: Nội dung
            cF = ws.cell(r, 6, value=item['noi_dung'])
            cF.font = font_tnr
            
            # Col G, H: TAB
            ws.cell(r, 7, value='TAB').font = font_tnr
            ws.cell(r, 8, value='TAB').font = font_tnr
            
            # Col I: Kê khai thuế GTGT - Đầu vào
            ws.cell(r, 9, value='Kê khai thuế GTGT - Đầu vào').font = font_tnr
            
            # Col J: TAB
            ws.cell(r, 10, value='TAB').font = font_tnr
            
            # Col K: Ký hiệu HĐ
            cK = ws.cell(r, 11, value=item['ky_hieu_hd'])
            cK.font = font_tnr
            
            # Col L: TAB
            ws.cell(r, 12, value='TAB').font = font_tnr
            
            # Col M: Số HĐ
            cM = ws.cell(r, 13, value=item['so_hd'])
            cM.font = font_tnr
            
            # Col N: TAB
            ws.cell(r, 14, value='TAB').font = font_tnr
            
            # Col O: Ngày HĐ (Định dạng chuẩn dd/mm/yyyy theo yêu cầu)
            cO = ws.cell(r, 15, value=item['ngay_hd'])
            cO.font = font_tnr
            cO.number_format = 'dd/mm/yyyy'
            
            # Col P: TAB
            ws.cell(r, 16, value='TAB').font = font_tnr
            
            # Col Q: 1
            ws.cell(r, 17, value=1).font = font_tnr
            
            # Col R: TAB
            ws.cell(r, 18, value='TAB').font = font_tnr
            
            # Col S: Tên KH
            cS = ws.cell(r, 19, value=item['ten_kh'])
            cS.font = font_tnr
            
            # Col T: TAB
            ws.cell(r, 20, value='TAB').font = font_tnr
            
            # Col U: MST
            cU = ws.cell(r, 21, value=item['mst'])
            cU.font = font_tnr
            
            # Col V: TAB
            ws.cell(r, 22, value='TAB').font = font_tnr
            
            # Col W: ND
            ws.cell(r, 23, value=item['nd']).font = font_tnr
            
            # Col X: TAB
            ws.cell(r, 24, value='TAB').font = font_tnr
            
            # Col Y: Số tiền trước thuế (định dạng số có dấu phân cách hàng nghìn)
            cY = ws.cell(r, 25, value=item['tien_truoc_thue'])
            cY.font = font_tnr
            cY.number_format = '#,##0'
            
            # Col Z: TAB
            ws.cell(r, 26, value='TAB').font = font_tnr
            
            # Col AA: Thuế suất
            cAA = ws.cell(r, 27, value=item['thue_suat'])
            cAA.font = font_tnr
            
            # Col AB, AC, AD, AE: TAB, CK, tab, ENT
            ws.cell(r, 28, value='TAB').font = font_tnr
            ws.cell(r, 29, value='CK').font = font_tnr
            ws.cell(r, 30, value='tab').font = font_tnr
            ws.cell(r, 31, value='ENT').font = font_tnr
            ws.cell(r, 32, value=None)
            
            # Col AG: test MST
            ws.cell(r, 33, value=f'=LEN(U{r})').font = font_tnr
            
            # Col AH: test VAT
            ws.cell(r, 34, value=f'=ROUND(D{r}-(Y{r}*(AA{r}/100)),0)').font = font_tnr

            # XỬ LÝ TÔ VÀNG:
            # 1. Nếu có thuế: tô vàng các cột chỉ định D, F, K, M, O, S, U, Y, AA
            # 2. Nếu KHÔNG CÓ THUẾ: tô vàng TOÀN BỘ CÁC CỘT (A -> AH) để anh nhận biết ngay
            if is_taxable:
                for c_target in [cD, cF, cK, cM, cO, cS, cU, cY, cAA]:
                    c_target.fill = yellow_fill
            else:
                for col_i in range(1, 35):
                    ws.cell(r, col_i).fill = yellow_fill
                    
    output_buffer = io.BytesIO()
    wb_new.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer

# ----------------- STREAMLIT INTERFACE ----------------- #

st.markdown('<div class="main-header">⚡ Xử Lý Hóa Đơn Điện Tử MTMN (.XML & .PDF)</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Tự động nạp vào mẫu DataLoad chuẩn, phân tách sheet theo tháng, định dạng ngày <b>dd/mm/yyyy</b> và phân cách số tiền hàng nghìn.</div>', unsafe_allow_html=True)

# Tự động nạp mẫu chuẩn
template_file_bytes = None
if os.path.exists(DEFAULT_TEMPLATE_PATH):
    with open(DEFAULT_TEMPLATE_PATH, 'rb') as f:
        template_file_bytes = f.read()
    st.markdown(f"""
    <div class="info-banner">
        ✓ <b>Mẫu Excel chuẩn đang áp dụng:</b> <code>{os.path.basename(DEFAULT_TEMPLATE_PATH)}</code> (Đã cấu hình sẵn, bạn chỉ cần tải hóa đơn lên).
    </div>
    """, unsafe_allow_html=True)
else:
    st.warning("⚠️ Không tìm thấy file mẫu mặc định, vui lòng tải file mẫu lên ở phần tùy chọn bên dưới.")
    uploaded_tmpl = st.file_uploader("Tải file mẫu Excel (.xlsx)", type=["xlsx"])
    if uploaded_tmpl:
        template_file_bytes = uploaded_tmpl.read()

# Input files
tab_upload, tab_folder = st.tabs(["📤 Tải Lên Hóa Đơn (.XML hoặc .PDF)", "📁 Quét Thư Mục Hóa Đơn"])

raw_files = [] # list of (filename, bytes)

with tab_upload:
    uploaded_files = st.file_uploader(
        "Kéo thả hoặc chọn các file hóa đơn (.xml hoặc .pdf):",
        type=["xml", "pdf"],
        accept_multiple_files=True,
        help="Hỗ trợ cả file .XML và .PDF hóa đơn điện tử. Nhấn Ctrl+A trong cửa sổ chọn file để chọn toàn bộ."
    )
    if uploaded_files:
        for uf in uploaded_files:
            raw_files.append((uf.name, uf.read()))

with tab_folder:
    default_folder = os.path.join(CURRENT_DIR, "XML Tháng 8 - Đợt 1")
    if not os.path.exists(default_folder):
        default_folder = r"d:\DATA\DATA_HOADON\XML Tháng 8 - Đợt 1"
        
    folder_path = st.text_input("Đường dẫn thư mục chứa hóa đơn trên máy:", value=default_folder)
    if st.button("🔍 Quét Thư Mục", use_container_width=False):
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.xml', '.pdf'))]
            if files:
                raw_files = []
                for fname in files:
                    try:
                        with open(os.path.join(folder_path, fname), 'rb') as f:
                            raw_files.append((fname, f.read()))
                    except Exception as e:
                        st.error(f"Lỗi đọc file {fname}: {e}")
                st.session_state['folder_raw_files'] = raw_files
                st.success(f"Đã tìm thấy **{len(raw_files)}** file hóa đơn trong thư mục!")
            else:
                st.warning("Không tìm thấy file .xml hoặc .pdf nào trong thư mục!")
        else:
            st.error("Thư mục không tồn tại. Vui lòng kiểm tra lại đường dẫn.")
            
    if 'folder_raw_files' in st.session_state and not raw_files:
        raw_files = st.session_state['folder_raw_files']

# Process Files
if raw_files:
    st.markdown("---")
    valid_items = []
    invalid_files = []
    
    for filename, content_bytes in raw_files:
        try:
            item = parse_invoice_file(filename, content_bytes)
            valid_items.append(item)
        except Exception as e:
            invalid_files.append({
                'Tên file': filename,
                'Kích thước': f"{len(content_bytes):,} bytes",
                'Chi tiết lỗi': str(e)
            })

    # Metrics
    c1, c2, c3, c4, c5 = st.columns(5)
    taxable_cnt = sum(1 for it in valid_items if it['is_taxable'])
    nontax_cnt = len(valid_items) - taxable_cnt
    
    with c1:
        st.metric("Tổng số file", len(raw_files))
    with c2:
        st.metric("Hóa đơn hợp lệ", len(valid_items))
    with c3:
        st.metric("File bị lỗi / Cần chép lại", len(invalid_files), delta=-len(invalid_files) if invalid_files else 0, delta_color="inverse")
    with c4:
        st.metric("HĐ Có thuế (Xếp trên)", taxable_cnt)
    with c5:
        st.metric("HĐ Không thuế (Tô vàng)", nontax_cnt)

    # Invalid files alert
    if invalid_files:
        st.markdown('<div class="error-card">', unsafe_allow_html=True)
        st.error(f"⚠️ **Có {len(invalid_files)} file bị lỗi hoặc hỏng nội dung!** Vui lòng kiểm tra và bổ sung lại:")
        df_invalid = pd.DataFrame(invalid_files)
        st.dataframe(df_invalid, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # Data preview & Export
    tab_view, tab_by_month, tab_dl = st.tabs(["👁️ Bảng Dữ Liệu Chi Tiết", "📅 Thống Kê Theo Tháng", "🚀 Xuất & Tải File Excel"])
    
    with tab_view:
        if valid_items:
            df_table = pd.DataFrame([
                {
                    'File': it['filename'],
                    'Loại': it['format'],
                    'Kỳ': it['period'],
                    'Phân loại thuế': '🟢 Có thuế (Xếp trên)' if it['is_taxable'] else '🟡 Không thuế (Tô vàng)',
                    'Ký hiệu': it['ky_hieu_hd'],
                    'Số HĐ': it['so_hd'],
                    'Ngày HĐ': it['ngay_hd'].strftime('%d/%m/%Y') if isinstance(it['ngay_hd'], (datetime.datetime, datetime.date)) else str(it['ngay_hd']),
                    'Tên người bán': it['ten_kh'],
                    'MST': it['mst'],
                    'Tiền trước thuế': f"{it['tien_truoc_thue']:,}".replace(',', ' '),
                    'Thuế suất': f"{it['thue_suat']}%" if isinstance(it['thue_suat'], int) else it['thue_suat'],
                    'Tiền thuế': f"{it['tien_thue']:,}".replace(',', ' '),
                }
                for it in valid_items
            ])
            st.dataframe(df_table, use_container_width=True, height=450)
            
    with tab_by_month:
        if valid_items:
            p_stats = {}
            for it in valid_items:
                p = it['period']
                if p not in p_stats:
                    p_stats[p] = {'Tổng HĐ': 0, 'Có thuế': 0, 'Không thuế': 0, 'Tiền trước thuế': 0, 'Tiền thuế': 0}
                p_stats[p]['Tổng HĐ'] += 1
                if it['is_taxable']:
                    p_stats[p]['Có thuế'] += 1
                else:
                    p_stats[p]['Không thuế'] += 1
                p_stats[p]['Tiền trước thuế'] += it['tien_truoc_thue']
                p_stats[p]['Tiền thuế'] += it['tien_thue']
                
            df_m = pd.DataFrame([
                {
                    'Kỳ phát điện': p,
                    'Số lượng HĐ': s['Tổng HĐ'],
                    'HĐ Có thuế': s['Có thuế'],
                    'HĐ Không thuế': s['Không thuế'],
                    'Tổng tiền trước thuế (VNĐ)': f"{s['Tiền trước thuế']:,}".replace(',', ' '),
                    'Tổng tiền thuế (VNĐ)': f"{s['Tiền thuế']:,}".replace(',', ' ')
                }
                for p, s in sorted(p_stats.items(), reverse=True)
            ])
            st.dataframe(df_m, use_container_width=True)

    with tab_dl:
        if valid_items and template_file_bytes:
            st.markdown("### 🎯 Xuất Dữ Liệu Ra File Excel Chuẩn")
            st.markdown("""
            * **Sheet tổng:** `DATALOAD TỔNG`
            * **Các sheet tháng:** `T08-2026`, `T07-2026`...
            * **Quy tắc sắp xếp:** Hóa đơn có thuế lên trên $\\rightarrow$ Hóa đơn không thuế xuống dưới (tô vàng toàn dòng).
            * **Định dạng:** Ngày `dd/mm/yyyy`, Tiền số có dấu phân cách `#,#0`.
            """)
            
            now_tag = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            export_name = st.text_input("Tên file Excel khi tải về:", value=f"DataLoad_MTMN_{now_tag}.xlsx")
            
            col_b1, col_b2 = st.columns([1, 2])
            with col_b1:
                if st.button("🚀 Bắt Đầu Tạo File Excel", type="primary", use_container_width=True):
                    with st.spinner("Đang áp dụng mẫu và tạo các sheet Excel..."):
                        try:
                            buf = generate_excel_bytes(valid_items, template_file_bytes)
                            st.session_state['dl_excel_buffer'] = buf
                            st.session_state['dl_excel_name'] = export_name
                            st.success("✅ Đã tạo file Excel thành công!")
                        except Exception as err:
                            st.error(f"Lỗi: {err}")
                            
            with col_b2:
                if 'dl_excel_buffer' in st.session_state:
                    st.download_button(
                        label=f"📥 TẢI XUỐNG FILE EXCEL: {st.session_state.get('dl_excel_name', 'DataLoad.xlsx')}",
                        data=st.session_state['dl_excel_buffer'],
                        file_name=st.session_state.get('dl_excel_name', 'DataLoad.xlsx'),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
else:
    st.info("👋 Hãy kéo thả hoặc chọn các file hóa đơn (.XML hoặc .PDF) để bắt đầu.")
