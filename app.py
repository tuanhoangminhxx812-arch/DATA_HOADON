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

# Set wide page layout & modern title
st.set_page_config(
    page_title="Hóa Đơn MTMN - Tách & Xuất DataLoad",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
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
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F3F4F6;
        padding: 1rem;
        border-radius: 8px;
        border-left: 5px solid #2563EB;
    }
    .error-card {
        background-color: #FEF2F2;
        padding: 1rem;
        border-radius: 8px;
        border-left: 5px solid #DC2626;
        margin-bottom: 1rem;
    }
    .success-badge {
        background-color: #DEF7EC;
        color: #03543F;
        padding: 4px 8px;
        border-radius: 4px;
        font-weight: 600;
    }
    .warning-badge {
        background-color: #FEF08A;
        color: #854D0E;
        padding: 4px 8px;
        border-radius: 4px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# ----------------- HELPER FUNCTIONS ----------------- #

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
    Xử lý thuế suất từ XML.
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

def parse_single_xml(content_bytes, filename=""):
    """
    Phân tích cú pháp một file XML hóa đơn điện tử.
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
        
    # Tiền trước thuế
    tgtcthue = root.find('.//TgTCThue')
    if tgtcthue is not None and tgtcthue.text:
        tien_truoc_thue = round(float(tgtcthue.text))
    else:
        thtien = root.find('.//ThTien')
        tien_truoc_thue = round(float(thtien.text)) if thtien is not None and thtien.text else 0
        
    # Tiền thuế
    tgtthue = root.find('.//TgTThue')
    if tgtthue is not None and tgtthue.text:
        tien_thue = round(float(tgtthue.text))
    else:
        tthue = root.find('.//TThue')
        tien_thue = round(float(tthue.text)) if tthue is not None and tthue.text else 0
        
    # Thuế suất
    tsuat = root.find('.//TSuat')
    tsuat_val = clean_tax_rate(tsuat.text if tsuat is not None else '8')
    
    # Xác định có thuế hay không có thuế
    # Có thuế: tiền thuế > 0 và thuế suất không phải KCT / 0%
    is_taxable = (tien_thue > 0) and (tsuat_val not in [0, '0', 'KCT', 'KKKNT'])
    
    noi_dung = f'Thuế GTGT điện MTMN {period}'
    nd = f'ĐIỆN MTMN {period}'
    
    return {
        'filename': filename,
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

def generate_excel_bytes(valid_items, template_bytes_or_path):
    """
    Sinh file Excel hoàn chỉnh định dạng openpyxl:
    - Có sheet 'DATALOAD TỔNG' và từng sheet cho từng tháng.
    - Sắp xếp: Hóa đơn có thuế lên trên, hóa đơn không thuế xuống dưới.
    - Hóa đơn không thuế được TÔ VÀNG TOÀN BỘ DÒNG để nhận biết rõ ràng.
    """
    if isinstance(template_bytes_or_path, str):
        wb_template = openpyxl.load_workbook(template_bytes_or_path)
    else:
        wb_template = openpyxl.load_workbook(io.BytesIO(template_bytes_or_path))
        
    ws_tmpl = wb_template.active
    
    wb_new = openpyxl.Workbook()
    wb_new.remove(wb_new.active) # Remove sheet mặc định
    
    # Sắp xếp các kỳ tháng giảm dần: T08, T07, T06...
    periods = sorted(list(set(item['period'] for item in valid_items)), reverse=True)
    
    sheets_to_create = [('DATALOAD TỔNG', valid_items)]
    for p in periods:
        items_p = [item for item in valid_items if item['period'] == p]
        sheets_to_create.append((p, items_p))
        
    # Styles
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

        # TÁCH VÀ SẮP XẾP:
        # Nhóm 1: Hóa đơn CÓ THUẾ (để lên trên)
        # Nhóm 2: Hóa đơn KHÔNG THUẾ (để xuống dưới)
        taxable_items = [it for it in raw_items if it['is_taxable']]
        nontaxable_items = [it for it in raw_items if not it['is_taxable']]
        sorted_items = taxable_items + nontaxable_items

        # Ghi các dòng dữ liệu
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
            
            # Col D: Thuế
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
            
            # Col O: Ngày HĐ
            cO = ws.cell(r, 15, value=item['ngay_hd'])
            cO.font = font_tnr
            cO.number_format = 'yyyy-mm-dd'
            
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
            
            # Col Y: Số tiền trước thuế
            cY = ws.cell(r, 25, value=item['tien_truoc_thue'])
            cY.font = font_tnr
            cY.number_format = '#,##0'
            
            # Col Z: TAB
            ws.cell(r, 26, value='TAB').font = font_tnr
            
            # Col AA: Thuế suất
            cAA = ws.cell(r, 27, value=item['thue_suat'])
            cAA.font = font_tnr
            
            # Col AB: TAB
            ws.cell(r, 28, value='TAB').font = font_tnr
            
            # Col AC, AD, AE: CK, tab, ENT
            ws.cell(r, 29, value='CK').font = font_tnr
            ws.cell(r, 30, value='tab').font = font_tnr
            ws.cell(r, 31, value='ENT').font = font_tnr
            
            # Col AF: None
            ws.cell(r, 32, value=None)
            
            # Col AG: test MST
            ws.cell(r, 33, value=f'=LEN(U{r})').font = font_tnr
            
            # Col AH: test VAT
            ws.cell(r, 34, value=f'=ROUND(D{r}-(Y{r}*(AA{r}/100)),0)').font = font_tnr

            # XỬ LÝ TÔ MÀU VÀNG THEO YÊU CẦU:
            # 1. Nếu là HÓA ĐƠN CÓ THUẾ: Tô vàng các cột D, F, K, M, O, S, U, Y, AA (chuẩn theo mẫu)
            # 2. Nếu là HÓA ĐƠN KHÔNG THUẾ (thuế = 0): Tô vàng TOÀN BỘ CÁC CỘT (A -> AH) để nhận biết ngay!
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

# ----------------- STREAMLIT UI ----------------- #

# Sidebar
with st.sidebar:
    st.image("https://img.icons8.com/color/96/solar-panel.png", width=70)
    st.title("⚡ Cấu Hình & Tùy Chọn")
    
    st.markdown("---")
    st.subheader("📄 Mẫu File Excel (Template)")
    default_template_path = r'd:\DATA\DATA_HOADON\DataLoad_MTMN_T09 (MAU).xlsx'
    template_option = st.radio(
        "Nguồn mẫu Excel:",
        ["Dùng mẫu mặc định trên máy", "Tải lên mẫu Excel khác"],
        index=0
    )
    
    template_file_bytes = None
    if template_option == "Dùng mẫu mặc định trên máy":
        if os.path.exists(default_template_path):
            st.caption(f"✓ Sử dụng mẫu: `{os.path.basename(default_template_path)}`")
            with open(default_template_path, 'rb') as f:
                template_file_bytes = f.read()
        else:
            st.error("Không tìm thấy file mẫu mặc định!")
    else:
        uploaded_template = st.file_uploader("Tải lên file mẫu Excel (.xlsx)", type=["xlsx"])
        if uploaded_template:
            template_file_bytes = uploaded_template.read()
            st.success("Đã tải mẫu Excel thành công!")
            
    st.markdown("---")
    st.markdown("### 📌 Ghi chú nghiệp vụ:")
    st.info("""
    - **HĐ có thuế**: Xếp lên trên trong từng sheet.
    - **HĐ không thuế (0% / KCT)**: Xếp xuống dưới cùng và **TÔ VÀNG TOÀN DÒNG** để nhận biết.
    - Tự động bóc tách **kỳ tiền điện** đưa vào từng sheet riêng theo tháng.
    """)

# Main Content
st.markdown('<div class="main-header">⚡ Xử Lý Hóa Đơn Điện Tử MTMN - Tách & Xuất DataLoad</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Đọc hóa đơn XML, kiểm tra tính hợp lệ, phân loại hóa đơn có thuế/không thuế và xuất file Excel theo chuẩn kế toán.</div>', unsafe_allow_html=True)

# Input method tabs
tab_upload, tab_folder = st.tabs(["📤 Tải File XML Lên Trực Tiếp", "📁 Đọc Từ Thư Mục Máy Tính"])

raw_xml_data = [] # List of tuples (filename, bytes_content)

with tab_upload:
    uploaded_files = st.file_uploader(
        "Kéo thả hoặc chọn các file XML hóa đơn cần xử lý:",
        type=["xml"],
        accept_multiple_files=True,
        help="Bạn có thể chọn cùng lúc nhiều file XML hoặc nhấn Ctrl+A để chọn toàn bộ."
    )
    if uploaded_files:
        for uf in uploaded_files:
            raw_xml_data.append((uf.name, uf.read()))

with tab_folder:
    folder_path = st.text_input(
        "Nhập đường dẫn thư mục chứa file XML trên máy tính:",
        value=r"d:\DATA\DATA_HOADON\XML Tháng 8 - Đợt 1"
    )
    col_btn, _ = st.columns([1, 3])
    with col_btn:
        read_folder_btn = st.button("🔍 Quét Thư Mục", use_container_width=True)
        
    if read_folder_btn:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            files = [f for f in os.listdir(folder_path) if f.lower().endswith('.xml')]
            if files:
                raw_xml_data = []
                for fname in files:
                    fpath = os.path.join(folder_path, fname)
                    try:
                        with open(fpath, 'rb') as f:
                            raw_xml_data.append((fname, f.read()))
                    except Exception as e:
                        st.error(f"Lỗi đọc file {fname}: {e}")
                st.session_state['folder_raw_xml'] = raw_xml_data
                st.success(f"Đã tìm thấy **{len(raw_xml_data)}** file XML trong thư mục!")
            else:
                st.warning("Không tìm thấy file .xml nào trong thư mục được chọn!")
        else:
            st.error("Thư mục không tồn tại. Vui lòng kiểm tra lại đường dẫn.")
            
    if 'folder_raw_xml' in st.session_state and not raw_xml_data:
        raw_xml_data = st.session_state['folder_raw_xml']

st.markdown("---")

# Process XML Data
if raw_xml_data:
    valid_items = []
    invalid_files = []
    
    for filename, content_bytes in raw_xml_data:
        try:
            item = parse_single_xml(content_bytes, filename)
            valid_items.append(item)
        except Exception as e:
            invalid_files.append({
                'Tên file': filename,
                'Kích thước': f"{len(content_bytes):,} bytes",
                'Chi tiết lỗi': str(e)
            })

    # Summary Statistics
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("Tổng số file XML", len(raw_xml_data))
    with col2:
        st.metric("Hóa đơn hợp lệ", len(valid_items))
    with col3:
        st.metric("File bị lỗi / Hỏng", len(invalid_files), delta=-len(invalid_files) if invalid_files else 0, delta_color="inverse")
    with col4:
        taxable_cnt = sum(1 for it in valid_items if it['is_taxable'])
        st.metric("HĐ Có thuế", taxable_cnt)
    with col5:
        nontax_cnt = len(valid_items) - taxable_cnt
        st.metric("HĐ Không thuế", nontax_cnt)
        
    # Warning for invalid files
    if invalid_files:
        st.markdown('<div class="error-card">', unsafe_allow_html=True)
        st.error(f"⚠️ **Phát hiện {len(invalid_files)} file bị lỗi cấu trúc XML!** Các file này sẽ bị bỏ qua, vui lòng kiểm tra và chép lại:")
        df_invalid = pd.DataFrame(invalid_files)
        st.dataframe(df_invalid, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.success("✅ Toàn bộ các file XML đều có cấu trúc hợp lệ!")

    # Tabs for View & Export
    tab_preview, tab_months, tab_export = st.tabs(["👁️ Xem Trước Dữ Liệu", "📅 Thống Kê Theo Tháng", "📥 Xuất File Excel"])
    
    with tab_preview:
        if valid_items:
            df_preview = pd.DataFrame([
                {
                    'File': it['filename'],
                    'Kỳ': it['period'],
                    'Loại HĐ': '🟢 Có thuế' if it['is_taxable'] else '🟡 Không thuế',
                    'Ký hiệu': it['ky_hieu_hd'],
                    'Số HĐ': it['so_hd'],
                    'Ngày HĐ': it['ngay_hd'].strftime('%Y-%m-%d') if isinstance(it['ngay_hd'], (datetime.datetime, datetime.date)) else str(it['ngay_hd']),
                    'Tên người bán': it['ten_kh'],
                    'MST': it['mst'],
                    'Tiền trước thuế': f"{it['tien_truoc_thue']:,}",
                    'Thuế suất': f"{it['thue_suat']}%" if isinstance(it['thue_suat'], int) else it['thue_suat'],
                    'Tiền thuế': f"{it['tien_thue']:,}",
                }
                for it in valid_items
            ])
            st.dataframe(df_preview, use_container_width=True, height=400)
            
    with tab_months:
        if valid_items:
            period_stats = {}
            for it in valid_items:
                p = it['period']
                if p not in period_stats:
                    period_stats[p] = {'Tổng HĐ': 0, 'Có thuế': 0, 'Không thuế': 0, 'Tổng tiền trước thuế': 0, 'Tổng thuế': 0}
                period_stats[p]['Tổng HĐ'] += 1
                if it['is_taxable']:
                    period_stats[p]['Có thuế'] += 1
                else:
                    period_stats[p]['Không thuế'] += 1
                period_stats[p]['Tổng tiền trước thuế'] += it['tien_truoc_thue']
                period_stats[p]['Tổng thuế'] += it['tien_thue']
                
            df_periods = pd.DataFrame([
                {
                    'Kỳ tháng': p,
                    'Số lượng HĐ': stats['Tổng HĐ'],
                    'HĐ Có thuế': stats['Có thuế'],
                    'HĐ Không thuế': stats['Không thuế'],
                    'Tổng tiền trước thuế (VNĐ)': f"{stats['Tổng tiền trước thuế']:,}",
                    'Tổng tiền thuế (VNĐ)': f"{stats['Tổng thuế']:,}"
                }
                for p, stats in sorted(period_stats.items(), reverse=True)
            ])
            st.dataframe(df_periods, use_container_width=True)

    with tab_export:
        if valid_items:
            st.markdown("### 🎯 Xuất Dữ Liệu Ra File Excel")
            st.write("""
            Khi bấm nút bên dưới, hệ thống sẽ:
            1. Tạo sheet **`DATALOAD TỔNG`** chứa toàn bộ các hóa đơn.
            2. Tạo các sheet riêng biệt cho từng tháng: **`T08-2026`**, **`T07-2026`**, v.v.
            3. Trong mỗi sheet, **hóa đơn có thuế được xếp lên trên**, **hóa đơn không thuế được xếp xuống dưới và tô vàng toàn bộ dòng**.
            4. Điền đầy đủ các cột tô vàng và các cột cố định (Mã TK, TAB, CK, ENT, công thức kiểm tra).
            """)
            
            if template_file_bytes is None:
                st.error("⚠️ Chưa có file mẫu Excel. Vui lòng kiểm tra thanh cài đặt bên trái!")
            else:
                col_exp1, col_exp2 = st.columns([1, 2])
                with col_exp1:
                    now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    default_out_name = f"DataLoad_MTMN_{now_str}.xlsx"
                    out_filename = st.text_input("Tên file tải về:", value=default_out_name)
                    
                    if st.button("🚀 Bắt Đầu Tạo File Excel", type="primary", use_container_width=True):
                        with st.spinner("Đang xử lý dữ liệu và tạo các sheet Excel..."):
                            try:
                                excel_buffer = generate_excel_bytes(valid_items, template_file_bytes)
                                st.session_state['generated_excel'] = excel_buffer
                                st.session_state['generated_filename'] = out_filename
                                st.success("✅ Đã tạo file Excel thành công!")
                            except Exception as ex:
                                st.error(f"Lỗi khi tạo file Excel: {ex}")
                                
                with col_exp2:
                    if 'generated_excel' in st.session_state:
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.download_button(
                            label=f"📥 TẢI XUỐNG FILE EXCEL: {st.session_state.get('generated_filename', 'DataLoad.xlsx')}",
                            data=st.session_state['generated_excel'],
                            file_name=st.session_state.get('generated_filename', 'DataLoad.xlsx'),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
else:
    st.info("👋 Hãy tải các file hóa đơn XML lên hoặc nhập đường dẫn thư mục để bắt đầu xử lý.")
