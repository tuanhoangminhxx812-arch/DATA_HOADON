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
    
    # 1. Tìm mẫu "tháng MM/YYYY" hoặc "tháng MM-YYYY"
    m_thang = re.search(r'(?:tháng|kỳ|T)\s*0?(\d{1,2})\s*[-/]\s*(\d{4})', cleaned, re.IGNORECASE)
    if m_thang:
        return f'T{int(m_thang.group(1)):02d}-{m_thang.group(2)}'
        
    # 2. Tìm mẫu "từ ngày ... đến ngày DD/MM/YYYY"
    m_range = re.search(r'(?:từ\s*ngày.*?đến\s*ngày|đến\s*ngày)\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})', cleaned, re.IGNORECASE)
    if m_range:
        return f'T{int(m_range.group(2)):02d}-{m_range.group(3)}'

    # 3. Dùng ngày lập hóa đơn NLap
    if nlap:
        if isinstance(nlap, (datetime.datetime, datetime.date)):
            return f'T{nlap.month:02d}-{nlap.year}'
        parts = str(nlap).split('-')
        if len(parts) >= 2 and len(parts[0]) == 4: # YYYY-MM-DD
            return f'T{int(parts[1]):02d}-{parts[0]}'
        parts_slash = str(nlap).split('/')
        if len(parts_slash) == 3 and len(parts_slash[2]) == 4: # DD/MM/YYYY
            return f'T{int(parts_slash[1]):02d}-{parts_slash[2]}'
            
    # 4. Fallback ngày cuối cùng tìm thấy
    dates = re.findall(r'(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})', cleaned)
    if dates:
        valid_dates = [f'T{int(m):02d}-{y}' for d, m, y in dates if int(y) >= 2020]
        if valid_dates:
            return valid_dates[-1]

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
    Phân tích cú pháp file hóa đơn điện tử XML đầy đủ chi tiết.
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
    
    # 1. Đơn vị bán (Chủ đầu tư)
    ten_nban = (root.find('.//NBan/Ten').text or '').strip() if root.find('.//NBan/Ten') is not None else ''
    mst_nban = (root.find('.//NBan/MST').text or '').strip() if root.find('.//NBan/MST') is not None else ''
    dchi_nban = (root.find('.//NBan/DChi').text or '').strip() if root.find('.//NBan/DChi') is not None else ''
    if mst_nban.isdigit() and not mst_nban.startswith('0'):
        mst_val = int(mst_nban)
    else:
        mst_val = mst_nban

    # 2. Khách hàng (Bên mua)
    ten_nmua = (root.find('.//NMua/Ten').text or '').strip() if root.find('.//NMua/Ten') is not None else ''
    if not ten_nmua and root.find('.//NMua/HVTNMHang') is not None:
        ten_nmua = (root.find('.//NMua/HVTNMHang').text or '').strip()
    mst_nmua = (root.find('.//NMua/MST').text or '').strip() if root.find('.//NMua/MST') is not None else ''
    dchi_nmua = (root.find('.//NMua/DChi').text or '').strip() if root.find('.//NMua/DChi') is not None else ''

    # 3. Danh sách hàng hóa, dịch vụ
    items = []
    for hhdvu in root.findall('.//DSHHDVu/HHDVu'):
        stt_i = (hhdvu.find('STT').text or '').strip() if hhdvu.find('STT') is not None else ''
        th_i = (hhdvu.find('THHDVu').text or '').strip() if hhdvu.find('THHDVu') is not None else ''
        dvt_i = (hhdvu.find('DVTinh').text or '').strip() if hhdvu.find('DVTinh') is not None else ''
        sl_i = (hhdvu.find('SLuong').text or '').strip() if hhdvu.find('SLuong') is not None else ''
        dg_i = (hhdvu.find('DGia').text or '').strip() if hhdvu.find('DGia') is not None else ''
        tt_i = (hhdvu.find('ThTien').text or '').strip() if hhdvu.find('ThTien') is not None else ''
        
        # Bỏ số 0 thừa sau dấu phẩy thập phân nếu là số nguyên
        try:
            if sl_i and float(sl_i).is_integer():
                sl_i = str(int(float(sl_i)))
            if dg_i and float(dg_i).is_integer():
                dg_i = f"{int(float(dg_i)):,}".replace(',', '.')
            if tt_i and float(tt_i).is_integer():
                tt_i = f"{int(float(tt_i)):,}".replace(',', '.')
        except:
            pass

        if th_i or tt_i:
            items.append({
                'stt': stt_i,
                'name': th_i,
                'unit': dvt_i,
                'quantity': sl_i,
                'price': dg_i,
                'amount': tt_i
            })

    # Nếu không có thẻ HHDVu chi tiết thì lấy thhdvu chung
    if not items and thhdvu:
        items.append({
            'stt': '1',
            'name': thhdvu,
            'unit': 'kWh',
            'quantity': '',
            'price': '',
            'amount': ''
        })

    # Ghép chuỗi để hiển thị 1 ô trong Excel
    if len(items) > 1:
        names_str = '\n'.join([f"{idx+1}. {it['name']}" for idx, it in enumerate(items)])
        units_str = '\n'.join([it['unit'] for it in items])
        qtys_str = '\n'.join([str(it['quantity']) for it in items])
        prices_str = '\n'.join([str(it['price']) for it in items])
        amounts_str = '\n'.join([str(it['amount']) for it in items])
    elif len(items) == 1:
        names_str = items[0]['name']
        units_str = items[0]['unit']
        qtys_str = str(items[0]['quantity'])
        prices_str = str(items[0]['price'])
        amounts_str = str(items[0]['amount'])
    else:
        names_str, units_str, qtys_str, prices_str, amounts_str = '', '', '', '', ''

    # 4. Tài chính
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

    tgtttbso = root.find('.//TgTTTBSo')
    if tgtttbso is not None and tgtttbso.text:
        tong_tien_tt = round(float(tgtttbso.text))
    else:
        tong_tien_tt = tien_truoc_thue + tien_thue

    tgtttbchu = root.find('.//TgTTTBChu')
    tien_bang_chu = (tgtttbchu.text or '').strip() if tgtttbchu is not None else ''
    
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
        'is_taxable': is_taxable,
        # Thông tin mở rộng đầy đủ
        'seller_name': ten_nban,
        'seller_mst': mst_val,
        'seller_addr': dchi_nban,
        'buyer_name': ten_nmua,
        'buyer_mst': mst_nmua,
        'buyer_addr': dchi_nmua,
        'items': items,
        'item_name_str': names_str,
        'item_unit_str': units_str,
        'item_qty_str': qtys_str,
        'item_price_str': prices_str,
        'item_amount_str': amounts_str,
        'tong_tien_tt': tong_tien_tt,
        'tien_bang_chu': tien_bang_chu,
        'ghi_chu': ''
    }

def parse_pdf_invoice(content_bytes, filename=""):
    """
    Phân tích cú pháp file hóa đơn điện tử PDF (TT78 / NĐ123) đầy đủ chi tiết.
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
    m_sh = re.search(r'(?:\(No\.\)[:\s]*Số|Số\s*hóa\s*đơn|Số\s*HĐ|Số\s*\(No\.\)|Số/No\.)[:\s]*0*(\d+)', full_text, re.IGNORECASE)
    if m_sh:
        sh_num = int(m_sh.group(1))
    else:
        m_file_sh = re.search(r'-\s*0*(\d+)(?:\.pdf)?$', filename, re.IGNORECASE)
        if m_file_sh:
            sh_num = int(m_file_sh.group(1))
        else:
            m_any_no = re.search(r'\bNo\.[:\s]*0*(\d+)', full_text, re.IGNORECASE)
            if m_any_no:
                sh_num = int(m_any_no.group(1))

    # 3. Ngày hóa đơn
    nlap_date = datetime.datetime.now()
    m_date1 = re.search(r'Ngày(?:\s*\(Date\))?\s*0?(\d{1,2})\s*tháng(?:\s*\(month\))?\s*0?(\d{1,2})\s*năm(?:\s*\(year\))?\s*(\d{4})', full_text, re.IGNORECASE)
    if m_date1:
        d, m, y = int(m_date1.group(1)), int(m_date1.group(2)), int(m_date1.group(3))
        nlap_date = datetime.datetime(y, m, d)
    else:
        m_date2 = re.search(r'(?:Ngày\s*lập|Ngày\s*HĐ|Ngày)[:\s]*0?(\d{1,2})[/.-]0?(\d{1,2})[/.-](\d{4})', full_text, re.IGNORECASE)
        if m_date2:
            d, m, y = int(m_date2.group(1)), int(m_date2.group(2)), int(m_date2.group(3))
            nlap_date = datetime.datetime(y, m, d)

    # 4. Bên bán (Chủ đầu tư)
    ten_nban = ""
    m_seller_start = re.search(r'^(.*?)\s*Mã số thuế', full_text, re.DOTALL)
    if m_seller_start:
        ten_nban = re.sub(r'\s+', ' ', m_seller_start.group(1).strip())
    else:
        m_ten = re.search(r'(?:Đơn\s*vị\s*bán(?:\s*hàng)?|Tên\s*người\s*bán|Người\s*bán)[:\s]*([^\n\r]+)', full_text, re.IGNORECASE)
        if m_ten:
            ten_nban = m_ten.group(1).strip()
        else:
            for line in lines[:15]:
                if line.upper().startswith(('CÔNG TY', 'CTY', 'DOANH NGHIỆP', 'DNTN', 'CHI NHÁNH')):
                    ten_nban = line
                    break

    m_seller_mst = re.search(r'Mã\s*số\s*thuế.*?:\s*([0-9]{10}(?:-[0-9]{3})?)', full_text)
    seller_mst_str = m_seller_mst.group(1) if m_seller_mst else ""
    if not seller_mst_str:
        m_mst_list = re.findall(r'(?:Mã\s*số\s*thuế|MST)(?:\s*\(Tax\s*code\))?[:\s]*([0-9]{10}(?:-[0-9]{3})?)', full_text, re.IGNORECASE)
        seller_mst_str = m_mst_list[0] if m_mst_list else ""

    mst_val = int(seller_mst_str) if seller_mst_str.isdigit() and not seller_mst_str.startswith('0') else seller_mst_str

    m_seller_addr = re.search(r'Địa\s*chỉ\s*(.*?)\s*\(Address\):', full_text, re.IGNORECASE)
    seller_addr = re.sub(r'\s+', ' ', m_seller_addr.group(1).strip()) if m_seller_addr else ""
    if not seller_addr:
        m_addr_gen = re.search(r'Địa\s*chỉ[:\s]*([^\n\r]+)', full_text, re.IGNORECASE)
        seller_addr = m_addr_gen.group(1).strip() if m_addr_gen else ""

    # 5. Khách hàng (Bên mua)
    buyer_name = ""
    m_buyer_name = re.search(r'Tên\s*đơn\s*vị\s*\(Company name\):\s*([^\n]+(?:\n[^\n]+)?)', full_text, re.IGNORECASE)
    if m_buyer_name:
        buyer_name = re.sub(r'\s+', ' ', m_buyer_name.group(1).strip())
    else:
        m_bname2 = re.search(r'(?:Tên\s*người\s*mua|Họ\s*tên\s*người\s*mua|Đơn\s*vị\s*mua)[:\s]*([^\n\r]+)', full_text, re.IGNORECASE)
        buyer_name = m_bname2.group(1).strip() if m_bname2 else ""

    buyer_mst = ""
    m_buyer_mst = re.search(r'\(Tax code\):\s*Mã\s*số\s*thuế\s*([0-9]{10}(?:-[0-9]{3})?)', full_text, re.IGNORECASE)
    if m_buyer_mst:
        buyer_mst = m_buyer_mst.group(1).strip()
    else:
        m_bmst2 = re.search(r'Mã\s*số\s*thuế\s*người\s*mua[:\s]*([0-9]{10}(?:-[0-9]{3})?)', full_text, re.IGNORECASE)
        buyer_mst = m_bmst2.group(1).strip() if m_bmst2 else ""

    buyer_addr = ""
    m_buyer_addr = re.search(r'Địa\s*chỉ\s*\(Address\):\s*([^\n]+)', full_text, re.IGNORECASE)
    if m_buyer_addr:
        buyer_addr = re.sub(r'\s+', ' ', m_buyer_addr.group(1).strip())

    # 6. Danh sách hàng hóa dịch vụ (Dạng bảng hoặc đơn lẻ)
    items = []
    pattern_row = r'^(\d+)\s+(.+?)\s+([A-Za-zÀ-ỹ\w]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)$'
    for line in lines:
        m = re.match(pattern_row, line)
        if m:
            stt_p, name_p, unit_p, qty_p, price_p, total_p = m.groups()
            items.append({
                'stt': stt_p,
                'name': name_p.strip(),
                'unit': unit_p.strip(),
                'quantity': qty_p.strip(),
                'price': price_p.strip(),
                'amount': total_p.strip()
            })

    # Tìm diễn giải / ghi chú trích yếu nếu có
    ghi_chu = ""
    m_gc = re.search(r'(Điều động VTTB[^\n]+(?:\n[^\n]+)*)', full_text)
    if m_gc:
        ghi_chu = re.sub(r'\s+', ' ', m_gc.group(1).split('1 Nối bọc')[0].strip())

    if len(items) > 1:
        names_str = '\n'.join([f"{idx+1}. {it['name']}" for idx, it in enumerate(items)])
        units_str = '\n'.join([it['unit'] for it in items])
        qtys_str = '\n'.join([str(it['quantity']) for it in items])
        prices_str = '\n'.join([str(it['price']) for it in items])
        amounts_str = '\n'.join([str(it['amount']) for it in items])
    elif len(items) == 1:
        names_str = items[0]['name']
        units_str = items[0]['unit']
        qtys_str = str(items[0]['quantity'])
        prices_str = str(items[0]['price'])
        amounts_str = str(items[0]['amount'])
    else:
        # Nếu không có bảng dạng này, dùng nội dung phát hiện
        names_str = ghi_chu if ghi_chu else f'ĐIỆN MTMN {nlap_date.strftime("%m/%Y")}'
        units_str, qtys_str, prices_str, amounts_str = '', '', '', ''

    # 7. Tiền trước thuế, Thuế, Thuế suất
    def parse_money(s):
        if not s: return 0
        cleaned = re.sub(r'[^\d]', '', str(s))
        return int(cleaned) if cleaned else 0

    tien_truoc_thue = 0
    tien_thue = 0
    thue_suat = 8
    
    m_ttt = re.search(r'(?:Cộng\s*tiền\s*hàng|Tổng\s*tiền\s*chưa\s*thuế|Tiền\s*trước\s*thuế|Tổng\s*cộng\s*tiền\s*hàng)(?:\s*\([^\)]*\))?[:\s]*([\d.,]+)', full_text, re.IGNORECASE)
    if m_ttt:
        tien_truoc_thue = parse_money(m_ttt.group(1))
    elif items:
        # Fallback tính tổng thành tiền từ các dòng hàng
        try:
            tien_truoc_thue = sum([parse_money(it['amount']) for it in items])
        except:
            pass

    m_ts = re.search(r'(?:Thuế\s*suất\s*GTGT|Thuế\s*suất|VAT\s*rate)(?:\s*\([^\)]*\))?[:\s]*([0-9%]+|KCT|KKKNT)', full_text, re.IGNORECASE)
    if m_ts:
        ts_raw = m_ts.group(1).upper()
        if 'KCT' in ts_raw or 'KKKNT' in ts_raw:
            thue_suat = 'KCT'
        else:
            m_num = re.search(r'(\d+)', ts_raw)
            thue_suat = int(m_num.group(1)) if m_num else 8

    m_tt = re.search(r'(?:Tiền\s*thuế\s*GTGT|Tiền\s*thuế|VAT\s*amount)(?:\s*\([^\)]*\))?[:\s]*([\d.,]+)', full_text, re.IGNORECASE)
    if m_tt:
        tien_thue = parse_money(m_tt.group(1))
    elif isinstance(thue_suat, (int, float)) and thue_suat > 0 and tien_truoc_thue > 0:
        tien_thue = round(tien_truoc_thue * thue_suat / 100)

    m_tong = re.search(r'Tổng cộng tiền thanh toán\s*([\d.,]+)', full_text, re.IGNORECASE)
    if not m_tong:
        m_tong = re.search(r'(?:Tổng\s*tiền\s*thanh\s*toán|Tổng\s*thanh\s*toán)[:\s]*([\d.,]+)', full_text, re.IGNORECASE)
    tong_tien_tt = parse_money(m_tong.group(1)) if m_tong else (tien_truoc_thue + tien_thue)

    m_chu = re.search(r'Số\s*tiền\s*bằng\s*chữ.*?:\s*([^\n\r]+)', full_text, re.IGNORECASE)
    tien_bang_chu = m_chu.group(1).strip() if m_chu else ''

    # 8. Kỳ phát điện & Nội dung
    period = extract_period(full_text, nlap_date.strftime('%Y-%m-%d'))
    noi_dung = f'Thuế GTGT điện MTMN {period}'
    nd = f'ĐIỆN MTMN {period}'

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
        'is_taxable': is_taxable,
        # Thông tin mở rộng đầy đủ
        'seller_name': ten_nban,
        'seller_mst': mst_val,
        'seller_addr': seller_addr,
        'buyer_name': buyer_name,
        'buyer_mst': buyer_mst,
        'buyer_addr': buyer_addr,
        'items': items,
        'item_name_str': names_str,
        'item_unit_str': units_str,
        'item_qty_str': qtys_str,
        'item_price_str': prices_str,
        'item_amount_str': amounts_str,
        'tong_tien_tt': tong_tien_tt,
        'tien_bang_chu': tien_bang_chu,
        'ghi_chu': ghi_chu
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

def parse_numeric(val):
    """
    Chuyển đổi chuỗi số (có dấu chấm hoặc phẩy) thành số int/float an toàn.
    """
    if val is None or val == '':
        return 0
    if isinstance(val, (int, float)):
        return val
    s = str(val).strip()
    cleaned = re.sub(r'[^\d.,]', '', s)
    if not cleaned:
        return 0
    if '.' in cleaned and ',' in cleaned:
        cleaned = cleaned.replace('.', '').replace(',', '.')
        try:
            return float(cleaned)
        except:
            return 0
    if ',' in cleaned and '.' not in cleaned:
        cleaned = cleaned.replace(',', '.')
        try:
            return float(cleaned)
        except:
            return 0
    if '.' in cleaned:
        parts = cleaned.split('.')
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
            return int(''.join(parts))
        try:
            return float(cleaned)
        except:
            return 0
    try:
        return int(cleaned) if cleaned.isdigit() else float(cleaned)
    except:
        return 0

def create_detail_invoice_sheet(wb, items):
    """
    Tạo sheet riêng biệt 'THÔNG TIN HÓA ĐƠN' chứa toàn bộ thông tin chi tiết:
    - Đơn vị bán (Chủ đầu tư) & Khách hàng (Bên mua): Tên, MST, Địa chỉ
    - Tách riêng mỗi mặt hàng / vật tư là 1 dòng độc lập trong Excel.
    - Thông tin hóa đơn (Số HĐ, Ký hiệu, Ngày HĐ, Bên bán, Bên mua) được giữ chung cho các dòng cùng 1 hóa đơn.
    """
    ws = wb.create_sheet(title='THÔNG TIN HÓA ĐƠN', index=0)
    
    headers = [
        'STT', 'Tên file', 'Ký hiệu HĐ', 'Số HĐ', 'Ngày HĐ',
        'Tên Đơn vị bán (Chủ đầu tư)', 'Mã số thuế ĐV bán', 'Địa chỉ Đơn vị bán',
        'Tên Khách hàng (Bên mua)', 'Mã số thuế Khách hàng', 'Địa chỉ Khách hàng',
        'Tên hàng hóa, dịch vụ', 'Đơn vị tính', 'Số lượng', 'Đơn giá',
        'Thành tiền (chưa thuế)', 'Thuế suất', 'Tiền thuế GTGT', 'Tổng tiền thanh toán',
        'Số tiền bằng chữ', 'Ghi chú / Diễn giải'
    ]
    
    col_widths = {
        1: 6,   # STT
        2: 24,  # Tên file
        3: 13,  # Ký hiệu HĐ
        4: 10,  # Số HĐ
        5: 14,  # Ngày HĐ
        6: 36,  # Tên ĐV bán
        7: 16,  # MST ĐV bán
        8: 36,  # Địa chỉ ĐV bán
        9: 36,  # Tên Khách hàng
        10: 17, # MST Khách hàng
        11: 36, # Địa chỉ Khách hàng
        12: 42, # Tên hàng hóa
        13: 12, # ĐVT
        14: 12, # SL
        15: 15, # Đơn giá
        16: 18, # Thành tiền
        17: 12, # Thuế suất
        18: 16, # Tiền thuế
        19: 20, # Tổng thanh toán
        20: 38, # Bằng chữ
        21: 25  # Ghi chú
    }
    
    for col_idx, width in col_widths.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width
        
    ws.row_dimensions[1].height = 32
    
    header_fill = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid') # Xanh Navy đậm sang trọng
    header_font = Font(name='Times New Roman', size=11, bold=True, color='FFFFFF')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    for col_idx, h_text in enumerate(headers, start=1):
        c = ws.cell(1, col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = header_align
        c.border = thin_border
        
    font_body = Font(name='Times New Roman', size=10)
    font_bold = Font(name='Times New Roman', size=10, bold=True)
    
    align_center_top = Alignment(horizontal='center', vertical='top', wrap_text=True)
    align_left_top = Alignment(horizontal='left', vertical='top', wrap_text=True)
    align_right_top = Alignment(horizontal='right', vertical='top', wrap_text=True)
    
    current_row = 1
    stt_counter = 0

    for item in items:
        item_rows = item.get('items', [])
        # Nếu không có danh sách items con (ví dụ hóa đơn XML chỉ có 1 dòng chung)
        if not item_rows:
            item_rows = [{
                'stt': '1',
                'name': item.get('item_name_str', '') or item.get('noi_dung', ''),
                'unit': item.get('item_unit_str', ''),
                'quantity': item.get('item_qty_str', ''),
                'price': item.get('item_price_str', ''),
                'amount': item.get('tien_truoc_thue', 0)
            }]

        num_sub_items = len(item_rows)
        thue_suat = item.get('thue_suat', 0)
        ts_display = f"{thue_suat}%" if isinstance(thue_suat, int) else str(thue_suat)

        for sub_idx, sub_item in enumerate(item_rows):
            stt_counter += 1
            current_row += 1
            r = current_row

            # Tính toán tài chính cho dòng hàng này
            raw_amount = sub_item.get('amount', 0)
            amount_val = parse_numeric(raw_amount)
            if amount_val == 0 and num_sub_items == 1:
                amount_val = item.get('tien_truoc_thue', 0)

            # Tiền thuế dòng
            if num_sub_items == 1:
                tax_val = item.get('tien_thue', 0)
            elif isinstance(thue_suat, (int, float)) and thue_suat > 0:
                tax_val = round(amount_val * thue_suat / 100)
            else:
                tax_val = 0

            # Tổng thanh toán dòng
            total_val = amount_val + tax_val

            # Số lượng và Đơn giá
            raw_qty = sub_item.get('quantity', '')
            raw_price = sub_item.get('price', '')
            qty_val = parse_numeric(raw_qty) if raw_qty != '' else ''
            price_val = parse_numeric(raw_price) if raw_price != '' else ''

            # 1. STT dòng
            c1 = ws.cell(r, 1, value=stt_counter)
            c1.alignment = align_center_top
            
            # 2. Tên file
            c2 = ws.cell(r, 2, value=item.get('filename', ''))
            c2.alignment = align_left_top
            
            # 3. Ký hiệu HĐ
            c3 = ws.cell(r, 3, value=item.get('ky_hieu_hd', ''))
            c3.alignment = align_center_top
            
            # 4. Số HĐ
            c4 = ws.cell(r, 4, value=item.get('so_hd', ''))
            c4.alignment = align_center_top
            
            # 5. Ngày HĐ
            c5 = ws.cell(r, 5, value=item.get('ngay_hd', ''))
            c5.alignment = align_center_top
            c5.number_format = 'dd/mm/yyyy'
            
            # 6. Tên Đơn vị bán (Chủ đầu tư)
            c6 = ws.cell(r, 6, value=item.get('seller_name', ''))
            c6.alignment = align_left_top
            
            # 7. MST ĐV bán
            c7 = ws.cell(r, 7, value=str(item.get('seller_mst', '')))
            c7.alignment = align_center_top
            
            # 8. Địa chỉ ĐV bán
            c8 = ws.cell(r, 8, value=item.get('seller_addr', ''))
            c8.alignment = align_left_top
            
            # 9. Tên Khách hàng (Bên mua)
            c9 = ws.cell(r, 9, value=item.get('buyer_name', ''))
            c9.alignment = align_left_top
            
            # 10. MST Khách hàng
            c10 = ws.cell(r, 10, value=str(item.get('buyer_mst', '')))
            c10.alignment = align_center_top
            
            # 11. Địa chỉ Khách hàng
            c11 = ws.cell(r, 11, value=item.get('buyer_addr', ''))
            c11.alignment = align_left_top
            
            # 12. Tên hàng hóa dịch vụ (riêng từng mặt hàng)
            c12 = ws.cell(r, 12, value=sub_item.get('name', ''))
            c12.alignment = align_left_top
            
            # 13. Đơn vị tính
            c13 = ws.cell(r, 13, value=sub_item.get('unit', ''))
            c13.alignment = align_center_top
            
            # 14. Số lượng
            c14 = ws.cell(r, 14, value=qty_val if qty_val != '' else raw_qty)
            c14.alignment = align_right_top
            if isinstance(qty_val, (int, float)) and qty_val != '':
                c14.number_format = '#,##0.##' if isinstance(qty_val, float) and not qty_val.is_integer() else '#,##0'

            # 15. Đơn giá
            c15 = ws.cell(r, 15, value=price_val if price_val != '' else raw_price)
            c15.alignment = align_right_top
            if isinstance(price_val, (int, float)) and price_val != '':
                c15.number_format = '#,##0.###' if isinstance(price_val, float) and not price_val.is_integer() else '#,##0'
            
            # 16. Thành tiền (chưa thuế)
            c16 = ws.cell(r, 16, value=amount_val)
            c16.alignment = align_right_top
            c16.number_format = '#,##0'
            
            # 17. Thuế suất
            c17 = ws.cell(r, 17, value=ts_display)
            c17.alignment = align_center_top
            
            # 18. Tiền thuế GTGT
            c18 = ws.cell(r, 18, value=tax_val)
            c18.alignment = align_right_top
            c18.number_format = '#,##0'
            
            # 19. Tổng tiền thanh toán
            c19 = ws.cell(r, 19, value=total_val)
            c19.alignment = align_right_top
            c19.number_format = '#,##0'
            
            # 20. Số tiền bằng chữ
            c20 = ws.cell(r, 20, value=item.get('tien_bang_chu', ''))
            c20.alignment = align_left_top
            
            # 21. Ghi chú
            c21 = ws.cell(r, 21, value=item.get('ghi_chu', ''))
            c21.alignment = align_left_top
            
            for c_idx in range(1, len(headers) + 1):
                cell = ws.cell(r, c_idx)
                cell.font = font_body
                cell.border = thin_border

    # Dòng Tổng Cộng
    total_row = current_row + 1
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=15)
    c_tot_label = ws.cell(total_row, 1, value="TỔNG CỘNG")
    c_tot_label.alignment = Alignment(horizontal='center', vertical='center')
    c_tot_label.font = font_bold
    
    # SUM Thành tiền (cột P)
    c_tot_ttt = ws.cell(total_row, 16, value=f"=SUM(P2:P{total_row-1})")
    c_tot_ttt.alignment = align_right_top
    c_tot_ttt.number_format = '#,##0'
    c_tot_ttt.font = font_bold
    
    # SUM Tiền thuế (cột R)
    c_tot_tt = ws.cell(total_row, 18, value=f"=SUM(R2:R{total_row-1})")
    c_tot_tt.alignment = align_right_top
    c_tot_tt.number_format = '#,##0'
    c_tot_tt.font = font_bold
    
    # SUM Tổng thanh toán (cột S)
    c_tot_all = ws.cell(total_row, 19, value=f"=SUM(S2:S{total_row-1})")
    c_tot_all.alignment = align_right_top
    c_tot_all.number_format = '#,##0'
    c_tot_all.font = font_bold
    
    tot_fill = PatternFill(start_color='F1F5F9', end_color='F1F5F9', fill_type='solid')
    for c_idx in range(1, len(headers) + 1):
        cell = ws.cell(total_row, c_idx)
        cell.fill = tot_fill
        cell.border = thin_border
        
    # Cố định dòng 1 và 4 cột đầu
    ws.freeze_panes = 'E2'

def generate_excel_bytes(valid_items, template_bytes_or_path):
    """
    Sinh file Excel hoàn chỉnh:
    - Sheet 1 (MỚI): 'THÔNG TIN HÓA ĐƠN' - Lấy toàn bộ thông tin chi tiết hóa đơn (mỗi HĐ là 1 dòng, mỗi thông tin là 1 cột).
    - Sheet 2: 'DATALOAD TỔNG' và từng sheet theo từng tháng (giữ nguyên chuẩn mẫu DataLoad hạch toán).
    """
    if isinstance(template_bytes_or_path, str):
        wb_template = openpyxl.load_workbook(template_bytes_or_path)
    else:
        wb_template = openpyxl.load_workbook(io.BytesIO(template_bytes_or_path))
        
    ws_tmpl = wb_template.active
    
    wb_new = openpyxl.Workbook()
    wb_new.remove(wb_new.active)
    
    # 1. TẠO SHEET MỚI: BẢNG KÊ TOÀN BỘ THÔNG TIN HÓA ĐƠN
    create_detail_invoice_sheet(wb_new, valid_items)
    
    # 2. TẠO CÁC SHEET DATALOAD
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
            
        # Sao chép dòng 1 (Header) - Bỏ toàn bộ màu tô vàng để giao diện sạch đẹp
        for col_idx in range(1, ws_tmpl.max_column + 1):
            src_cell = ws_tmpl.cell(1, col_idx)
            dest_cell = ws.cell(1, col_idx, value=src_cell.value)
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

            # XỬ LÝ MÀU THEO YÊU CẦU:
            # 1. Hóa đơn CÓ THUẾ: Bỏ toàn bộ màu vàng ở tất cả các cột để dữ liệu sạch đẹp tự nhiên.
            # 2. Hóa đơn KHÔNG THUẾ: Tô vàng toàn bộ dòng (A -> AH) để anh dễ dàng nhận biết các hóa đơn không thuế.
            if not is_taxable:
                for col_i in range(1, 35):
                    ws.cell(r, col_i).fill = yellow_fill
                    
    output_buffer = io.BytesIO()
    wb_new.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer

# ----------------- STREAMLIT INTERFACE ----------------- #

st.markdown('<div class="main-header">⚡ Xử Lý Hóa Đơn Điện Tử MTMN (.XML & .PDF)</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Trích xuất toàn bộ thông tin hóa đơn (Bên bán, Bên mua, Hàng hóa, Thuế, Tổng tiền), tự động tạo sheet riêng và nạp mẫu DataLoad chuẩn.</div>', unsafe_allow_html=True)

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
    tab_detail, tab_view, tab_by_month, tab_dl = st.tabs([
        "📑 Bảng Kê Toàn Bộ Hóa Đơn (Mới)",
        "👁️ Bảng Dữ Liệu DataLoad",
        "📅 Thống Kê Theo Tháng",
        "🚀 Xuất & Tải File Excel"
    ])
    
    with tab_detail:
        st.markdown("##### 📋 Toàn bộ thông tin chi tiết hóa đơn (Tách riêng mỗi mặt hàng / vật tư là 1 dòng)")
        st.caption("Mỗi mặt hàng được tách thành 1 dòng riêng biệt; thông tin hóa đơn (Số HĐ, Ký hiệu, Ngày HĐ, Bên bán, Bên mua) được giữ chung.")
        if valid_items:
            flattened_rows = []
            for it in valid_items:
                sub_items = it.get('items', [])
                if not sub_items:
                    sub_items = [{
                        'stt': '1',
                        'name': it.get('item_name_str', '') or it.get('noi_dung', ''),
                        'unit': it.get('item_unit_str', ''),
                        'quantity': it.get('item_qty_str', ''),
                        'price': it.get('item_price_str', ''),
                        'amount': it.get('tien_truoc_thue', 0)
                    }]
                num_sub = len(sub_items)
                ts = it.get('thue_suat', 0)
                
                for sub in sub_items:
                    raw_amt = sub.get('amount', 0)
                    amt = parse_numeric(raw_amt)
                    if amt == 0 and num_sub == 1:
                        amt = it.get('tien_truoc_thue', 0)
                    
                    if num_sub == 1:
                        tax = it.get('tien_thue', 0)
                    elif isinstance(ts, (int, float)) and ts > 0:
                        tax = round(amt * ts / 100)
                    else:
                        tax = 0
                    tot = amt + tax
                    
                    raw_q = sub.get('quantity', '')
                    raw_p = sub.get('price', '')
                    q_val = parse_numeric(raw_q) if raw_q != '' else ''
                    p_val = parse_numeric(raw_p) if raw_p != '' else ''
                    
                    q_disp = f"{q_val:,}".replace(',', '.') if isinstance(q_val, int) else (f"{q_val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if isinstance(q_val, float) else str(raw_q))
                    p_disp = f"{p_val:,}".replace(',', '.') if isinstance(p_val, int) else (f"{p_val:,.3f}".replace(',', 'X').replace('.', ',').replace('X', '.') if isinstance(p_val, float) else str(raw_p))

                    flattened_rows.append({
                        'STT': len(flattened_rows) + 1,
                        'Tên file': it['filename'],
                        'Ký hiệu HĐ': it['ky_hieu_hd'],
                        'Số HĐ': it['so_hd'],
                        'Ngày HĐ': it['ngay_hd'].strftime('%d/%m/%Y') if isinstance(it['ngay_hd'], (datetime.datetime, datetime.date)) else str(it['ngay_hd']),
                        'ĐV bán (Chủ đầu tư)': it['seller_name'],
                        'MST ĐV bán': it['seller_mst'],
                        'Địa chỉ ĐV bán': it['seller_addr'],
                        'Khách hàng (Bên mua)': it['buyer_name'],
                        'MST Khách hàng': it['buyer_mst'],
                        'Địa chỉ Khách hàng': it['buyer_addr'],
                        'Tên hàng hóa dịch vụ': sub.get('name', ''),
                        'ĐVT': sub.get('unit', ''),
                        'Số lượng': q_disp,
                        'Đơn giá': p_disp,
                        'Thành tiền (chưa thuế)': f"{amt:,}".replace(',', ' '),
                        'Thuế suất': f"{ts}%" if isinstance(ts, int) else str(ts),
                        'Tiền thuế': f"{tax:,}".replace(',', ' '),
                        'Tổng thanh toán': f"{tot:,}".replace(',', ' '),
                        'Ghi chú': it.get('ghi_chu', '')
                    })
            df_detail = pd.DataFrame(flattened_rows)
            st.dataframe(df_detail, use_container_width=True, height=500)

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
            * **Sheet 1 (MỚI):** `THÔNG TIN HÓA ĐƠN` - Chứa toàn bộ thông tin chi tiết: Đơn vị bán & Khách hàng (Tên, MST, Địa chỉ), Hàng hóa, ĐVT, Số lượng, Đơn giá, Thành tiền, Thuế suất, Tiền thuế, Tổng tiền.
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
