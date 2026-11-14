import os
import re
import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import legal, A4
from reportlab.pdfgen import canvas
from datetime import datetime
from reportlab.platypus import Paragraph, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Explicitly register Algerian Font from Windows Fonts directory
try:
    font_path = r"C:\Windows\Fonts\alger.ttf"
    if not os.path.exists(font_path):
        font_path = r"C:\Windows\Fonts\ALGER.TTF"
    
    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont('Algerian', font_path))
    else:
        windir = os.environ.get('WINDIR', 'C:\\Windows')
        alt_path = os.path.join(windir, 'Fonts', 'alger.ttf')
        if os.path.exists(alt_path):
            pdfmetrics.registerFont(TTFont('Algerian', alt_path))
except Exception as e:
    print(f"Font registration error: {e}")

st.set_page_config(page_title="Holy Child School - Management Hub", layout="wide")

# ==========================================
# 📊 EXCEL COLUMN MAPPING CONFIGURATION
# ==========================================
COL_REG_NO = 'C'
COL_CLASS = 'H'  
COL_DATE_ADMC = 'D'
COL_NAME = 'M'
COL_DOB = 'N'
COL_SCHOOL = 'F'
COL_STATE = 'DV'
COL_HOMETOWN = 'DU'
COL_MOTHER_TONGUE = 'DN'
COL_RELIGION = 'P'
COL_CASTE = 'Q'
COL_APAR = 'O'
COL_CATEGORY = 'DM'
COL_AADHAR = 'R'
COL_BLOOD = 'DP'
COL_CONTACT = 'L'
COL_EMAIL = 'K'

COL_FATHER_NAME = 'S'
COL_FATHER_QUAL = 'AA'
COL_FATHER_OCC = 'AC'        
COL_FATHER_OCC_EXTRA = 'X'  
COL_FATHER_OFFICE = 'AE'
COL_FATHER_CONT = 'AM'
COL_FATHER_AADHAAR = 'U'      

COL_MOTHER_NAME = 'AQ'
COL_MOTHER_QUAL = 'AU'        
COL_MOTHER_OCC = 'AW'
COL_MOTHER_OFFICE = 'AZ'
COL_MOTHER_CONT = 'BG'
COL_MOTHER_AADHAR = 'AS'      

COL_PERM_ADDR = 'DT'
COL_PRES_ADDR = 'DQ'
COL_GUARDIAN = 'DW'
COL_GUARDIAN_REL = 'CS'
COL_GUARDIAN_CONT = 'DX'      
COL_SIB_NAME = 'BK'
COL_SIB_ADM = 'BL'
COL_SIB_CLASS = 'BM'

def col_to_idx(letter):
    if not letter or not str(letter).strip():
        return -1
    col = 0
    for char in str(letter).upper():
        if 'A' <= char <= 'Z':
            col = col * 26 + (ord(char) - ord('A')) + 1
    return col - 1

def format_address(addr_str):
    if not addr_str:
        return ""
    text = str(addr_str).strip()
    text = re.sub(r'^(?:S/O|C/O|D/O|W/O)\s*[:\-]?\s*[^,\n]+,\s*', '', text, flags=re.IGNORECASE)
    formatted = re.sub(r'uttar\s+pradesh', 'U.P.', text, flags=re.IGNORECASE)
    return formatted.strip().title()

def clean_office_address(addr_str):
    if not addr_str:
        return ""
    text = format_address(addr_str)
    text = re.sub(r'\b\d{6}\b', '', text)
    text = re.sub(r'\s*,\s*,\s*', ', ', text)
    return text.strip(" ,-/")

def extract_address_and_pin(addr_str):
    if not addr_str:
        return "", ""
    cleaned_addr = format_address(addr_str)
    match = re.search(r'\b\d{6}\b', cleaned_addr)
    if match:
        pin = match.group(0)
        clean_addr = cleaned_addr.replace(pin, "").strip(" ,-/")
        return clean_addr, pin
    return cleaned_addr, ""

def format_email(email_str):
    if not email_str:
        return ""
    return str(email_str).strip().lower()

def number_to_words(date_str):
    try:
        dt = datetime.strptime(str(date_str).strip(), '%d-%m-%Y')
        day = dt.day
        year = dt.year
        month_names = ["January", "February", "March", "April", "May", "June", 
                       "July", "August", "September", "October", "November", "December"]
        month = month_names[dt.month - 1]
        
        def num_to_w(n):
            d1 = ["", "First", "Second", "Third", "Fourth", "Fifth", "Sixth", "Seventh", "Eighth", "Ninth", "Tenth", 
                  "Eleventh", "Twelfth", "Thirteenth", "Fourteenth", "Fifteenth", "Sixteenth", "Seventeenth", "Eighteenth", "Nineteenth"]
            d2 = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
            if n < 20:
                return d1[n]
            else:
                return d2[n // 10] + (" " + d1[n % 10] if n % 10 != 0 else "")

        day_word = num_to_w(day)
        y_first = year // 100
        y_second = year % 100
        
        y_words = []
        if y_first == 20:
            y_words.append("Two Thousand")
            if y_second > 0:
                y_words.append(num_to_w(y_second).lower())
        else:
            y_words.append(str(year))
            
        return f"{day_word} {month} {' '.join(y_words)}"
    except Exception:
        return str(date_str)

def get_row_data(ref_number):
    excel_file = "Online_Registration.xlsx"
    if not os.path.exists(excel_file):
        excel_file = "Online_Registration (2).xls"
        if not os.path.exists(excel_file):
            return None
    try:
        df = pd.read_excel(excel_file, header=None, engine='openpyxl')
    except Exception:
        try:
            df = pd.read_excel(excel_file, header=None, engine='xlrd')
        except Exception:
            return None
            
    reg_idx = col_to_idx(COL_REG_NO)
    for idx, row in df.iloc[1:].iterrows():
        if len(row) > reg_idx and str(row.iloc[reg_idx]).strip() == str(ref_number).strip():
            return row
    return None

def draw_fitted_string(c, x, y, text, max_width, font_name="Helvetica-Bold", start_size=10, min_size=6):
    size = start_size
    while size > min_size and c.stringWidth(text, font_name, size) > max_width:
        size -= 0.5
    c.setFont(font_name, size)
    c.drawString(x, y, text)

def generate_admission_pdf(ref_number):
    matched_row = get_row_data(ref_number)
    if matched_row is None:
        return None

    def val(col_letter):
        idx = col_to_idx(col_letter)
        if idx < 0 or matched_row is None:
            return ''
        v = matched_row.iloc[idx] if idx < len(matched_row) else ''
        if pd.isna(v):
            return ''
        if isinstance(v, datetime) or hasattr(v, 'strftime'):
            return v.strftime('%d-%m-%Y')
        
        val_str = str(v).strip()
        if col_letter == COL_CLASS:
            if val_str.lower() in ['nur', 'nursery']:
                return 'NURSERY'
            return val_str.upper()
        
        # Keep Caste blank if it accidentally contains 'FEMALE' or 'MALE'
        if col_letter == COL_CASTE:
            if val_str.upper() in ['FEMALE', 'MALE']:
                return ''
                
        return val_str

    output_dir = os.path.abspath("Generated_Admission_Forms")
    os.makedirs(output_dir, exist_ok=True)
    pdf_path = os.path.join(output_dir, f"Admission_Form_{str(ref_number).replace('/', '_')}.pdf")
    
    c = canvas.Canvas(pdf_path, pagesize=legal)
    width, height = legal
    
    left_margin = 42  
    right_margin = width - 30  
    top_offset = height - 65   

    logo_path = "school_logo.png"
    if os.path.exists(logo_path):
        try:
            c.drawImage(logo_path, left_margin, top_offset - 55, width=65, height=65, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    try:
        c.setFont("Algerian", 21)
    except Exception:
        c.setFont("Helvetica-Bold", 21)
        
    c.drawCentredString(width / 2.0, top_offset + 2, "HOLY CHILD SCHOOL, GHAZIABAD")
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width / 2.0, top_offset - 18, "ADMISSION FORM")

    c.setFont("Helvetica", 10.5)
    c.drawString(430, top_offset - 40, "Ref. No. :")
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(490, top_offset - 40, f" {str(ref_number)}")
    c.setDash(1, 2)
    c.line(480, top_offset - 43, right_margin, top_offset - 43)

    c.setFont("Helvetica", 10.5)
    c.drawString(430, top_offset - 60, "PEN #")
    c.line(470, top_offset - 63, right_margin, top_offset - 63)

    c.setDash()
    c.rect(left_margin, top_offset - 120, right_margin - left_margin, 45)
    
    c.setFont("Helvetica", 10.5)
    c.drawString(left_margin + 10, top_offset - 92, "Admission to Class :")
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(left_margin + 110, top_offset - 92, f" {val(COL_CLASS)}")

    c.setFont("Helvetica", 10.5)
    c.drawString(380, top_offset - 92, "Year :")
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(420, top_offset - 92, " 2027-28")
    
    c.setFont("Helvetica", 10.5)
    c.drawString(left_margin + 10, top_offset - 112, "Date of Admission :")
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(left_margin + 110, top_offset - 112, f" {datetime.now().strftime('%d-%m-%Y')}")
    
    c.setFont("Helvetica", 10.5)
    c.drawString(380, top_offset - 112, "Admission No :")

    y = top_offset - 147
    line_gap = 19

    # 1. Name of Pupil
    c.setFont("Helvetica", 10)
    c.drawString(left_margin, y, "1")
    c.drawString(left_margin + 15, y, "a) Name of Pupil :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 97, y, f" {val(COL_NAME)}")
    c.setDash(1, 2)
    c.line(left_margin + 94, y - 3, right_margin, y - 3)
    
    y -= line_gap
    dob_val = val(COL_DOB)
    dob_words = number_to_words(dob_val)
    
    # b) Date of Birth
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, "b) Date of Birth :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 97, y, f" {dob_val}")
    c.setDash(1, 2)
    c.line(left_margin + 94, y - 3, 245, y - 3)  
    
    c.setFont("Helvetica", 10)
    c.drawString(252, y, "(in words)")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(312, y, f" {dob_words}")
    c.setDash(1, 2)
    c.line(309, y - 3, right_margin, y - 3)

    y -= 15
    c.setFont("Helvetica-Oblique", 9.5)
    c.drawString(left_margin + 75, y, "N.B. i. For first admission a birth certificate from the Municipal Corporation must be attached.")
    y -= 13
    c.drawString(left_margin + 87, y, "ii. The \"Declaration on Statement of Age\" should also be signed.")

    y -= line_gap
    # c) School last attended
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, "c) Name of the school last attended")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 177, y, f" {val(COL_SCHOOL)}")
    c.setDash(1, 2)
    c.line(left_margin + 174, y - 3, right_margin, y - 3)
    
    y -= line_gap
    # d) Nationality & e) State
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, "d) Nationality :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 87, y, " INDIAN")
    c.setDash(1, 2)
    c.line(left_margin + 84, y - 3, 310, y - 3)
    
    c.setFont("Helvetica", 10)
    c.drawString(330, y, "e) State :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(372, y, f" {val(COL_STATE)}")
    c.line(369, y - 3, right_margin, y - 3)

    y -= line_gap
    # f) Home town & g) Mother Tongue
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, "f) Home town :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 92, y, f" {val(COL_HOMETOWN)}")
    c.setDash(1, 2)
    c.line(left_margin + 89, y - 3, 310, y - 3)
    
    c.setFont("Helvetica", 10)
    c.drawString(330, y, "g) Mother Tongue :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(422, y, f" {val(COL_MOTHER_TONGUE)}")
    c.line(419, y - 3, right_margin, y - 3)

    y -= line_gap
    # h) Religion & i) Caste
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, "h) Religion :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 82, y, f" {val(COL_RELIGION)}")
    c.setDash(1, 2)
    c.line(left_margin + 79, y - 3, 310, y - 3)
    
    c.setFont("Helvetica", 10)
    c.drawString(330, y, "i) Caste :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(372, y, f" {val(COL_CASTE)}")
    c.line(369, y - 3, right_margin, y - 3)

    y -= line_gap
    # j) Apar No. & k) Social Category
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, "j) Apar No. :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 82, y, "")  
    c.setDash(1, 2)
    c.line(left_margin + 79, y - 3, 310, y - 3)
    
    c.setFont("Helvetica", 10)
    c.drawString(330, y, "k) Social Category :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(417, y, f" {val(COL_CATEGORY)}")
    c.setDash(1, 2)
    c.line(414, y - 3, 475, y - 3)
    c.setFont("Helvetica", 9.5)
    c.drawString(480, y, "(SC/ST/OBC/GEN)")

    y -= line_gap
    # l) Aadhar No & m) Blood Group
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, "l) Aadhar No :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 87, y, f" {val(COL_AADHAR)}")
    c.setDash(1, 2)
    c.line(left_margin + 84, y - 3, 310, y - 3)
    
    c.setFont("Helvetica", 10)
    c.drawString(330, y, "m) Blood Group :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(417, y, f" {val(COL_BLOOD)}")
    c.line(414, y - 3, right_margin, y - 3)

    y -= line_gap
    # n) Contact No & o) E-mail
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, "n) Contact No :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 87, y, f" {val(COL_CONTACT)}")
    c.setDash(1, 2)
    c.line(left_margin + 84, y - 3, 310, y - 3)
    
    c.setFont("Helvetica", 10)
    c.drawString(330, y, "o) E-mail :")
    draw_fitted_string(c, 377, y, f" {format_email(val(COL_EMAIL))}", right_margin - 377, font_name="Helvetica-Bold", start_size=10, min_size=7)
    c.line(374, y - 3, right_margin, y - 3)

    # ==========================================
    # 2. FATHER & MOTHER DETAILS SECTION
    # ==========================================
    y -= 25
    c.setFont("Helvetica", 10)
    c.drawString(left_margin, y, "2")
    c.drawString(left_margin + 15, y, "a) Father's Name :")
    
    f_name = val(COL_FATHER_NAME)
    f_qual = val(COL_FATHER_QUAL)
    
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 112, y, f" {f_name}")
    c.setDash(1, 2)
    c.line(left_margin + 109, y - 3, 310, y - 3)
    
    c.setFont("Helvetica", 10)
    c.drawString(318, y, "& Edu. Qual. :")
    draw_fitted_string(c, 397, y, f" {f_qual}", right_margin - 397, font_name="Helvetica-Bold", start_size=10, min_size=6)
    c.line(394, y - 3, right_margin, y - 3)
    
    y -= line_gap
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 25, y, "Contact No :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 97, y, f" {val(COL_FATHER_CONT)}")
    c.setDash(1, 2)
    c.line(left_margin + 94, y - 3, 310, y - 3)

    c.setFont("Helvetica", 10)
    c.drawString(320, y, "Aadhaar No :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(392, y, f" {val(COL_FATHER_AADHAAR)}")  
    c.line(389, y - 3, right_margin, y - 3)

    y -= line_gap
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 25, y, "Occupation / Designation :")
    
    f_occ = val(COL_FATHER_OCC)
    f_occ_extra = val(COL_FATHER_OCC_EXTRA)
    f_occ_display = f"{f_occ} ({f_occ_extra})" if f_occ_extra else f_occ
    
    draw_fitted_string(c, left_margin + 152, y, f" {f_occ_display}", right_margin - (left_margin + 152), font_name="Helvetica-Bold", start_size=10, min_size=6)
    c.setDash(1, 2)
    c.line(left_margin + 149, y - 3, right_margin, y - 3)

    y -= line_gap
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 25, y, "Office Address :")
    draw_fitted_string(c, left_margin + 102, y, f" {clean_office_address(val(COL_FATHER_OFFICE))}", right_margin - (left_margin + 102), font_name="Helvetica-Bold", start_size=10, min_size=7)
    c.setDash(1, 2)
    c.line(left_margin + 99, y - 3, right_margin, y - 3)

    y -= 22
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, "b) Mother's Name :")
    
    m_name = val(COL_MOTHER_NAME)
    m_qual = val(COL_MOTHER_QUAL)
    
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 112, y, f" {m_name}")
    c.setDash(1, 2)
    c.line(left_margin + 109, y - 3, 310, y - 3)

    c.setFont("Helvetica", 10)
    c.drawString(318, y, "& Edu. Qual. :")
    draw_fitted_string(c, 397, y, f" {m_qual}", right_margin - 397, font_name="Helvetica-Bold", start_size=10, min_size=6)  
    c.line(394, y - 3, right_margin, y - 3)

    y -= line_gap
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 25, y, "Contact No :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 97, y, f" {val(COL_MOTHER_CONT)}")
    c.setDash(1, 2)
    c.line(left_margin + 94, y - 3, 310, y - 3)

    c.setFont("Helvetica", 10)
    c.drawString(320, y, "Aadhaar No :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(392, y, f" {val(COL_MOTHER_AADHAR)}")  
    c.line(389, y - 3, right_margin, y - 3)

    y -= line_gap
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 25, y, "Occupation / Designation :")
    draw_fitted_string(c, left_margin + 152, y, f" {val(COL_MOTHER_OCC)}", right_margin - (left_margin + 152), font_name="Helvetica-Bold", start_size=10, min_size=6)
    c.setDash(1, 2)
    c.line(left_margin + 149, y - 3, right_margin, y - 3)

    y -= line_gap
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 25, y, "Office Address :")
    draw_fitted_string(c, left_margin + 102, y, f" {clean_office_address(val(COL_MOTHER_OFFICE))}", right_margin - (left_margin + 102), font_name="Helvetica-Bold", start_size=10, min_size=7)
    c.setDash(1, 2)
    c.line(left_margin + 99, y - 3, right_margin, y - 3)

    # ==========================================
    # 3. ADDRESS & GUARDIAN SECTION
    # ==========================================
    y -= 22
    c.setFont("Helvetica", 10)
    c.drawString(left_margin, y, "3")
    c.drawString(left_margin + 15, y, "a) Permanent Address :")
    
    perm_clean, perm_pin = extract_address_and_pin(val(COL_PERM_ADDR))
    draw_fitted_string(c, left_margin + 122, y, f" {perm_clean}", right_margin - (left_margin + 122), font_name="Helvetica-Bold", start_size=10, min_size=7)
    c.setDash(1, 2)
    c.line(left_margin + 119, y - 3, right_margin, y - 3)
    
    y -= 16
    c.setFont("Helvetica", 10)
    c.drawString(440, y, "Pin No :")
    if perm_pin:
        c.setFont("Helvetica-Bold", 10)
        c.drawString(487, y, f" {perm_pin}")
    c.setDash(1, 2)
    c.line(484, y - 3, right_margin, y - 3)

    y -= line_gap
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, "b) Present Address :")
    
    pres_clean, pres_pin = extract_address_and_pin(val(COL_PRES_ADDR))
    draw_fitted_string(c, left_margin + 112, y, f" {pres_clean}", right_margin - (left_margin + 112), font_name="Helvetica-Bold", start_size=10, min_size=7)
    c.setDash(1, 2)
    c.line(left_margin + 109, y - 3, right_margin, y - 3)
    
    y -= 16
    c.setFont("Helvetica", 10)
    c.drawString(440, y, "Pin No :")
    if pres_pin:
        c.setFont("Helvetica-Bold", 10)
        c.drawString(487, y, f" {pres_pin}")
    c.setDash(1, 2)
    c.line(484, y - 3, right_margin, y - 3)

    y -= line_gap
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, "c) Local Guardian :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 102, y, f" {val(COL_GUARDIAN)}")
    c.setDash(1, 2)
    c.line(left_margin + 99, y - 3, right_margin, y - 3)

    y -= line_gap
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 25, y, "Relationship with pupil :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 152, y, f" {val(COL_GUARDIAN_REL)}")
    c.setDash(1, 2)
    c.line(left_margin + 149, y - 3, 350, y - 3)

    c.setFont("Helvetica", 10)
    c.drawString(360, y, "Ph. No. :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(412, y, f" {val(COL_GUARDIAN_CONT)}")
    c.setDash(1, 2)
    c.line(409, y - 3, right_margin, y - 3)

    y -= line_gap
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 25, y, "Siblings Details :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 107, y, f" {val(COL_SIB_NAME)}")
    c.setDash(1, 2)
    c.line(left_margin + 104, y - 3, 330, y - 3)
    
    c.setFont("Helvetica", 10)
    c.drawString(340, y, "Adm. No.")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(387, y, f" {val(COL_SIB_ADM)}")
    c.setDash(1, 2)
    c.line(384, y - 3, 450, y - 3)
    
    c.setFont("Helvetica", 10)
    c.drawString(460, y, "Class & Div")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(517, y, f" {val(COL_SIB_CLASS)}")
    c.setDash(1, 2)
    c.line(514, y - 3, right_margin, y - 3)

    y -= 12
    c.setDash()
    c.setLineWidth(0.75)
    c.line(left_margin, y, right_margin, y)

    y -= 5
    styles = getSampleStyleSheet()
    decl_style = ParagraphStyle('DeclarationStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=10.5, leading=15, alignment=4)
    p1_text = "I have read the rules and regulations of Holy Child School, Ghaziabad and wish my daughter/ward to be brought up accordingly. I clearly understand that my daughter/ward is liable to dismissal for misconduct, habitual idleness, non-payment of fees when due, or for any other reason which in the opinion of school authorities renders her continuance in school undesirable."
    p1 = Paragraph(p1_text, decl_style)
    t1 = Table([[p1]], colWidths=[right_margin - left_margin])
    t1.setStyle(TableStyle([('TOPPADDING', (0,0), (-1,-1), 0), ('BOTTOMPADDING', (0,0), (-1,-1), 0), ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0)]))
    
    t1_w, t1_h = t1.wrap(right_margin - left_margin, 100)
    t1.drawOn(c, left_margin, y - t1_h)

    y = y - t1_h - 25
    c.setDash(1, 2)
    c.setLineWidth(1)
    p1_sig_x1 = 380
    p1_sig_x2 = right_margin - 20
    c.line(p1_sig_x1, y, p1_sig_x2, y)
    c.setFont("Helvetica", 10)
    # Slash positioned right on top of the line
    c.drawCentredString((p1_sig_x1 + p1_sig_x2) / 2.0, y, "/")
    
    y -= 12
    c.setFont("Helvetica", 10)
    c.drawCentredString(380 + (right_margin - 20 - 380) / 2.0, y, "Signature of Parent")

    y -= 12
    c.setDash()
    c.setLineWidth(0.75)
    c.line(left_margin, y, right_margin, y)

    y -= 16
    c.setFont("Helvetica-Bold", 10.5)
    c.drawCentredString(width / 2.0, y, "DECLARATION ON STATEMENT OF AGE AND NAME")

    y -= 15
    c.setFont("Helvetica", 10)
    c.drawString(left_margin, y, "I hereby declare that the statement of the date of birth of my ward.")

    y -= 20
    student_name = val(COL_NAME)
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(265, y + 2, f" {student_name}")
    c.setDash(1, 2)
    c.line(left_margin, y, 420, y)
    
    c.setFont("Helvetica", 9.5)
    c.drawString(425, y, "(Name of the child in Block Letters)")

    y -= 16
    c.setFont("Helvetica", 10)
    c.drawString(left_margin, y, "Date of Birth :")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 77, y, f" {dob_val}")
    c.setDash(1, 2)
    c.line(left_margin + 74, y - 3, 270, y - 3)
    c.setFont("Helvetica", 10)
    c.drawString(285, y, "(in words)")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(337, y, f" {dob_words}")
    c.line(334, y - 3, right_margin, y - 3)

    y -= 8
    p2_text = "is entirely correct and I fully understand that the school authorities will at no time entertain any application from me or from any one else connected with my ward for a change in this statement under any circumstances."
    p2 = Paragraph(p2_text, decl_style)
    t2 = Table([[p2]], colWidths=[right_margin - left_margin])
    t2.setStyle(TableStyle([('TOPPADDING', (0,0), (-1,-1), 0), ('BOTTOMPADDING', (0,0), (-1,-1), 0), ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0)]))
    
    t2_w, t2_h = t2.wrap(right_margin - left_margin, 100)
    t2.drawOn(c, left_margin, y - t2_h)

    y = y - t2_h - 18

    c.setFont("Helvetica", 10)
    c.drawString(left_margin, y, "Name of parents :")
    
    f_raw = val(COL_FATHER_NAME).replace("MR.", "").replace("Mr.", "").replace("MR ", "").strip()
    m_raw = val(COL_MOTHER_NAME).replace("MRS.", "").replace("Mrs.", "").replace("MS.", "").replace("Ms.", "").strip()
    parents_str = f"MR. {f_raw} / MRS. {m_raw}"
    
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 92, y, f" {parents_str}")
    c.setDash(1, 2)
    c.line(left_margin + 89, y - 3, 430, y - 3)
    
    y -= 35
    c.setDash(1, 2)
    c.setLineWidth(1)
    
    # Bottom Parent Signature Line & Slash right on top of the line
    p_sig_x1 = left_margin
    p_sig_x2 = left_margin + 180
    c.line(p_sig_x1, y, p_sig_x2, y)
    c.setFont("Helvetica", 10)
    c.drawCentredString((p_sig_x1 + p_sig_x2) / 2.0, y, "/")

    # Bottom Principal Signature Line & Slash right on top of the line
    pr_sig_x1 = 380
    pr_sig_x2 = right_margin - 20
    c.line(pr_sig_x1, y, pr_sig_x2, y)
    c.drawCentredString((pr_sig_x1 + pr_sig_x2) / 2.0, y, "/")
    
    y -= 12
    c.setFont("Helvetica", 10)
    c.drawCentredString(left_margin + 90, y, "Signature of Parent")
    c.drawCentredString(380 + (right_margin - 20 - 380) / 2.0, y, "Signature of Principal")

    c.setDash()
    c.save()
    return pdf_path

def draw_single_receipt(c, x, y_bottom, width, height, matched_row, ref_number):
    if matched_row is None:
        return
    def val(col_letter):
        idx = col_to_idx(col_letter)
        if idx < 0:
            return ''
        v = matched_row.iloc[idx] if idx < len(matched_row) else ''
        if pd.isna(v):
            return ''
        val_str = str(v).strip()
        if col_letter == COL_CLASS:
            if val_str.lower() in ['nur', 'nursery']:
                return 'NURSERY'
            return val_str.upper()
        return val_str

    c.setLineWidth(1)
    c.rect(x, y_bottom, width, height)

    logo_path = "school_logo.png"

    if os.path.exists(logo_path):
        try:
            c.saveState()
            c.setFillAlpha(0.12)
            watermark_size = 150
            c.drawImage(logo_path, x + (width - watermark_size)/2.0, y_bottom + (height - watermark_size)/2.0, width=watermark_size, height=watermark_size, preserveAspectRatio=True, mask='auto')
            c.restoreState()
        except Exception:
            pass

    top_margin_pts = 71
    header_y = y_bottom + height - top_margin_pts

    if os.path.exists(logo_path):
        try:
            logo_w = 42
            logo_h = 42
            c.drawImage(logo_path, x + (width - logo_w)/2.0, header_y + 16, width=logo_w, height=logo_h, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    c.setFont("Helvetica-Bold", 10.5)
    c.drawCentredString(x + width/2.0, header_y, "HOLY CHILD SCHOOL")
    c.setFont("Helvetica", 7.5)
    c.drawCentredString(x + width/2.0, header_y - 11, "III B NEHRU NAGAR, GHAZIABAD (U.P.)-201001")
    
    c.setFont("Helvetica-Bold", 9.5)
    c.drawCentredString(x + width/2.0, header_y - 26, "Acknowledgement Receipt")

    table_left_margin = 14
    table_right_margin = 14
    table_width = width - (table_left_margin + table_right_margin)

    c.setFont("Helvetica", 8)
    c.drawString(x + table_left_margin, header_y - 42, "Reg. No. :")
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(x + table_left_margin + 45, header_y - 42, f" {str(ref_number)}")

    date_str = datetime.now().strftime('%d-%b-%Y')
    c.setFont("Helvetica", 8)
    c.drawString(x + width - table_right_margin - 105, header_y - 42, "Date:")
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(x + width - table_right_margin - 78, header_y - 42, f" {date_str}")

    cell_style = ParagraphStyle('ReceiptCell', fontName='Helvetica', fontSize=7.2, leading=9.5)
    cell_bold = ParagraphStyle('ReceiptCellBold', fontName='Helvetica-Bold', fontSize=7.2, leading=9.5)

    class_full = val(COL_CLASS)
    addr_val = format_address(val(COL_PRES_ADDR) if val(COL_PRES_ADDR) else val(COL_PERM_ADDR))

    t_data = [
        [Paragraph("Name:", cell_bold), Paragraph(f" {val(COL_NAME)}", cell_style)],
        [Paragraph("Father's Name:", cell_bold), Paragraph(f" MR. {val(COL_FATHER_NAME)}", cell_style)],
        [Paragraph("Admission No:", cell_bold), Paragraph("", cell_style)],
        [Paragraph("Class & Div:", cell_bold), Paragraph(f" {class_full}", cell_style)],
        [Paragraph("Address:", cell_bold), Paragraph(f" {addr_val}", cell_style)]
    ]
    
    t = Table(t_data, colWidths=[75, table_width - 75])
    t.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#333333')), ('TOPPADDING', (0,0), (-1,-1), 2.5), ('BOTTOMPADDING', (0,0), (-1,-1), 2.5), ('LEFTPADDING', (0,0), (-1,-1), 3), ('RIGHTPADDING', (0,0), (-1,-1), 3), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    
    t.wrapOn(c, table_width, height)
    t.drawOn(c, x + table_left_margin, header_y - 148)

    p_style = ParagraphStyle('ConfStyle', fontName='Helvetica-Oblique', fontSize=7.8, leading=10.5)
    conf_text = "This is to confirm your Admission with Holy Child School as per above details and Received Rs: 1,000/- (One Thousand only) cash on account of Registration for new admission."
    p = Paragraph(conf_text, p_style)
    
    t_conf = Table([[p]], colWidths=[table_width])
    t_conf.setStyle(TableStyle([('TOPPADDING', (0,0), (-1,-1), 1), ('BOTTOMPADDING', (0,0), (-1,-1), 1), ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0)]))
    t_conf.wrapOn(c, table_width, 40)
    t_conf.drawOn(c, x + table_left_margin, header_y - 195)

    bottom_margin_pts = 71
    c.setDash(1, 1)
    line_x1 = x + width - 115
    line_x2 = x + width - 12
    line_y = y_bottom + bottom_margin_pts
    c.line(line_x1, line_y, line_x2, line_y)
    c.setDash()
    c.setFont("Helvetica-Bold", 7.5)
    
    sig_x = x + width - 98
    sig_y = y_bottom + bottom_margin_pts - 11
    c.drawString(sig_x, sig_y, "Authorised Signatory")

def generate_4_receipts_pdf(ref_list):
    output_dir = os.path.abspath("Generated_Admission_Forms")
    os.makedirs(output_dir, exist_ok=True)
    pdf_path = os.path.join(output_dir, "Acknowledgement_Receipts_4in1.pdf")
    
    c = canvas.Canvas(pdf_path, pagesize=A4)
    a4_w, a4_h = A4
    
    margin_x = 18
    margin_y = 18
    box_w = (a4_w - (margin_x * 3)) / 2.0
    box_h = (a4_h - (margin_y * 3)) / 2.0
    
    positions = [
        (margin_x, a4_h - margin_y - box_h),
        (margin_x * 2 + box_w, a4_h - margin_y - box_h),
        (margin_x, margin_y),
        (margin_x * 2 + box_w, margin_y)
    ]
    
    for i, ref in enumerate(ref_list):
        if i >= 4:
            break
        pos_x, pos_y = positions[i]
        matched_row = get_row_data(ref)
        if matched_row is not None:
            draw_single_receipt(c, pos_x, pos_y, box_w, box_h, matched_row, ref)
            
    c.setDash(2, 2)
    c.setLineWidth(0.5)
    c.setStrokeColor(colors.HexColor("#888888"))
    c.line(a4_w / 2.0, margin_y, a4_w / 2.0, a4_h - margin_y)
    c.line(margin_x, a4_h / 2.0, a4_w - margin_x, a4_h / 2.0)
    c.setDash()

    c.save()
    return pdf_path

st.sidebar.title("Holy Child School Hub")
menu = st.sidebar.radio("Navigation", ["Admission Form Generator", "Acknowledgement Receipts (4-in-1)"])

excel_file = "Online_Registration.xlsx"
if not os.path.exists(excel_file):
    excel_file = "Online_Registration (2).xls"

if menu == "Admission Form Generator":
    df_excel = None
    if os.path.exists(excel_file):
        try:
            df_excel = pd.read_excel(excel_file, header=None, engine='openpyxl')
        except Exception as e:
            st.error(f"⚠️ Excel file read karne mein error aayi: {e}")

    if df_excel is not None:
        reg_idx = col_to_idx(COL_REG_NO)
        ref_list = df_excel.iloc[1:, reg_idx].dropna().astype(str).tolist()
        
        st.markdown("""<h2 style='text-align: center; color: #1e3d59;'>Holy Child School — Admission Form Generator</h2>""", unsafe_allow_html=True)
        st.write("---")
        
        col1, col2 = st.columns([3, 1])
        with col1:
            selected_ref = st.selectbox("Select Registration Number (Column C):", ref_list)
        with col2:
            st.write("")
            if st.button("Generate Admission Form PDF", type="primary"):
                pdf_file = generate_admission_pdf(selected_ref)
                if pdf_file and os.path.exists(pdf_file):
                    st.success("✅ PDF Successfully Generated!")
                    with open(pdf_file, "rb") as f:
                        st.download_button("📥 Download Admission Form", f, file_name=os.path.basename(pdf_file), mime="application/pdf")
                else:
                    st.error("⚠️ PDF generate karne mein koi problem aayi.")

elif menu == "Acknowledgement Receipts (4-in-1)":
    df_excel = None
    if os.path.exists(excel_file):
        try:
            df_excel = pd.read_excel(excel_file, header=None, engine='openpyxl')
        except Exception as e:
            st.error(f"⚠️ Excel file read karne mein error aayi: {e}")

    if df_excel is not None:
        reg_idx = col_to_idx(COL_REG_NO)
        ref_list = df_excel.iloc[1:, reg_idx].dropna().astype(str).tolist()
        
        st.markdown("""<h2 style='text-align: center; color: #1e3d59;'>Holy Child School — Acknowledgement Receipts (4-in-1)</h2>""", unsafe_allow_html=True)
        st.write("---")
        
        selected_refs = st.multiselect("Select up to 4 Registration Numbers (Column C):", ref_list, max_selections=4)
        
        if st.button("Generate 4-in-1 Receipts PDF", type="primary"):
            if selected_refs:
                pdf_file = generate_4_receipts_pdf(selected_refs)
                if pdf_file and os.path.exists(pdf_file):
                    st.success("✅ Receipts PDF Successfully Generated!")
                    with open(pdf_file, "rb") as f:
                        st.download_button("📥 Download Receipts PDF", f, file_name=os.path.basename(pdf_file), mime="application/pdf")
                else:
                    st.error("⚠️ Receipts generate karne mein koi problem aayi.")
            else:
                st.warning("⚠️ Kripya kam se kam ek Registration Number select karein.")
