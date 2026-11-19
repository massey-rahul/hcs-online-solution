import os
import re
import base64
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

# Bulletproof Font Registration (Works on both Windows Desktop & Streamlit Cloud)
try:
    font_registered = False
    possible_paths = [
        "alger.ttf", 
        "ALGER.TTF", 
        r"C:\Windows\Fonts\alger.ttf", 
        r"C:\Windows\Fonts\ALGER.TTF"
    ]
    for path in possible_paths:
        if os.path.exists(path):
            pdfmetrics.registerFont(TTFont('Algerian', path))
            font_registered = True
            break
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

def clean_father_name(name_str):
    if not name_str:
        return ""
    cleaned = re.sub(r'^(?:MR\.?|SHRI|SH\.?)\s+', '', str(name_str).strip(), flags=re.IGNORECASE).strip()
    return f"MR. {cleaned}" if cleaned else ""

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
        
        def num_to_w(n, is_ordinal=False):
            d1 = ["", "First", "Second", "Third", "Fourth", "Fifth", "Sixth", "Seventh", "Eighth", "Ninth", "Tenth", 
                  "Eleventh", "Twelfth", "Thirteenth", "Fourteenth", "Fifteenth", "Sixteenth", "Seventeenth", "Eighteenth", "Nineteenth"]
            d2 = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
            ord_d1 = ["", "First", "Second", "Third", "Fourth", "Fifth", "Sixth", "Seventh", "Eighth", "Ninth", "Tenth", 
                      "Eleventh", "Twelfth", "Thirteenth", "Fourteenth", "Fifteenth", "Sixteenth", "Seventeenth", "Eighteenth", "Nineteenth"]
            ord_d2 = ["", "", "Twentieth", "Thirtieth", "Fortieth", "Fiftieth", "Sixtieth", "Seventieth", "Eightieth", "Ninetieth"]
            
            if is_ordinal:
                if n < 20:
                    return ord_d1[n]
                else:
                    tens = n // 10
                    rem = n % 10
                    if rem == 0:
                        return ord_d2[tens]
                    else:
                        return d2[tens] + " " + ord_d1[rem]
            else:
                if n < 20:
                    return d1[n]
                else:
                    return d2[n // 10] + (" " + d1[n % 10] if n % 10 != 0 else "")

        day_word = num_to_w(day, is_ordinal=True)
        y_first = year // 100
        y_second = year % 100
        
        y_words = []
        if y_first == 20:
            y_words.append("Two Thousand")
            if y_second > 0:
                y_words.append(num_to_w(y_second))
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

def get_admission_details(ref_number):
    """Strictly reads ONLY Admission_Details.xlsx and matches exact columns"""
    details_file = "Admission_Details.xlsx"
    
    if not os.path.exists(details_file):
        print(f"❌ Error: {details_file} does not exist in folder!")
        return "", ""
            
    try:
        df = pd.read_excel(details_file, header=None)
        target_reg = str(ref_number).strip().upper()
        
        for idx, row in df.iterrows():
            if len(row) > 4 and pd.notna(row.iloc[4]):
                excel_reg = str(row.iloc[4]).strip().upper()
                
                if excel_reg == target_reg:
                    class_sec = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ""
                    adm_no = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else ""
                    return adm_no, class_sec
    except Exception as e:
        print(f"❌ Exception while reading file: {e}")
        
    return "", ""

def draw_fitted_string(c, x, y, text, max_width, font_name="Helvetica-Bold", start_size=10, min_size=6):
    size = start_size
    while size > min_size and c.stringWidth(text, font_name, size) > max_width:
        size -= 0.5
    c.setFont(font_name, size)
    c.drawString(x, y, text)

def draw_dotted_after_text(c, text, x_start, font_name, font_size, y_line, x_end, space_gap=4):
    """Starts a dotted line after a uniform gap from the end of `text`."""
    if x_end <= x_start:
        return
    line_start = c.stringWidth(text, font_name, font_size) + x_start + space_gap
    if line_start < x_end:
        c.saveState()
        c.setDash(1, 2)
        c.line(line_start, y_line, x_end, y_line)
        c.restoreState()

def show_pdf_preview(pdf_path):
    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()
        base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
    
    is_cloud = "streamlit.app" in st.context.headers.get("Host", "") if hasattr(st, "context") else False
    
    if not is_cloud:
        pdf_display = f'''
        <div style="border: 2px solid #1e3d59; border-radius: 8px; padding: 5px; background-color: #ffffff; margin-top: 15px; margin-bottom: 15px;">
            <embed src="data:application/pdf;base64,{base64_pdf}#view=FitH" type="application/pdf" width="100%" height="800px" style="border: none;" />
        </div>
        '''
        st.markdown(pdf_display, unsafe_allow_html=True)
    else:
        st.info("☁️ Running on Cloud: Direct inline frames are restricted by browser security. Use the options below to view or download:")
        
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="📥 Download PDF",
                data=pdf_bytes,
                file_name="Document.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        with col2:
            st.link_button(
                label="🔍 Open / Preview PDF in New Tab",
                url=f"data:application/pdf;base64,{base64_pdf}",
                use_container_width=True
            )

def generate_admission_pdf(ref_number):
    matched_row = get_row_data(ref_number)
    if matched_row is None:
        return None

    mapped_adm_no, mapped_class_sec = get_admission_details(ref_number)

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
            if mapped_class_sec:
                return mapped_class_sec.upper()
            if val_str.lower() in ['nur', 'nursery']:
                return 'NURSERY'
            return val_str.upper()
        
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
    top_offset = height - 75   

    logo_path = "school_logo.png"
    if os.path.exists(logo_path):
        try:
            c.drawImage(logo_path, left_margin, top_offset - 55, width=65, height=65, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    try:
        c.setFont("Algerian", 24)
    except Exception:
        c.setFont("Helvetica-Bold", 24)
        
    c._code.append('2 Tr')  
    c.setLineWidth(0.5)     
    c.drawCentredString(width / 2.0, top_offset + 2, "HOLY CHILD SCHOOL, GHAZIABAD")
    c._code.append('0 Tr')  
    
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(width / 2.0, top_offset - 20, "ADMISSION FORM")

    c.setFont("Helvetica", 10.5)
    ref_label = "Ref. No. :"
    c.drawString(430, top_offset - 40, ref_label)
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(490, top_offset - 40, f" {str(ref_number)}")
    draw_dotted_after_text(c, ref_label, 430, "Helvetica", 10.5, top_offset - 43, right_margin)

    c.setFont("Helvetica", 10.5)
    pen_label = "PEN #:"
    c.drawString(430, top_offset - 60, pen_label)
    draw_dotted_after_text(c, pen_label, 430, "Helvetica", 10.5, top_offset - 63, right_margin)

    c.setDash()
    c.rect(left_margin, top_offset - 120, right_margin - left_margin, 45)
    
    # Class & Section box
    c.setFont("Helvetica", 10.5)
    c.drawString(left_margin + 10, top_offset - 92, "Class & Section :")
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(142, top_offset - 92, f" {val(COL_CLASS)}")

    c.setFont("Helvetica", 10.5)
    c.drawString(405, top_offset - 92, "Academic Year :")
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(485, top_offset - 92, " 2027-28")
    
    c.setFont("Helvetica", 10.5)
    c.drawString(left_margin + 10, top_offset - 112, "Date of Admission :")
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(142, top_offset - 112, f" {datetime.now().strftime('%d-%m-%Y')}")
    
    c.setFont("Helvetica", 10.5)
    c.drawString(405, top_offset - 112, "Admission No. :")
    if mapped_adm_no:
        c.setFont("Helvetica-Bold", 10.5)
        c.drawString(485, top_offset - 112, f" {mapped_adm_no}")

    y = top_offset - 147
    line_gap = 19

    # 1 a) Name of Pupil
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(left_margin, y, "1")
    lbl_pupil = "a) Name of Pupil :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, lbl_pupil)
    
    c.saveState()
    c.setFont("Helvetica-Bold", 11)
    c.setFillColor(colors.black)
    c.drawString(left_margin + 102, y, f" {val(COL_NAME)}")
    c.restoreState()
    draw_dotted_after_text(c, lbl_pupil, left_margin + 15, "Helvetica", 10, y - 3, right_margin)
    
    y -= line_gap
    dob_val = val(COL_DOB)
    dob_words = number_to_words(dob_val)
    
    # 1 b) Date of Birth
    lbl_dob = "b) Date of Birth :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, lbl_dob)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 97, y, f" {dob_val}")
    draw_dotted_after_text(c, lbl_dob, left_margin + 15, "Helvetica", 10, y - 3, 245)
    
    lbl_in_words = "(in words)"
    c.setFont("Helvetica", 10)
    c.drawString(252, y, lbl_in_words)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(312, y, f" {dob_words}")
    draw_dotted_after_text(c, lbl_in_words, 252, "Helvetica", 10, y - 3, right_margin)

    y -= 15
    c.setFont("Helvetica-Oblique", 9.5)
    c.drawString(left_margin + 75, y, "N.B. i. For first admission a birth certificate from the Municipal Corporation must be attached.")
    y -= 13
    c.drawString(left_margin + 87, y, "ii. The \"Declaration on Statement of Age\" should also be signed.")

    y -= line_gap
    # 1 c) Name of the school last attended
    lbl_school = "c) Name of the school last attended :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, lbl_school)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 182, y, f" {val(COL_SCHOOL)}")
    draw_dotted_after_text(c, lbl_school, left_margin + 15, "Helvetica", 10, y - 3, right_margin)
    
    y -= line_gap
    # 1 d) Nationality & e) State
    lbl_nat = "d) Nationality :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, lbl_nat)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 87, y, " INDIAN")
    draw_dotted_after_text(c, lbl_nat, left_margin + 15, "Helvetica", 10, y - 3, 310)
    
    lbl_state = "e) State :"
    c.setFont("Helvetica", 10)
    c.drawString(330, y, lbl_state)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(372, y, f" {val(COL_STATE)}")
    draw_dotted_after_text(c, lbl_state, 330, "Helvetica", 10, y - 3, right_margin)

    y -= line_gap
    # 1 f) Home town & g) Mother Tongue
    lbl_hometown = "f) Home town :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, lbl_hometown)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 87, y, f" {val(COL_HOMETOWN)}")
    draw_dotted_after_text(c, lbl_hometown, left_margin + 15, "Helvetica", 10, y - 3, 310)
    
    lbl_mt = "g) Mother Tongue :"
    c.setFont("Helvetica", 10)
    c.drawString(330, y, lbl_mt)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(422, y, f" {val(COL_MOTHER_TONGUE)}")
    draw_dotted_after_text(c, lbl_mt, 330, "Helvetica", 10, y - 3, right_margin)

    y -= line_gap
    # 1 h) Religion & i) Caste
    lbl_rel = "h) Religion :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, lbl_rel)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 77, y, f" {val(COL_RELIGION)}")
    draw_dotted_after_text(c, lbl_rel, left_margin + 15, "Helvetica", 10, y - 3, 310)
    
    lbl_caste = "i) Caste :"
    c.setFont("Helvetica", 10)
    c.drawString(330, y, lbl_caste)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(372, y, f" {val(COL_CASTE)}")
    draw_dotted_after_text(c, lbl_caste, 330, "Helvetica", 10, y - 3, right_margin)

    y -= line_gap
    # 1 j) Apaar No. & k) Social Category
    lbl_apaar = "j) Apaar No. :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, lbl_apaar)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 82, y, "")  
    draw_dotted_after_text(c, lbl_apaar, left_margin + 15, "Helvetica", 10, y - 3, 310)
    
    lbl_cat = "k) Social Category :"
    c.setFont("Helvetica", 10)
    c.drawString(330, y, lbl_cat)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(422, y, f" {val(COL_CATEGORY)}")
    lbl_cat_opt = "(SC/ST/OBC/GEN)"
    c.setFont("Helvetica", 9.5)
    opt_w = c.stringWidth(lbl_cat_opt, "Helvetica", 9.5)
    opt_x = right_margin - opt_w
    c.drawString(opt_x, y, lbl_cat_opt)
    draw_dotted_after_text(c, lbl_cat, 330, "Helvetica", 10, y - 3, opt_x - 6)

    y -= line_gap
    # 1 l) Aadhar No. & m) Blood Group
    lbl_aadhar = "l) Aadhar No. :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, lbl_aadhar)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 87, y, f" {val(COL_AADHAR)}")
    draw_dotted_after_text(c, lbl_aadhar, left_margin + 15, "Helvetica", 10, y - 3, 310)
    
    lbl_blood = "m) Blood Group :"
    c.setFont("Helvetica", 10)
    c.drawString(330, y, lbl_blood)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(412, y, f" {val(COL_BLOOD)}")
    draw_dotted_after_text(c, lbl_blood, 330, "Helvetica", 10, y - 3, right_margin)

    y -= line_gap
    # 1 n) Contact No. & o) E-mail
    lbl_cont = "n) Contact No. :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, lbl_cont)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 92, y, f" {val(COL_CONTACT)}")
    draw_dotted_after_text(c, lbl_cont, left_margin + 15, "Helvetica", 10, y - 3, 310)
    
    lbl_email = "o) E-mail :"
    c.setFont("Helvetica", 10)
    c.drawString(330, y, lbl_email)
    draw_fitted_string(c, 377, y, f" {format_email(val(COL_EMAIL))}", right_margin - 377, font_name="Helvetica-Bold", start_size=10, min_size=7)
    draw_dotted_after_text(c, lbl_email, 330, "Helvetica", 10, y - 3, right_margin)

    # ==========================================
    # 2. FATHER & MOTHER DETAILS SECTION
    # ==========================================
    y -= 25
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(left_margin, y, "2")
    lbl_fname = "a) Father's Name :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, lbl_fname)
    
    f_name = val(COL_FATHER_NAME)
    f_qual = val(COL_FATHER_QUAL)
    
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 107, y, f" {f_name}")
    draw_dotted_after_text(c, lbl_fname, left_margin + 15, "Helvetica", 10, y - 3, 310)
    
    lbl_fqual = "Edu. Qual. :"
    c.setFont("Helvetica", 10)
    c.drawString(318, y, lbl_fqual)
    draw_fitted_string(c, 372, y, f" {f_qual}", right_margin - 372, font_name="Helvetica-Bold", start_size=10, min_size=6)
    draw_dotted_after_text(c, lbl_fqual, 318, "Helvetica", 10, y - 3, right_margin)
    
    y -= line_gap
    lbl_fcont = "Contact No. :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 25, y, lbl_fcont)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 87, y, f" {val(COL_FATHER_CONT)}")
    draw_dotted_after_text(c, lbl_fcont, left_margin + 25, "Helvetica", 10, y - 3, 310)

    lbl_faadhaar = "Aadhaar No. :"
    c.setFont("Helvetica", 10)
    c.drawString(320, y, lbl_faadhaar)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(382, y, f" {val(COL_FATHER_AADHAAR)}")  
    draw_dotted_after_text(c, lbl_faadhaar, 320, "Helvetica", 10, y - 3, right_margin)

    y -= line_gap
    lbl_focc = "Occupation / Designation :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 25, y, lbl_focc)
    
    f_occ = val(COL_FATHER_OCC)
    f_occ_extra = val(COL_FATHER_OCC_EXTRA)
    f_occ_display = f"{f_occ} ({f_occ_extra})" if f_occ_extra else f_occ
    
    draw_fitted_string(c, left_margin + 147, y, f" {f_occ_display}", right_margin - (left_margin + 147), font_name="Helvetica-Bold", start_size=10, min_size=6)
    draw_dotted_after_text(c, lbl_focc, left_margin + 25, "Helvetica", 10, y - 3, right_margin)

    y -= line_gap
    lbl_foff = "Office Address :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 25, y, lbl_foff)
    draw_fitted_string(c, left_margin + 102, y, f" {clean_office_address(val(COL_FATHER_OFFICE))}", right_margin - (left_margin + 102), font_name="Helvetica-Bold", start_size=10, min_size=7)
    draw_dotted_after_text(c, lbl_foff, left_margin + 25, "Helvetica", 10, y - 3, right_margin)

    y -= 22
    lbl_mname = "b) Mother's Name :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, lbl_mname)
    
    m_name = val(COL_MOTHER_NAME)
    m_qual = val(COL_MOTHER_QUAL)
    
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 112, y, f" {m_name}")
    draw_dotted_after_text(c, lbl_mname, left_margin + 15, "Helvetica", 10, y - 3, 310)

    lbl_mqual = "Edu. Qual. :"
    c.setFont("Helvetica", 10)
    c.drawString(318, y, lbl_mqual)
    draw_fitted_string(c, 372, y, f" {m_qual}", right_margin - 372, font_name="Helvetica-Bold", start_size=10, min_size=6)  
    draw_dotted_after_text(c, lbl_mqual, 318, "Helvetica", 10, y - 3, right_margin)

    y -= line_gap
    lbl_mcont = "Contact No. :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 25, y, lbl_mcont)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 87, y, f" {val(COL_MOTHER_CONT)}")
    draw_dotted_after_text(c, lbl_mcont, left_margin + 25, "Helvetica", 10, y - 3, 310)

    lbl_maadhaar = "Aadhaar No. :"
    c.setFont("Helvetica", 10)
    c.drawString(320, y, lbl_maadhaar)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(382, y, f" {val(COL_MOTHER_AADHAR)}")  
    draw_dotted_after_text(c, lbl_maadhaar, 320, "Helvetica", 10, y - 3, right_margin)

    y -= line_gap
    lbl_mocc = "Occupation / Designation :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 25, y, lbl_mocc)
    draw_fitted_string(c, left_margin + 147, y, f" {val(COL_MOTHER_OCC)}", right_margin - (left_margin + 147), font_name="Helvetica-Bold", start_size=10, min_size=6)
    draw_dotted_after_text(c, lbl_mocc, left_margin + 25, "Helvetica", 10, y - 3, right_margin)

    y -= line_gap
    lbl_moff = "Office Address :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 25, y, lbl_moff)
    draw_fitted_string(c, left_margin + 102, y, f" {clean_office_address(val(COL_MOTHER_OFFICE))}", right_margin - (left_margin + 102), font_name="Helvetica-Bold", start_size=10, min_size=7)
    draw_dotted_after_text(c, lbl_moff, left_margin + 25, "Helvetica", 10, y - 3, right_margin)

    # ==========================================
    # 3. ADDRESS & GUARDIAN SECTION
    # ==========================================
    y -= 22
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(left_margin, y, "3")
    lbl_perm = "a) Permanent Address :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, lbl_perm)
    
    perm_clean, perm_pin = extract_address_and_pin(val(COL_PERM_ADDR))
    draw_fitted_string(c, left_margin + 122, y, f" {perm_clean}", right_margin - (left_margin + 122), font_name="Helvetica-Bold", start_size=10, min_size=7)
    draw_dotted_after_text(c, lbl_perm, left_margin + 15, "Helvetica", 10, y - 3, right_margin)
    
    y -= 16
    lbl_ppin = "Pin No. :"
    c.setFont("Helvetica", 10)
    c.drawString(440, y, lbl_ppin)
    if perm_pin:
        c.setFont("Helvetica-Bold", 10)
        c.drawString(482, y, f" {perm_pin}")
    draw_dotted_after_text(c, lbl_ppin, 440, "Helvetica", 10, y - 3, right_margin)

    y -= line_gap
    lbl_pres = "b) Present Address :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, lbl_pres)
    
    pres_clean, pres_pin = extract_address_and_pin(val(COL_PRES_ADDR))
    draw_fitted_string(c, left_margin + 112, y, f" {pres_clean}", right_margin - (left_margin + 112), font_name="Helvetica-Bold", start_size=10, min_size=7)
    draw_dotted_after_text(c, lbl_pres, left_margin + 15, "Helvetica", 10, y - 3, right_margin)
    
    y -= 16
    lbl_prpin = "Pin No. :"
    c.setFont("Helvetica", 10)
    c.drawString(440, y, lbl_prpin)
    if pres_pin:
        c.setFont("Helvetica-Bold", 10)
        c.drawString(482, y, f" {pres_pin}")
    draw_dotted_after_text(c, lbl_prpin, 440, "Helvetica", 10, y - 3, right_margin)

    y -= line_gap
    lbl_guard = "c) Local Guardian :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 15, y, lbl_guard)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 107, y, f" {val(COL_GUARDIAN)}")
    draw_dotted_after_text(c, lbl_guard, left_margin + 15, "Helvetica", 10, y - 3, right_margin)

    y -= line_gap
    lbl_grel = "Relationship with Pupil :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 25, y, lbl_grel)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 137, y, f" {val(COL_GUARDIAN_REL)}")
    draw_dotted_after_text(c, lbl_grel, left_margin + 25, "Helvetica", 10, y - 3, 350)

    lbl_gcont = "Contact No. :"
    c.setFont("Helvetica", 10)
    c.drawString(360, y, lbl_gcont)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(422, y, f" {val(COL_GUARDIAN_CONT)}")
    draw_dotted_after_text(c, lbl_gcont, 360, "Helvetica", 10, y - 3, right_margin)

    y -= line_gap
    lbl_sib = "Siblings Details :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin + 25, y, lbl_sib)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 102, y, f" {val(COL_SIB_NAME)}")
    draw_dotted_after_text(c, lbl_sib, left_margin + 25, "Helvetica", 10, y - 3, 330)
    
    lbl_sadm = "Adm. No. :"
    c.setFont("Helvetica", 10)
    c.drawString(335, y, lbl_sadm)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(392, y, f" {val(COL_SIB_ADM)}")
    draw_dotted_after_text(c, lbl_sadm, 335, "Helvetica", 10, y - 3, 450)
    
    lbl_scls = "Class & Sec. :"
    c.setFont("Helvetica", 10)
    c.drawString(455, y, lbl_scls)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(517, y, f" {val(COL_SIB_CLASS)}")
    draw_dotted_after_text(c, lbl_scls, 455, "Helvetica", 10, y - 3, right_margin)

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
    c.drawCentredString((p1_sig_x1 + p1_sig_x2) / 2.0, y, "/")
    
    y -= 12
    c.setFont("Helvetica", 10)
    c.drawCentredString(380 + (right_margin - 20 - 380) / 2.0, y, "Signature of Parents")

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
    lbl_block = "(Name of the child in block letters)"
    c.setFont("Helvetica", 9.5)
    block_w = c.stringWidth(lbl_block, "Helvetica", 9.5)
    block_x = right_margin - block_w
    line_end_x = block_x - 8

    c.saveState()
    c.setFont("Helvetica-Bold", 11)
    c.setFillColor(colors.black)
    c.drawCentredString((left_margin + line_end_x) / 2.0, y + 2, f" {student_name}")
    c.restoreState()

    c.saveState()
    c.setDash(1, 2)
    c.line(left_margin, y, line_end_x, y)
    c.restoreState()

    c.setFont("Helvetica", 9.5)
    c.drawString(block_x, y, lbl_block)
    draw_dotted_after_text(c, lbl_block, 425, "Helvetica", 9.5, y, right_margin)

    y -= 16
    lbl_decl_dob = "Date of Birth :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin, y, lbl_decl_dob)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 77, y, f" {dob_val}")
    draw_dotted_after_text(c, lbl_decl_dob, left_margin, "Helvetica", 10, y - 3, 270)
    
    lbl_decl_words = "(in words)"
    c.setFont("Helvetica", 10)
    c.drawString(285, y, lbl_decl_words)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(337, y, f" {dob_words}")
    draw_dotted_after_text(c, lbl_decl_words, 285, "Helvetica", 10, y - 3, right_margin)

    y -= 8
    p2_text = "is entirely correct and I fully understand that the school authorities will at no time entertain any application from me or from anyone else connected with my ward for a change in this statement under any circumstances."
    p2 = Paragraph(p2_text, decl_style)
    t2 = Table([[p2]], colWidths=[right_margin - left_margin])
    t2.setStyle(TableStyle([('TOPPADDING', (0,0), (-1,-1), 0), ('BOTTOMPADDING', (0,0), (-1,-1), 0), ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0)]))
    
    t2_w, t2_h = t2.wrap(right_margin - left_margin, 100)
    t2.drawOn(c, left_margin, y - t2_h)

    y = y - t2_h - 18

    lbl_parents = "Name of parents :"
    c.setFont("Helvetica", 10)
    c.drawString(left_margin, y, lbl_parents)
    
    f_raw = val(COL_FATHER_NAME).replace("MR.", "").replace("Mr.", "").replace("MR ", "").strip()
    m_raw = val(COL_MOTHER_NAME).replace("MRS.", "").replace("Mrs.", "").replace("MS.", "").replace("Ms.", "").strip()
    parents_str = f"MR. {f_raw} & MRS. {m_raw}"
    
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin + 92, y, f" {parents_str}")
    draw_dotted_after_text(c, lbl_parents, left_margin, "Helvetica", 10, y - 3, 430)
    
    y -= 35
    c.setDash(1, 2)
    c.setLineWidth(1)
    
    p_sig_x1 = left_margin
    p_sig_x2 = left_margin + 180
    c.line(p_sig_x1, y, p_sig_x2, y)
    c.setFont("Helvetica", 10)
    c.drawCentredString((p_sig_x1 + p_sig_x2) / 2.0, y, "/")

    pr_sig_x1 = 380
    pr_sig_x2 = right_margin - 20
    c.line(pr_sig_x1, y, pr_sig_x2, y)
    
    y -= 12
    c.setFont("Helvetica", 10)
    c.drawCentredString(left_margin + 90, y, "Signature of Parents")
    c.drawCentredString(380 + (right_margin - 20 - 380) / 2.0, y, "Signature of Principal")

    c.setDash()
    c.save()
    return pdf_path

# ==========================================
# 🎫 ACKNOWLEDGEMENT RECEIPT (4-in-1) FUNCTIONS
# ==========================================
def draw_single_receipt(c, x, y_bottom, width, height, matched_row, ref_number):
    if matched_row is None:
        return
    mapped_adm_no, mapped_class_sec = get_admission_details(ref_number)

    def val(col_letter):
        idx = col_to_idx(col_letter)
        if idx < 0:
            return ''
        v = matched_row.iloc[idx] if idx < len(matched_row) else ''
        if pd.isna(v):
            return ''
        val_str = str(v).strip()
        if col_letter == COL_CLASS:
            if mapped_class_sec:
                return mapped_class_sec.upper()
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

    cell_style = ParagraphStyle('ReceiptCell', fontName='Helvetica', fontSize=8.2, leading=10.5)
    cell_bold = ParagraphStyle('ReceiptCellBold', fontName='Helvetica-Bold', fontSize=8.2, leading=10.5)

    father_display = clean_father_name(val(COL_FATHER_NAME))
    class_sec_val = val(COL_CLASS)
    addr_val = format_address(val(COL_PRES_ADDR) if val(COL_PRES_ADDR) else val(COL_PERM_ADDR))

    t_data = [
        [Paragraph("Name:", cell_bold), Paragraph(f" {val(COL_NAME)}", cell_style)],
        [Paragraph("Father's Name:", cell_bold), Paragraph(f" {father_display}", cell_style)],
        [Paragraph("Admission No:", cell_bold), Paragraph(f" {mapped_adm_no}", cell_style)],
        [Paragraph("Class & Section:", cell_bold), Paragraph(f" {class_sec_val}", cell_style)],
        [Paragraph("Address:", cell_bold), Paragraph(f" {addr_val}", cell_style)]
    ]
    
    t = Table(t_data, colWidths=[86, table_width - 86])
    t.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#333333')),
        ('TOPPADDING', (0, 0), (-1, -1), 3.8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3.8),
        ('LEFTPADDING', (0, 0), (-1, -1), 3.5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3.5),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
    ]))
    
    t.wrapOn(c, table_width, height)
    t_w, t_h = t.wrap(table_width, height)
    table_bottom_y = header_y - 48 - t_h
    t.drawOn(c, x + table_left_margin, table_bottom_y)

    p_style = ParagraphStyle('ConfStyle', fontName='Helvetica-Oblique', fontSize=8.8, leading=11.5)
    conf_text = "This is to confirm your Admission with Holy Child School as per above details and Received Rs: 1,000/- (One Thousand only) cash on account of Registration for new admission."
    p = Paragraph(conf_text, p_style)
    
    t_conf = Table([[p]], colWidths=[table_width])
    t_conf.setStyle(TableStyle([
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0)
    ]))
    _, conf_h = t_conf.wrap(table_width, 60)
    t_conf.drawOn(c, x + table_left_margin, table_bottom_y - 8 - conf_h)

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


# ==========================================
# 🚀 STREAMLIT APP UI INTERFACE & NAVIGATION
# ==========================================
def main():
    st.sidebar.title("Navigation")
    page = st.sidebar.radio("Go to", ["Admission Form Generator", "Acknowledgement Receipts (4-in-1)"])
    
    excel_file = "Online_Registration.xlsx"
    if not os.path.exists(excel_file):
        excel_file = "Online_Registration (2).xls"
        
    if page == "Admission Form Generator":
        st.title("Holy Child School — Admission Form Generator")
        
        if os.path.exists(excel_file):
            try:
                df = pd.read_excel(excel_file, header=None, engine='openpyxl' if excel_file.endswith('.xlsx') else 'xlrd')
                reg_idx = col_to_idx(COL_REG_NO)
                reg_numbers = df.iloc[1:, reg_idx].dropna().astype(str).tolist() if len(df) > 1 else []
            except Exception:
                reg_numbers = []
                
            selected_reg = st.selectbox("Select Registration Number (Column C):", reg_numbers)
            
            if st.button("Generate Admission Form PDF", type="primary"):
                if selected_reg:
                    pdf_path = generate_admission_pdf(selected_reg)
                    if pdf_path and os.path.exists(pdf_path):
                        st.success("PDF Successfully Generated! Professional Preview is loaded below:")
                        show_pdf_preview(pdf_path)
                    else:
                        st.error("Failed to generate PDF. Please check data for this registration number.")
        else:
            st.warning("Excel data file not found in the project directory.")
            
    elif page == "Acknowledgement Receipts (4-in-1)":
        st.title("Acknowledgement Receipts Hub (4-in-1)")
        
        if os.path.exists(excel_file):
            try:
                df = pd.read_excel(excel_file, header=None, engine='openpyxl' if excel_file.endswith('.xlsx') else 'xlrd')
                reg_idx = col_to_idx(COL_REG_NO)
                reg_numbers = df.iloc[1:, reg_idx].dropna().astype(str).tolist() if len(df) > 1 else []
            except Exception:
                reg_numbers = []
                
            st.write("Select up to 4 Registration Numbers for the 4-in-1 Acknowledgement Receipt page:")
            selected_ack_regs = st.multiselect("Select Registration Numbers (Max 4):", reg_numbers, max_selections=4)
            
            if st.button("Generate 4-in-1 Acknowledgement Receipt PDF", type="primary"):
                if selected_ack_regs:
                    receipt_pdf_path = generate_4_receipts_pdf(selected_ack_regs)
                    if receipt_pdf_path and os.path.exists(receipt_pdf_path):
                        st.success("4-in-1 Acknowledgement Receipts PDF Successfully Generated!")
                        show_pdf_preview(receipt_pdf_path)
                    else:
                        st.error("Failed to generate 4-in-1 receipts PDF.")
                else:
                    st.warning("Please select at least one registration number.")
        else:
            st.warning("Excel data file not found in the project directory.")

if __name__ == "__main__":
    main()
