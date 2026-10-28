import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# مكتبات معالجة وإصلاح اللغة العربية
import arabic_reshaper
from bidi.algorithm import get_display
from pdfminer.high_level import extract_text as pdfminer_extract
import re

# PAGE CONFIG
st.set_page_config(page_title="PDF to Excel", page_icon="📂", layout="wide")

st.title("📂 PDF Emails to Excel Converter (Arabic Supported)")
st.write("Upload your Outlook emails PDF file to extract texts and tables correctly.")

# FILE UPLOADER
uploaded_file = st.file_uploader("Choose PDF File", type=["pdf"])

def clean_cell_value(val):
    if val is None:
        return ""
    # تنظيف النصوص من الرموز والأسهم والأسطر الجديدة لمنع أخطاء الإكسيل
    cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\n\r\t\^]', ' ', str(val))
    cleaned = cleaned.strip()
    # إذا كانت الخلية تحتوي على نص طويل جداً (مثل جملة التحذير)، نختصرها لحماية الشيت
    if len(cleaned) > 150:
        return cleaned[:150] + "..."
    return cleaned

def fix_arabic_text(text):
    if not text:
        return ""
    lines = []
    for line in text.split('\n'):
        if "(cid:" in line:
            line = re.sub(r'\(cid:\d+\)', '', line)
        if line.strip():
            try:
                reshaped_text = arabic_reshaper.reshape(line)
                bidi_text = get_display(reshaped_text)
                lines.append(bidi_text)
            except:
                lines.append(line)
    return '\n'.join(lines)

def process_pdf(pdf_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    text_font = Font(name="Calibri", size=11, color="333333")
    
    current_row = 1
    
    # قراءة نصوص الإيميلات
    pdf_bytes = pdf_file.read()
    pdf_file.seek(0)
    
    try:
        raw_text = pdfminer_extract(BytesIO(pdf_bytes))
        fixed_text = fix_arabic_text(raw_text)
    except:
        fixed_text = ""

    # كتابة النصوص المستخرجة في أول الشيت
    if fixed_text:
        ws.cell(row=current_row, column=1, value="--- نصوص ومواضيع الإيميلات المستخرجة ---").font = Font(bold=True, color="FF0000")
        current_row += 1
        for line in fixed_text.split('\n'):
            if line.strip():
                cell = ws.cell(row=current_row, column=1, value=line)
                cell.font = text_font
                cell.alignment = Alignment(horizontal="right")
                current_row += 1
        current_row += 2

    # استخراج الجداول يدوياً صفاً بصف وخلية بخلية بشكل آمن
    with pdfplumber.open(pdf_file) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            for table_idx, table in enumerate(tables, start=1):
                if not table or len(table) == 0:
                    continue
                
                ws.cell(row=current_row, column=1, value=f"[Table {table_idx} - Page {page_num}]").font = Font(bold=True)
                current_row += 1
                
                # كتابة صفوف الجدول بالكامل
                for row_idx, row_data in enumerate(table):
                    for col_idx, cell_val in enumerate(row_data, start=1):
                        clean_val = clean_cell_value(cell_val)
                        if clean_val:
                            try:
                                clean_val = get_display(arabic_reshaper.reshape(clean_val))
                            except:
                                pass
                        
                        cell = ws.cell(row=current_row, column=col_idx, value=clean_val)
                        
                        # تنسيق الصف الأول كترويسة للجدول، وبقية الصفوف كبيانات عادية
                        if row_idx == 0:
                            cell.font = header_font
                            cell.fill = header_fill
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        else:
                            cell.font = text_font
                            cell.alignment = Alignment(horizontal="right", vertical="center")
                    current_row += 1
                
                current_row += 2
                
    # ضبط تلقائي وآمن لعرض الأعمدة
    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        col_letter = get_column_letter(col_idx)
        for row_idx in range(1, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value:
                max_len = max(max_len, len(str(cell_value)))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
        
    excel_buffer = BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer

if uploaded_file is not None:
    with st.spinner("Processing Arabic PDF... Please wait..."):
        try:
            excel_data = process_pdf(uploaded_file)
            st.success("Success!")
            st.download_button(
                label="📥 Download Excel Report",
                data=excel_data,
                file_name="Converted_Arabic_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error during processing: {e}")
