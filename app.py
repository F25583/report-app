import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
import re

# PAGE CONFIG
st.set_page_config(page_title="PDF to CSV", page_icon="📂", layout="wide")

st.title("📂 PDF Emails to CSV UTF-8 Converter")
st.write("Upload your Outlook emails PDF file to extract texts and tables into a clean CSV format.")

# FILE UPLOADER
uploaded_file = st.file_uploader("Choose PDF File", type=["pdf"])

def fix_visual_arabic_block(text):
    if not text:
        return ""
    
    if "(cid:" in text:
        text = re.sub(r'\(cid:\d+\)', '', text)
        
    # قلب كتل النصوص العربية فقط لتصحيح الترتيب البصري المعكوس
    reversed_line = text[::-1]
    
    # إعادة الكلمات الإنجليزية والأرقام والتواريخ لاتجاهها الصحيح المستدير
    def restore_non_arabic(match):
        return match.group(0)[::-1]
        
    non_arabic_pattern = r'[A-Za-z0-9\/:\-\_\+\=\@\.\!\?\#\$\%\&\*\(\)]+'
    fixed_line = re.sub(non_arabic_pattern, restore_non_arabic, reversed_line)
    
    return fixed_line.strip()

def clean_cell_value(val):
    if val is None:
        return ""
    # تنظيف الفواصل لمنع تداخل الأعمدة في ملف الـ CSV
    cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\n\r\t]', ' ', str(val))
    return cleaned.strip().replace(';', ',')

def process_pdf_to_csv(pdf_file):
    all_rows = []
    
    # إضافة ترويسة رئيسية للملف
    all_rows.append(["--- نصوص ومواضيع الإيميلات والجداول المستخرجة ---"])
    all_rows.append([]) # سطر فارغ

    with pdfplumber.open(pdf_file) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            
            # 1. استخراج النصوص العادية للإيميل وإصلاحها
            try:
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        line_cleaned = line.strip()
                        if len(line_cleaned) > 150 or "intended recipient" in line_cleaned.lower():
                            continue
                            
                        if line_cleaned:
                            proper_arabic_line = fix_visual_arabic_block(line_cleaned)
                            all_rows.append([proper_arabic_line])
                    all_rows.append([]) # سطر فارغ بين الصفحات
            except:
                pass

            # 2. استخراج الجداول وتحويل صفوفها
            try:
                tables = page.extract_tables()
                for table_idx, table in enumerate(tables, start=1):
                    if not table or len(table) == 0:
                        continue
                    
                    all_rows.append([f"[Table {table_idx} - Page {page_num}]"])
                    
                    for row_data in table:
                        cleaned_row = []
                        for cell_val in row_data:
                            val = clean_cell_value(cell_val)
                            if val:
                                val = fix_visual_arabic_block(val)
                            cleaned_row.append(val)
                        all_rows.append(cleaned_row)
                    
                    all_rows.append([]) # سطر فارغ بعد كل جدول
            except:
                pass

    # تحويل البيانات إلى DataFrame داتا فريم مدمج
    df_final = pd.DataFrame(all_rows)
    
    # تحويل التقرير إلى سيل نصي بصيغة CSV UTF-8 مع إضافة معرف الترميز للإكسيل (utf-8-sig)
    csv_buffer = BytesIO()
    df_final.to_csv(csv_buffer, index=False, header=False, encoding='utf-8-sig', sep=';')
    csv_buffer.seek(0)
    return csv_buffer

if uploaded_file is not None:
    with st.spinner("Processing PDF and converting to CSV UTF-8... Please wait..."):
        try:
            csv_data = process_pdf_to_csv(uploaded_file)
            st.success("Success!")
            st.download_button(
                label="📥 Download CSV UTF-8 Report",
                data=csv_data,
                file_name="Converted_Arabic_Report.csv",
                mime="text/csv"
            )
        except Exception as e:
            st.error(f"Error during processing: {e}")
