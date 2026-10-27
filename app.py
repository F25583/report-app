import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment

# إعدادات الصفحة الأساسية للواجهة
st.set_page_config(
    page_title="PDF to Excel Converter", 
    page_icon="📂", 
    layout="wide"
)

# واجهة المستخدم الرسومية بالألوان والتنسيق
st.markdown("""
    <style>
    .main-title {
        font-size: 32px;
        font-weight: bold;
        color: #1F4E78;
        text-align: center;
        margin-bottom: 20px;
    }
    .sub-title {
        font-size: 18px;
        color: #555555;
        text-align: center;
        margin-bottom: 40px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📂 أداة تحويل إيميلات PDF إلى تقرير إكسيل منظم</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">قم برفع ملف الـ PDF المستخرج من أوت لوك لتحويله إلى تقرير Excel مقسم وميُسر للقراءة</div>', unsafe_allow_html=True)

# أداة رفع الملفات من جهازك
uploaded_file = st.file_uploader("اختر ملف الـ PDF الخاص بالإيميلات:", type=["pdf"])

def process_pdf(pdf_file):
    # إنشاء ملف إكسيل جديد في الذاكرة
    wb = Workbook()
    ws = wb.active
    ws.title = "Email Report"
    
    # تنسيقات الخلايا والجداول
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    text_font = Font(name="Calibri", size=11, italic=True, color="333333")
    
    current_row = 1
    
    with pdfplumber.open(pdf_file) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            # 1. استخراج النص (موضوع الإيميل)
            text = page.extract_text(layout=True)
            if text:
                ws.cell(row=current_row, column=1, value=f"--- صفحة رقم {page_num} (نص الإيميل) ---").font = Font(bold=True, color="FF0000")
                current_row += 1
                for line in text.split('\n'):
                    if line.strip():
                        ws.cell(row=current_row, column=1, value=line).font = text_font
                        current_row += 1
                current_row += 1
            
            # 2. استخراج الجداول المرافقة
            tables = page.extract_tables()
            for table_idx, table in enumerate(tables, start=1):
                if not table:
                    continue
                ws.cell(row=current_row, column=1, value=f"[جدول {table_idx} - صفحة {page_num}]").font = Font(bold=True)
                current_row += 1
                
                df = pd.DataFrame(table)
                if len(df) > 1:
                    df.columns = df.iloc[0]
                    df = df[1:]
                
                # كتابة ترويسة الجدول وتلوينها
                for col_idx, header in enumerate(list(df.columns), start=1):
                    cell = ws.cell(row=current_row, column=col_idx, value=str(header))
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center")
                current_row += 1
                
                # كتابة محتويات الجدول
                for r in dataframe_to_rows(df, index=False, header=False):
                    for col_idx, value in enumerate(r, start=1):
                        cell = ws.cell(row=current_row, column=col_idx, value=value)
                        cell.alignment = Alignment(horizontal="left")
                    current_row += 1
                current_row += 2
                
    # ضبط أوتوماتيكي لعرض الأعمدة
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col.column_letter].width = max(max_len + 3, 12)
        
    excel_buffer = BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer

# تفعيل زر المعالجة والتنزيل عند رفع الملف
if uploaded_file is not None:
    with st.spinner("جاري قراءة الـ PDF وتنسيق الجداول... بروجاء الانتظار"):
        try:
            excel_data = process_pdf(uploaded_file)
            st.success("✨ تم استخراج البيانات وتنسيق التقرير بنجاح!")
            
            # زر التنزيل المباشر من المتصفح
            st.download_button(
                label="📥 تحميل ملف Excel المنسق",
                data=excel_data,
                file_name="Outlook_Converted_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"حدث خطأ أثناء المعالجة: {e}")
