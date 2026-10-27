import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# PAGE CONFIG
st.set_page_config(page_title="PDF to Excel", page_icon="📂", layout="wide")

st.title("📂 PDF Emails to Excel Converter")
st.write("Upload your Outlook emails PDF file to extract text and tables automatically.")

# FILE UPLOADER
uploaded_file = st.file_uploader("Choose PDF File", type=["pdf"])

def process_pdf(pdf_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    text_font = Font(name="Calibri", size=11, italic=True, color="333333")
    
    current_row = 1
    
    with pdfplumber.open(pdf_file) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            # Extract Text
            text = page.extract_text()
            if text:
                ws.cell(row=current_row, column=1, value=f"--- Page {page_num} Text ---").font = Font(bold=True, color="FF0000")
                current_row += 1
                for line in text.split('\n'):
                    if line.strip():
                        ws.cell(row=current_row, column=1, value=line).font = text_font
                        current_row += 1
                current_row += 1
            
            # Extract Tables
            tables = page.extract_tables()
            for table_idx, table in enumerate(tables, start=1):
                if not table:
                    continue
                ws.cell(row=current_row, column=1, value=f"[Table {table_idx} - Page {page_num}]").font = Font(bold=True)
                current_row += 1
                
                df = pd.DataFrame(table)
                if len(df) > 1:
                    df.columns = df.iloc[0]
                    df = df[1:]
                
                for col_idx, header in enumerate(list(df.columns), start=1):
                    cell = ws.cell(row=current_row, column=col_idx, value=str(header if header else ''))
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center")
                current_row += 1
                
                for r in dataframe_to_rows(df, index=False, header=False):
                    for col_idx, value in enumerate(r, start=1):
                        ws.cell(row=current_row, column=col_idx, value=value)
                    current_row += 1
                current_row += 2
                
    # Safe Auto-fit column widths
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
    with st.spinner("Processing PDF... Please wait..."):
        try:
            excel_data = process_pdf(uploaded_file)
            st.success("Success!")
            st.download_button(
                label="📥 Download Excel Report",
                data=excel_data,
                file_name="Converted_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error during processing: {e}")
