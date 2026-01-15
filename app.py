import streamlit as st
import pdfplumber
import openpyxl
import re
import zipfile
import shutil
import os
import xml.etree.ElementTree as ET
from io import BytesIO
import pandas as pd
import tempfile

# Namespaces
NS = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
ET.register_namespace('', NS['x'])

def get_excel_headers(template_path):
    """Extract headers from the Excel template (Row 5)"""
    wb = openpyxl.load_workbook(template_path)
    if len(wb.sheetnames) > 1:
        ws = wb[wb.sheetnames[1]]
    else:
        ws = wb.active
        
    headers = []
    header_row = 5
    for col in range(1, ws.max_column + 1):
        val = ws.cell(header_row, col).value
        if val:
            # Clean header: remove newlines, extra spaces
            clean_val = str(val).replace('\n', ' ').strip()
            headers.append((col, clean_val))
    return headers

def get_pdf_headers(pdf_file):
    """Extract headers from the first table in the PDF"""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    row_clean = [str(cell).strip() for cell in row if cell]
                    if len(row_clean) > 2:
                        return [str(cell).strip() if cell else f"Col_{i}" for i, cell in enumerate(row)]
    return []

def extract_pdf_data(pdf_file, selected_pdf_headers):
    """Extract data from PDF file object"""
    data = []
    
    # We need to find a table that contains the selected headers
    # If no headers selected, we can't find the table easily.
    if not selected_pdf_headers:
        return []

    with pdfplumber.open(pdf_file) as pdf:
        for i, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            
            for table in tables:
                header_row_idx = -1
                headers = []
                
                # Find header row
                for idx, row in enumerate(table):
                    row_values = [str(cell).strip() for cell in row if cell]
                    # Check if at least one selected header is present
                    # (Making it lenient: if ANY selected header is found, assume it's the table)
                    # Or better: check if a significant subset is found?
                    # Let's check for the first selected header found.
                    matches = sum(1 for h in selected_pdf_headers if h in row_values)
                    if matches > 0:
                        header_row_idx = idx
                        headers = [str(cell).strip() if cell else f"Col_{c_i}" for c_i, cell in enumerate(row)]
                        break
                
                if header_row_idx != -1:
                    # Map column names to indices
                    col_indices = {h: i for i, h in enumerate(headers)}
                    
                    for row in table[header_row_idx+1:]:
                        if not row or all(cell is None or cell == "" for cell in row):
                            continue
                            
                        row_data = {}
                        # Extract all columns
                        for h, idx in col_indices.items():
                            if idx < len(row):
                                row_data[h] = row[idx]
                        
                        # Only add if we have some data
                        if any(row_data.values()):
                            data.append(row_data)
    
    return data

def clean_number(value):
    """Clean numeric values"""
    if value is None:
        return 0
    if isinstance(value, (int, float)):
        return value
    
    val_str = str(value).strip()
    val_str = val_str.replace(',', '.')
    val_str = re.sub(r'[^\d.]', '', val_str)
    
    try:
        return float(val_str)
    except ValueError:
        return 0

def get_col_letter(col_idx):
    """Convert 1-based column index to letter"""
    string = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        string = chr(65 + remainder) + string
    return string

def update_cell(row, row_idx, col_idx, value, val_type):
    """Update or create a cell in the row"""
    col_letter = get_col_letter(col_idx)
    cell_ref = f"{col_letter}{row_idx}"
    
    # Find existing cell
    cell = None
    for c in row.findall('x:c', NS):
        if c.get('r') == cell_ref:
            cell = c
            break
            
    if cell is None:
        cell = ET.Element(f"{{{NS['x']}}}c", {'r': cell_ref})
        row.append(cell)
    
    # Clear children
    for child in list(cell):
        cell.remove(child)
        
    # Set value
    if val_type == 'str':
        cell.set('t', 'inlineStr')
        is_elem = ET.SubElement(cell, f"{{{NS['x']}}}is")
        t_elem = ET.SubElement(is_elem, f"{{{NS['x']}}}t")
        t_elem.text = str(value)
    else:
        # Numeric
        if 't' in cell.attrib:
            del cell.attrib['t']
        v_elem = ET.SubElement(cell, f"{{{NS['x']}}}v")
        v_elem.text = str(value)

def populate_excel(data, template_path, mapping):
    """Populate Excel file using direct XML patching
    mapping: dict {excel_col_idx: [pdf_col_names]}
    """
    
    # Create a temp file for the output
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        output_path = tmp.name
    
    shutil.copy(template_path, output_path)
    
    with zipfile.ZipFile(output_path, 'r') as zin:
        sheet_xml = zin.read('xl/worksheets/sheet2.xml')
    
    root = ET.fromstring(sheet_xml)
    sheetData = root.find('x:sheetData', NS)
    
    if sheetData is None:
        st.error("Error: sheetData not found in template")
        return None
        
    start_row = 6
    rows = {int(r.get('r')): r for r in sheetData.findall('x:row', NS)}
    
    for i, item in enumerate(data):
        row_idx = start_row + i
        
        if row_idx in rows:
            row = rows[row_idx]
        else:
            row = ET.Element(f"{{{NS['x']}}}row", {'r': str(row_idx)})
            sheetData.append(row)
            rows[row_idx] = row

        # Apply mapping
        for excel_col_idx, pdf_cols in mapping.items():
            if not pdf_cols:
                continue
                
            # Concatenate values
            parts = []
            for pdf_col in pdf_cols:
                val = item.get(pdf_col)
                if val:
                    parts.append(str(val).strip())
            
            final_val = " ".join(parts)
            
            # Determine type
            # Heuristic: if column name implies number, try to clean
            # For now, let's try to convert to float. If it works, use number.
            # Unless it's a code that looks like a number but should be string?
            # Excel handles numbers best as numbers.
            
            val_type = 'str'
            num_val = clean_number(final_val)
            
            # Check if it looks like a number (and not empty)
            if final_val and re.match(r'^-?\d+(\.\d+)?$', final_val.replace(',', '.')):
                 final_val = num_val
                 val_type = 'num'
            
            update_cell(row, row_idx, excel_col_idx, final_val, val_type)
        
    temp_zip = output_path + ".tmp"
    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(temp_zip, 'w') as zout:
            for item in zin.infolist():
                if item.filename == 'xl/worksheets/sheet2.xml':
                    zout.writestr(item, ET.tostring(root, encoding='UTF-8', xml_declaration=True))
                else:
                    zout.writestr(item, zin.read(item.filename))
    
    shutil.move(temp_zip, output_path)
    
    with open(output_path, 'rb') as f:
        output = BytesIO(f.read())
        
    os.unlink(output_path)
    return output

def main():
    st.set_page_config(page_title="PDF to Excel Converter", layout="wide")
    
    st.title("📄 PDF to Excel Converter")
    st.markdown("""
    Upload your PDF invoice to automatically fill the data into the IDI template.
    """)
    
    template_path = "IDI VIDE.xlsx"
    try:
        open(template_path, 'rb')
    except FileNotFoundError:
        st.error(f"Template file '{template_path}' not found in the directory!")
        return
    
    pdf_file = st.file_uploader("Upload PDF Invoice", type="pdf")
        
    if pdf_file:
        pdf_headers = get_pdf_headers(pdf_file)
        excel_headers = get_excel_headers(template_path)
        
        if pdf_headers and excel_headers:
            st.subheader("Column Mapping")
            st.info("Map the PDF columns to the Excel fields. You can select multiple PDF columns to combine them into one Excel field.")
            
            mapping = {}
            selected_pdf_headers = set()
            
            # Create 3 columns for layout
            cols = st.columns(3)
            
            for i, (col_idx, col_name) in enumerate(excel_headers):
                with cols[i % 3]:
                    # Default selection logic
                    default = []
                    col_name_lower = col_name.lower()
                    
                    for ph in pdf_headers:
                        ph_lower = ph.lower()
                        # Heuristics for defaults
                        if "description" in col_name_lower and ("model" in ph_lower or "material" in ph_lower):
                            default.append(ph)
                        elif "quantité" in col_name_lower and "qty" in ph_lower:
                            default.append(ph)
                        elif "valeur" in col_name_lower and "amount" in ph_lower:
                            default.append(ph)
                    
                    selection = st.multiselect(
                        f"{col_name}",
                        options=pdf_headers,
                        default=default,
                        key=f"map_{col_idx}"
                    )
                    mapping[col_idx] = selection
                    selected_pdf_headers.update(selection)
            
            if st.button("Process File", type="primary"):
                with st.spinner("Processing..."):
                    try:
                        # Extract data
                        data = extract_pdf_data(pdf_file, list(selected_pdf_headers))
                        st.success(f"Extracted {len(data)} items from PDF.")
                        
                        if data:
                            processed_excel = populate_excel(data, template_path, mapping)
                            
                            st.subheader("Preview of Data to be Written")
                            
                            # Create preview dataframe with mapped values
                            preview_rows = []
                            # Create a map of col_idx -> col_name for easy lookup
                            col_name_map = {idx: name for idx, name in excel_headers}
                            
                            for item in data:
                                row_data = {}
                                for col_idx, pdf_cols in mapping.items():
                                    if not pdf_cols:
                                        continue
                                    
                                    # Replicate the concatenation logic
                                    parts = []
                                    for pdf_col in pdf_cols:
                                        val = item.get(pdf_col)
                                        if val:
                                            parts.append(str(val).strip())
                                    
                                    final_val = " ".join(parts)
                                    col_name = col_name_map.get(col_idx, f"Col {col_idx}")
                                    row_data[col_name] = final_val
                                
                                if row_data:
                                    preview_rows.append(row_data)
                            
                            if preview_rows:
                                st.dataframe(pd.DataFrame(preview_rows))
                            else:
                                st.info("No data mapped yet.")
                            
                            st.download_button(
                                label="📥 Download Filled Excel",
                                data=processed_excel,
                                file_name="IDI_FILLED.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.warning("No data found in the PDF matching the selected columns.")
                            
                    except Exception as e:
                        st.error(f"An error occurred: {str(e)}")
        else:
            if not pdf_headers:
                st.warning("Could not detect headers in the PDF.")
            if not excel_headers:
                st.error("Could not read headers from Excel template.")
                    
    elif not pdf_file:
        st.info("Please upload a PDF file to start.")

if __name__ == "__main__":
    main()
