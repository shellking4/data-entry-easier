import zipfile
import shutil
import os
import xml.etree.ElementTree as ET
import re
from io import BytesIO

# Namespaces
NS = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
ET.register_namespace('', NS['x'])

def get_col_letter(col_idx):
    """Convert 1-based column index to letter (e.g., 1->A, 3->C)"""
    string = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        string = chr(65 + remainder) + string
    return string

def patch_excel_sheet(template_path, output_path, data):
    """
    Patch the Excel sheet XML directly to preserve all features.
    data: list of dicts with 'model', 'qty', 'amount', 'material'
    """
    print(f"Patching {template_path} -> {output_path}")
    
    # 1. Copy template to output
    shutil.copy(template_path, output_path)
    
    # 2. Read sheet2.xml
    with zipfile.ZipFile(output_path, 'r') as zin:
        sheet_xml = zin.read('xl/worksheets/sheet2.xml')
    
    # 3. Parse XML
    root = ET.fromstring(sheet_xml)
    sheetData = root.find('x:sheetData', NS)
    
    if sheetData is None:
        print("Error: sheetData not found")
        return False
        
    # 4. Update data
    # Target columns: 
    # Description (Col 3/C)
    # Valeur (Col 4/D)
    # Quantité (Col 5/E)
    
    start_row = 6
    
    # Create a map of existing rows for quick access
    rows = {int(r.get('r')): r for r in sheetData.findall('x:row', NS)}
    
    for i, item in enumerate(data):
        row_idx = start_row + i
        
        # Get or create row
        if row_idx in rows:
            row = rows[row_idx]
        else:
            row = ET.Element(f"{{{NS['x']}}}row", {'r': str(row_idx)})
            # We need to insert it in the correct order... 
            # For simplicity, if we are appending, we can append.
            # But if we are inserting in the middle, we need to find the spot.
            # Assuming we are filling empty rows or appending.
            # Let's just append for now and sort later if needed? 
            # Excel is picky about order.
            # Finding insertion point:
            inserted = False
            for idx, child in enumerate(sheetData):
                if child.tag.endswith('row'):
                    r_num = int(child.get('r'))
                    if r_num > row_idx:
                        sheetData.insert(idx, row)
                        inserted = True
                        break
            if not inserted:
                sheetData.append(row)
            rows[row_idx] = row

        # Prepare values
        description = f"{item['model']} {item['material']}".strip()
        qty = item['qty']
        amount = item['amount']
        
        # Update cells
        update_cell(row, row_idx, 3, description, 'str')
        update_cell(row, row_idx, 4, amount, 'num')
        update_cell(row, row_idx, 5, qty, 'num')
        
    # 5. Write back to zip
    # We need to replace the file in the zip. Python's zipfile doesn't support overwrite easily.
    # We have to create a new zip.
    
    temp_zip = output_path + ".tmp"
    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(temp_zip, 'w') as zout:
            for item in zin.infolist():
                if item.filename == 'xl/worksheets/sheet2.xml':
                    zout.writestr(item, ET.tostring(root, encoding='UTF-8', xml_declaration=True))
                else:
                    zout.writestr(item, zin.read(item.filename))
    
    shutil.move(temp_zip, output_path)
    print("Patching complete.")
    return True

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
        # Insert in order
        inserted = False
        for idx, child in enumerate(row):
            if child.tag.endswith('c'):
                # Compare column indices
                # This is tricky with letters (A, B, ... AA). 
                # But we know we are only dealing with C, D, E.
                # Let's just append if simple.
                pass
        row.append(cell) # appending for now, assuming order is roughly correct or we are filling empty row
    
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
        # Remove 't' attribute if present (default is number)
        if 't' in cell.attrib:
            del cell.attrib['t']
        v_elem = ET.SubElement(cell, f"{{{NS['x']}}}v")
        v_elem.text = str(value)

if __name__ == "__main__":
    # Test data
    data = [
        {'model': 'TEST MODEL', 'material': 'TEST MAT', 'qty': 100, 'amount': 5000},
        {'model': 'TEST MODEL 2', 'material': 'TEST MAT 2', 'qty': 200, 'amount': 10000}
    ]
    
    patch_excel_sheet("IDI VIDE.xlsx", "IDI_PATCHED.xlsx", data)
