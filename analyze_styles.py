import zipfile
import xml.etree.ElementTree as ET

def analyze_styles():
    template = "IDI VIDE.xlsx"
    
    with zipfile.ZipFile(template, 'r') as z:
        # 1. Get style of C6 from sheet2.xml
        sheet_xml = z.read('xl/worksheets/sheet2.xml')
        root = ET.fromstring(sheet_xml)
        ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        
        c6_style = None
        sheetData = root.find('x:sheetData', ns)
        for row in sheetData.findall('x:row', ns):
            if row.get('r') == '6':
                for c in row.findall('x:c', ns):
                    if c.get('r') == 'C6':
                        c6_style = c.get('s')
                        break
                break
        
        print(f"Style index for C6: {c6_style}")
        
        if c6_style is None:
            print("C6 not found or has no style.")
            return

        # 2. Read styles.xml
        styles_xml = z.read('xl/styles.xml')
        styles_root = ET.fromstring(styles_xml)
        cellXfs = styles_root.find('x:cellXfs', ns)
        xfs = list(cellXfs.findall('x:xf', ns))
        
        target_idx = int(c6_style)
        target_xf = xfs[target_idx]
        
        print(f"Target XF attributes: {target_xf.attrib}")
        target_align = target_xf.find('x:alignment', ns)
        if target_align is not None:
            print(f"Target Alignment: {target_align.attrib}")
            
        # Key attributes to preserve
        border_id = target_xf.get('borderId')
        fill_id = target_xf.get('fillId')
        font_id = target_xf.get('fontId')
        num_fmt_id = target_xf.get('numFmtId')
        
        print(f"Looking for style with: Border={border_id}, Fill={fill_id}, Font={font_id}, NumFmt={num_fmt_id}, Align=Left")
        
        # 3. Search for matching left-aligned style
        candidates = []
        for i, xf in enumerate(xfs):
            if (xf.get('borderId') == border_id and 
                xf.get('fillId') == fill_id and 
                xf.get('fontId') == font_id):
                
                align = xf.find('x:alignment', ns)
                if align is not None and align.get('horizontal') == 'left':
                    candidates.append(i)
                    print(f"Found candidate at index {i}: {align.attrib}")
                    
        if candidates:
            print(f"Best candidate index: {candidates[0]}")
        else:
            print("No exact match found.")

if __name__ == "__main__":
    analyze_styles()
