import zipfile
import xml.etree.ElementTree as ET

ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

targets = {
    'ImpactThermMatrix': ('xl/worksheets/sheet7.xml',  ['AB1']),
    'ImpactToxMatrix':   ('xl/worksheets/sheet8.xml',  ['AB1', 'H1', 'J1']),
    'ImpactFFMatrix':    ('xl/worksheets/sheet9.xml',  ['AB1']),
    'ImpactExpMatrix':   ('xl/worksheets/sheet10.xml', ['AB1']),
    'ImpactJFMatrix':    ('xl/worksheets/sheet11.xml', ['AL1']),
}

def col_to_num(col):
    n = 0
    for c in col:
        n = n * 26 + (ord(c) - ord('A') + 1)
    return n

def cell_ref_to_idx(ref):
    import re
    m = re.match(r'([A-Z]+)(\d+)', ref)
    col = col_to_num(m.group(1))
    row = int(m.group(2))
    return row, col

with zipfile.ZipFile('KernelV0_v2_copy.xlsx', 'r') as z:
    for sheet_name, (xml_path, cells) in targets.items():
        print(f'\n{"="*70}')
        print(f'SHEET: {sheet_name}')
        print(f'{"="*70}')
        xml_data = z.read(xml_path)
        root = ET.fromstring(xml_data)

        # Find all cells in row 1 with formulas
        for row_elem in root.findall('.//x:row', ns):
            row_num = int(row_elem.get('r', 0))
            if row_num != 1:
                continue
            for c_elem in row_elem.findall('x:c', ns):
                ref = c_elem.get('r', '')
                f_elem = c_elem.find('x:f', ns)
                if f_elem is not None and f_elem.text:
                    ftype = f_elem.get('t', 'normal')
                    ref_range = f_elem.get('ref', '')
                    print(f'\n  Cell {ref} [type={ftype}, ref={ref_range}]:')
                    print(f'  {f_elem.text}')
