import zipfile
import xml.etree.ElementTree as ET

ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

sheet_map = {
    'ImpactThermMatrix': ('xl/worksheets/sheet7.xml', 'AB1'),
    'ImpactToxMatrix':   ('xl/worksheets/sheet8.xml', 'AB1'),
    'ImpactFFMatrix':    ('xl/worksheets/sheet9.xml', 'AB1'),
    'ImpactExpMatrix':   ('xl/worksheets/sheet10.xml', 'AB1'),
    'ImpactJFMatrix':    ('xl/worksheets/sheet11.xml', 'AL1'),
}

# First discover which xml files correspond to which sheets
with zipfile.ZipFile('KernelV0_v2_copy.xlsx', 'r') as z:
    # Read workbook to get sheet order
    wb_xml = z.read('xl/workbook.xml')
    wb_root = ET.fromstring(wb_xml)
    sheets_list = wb_root.findall('.//x:sheet', ns)
    print('Sheet order:')
    for i, s in enumerate(sheets_list, 1):
        name = s.get('name')
        rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        print(f'  sheet{i}: {name}  (rId={rid})')

    # Read relationships to map rId to file
    rels_xml = z.read('xl/_rels/workbook.xml.rels')
    rels_root = ET.fromstring(rels_xml)
    rId_map = {}
    for r in rels_root:
        rId_map[r.get('Id')] = r.get('Target')
    print('\nRelationships:')
    for k, v in rId_map.items():
        print(f'  {k} -> {v}')
