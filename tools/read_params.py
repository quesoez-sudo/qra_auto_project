import openpyxl
import zipfile
import xml.etree.ElementTree as ET

wb_f = openpyxl.load_workbook('KernelV0_v2_copy.xlsx', data_only=False)
wb_d = openpyxl.load_workbook('KernelV0_v2_copy.xlsx', data_only=True)

ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

# Read AA column (rows 1-30) for each sheet, and AK column for JF
sheets_params = {
    'ImpactThermMatrix': ('AA', 'xl/worksheets/sheet7.xml'),
    'ImpactToxMatrix':   ('AA', 'xl/worksheets/sheet8.xml'),
    'ImpactFFMatrix':    ('AA', 'xl/worksheets/sheet9.xml'),
    'ImpactExpMatrix':   ('AA', 'xl/worksheets/sheet10.xml'),
    'ImpactJFMatrix':    ('AK', 'xl/worksheets/sheet11.xml'),
}

with zipfile.ZipFile('KernelV0_v2_copy.xlsx', 'r') as z:
    for sheet_name, (pcol, xml_path) in sheets_params.items():
        ws_d = wb_d[sheet_name]
        ws_f = wb_f[sheet_name]
        print(f'\n{"="*60}')
        print(f'SHEET: {sheet_name}  (param col: {pcol})')
        print(f'{"="*60}')
        # Read rows 1-30 of the param column
        for r in range(1, 35):
            cell_f = ws_f[f'{pcol}{r}']
            cell_d = ws_d[f'{pcol}{r}']
            fval = cell_f.value
            dval = cell_d.value
            if fval is not None or dval is not None:
                # Also get the label from adjacent column (prev col)
                prev_col = chr(ord(pcol[0]) - 1) if len(pcol) == 1 else None
                label = ''
                if prev_col:
                    label_cell = ws_d[f'{prev_col}{r}']
                    label = label_cell.value or ''
                print(f'  {pcol}{r}: value={repr(dval):<25}  formula={repr(str(fval))[:60]}  | label={repr(str(label))[:40]}')

        # Also read some labels from Z column (or AJ for JF)
        label_col = 'Z' if pcol == 'AA' else 'AJ'
        print(f'\n  --- {label_col} column (labels) ---')
        for r in range(1, 35):
            cell = ws_d[f'{label_col}{r}']
            if cell.value is not None:
                print(f'  {label_col}{r}: {repr(cell.value)}')
