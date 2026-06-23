import openpyxl

wb = openpyxl.load_workbook('KernelV0_v2_copy.xlsx', data_only=True)
ws = wb['ANCHOR']

print(f'ANCHOR sheet dimensions: {ws.dimensions}')
print('\n--- Full content of ANCHOR sheet ---')
for row in ws.iter_rows():
    for cell in row:
        if cell.value is not None:
            print(f'  {cell.coordinate}: {repr(cell.value)}')

# Also read the SumMatrix to understand accumulation
ws2 = wb['SumMatrix']
print(f'\n\nSumMatrix dimensions: {ws2.dimensions}')
print('--- First few rows of SumMatrix ---')
for r in range(1, 6):
    for cell in ws2[r]:
        if cell.value is not None:
            print(f'  {cell.coordinate}: {repr(str(cell.value))[:80]}')
