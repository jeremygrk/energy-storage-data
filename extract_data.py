import openpyxl, json, sys

wb = openpyxl.load_workbook('data.xlsx')
all_data = {}

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    headers = []
    for cell in ws[1]:
        headers.append(str(cell.value) if cell.value is not None else '')

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(c is None or str(c).strip() == '' for c in row):
            continue
        clean_row = []
        for c in row:
            if c is None:
                clean_row.append(None)
            elif isinstance(c, str) and c.startswith('='):
                clean_row.append(None)
            elif isinstance(c, (int, float)):
                clean_row.append(round(c, 4))
            else:
                clean_row.append(str(c))
        rows.append(clean_row)

    all_data[sheet_name] = {'headers': headers, 'rows': rows}
    print(f'{sheet_name}: {len(rows)} rows, {len(headers)} cols')

with open('data.json', 'w', encoding='utf-8') as f:
    json.dump(all_data, f, ensure_ascii=False)

print('Done - data.json updated.')
