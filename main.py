from openpyxl import load_workbook

file_path = 'pyscp.xlsx'

wb = load_workbook(file_path)

ws = wb.active

ws['A1'] = "harvey"
ws['B1'] = "jiang"
ws['A2'] = "wilson"
ws['B2'] = "boooochen"
ws['A3'] = "rayn"
ws['B3'] = "zanhe"

final_name_row = 1

ws.insert_cols(3)

for index, row in enumerate(ws.iter_rows(values_only=True)):
    first_initial = row[0][0]
    last_name = row[1]
    combined_name = [first_initial, last_name]
    end_string = ''.join(combined_name)
    print(end_string)
    ws[f"C{final_name_row}"] = end_string
    print(final_name_row)
    final_name_row += 1

ws.delete_cols(4)
for row in ws.iter_rows(min_row=1,  max_row=3, values_only=True):
    print(row)

wb.save(file_path)


