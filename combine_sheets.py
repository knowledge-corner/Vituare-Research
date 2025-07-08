import xlwings as xw

import os 
os.chdir(r"C:\Users\vaide\OneDrive - knowledgecorner.in\Course Material\Clients\Virtua Search\Excel Python Integration\Files")

wb = xw.Book("monthly_data.xlsx")
if "Combined" in [sheet.name for sheet in wb.sheets]:
    wb.sheets["Combined"].delete()

output_sheet = wb.sheets.add("Combined", after=wb.sheets[-1])

output_sheet.range("A1:B1").value = wb.sheets[0].range("A1:B1").value  # Copy headers from the first sheet

row_index = 2
for sheet in wb.sheets:
    if sheet.name == "Combined":
        continue

    data = sheet.used_range.value[1:]
    if data :
        output_sheet.range(f"A{row_index}").value = data
        row_index += len(data)

# Save the workbook after combining sheets
wb.save()
wb.close()
