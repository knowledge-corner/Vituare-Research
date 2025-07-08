import xlwings as xw
from xlwings.utils import col_name

import os 
os.chdir(r"C:\Users\vaide\OneDrive - knowledgecorner.in\Course Material\Clients\Virtua Search\Excel Python Integration\Files")

wb = xw.Book("Employees.xlsx")

emp_sheet = wb.sheets["Employee"]1

last_col = col_name(emp_sheet.used_range.columns.count + 1)

emp_sheet.range(f"{last_col}1").value = "Salary"

for row in range(2, emp_sheet.used_range.rows.count + 1):
    emp_sheet.range(f"{last_col}{row}").formula = f"=Vlookup(B{row}, Salary!A:B, 2, False)"


wb.save()