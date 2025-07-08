import xlwings as xw
from xlwings.utils import col_name
import pandas as pd

import os 
os.chdir(r"C:\Users\vaide\OneDrive - knowledgecorner.in\Course Material\Clients\Virtua Search\Excel Python Integration\Files")

df = pd.read_excel("employee_car.xlsx", sheet_name="data")
# Create a pivot table
pivot_table = df["Designation"].value_counts().reset_index()

wb = xw.Book("employee_car.xlsx")
ws = wb.sheets["data"]

# Create a new sheet for the dashboard
if "Dashboard" in [sheet.name for sheet in wb.sheets]:
    wb.sheets["Dashboard"].delete()
dashboard_sheet = wb.sheets.add("Dashboard", after=ws)

# Write the pivot table to the dashboard
dashboard_sheet.range("A1").options(index = False).value = pivot_table

# Create a chart

# chart = dashboard_sheet.charts.add(
#     left= dashboard_sheet.range("E2").left,
#     top= dashboard_sheet.range("E2").top,
#     width=300,
#     height=200
# )
# chart.chart_type = "column_clustered"
# chart.set_source_data(dashboard_sheet.range("B2").expand())

wb.save()
wb.close()


