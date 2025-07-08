import xlwings as xw
import pandas as pd

def value_count():
    wb = xw.Book.caller()
    # wb = xw.Book('my_udfs.xlsm')
    # sheet = wb.sheets["data"]

    rng = wb.app.selection

    if "Pivot Data" in [sheet.name for sheet in wb.sheets]:
        wb.sheets["Pivot Data"].delete()
    pivot = wb.sheets.add("Pivot Data", after=wb.sheets[-1])       
    
    # Validation checks
    try :
        if rng.columns.count != 1:
            raise ValueError("Please select a single column.")
        if rng.rows.count < 2:
            raise ValueError("Selection must include a header and at least one data row.")

        data = rng.options(ndim=1).expand('down').value[1:]  

        df = pd.Series(data)
        counts = df.value_counts().reset_index()
        counts.columns = ['Value', 'Count']  

        pivot.range("A1").options(index = False).value = counts.values.tolist()

    except ValueError as e:
        print(f"Error: {e}")
        pivot.range("A1").value = [[str(e)]]


