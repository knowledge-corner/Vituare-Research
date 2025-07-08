import xlwings as xw

import os 
os.chdir(r"C:\Users\vaide\OneDrive - knowledgecorner.in\Course Material\Clients\Virtua Search\Excel Python Integration\Files")

wb = xw.Book("employee_car_macro.xlsm")

macro = wb.macro("CreatePivotCarOwnershipByGender")
macro()

wb.save()
wb.close()

# Install an excel add-in (macro function) as a command button in custom ribbon
'''
To Turn This Into an Add-in - 
- Open Excel → Press Alt + F11
- Paste the macro in a new module in a blank workbook
- Save the workbook as an Excel Add-in (.xlam) file and close the VBA editor and file without saving
- Load the add-in in Developer → Excel Add-ins Browse and add the add-in and click on check box to enable it
- Add the macro to a custom tab in the ribbon via Options → Customize Ribbon
- Add a new tab and group, then add the macro to the group
- Now you can run the macro from the custom ribbon tab

Note - This cannot trigger a python code directly, but you can use the macro to call a python script using xlwings or other methods.

'''