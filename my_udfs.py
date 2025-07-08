'''
Creating a Python UDF - 

- Creat a Python file with the function you want to use in Excel
- Use xlwings to create a UDF (User Defined Function) that can be called from Excel
- Save the Python file and use it in Excel by calling the function in a cell

Steps - 
1. pip install xlwings
2. Create a Python file (e.g., my_udfs.py) with the function you want to use
    - Use the @xw.func decorator to define the function as a UDF
3. In your terminal/command prompt:- xlwings addin install
4. Open Excel → Go to xlwings tab → Click Import Functions
5. Set Interpreter Path
    - Go to the xlwings tab
    - Click Settings
    - Set the Python Interpreter path
4. Call the Python Function in Excel
'''
import xlwings as xw

import pandas as pd

@xw.func
def create_value_counts(col_range):
    data = pd.Series(col_range)
    counts = data.value_counts().reset_index()
    counts.columns = ['Value', 'Count']
    return counts.values.tolist()

@xw.func
def hello():
    return "Hello from Python!"