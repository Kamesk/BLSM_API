from collections import OrderedDict
import openpyxl
from openpyxl import workbook



path = r"C:\Users\Public\AppData\TT_Hygiene_Check.xlsx"
workbook = openpyxl.load_workbook(path)
sheet = workbook.get_sheet_by_name('Hygiene')
sheet.cell(row=2, column=17).value = "data"
workbook.save(path)



