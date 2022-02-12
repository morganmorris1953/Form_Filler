import openpyxl
import os
currentFilePath = os.path.dirname(os.path.abspath(__file__))
referencePath = os.path.join(currentFilePath, 'reference')
excelFilePath = os.path.join(referencePath, 'ALPHA_ROSTER_FIELDS.xlsm')

def getExcelFileInfo(excelFilePath):
    wb_obj = openpyxl.load_workbook(excelFilePath, data_only = True)
    sheet_obj = wb_obj.active
    return sheet_obj