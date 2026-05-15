import os
import openpyxl
folder = "C:/Users/singhha1/Downloads/files"
for filename in os.listdir(folder):
    if filename.endswith('.xlsx'):
        file_path=os.path.join(folder,filename)
        workbook=openpyxl.load_workbook(file_path)
        sheet=workbook.active
        sheet['B14']= '=Sum(B8:B13)'
        sheet['C14']= '=Sum(C8:C13)'
        sheet['B14'].number_format = '#,##0_ ;-##0'
        sheet['C14'].number_format = '#,##0_ ;-##0'
        workbook.save(file_path)