import os
import openpyxl
folder="C:/Users/singhha1/Downloads/files"
for filename in os.listdir(folder):
    if filename.endswith('.xlsx'):
        file_path=os.path.join(folder,filename)
        workbook=openpyxl.load_workbook(file_path)
        sheet=workbook.active
        sheet['A11']='Punjabi'
        sheet['A12']='Hindi'
        workbook.save(file_path)
        