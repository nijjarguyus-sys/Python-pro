import os
import openpyxl
folder = 'C:/Users/singhha1/Downloads/files'
output_file = 'C:/Users/singhha1/Downloads/files/Candidates Scores.xlsx'

output_wb = openpyxl.Workbook()
output_sheet = output_wb.active
output_sheet.title = 'candidate Data'

cells = ['B2', 'B3', 'C8','C9','C10','C11','C12','C13']
for filename in os.listdir(folder):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(folder, filename)
        workbook = openpyxl.load_workbook(file_path)
        values = [workbook.active[cell].value for cell in cells]
        output_sheet.append(values)
output_wb.save(output_file)
