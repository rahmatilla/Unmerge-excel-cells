import pandas as pd
import glob, os

path = 'C:/Users/Rakhmatilla/PycharmProjects/webScrap/excel_files/'
target_path = 'C:/Users/Rakhmatilla/PycharmProjects/webScrap/excel_files/result_directory/'
files = glob.glob(os.path.join(path, '*.xlsx'))
if not os.path.exists(target_path):
    os.makedirs(target_path)

print(files)

for file in files:
    wb = pd.ExcelFile(file)
    file_name = os.path.basename(file)

    writer = pd.ExcelWriter(target_path + file_name, engine='xlsxwriter')

    for i in range(len(wb.sheet_names)):
        sheet_name = wb.sheet_names[i]
        ws = wb.parse(sheet_name=sheet_name)
        ws = ws.fillna(method='ffill')
        ws.to_excel(writer, sheet_name=sheet_name, index=False)

    writer.save()
