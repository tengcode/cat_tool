import os
import openpyxl as op


def cat_excel(file_path: str, new_file_name: str):
    excel_file = []
    for file in os.listdir(file_path):
        if '.xlsx' in file:
            excel_file.append(file)

    new_file = f"{file_path}\\{new_file_name}.xlsx"
    new_xlsx = op.Workbook()
    new_sheets = []
    for file in excel_file:
        old_file = f"{file_path}/{file}"
        old_xlsx = op.load_workbook(filename=old_file)
        old_sheets = old_xlsx.sheetnames

        for old_sheet in old_sheets:
            new_sheets.append(old_sheet)
            count = new_sheets.count(old_sheet)
            if count <= 1:
                new_sheet = old_sheet
            else:
                new_sheet = f"{old_sheet}-{count}"
            new_xlsx.create_sheet(new_sheet)
            data = list(old_xlsx[old_sheet].values)
            print(data)
            for row in data:
                new_xlsx[new_sheet].append(row)

    new_xlsx.save(new_file)