import openpyxl
from datetime import datetime
from openpyxl.drawing.image import Image
import os

wb = openpyxl.load_workbook("files/invoice_data.xlsx",data_only=True)
ws = wb.active

values = list(ws.values)

lastrow = len(values)

wb = openpyxl.load_workbook("files/invoice.xlsx")
ws = wb.active


current_date = datetime.now()


year_month = current_date.strftime("%Y%m")

invoice_month = current_date.month

output_folder = f"請求書_{current_date.strftime('%Y%m月')}"

os.makedirs(output_folder,exist_ok=True)

output_file = f"{output_folder}/請求書_{current_date.strftime('%Y年%m月')}.xlsx"

invoice_number = 1

for index in range(lastrow):
    if not index ==0:
        if values[index][12]is None:
            continue

        sheet_name = str(values[index][0])
        #ワークシートを新規で追加
        copy_ws = wb.copy_worksheet(ws)
        copy_ws.title = sheet_name
        copy_ws["A2"].value = sheet_name
        copy_ws["A4"].value = values[index][10]
        copy_ws["B7"].value = f"{invoice_month}月分請"
        copy_ws["N2"].value = f"{year_month}-{invoice_number:03d}"
        invoice_number += 1

        copy_ws["A14"].value = values[index][13]

        img = Image("files/角印.png")
        img.width = 100
        img.height = 100
        copy_ws.add_image(img,"P5")

ws = wb["請求書"]
wb.remove(ws)
wb.save(output_file)
