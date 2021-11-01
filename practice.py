# import openpyxl
# import datetime
import pyautogui as pag

# # ここからExcel操作
# workbook = "GMBRank.xlsx" # ファイル名を指定
# worksheet = "シャンゴ問屋町本店" # シート名を指定

# wb = openpyxl.load_workbook(workbook)
# sheet = wb[worksheet]

# max = sheet.max_row


# maxcol = sheet.max_column
# # for key in range(maxcol - 1):
# #   print(sheet.cell(4, key + 2).value)

# if True:
#   rank = 1
# print(rank)


# # sheet.cell(row = max + 1, column = 1).value = str(datetime.date.today())

# # 保存する
# wb.save(workbook)

pag.moveTo(300,300)
pag.scroll(-2000)