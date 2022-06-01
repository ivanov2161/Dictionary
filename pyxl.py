from openpyxl import load_workbook
import datetime

wb = load_workbook('test.xlsx')
sh = wb["Лист1"]
start = 1
end = 178

while start <= end:
    sh['N' + str(start)] = datetime.date.today()
    start += 1


wb.save(filename = 'test.xlsx')