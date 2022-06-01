import sys
from openpyxl import load_workbook
import datetime

wb = load_workbook('fog.xlsx')
sh = wb["Лист1"]

a = datetime.date.today()
bb = datetime.timedelta(days=3)
cc = a+bb

# sh['A3'] = cc
xx = sh['A3'].value.date()

dd = xx - a

print(dd.days)

# if a < sh['A2']:
#     print('True')
# else:
#     print('False')

wb.save(filename = 'fog.xlsx')