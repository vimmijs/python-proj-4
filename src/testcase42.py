import openpyxl

wk = openpyxl.Workbook()

sh = wk.active
sh.title = 'MySheetHello'
print(sh.title)

sh['B2'].value = "something"
wk.create_sheet(title="MySheet2")
sh1 = wk['MySheet2']
sh1['A2'] = "new data"

wk.remove(wk['MySheetHello'])


wk.save('C:\\Vimmi\\Job\\Study\\Python\\files\\writedata.xlsx')