from xlsx import workbook

Workbook = workbook("Book1.xlsx")

print Workbook
print Workbook.Sheets
print Workbook.Sheets.Sheet1
print Workbook.Sheets["Sheet1"]
print Workbook.Sheets["Sheet23"]
print Workbook.Sheets.Sheet1.A1
print Workbook.Sheets.Sheet1["A1"]
print Workbook.Sheets.Sheet1["FF"]


total = 0
for sheet in Workbook.Sheets:
    for cell in sheet:
        total = total + int(sheet[cell])
        
print total