from xlsx import workbook

Workbook = workbook("test.xlsx")

@Workbook.extend
def SUM(lst):
    _n = 0
    for n in lst:
        try:
            _n += n
        except:
            pass
    return _n

Workbook.Sheets.Sheet1.F1.fn = "A1:D1"
print Workbook.Sheets.Sheet1.F1.evaluate()

Workbook.Sheets.Sheet1.F1.fn = "SUM(A1:D1)*4"
print Workbook.Sheets.Sheet1.F1.evaluate()