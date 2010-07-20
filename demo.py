from xlsx import workbook

Workbook = workbook("test.xlsx")

s = Workbook.Sheets.Sheet1
print "Total Formulas Tested: %s"%len(s.keys())
print "CELL:\tFunction,\t\tEval,\t\tOrig"

for c in s:
    cell = s[c]
    orig = cell.val
    cell.evaluate()
    if cell.fn:
        print "%s:\t%s,\t\t%s,\t\t%s"%(c,cell.fn if hasattr(cell, "fn") else "", cell.val, orig)