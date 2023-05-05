import win32com.client as win32

# open the workbook
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open('example.xlsx')

# select the worksheet with the pivot table
ws = wb.Worksheets('PivotTable')

# select the pivot table
pivot = ws.PivotTables(1)

# set the border style
border = win32.constants.xlThin

# set the font style
font = 'Calibri'

# set the alignment
align = win32.constants.xlCenter

# set the fill color
fill = win32.constants.xlColorIndexNone

# apply styles to the pivot table cells
for cell in pivot.TableRange1:
    cell.Borders.Weight = border
    cell.Font.Name = font
    cell.HorizontalAlignment = align
    cell.Interior.ColorIndex = fill

# save the workbook
wb.Save()
wb.Close()
excel.Quit()
