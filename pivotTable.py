import win32com.client as win32

def clear_pts(ws):
    for pt in ws.PivotTables():
        pt.TableRange2.Clear()

def insert_pt_field_set1(pt):
    field_rows = {}
    field_rows['program area'] = pt.PivotFields("Program Area")
    field_rows['agency name'] = pt.PivotFields("Agency Name")

    field_values = {}
    field_values['total'] = pt.PivotFields("Total")
    field_values['program area'] = pt.PivotFields("Program Area")

    # insert row fields
    # https://docs.microsoft.com/en-us/office/vba/api/excel.xlpivotfieldorientation
    field_rows['program area'].Orientation = 1
    field_rows['program area'].Position = 1

    field_rows['agency name'].Orientation = 1
    field_rows['agency name'].Position = 2

    # insert values field
    # https://docs.microsoft.com/en-us/office/vba/api/excel.xlconsolidationfunction
    field_values['program area'].Orientation = 4
    field_values['program area'].Function = -4112 # value reference
    field_values['program area'].NumberFormat = "#,##0"

    field_values['total'].Orientation = 4
    field_values['total'].Function = -4157 # value reference
    field_values['total'].NumberFormat = "$#,##0"


# construct the Excel application object
xlApp = win32.Dispatch('Excel.Application')
xlApp.Visible = True

# Create wb and ws 
wb = xlApp.Workbooks.Open(r"C:\Users\Jie\Google Drive\_To Upload\__Ready Folder\Automate Excel Pivot Table With Python\exercise.xlsm")
# wb = xlApp.ActiveWorkbook
ws_data = wb.Worksheets("Data")
ws_report = wb.Worksheets("Demo")


# clear pivot tables on Report tab
clear_pts(ws_report)

# create pt cache connection
# https://docs.microsoft.com/en-us/office/vba/api/excel.xlpivottablesourcetype
pt_cache = wb.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)

# insert pivot table designer/editor
pt = pt_cache.CreatePivotTable(ws_report.Range("B3"), "Dev_Proj_Summary")

# toggle grand totals
pt.ColumnGrand = True
pt.RowGrand = False

# change subtotal location
# https://docs.microsoft.com/en-us/office/vba/api/excel.xlsubtotallocationtype
pt.SubtotalLocation(2) # bottom

# use Tabular as the report layout
# https://docs.microsoft.com/en-us/office/vba/api/excel.xllayoutrowtype
pt.RowAxisLayout(1)

# update table style
pt.TableStyle2 = "PivotStyleMedium9"

insert_pt_field_set1(pt)

# wrap up
ws_report.Columns("D:E").ColumnWidth = 35
