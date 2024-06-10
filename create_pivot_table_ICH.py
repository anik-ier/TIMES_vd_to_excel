import win32com.client as win32
win32c = win32.constants

# function to clear any existing pivot table
def clear_pt(ws):
    for pt in ws.PivotTables():
        pt.TableRange2.Clear()

# function to toggle grand totals
def toggle_grand_totals(pt):
    pt.ColumnGrand = True
    pt.RowGrand = False

# function for pivot table row
def pivot_table_row(pt, row_name: str):
    pt_rows = pt.PivotFields(row_name)
    pt_rows.Orientation = 1  # win32c.xlRowField  # 1 for row of the pivot table
    pt_rows.Position = 1

# Pivot table column
def pivot_table_column(pt):
    pt_col = pt.PivotFields('Period')
    pt_col.Orientation = 2  # win32c.xlColumnField  # 2 for column field

# function for insert data fields/Pivot table data
def pivot_table_insert_data(pt, data_PV:str):
    pt_data = pt.PivotFields(data_PV)
    pt_data.Orientation = 4  # win32c.xlDataField
    # pt_data.Function = win32c.xlSum

# Pivot table filters
def pivot_table_filter(pt):
    pt_filter_fields = [('Scenario', 3), ('Attribute', 3), ('Commodity', 3)]  # 3 for xlPageField(filter)
    for field, orientation in pt_filter_fields:
        pt_filter = pt.PivotFields(field)
        pt_filter.Orientation = orientation
    print('Pivot table is created')


# instance of Excel application
xlApp = win32.Dispatch('Excel.Application')
xlApp.Visible = True

'''Please, provide location of Excel file where you want to create PIVOT TABLE; 
this Excel should also contain results data and Pivot table sheet'''
# excel workbook which contains results data and pivot table sheet
wb = xlApp.Workbooks.Open(r'C:/Users/ac141435/Desktop/TAM_Idustrie/TAM-Industry/results/TAM_ICH.xlsx')

'''Please, write worksheet name that contains all_results data'''
ws_data = wb.Sheets('all_results')  # reference (all_result) data worksheets
ws_data_range = ws_data.Range('A1').CurrentRegion

'''Please, write worksheet name where you would like to create pivot table'''
pivot_table_ws = wb.Worksheets('pivot_table')  # pivot table worksheet

# create pt cache connection
pt_cache = wb.PivotCaches().Create(1, ws_data_range)

# clear pivot tables
clear_pt(pivot_table_ws)

# create pivot table editor
pivot_table = pt_cache.CreatePivotTable(pivot_table_ws .Range('A6'), 'AI_Trial')

# toggle grand totals
toggle_grand_totals(pivot_table)

# Pivot table row
pivot_table_row(pivot_table, 'Process')

# pivot table columns
pivot_table_column(pivot_table)

# insert data fields/Pivot table data
pivot_table_insert_data(pivot_table, 'PV')

# pivot table filter
pivot_table_filter(pivot_table)




