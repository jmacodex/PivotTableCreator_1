import pandas as pd
import win32com.client as win32

df = pd.read_csv('consumptionstatemon.csv')
df.columns.str.strip()
df.rename(columns={df.columns[5]:"ENERGY SOURCE(UNITS)"},inplace=True)

def pivotFields(pt):
    row_field = {}
    row_field['STATE'] = pt.PivotFields("STATE")
    row_field['STATE'].Orientation = 1
    row_field['STATE'].Position = 1
    row_field['TYPE OF PRODUCER'] = pt.PivotFields("TYPE OF PRODUCER")
    row_field['TYPE OF PRODUCER'].Orientation = 1
    row_field['TYPE OF PRODUCER'].Position = 2
    row_field['ENERGY SOURCE(UNITS)'] = pt.PivotFields("ENERGY SOURCE(UNITS)")
    row_field['ENERGY SOURCE(UNITS)'].Orientation = 1
    row_field['ENERGY SOURCE(UNITS)'].Position = 3
    
    col_field ={}
    col_field['MONTH'] = pt.PivotFields("MONTH")
    col_field['MONTH'].Orientation = 2
    col_field['MONTH'].Position = 1
    
    filter_field = {}
    filter_field['YEAR'] = pt.PivotFields("YEAR")
    filter_field['YEAR'].Orientation = 3
    filter_field['YEAR'].Position = 1
    # Auto select the last item on filter field
    auto_select = filter_field['YEAR'].Pivotitems(filter_field['YEAR'].Pivotitems.Count).Name
    filter_field['YEAR'].CurrentPage = auto_select
    
    value_field = {}
    value_field['CONSUMPTION'] = pt.PivotFields("CONSUMPTION")
    value_field['CONSUMPTION'].Orientation = 4
    value_field['CONSUMPTION'].Function = -4157
    value_field['CONSUMPTION'].NumberFormat = "#,##0"
    
for i in df['STATE'].drop_duplicates():
    filtered_state= df[(df['STATE'] == i )]
    # Save excel file
    with pd.ExcelWriter('Reports\\Power Consumption - {}.xlsx'.format(i)) as writer:
        filtered_state.to_excel(writer,index=False,sheet_name='Data')
        
    # Open the excel file using pywin32
    xlApp = win32.Dispatch('Excel.Application')
    xlApp.Visible = True
    wb = xlApp.Workbooks.Open(r'C:\\Users\\samso\\Projects\\PivotTableCreator_1\\Reports\\Power Consumption - {}.xlsx'.format(i)) # path name the reports
    # Add new sheet and prepare the Data for Pivot Table
    add_new_sheet = wb.Sheets.Add(Before=None,After=wb.Sheets(wb.Sheets.count))
    new_sheet = add_new_sheet.Name = 'Pivot_Table'
    ws_data = wb.Worksheets('Data')
    ws_report = wb.Worksheets(new_sheet)
    
    pt_cache = wb.PivotCaches().Create(1, ws_data.Range("A1").CurrentRegion)
    pt = pt_cache.CreatePivotTable(ws_report.Range("B3"),'Pivot Table')
    
    pt.ColumnGrand = True
    pt.RowGrand = True
    
    pt.RowAxisLayout(1)
    
    pt.TableStyle2 = "PivotStyleMedium9"
    
    pivotFields(pt)
    ws_report.Columns.Autofit
    wb.Save()
    wb.Close()
    xlApp.Quit()
    print(i," --saved.")

    
    
    