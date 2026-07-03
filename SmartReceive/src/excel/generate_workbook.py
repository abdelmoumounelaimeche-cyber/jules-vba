import xlsxwriter
import os

def create_workbook():
    workbook_path = 'SmartReceive_Excel_Pro.xlsm'
    workbook = xlsxwriter.Workbook(workbook_path)

    # Enable VBA
    if os.path.exists('vbaProject.bin'):
        workbook.add_vba_project('vbaProject.bin')

    # 1. Dashboard
    ws_dashboard = workbook.add_worksheet('Dashboard')
    ws_dashboard.set_tab_color('red')
    ws_dashboard.write('B2', 'SmartReceive Pro', workbook.add_format({'bold': True, 'font_size': 24}))
    ws_dashboard.write('B4', 'Operator Name:')
    ws_dashboard.write('D4', '', workbook.add_format({'bg_color': '#EBF1DE', 'border': 1}))
    ws_dashboard.write('B6', 'Instructions:')
    ws_dashboard.write('B7', '1. Enter name above.')
    ws_dashboard.write('B8', '2. Go to "Scanner UI" sheet.')

    # 2. Scanner UI
    ws_scanner = workbook.add_worksheet('Scanner UI')
    ws_scanner.set_tab_color('blue')
    ws_scanner.write('B2', 'SCANNER INTERFACE', workbook.add_format({'bold': True, 'font_size': 18}))
    ws_scanner.write('B4', 'Scanner Input:')
    ws_scanner.write('C4', '', workbook.add_format({'bg_color': '#FFFFCC', 'border': 1}))
    ws_scanner.write('B6', 'Current Mode:')
    ws_scanner.write('C6', 'First Scan')
    ws_scanner.write('B7', 'Status:')
    ws_scanner.write('C7', 'Ready')
    ws_scanner.write('B8', 'Target Box (Partial):')
    ws_scanner.write('C8', '')

    # Data Sheets
    sheets_config = [
        ('Original PL', 'tblOriginalPL', ['Carton ID', 'SKU', 'Description', 'Expected Qty', 'Store', 'Status', 'Type']),
        ('First scan', 'tblFirstScan', ['Timestamp', 'Carton ID', 'Operator', 'Classification']),
        ('data Whole', 'tblDataWhole', ['Carton ID', 'SKU', 'Qty', 'Store', 'Scan Timestamp']),
        ('segregation', 'tblSegregation', ['Source Carton', 'SKU', 'Qty', 'Destination Box', 'Status', 'Operator']),
        ('scan partial', 'tblScanPartial', ['Box ID', 'SKU', 'Scanned Qty', 'Timestamp', 'Operator']),
        ('data after reception', 'tblDataAfterReception', ['Box/Carton ID', 'SKU', 'Final Qty', 'Store', 'Type', 'Reception Date']),
        ('OPEN Boxes', 'tblOpenBoxes', ['Box Name', 'Current SKU Count', 'Status']),
        ('Missing', 'tblMissing', ['Carton ID', 'SKU', 'Missing Qty', 'Comment', 'Reported By']),
        ('Settings', 'tblSettings', ['Parameter', 'Value'])
    ]

    for sheet_name, table_name, cols in sheets_config:
        ws = workbook.add_worksheet(sheet_name)
        ws.add_table('A1:' + chr(64 + len(cols)) + '10', {
            'name': table_name,
            'columns': [{'header': c} for c in cols]
        })

    workbook.close()
    print(f"Workbook {workbook_path} created with VBA project shell.")

if __name__ == "__main__":
    create_workbook()
