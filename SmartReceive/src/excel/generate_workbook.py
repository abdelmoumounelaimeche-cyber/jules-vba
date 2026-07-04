import xlsxwriter
import os

def create_workbook():
    workbook_path = 'SmartReceive_Excel_Pro.xlsm'
    workbook = xlsxwriter.Workbook(workbook_path)

    # Download/Ensure vbaProject.bin exists from previous steps or external source
    if os.path.exists('vbaProject.bin'):
        workbook.add_vba_project('vbaProject.bin')

    # Formats
    btn_fmt = workbook.add_format({
        'bg_color': '#4F81BD',
        'font_color': 'white',
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    title_fmt = workbook.add_format({'bold': True, 'font_size': 24, 'font_color': '#1F4E78'})

    # 1. Dashboard
    ws_dashboard = workbook.add_worksheet('Dashboard')
    ws_dashboard.write('B2', 'SmartReceive Pro Dashboard', title_fmt)
    ws_dashboard.write('B4', 'Operator Name:')
    ws_dashboard.write('D4', '', workbook.add_format({'bg_color': '#EBF1DE', 'border': 1}))

    # Dashboard Buttons (Cells)
    ws_dashboard.write('B6', 'START SCANNING', btn_fmt)
    ws_dashboard.set_row(5, 30) # Row 6 height
    ws_dashboard.set_column('B:B', 20)

    ws_dashboard.write('B8', 'IMPORT SUPPLIER PL', btn_fmt)
    ws_dashboard.set_row(7, 30)

    ws_dashboard.write('B10', 'GENERATE REPORTS', btn_fmt)
    ws_dashboard.set_row(9, 30)

    # 2. Scanner UI
    ws_scanner = workbook.add_worksheet('Scanner UI')
    ws_scanner.write('B2', 'SCANNER INTERFACE', title_fmt)
    ws_scanner.write('B4', 'SCAN BARCODE HERE:', workbook.add_format({'bold': True}))
    ws_scanner.write('D4', '', workbook.add_format({'bg_color': '#FFFFCC', 'border': 2}))

    ws_scanner.write('B6', 'CURRENT MODE:', workbook.add_format({'bold': True}))
    ws_scanner.write('D6', 'First Scan', workbook.add_format({'bold': True, 'font_color': 'red'}))

    ws_scanner.write('B8', 'STATUS:', workbook.add_format({'bold': True}))
    ws_scanner.write('D8', 'Ready', workbook.add_format({'italic': True}))

    # Scanner UI Buttons
    ws_scanner.write('F6', 'SWITCH MODE', btn_fmt)
    ws_scanner.write('F8', 'CLEAR INPUT', btn_fmt)

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

if __name__ == "__main__":
    create_workbook()
