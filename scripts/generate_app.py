import xlsxwriter

def create_workbook():
    workbook = xlsxwriter.Workbook('SmartReceive_Excel_Pro.xlsx')

    # Formats
    title_fmt = workbook.add_format({'bold': True, 'font_size': 24, 'font_color': '#FFFFFF', 'bg_color': '#1F4E78', 'align': 'center', 'valign': 'vcenter', 'border': 2})
    kpi_label_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center', 'font_size': 12})
    kpi_val_fmt = workbook.add_format({'bold': True, 'font_size': 18, 'border': 1, 'align': 'center', 'font_color': '#1F4E78'})
    btn_fmt = workbook.add_format({'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white', 'border': 2, 'align': 'center', 'valign': 'vcenter'})

    # High Contrast for Operator Screens
    op_label_fmt = workbook.add_format({'bold': True, 'font_size': 20, 'bg_color': '#000000', 'font_color': '#FFFF00', 'align': 'center'})
    op_val_fmt = workbook.add_format({'bold': True, 'font_size': 48, 'bg_color': '#000000', 'font_color': '#FFFFFF', 'align': 'center', 'border': 5, 'border_color': '#FFFF00'})
    scan_box_fmt = workbook.add_format({'bg_color': '#CCFFCC', 'border': 2, 'font_size': 20})

    # 1. Dashboard
    ws_dash = workbook.add_worksheet('Dashboard')
    ws_dash.set_column('B:G', 20)
    ws_dash.merge_range('B2:G3', 'SmartReceive Excel Pro', title_fmt)

    kpis = [
        ('Total Cartons', 'KPI_TotalCartons', 'C5'),
        ('Whole Cartons', 'KPI_WholeCartons', 'C7'),
        ('Scanned Whole', 'KPI_ScannedWhole', 'C9'),
        ('Partial Cartons', 'KPI_PartialCartons', 'E5'),
        ('Missing Items', 'KPI_Missing', 'E7'),
        ('Damaged Items', 'KPI_Damaged', 'E9'),
    ]

    for label, name, cell in kpis:
        ws_dash.write(cell, label, kpi_label_fmt)
        ws_dash.write(cell[0] + str(int(cell[1:])+1), 0, kpi_val_fmt)
        workbook.define_name(name, f'=Dashboard!{cell[0]}${int(cell[1:])+1}')

    buttons = [('IMPORT PL', 'G5'), ('FIRST SCAN', 'G6'), ('WHOLE SCAN', 'G7'), ('SEGREGATION', 'G8'), ('REPORTS', 'G9')]
    for text, cell in buttons:
        ws_dash.write(cell, text, btn_fmt)

    # 2. Original PL (Technical - Hidden)
    ws_pl = workbook.add_worksheet('Original PL')
    ws_pl.add_table('A1:E2', {
        'name': 'tblOriginalPL',
        'columns': [{'header': 'Store'}, {'header': 'Carton_ID'}, {'header': 'SKU'}, {'header': 'Description'}, {'header': 'Qty'}]
    })
    ws_pl.hide()

    # 3. Data Whole (Technical - Hidden)
    ws_whole = workbook.add_worksheet('Data Whole')
    ws_whole.add_table('A1:F2', {
        'name': 'tblDataWhole',
        'columns': [{'header': 'Carton_ID'}, {'header': 'Store'}, {'header': 'SKU'}, {'header': 'Qty'}, {'header': 'Status'}, {'header': 'Pallet_ID'}]
    })
    ws_whole.hide()

    # 4. Partial (Technical - Hidden)
    ws_partial = workbook.add_worksheet('Partial')
    ws_partial.add_table('A1:E2', {
        'name': 'tblPartial',
        'columns': [{'header': 'Carton_ID'}, {'header': 'Store'}, {'header': 'SKU'}, {'header': 'Qty'}, {'header': 'Status'}]
    })
    ws_partial.hide()

    # 5. Open Boxes
    ws_open = workbook.add_worksheet('Open Boxes')
    ws_open.add_table('A1:E2', {
        'name': 'tblOpenBoxes',
        'columns': [{'header': 'Carton_ID'}, {'header': 'SKU'}, {'header': 'Qty'}, {'header': 'Condition'}, {'header': 'Comments'}]
    })
    ws_open.conditional_format('D2:D1000', {'type': 'text', 'criteria': 'containing', 'value': 'DAMAGED', 'format': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})})
    ws_open.conditional_format('D2:D1000', {'type': 'text', 'criteria': 'containing', 'value': 'MISSING', 'format': workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})})

    # 6. First Scan
    ws_fscan = workbook.add_worksheet('First Scan')
    ws_fscan.write('E1', 'SCAN CARTON HERE:', kpi_label_fmt)
    ws_fscan.write('E2', '', scan_box_fmt)
    ws_fscan.add_table('A1:C2', {
        'name': 'tblFirstScan',
        'columns': [{'header': 'Timestamp'}, {'header': 'Carton_ID'}, {'header': 'Status'}]
    })

    # 7. SCAN WHOLE
    ws_swhole = workbook.add_worksheet('SCAN WHOLE')
    ws_swhole.set_column('F:H', 35)
    ws_swhole.write('E1', 'SCAN CARTON HERE:', kpi_label_fmt)
    ws_swhole.write('E2', '', scan_box_fmt)
    ws_swhole.write('F4', 'STORE', op_label_fmt)
    ws_swhole.write('F5', '', op_val_fmt)
    workbook.define_name('UI_StoreDisplay', "='SCAN WHOLE'!$F$5")
    ws_swhole.write('F7', 'PALLET', op_label_fmt)
    ws_swhole.write('F8', '', op_val_fmt)
    workbook.define_name('UI_PalletDisplay', "='SCAN WHOLE'!$F$8")
    ws_swhole.write('F10', 'NEXT PALLET', btn_fmt)
    ws_swhole.add_table('A1:D2', {
        'name': 'tblScanWhole',
        'columns': [{'header': 'Timestamp'}, {'header': 'Carton_ID'}, {'header': 'Store'}, {'header': 'Pallet_ID'}]
    })

    # 8. Segregation (Card Only, Table Hidden)
    ws_seg = workbook.add_worksheet('Segregation')
    ws_seg.set_column('G:I', 30)
    labels = [('CARTON', 'G4'), ('STORE', 'G6'), ('SKU', 'G8'), ('QTY', 'G10')]
    for lbl, cell in labels:
        ws_seg.write(cell, lbl, op_label_fmt)
        ws_seg.write(cell[0] + str(int(cell[1:])+1), '', op_val_fmt)
    workbook.define_name('Card_CartonID', '=Segregation!$G$5')
    workbook.define_name('Card_Store', '=Segregation!$G$7')
    workbook.define_name('Card_SKU', '=Segregation!$G$9')
    workbook.define_name('Card_Qty', '=Segregation!$G$11')
    ws_seg.add_table('A1:E2', {
        'name': 'tblSegregation',
        'columns': [{'header': 'Carton_ID'}, {'header': 'Store'}, {'header': 'SKU'}, {'header': 'Qty'}, {'header': 'Status'}]
    })

    # 9. Scan Partial
    ws_spart = workbook.add_worksheet('Scan Partial')
    ws_spart.write('E1', 'REPACK BOX ID:', kpi_label_fmt)
    ws_spart.write('E2', 'BOX-001', scan_box_fmt)
    ws_spart.write('E3', 'SCAN SKU HERE:', kpi_label_fmt)
    ws_spart.write('E4', '', scan_box_fmt)
    ws_spart.add_table('A1:E2', {
        'name': 'tblScanPartial',
        'columns': [{'header': 'Timestamp'}, {'header': 'Repacked_Box'}, {'header': 'Store'}, {'header': 'SKU'}, {'header': 'Qty'}]
    })

    # 10. Data After Reception (Hidden)
    ws_final = workbook.add_worksheet('Data After Reception')
    ws_final.add_table('A1:E2', {
        'name': 'tblFinalData',
        'columns': [{'header': 'Store'}, {'header': 'SKU'}, {'header': 'Qty'}, {'header': 'Carton_Type'}, {'header': 'Reference'}]
    })
    ws_final.hide()

    # 11. Missing
    ws_miss = workbook.add_worksheet('Missing')
    ws_miss.add_table('A1:D2', {
        'name': 'tblMissing',
        'columns': [{'header': 'Carton_ID'}, {'header': 'SKU'}, {'header': 'Qty'}, {'header': 'Reason'}]
    })

    # 12. Settings (Hidden)
    ws_set = workbook.add_worksheet('Settings')
    ws_set.add_table('A1:B10', {
        'name': 'tblSettings',
        'columns': [{'header': 'Setting_Name'}, {'header': 'Setting_Value'}]
    })
    ws_set.hide()

    workbook.close()
    print("Workbook SmartReceive_Excel_Pro.xlsx updated with hidden sheets and scan areas.")

if __name__ == '__main__':
    create_workbook()
