# Assembly Instructions for SmartReceive Excel Pro

Since this application relies on VBA macros, you must follow these steps to assemble the final `.xlsm` file from the provided components.

## 1. Generate the Base Workbook
Run the following command to generate the latest Excel structure:
\`\`\`bash
python3 scripts/generate_app.py
\`\`\`
This creates `SmartReceive_Excel_Pro.xlsx`.

## 2. Convert to Macro-Enabled Workbook
1. Open `SmartReceive_Excel_Pro.xlsx` in Microsoft Excel.
2. Go to **File > Save As**.
3. Choose **Excel Macro-Enabled Workbook (*.xlsm)**.
4. Save as `SmartReceive_Excel_Pro.xlsm`.

## 3. Import VBA Modules
1. Press **ALT + F11** to open the VBA Editor.
2. In the Project Explorer, right-click on `VBAProject (SmartReceive_Excel_Pro.xlsm)`.
3. Select **Import File...**.
4. Import all `.bas` files from the `vba/` directory:
   - `modImport.bas`
   - `modScanner.bas`
   - `modUI.bas`
   - `modReport.bas`
   - `modUtils.bas`

## 4. Final Setup
1. Assign the macros to the Dashboard buttons (Right-click button > Assign Macro):
   - **IMPORT PL** -> `modImport.ImportAndProcessPL`
   - **REPORTS** -> `modReport.GenerateWMSExport`
   - **NEXT PALLET** -> `modUI.NextPallet`
2. Save the workbook.

Your professional warehouse application is now ready for use!
