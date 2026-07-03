# SmartReceive Pro Documentation

## Installation
1. Create `SmartReceive_Pro.xlsm`.
2. Import all `.bas` files from the `VBA` directory.
3. **Automated Event Handling**:
   - Open the `ThisWorkbook` module in the VBA editor.
   - Paste the code block from the top of `Mod_Utils.bas` (commented out section).
   - This ensures all scanning sheets work without individual sheet code.
4. Run `Mod_Utils.InitialSetup`.

## Key Features
- **Persisted State**: `CurrentCartonID` is tracked via the Settings sheet.
- **Live Validation**: Scans are checked against PL quantities in real-time.
- **WMS Export**: Generates a standardized CSV for external systems.
- **Clean UI**: 24pt fonts for scanning cells, large buttons for operators.
