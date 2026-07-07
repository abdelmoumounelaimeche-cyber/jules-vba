# SmartReceive Excel Pro Architecture

## Worksheets & Tables

| Worksheet | Table Name | Columns | Description |
|-----------|------------|---------|-------------|
| Dashboard | - | - | User interface with buttons and KPIs |
| Original PL | tblOriginalPL | Store, Carton_ID, SKU, Description, Qty | Raw data imported from supplier |
| Data Whole | tblDataWhole | Carton_ID, Store, SKU, Qty, Status, Pallet_ID | List of whole cartons to be scanned |
| Partial | tblPartial | Carton_ID, Store, SKU, Qty, Status | List of cartons requiring segregation |
| Open Boxes | tblOpenBoxes | Carton_ID, SKU, Qty, Condition, Comments | Inspection log for open boxes |
| First Scan | tblFirstScan | Timestamp, Carton_ID, Status | Log of all cartons entering the warehouse |
| SCAN WHOLE | tblScanWhole | Timestamp, Carton_ID, Store, Pallet_ID | Log of whole cartons scanned to pallets |
| Segregation | tblSegregation | Carton_ID, Store, SKU, Qty, Status | Assistance for sorting partial cartons |
| Scan Partial | tblScanPartial | Timestamp, Repacked_Box, Store, SKU, Qty | Log of items scanned into new boxes |
| Data After Reception | tblFinalData | Store, SKU, Qty, Carton_Type, Reference | Final aggregated data for WMS |
| Missing | tblMissing | Carton_ID, SKU, Qty, Reason | Report of missing items |
| Settings | tblSettings | Setting_Name, Setting_Value | Configuration parameters |

## VBA Modules

- **modImport**: Handles Excel/CSV import of the Purchase List. Performs data cleaning and logic to split cartons into 'Whole' or 'Partial'.
- **modScanner**: Centralizes logic for barcode scans. Validates inputs against expected data and updates status in real-time.
- **modUI**: Manages Dashboard updates, sheet protection, navigation, and the high-visibility "Segregation Card" display.
- **modReport**: Aggregates data from various tables to generate formatted PDF/Excel reports and the WMS CSV export.
- **modUtils**: Contains shared utility functions, table handlers, error logging, and global constants.

## Business Logic Rules

1. **Whole vs Partial**: A carton is 'Whole' if it contains only one SKU and the quantity matches the standard pack (or simply if it's marked as such in PL).
2. **Scanner Validation**: Every scan must check for duplicates and existence in the PL.
3. **Palletization**: During Whole Scan, items are assigned to Pallets (e.g., C-1, C-2) which are store-specific.
4. **Partial Aggregation**: In Partial Scan, multiple scans of the same SKU in the same Repacked Box increment the quantity rather than adding rows.
