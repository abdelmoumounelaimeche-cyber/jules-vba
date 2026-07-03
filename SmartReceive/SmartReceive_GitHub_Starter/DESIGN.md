# SmartReceive Data Model & Design

## Sheets & Tables
| Sheet Name | Table Name | Key Columns |
|------------|------------|-------------|
| Original PL | tblOriginalPL | CartonID, SKU, Description, Qty, Store, Provider, Status, Class |
| First scan | tblFirstScan | CartonID, Timestamp, Result, Operator |
| data Whole | tblDataWhole | CartonID, SKU, Qty, Store, Timestamp |
| SCAN WHOLE | - | UI Sheet |
| PARTIAL | tblPartial | CartonID, SKU, Qty, Store, SegregationStatus |
| segregation | - | UI Sheet |
| scan partial | - | UI Sheet |
| data after reception | tblFinalReception | CartonID, SKU, Qty, Store, Condition, Timestamp |
| OPEN Boxes | tblOpenBoxes | CartonID, Status, StartTime |
| Missing | tblMissing | CartonID, SKU, Expected, Actual, Comments |
| Settings | tblSettings | Key, Value |
| Settings | tblStores | StoreName, Code |
| Settings | tblProviders | ProviderName, Code |

## VBA Modules
- `Mod_Globals`: Constants for sheet names, table names, and status codes.
- `Mod_Settings`: Functions to get/set configuration.
- `Mod_Import`: Logic for loading PL and initial classification (Whole if 1 SKU/1 Store, Partial otherwise).
- `Mod_Scanner`: Event-driven scanning logic for the three scanning stages.
- `Mod_Workflow`: State transitions (Open -> Scanned -> Segregated -> Completed).
- `Mod_Reporting`: Summary generation.
- `Mod_UI`: Button click handlers and sheet protection management.

## Business Logic
1. **Import**: Clear previous data, load new PL, calculate if Carton is Whole or Partial.
2. **First Scan**: Scanner reads Carton ID. If Whole, go to Whole Scan. If Partial, go to Open Boxes/Segregation.
3. **Whole Scan**: Fast confirmation of whole cartons.
4. **Segregation**: Identify which SKUs go where from a partial carton.
5. **Partial Scan**: Scan individual SKUs to verify counts.
6. **Validation**: Prevent duplicate scans, alert on unknown barcodes, ensure quantities don't exceed PL.
