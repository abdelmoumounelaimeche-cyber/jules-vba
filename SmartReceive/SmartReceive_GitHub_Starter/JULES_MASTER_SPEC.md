# SmartReceive Excel Pro

## Mission
Build a professional Excel (.xlsm) warehouse receiving application that automates the existing workbook while preserving business logic.

### Workflow
Supplier PL -> Import -> Whole/Partial Classification -> Open Box -> First Scan -> Whole Scan -> Segregation -> Partial Scan -> Reports -> WMS Export

### Objectives
- Eliminate manual VLOOKUPs and copy/paste
- Scanner-first workflow
- Fast operator UI
- Automatic validation
- Professional reports

### Modules
1. Dashboard
2. PL Import
3. Open Boxes
4. First Scan
5. Whole Scan
6. Segregation
7. Partial Scan
8. Reports
9. Settings

### Existing Sheets
Original PL
First scan
data Whole
SCAN WHOLE
PARTIAL
segregation
scan partial
data after reception
OPEN Boxes
Missing

Treat these as the reference business process.

### Technical Requirements
- Microsoft Excel 365
- VBA
- Power Query
- Excel Tables
- Dynamic Arrays
- No external add-ins

### Key Business Rules
- Carton unique per shipment
- CTN may contain multiple SKUs
- Whole cartons go directly to store pallets
- Partial cartons require segregation
- Barcode scanner acts as keyboard
- Partial scans aggregate quantities
- Repacked box names must be unique
- Missing/Damaged require comments

### Deliverables
- Production-ready .xlsm
- VBA modules
- UserForms
- Documentation
- Test checklist
