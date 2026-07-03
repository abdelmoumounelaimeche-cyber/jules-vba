# SmartReceive Excel Pro - User Guide

## Introduction
SmartReceive Excel Pro is a professional Excel application designed to automate warehouse receiving processes. It features scanner-first workflows, automated carton classification, and real-time validation.

## Getting Started
1. Open `SmartReceive_Excel_Pro.xlsm`.
2. Enable Macros if prompted.
3. Import the VBA modules provided in the `vba/` folder if they are not already present.

## Dashboard
The Dashboard provides a high-level overview of the shipment progress, including KPIs for total cartons, whole/partial classification, and scan progress. Use the navigation buttons to move between modules.

## Workflow
### 1. PL Import
Go to the **Original PL** sheet and click **Import PL**. The system will automatically:
- Clean the data.
- Detect duplicates.
- Classify cartons into **Whole** (single SKU) or **Partial** (multi-SKU).

### 2. First Scan
As cartons enter the warehouse, scan their barcodes in the **First Scan** sheet. The system validates that the carton belongs to the current shipment.

### 3. Whole Scan
For whole cartons, use the **Whole Scan** sheet.
- Scan the carton ID.
- The system will display the target **Store** and assigned **Pallet** (e.g., C-1) in large, high-contrast fonts.
- Click **NEXT PALLET** when a pallet is full to increment the numbering.

### 4. Segregation & Partial Scan
Partial cartons are sorted in the **Segregation** area.
- The **Segregation Card** displays the items that should be in the carton.
- Use the **Partial Scan** sheet to scan items into new repacked boxes.
- Scans are aggregated by SKU (multiple scans increment quantity).

## Reports
Generate reports with one click from the Dashboard:
- **WMS Export**: CSV file ready for WMS import.
- **Missing Report**: Identifies items expected in the PL but not scanned.

## Support
For technical issues, refer to the `vba/` source modules for error handling details.
