# SmartReceive Pro - Operator Guide

## Introduction
SmartReceive Pro is an automated warehouse receiving application built in Excel. It streamlines the process from Supplier Packing List (PL) import to WMS Export.

## Getting Started
1. Open `SmartReceive_Pro.xlsm`.
2. Ensure macros are enabled.
3. On the **Dashboard**, enter your name in the designated cell.

## Workflow

### 1. PL Import
- Use the **Settings** sheet or the Dashboard instruction to trigger **Import PL**.
- Select the supplier's packing list file (.csv).
- The system will automatically classify cartons as **Whole** or **Partial**.

### 2. Scanning (Scanner UI)
- All scanning is done on the **Scanner UI** sheet.
- **First Scan**: Scan the carton barcode. The system will tell you if it's Whole or Partial.
- **Whole Scan**: Change mode to "Whole Scan" and scan the carton.
- **Segregation**: Change mode to "Segregation". Scan the carton, then scan each item. The system will suggest a box.
- **Partial Scan**: Change mode to "Partial Scan". Scan the box name, then scan all items inside it.

### 3. Reports & Export
- Once finished, go to the main menu/buttons to **Generate WMS Export**.
- The file will be saved in your configured folder.

## Security
- Technical sheets are hidden and protected.
- Admin Password: `SR2026`
