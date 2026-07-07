# SmartReceive Excel Pro - Testing Checklist

- [ ] **Data Cleaning**: Verify that trailing spaces and casing are normalized during import.
- [ ] **Classification**: Confirm that cartons with >1 SKU are correctly sent to 'Partial' table.
- [ ] **Scanner Validation**:
    - [ ] Scan an unknown carton (should trigger error).
    - [ ] Scan a duplicate carton (should trigger error).
- [ ] **Pallet Logic**: Verify that 'Next Pallet' correctly increments (e.g., C-1 to C-2).
- [ ] **Partial Aggregation**: Scan the same SKU 3 times in Partial Scan (should result in 1 row with Qty 3).
- [ ] **WMS Export**: Verify that the generated CSV contains aggregated data from both Whole and Partial scans.
- [ ] **UI/UX**: Verify that 'Segregation Cards' and 'Assignment Displays' are clearly visible.
