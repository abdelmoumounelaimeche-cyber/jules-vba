# SmartReceive Pro - Test Checklist

## Module: Dashboard
- [ ] Dashboard displays correctly on open.
- [ ] Operator name can be entered.
- [ ] Navigation buttons work.

## Module: PL Import & Classification
- [ ] Importing a PL populates `tblOriginalPL`.
- [ ] Cartons with 1 SKU are marked "Whole".
- [ ] Cartons with >1 SKU are marked "Partial".
- [ ] Duplicate Carton IDs in PL are flagged.

## Module: First Scan
- [ ] Scanning a valid Carton ID records a row in `tblFirstScan`.
- [ ] Scanning an unknown Carton ID shows an error.
- [ ] Correct classification is displayed to the operator.

## Module: Whole Scan
- [ ] Scanning a "Whole" carton records data in `tblDataWhole` and `tblDataAfterReception`.
- [ ] Scanning a "Partial" carton in the "Whole Scan" screen shows an error.
- [ ] Prevent duplicate scans of the same carton.

## Module: Segregation & Open Boxes
- [ ] Can create a new repacked box in `tblOpenBoxes`.
- [ ] Assigning items updates the `Current SKU Count`.
- [ ] Suggestion logic returns a valid box name.

## Module: Partial Scan
- [ ] Scanning items into a box aggregates the quantity.
- [ ] Finalizing a box moves data to `tblDataAfterReception` and clears the partial scan table.
- [ ] Box status changes to "Closed".

## Module: Reports
- [ ] WMS Export CSV is generated with correct data.
- [ ] Discrepancy report correctly identifies missing SKUs/quantities.

## Technical
- [ ] VBA project compiles without errors.
- [ ] All tables have correct headers and names.
- [ ] Error handling catches common warehouse mistakes (wrong scanner input, etc).
