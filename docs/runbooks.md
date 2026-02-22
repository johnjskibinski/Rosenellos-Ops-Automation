# Runbooks

## If CompanyCam duplicates appear
- Verify address normalization logic.
- Confirm street-key parsing.
- Ensure project title is "Firstname Lastname" only.
- Confirm address fields are structured (street, city, state, zip).

---

## If calendar automation runs repeatedly
- Check debounce guard.
- Verify LAST_PATCH_AT logic.
- Confirm no metadata-only patch loop.
- Ensure no duplicate triggers installed.

---

## If Measure Packet uploads fail
- Verify LP token is valid.
- Confirm correct doc type ID mapping.
- Ensure JobID in cell C1 is numeric.
- Confirm LP API credentials in Script Properties.

---

## If Balance Due does not update
- Check 12-hour refresh guard logic.
- Confirm SalesApi/GetSalesJobDetail endpoint access.
- Verify LP_BALANCE_LAST_SYNC stored in event properties.

---

## If Sheet functions don’t appear in Apps Script
- Check for syntax errors.
- Ensure file is saved.
- Refresh the spreadsheet.
