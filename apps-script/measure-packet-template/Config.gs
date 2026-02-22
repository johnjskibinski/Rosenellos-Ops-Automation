/**
 * MEASURE TEMPLATE CONFIG SETUP
 * Creates/overwrites a hidden CONFIG sheet with non-secret settings.
 *
 * Run: setupMeasureTemplateConfig
 */
function setupMeasureTemplateConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  upsertConfigSheet_(ss);
  SpreadsheetApp.flush();
}

function upsertConfigSheet_(ss) {
  var sh = getOrCreateSheet_(ss, "CONFIG", 0, true);

  // Header row
  sh.setFrozenRows(1);
  sh.getRange(1, 1).setValue("Key").setFontWeight("bold");
  sh.getRange(1, 2).setValue("Value").setFontWeight("bold");
  sh.getRange(1, 3).setValue("Notes").setFontWeight("bold");

  var rows = [];

  // Cell targets
  rows.push(["CELL_JOB_ID", "C1", "Numeric LeadPerfection JobID"]);
  rows.push(["CELL_GROSS_AMOUNT", "D5", "LP SalesApi/GetSalesJobDetail -> GrossAmount"]);
  rows.push(["CELL_BALANCE_DUE", "G5", "LP SalesApi/GetSalesJobDetail -> BalanceDue"]);

  // Tab names (match your sheet tabs exactly)
  rows.push(["TAB_Costing", "Costing", "Sheet tab name"]);
  rows.push(["TAB_LaborCalc", "LaborCalc", "Sheet tab name"]);
  rows.push(["TAB_WorkOrder", "Work Order", "Sheet tab name"]);
  rows.push(["TAB_WindowMeasure", "Window Measure", "Sheet tab name"]);
  rows.push(["TAB_Checklist", "Checklist", "Sheet tab name"]);

  // Print orientation
  rows.push(["PRINT_LANDSCAPE_TABS", "Costing,Window Measure", "Export as PDF landscape"]);
  rows.push(["PRINT_PORTRAIT_TABS", "Work Order,LaborCalc,Checklist", "Export as PDF portrait"]);

  // LP doc type mapping (Option A) - confirmed
  rows.push(["LP_DOC_TYPE_Window Measure", "16", "LP document type ID"]);
  rows.push(["LP_DOC_TYPE_Work Order", "26", "LP document type ID"]);
  rows.push(["LP_DOC_TYPE_Checklist", "37", "LP document type ID"]);
  rows.push(["LP_DOC_TYPE_Costing", "36", "LP document type ID"]);
  rows.push(["LP_DOC_TYPE_LaborCalc", "35", "LP document type ID"]);

  // Refresh cadence
  rows.push(["LP_BALANCE_REFRESH_MS", String(12 * 60 * 60 * 1000), "12 hours"]);

  // Calendar rule reminder
  rows.push(["CAL_BALANCE_LINE_INSTALLS_ONLY", "TRUE", "Do NOT add balance due line on measure events"]);

  // Editor emails (6 slots for silent sharing)
  rows.push(["EDITOR_EMAIL_1", "", "Add as editor silently (no email)"]);
  rows.push(["EDITOR_EMAIL_2", "", "Add as editor silently (no email)"]);
  rows.push(["EDITOR_EMAIL_3", "", "Add as editor silently (no email)"]);
  rows.push(["EDITOR_EMAIL_4", "", "Add as editor silently (no email)"]);
  rows.push(["EDITOR_EMAIL_5", "", "Add as editor silently (no email)"]);
  rows.push(["EDITOR_EMAIL_6", "", "Add as editor silently (no email)"]);

  sh.getRange(2, 1, rows.length, 3).setValues(rows);
  sh.autoResizeColumns(1, 3);

  // Hide after writing
  sh.hideSheet();
}

function getOrCreateSheet_(ss, sheetName, indexZeroBased, clearAll) {
  var sh = ss.getSheetByName(sheetName);

  if (!sh) {
    // insertSheet(index) uses 0-based index
    sh = ss.insertSheet(sheetName, indexZeroBased);
  } else {
    // Move to desired index
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(indexZeroBased + 1); // moveActiveSheet is 1-based
  }

  if (clearAll) {
    sh.clear(); // contents + formats + notes
  }

  return sh;
}

function _sanityCheck() {
  Logger.log("Sanity check OK");
}
