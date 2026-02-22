/***************
 * ROSENELLO SIDEBAR UI (REAL)
 * - Reads CONFIG dynamically
 * - UI is live now
 * - Upload/LP calls are stubs for safety (we wire API next)
 ***************/

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("Rosenello Tools")
    .addItem("Open LP Upload Sidebar", "openLpSidebar")
    .addToUi();
}

/**
 * Assign THIS function to your Drawing button on each tab.
 */
function openLpSidebar() {
  SpreadsheetApp.getUi()
    .showSidebar(
      HtmlService.createHtmlOutputFromFile("Sidebar")
        .setTitle("LP Upload / Print")
    );
}

/**
 * Sidebar model: CONFIG + tab list + mapping list.
 */
function getSidebarModel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getConfigSheet_(ss);

  const config = readConfig_(sh);

  const activeTabName = ss.getActiveSheet().getName();
  const sheetTabs = ss.getSheets().map(s => s.getName());

  // Build mappings dynamically:
  // We take all keys like LP_DOC_TYPE_{TABNAME} and map them.
  const mappings = [];

  const landscapeSet = csvToSet_(config.PRINT_LANDSCAPE_TABS || "");
  const portraitSet  = csvToSet_(config.PRINT_PORTRAIT_TABS || "");

  Object.keys(config).forEach(k => {
    if (k.indexOf("LP_DOC_TYPE_") === 0) {
      const tab = k.substring("LP_DOC_TYPE_".length);
      const docTypeId = String(config[k] || "").trim();

      if (!docTypeId) return;

      let orientation = "";
      if (landscapeSet.has(tab)) orientation = "LANDSCAPE";
      else if (portraitSet.has(tab)) orientation = "PORTRAIT";

      mappings.push({ tab, docTypeId, orientation });
    }
  });

  return {
    configSheetName: sh.getName(),
    lpBaseUrl: config.LP_BASE_URL || "", // configurable later
    activeTabName,
    sheetTabs,
    mappings
  };
}

/**
 * Save/update a mapping:
 * - sets LP_DOC_TYPE_{TAB} = id
 * - updates PRINT_LANDSCAPE_TABS / PRINT_PORTRAIT_TABS to include TAB
 */
function saveTabMapping(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getConfigSheet_(ss);
  const cfg = readConfig_(sh);

  const tab = String(payload.tab || "").trim();
  const docTypeId = String(payload.docTypeId || "").trim();
  const orientation = String(payload.orientation || "").trim().toUpperCase();
  const notes = String(payload.notes || "").trim();

  if (!tab) throw new Error("Tab is required.");
  if (!/^\d+$/.test(docTypeId)) throw new Error("DocTypeId must be numeric.");
  if (orientation !== "LANDSCAPE" && orientation !== "PORTRAIT") throw new Error("Orientation must be LANDSCAPE or PORTRAIT.");

  // Write LP_DOC_TYPE_{TAB}
  upsertConfigRow_(sh, `LP_DOC_TYPE_${tab}`, docTypeId, notes || "LP document type ID");

  // Update orientation CSV lists
  const landscape = csvToSet_(cfg.PRINT_LANDSCAPE_TABS || "");
  const portrait  = csvToSet_(cfg.PRINT_PORTRAIT_TABS || "");

  // remove from both first
  landscape.delete(tab);
  portrait.delete(tab);

  if (orientation === "LANDSCAPE") landscape.add(tab);
  if (orientation === "PORTRAIT") portrait.add(tab);

  upsertConfigRow_(sh, "PRINT_LANDSCAPE_TABS", setToCsv_(landscape), "Export as PDF landscape");
  upsertConfigRow_(sh, "PRINT_PORTRAIT_TABS", setToCsv_(portrait), "Export as PDF portrait");

  return { ok: true, message: `Saved mapping: ${tab} → DocType ${docTypeId} (${orientation}).` };
}

/**
 * Remove a mapping:
 * - deletes LP_DOC_TYPE_{TAB}
 * - removes TAB from both orientation lists
 */
function removeTabMapping(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getConfigSheet_(ss);
  const cfg = readConfig_(sh);

  const tab = String(tabName || "").trim();
  if (!tab) throw new Error("Tab is required.");

  deleteConfigKey_(sh, `LP_DOC_TYPE_${tab}`);

  const landscape = csvToSet_(cfg.PRINT_LANDSCAPE_TABS || "");
  const portrait  = csvToSet_(cfg.PRINT_PORTRAIT_TABS || "");
  landscape.delete(tab);
  portrait.delete(tab);

  upsertConfigRow_(sh, "PRINT_LANDSCAPE_TABS", setToCsv_(landscape), "Export as PDF landscape");
  upsertConfigRow_(sh, "PRINT_PORTRAIT_TABS", setToCsv_(portrait), "Export as PDF portrait");

  return { ok: true, message: `Removed mapping for: ${tab}` };
}

/***************
 * STUB ACTIONS (safe placeholders)
 * We wire these to LP API upload + token + PDF export next.
 ***************/

function refreshLpAmounts() {
  // Later: use JobID from C1, call SalesApi/GetSalesJobDetail, write GrossAmount to D5 and BalanceDue to G5.
  return {
    ok: true,
    message:
      "Refresh stub ✅\n" +
      "Next wiring step:\n" +
      "• Read JobID from CONFIG CELL_JOB_ID (C1)\n" +
      "• Call LP SalesApi/GetSalesJobDetail\n" +
      "• Write GrossAmount → D5, BalanceDue → G5"
  };
}

function uploadCurrentTab() {
  const tab = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  return uploadSingleTab(tab);
}

function uploadAllTabs() {
  // Later: iterate mappings in CONFIG in your desired order (Costing, Window Measure, Work Order, LaborCalc, Checklist)
  return {
    ok: true,
    message:
      "Upload FULL packet stub ✅\n" +
      "Next wiring step:\n" +
      "• Read mappings LP_DOC_TYPE_* from CONFIG\n" +
      "• Export each mapped tab to 1-page PDF with correct orientation\n" +
      "• Upload each PDF using LP API UploadDocument (Option A)"
  };
}

function uploadSingleTab(tabName) {
  // Later: export only this tab and upload with its doc type.
  return {
    ok: true,
    message:
      `Upload stub ✅\nTab: ${tabName}\n` +
      "Next wiring step:\n" +
      "• Find LP_DOC_TYPE_{TAB} in CONFIG\n" +
      "• Export tab → PDF (portrait/landscape from CONFIG)\n" +
      "• Upload via LP API UploadDocument (NOT web endpoint)"
  };
}

/***************
 * CONFIG HELPERS
 ***************/

function getConfigSheet_(ss) {
  let sh = ss.getSheetByName("CONFIG");
  if (!sh) throw new Error("CONFIG sheet not found. Run your setupMeasureTemplateConfig_() first.");
  return sh;
}

function readConfig_(sh) {
  const values = sh.getDataRange().getValues();
  // Expect header row: Key | Value | Notes
  const out = {};
  for (let r = 1; r < values.length; r++) {
    const key = String(values[r][0] || "").trim();
    const val = values[r][1];
    if (!key) continue;
    out[key] = (val === null || typeof val === "undefined") ? "" : String(val);
  }
  return out;
}

/**
 * Upsert key/value/notes into CONFIG.
 */
function upsertConfigRow_(sh, key, value, notes) {
  const data = sh.getDataRange().getValues();
  const keyStr = String(key || "").trim();
  if (!keyStr) throw new Error("Config key required.");

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0] || "").trim() === keyStr) {
      sh.getRange(r + 1, 2).setValue(value);
      if (typeof notes !== "undefined") sh.getRange(r + 1, 3).setValue(notes);
      return;
    }
  }

  // append
  const newRow = sh.getLastRow() + 1;
  sh.getRange(newRow, 1, 1, 3).setValues([[keyStr, value, notes || ""]]);
}

function deleteConfigKey_(sh, key) {
  const data = sh.getDataRange().getValues();
  const keyStr = String(key || "").trim();
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0] || "").trim() === keyStr) {
      sh.deleteRow(r + 1);
      return;
    }
  }
}

/***************
 * CSV helpers
 ***************/
function csvToSet_(csv) {
  const s = new Set();
  String(csv || "")
    .split(",")
    .map(x => x.trim())
    .filter(Boolean)
    .forEach(x => s.add(x));
  return s;
}

function setToCsv_(set) {
  return Array.from(set.values()).sort().join(",");
}

/**
 * Always keep one compile-visible function for debugging.
 */
function _sanityCheck() {
  Logger.log("Sanity check OK");
}
