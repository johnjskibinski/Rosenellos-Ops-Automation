/***************
 * ROSENELLO SIDEBAR UI (REAL)
 * - Reads CONFIG dynamically
 * - UI is live now
 * - Upload/LP calls are stubs for safety (we wire API next)
 *
 * DIAGNOSTICS:
 * - This file expects Diagnostics.gs to exist in the SAME Apps Script project
 *   and provide safeExecute_(taskName, fn).
 ***************/

function onOpen(e) {
  return safeExecute_("SidebarUI.onOpen", function () {
    SpreadsheetApp.getUi()
      .createMenu("Rosenello Tools")
      .addItem("Open LP Upload Sidebar", "openLpSidebar")
      .addToUi();
    return true;
  });
}

/**
 * Assign THIS function to your Drawing button on each tab.
 */
function openLpSidebar() {
  return safeExecute_("SidebarUI.openLpSidebar", function () {
    SpreadsheetApp.getUi().showSidebar(
      HtmlService.createHtmlOutputFromFile("Sidebar").setTitle("LP Upload / Print")
    );
    return true;
  });
}

/**
 * Sidebar model: CONFIG + tab list + mapping list.
 */
function getSidebarModel() {
  return safeExecute_("SidebarUI.getSidebarModel", function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getConfigSheet_(ss);

    var config = readConfig_(sh);

    var activeTabName = ss.getActiveSheet().getName();
    var sheetTabs = ss.getSheets().map(function (s) {
      return s.getName();
    });

    // Build mappings dynamically:
    // We take all keys like LP_DOC_TYPE_{TABNAME} and map them.
    var mappings = [];

    var landscapeSet = csvToSet_(config.PRINT_LANDSCAPE_TABS || "");
    var portraitSet = csvToSet_(config.PRINT_PORTRAIT_TABS || "");

    Object.keys(config).forEach(function (k) {
      if (k.indexOf("LP_DOC_TYPE_") === 0) {
        var tab = k.substring("LP_DOC_TYPE_".length);
        var docTypeId = String(config[k] || "").trim();
        if (!docTypeId) return;

        var orientation = "";
        if (landscapeSet.has(tab)) orientation = "LANDSCAPE";
        else if (portraitSet.has(tab)) orientation = "PORTRAIT";

        mappings.push({ tab: tab, docTypeId: docTypeId, orientation: orientation });
      }
    });

    return {
      configSheetName: sh.getName(),
      lpBaseUrl: config.LP_BASE_URL || "", // configurable later
      activeTabName: activeTabName,
      sheetTabs: sheetTabs,
      mappings: mappings,
    };
  });
}

/**
 * Save/update a mapping:
 * - sets LP_DOC_TYPE_{TAB} = id
 * - updates PRINT_LANDSCAPE_TABS / PRINT_PORTRAIT_TABS to include TAB
 */
function saveTabMapping(payload) {
  return safeExecute_("SidebarUI.saveTabMapping", function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getConfigSheet_(ss);
    var cfg = readConfig_(sh);

    var tab = String((payload && payload.tab) || "").trim();
    var docTypeId = String((payload && payload.docTypeId) || "").trim();
    var orientation = String((payload && payload.orientation) || "").trim().toUpperCase();
    var notes = String((payload && payload.notes) || "").trim();

    if (!tab) throw new Error("Tab is required.");
    if (!/^\d+$/.test(docTypeId)) throw new Error("DocTypeId must be numeric.");
    if (orientation !== "LANDSCAPE" && orientation !== "PORTRAIT") {
      throw new Error("Orientation must be LANDSCAPE or PORTRAIT.");
    }

    // Write LP_DOC_TYPE_{TAB}
    upsertConfigRow_(sh, "LP_DOC_TYPE_" + tab, docTypeId, notes || "LP document type ID");

    // Update orientation CSV lists
    var landscape = csvToSet_(cfg.PRINT_LANDSCAPE_TABS || "");
    var portrait = csvToSet_(cfg.PRINT_PORTRAIT_TABS || "");

    // remove from both first
    landscape.delete(tab);
    portrait.delete(tab);

    if (orientation === "LANDSCAPE") landscape.add(tab);
    if (orientation === "PORTRAIT") portrait.add(tab);

    upsertConfigRow_(sh, "PRINT_LANDSCAPE_TABS", setToCsv_(landscape), "Export as PDF landscape");
    upsertConfigRow_(sh, "PRINT_PORTRAIT_TABS", setToCsv_(portrait), "Export as PDF portrait");

    return { ok: true, message: "Saved mapping: " + tab + " → DocType " + docTypeId + " (" + orientation + ")." };
  });
}

/**
 * Remove a mapping:
 * - deletes LP_DOC_TYPE_{TAB}
 * - removes TAB from both orientation lists
 */
function removeTabMapping(tabName) {
  return safeExecute_("SidebarUI.removeTabMapping", function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getConfigSheet_(ss);
    var cfg = readConfig_(sh);

    var tab = String(tabName || "").trim();
    if (!tab) throw new Error("Tab is required.");

    deleteConfigKey_(sh, "LP_DOC_TYPE_" + tab);

    var landscape = csvToSet_(cfg.PRINT_LANDSCAPE_TABS || "");
    var portrait = csvToSet_(cfg.PRINT_PORTRAIT_TABS || "");
    landscape.delete(tab);
    portrait.delete(tab);

    upsertConfigRow_(sh, "PRINT_LANDSCAPE_TABS", setToCsv_(landscape), "Export as PDF landscape");
    upsertConfigRow_(sh, "PRINT_PORTRAIT_TABS", setToCsv_(portrait), "Export as PDF portrait");

    return { ok: true, message: "Removed mapping for: " + tab };
  });
}

/***************
 * STUB ACTIONS (safe placeholders)
 * We wire these to LP API upload + token + PDF export next.
 ***************/

function refreshLpAmounts() {
  return safeExecute_("SidebarUI.refreshLpAmounts_stub", function () {
    // Later: use JobID from C1, call SalesApi/GetSalesJobDetail, write GrossAmount to D5 and BalanceDue to G5.
    return {
      ok: true,
      message:
        "Refresh stub ✅\n" +
        "Next wiring step:\n" +
        "• Read JobID from CONFIG CELL_JOB_ID (C1)\n" +
        "• Call LP SalesApi/GetSalesJobDetail\n" +
        "• Write GrossAmount → D5, BalanceDue → G5",
    };
  });
}

function uploadCurrentTab() {
  return safeExecute_("SidebarUI.uploadCurrentTab_stub", function () {
    var tab = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    return uploadSingleTab(tab);
  });
}

function uploadAllTabs() {
  return safeExecute_("SidebarUI.uploadAllTabs_stub", function () {
    // Later: iterate mappings in CONFIG in your desired order (Costing, Window Measure, Work Order, LaborCalc, Checklist)
    return {
      ok: true,
      message:
        "Upload FULL packet stub ✅\n" +
        "Next wiring step:\n" +
        "• Read mappings LP_DOC_TYPE_* from CONFIG\n" +
        "• Export each mapped tab to 1-page PDF with correct orientation\n" +
        "• Upload each PDF using LP API UploadDocument (Option A)",
    };
  });
}

function uploadSingleTab(tabName) {
  return safeExecute_("SidebarUI.uploadSingleTab_stub:" + String(tabName || ""), function () {
    // Later: export only this tab and upload with its doc type.
    return {
      ok: true,
      message:
        "Upload stub ✅\nTab: " +
        tabName +
        "\n" +
        "Next wiring step:\n" +
        "• Find LP_DOC_TYPE_{TAB} in CONFIG\n" +
        "• Export tab → PDF (portrait/landscape from CONFIG)\n" +
        "• Upload via LP API UploadDocument (NOT web endpoint)",
    };
  });
}

/***************
 * CONFIG HELPERS
 ***************/

function getConfigSheet_(ss) {
  return safeExecute_("SidebarUI.getConfigSheet_", function () {
    var sh = ss.getSheetByName("CONFIG");
    if (!sh) throw new Error("CONFIG sheet not found. Run your setupMeasureTemplateConfig_() first.");
    return sh;
  });
}

function readConfig_(sh) {
  return safeExecute_("SidebarUI.readConfig_", function () {
    var values = sh.getDataRange().getValues();
    // Expect header row: Key | Value | Notes
    var out = {};
    for (var r = 1; r < values.length; r++) {
      var key = String(values[r][0] || "").trim();
      var val = values[r][1];
      if (!key) continue;
      out[key] = val === null || typeof val === "undefined" ? "" : String(val);
    }
    return out;
  });
}

/**
 * Upsert key/value/notes into CONFIG.
 */
function upsertConfigRow_(sh, key, value, notes) {
  return safeExecute_("SidebarUI.upsertConfigRow_:" + String(key || ""), function () {
    var data = sh.getDataRange().getValues();
    var keyStr = String(key || "").trim();
    if (!keyStr) throw new Error("Config key required.");

    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0] || "").trim() === keyStr) {
        sh.getRange(r + 1, 2).setValue(value);
        if (typeof notes !== "undefined") sh.getRange(r + 1, 3).setValue(notes);
        return true;
      }
    }

    // append
    var newRow = sh.getLastRow() + 1;
    sh.getRange(newRow, 1, 1, 3).setValues([[keyStr, value, notes || ""]]);
    return true;
  });
}

function deleteConfigKey_(sh, key) {
  return safeExecute_("SidebarUI.deleteConfigKey_:" + String(key || ""), function () {
    var data = sh.getDataRange().getValues();
    var keyStr = String(key || "").trim();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0] || "").trim() === keyStr) {
        sh.deleteRow(r + 1);
        return true;
      }
    }
    return false;
  });
}

/***************
 * CSV helpers
 ***************/
function csvToSet_(csv) {
  return safeExecute_("SidebarUI.csvToSet_", function () {
    var s = new Set();
    String(csv || "")
      .split(",")
      .map(function (x) {
        return String(x || "").trim();
      })
      .filter(function (x) {
        return !!x;
      })
      .forEach(function (x) {
        s.add(x);
      });
    return s;
  });
}

function setToCsv_(set) {
  return safeExecute_("SidebarUI.setToCsv_", function () {
    return Array.from(set.values()).sort().join(",");
  });
}

/**
 * Always keep one compile-visible function for debugging.
 */
function _sanityCheck() {
  return safeExecute_("SidebarUI._sanityCheck", function () {
    Logger.log("Sanity check OK");
    return true;
  });
}
