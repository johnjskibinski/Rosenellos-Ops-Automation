/*******************************
 * ROSENELLO SIDEBAR UI (Server)
 * (Diagnostics-wrapped)
 *
 * NOTE:
 * - This file expects Diagnostics.gs to exist in the SAME Apps Script project
 *   and provide safeExecute_(taskName, fn).
 *******************************/

/**
 * Assign THIS function to the small drawing/button on each tab.
 * It opens the sidebar.
 */
function openLpSidebar() {
  return safeExecute_("openLpSidebar", function () {
    SpreadsheetApp.getUi()
      .showSidebar(
        HtmlService.createTemplateFromFile("Sidebar")
          .evaluate()
          .setTitle("LP Upload / Print")
      );
    return true;
  });
}

/**
 * Optional: adds a top menu as a backup opener (desktop friendly).
 */
function onOpen(e) {
  return safeExecute_("onOpen", function () {
    SpreadsheetApp.getUi()
      .createMenu("Rosenello Tools")
      .addItem("Open LP Upload Sidebar", "openLpSidebar")
      .addSeparator()
      .addItem("Refresh LP (Gross D5 / Balance G5)", "refreshLpAmounts")
      .addToUi();
    return true;
  });
}

/**
 * Sidebar calls this to get current context + CONFIG-driven options.
 */
function getSidebarModel() {
  return safeExecute_("getSidebarModel", function () {
    var ss = SpreadsheetApp.getActive();
    var activeSheet = ss.getActiveSheet();

    var cfg = readConfig_();

    return {
      spreadsheetId: ss.getId(),
      activeTabName: activeSheet.getName(),
      tabs: cfg.tabs, // array of tab names
      landscapeTabs: cfg.landscapeTabs, // array
      portraitTabs: cfg.portraitTabs, // array
      docTypeByTab: cfg.docTypeByTab, // map tabName -> docTypeId (string/number)
      cellJobId: cfg.cellJobId || "C1",
      cellGross: cfg.cellGross || "D5",
      cellBalance: cfg.cellBalance || "G5",
    };
  });
}

/**
 * Refresh GrossAmount (D5) and BalanceDue (G5) from LP for the JobID in C1.
 * Called from sidebar button + menu item.
 */
function refreshLpAmounts() {
  return safeExecute_("refreshLpAmounts", function () {
    var ss = SpreadsheetApp.getActive();
    var cfg = readConfig_();

    var jobId = String(ss.getRange(cfg.cellJobId).getValue() || "").trim();
    if (!jobId) throw new Error("No JobID found in " + cfg.cellJobId + ".");

    // TODO: Implement token + SalesApi/GetSalesJobDetail using your LP API credentials.
    // For now, this is a stub that throws a clear message until you add credentials.
    var detail = lpGetSalesJobDetail_(jobId); // { grossAmount: number, balanceDue: number }

    // Write values
    ss.getRange(cfg.cellGross).setValue(detail.grossAmount);
    ss.getRange(cfg.cellBalance).setValue(detail.balanceDue);

    return {
      ok: true,
      jobId: jobId,
      grossAmount: detail.grossAmount,
      balanceDue: detail.balanceDue,
      wrote: { grossCell: cfg.cellGross, balanceCell: cfg.cellBalance },
    };
  });
}

/**
 * Export + upload ONLY the currently active tab.
 * Called from sidebar button.
 */
function uploadCurrentTab() {
  return safeExecute_("uploadCurrentTab", function () {
    var ss = SpreadsheetApp.getActive();
    var activeTab = ss.getActiveSheet().getName();
    return uploadTabByName_(activeTab);
  });
}

/**
 * Export + upload ALL configured tabs (the 5 packet tabs).
 * Called from sidebar button.
 */
function uploadAllTabs() {
  return safeExecute_("uploadAllTabs", function () {
    var cfg = readConfig_();
    var results = [];
    for (var i = 0; i < cfg.tabs.length; i++) {
      var tabName = cfg.tabs[i];
      results.push(uploadTabByName_(tabName));
    }
    return { ok: true, results: results };
  });
}

/**
 * Core: export one tab to PDF with correct orientation, then upload to LP with correct docType.
 */
function uploadTabByName_(tabName) {
  // Internal helper; still wrapped so we get per-tab diagnostics too.
  return safeExecute_("uploadTabByName_:" + tabName, function () {
    var ss = SpreadsheetApp.getActive();
    var cfg = readConfig_();

    var jobId = String(ss.getRange(cfg.cellJobId).getValue() || "").trim();
    if (!jobId) {
      throw new Error("No JobID found in " + cfg.cellJobId + ". Enter it, then retry.");
    }

    var docType = cfg.docTypeByTab[tabName];
    if (!docType) {
      throw new Error('No LP doc type mapped for tab "' + tabName + '" in CONFIG.');
    }

    var orientation = cfg.landscapeTabs.indexOf(tabName) !== -1 ? "landscape" : "portrait";

    var pdfBlob = exportSheetTabToPdf_(ss, tabName, orientation);

    // NOTE: This is where the LP upload happens. Stubbed until LP method is finalized.
    var uploadInfo = lpUploadDocument_(jobId, docType, pdfBlob, tabName, orientation);

    return {
      ok: true,
      tabName: tabName,
      jobId: jobId,
      docType: docType,
      orientation: orientation,
      uploadInfo: uploadInfo,
    };
  });
}

/*******************************
 * PDF EXPORT (Google Sheets)
 *******************************/

/**
 * Exports a single tab as a 1-page PDF (fit to width/page) using a standard Sheets export URL.
 * We’ll tune exact parameters once we mirror your sample PDFs 1:1.
 */
function exportSheetTabToPdf_(ss, tabName, orientation) {
  // Internal helper; still wrapped so export failures are clearly logged.
  return safeExecute_("exportSheetTabToPdf_:" + tabName, function () {
    var sh = ss.getSheetByName(tabName);
    if (!sh) throw new Error("Tab not found: " + tabName);

    var gid = sh.getSheetId();
    var base = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export";

    // Basic print params. We will refine these to match your “Print to PDF” exactly.
    var params = {
      format: "pdf",
      gid: gid,
      portrait: orientation === "portrait" ? "true" : "false",
      fitw: "true", // fit to width
      sheetnames: "false",
      printtitle: "false",
      pagenumbers: "false",
      gridlines: "false",
      fzr: "false", // repeat frozen rows
      size: "letter",
      top_margin: "0.50",
      bottom_margin: "0.50",
      left_margin: "0.50",
      right_margin: "0.50",
    };

    var query = Object.keys(params)
      .map(function (k) {
        return encodeURIComponent(k) + "=" + encodeURIComponent(params[k]);
      })
      .join("&");

    var url = base + "?" + query;

    var token = ScriptApp.getOAuthToken();
    var resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true,
    });

    if (resp.getResponseCode() !== 200) {
      throw new Error(
        'PDF export failed for "' +
          tabName +
          '" (HTTP ' +
          resp.getResponseCode() +
          "): " +
          resp.getContentText()
      );
    }

    var blob = resp.getBlob().setName(sanitizeFilePart_(tabName) + ".pdf");
    return blob;
  });
}

function sanitizeFilePart_(s) {
  return safeExecute_("sanitizeFilePart_", function () {
    return String(s).replace(/[^\w\-]+/g, "_").substring(0, 80);
  });
}

/*******************************
 * CONFIG READER
 *******************************/

function readConfig_() {
  // Internal helper; wrapped so missing CONFIG errors are logged.
  return safeExecute_("readConfig_", function () {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName("CONFIG");
    if (!sh) throw new Error("CONFIG sheet not found. Run setupMeasureTemplateConfig first.");

    var values = sh.getDataRange().getValues(); // [Key, Value, Notes]
    var map = {};
    for (var r = 1; r < values.length; r++) {
      var k = String(values[r][0] || "").trim();
      var v = String(values[r][1] || "").trim();
      if (k) map[k] = v;
    }

    var tabs = [
      map["TAB_Costing"],
      map["TAB_LaborCalc"],
      map["TAB_WorkOrder"],
      map["TAB_WindowMeasure"],
      map["TAB_Checklist"],
    ].filter(function (x) {
      return !!x;
    });

    var landscapeTabs = splitCsv_(map["PRINT_LANDSCAPE_TABS"]);
    var portraitTabs = splitCsv_(map["PRINT_PORTRAIT_TABS"]);

    // doc type keys in CONFIG are stored as: LP_DOC_TYPE_<TabName>
    var docTypeByTab = {};
    for (var i = 0; i < tabs.length; i++) {
      var tab = tabs[i];
      var key = "LP_DOC_TYPE_" + tab;
      if (map[key]) docTypeByTab[tab] = map[key];
    }

    return {
      cellJobId: map["CELL_JOB_ID"] || "C1",
      cellGross: map["CELL_GROSS_AMOUNT"] || "D5",
      cellBalance: map["CELL_BALANCE_DUE"] || "G5",
      tabs: tabs,
      landscapeTabs: landscapeTabs,
      portraitTabs: portraitTabs,
      docTypeByTab: docTypeByTab,
    };
  });
}

function splitCsv_(s) {
  return safeExecute_("splitCsv_", function () {
    return String(s || "")
      .split(",")
      .map(function (x) {
        return String(x || "").trim();
      })
      .filter(function (x) {
        return !!x;
      });
  });
}

/*******************************
 * LP STUBS (intentionally)
 *******************************/

/**
 * Placeholder: LP SalesApi/GetSalesJobDetail
 * You already confirmed we’ll write:
 *  - GrossAmount -> D5
 *  - BalanceDue  -> G5
 */
function lpGetSalesJobDetail_(jobId) {
  // Keep this wrapped so when you test refresh it logs cleanly.
  return safeExecute_("lpGetSalesJobDetail_:" + jobId, function () {
    // TODO: implement with your LP API token flow + POST to /api/SalesApi/GetSalesJobDetail
    // using Script Properties:
    // LP_API_BASE_URL, LP_API_USERNAME, LP_API_PASSWORD, LP_CLIENTID, LP_APPKEY (if required)
    throw new Error(
      "LP API not configured yet (lpGetSalesJobDetail_). Add credentials in Script Properties, then we’ll wire this."
    );
  });
}

/**
 * Placeholder: upload a PDF to LP under a docType for a jobId.
 * This is where we will plug in the exact “bulletproof” upload method.
 */
function lpUploadDocument_(jobId, docType, pdfBlob, tabName, orientation) {
  return safeExecute_("lpUploadDocument_:" + jobId + ":" + tabName, function () {
    // TODO: implement actual upload
    // IMPORTANT: we will use the method you already mapped for Rosenello’s tenant,
    // and keep tokens/cookies out of the sheet.
    throw new Error(
      "LP upload not configured yet (lpUploadDocument_). Next step is wiring your tenant’s upload method."
    );
  });
}

/**
 * Sanity function (always visible if script compiles).
 */
function _sanityCheck() {
  return safeExecute_("_sanityCheck", function () {
    Logger.log("Sanity check OK");
    return true;
  });
}
