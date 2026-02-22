/*******************************
 * ROSENELLO SIDEBAR UI (Server)
 *******************************/

/**
 * Assign THIS function to the small drawing/button on each tab.
 * It opens the sidebar.
 */
function openLpSidebar() {
  SpreadsheetApp.getUi()
    .showSidebar(HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('LP Upload / Print'));
}

/**
 * Optional: adds a top menu as a backup opener (desktop friendly).
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Rosenello Tools')
    .addItem('Open LP Upload Sidebar', 'openLpSidebar')
    .addSeparator()
    .addItem('Refresh LP (Gross D5 / Balance G5)', 'refreshLpAmounts')
    .addToUi();
}

/**
 * Sidebar calls this to get current context + CONFIG-driven options.
 */
function getSidebarModel() {
  const ss = SpreadsheetApp.getActive();
  const activeSheet = ss.getActiveSheet();

  const cfg = readConfig_();

  return {
    spreadsheetId: ss.getId(),
    activeTabName: activeSheet.getName(),
    tabs: cfg.tabs, // array of tab names
    landscapeTabs: cfg.landscapeTabs, // array
    portraitTabs: cfg.portraitTabs, // array
    docTypeByTab: cfg.docTypeByTab, // map tabName -> docTypeId (string/number)
    cellJobId: cfg.cellJobId || 'C1',
    cellGross: cfg.cellGross || 'D5',
    cellBalance: cfg.cellBalance || 'G5',
  };
}

/**
 * Refresh GrossAmount (D5) and BalanceDue (G5) from LP for the JobID in C1.
 * Called from sidebar button + menu item.
 */
function refreshLpAmounts() {
  const ss = SpreadsheetApp.getActive();
  const cfg = readConfig_();

  const jobId = String(ss.getRange(cfg.cellJobId).getValue() || '').trim();
  if (!jobId) throw new Error(`No JobID found in ${cfg.cellJobId}.`);

  // TODO: Implement token + SalesApi/GetSalesJobDetail using your LP API credentials.
  // For now, this is a stub that throws a clear message until you add credentials.
  const detail = lpGetSalesJobDetail_(jobId); // { grossAmount: number, balanceDue: number }

  // Write values
  ss.getRange(cfg.cellGross).setValue(detail.grossAmount);
  ss.getRange(cfg.cellBalance).setValue(detail.balanceDue);

  return {
    ok: true,
    jobId,
    grossAmount: detail.grossAmount,
    balanceDue: detail.balanceDue,
    wrote: { grossCell: cfg.cellGross, balanceCell: cfg.cellBalance }
  };
}

/**
 * Export + upload ONLY the currently active tab.
 * Called from sidebar button.
 */
function uploadCurrentTab() {
  const ss = SpreadsheetApp.getActive();
  const activeTab = ss.getActiveSheet().getName();
  return uploadTabByName_(activeTab);
}

/**
 * Export + upload ALL configured tabs (the 5 packet tabs).
 * Called from sidebar button.
 */
function uploadAllTabs() {
  const cfg = readConfig_();
  const results = [];
  for (const tabName of cfg.tabs) {
    results.push(uploadTabByName_(tabName));
  }
  return { ok: true, results };
}

/**
 * Core: export one tab to PDF with correct orientation, then upload to LP with correct docType.
 */
function uploadTabByName_(tabName) {
  const ss = SpreadsheetApp.getActive();
  const cfg = readConfig_();

  const jobId = String(ss.getRange(cfg.cellJobId).getValue() || '').trim();
  if (!jobId) throw new Error(`No JobID found in ${cfg.cellJobId}. Enter it, then retry.`);

  const docType = cfg.docTypeByTab[tabName];
  if (!docType) throw new Error(`No LP doc type mapped for tab "${tabName}" in CONFIG.`);

  const orientation = cfg.landscapeTabs.includes(tabName) ? 'landscape' : 'portrait';

  const pdfBlob = exportSheetTabToPdf_(ss, tabName, orientation);

  // NOTE: This is where the LP upload happens. Stubbed until LP method is finalized.
  const uploadInfo = lpUploadDocument_(jobId, docType, pdfBlob, tabName, orientation);

  return {
    ok: true,
    tabName,
    jobId,
    docType,
    orientation,
    uploadInfo
  };
}

/*******************************
 * PDF EXPORT (Google Sheets)
 *******************************/

/**
 * Exports a single tab as a 1-page PDF (fit to width/page) using a standard Sheets export URL.
 * We’ll tune exact parameters once we mirror your sample PDFs 1:1.
 */
function exportSheetTabToPdf_(ss, tabName, orientation) {
  const sh = ss.getSheetByName(tabName);
  if (!sh) throw new Error(`Tab not found: ${tabName}`);

  const gid = sh.getSheetId();
  const base = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export`;

  // Basic print params. We will refine these to match your “Print to PDF” exactly.
  const params = {
    format: 'pdf',
    gid: gid,
    portrait: (orientation === 'portrait') ? 'true' : 'false',
    fitw: 'true',          // fit to width
    sheetnames: 'false',
    printtitle: 'false',
    pagenumbers: 'false',
    gridlines: 'false',
    fzr: 'false',          // repeat frozen rows
    size: 'letter',
    top_margin: '0.50',
    bottom_margin: '0.50',
    left_margin: '0.50',
    right_margin: '0.50'
  };

  const query = Object.keys(params)
    .map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`)
    .join('&');

  const url = `${base}?${query}`;

  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error(`PDF export failed for "${tabName}" (HTTP ${resp.getResponseCode()}): ${resp.getContentText()}`);
  }

  const blob = resp.getBlob().setName(`${sanitizeFilePart_(tabName)}.pdf`);
  return blob;
}

function sanitizeFilePart_(s) {
  return String(s).replace(/[^\w\-]+/g, '_').substring(0, 80);
}

/*******************************
 * CONFIG READER
 *******************************/

function readConfig_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('CONFIG');
  if (!sh) throw new Error('CONFIG sheet not found. Run setupMeasureTemplateConfig_ first.');

  const values = sh.getDataRange().getValues(); // [Key, Value, Notes]
  const map = {};
  for (let r = 1; r < values.length; r++) {
    const k = String(values[r][0] || '').trim();
    const v = String(values[r][1] || '').trim();
    if (k) map[k] = v;
  }

  const tabs = [
    map['TAB_Costing'],
    map['TAB_LaborCalc'],
    map['TAB_WorkOrder'],
    map['TAB_WindowMeasure'],
    map['TAB_Checklist'],
  ].filter(Boolean);

  const landscapeTabs = splitCsv_(map['PRINT_LANDSCAPE_TABS']);
  const portraitTabs = splitCsv_(map['PRINT_PORTRAIT_TABS']);

  // doc type keys in CONFIG are stored as: LP_DOC_TYPE_<TabName>
  const docTypeByTab = {};
  for (const tab of tabs) {
    const key = `LP_DOC_TYPE_${tab}`;
    if (map[key]) docTypeByTab[tab] = map[key];
  }

  return {
    cellJobId: map['CELL_JOB_ID'] || 'C1',
    cellGross: map['CELL_GROSS_AMOUNT'] || 'D5',
    cellBalance: map['CELL_BALANCE_DUE'] || 'G5',
    tabs,
    landscapeTabs,
    portraitTabs,
    docTypeByTab,
  };
}

function splitCsv_(s) {
  return String(s || '')
    .split(',')
    .map(x => x.trim())
    .filter(Boolean);
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
  // TODO: implement with your LP API token flow + POST to /api/SalesApi/GetSalesJobDetail
  // using Script Properties:
  // LP_API_BASE_URL, LP_API_USERNAME, LP_API_PASSWORD, LP_CLIENTID, LP_APPKEY (if required)
  throw new Error('LP API not configured yet (lpGetSalesJobDetail_). Add credentials in Script Properties, then we’ll wire this.');
}

/**
 * Placeholder: upload a PDF to LP under a docType for a jobId.
 * This is where we will plug in the exact “bulletproof” upload method.
 */
function lpUploadDocument_(jobId, docType, pdfBlob, tabName, orientation) {
  // TODO: implement actual upload
  // IMPORTANT: we will use the method you already mapped for Rosenello’s tenant,
  // and keep tokens/cookies out of the sheet.
  throw new Error('LP upload not configured yet (lpUploadDocument_). Next step is wiring your tenant’s upload method.');
}

/**
 * Sanity function (always visible if script compiles).
 */
function _sanityCheck() {
  Logger.log('Sanity check OK');
  return true;
}
