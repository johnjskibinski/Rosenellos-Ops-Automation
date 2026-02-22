function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create/ensure Tracker tab exists + headers
  let sh = ss.getSheetByName("Tracker");
  if (!sh) sh = ss.insertSheet("Tracker");

  if (sh.getLastRow() === 0) {
    sh.appendRow(["eventId","eventStart","eventTitle","address","phone","packetUrl","status","createdAt","notes"]);
  }

  // Remove existing sync trigger(s) to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "syncmeasure" || t.getHandlerFunction() === "syncMeasures") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create a new time trigger
  ScriptApp.newTrigger("syncmeasure")  // <-- matches what YOU see in dropdown
    .timeBased()
    .everyMinutes(5)
    .create();
}
