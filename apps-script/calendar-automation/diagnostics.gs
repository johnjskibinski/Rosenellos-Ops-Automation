/*********************************************************
 DIAGNOSTICS + SAFE EXECUTION WRAPPER
 Shared across ALL automation scripts
**********************************************************/

function safeExecute_(taskName, fn) {

  const startTime = new Date();

  try {

    const result = fn();

    logSuccess_(taskName, startTime);

    return result;

  } catch (err) {

    logError_(taskName, err, startTime);

    throw err; // still throw so execution shows failed if needed
  }
}


/**********************
 SUCCESS LOGGER
**********************/
function logSuccess_(taskName, startTime) {

  console.log(JSON.stringify({
    type: "SUCCESS",
    task: taskName,
    durationMs: new Date() - startTime,
    timestamp: new Date().toISOString()
  }));
}


/**********************
 ERROR LOGGER
**********************/
function logError_(taskName, err, startTime) {

  const errorObj = {
    type: "ERROR",
    task: taskName,
    message: err.message || err,
    stack: err.stack || "",
    durationMs: new Date() - startTime,
    timestamp: new Date().toISOString()
  };

  console.error(JSON.stringify(errorObj));

  // FUTURE: add alerting logic here
}

function diagnosticsSanity() { Logger.log("ok"); }

function diagnosticsListGlobals() {
  Logger.log(Object.keys(this).filter(k => k.includes("safeExecute") || k.includes("logSuccess") || k.includes("logError")));
}
