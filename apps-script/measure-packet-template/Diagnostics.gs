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
    throw err;
  }
}

function logSuccess_(taskName, startTime) {
  console.log(
    JSON.stringify({
      type: "SUCCESS",
      task: taskName,
      durationMs: new Date() - startTime,
      timestamp: new Date().toISOString(),
    })
  );
}

function logError_(taskName, err, startTime) {
  const errorObj = {
    type: "ERROR",
    task: taskName,
    message: err && err.message ? err.message : String(err),
    stack: err && err.stack ? err.stack : "",
    durationMs: new Date() - startTime,
    timestamp: new Date().toISOString(),
  };

  console.error(JSON.stringify(errorObj));
}

function diagnosticsSanity() {
  return safeExecute_("diagnosticsSanity", function () {
    Logger.log("Diagnostics wrapper is working");
  });
}
