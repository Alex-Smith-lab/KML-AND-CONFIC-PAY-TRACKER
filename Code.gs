/* ============================================================
   R.C PAY GOOGLE SHEETS BACKEND
   ============================================================ */

const CONFIG = {

  /*
     CHANGE THIS TOKEN.

     Use the same token in the R.C PAY web application.
  */
  TOKEN: "CHANGE_THIS_TO_YOUR_PRIVATE_TOKEN",

  /*
     If you leave SHEET_ID empty, the script will use the
     spreadsheet that this Apps Script is attached to.
  */
  SHEET_ID: "",

  REPORT_SHEET: "RC PAY REPORT",

  KML_SHEET: "KML BLOCKS",

  LOG_SHEET: "PAY LOG"

};


/* ============================================================
   GET
   ============================================================ */

function doGet(e) {

  return ContentService
    .createTextOutput(
      JSON.stringify({
        success: true,
        application: "R.C PAY",
        status: "online"
      })
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}


/* ============================================================
   POST
   ============================================================ */

function doPost(e) {

  try {

    const raw =
      e.postData &&
      e.postData.contents
        ? e.postData.contents
        : "{}";

    const data =
      JSON.parse(raw);

    if (
      CONFIG.TOKEN &&
      data.token !== CONFIG.TOKEN
    ) {

      return jsonResponse({
        success: false,
        error: "Invalid token"
      });

    }

    if (
      data.action === "test"
    ) {

      writeLog(
        "TEST",
        data.worker || ""
      );

      return jsonResponse({
        success: true,
        message: "R.C PAY connection successful"
      });

    }

    if (
      data.action === "saveReport"
    ) {

      saveReport(data);

      return jsonResponse({
        success: true,
        message: "Report saved"
      });

    }

    return jsonResponse({

      success: false,

      error:
        "Unknown action"

    });

  } catch (error) {

    return jsonResponse({

      success: false,

      error:
        String(error)

    });

  }

}


/* ============================================================
   SAVE REPORT
   ============================================================ */

function saveReport(data) {

  const spreadsheet =
    getSpreadsheet(
      data.sheetId
    );

  const worker =
    data.worker || {};

  const reportRows =
    data.rows || [];

  const kmlRows =
    data.kmlRows || [];

  const reportSheet =
    getOrCreateSheet(
      spreadsheet,
      CONFIG.REPORT_SHEET
    );

  const kmlSheet =
    getOrCreateSheet(
      spreadsheet,
      CONFIG.KML_SHEET
    );

  const now =
    new Date();

  /*
     REPORT HEADER
  */

  ensureReportHeader(
    reportSheet
  );

  /*
     FINAL PAY ROWS
  */

  reportRows.forEach(row => {

    reportSheet.appendRow([

      now,

      worker.name || "",

      worker.id || "",

      row.Source || "",

      row.Block || "",

      row.Location || "",

      row.Annotation || "",

      row.Quantity || 0,

      row.Rate || "",

      row.Assigned || "",

      row.Pay || 0

    ]);

  });


  /*
     KML BLOCK REPORT

     IMPORTANT:
     One row per block.

     Bus Stop, Bus Lane, Bike Lane,
     R0 and L0 stay on the SAME row.
  */

  ensureKmlHeader(
    kmlSheet
  );

  kmlRows.forEach(row => {

    kmlSheet.appendRow([

      now,

      worker.name || "",

      worker.id || "",

      row.Block || "",

      row.Location || "",

      row.BusStop || 0,

      row.BusLane || 0,

      row.BikeLane || 0,

      row.R0 || 0,

      row.L0 || 0,

      row.LocationDescription || 0,

      row.Additional || 0,

      row.AdditionalNames || "",

      row.BusStopPay || 0,

      row.BusLanePay || 0,

      row.BikeLanePay || 0,

      row.R0Pay || 0,

      row.L0Pay || 0,

      row.LocationPay || 0,

      row.AdditionalPay || 0,

      row.Pay || 0

    ]);

  });


  /*
     SUMMARY
  */

  writeSummary(
    spreadsheet,
    data,
    worker
  );


  /*
     LOG
  */

  writeLog(
    "REPORT SAVED",
    worker.name || ""
  );

}


/* ============================================================
   SUMMARY
   ============================================================ */

function writeSummary(
  spreadsheet,
  data,
  worker
) {

  const sheet =
    getOrCreateSheet(
      spreadsheet,
      CONFIG.PAY_LOG
    );

  if (
    sheet.getLastRow() === 0
  ) {

    sheet.appendRow([

      "Timestamp",
      "Worker",
      "Worker ID",
      "Annotation Pay",
      "KML Pay",
      "Final Pay"

    ]);

  }

  sheet.appendRow([

    new Date(),

    worker.name || "",

    worker.id || "",

    data.annotationTotal || 0,

    data.kmlTotal || 0,

    data.finalTotal || 0

  ]);

}


/* ============================================================
   REPORT HEADER
   ============================================================ */

function ensureReportHeader(sheet) {

  if (
    sheet.getLastRow() > 0
  ) {

    return;

  }

  sheet.appendRow([

    "Timestamp",
    "Worker",
    "Worker ID",
    "Source",
    "Block",
    "Location",
    "Annotation",
    "Quantity",
    "Rate",
    "Assigned",
    "Pay"

  ]);

}


/* ============================================================
   KML HEADER
   ============================================================ */

function ensureKmlHeader(sheet) {

  if (
    sheet.getLastRow() > 0
  ) {

    return;

  }

  sheet.appendRow([

    "Timestamp",
    "Worker",
    "Worker ID",
    "Block",
    "Location",

    "Bus Stop",
    "Bus Lane",
    "Bike Lane",
    "R0",
    "L0",

    "Location Description",
    "Additional",
    "Additional Names",

    "Bus Stop Pay",
    "Bus Lane Pay",
    "Bike Lane Pay",
    "R0 Pay",
    "L0 Pay",
    "Location Pay",
    "Additional Pay",

    "Block Total Pay"

  ]);

}


/* ============================================================
   LOG
   ============================================================ */

function writeLog(
  action,
  worker
) {

  const spreadsheet =
    getSpreadsheet("");

  const sheet =
    getOrCreateSheet(
      spreadsheet,
      CONFIG.PAY_LOG
    );

  if (
    sheet.getLastRow() === 0
  ) {

    sheet.appendRow([
      "Timestamp",
      "Action",
      "Worker"
    ]);

  }

  sheet.appendRow([

    new Date(),

    action,

    worker

  ]);

}


/* ============================================================
   SPREADSHEET
   ============================================================ */

function getSpreadsheet(
  suppliedId
) {

  const id =
    suppliedId ||
    CONFIG.SHEET_ID;

  if (id) {

    return SpreadsheetApp
      .openById(id);

  }

  return SpreadsheetApp
    .getActiveSpreadsheet();

}


/* ============================================================
   SHEET
   ============================================================ */

function getOrCreateSheet(
  spreadsheet,
  name
) {

  let sheet =
    spreadsheet.getSheetByName(name);

  if (!sheet) {

    sheet =
      spreadsheet.insertSheet(name);

  }

  return sheet;

}


/* ============================================================
   JSON
   ============================================================ */

function jsonResponse(data) {

  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}
