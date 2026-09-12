/*
============================================================
R.C PAY - GOOGLE SHEETS BACKEND
============================================================

This receives reports from the R.C PAY web application
and appends them to a Google Sheet.

IMPORTANT:
Change SECRET below to your own private secret.
Do NOT use the example value in production.
============================================================
*/


const CONFIG = {

  SECRET: "CHANGE_THIS_TO_YOUR_PRIVATE_SECRET",

  DEFAULT_SHEET: "RC_PAY_REPORTS"

};


/* ============================================================
   GET
   ============================================================ */

function doGet(e) {

  return jsonResponse({

    ok: true,

    service: "R.C PAY",

    message:
      "R.C PAY Google Sheets integration is running."

  });

}


/* ============================================================
   POST
   ============================================================ */

function doPost(e) {

  try {

    if (
      !e ||
      !e.postData ||
      !e.postData.contents
    ) {

      return jsonResponse({

        ok: false,

        error:
          "No POST data received."

      });

    }


    const body =
      JSON.parse(
        e.postData.contents
      );


    /* --------------------------------------------------------
       SECURITY CHECK
       -------------------------------------------------------- */

    if (
      !body.secret ||
      body.secret !== CONFIG.SECRET
    ) {

      return jsonResponse({

        ok: false,

        error:
          "Unauthorized."

      });

    }


    /* --------------------------------------------------------
       SPREADSHEET
       -------------------------------------------------------- */

    const spreadsheetId =
      body.spreadsheetId;


    if (!spreadsheetId) {

      return jsonResponse({

        ok: false,

        error:
          "Spreadsheet ID is required."

      });

    }


    const spreadsheet =
      SpreadsheetApp.openById(
        spreadsheetId
      );


    const sheetName =
      body.sheetName ||
      CONFIG.DEFAULT_SHEET;


    let sheet =
      spreadsheet.getSheetByName(
        sheetName
      );


    if (!sheet) {

      sheet =
        spreadsheet.insertSheet(
          sheetName
        );

    }


    /* --------------------------------------------------------
       ROWS
       -------------------------------------------------------- */

    const rows =
      body.rows;


    if (
      !Array.isArray(rows) ||
      rows.length === 0
    ) {

      return jsonResponse({

        ok: false,

        error:
          "No report rows supplied."

      });

    }


    /* --------------------------------------------------------
       CREATE HEADERS
       -------------------------------------------------------- */

    const existingLastColumn =
      sheet.getLastColumn();


    const existingLastRow =
      sheet.getLastRow();


    let headers = [];


    if (
      existingLastRow > 0 &&
      existingLastColumn > 0
    ) {

      headers =
        sheet
          .getRange(
            1,
            1,
            1,
            existingLastColumn
          )
          .getValues()[0]
          .map(
            value =>
              String(value)
          )
          .filter(
            value =>
              value.trim() !== ""
          );

    }


    /* --------------------------------------------------------
       ADD NEW HEADERS
       -------------------------------------------------------- */

    const incomingHeaders = [];


    rows.forEach(row => {

      Object.keys(row).forEach(key => {

        if (
          !incomingHeaders.includes(key)
        ) {

          incomingHeaders.push(key);

        }

      });

    });


    incomingHeaders.forEach(header => {

      if (
        !headers.includes(header)
      ) {

        headers.push(header);

      }

    });


    if (headers.length === 0) {

      return jsonResponse({

        ok: false,

        error:
          "No columns found."

      });

    }


    /* --------------------------------------------------------
       WRITE HEADERS
       -------------------------------------------------------- */

    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .setValues([
        headers
      ]);


    /* --------------------------------------------------------
       PREPARE VALUES
       -------------------------------------------------------- */

    const values =
      rows.map(row => {

        return headers.map(
          header => {

            const value =
              row[header];

            if (
              value === null ||
              value === undefined
            ) {

              return "";

            }

            return value;

          }
        );

      });


    /* --------------------------------------------------------
       APPEND
       -------------------------------------------------------- */

    const startRow =
      Math.max(
        sheet.getLastRow() + 1,
        2
      );


    sheet
      .getRange(
        startRow,
        1,
        values.length,
        headers.length
      )
      .setValues(values);


    /* --------------------------------------------------------
       FORMATTING
       -------------------------------------------------------- */

    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .setFontWeight("bold");


    sheet.setFrozenRows(1);


    return jsonResponse({

      ok: true,

      count:
        values.length,

      sheet:
        sheetName,

      message:
        "R.C PAY report saved successfully."

    });


  } catch (error) {

    return jsonResponse({

      ok: false,

      error:
        String(error)

    });

  }

}


/* ============================================================
   JSON RESPONSE
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
