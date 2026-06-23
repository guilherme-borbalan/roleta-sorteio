// ============================================================
// COLE ESTE CODIGO NO GOOGLE APPS SCRIPT (script.google.com)
// ============================================================

function doGet(e) {
  var sheetId = e.parameter.id;
  var sheetName = e.parameter.sheet || "Planilha1";
  var col = e.parameter.col || "A";

  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return jsonResponse({ error: "Aba '" + sheetName + "' nao encontrada." });
    }

    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      return jsonResponse({ error: "Planilha vazia." });
    }

    var range = sheet.getRange(col + "1:" + col + lastRow);
    var values = range.getValues();

    var names = [];
    for (var i = 0; i < values.length; i++) {
      var val = String(values[i][0]).trim();
      if (val !== "" && val !== "undefined" && val !== "null") {
        names.push(val);
      }
    }

    return jsonResponse({ names: names });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
