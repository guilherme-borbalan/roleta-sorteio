// ============================================================
// COLE ESTE CODIGO NO GOOGLE APPS SCRIPT (script.google.com)
// ============================================================
// IMPORTANTE: Apos colar, va em Implantar > Gerenciar implantacoes
//   Clique no lapis > Versao: "Nova versao" > Implantar
// ============================================================

function doGet(e) {
  var params = e.parameter || {};
  var sheetId = params.id || "";
  var sheetName = params.sheet || "Planilha1";
  var col = params.col || "A";
  var callback = params.callback || "";

  if (!sheetId) {
    return sendResponse(callback, { error: "ID nao informado." });
  }

  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return sendResponse(callback, { error: "Aba '" + sheetName + "' nao encontrada." });
    }

    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      return sendResponse(callback, { error: "Planilha vazia." });
    }

    var values = sheet.getRange(col + "1:" + col + lastRow).getValues();
    var names = [];

    for (var i = 0; i < values.length; i++) {
      var val = String(values[i][0]).trim();
      if (val !== "" && val !== "undefined" && val !== "null") {
        names.push(val);
      }
    }

    if (names.length === 0) {
      return sendResponse(callback, { error: "Nenhum nome encontrado na coluna " + col + "." });
    }

    return sendResponse(callback, { names: names });
  } catch (err) {
    return sendResponse(callback, { error: "Erro: " + err.message });
  }
}

function sendResponse(callback, data) {
  var json = JSON.stringify(data);

  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
