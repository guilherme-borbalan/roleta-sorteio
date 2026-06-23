// ============================================================
// COLE ESTE CODIGO NO GOOGLE APPS SCRIPT (script.google.com)
// ============================================================
// IMPORTANTE: Apos colar, va em Implantar > Nova implantacao
//   Tipo: App da Web
//   Executar como: Eu
//   Quem tem acesso: Qualquer pessoa
// ============================================================

function doGet(e) {
  var params = e.parameter || {};
  var sheetId = params.id || "";
  var sheetName = params.sheet || "Planilha1";
  var col = params.col || "A";

  if (!sheetId) {
    return jsonResponse({
      error: "Parametro 'id' nao informado. Este script deve ser chamado pela Roleta de Sorteio."
    });
  }

  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return jsonResponse({ error: "Aba '" + sheetName + "' nao encontrada na planilha." });
    }

    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      return jsonResponse({ error: "A planilha esta vazia." });
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
      return jsonResponse({ error: "Nenhum nome encontrado na coluna " + col + "." });
    }

    return jsonResponse({ names: names });
  } catch (err) {
    return jsonResponse({ error: "Erro: " + err.message });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
