function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('SIPTAMA Batola')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function cekDataMadrasah(nsm) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  SpreadsheetApp.flush(); 
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var nsmColIndex = headers.indexOf("NSM");
  if (nsmColIndex === -1) return null;
  var nsmCari = nsm.toString().trim();
  for (var i = data.length - 1; i > 0; i--) {
    if (data[i][nsmColIndex].toString().trim() === nsmCari) {
      var result = {};
      headers.forEach((header, index) => {
        var value = data[i][index];
        result[header] = (value instanceof Date) ? Utilities.formatDate(value, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss") : value;
      });
      return result;
    }
  }
  return null;
}

function prosesFormulir(dataObj) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var nsmColIndex = headers.indexOf("NSM");
  var rowIndex = -1;
  if (nsmColIndex !== -1 && dataObj["NSM"]) {
    var nsmCari = dataObj["NSM"].toString().trim();
    for (var i = 1; i < data.length; i++) {
      if (data[i][nsmColIndex].toString().trim() == nsmCari) { rowIndex = i + 1; break; }
    }
  }
  var rowValue = headers.map(function(header, index) {
    if (index === 0) return new Date();
    return dataObj[header] || "";
  });
  if (rowIndex !== -1) {
    sheet.getRange(rowIndex, 1, 1, rowValue.length).setValues([rowValue]);
    return "Data Madrasah Berhasil Diperbarui!";
  } else {
    sheet.appendRow(rowValue);
    return "Data Madrasah Baru Berhasil Disimpan!";
  }
}