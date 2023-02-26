function getDataInDataBaseForSendForEachUser(
  id,
  hour,
  nameValue,
  funcaoValue,
  setorValue,
  categoriaValue,
  subCategoriaValue
) {
  const sheetNames = ["RODRIGO", "RAFAEL", "OTAVIO", "JOAO"];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetData = sheetNames.map(
    (name) => ss.getSheetByName(name).getRange("A2:G").getValues().length
  );
  const minIndex = sheetData.indexOf(Math.min(...sheetData));
  const destSheet = ss.getSheetByName(sheetNames[minIndex]);
  const value = [
    id,
    hour,
    nameValue,
    funcaoValue,
    setorValue,
    categoriaValue,
    subCategoriaValue,
  ];
  destSheet.appendRow(value);
}

function tradeDataAmongUsers(e) {
  // Obtém a planilha ativa e a célula selecionada
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cell = sheet.getActiveCell();

  // Verifica se a célula selecionada está na coluna 12 (L)
  if (cell.getColumn() != 12) {
    return;
  }

  // Define o nome da aba de destino com base no nome selecionado
  var destSheetNames = {
    RODRIGO: "RODRIGO",
    JOAO: "JOAO",
    OTAVIO: "OTAVIO",
    RAFAEL: "RAFAEL",
  };
  var name = cell.getValue();
  var destSheetName = destSheetNames[name];

  // Verifica se o nome da aba de destino é válido
  if (!destSheetName || destSheetName === sheet.getName()) {
    return;
  }

  // Obtém os valores da linha selecionada pela coluna A até H
  var range = sheet.getRange(cell.getRow(), 1, 1, 8);
  var values = range.getValues()[0];

  // Adiciona os valores na última linha vazia em H da aba de destino
  var destSheet = ss.getSheetByName(destSheetName);
  if (!destSheet) {
    throw new Error("A aba de destino '" + destSheetName + "' não existe.");
  }
  var lastRow = destSheet.getLastRow();
  destSheet.getRange(lastRow + 1, 1, 1, values.length).setValues([values]);

  // Deleta a linha original
  sheet.deleteRow(cell.getRow());

  // Exibe uma mensagem de confirmação para o usuário
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    "Os dados foram transferidos com sucesso para a aba '" +
      destSheetName +
      "' e a linha original foi excluída."
  );
}

function onEdit(event) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nameTeste = ss.getActiveSheet().getName();

  const START_ROW = 2;
  const TARGET_COLUMN = 8;
  const TARGET_COLUMN_FINAL = 10;
  const ABA = nameTeste;

  const [row, col, nameSheet, rangee] = [
    event.range.getRow(),
    event.range.getColumn(),
    event.source.getActiveSheet().getName(),
    event.range.getA1Notation(),
  ];

  if (col == TARGET_COLUMN && row >= START_ROW && nameSheet == ABA) {
    const data = new Date();
    event.source
      .getActiveSheet()
      .getRange(rangee)
      .offset(0, 1)
      .setValue(Utilities.formatDate(data, "GMT-3", "HH:mm:ss"));
  }

  if (col == TARGET_COLUMN_FINAL && row >= START_ROW && nameSheet == ABA) {
    const data = new Date();
    event.source
      .getActiveSheet()
      .getRange(rangee)
      .offset(0, 1)
      .setValue(Utilities.formatDate(data, "GMT-3", "HH:mm:ss"));
  }
}
