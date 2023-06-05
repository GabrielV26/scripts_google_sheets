function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Preencher Tabela')
    .addItem('Abrir Caixa de Diálogo', 'showDialog')
    .addToUi();
}

function showDialog() {
  var htmlOutput = HtmlService
    .createHtmlOutputFromFile('formPreenchimento.html')
    .setWidth(600)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Preencher Dados');
}

function saveData(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rowData = [data.name, data.email, data.phone, data.cpf, data.endereco, data.ambiente, data.orcamento];
  sheet.appendRow(rowData);
}