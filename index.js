function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ACTIONS')
    .addSubMenu(ui.createMenu('Récupérer les annonces')
      .addItem(`annonces.mesinfos.fr`, 'fetchAnnoncesMesInfos')
      .addItem(`cessions.immobilier-etat.gouv.fr`, 'fetchCessionsImmobilierGouv')
    )
    .addToUi();
  fetchAnnoncesMesInfos();
}

function getByName(colName, line, data) {
  var col = data[0].indexOf(colName);
  if (col != -1) return data[line - 1][col];
}

function getColNumberByName(colName, data) {
  var col = data[0].indexOf(colName);
  if (col != -1) return col + 1;
}

function getSheetByName(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}
