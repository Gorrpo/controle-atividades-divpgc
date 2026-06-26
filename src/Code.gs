function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('DIVPGC')
    .addItem('Atualizar', 'atualizarDados')
    .addToUi();
}

function atualizarDados() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Dados atualizados!', 'DIVPGC', 3);
}
