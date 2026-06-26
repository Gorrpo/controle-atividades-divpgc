function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DIVPGC')
    .addItem('Atualizar Dashboard', 'atualizarDashboard')
    .addSeparator()
    .addItem('Enviar Alertas de Prazo', 'verificarPrazos')
    .addItem('Gerar Relatório Mensal', 'gerarRelatorioMensal')
    .addSeparator()
    .addItem('Validar Dados', 'validarDados')
    .addSeparator()
    .addItem('⚙ Configurar Planilha (1ª vez)', 'configurarPlanilha')
    .addItem('Instalar Trigger Diário', 'instalarTriggerDiario')
    .addToUi();
}

function validarDados() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var aba   = ss.getSheetByName(ABA_ATIVIDADES);
  if (!aba) {
    SpreadsheetApp.getUi().alert('Aba "' + ABA_ATIVIDADES + '" não encontrada.');
    return;
  }

  var dados  = aba.getDataRange().getValues();
  var erros  = [];

  for (var i = 1; i < dados.length; i++) {
    var linha  = dados[i];
    var num    = i + 1;
    var status = linha[COL.STATUS];
    var tipo   = linha[COL.TIPO];

    if (status && STATUS_VALIDOS.indexOf(status) === -1) {
      erros.push('Linha ' + num + ': status inválido "' + status + '"');
    }
    if (tipo && TIPOS_VALIDOS.indexOf(tipo) === -1) {
      erros.push('Linha ' + num + ': tipo inválido "' + tipo + '"');
    }
  }

  if (erros.length === 0) {
    ss.toast('Nenhum erro encontrado.', 'DIVPGC — Validação', 4);
  } else {
    SpreadsheetApp.getUi().alert('Erros encontrados:\n\n' + erros.join('\n'));
  }
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('DIVPGC — Controle de Atividades')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function instalarTriggerDiario() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'verificarPrazos') {
      SpreadsheetApp.getActiveSpreadsheet()
        .toast('Trigger já instalado.', 'DIVPGC', 3);
      return;
    }
  }
  ScriptApp.newTrigger('verificarPrazos')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  SpreadsheetApp.getActiveSpreadsheet()
    .toast('Trigger diário instalado (08h).', 'DIVPGC', 4);
}
