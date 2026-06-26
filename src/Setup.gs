function configurarPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  _configurarAbaAtividades(ss);
  _configurarAbaDashboard(ss);

  ss.setActiveSheet(ss.getSheetByName(ABA_ATIVIDADES));
  ss.toast('Planilha configurada com sucesso!', 'DIVPGC', 5);
}

function _configurarAbaAtividades(ss) {
  var aba = ss.getSheetByName(ABA_ATIVIDADES);
  if (!aba) {
    aba = ss.insertSheet(ABA_ATIVIDADES);
  }

  // Cabeçalhos
  var cabecalhos = [
    'ID', 'Data de Abertura', 'Descrição / Título', 'Tipo',
    'Responsável Interno', 'Setor Externo Dependente', 'Status',
    'Prazo / Previsão', 'Data de Conclusão', 'Observações'
  ];
  var rangeCab = aba.getRange(1, 1, 1, cabecalhos.length);
  rangeCab.setValues([cabecalhos]);
  rangeCab.setFontWeight('bold')
          .setBackground('#4a86e8')
          .setFontColor('#ffffff')
          .setHorizontalAlignment('center');

  // Larguras de coluna
  var larguras = [80, 130, 300, 130, 160, 170, 130, 130, 130, 250];
  larguras.forEach(function(l, i) { aba.setColumnWidth(i + 1, l); });

  // Linha de exemplo
  if (aba.getLastRow() < 2) {
    aba.getRange(2, 1, 1, 10).setValues([[
      'DIV-001',
      new Date(),
      'Exemplo: revisão de regulamento stricto sensu',
      'Resolução',
      'Servidor A',
      'Procuradoria',
      'Em Andamento',
      new Date(Date.now() + 15 * 86400000),
      '',
      'Aguardando parecer jurídico'
    ]]);
  }

  // Validação de dados: Tipo (coluna D)
  var regTipo = SpreadsheetApp.newDataValidation()
    .requireValueInList(TIPOS_VALIDOS, true)
    .setAllowInvalid(false)
    .build();
  aba.getRange(2, COL.TIPO + 1, 1000, 1).setDataValidation(regTipo);

  // Validação de dados: Status (coluna G)
  var regStatus = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_VALIDOS, true)
    .setAllowInvalid(false)
    .build();
  aba.getRange(2, COL.STATUS + 1, 1000, 1).setDataValidation(regStatus);

  // Formato de data nas colunas B, H, I
  var fmtData = 'dd/MM/yyyy';
  [COL.DATA_ABERTURA, COL.PRAZO, COL.DATA_CONCLUSAO].forEach(function(c) {
    aba.getRange(2, c + 1, 1000, 1).setNumberFormat(fmtData);
  });

  // Congelar linha de cabeçalho
  aba.setFrozenRows(1);
}

function _configurarAbaDashboard(ss) {
  var aba = ss.getSheetByName(ABA_DASHBOARD);
  if (!aba) {
    aba = ss.insertSheet(ABA_DASHBOARD);
  } else {
    aba.clearContents();
    aba.clearFormats();
  }

  aba.getRange(1, 1).setValue('Execute "Atualizar Dashboard" no menu DIVPGC para gerar o painel.')
    .setFontStyle('italic')
    .setFontColor('#888888');
}
