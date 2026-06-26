function listarAtividades() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(ABA_ATIVIDADES);
  if (!aba) return [];

  var dados = aba.getDataRange().getValues();
  var resultado = [];

  for (var i = 1; i < dados.length; i++) {
    var l = dados[i];
    if (!l[COL.ATIVIDADE] && !l[COL.PROJETO]) continue;
    resultado.push({
      linha:       i + 1,
      status:      l[COL.STATUS],
      projeto:     l[COL.PROJETO],
      atividade:   l[COL.ATIVIDADE],
      data:        l[COL.DATA],
      responsavel: l[COL.RESPONSAVEL],
      unidades:    l[COL.UNIDADES],
      sei:         l[COL.SEI]
    });
  }

  return resultado;
}

function salvarAtividade(dados) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(ABA_ATIVIDADES);
  if (!aba) return { ok: false, msg: 'Aba não encontrada.' };

  var linha = [
    dados.status,
    dados.projeto,
    dados.atividade,
    dados.data,
    dados.responsavel,
    dados.unidades,
    dados.sei
  ];

  if (dados.linhaEdicao) {
    aba.getRange(dados.linhaEdicao, 1, 1, 7).setValues([linha]);
    aba.getRange(dados.linhaEdicao, 1, 1, 7).setWrap(true).setVerticalAlignment('top');
  } else {
    aba.appendRow(linha);
    var ultima = aba.getLastRow();
    aba.getRange(ultima, 1, 1, 7).setWrap(true).setVerticalAlignment('top').setRowHeight ? null : null;
    aba.setRowHeight(ultima, 80);
    if (ultima % 2 === 0) aba.getRange(ultima, 1, 1, 7).setBackground('#f8f9fa');
  }

  return { ok: true };
}

function excluirAtividade(numeroLinha) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(ABA_ATIVIDADES);
  if (!aba) return { ok: false };
  aba.deleteRow(numeroLinha);
  return { ok: true };
}

function getConfig() {
  return { statusValidos: STATUS_VALIDOS };
}
