function listarAtividades() {
  var ss  = getSpreadsheet();
  var aba = ss.getSheetByName(ABA_ATIVIDADES);
  if (!aba) return { erro: 'Aba "' + ABA_ATIVIDADES + '" não encontrada. Execute DIVPGC → Configurar Planilha primeiro.' };

  var dados = aba.getDataRange().getValues();
  var resultado = [];

  for (var i = 1; i < dados.length; i++) {
    var l = dados[i];
    if (!l[COL.ATIVIDADE] && !l[COL.PROJETO]) continue;
    resultado.push({
      linha:       i + 1,
      status:      String(l[COL.STATUS]      || ''),
      projeto:     String(l[COL.PROJETO]     || ''),
      atividade:   String(l[COL.ATIVIDADE]   || ''),
      data:        String(l[COL.DATA]        || ''),
      responsavel: String(l[COL.RESPONSAVEL] || ''),
      unidades:    String(l[COL.UNIDADES]    || ''),
      sei:         String(l[COL.SEI]         || '')
    });
  }

  return resultado;
}

function salvarAtividade(dados) {
  var ss  = getSpreadsheet();
  var aba = ss.getSheetByName(ABA_ATIVIDADES);
  if (!aba) return { ok: false, msg: 'Aba não encontrada.' };

  var linha = [
    dados.status, dados.projeto, dados.atividade,
    dados.data, dados.responsavel, dados.unidades, dados.sei
  ];

  if (dados.linhaEdicao) {
    aba.getRange(dados.linhaEdicao, 1, 1, 7).setValues([linha])
       .setWrap(true).setVerticalAlignment('top');
  } else {
    aba.appendRow(linha);
    var ultima = aba.getLastRow();
    aba.getRange(ultima, 1, 1, 7).setWrap(true).setVerticalAlignment('top');
    aba.setRowHeight(ultima, 80);
    if (ultima % 2 === 0) aba.getRange(ultima, 1, 1, 7).setBackground('#f8f9fa');
  }

  return { ok: true };
}

function excluirAtividade(numeroLinha) {
  var ss  = getSpreadsheet();
  var aba = ss.getSheetByName(ABA_ATIVIDADES);
  if (!aba) return { ok: false };
  aba.deleteRow(numeroLinha);
  return { ok: true };
}

function getConfig() {
  return { statusValidos: STATUS_VALIDOS };
}
