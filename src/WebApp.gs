function listarAtividades() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(ABA_ATIVIDADES);
  if (!aba) return [];

  var dados = aba.getDataRange().getValues();
  var resultado = [];

  for (var i = 1; i < dados.length; i++) {
    var l = dados[i];
    if (!l[COL.DESCRICAO]) continue;
    resultado.push({
      linha:         i + 1,
      id:            l[COL.ID],
      dataAbertura:  _fmtData(l[COL.DATA_ABERTURA]),
      descricao:     l[COL.DESCRICAO],
      tipo:          l[COL.TIPO],
      responsavel:   l[COL.RESPONSAVEL],
      setorExterno:  l[COL.SETOR_EXTERNO],
      status:        l[COL.STATUS],
      prazo:         _fmtData(l[COL.PRAZO]),
      dataConclusao: _fmtData(l[COL.DATA_CONCLUSAO]),
      observacoes:   l[COL.OBSERVACOES]
    });
  }

  return resultado;
}

function salvarAtividade(dados) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(ABA_ATIVIDADES);
  if (!aba) return { ok: false, msg: 'Aba não encontrada.' };

  var linha = [
    dados.id,
    dados.dataAbertura  ? new Date(dados.dataAbertura)  : '',
    dados.descricao,
    dados.tipo,
    dados.responsavel,
    dados.setorExterno,
    dados.status,
    dados.prazo         ? new Date(dados.prazo)         : '',
    dados.dataConclusao ? new Date(dados.dataConclusao) : '',
    dados.observacoes
  ];

  if (dados.linhaEdicao) {
    aba.getRange(dados.linhaEdicao, 1, 1, 10).setValues([linha]);
  } else {
    // Gera ID automático se vazio
    if (!linha[COL.ID]) {
      var total = aba.getLastRow();
      linha[COL.ID] = 'DIV-' + String(total).padStart(3, '0');
    }
    aba.appendRow(linha);
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
  return {
    statusValidos: STATUS_VALIDOS,
    tiposValidos:  TIPOS_VALIDOS
  };
}

function _fmtData(val) {
  if (!val || !(val instanceof Date)) return '';
  return Utilities.formatDate(val, 'America/Sao_Paulo', 'yyyy-MM-dd');
}
