function atualizarDashboard() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var abaAtiv = ss.getSheetByName(ABA_ATIVIDADES);
  if (!abaAtiv) {
    SpreadsheetApp.getUi().alert('Aba "' + ABA_ATIVIDADES + '" não encontrada.');
    return;
  }

  var dados = abaAtiv.getDataRange().getValues();
  var hoje  = new Date();

  // Contadores
  var porStatus = {};
  var porTipo   = {};
  var aguardandoAntigos = [];

  STATUS_VALIDOS.forEach(function(s) { porStatus[s] = 0; });
  TIPOS_VALIDOS.forEach(function(t)  { porTipo[t]   = 0; });

  for (var i = 1; i < dados.length; i++) {
    var linha  = dados[i];
    if (!linha[COL.DESCRICAO]) continue; // linha vazia

    var status = linha[COL.STATUS] || 'Não Iniciada';
    var tipo   = linha[COL.TIPO]   || 'Outro';

    if (porStatus[status] !== undefined) porStatus[status]++;
    else porStatus['Não Iniciada']++;

    if (porTipo[tipo] !== undefined) porTipo[tipo]++;
    else porTipo['Outro']++;

    // Atividades "Aguardando" há mais de 30 dias sem conclusão
    if (status === 'Aguardando') {
      var dataAbertura = linha[COL.DATA_ABERTURA];
      if (dataAbertura instanceof Date) {
        var diasEspera = Math.floor((hoje - dataAbertura) / 86400000);
        if (diasEspera > 30) {
          aguardandoAntigos.push({
            id:          linha[COL.ID],
            descricao:   linha[COL.DESCRICAO],
            setor:       linha[COL.SETOR_EXTERNO],
            diasEspera:  diasEspera
          });
        }
      }
    }
  }

  var total = dados.length - 1; // desconta cabeçalho

  // Cria ou limpa aba Dashboard
  var abaDash = ss.getSheetByName(ABA_DASHBOARD);
  if (!abaDash) {
    abaDash = ss.insertSheet(ABA_DASHBOARD);
  } else {
    abaDash.clearContents();
    abaDash.clearFormats();
  }

  var linha = 1;

  // Título
  abaDash.getRange(linha, 1).setValue('DIVPGC — Dashboard de Atividades');
  abaDash.getRange(linha, 1).setFontWeight('bold').setFontSize(14);
  abaDash.getRange(linha, 2).setValue('Atualizado em: ' + Utilities.formatDate(hoje, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'));
  linha += 2;

  // Bloco: Resumo por Status
  abaDash.getRange(linha, 1).setValue('Status').setFontWeight('bold');
  abaDash.getRange(linha, 2).setValue('Qtd').setFontWeight('bold');
  abaDash.getRange(linha, 3).setValue('%').setFontWeight('bold');
  linha++;

  STATUS_VALIDOS.forEach(function(s) {
    var qtd = porStatus[s];
    var pct = total > 0 ? (qtd / total * 100).toFixed(1) + '%' : '0%';
    abaDash.getRange(linha, 1).setValue(s);
    abaDash.getRange(linha, 2).setValue(qtd);
    abaDash.getRange(linha, 3).setValue(pct);
    if (s === 'Aguardando') {
      abaDash.getRange(linha, 1, 1, 3).setBackground('#fff2cc');
    }
    linha++;
  });

  abaDash.getRange(linha, 1).setValue('TOTAL').setFontWeight('bold');
  abaDash.getRange(linha, 2).setValue(total).setFontWeight('bold');
  linha += 2;

  // Bloco: Resumo por Tipo
  abaDash.getRange(linha, 1).setValue('Tipo').setFontWeight('bold');
  abaDash.getRange(linha, 2).setValue('Qtd').setFontWeight('bold');
  linha++;

  TIPOS_VALIDOS.forEach(function(t) {
    abaDash.getRange(linha, 1).setValue(t);
    abaDash.getRange(linha, 2).setValue(porTipo[t]);
    linha++;
  });
  linha++;

  // Bloco: Gargalos — Aguardando há mais de 30 dias
  abaDash.getRange(linha, 1).setValue('Gargalos — Aguardando há mais de 30 dias').setFontWeight('bold').setFontColor('#cc0000');
  linha++;

  if (aguardandoAntigos.length === 0) {
    abaDash.getRange(linha, 1).setValue('Nenhum gargalo identificado.');
    linha++;
  } else {
    abaDash.getRange(linha, 1).setValue('ID').setFontWeight('bold');
    abaDash.getRange(linha, 2).setValue('Descrição').setFontWeight('bold');
    abaDash.getRange(linha, 3).setValue('Setor Externo').setFontWeight('bold');
    abaDash.getRange(linha, 4).setValue('Dias Aguardando').setFontWeight('bold');
    linha++;
    aguardandoAntigos.forEach(function(a) {
      abaDash.getRange(linha, 1).setValue(a.id);
      abaDash.getRange(linha, 2).setValue(a.descricao);
      abaDash.getRange(linha, 3).setValue(a.setor);
      abaDash.getRange(linha, 4).setValue(a.diasEspera);
      abaDash.getRange(linha, 1, 1, 4).setBackground('#fce8e6');
      linha++;
    });
  }

  abaDash.autoResizeColumns(1, 4);
  ss.toast('Dashboard atualizado!', 'DIVPGC', 4);
}
