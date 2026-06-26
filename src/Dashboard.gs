function atualizarDashboard() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var abaAtiv = ss.getSheetByName(ABA_ATIVIDADES);
  if (!abaAtiv) { SpreadsheetApp.getUi().alert('Aba "' + ABA_ATIVIDADES + '" não encontrada.'); return; }

  var dados = abaAtiv.getDataRange().getValues();
  var hoje  = new Date();

  var porStatus = { 'Em desenvolvimento': 0, 'A iniciar a execução': 0, 'Concluído': 0, 'SUSPENSO': 0, 'Outros': 0 };
  var porProjeto = {};

  for (var i = 1; i < dados.length; i++) {
    var l = dados[i];
    if (!l[COL.ATIVIDADE] && !l[COL.PROJETO]) continue;

    var s = (l[COL.STATUS] || '').trim();
    var sl = s.toLowerCase();
    if (sl.indexOf('desenvolvimento') > -1 || sl.indexOf('deenvolvimento') > -1) porStatus['Em desenvolvimento']++;
    else if (sl.indexOf('iniciar') > -1) porStatus['A iniciar a execução']++;
    else if (sl.indexOf('conclu') > -1) porStatus['Concluído']++;
    else if (sl.indexOf('suspenso') > -1) porStatus['SUSPENSO']++;
    else porStatus['Outros']++;

    var proj = (l[COL.PROJETO] || 'Sem projeto').substring(0, 50);
    porProjeto[proj] = (porProjeto[proj] || 0) + 1;
  }

  var total = dados.length - 1;

  var abaDash = ss.getSheetByName(ABA_DASHBOARD);
  if (!abaDash) abaDash = ss.insertSheet(ABA_DASHBOARD);
  else { abaDash.clearContents(); abaDash.clearFormats(); }

  var row = 1;

  abaDash.getRange(row, 1).setValue('DIVPGC — Dashboard de Atividades 2026')
    .setFontWeight('bold').setFontSize(14);
  abaDash.getRange(row, 3).setValue('Atualizado: ' +
    Utilities.formatDate(hoje, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'));
  row += 2;

  // Resumo por status
  _cabecalho(abaDash, row, ['Status', 'Qtd', '%']); row++;
  var statusKeys = ['Em desenvolvimento', 'A iniciar a execução', 'Concluído', 'SUSPENSO', 'Outros'];
  statusKeys.forEach(function(s) {
    var qtd = porStatus[s];
    var pct = total > 0 ? (qtd / total * 100).toFixed(1) + '%' : '0%';
    abaDash.getRange(row, 1).setValue(s);
    abaDash.getRange(row, 2).setValue(qtd);
    abaDash.getRange(row, 3).setValue(pct);
    var cor = { 'Em desenvolvimento': '#dbeafe', 'A iniciar a execução': '#fef9c3',
                'Concluído': '#dcfce7', 'SUSPENSO': '#fee2e2' }[s];
    if (cor) abaDash.getRange(row, 1, 1, 3).setBackground(cor);
    row++;
  });
  abaDash.getRange(row, 1).setValue('TOTAL').setFontWeight('bold');
  abaDash.getRange(row, 2).setValue(total).setFontWeight('bold');
  row += 2;

  // Top projetos
  abaDash.getRange(row, 1).setValue('Projetos com mais atividades').setFontWeight('bold').setFontSize(11);
  row++;
  _cabecalho(abaDash, row, ['Projeto', 'Qtd']); row++;
  var projs = Object.keys(porProjeto).sort(function(a,b){ return porProjeto[b]-porProjeto[a]; }).slice(0, 15);
  projs.forEach(function(p) {
    abaDash.getRange(row, 1).setValue(p);
    abaDash.getRange(row, 2).setValue(porProjeto[p]);
    row++;
  });

  abaDash.setColumnWidth(1, 320);
  abaDash.setColumnWidth(2, 80);
  abaDash.setColumnWidth(3, 80);
  ss.toast('Dashboard atualizado!', 'DIVPGC', 4);
}

function _cabecalho(aba, row, cols) {
  var r = aba.getRange(row, 1, 1, cols.length);
  r.setValues([cols]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
}
