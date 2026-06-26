function verificarPrazos() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var aba     = ss.getSheetByName(ABA_ATIVIDADES);
  if (!aba) return;

  var dados  = aba.getDataRange().getValues();
  var hoje   = new Date();
  hoje.setHours(0, 0, 0, 0);
  var limite = new Date(hoje.getTime() + DIAS_ALERTA * 86400000);

  var alertas = [];

  for (var i = 1; i < dados.length; i++) {
    var linha  = dados[i];
    if (!linha[COL.DESCRICAO]) continue;

    var status = linha[COL.STATUS];
    if (status === 'Concluída' || status === 'Arquivada') continue;

    var prazo = linha[COL.PRAZO];
    if (!(prazo instanceof Date)) continue;

    prazo.setHours(0, 0, 0, 0);
    var diasRestantes = Math.floor((prazo - hoje) / 86400000);

    if (diasRestantes <= DIAS_ALERTA) {
      alertas.push({
        id:            linha[COL.ID],
        descricao:     linha[COL.DESCRICAO],
        tipo:          linha[COL.TIPO],
        responsavel:   linha[COL.RESPONSAVEL],
        status:        status,
        prazo:         Utilities.formatDate(prazo, 'America/Sao_Paulo', 'dd/MM/yyyy'),
        diasRestantes: diasRestantes
      });
    }
  }

  if (alertas.length === 0) {
    ss.toast('Nenhum prazo crítico nos próximos ' + DIAS_ALERTA + ' dias.', 'DIVPGC', 4);
    return;
  }

  var corpo = _montarEmailAlertas(alertas);
  var assunto = '[DIVPGC] ' + alertas.length + ' atividade(s) com prazo crítico';

  EMAIL_DESTINATARIOS.forEach(function(email) {
    MailApp.sendEmail({ to: email, subject: assunto, htmlBody: corpo });
  });

  ss.toast(alertas.length + ' alerta(s) enviado(s) por e-mail.', 'DIVPGC', 5);
}

function _montarEmailAlertas(alertas) {
  var linhas = alertas.map(function(a) {
    var cor    = a.diasRestantes < 0 ? '#fce8e6' : (a.diasRestantes <= 3 ? '#fff2cc' : '#ffffff');
    var aviso  = a.diasRestantes < 0
      ? '<strong style="color:#c0392b">VENCIDO (' + Math.abs(a.diasRestantes) + ' dias)</strong>'
      : (a.diasRestantes === 0 ? '<strong style="color:#e67e22">HOJE</strong>'
        : 'em ' + a.diasRestantes + ' dia(s)');
    return '<tr style="background:' + cor + '">'
      + '<td>' + (a.id || '—') + '</td>'
      + '<td>' + a.descricao + '</td>'
      + '<td>' + (a.tipo || '—') + '</td>'
      + '<td>' + (a.responsavel || '—') + '</td>'
      + '<td>' + a.status + '</td>'
      + '<td>' + a.prazo + ' (' + aviso + ')</td>'
      + '</tr>';
  });

  return '<h2 style="font-family:sans-serif">DIVPGC — Alertas de Prazo</h2>'
    + '<p style="font-family:sans-serif">As seguintes atividades estão com prazo crítico (vencidas ou nos próximos '
    + DIAS_ALERTA + ' dias):</p>'
    + '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:sans-serif;font-size:13px">'
    + '<thead style="background:#4a86e8;color:#fff">'
    + '<tr><th>ID</th><th>Descrição</th><th>Tipo</th><th>Responsável</th><th>Status</th><th>Prazo</th></tr>'
    + '</thead><tbody>'
    + linhas.join('')
    + '</tbody></table>'
    + '<p style="font-family:sans-serif;color:#777;font-size:12px">Mensagem automática — DIVPGC/UFT</p>';
}
