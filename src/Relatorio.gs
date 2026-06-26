function gerarRelatorioMensal() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var aba   = ss.getSheetByName(ABA_ATIVIDADES);
  if (!aba) {
    SpreadsheetApp.getUi().alert('Aba "' + ABA_ATIVIDADES + '" não encontrada.');
    return;
  }

  var dados = aba.getDataRange().getValues();
  var hoje  = new Date();
  var mes   = hoje.getMonth();
  var ano   = hoje.getFullYear();

  var atividadesMes   = [];
  var concluidasMes   = [];
  var porStatus       = {};
  var porTipo         = {};

  STATUS_VALIDOS.forEach(function(s) { porStatus[s] = 0; });
  TIPOS_VALIDOS.forEach(function(t)  { porTipo[t]   = 0; });

  for (var i = 1; i < dados.length; i++) {
    var linha = dados[i];
    if (!linha[COL.DESCRICAO]) continue;

    var status = linha[COL.STATUS] || 'Não Iniciada';
    var tipo   = linha[COL.TIPO]   || 'Outro';

    // Contagem geral
    if (porStatus[status] !== undefined) porStatus[status]++;
    if (porTipo[tipo] !== undefined)     porTipo[tipo]++;

    // Atividades abertas no mês
    var dataAbertura = linha[COL.DATA_ABERTURA];
    if (dataAbertura instanceof Date &&
        dataAbertura.getMonth() === mes &&
        dataAbertura.getFullYear() === ano) {
      atividadesMes.push(linha);
    }

    // Concluídas no mês
    var dataConclusao = linha[COL.DATA_CONCLUSAO];
    if (dataConclusao instanceof Date &&
        dataConclusao.getMonth() === mes &&
        dataConclusao.getFullYear() === ano) {
      concluidasMes.push(linha);
    }
  }

  var nomeMes   = _nomeMes(mes);
  var assunto   = '[DIVPGC] Relatório Mensal — ' + nomeMes + '/' + ano;
  var corpo     = _montarEmailRelatorio(nomeMes, ano, porStatus, porTipo, atividadesMes, concluidasMes, dados.length - 1);

  EMAIL_DESTINATARIOS.forEach(function(email) {
    MailApp.sendEmail({ to: email, subject: assunto, htmlBody: corpo });
  });

  ss.toast('Relatório mensal enviado por e-mail.', 'DIVPGC', 4);
}

function _montarEmailRelatorio(nomeMes, ano, porStatus, porTipo, abertas, concluidas, total) {
  var linhasStatus = STATUS_VALIDOS.map(function(s) {
    var qtd = porStatus[s];
    var pct = total > 0 ? (qtd / total * 100).toFixed(1) + '%' : '0%';
    var bg  = s === 'Aguardando' ? '#fff2cc' : '#ffffff';
    return '<tr style="background:' + bg + '"><td>' + s + '</td><td>' + qtd + '</td><td>' + pct + '</td></tr>';
  });

  var linhasTipo = TIPOS_VALIDOS.map(function(t) {
    return '<tr><td>' + t + '</td><td>' + porTipo[t] + '</td></tr>';
  });

  function tabelaAtividades(lista) {
    if (!lista.length) return '<p><em>Nenhuma atividade no período.</em></p>';
    var rows = lista.map(function(l) {
      return '<tr>'
        + '<td>' + (l[COL.ID] || '—') + '</td>'
        + '<td>' + l[COL.DESCRICAO] + '</td>'
        + '<td>' + (l[COL.TIPO] || '—') + '</td>'
        + '<td>' + (l[COL.RESPONSAVEL] || '—') + '</td>'
        + '<td>' + (l[COL.STATUS] || '—') + '</td>'
        + '</tr>';
    });
    return '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;font-size:12px">'
      + '<thead style="background:#4a86e8;color:#fff"><tr><th>ID</th><th>Descrição</th><th>Tipo</th><th>Responsável</th><th>Status</th></tr></thead>'
      + '<tbody>' + rows.join('') + '</tbody></table>';
  }

  return '<div style="font-family:sans-serif">'
    + '<h2>DIVPGC — Relatório Mensal: ' + nomeMes + '/' + ano + '</h2>'

    + '<h3>Panorama Geral (acumulado)</h3>'
    + '<p>Total de atividades registradas: <strong>' + total + '</strong></p>'
    + '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse">'
    + '<thead style="background:#4a86e8;color:#fff"><tr><th>Status</th><th>Qtd</th><th>%</th></tr></thead>'
    + '<tbody>' + linhasStatus.join('') + '</tbody></table>'
    + '<br>'
    + '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse">'
    + '<thead style="background:#4a86e8;color:#fff"><tr><th>Tipo</th><th>Qtd</th></tr></thead>'
    + '<tbody>' + linhasTipo.join('') + '</tbody></table>'

    + '<h3>Abertas em ' + nomeMes + '/' + ano + ' (' + abertas.length + ')</h3>'
    + tabelaAtividades(abertas)

    + '<h3>Concluídas em ' + nomeMes + '/' + ano + ' (' + concluidas.length + ')</h3>'
    + tabelaAtividades(concluidas)

    + '<p style="color:#777;font-size:11px">Mensagem automática — DIVPGC/UFT</p>'
    + '</div>';
}

function _nomeMes(indice) {
  var nomes = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
               'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  return nomes[indice];
}
