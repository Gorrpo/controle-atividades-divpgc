var ABA_ATIVIDADES = 'Atividades 2026';
var ABA_DASHBOARD  = 'Dashboard';

// Índices de colunas (base 0) — espelham o xlsx original
var COL = {
  STATUS:       0,
  PROJETO:      1,
  ATIVIDADE:    2,
  DATA:         3,
  RESPONSAVEL:  4,
  UNIDADES:     5,
  SEI:          6
};

var STATUS_VALIDOS = [
  'Em desenvolvimento',
  'A iniciar a execução',
  'Concluído',
  'SUSPENSO'
];

// Dias de antecedência para disparar alerta (coluna Data usada como referência)
var DIAS_ALERTA = 7;

var EMAIL_DESTINATARIOS = ['leal.leo@gmail.com'];
