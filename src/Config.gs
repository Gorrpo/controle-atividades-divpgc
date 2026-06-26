// Aba principal de dados
var ABA_ATIVIDADES = 'Atividades';
var ABA_DASHBOARD  = 'Dashboard';

// Índices de colunas (base 0)
var COL = {
  ID:               0,
  DATA_ABERTURA:    1,
  DESCRICAO:        2,
  TIPO:             3,
  RESPONSAVEL:      4,
  SETOR_EXTERNO:    5,
  STATUS:           6,
  PRAZO:            7,
  DATA_CONCLUSAO:   8,
  OBSERVACOES:      9
};

var STATUS_VALIDOS = [
  'Não Iniciada',
  'Em Andamento',
  'Aguardando',
  'Concluída',
  'Arquivada'
];

var TIPOS_VALIDOS = [
  'APCN',
  'Minter/Dinter',
  'Resolução',
  'Edital',
  'Fluxo Contínuo',
  'Outro'
];

// Dias de antecedência para disparar alerta de prazo
var DIAS_ALERTA = 7;

// E-mails que recebem alertas e relatórios (separados por vírgula se múltiplos)
var EMAIL_DESTINATARIOS = ['leal.leo@gmail.com'];
