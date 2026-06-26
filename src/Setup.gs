function configurarPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  _configurarAbaAtividades(ss);
  _configurarAbaDashboard(ss);
  ss.setActiveSheet(ss.getSheetByName(ABA_ATIVIDADES));
  ss.toast('Planilha configurada e dados importados!', 'DIVPGC', 5);
}

function _configurarAbaAtividades(ss) {
  var aba = ss.getSheetByName(ABA_ATIVIDADES);
  if (!aba) aba = ss.insertSheet(ABA_ATIVIDADES);
  aba.clearContents();
  aba.clearFormats();

  // Cabeçalhos — mesma ordem do xlsx original
  var cab = [
    'Status',
    'Projeto / Ações / Serviços',
    'Atividade',
    'Data de realização',
    'Facilitador/a de aprendizagem / Responsável',
    'Unidades Responsáveis',
    'SEI'
  ];
  var rangeCab = aba.getRange(1, 1, 1, cab.length);
  rangeCab.setValues([cab])
          .setFontWeight('bold')
          .setBackground('#1a73e8')
          .setFontColor('#ffffff')
          .setHorizontalAlignment('center')
          .setWrap(true);

  // Larguras
  aba.setColumnWidth(1, 140);
  aba.setColumnWidth(2, 280);
  aba.setColumnWidth(3, 380);
  aba.setColumnWidth(4, 200);
  aba.setColumnWidth(5, 180);
  aba.setColumnWidth(6, 160);
  aba.setColumnWidth(7, 180);

  // Dados — todos os 67 registros do xlsx 2026
  var dados = [
    ["Em desenvolvimento", "Carteira de Identidade Nacional", "Ofício;\nTratativas para a atividade; \nDespacho para DTINF / DIADM/ Organização de pessoas para  o atendimetno/suporte/ Expedição dec CIN\nEvento\nEntregas das CIN e relatórios", "Comarcas de Gurupi, nos dias 3 e 4 de agosto de 2026;\n\nAraguaína, nos dias 6 e 7 de agosto de 2026;\n\ne Palmas (TJ e Fórum), nos dias 27 e 28 de agosto de 2026.", "Leonardo Leal", "DIVPGC", "26.0.000004764-2"],
    ["Em desenvolvimento", "Projeto Justiça em Dialogo", "Cartilha", "2025/2026", "Janaína", "DIVPGC", "25.0.000026747-6"],
    ["Concluído", "Composição da Comissão CASSEDIO- TJTO", "Eleição de estagiários para composição da Comissão CASSEDIO", "23/03 a 04/05/2026 - edital no evento 7042902", "Leonardo e Alessandra", "DIVPGC", "23.0.000028360-6"],
    ["Concluído", "Projeto Oficinas Literárias", "I Oficina Literária I - Obra: \"O estranho caso do cachorro morto\", de Mark Haddon; Tema: Direitos das pessoas com deficiência, acessibilidade e inclusão;  Comissão parceira: Comissão Permanente de Acessibilidade e Inclusão - CPAI.", "Maio de 2026- Em 14,21 e 28/05\nObs. inscrições ocorrerão no período de 27 de abril a 7 de maio de 2026", "Janaína", "DIVPGC", "26.0.000001284-9 / 26.0.000008400-9"],
    ["Em desenvolvimento", "Projeto Oficinas Literárias", "II - Oficina Literária II - Obra: Tudo é Rio, de Carla Madeira; Tema: Violência de gênero e direitos das mulheres;  Comissão parceira: Comissão de Política de Equidade de Gênero e Racial - CGPEG .", "10/06/2026;\n17/06/2026;\n24/06/2026", "Janaína", "DIVPGC", "26.0.000001284-9"],
    ["A iniciar a execução", "Projeto Oficinas Literárias", "III -Oficina Literária III - Obra: A Contagem dos Sonhos, de Chimamanda Ngozi Adichie; Tema: Identidade negra, pandemia, colonialismo e imigração;Comissão parceira:  Observatório de Direitos Humanos;", "Setembro de 2026;", "Janaína", "DIVPGC", "26.0.000001284-9"],
    ["A iniciar a execução", "Projeto Oficinas Literárias", "IV - Oficina Literária IV - Obra: Canção para Ninar Menino Grande, de Conceição Evaristo; Tema: Desigualdade social e gênero; Comissão parceira:  Comissão de Prevenção e Enfrentamento do Assédio;", "Outubro de 2026;", "Janaína", "DIVPGC", "26.0.000001284-9"],
    ["Em Desenvolvimento", "Projeto Oficinas Literárias", "V - Sarau das Oficinas Literárias - Atividade: Encontro anual com participantes dos círculos de leitura e público em geral;  Parceiros: GGEM, ESMAT, Comissões participantes das oficinas e Comitê Gestor Local de Gestão de Pessoas - CGLGP", "Novembro de 2026;", "Janaína", "DIVPGC", "26.0.000001284-9 e 26.0.000008987-6"],
    ["Em desenvolvimento", "Metas  CPAI", "Executar oficina literária nº 1", "03/2026 a 07/2026", "UNIDADE DIVPGC- PRIORIDADE", "DIVPGC", "24.0.000001639-6\nEvento 6938384"],
    ["Em desenvolvimento", "Projeto de Aperfeiçoamento da Gestão de Pessoas por Competências- Ano 2025/2026 ( Meta 19)", "Curso Autoinstrucional- Gestão por Competências: fundamentos e práticas - 1º GRAU", "18 de maio a 15 de novembro de 2026.", "Tania, Leonardo e Silvia", "DIVPGC/SGPC", "26.0.000001603-8 - Esmat,  26.0.000005651-0 (contratação Esmat) e 26.0.000006806-2 (convocação)"],
    ["Em desenvolvimento", "Projeto de Aperfeiçoamento da Gestão de Pessoas por Competências- Ano 2025/2026 ( Meta 19)", "Curso Autoinstrucional- Gestão por Competências: fundamentos e práticas - 2º GRAU", "18 de maio a 15 de novembro de 2026.", "Tania, Leonardo e Silvia", "DIVPGC/SGPC", "26.0.000001603-8- Esmat e 26.0.000006800-3 (contratação Esmat), 26.0.000006806-2(Convocação)."],
    ["Em desenvolvimento", "Projeto de Aperfeiçoamento da Gestão de Pessoas por Competências- Ano 2025/2026 ( Meta 19)", "Estabelecimento de fluxo para criação, alteração e extinção de unidades", "2025/2026", "Tania, Leonardo e Silvia", "DIVPGC/SGPC", "25.0.000006505-9 (7032611)"],
    ["Em desenvolvimento", "Projeto de Aperfeiçoamento da Gestão de Pessoas por Competências- Ano 2025/2026 ( Meta 19)", "Reeedição de videos tutoriais para  uso no mapeamento de unidades", "03/2026 a 07/2026", "Tania, Leonardo e Silvia", "DIGEP/CECOM", "26.0.000006271-4"],
    ["CONCLUÍDO", "Projeto de Aperfeiçoamento da Gestão de Pessoas por Competências- Ano 2025/2026 ( Meta 19)", "Mapeamento da unidade: NÚCLEO DE GERENCIAMENTO DE PRECEDENTES E AÇÕES COLETIVAS (NUGEPAC),", "01/2026 a 07/2026", "Silvia, Leo e Tania", "DIVPGC/SGPC", "26.0.000000120-0"],
    ["Em desenvolvimento", "Projeto de Aperfeiçoamento da Gestão de Pessoas por Competências- Ano 2025/2026 ( Meta 19)", "Mapeamento da unidade: DIVISÃO DE CONTROLE E ANÁLISE PROCESSUAL (DCAPDIJUD)", "01/2026 a 07/2026", "Silvia, Leo e Tania", "DIVPGC/SGPC", "26.0.000000114-6 /25.0.000018639-5"],
    ["CONCLUÍDO", "Projeto de Aperfeiçoamento da Gestão de Pessoas por Competências- Ano 2025/2026 ( Meta 19)", "Mapeamento da unidade: DIVISÃO DE BAIXA E PUBLICIDADE (DBPDIJUD)", "01/2026 a 07/2026", "Silvia, Leo e Tania", "DIVPGC/SGPC", "26.0.000000108-1"],
    ["Em desenvolvimento", "Projeto de Aperfeiçoamento da Gestão de Pessoas por Competências- Ano 2025/2026 ( Meta 19)", "Mapeamento da unidade: Gabinete do Controlador Interno (GCONTI)", "01/2026 a 07/2026", "Silvia, Leo e Tania", "DIVPGC/SGPC", "25.0.000025051-4"],
    ["CONCLUÍDO", "Projeto de Aperfeiçoamento da Gestão de Pessoas por Competências- Ano 2025/2026 ( Meta 19)", "Mapeamento de Unidade: Serviço de Suprimento de Fundo (SSFDIFIN) - Indicada: Celma", "01/2026 a 07/2026", "Silvia, Leo e Tania", "DIVPGC/SGPC", "25.0.000022506-4"],
    ["Em desenvolvimento", "Projeto de Aperfeiçoamento da Gestão de Pessoas por Competências- Ano 2025/2026 ( Meta 19)", "Apresentação das unidades não mapeadas(15) e procedimentos para mapeamento", "01/2026 a 12/2026", "Silvia, Leo e Tania", "DIVPGC/SGPC", "25.0.000018639-5"],
    ["Em desenvolvimento", "GPC- Operacional - Encaminhamento para a Capacitação", "Trabalho em equipe: integração e inteligência relacional nas equipes do Judiciáiro  - EAD - 10 horas", "12 de junho a 15 de novembro de 2026,", "Silvia, Leo e Tania", "DIVPGC/ESMAT", "26.0.000006386-9 e 26.0.000008040-2(Esmat) e 26.0.000009410-1"],
    ["Em desenvolvimento", "GPC- Operacional - Encaminhamento para a Capacitação", "Curso Ética no Poder Judiciário -EAD- 10 horas", "12 de junho a 15 de novembro de 2026,", "Silvia, Leo e Tania", "DIVPGC/ESMAT", "26.0.000008044-5(Esmat)\n26.0.000009406-3"],
    ["Em desenvolvimento", "GPC- Operacional - Encaminhamento para a Capacitação", "Curso Inovação e Criatividade no Poder Judiciário  0- EAD- 10 horas", "12 de junho a 15 de novembro de 2026,", "Silvia, Leo e Tania", "DIVPGC/ESMAT", "26.0.000008131-0(Esmat) e 26.0.000009408-0"],
    ["Em desenvolvimento", "GPC- Operacional - Encaminhamento para a Capacitação", "curso Flexibilidade no Serviço Público: Autoconhecimento e Autogerenciamento para Autoperformance no Judiciário  - EAD -10 horas", "12 de junho a 15 de novembro de 2026,", "Silvia, Leo e Tania", "DIVPGC/ESMAT", "26.0.000006055-0, 26.0.000008130-1\n(Esmat) e 26.0.000009405-5"],
    ["Em desenvolvimento", "GPC- Operacional - Encaminhamento para a Capacitação", "Gestão de Projetos no Judiciário", "12 de junho a 15 de novembro de 2026,", "Silvia, Leo e Tania", "DIVPGC/ESMAT", "26.0.000009910-3 (Esmat) e  26.0.000010482-4"],
    ["Em desenvolvimento", "Projeto de Acompanhamento e Monitoramento dos Operadores do Teletrabalho- 2025", "a) Mediação para capacitação (Google Workspace) - Realização Esmat Período: fevereiro a outubro de 2026; Público: gestores(as) e teletrabalhadores(as);", "01/2026 a 10/2026", "Tania/João/Barbara", "DIVPGC- SATT/DIGEP", "26.0.000002062-0/ 26.0.000008365-7"],
    ["Em desenvolvimento", "Projeto de Acompanhamento e Monitoramento dos Operadores do Teletrabalho- 2025", "b) Webinário: Entendendo o Teletrabalho no TJTO: Data: 14/09/2026; Horário: 13h às 15h; Modalidade: on-line; Público: Toda a força de trabalho do Poder Judiciário", "14/09/2026", "Tania/João/Barbara", "DIVPGC- SATT/DIGEP", "26.0.000002062-0"],
    ["Em desenvolvimento", "Projeto de Acompanhamento e Monitoramento dos Operadores do Teletrabalho- 2025", "c) Workshop – Teletrabalho, Saúde e Qualidade de Vida: Data: 15/09/2026; Horário: 13h às 15h; Modalidade: on-line; Público: Servidores(as) teletrabalhadores(as) e Gestores(as) do teletrabalho", "15/09/2026", "Tania/João/Barbara", "DIVPGC- SATT/DIGEP", "26.0.000002062-0"],
    ["Em desenvolvimento", "Projeto de Acompanhamento e Monitoramento dos Operadores do Teletrabalho- 2025", "d) Aplicação de formulários de saúde e qualidade de vida: Período: fevereiro a maio de 2026; Público: novos(as) teletrabalhadores(as)", "01/2026 a 10/2026", "Tania/João/Barbara", "DIVPGC- SATT/DIGEP", "26.0.000002062-0"],
    ["A iniciar a execução", "Projeto de Acompanhamento e Monitoramento dos Operadores do Teletrabalho- 2025", "e) Entrevistas sociais: Período: maio a julho de 2026; Público: novos(as) teletrabalhadores(as)", "01/2026 a 10/2026", "Tania/João/Barbara", "DIVPGC- SATT/DIGEP", "26.0.000002062-0"],
    ["CONCLUÍDO", "Solicitação de reedição do Curso Autoinstrucional Preparatório para o Teletrabalho – Turma II", "Disponiblização do Curso de teletrabalho - EAD/Autoinstrucional", "3 de março a 30 de Outubro de 2026.", "João e Tania", "DIVPGC/DIGEP/CGTT", "26.0.000002941-5"],
    ["Em desenvolvimento", "Módulo e-GESP", "Desenvolvimento de um módulo no e-GESP, destinado à realização e à gestão de pesquisas, formulários e enquetes,", "01/2026 a 12/2026", "Leonardo e Tania", "DIVPGC/DIGEP", "25.0.000018166-0"],
    ["SUSPENSO", "Módulo e-GESP", "Módulo do Teletrabalho", "suspenso", "Tania, João e Leonardo", "DIVPGC/DIGEP", "24.0.000010924-6\n25.0.000024314-3"],
    ["CONCLUÍDO", "Programa de Preparação para a Aposentadoria", "Oficinas do PPA", "08, 09 e 10 de junho de 2026", "Janaína", "DIGER/ESMAT/DIGEP/CGJUS/DTINF", "25.0.000000001-1"],
    ["CONCLUÍDO", "Programa de Preparação para a Aposentadoria", "Criação do hotsite", "01/2026 a 07/2026", "Janaína", "DIGER/ESMAT/DIGEP/CGJUS/DTINF", "25.0.000000001-1"],
    ["Em desenvolvimento", "Programa de Preparação para a Aposentadoria", "Monitoramento do cumprimento do Plano de Ação", "01/2026 a 12/2026", "Janaína", "DIGER/ESMAT/DIGEP/CGJUS/DTINF", "25.0.000000001-1 - evento7047899\n26.0.000012451-5 - pendencias do plano de ação"],
    ["Em desenvolvimento", "Projeto Mapeamento dos Fluxos de Trabalho da DIGEP", "Mapeamento de processos das Unidades da DIGEP", "03/2025 a 12/2026", "Leo e Alessandra", "DIVPGC/DIGEP", "25.0.000002423-9"],
    ["Em desenvolvimento", "Plano de Gestão biênio 2025/2027 -Projeto de Bem-Estar Psíquico de Servidores e Magistrados -  Meta 21", "Contratação de Serviço Digital Integrado de Atividades de Bem-Estar - Ação do Projeto Bem-Estar Psíquico", "2025/2026", "Intersetorial", "DIGEP/CESAU/JMED/NApsi (Multisetorial)", "25.0.000018724-3"],
    ["A iniciar a execução", "Plano de Gestão biênio 2025/2027 -Projeto de Bem-Estar Psíquico de Servidores e Magistrados -  Meta 21", "Mentoria: Autocuidado com saberes naturais", "03/2026 a 12/2026", "Sandra Carvalho", "CESAU", "25.0.000018809-6"],
    ["Em deenvolvimento", "Plano de Gestão biênio 2025/2027 -Projeto de Bem-Estar Psíquico de Servidores e Magistrados -  Meta 21", "Microfisioterapia", "03/2026 a 12/2026", "Virlene Maria Pereira Queiroz Torres", "CESAU", "25.0.000018788-0"],
    ["A iniciar a execução", "Plano de Gestão biênio 2025/2027 -Projeto de Bem-Estar Psíquico de Servidores e Magistrados -  Meta 21", "Dados Estatísticos da Junta Médica", "01/2026 a 12/2026", "Jessyca Lira de Carvalho Ferreira", "JMED", "25.0.000018876-2"],
    ["A iniciar a execução", "Plano de Gestão biênio 2025/2027 -Projeto de Bem-Estar Psíquico de Servidores e Magistrados -  Meta 21", "NAPsi Itinerante", "02/2026 a 06/2026", "Wordney Carvalho Camarço", "NAPsi", "25.0.000019064-3"],
    ["A iniciar a execução", "Plano de Gestão biênio 2025/2027 -Projeto de Bem-Estar Psíquico de Servidores e Magistrados -  Meta 21", "Grupo Psicoeducativo com Mulheres do Judiciário do Estado do Tocantins", "01/2026 a 11/2026", "Fernanda Maria dos Santos Abreu", "NAPsi", "25.0.000019059-7"],
    ["A iniciar a execução", "Plano de Gestão biênio 2025/2027 -Projeto de Bem-Estar Psíquico de Servidores e Magistrados -  Meta 21", "Rodas de Conversa sobre Saúde Mental", "02/2026 a 07/2026", "Sérgio Baggio", "NAPsi", "25.0.000019071-6"],
    ["A iniciar a execução", "Plano de Gestão biênio 2025/2027 -Projeto de Bem-Estar Psíquico de Servidores e Magistrados -  Meta 21", "Campanha Anual em alusão ao Setembro Amarelo", "03/2026 a 09/2026", "Janaína Rodrigues Araújo", "NAPsi", "25.0.000019061-9"],
    ["A iniciar a execução", "Plano de Gestão biênio 2025/2027 -Projeto de Bem-Estar Psíquico de Servidores e Magistrados -  Meta 21", "Grupo Terapêutico com Cuidadores (Pais e Mães) de Pessoas com Neurodiversidades", "01/2026 a 09/2026", "Virgínia de Moura Fragoso", "NAPsi", "25.0.000019054-6"],
    ["A iniciar a execução", "Plano de Gestão biênio 2025/2027 -Projeto de Bem-Estar Psíquico de Servidores e Magistrados -  Meta 21", "Dados Estatísticos do NAPsi", "01/2026 a 12/2026", "Sérgio Baggio", "NAPsi", "25.0.000019062-7"],
    ["A iniciar a execução", "Plano de Gestão biênio 2025/2027 -Projeto de Bem-Estar Psíquico de Servidores e Magistrados -  Meta 21", "Conexão e Bem-Estar no Trabalho ( Fluxo de condições especiais de trabalho)", "03/2026 a 08/2026", "Leonardo Andrade Leal", "DIGEP/CESAU/JMED/NApsi (Multisetorial)", "25.0.000018725-1"],
    ["Em deenvolvimento", "Plano de Gestão biênio 2025/2027 -Projeto de Bem-Estar Psíquico de Servidores e Magistrados -  Meta 21", "Mapeamento de Fluxo Operacional de Readaptação( Fluxo para retorno de licença médica)", "03/2026 a 12/2026", "Tania Mara Alves Barbosa", "DIGEP/CESAU/JMED/NApsi (Multisetorial)", "25.0.000018744-8"],
    ["Em deenvolvimento", "Plano de Gestão biênio 2025/2027 -Projeto de Bem-Estar Psíquico de Servidores e Magistrados -  Meta 21", "Protocolo de Ações em Situações de Emergência em Saúde Mental", "02/2026 a 12/2026", "Alessandra Santos Santana", "DIGEP/CESAU/JMED/NApsi (Multisetorial)", "25.0.000018727-8"],
    ["Em desenvolvimento", "Plano de Gestão biênio 2025/2027 -Programa de Bem-Estar Psíquico de Servidores e Magistrados -  Meta 21", "Projeto com 12 planos de ação", "02/2026 a 12/2026", "Ver", "DIVPGC/DIGEP", "25.0.000006311-0"],
    ["Em desenvolvimento", "Evento Dia do Servidor 2026", "levantamento de valores e temas para a atividade.", "outubro de 2026", "Leonardo/Silvia/Tania", "DIVPGC/DIGEP", "26.0.000012030-7"],
    ["Em desenvolvimento", "Projeto Plantar  para Compensar", "Realizar relatórios de monitoramento (semestral e anual) do projeto \"Plantar para comPENSAR 2025\"", "relatório semestral deverá ser apresentado no mês de junho de 2026 e o relatório anual em dezembro de 2026", "Alessandra e Tania", "DIVPGC", "25.0.000025337-8"],
    ["CONCLUÍDO", "II Encontro de Secretário de Juízo do TJTO.", "Organizar os temas e atividades da DIVPGC para compartilhar no evento", "08 e 09/04/2026", "Bárbara", "DIGEP e DIVPGC", "26.0.000010275-9 / 26.0.000003428-1"],
    ["Em desenvolvimento", "Plano de Gestão biênio 2025/2027- Projeto Permanente de Formação de Lideres - Meta 22", "Projeto elaborado e aprovado pela decisão do evento 7024199", "03/2025 a 02/2027", "Alessandra", "DIVPGC/DIGEP", "25.0.000006300-5"],
    ["Em desenvolvimento", "Plano de Gestão biênio 2025/2027- Projeto Permanente de Formação de Lideres - Meta 22", "Processo para solicitar a aquisição dos botons", "03/2025 a 02/2027", "Alessandra", "DIVPGC/DIGEP", "26.0.000006586-1"],
    ["Em desenvolvimento", "Plano de Gestão biênio 2025/2027- Programa Permanente de Formação Inicial para Novos Servidores  - Meta 24", "Projeto elaborado e aprovado pela decisão do evento 7024199", "Prazo final-12/2026", "Matheus", "GABDIGEP", "25.0.000006281-5"],
    ["CONCLUÍDO", "LANÇAMENTO DO Programa Permanente de Formação Inicial para Novos Servidores  - Meta 24 e do Projeto Permanente de Formação de Lideres - Meta 22", "LANÇAMENTO - Contratação de palestrantes de evento de abertura para a apresentação-Projeto Permanente de Formação de  Lideres e  Formação Inciial para Novos Servidores", "12 de junho de 2026, às 14 horas, no Auditório do Tribunal de Justiça do Estado do Tocantins.", "Matheus", "GABDIGEP", "26.0.000006436-9"],
    ["CONCLUÍDO", "Comitê de Governança e Gestão das Contratações -Representação da DIGEP", "Participar das reuniões e cumprir plano de ação", "Até 07/2026", "Tania/Vinicius e Paula Márcia", "DIVPGC", "23.0.000021887-1 - 6485101"],
    ["A iniciar a execução", "Plano Diretor da DIGEP", "Projeto DIGEP por Dentro - eixo comunicação", "01/12/2026", "DIVPGC", "DIVPGC", "23.0.000047556-4/ 26.0.000011819-1"],
    ["A iniciar a execução", "Projeto de Educação Financeira", "Acompanhamento das ações", "01/05/2026", "Leonardo/Tania/Barbara", "DIVPGC", "26.0.000005448-7"],
    ["Em desenvolvimento", "Fluxo de avaliação biopsicossocial da Pessoa com Deficiência", "eEaboração de fluxograma", "avaliação biopsicossocial", "a definir", "/ DIVPGC /  junta medica", "26.0.000001182-6"],
    ["Em desenvolvimento", "Comitê Estadual", "Comitê Interinstitucional de Promoção de Políticas Públicas Judiciais de Atenção às Pessoas Idosas do Tocantins", "representação no comitê", "João Gabriel", "DIVPGC", "26.0.000003501-6 / 26.0.000007341-4"],
    ["Em desenvolvimento", "Realização de evento em atendimento à determinação da CGJUS", "Evento referente à Saúde de Oficiais - Jornada de Saúde e Qualidade de Vida dos Oficiais de Justiça do TJTO - Quem cuida da Justiça precisa ser cuidado!", "Atividade  será realizada por DIGEP/CESAU e Napsi - previsão para o mês de agosto", "a definir", "Jm/ Napsi/ cesau/DIGEP", "25.0.000026494-9 / 26.0.000012890-1"],
    ["Em desenvolvimento", "Pesquisa de interesse em plataforma digital", "Requerimento do Comitê de Saúde", "Atividade será realizada pela DIVPGC", "a definir", "DIVPGC/DIGEP", "26.0.000011896-5"],
    ["A iniciar a execução", "Realização do concurso para analista", "aguardando elaboração", "", "Paula Jorge", "", "26.0.000012074-9"],
    ["A iniciar a execução", "Programa de Preparação para a Aposentadoria", "Relatório do CNJ sobre o PPA", "Aguardando a realização das atividades", "a definir", "DIVPGC/DIGEP", "26.0.000011213-4"],
    ["Em desenvolvimento", "Comitê de Governança das Contratações", "Realização das competencias das unidades de contratações", "Aguardando retorno do GESTCOM", "a definir", "DIVPGC/DIGEP", "26.0.000009294-0"]
  ];

  aba.getRange(2, 1, dados.length, 7).setValues(dados);

  // Formatação zebrada e wrap
  for (var i = 0; i < dados.length; i++) {
    var r = aba.getRange(i + 2, 1, 1, 7);
    r.setWrap(true).setVerticalAlignment('top');
    if (i % 2 === 1) r.setBackground('#f8f9fa');
  }

  // Coloração por status
  var statusCores = {
    'concluído': '#d9ead3',
    'concluido': '#d9ead3',
    'em desenvolvimento': '#dae8fc',
    'em deenvolvimento': '#dae8fc',
    'a iniciar a execução': '#fff2cc',
    'a iniciar a execucao': '#fff2cc',
    'suspenso': '#f4cccc'
  };
  for (var j = 0; j < dados.length; j++) {
    var statusVal = (dados[j][0] || '').toLowerCase().trim()
                   .normalize('NFD').replace(/[̀-ͯ]/g, '');
    var cor = statusCores[statusVal];
    if (cor) aba.getRange(j + 2, 1).setBackground(cor);
  }

  aba.setFrozenRows(1);
  aba.setRowHeights(2, dados.length, 80);
}

function _configurarAbaDashboard(ss) {
  var aba = ss.getSheetByName(ABA_DASHBOARD);
  if (!aba) {
    aba = ss.insertSheet(ABA_DASHBOARD);
  } else {
    aba.clearContents();
    aba.clearFormats();
  }
  aba.getRange(1, 1)
    .setValue('Execute "Atualizar Dashboard" no menu DIVPGC para gerar o painel.')
    .setFontStyle('italic').setFontColor('#888888');
}
