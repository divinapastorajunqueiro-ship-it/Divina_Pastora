/**
 * ============================================================
 *  SISTEMA DE CONTROLE DE FREQUÊNCIA ESCOLAR
 *  Script Google Apps Script – Versão 3.0 (2026)
 *  Escola: Escola Municipal de Educação Básica Nossa Senhora Divina Pastora
 *
 *  COMO USAR:
 *  1. Abra o Google Sheets e crie uma planilha em branco
 *  2. Acesse Extensões > Apps Script
 *  3. Cole todo este código e salve (Ctrl+S)
 *  4. Execute a função: criarPlanilhaCompleta()
 *  5. Após criada, publique como Web App (Implantar > Novo Implantação)
 *     - Tipo: App da Web
 *     - Executar como: Eu
 *     - Quem tem acesso: Qualquer pessoa
 *  6. Cole a URL gerada na aba ⚙️ Config e no App Web
 * ============================================================
 */

// ============================================================
//  HELPERS DE LOCALE – FÓRMULAS pt-BR
//  Converte separadores de argumento (',' -> ';') preservando
//  vírgulas dentro de strings entre aspas. Necessário porque
//  setFormula NÃO traduz automaticamente em planilhas pt-BR
//  com strings literais contendo vírgulas, gerando #ERROR!.
// ============================================================
function _lf(formula) {
  if (typeof formula !== 'string') return formula;
  let out = '', inStr = false;
  for (let i = 0; i < formula.length; i++) {
    const c = formula[i];
    if (c === '"') inStr = !inStr;
    if (!inStr && c === '.' && /\d/.test(formula[i - 1] || '') && /\d/.test(formula[i + 1] || '')) {
      out += ',';
    } else {
      out += (c === ',' && !inStr) ? ';' : c;
    }
  }
  return out;
}
function _lf2(arr) {
  return arr.map(function(row){ return row.map(_lf); });
}

// ============================================================
//  PALETA DE CORES – VISUAL MODERNO 2026
// ============================================================
const COR = {
  AZUL_PRIMARIO:   '#1A73E8',
  AZUL_ESCURO:     '#0D47A1',
  AZUL_CLARO:      '#E8F0FE',
  AZUL_MEDIO:      '#4285F4',
  VERDE_OK:        '#137333',
  VERDE_FUNDO:     '#E6F4EA',
  VERDE_MEDIO:     '#34A853',
  AMARELO_AVISO:   '#F9AB00',
  AMARELO_FUNDO:   '#FEF7E0',
  VERMELHO:        '#C5221F',
  VERMELHO_FUNDO:  '#FCE8E6',
  CINZA_ESCURO:    '#3C4043',
  CINZA_MEDIO:     '#5F6368',
  CINZA_CLARO:     '#F8F9FA',
  CINZA_BORDA:     '#DADCE0',
  BRANCO:          '#FFFFFF',
  HEADER_GRAD:     '#0D47A1',
  LINHA_PAR:       '#F8F9FA',
  LINHA_IMPAR:     '#FFFFFF',
  LARANJA:         '#E37400',
  LARANJA_FUNDO:   '#FEF3E2',
  ROXO:            '#7B1FA2',
  ROXO_FUNDO:      '#F3E5F5',
};

// ============================================================
//  CONSTANTES DO SISTEMA
// ============================================================
const ESCOLA      = 'Escola Municipal de Educação Básica Nossa Senhora Divina Pastora';
const ANO         = 2026;
const MAX_ALUNOS  = 1000;
const DIAS_LETIVOS = 200;
const FREQ_MINIMA  = 0.75;

const TURMAS = [
  // Insira aqui as turmas da sua escola, ex: '6A','6B','7A'
];

const ABAS = [
  { nome: '🗺️ Navegação',       cor: COR.AZUL_PRIMARIO  },
  { nome: '🏠 Início',           cor: COR.AZUL_ESCURO    },
  { nome: '👥 Alunos',           cor: COR.VERDE_MEDIO    },
  { nome: '✅ Frequência',       cor: COR.VERDE_OK       },
  { nome: '🔲 QR Codes',         cor: COR.LARANJA        },
  { nome: '📋 Status',           cor: COR.AZUL_MEDIO     },
  { nome: '📊 Dashboard',        cor: COR.ROXO           },
  { nome: '⚠️ Alertas',         cor: COR.VERMELHO       },
  { nome: '📅 Rel. Mensal',      cor: COR.CINZA_ESCURO   },
  { nome: '📚 Rel. por UD',      cor: COR.CINZA_MEDIO    },
  { nome: '⚙️ Config',          cor: COR.AZUL_PRIMARIO  },
];

// ============================================================
//  FUNÇÃO PRINCIPAL – PONTO DE ENTRADA
// ============================================================
function criarPlanilhaCompleta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remover abas padrão existentes
  _limparAbas(ss);

  // Criar todas as abas na ordem
  const abas = {};
  ABAS.forEach((a, i) => {
    const sheet = (i === 0) ? ss.getSheets()[0] : ss.insertSheet(i);
    sheet.setName(a.nome);
    sheet.setTabColor(a.cor);
    abas[a.nome] = sheet;
  });

  // Preencher cada aba (flush() entre abas pesadas para evitar timeout)
  _criarNavegacao(abas['🗺️ Navegação'], ss);
  _criarInicio(abas['🏠 Início']);
  SpreadsheetApp.flush();

  _criarAlunos(abas['👥 Alunos']);
  SpreadsheetApp.flush();

  _criarFrequencia(abas['✅ Frequência']);
  SpreadsheetApp.flush();

  _criarQRCodes(abas['🔲 QR Codes']);
  SpreadsheetApp.flush();

  _criarStatus(abas['📋 Status']);
  SpreadsheetApp.flush();

  _criarDashboard(abas['📊 Dashboard']);
  SpreadsheetApp.flush();

  _criarAlertas(abas['⚠️ Alertas']);
  SpreadsheetApp.flush();

  _criarRelatorioMensal(abas['📅 Rel. Mensal']);
  SpreadsheetApp.flush();

  _criarRelatorioPorUD(abas['📚 Rel. por UD']);
  SpreadsheetApp.flush();

  _criarConfig(abas['⚙️ Config']);

  // Ativar aba de início
  abas['🏠 Início'].activate();

  SpreadsheetApp.getUi().alert(
    '✅ Planilha criada com sucesso!\n\n' +
    '• 1.000 linhas reservadas por aba de dados\n' +
    '• Visual moderno aplicado\n' +
    '• Configurações de impressão definidas\n\n' +
    'Próximo passo: publique como Web App e cole a URL na aba ⚙️ Config.'
  );
}

/**
 * CRIAÇÃO EM ETAPAS – use estas funções individuais pelo menu
 * caso criarPlanilhaCompleta() ainda expire.
 * Execute-as na ordem: Etapa 1 → 2 → 3
 */
function criarEtapa1_EstruturaeAlunos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _limparAbas(ss);
  const abas = {};
  ABAS.forEach((a, i) => {
    const sheet = (i === 0) ? ss.getSheets()[0] : ss.insertSheet(i);
    sheet.setName(a.nome);
    sheet.setTabColor(a.cor);
    abas[a.nome] = sheet;
  });
  _criarNavegacao(abas['🗺️ Navegação'], ss);
  _criarInicio(abas['🏠 Início']);
  SpreadsheetApp.flush();
  _criarAlunos(abas['👥 Alunos']);
  SpreadsheetApp.flush();
  _criarFrequencia(abas['✅ Frequência']);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('✅ Etapa 1/3 concluída! Execute agora: Etapa 2 – QR Codes e Status.');
}

function criarEtapa2_QRCodesStatusDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {};
  ss.getSheets().forEach(s => { sheets[s.getName()] = s; });
  _criarQRCodes(sheets['🔲 QR Codes']);
  SpreadsheetApp.flush();
  _criarStatus(sheets['📋 Status']);
  SpreadsheetApp.flush();
  _criarDashboard(sheets['📊 Dashboard']);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('✅ Etapa 2/3 concluída! Execute agora: Etapa 3 – Alertas e Relatórios.');
}

function criarEtapa3_AlertasRelatoriosConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {};
  ss.getSheets().forEach(s => { sheets[s.getName()] = s; });
  _criarAlertas(sheets['⚠️ Alertas']);
  SpreadsheetApp.flush();
  _criarRelatorioMensal(sheets['📅 Rel. Mensal']);
  SpreadsheetApp.flush();
  _criarRelatorioPorUD(sheets['📚 Rel. por UD']);
  SpreadsheetApp.flush();
  _criarConfig(sheets['⚙️ Config']);
  sheets['🏠 Início'].activate();
  SpreadsheetApp.getUi().alert('✅ Planilha completa! Todas as abas foram criadas com sucesso.');
}

// ============================================================
//  HELPER – LIMPAR ABAS EXISTENTES
// ============================================================
function _limparAbas(ss) {
  const sheets = ss.getSheets();
  // Limpa completamente a primeira aba (incluindo merges e formatação)
  const first = sheets[0];
  first.getRange(1, 1, first.getMaxRows(), first.getMaxColumns()).breakApart();
  first.clearContents();
  first.clearFormats();
  first.setName('_temp');
  // Exclui as demais
  for (let i = sheets.length - 1; i >= 1; i--) {
    ss.deleteSheet(sheets[i]);
  }
}

// ============================================================
//  HELPERS DE ESTILO
// ============================================================
function _getNomeEscola() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cfg = ss.getSheetByName('⚙️ Config');
    if (cfg) {
      // Nome da escola está em B6 (primeira linha de dados da seção CONFIGURAÇÃO GERAL)
      const val = cfg.getRange('B6').getValue();
      if (val && String(val).trim() !== '') return String(val).trim();
    }
  } catch(e) {}
  return ESCOLA; // fallback para a constante padrão
}

function _cabecalhoPrincipal(sheet, titulo, subtitulo, numCols, incluirBotaoVoltar) {
  numCols = numCols || 9;
  if (incluirBotaoVoltar === undefined) incluirBotaoVoltar = true;

  // Desfaz merges existentes nas linhas do cabeçalho antes de recriar
  sheet.getRange(1, 1, 3, numCols).breakApart();

  // Lê o nome da escola dinamicamente da aba Config (ou usa a constante padrão)
  const nomeEscola = _getNomeEscola();

  // Linha 1 – nome da escola (+ link "Voltar à Navegação" quando aplicável)
  const sepBotao = '     🗺️ ← Voltar à Navegação';
  const textoL1 = incluirBotaoVoltar ? (nomeEscola + sepBotao) : nomeEscola;
  const rtBuilder = SpreadsheetApp.newRichTextValue()
    .setText(textoL1)
    .setTextStyle(0, nomeEscola.length,
      SpreadsheetApp.newTextStyle()
        .setFontFamily('Google Sans').setFontSize(11)
        .setForegroundColor(COR.BRANCO).setBold(true).build());
  if (incluirBotaoVoltar) {
    const ss  = sheet.getParent();
    const nav = ss.getSheetByName('🗺️ Navegação');
    if (nav) {
      const url = '#gid=' + nav.getSheetId();
      rtBuilder.setTextStyle(nomeEscola.length, textoL1.length,
        SpreadsheetApp.newTextStyle()
          .setFontFamily('Google Sans').setFontSize(11)
          .setForegroundColor('#FFEB3B').setBold(true)
          .setUnderline(true).build());
      rtBuilder.setLinkUrl(nomeEscola.length, textoL1.length, url);
    }
  }
  sheet.getRange(1,1).setRichTextValue(rtBuilder.build());
  sheet.getRange(1,1,1,numCols).merge()
    .setBackground(COR.HEADER_GRAD)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 32);

  // Linha 2 – título da aba
  sheet.getRange(2,1).setValue(titulo)
    .setFontFamily('Google Sans')
    .setFontSize(16)
    .setFontColor(COR.BRANCO)
    .setBackground(COR.AZUL_PRIMARIO)
    .setFontWeight('bold');
  sheet.getRange(2,1,1,numCols).merge()
    .setBackground(COR.AZUL_PRIMARIO)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 40);

  // Linha 3 – subtítulo / info
  if (subtitulo) {
    sheet.getRange(3,1).setValue(subtitulo)
      .setFontFamily('Google Sans')
      .setFontSize(10)
      .setFontColor(COR.CINZA_ESCURO)
      .setBackground(COR.AZUL_CLARO)
      .setFontStyle('italic');
    sheet.getRange(3,1,1,numCols).merge()
      .setBackground(COR.AZUL_CLARO)
      .setHorizontalAlignment('center');
    sheet.setRowHeight(3, 24);
  }
}

function _cabecalhoTabela(sheet, row, colunas) {
  const range = sheet.getRange(row, 1, 1, colunas.length);
  range.setValues([colunas])
    .setBackground(COR.AZUL_ESCURO)
    .setFontColor(COR.BRANCO)
    .setFontWeight('bold')
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(row, 36);
}

function _zebraStripes(sheet, startRow, endRow, numCols) {
  // OTIMIZADO: uma única chamada de API com array 2D de cores (batch)
  const numRows = endRow - startRow + 1;
  const bgs = [];
  const heights = [];
  for (let r = 0; r < numRows; r++) {
    const bg = ((startRow + r) % 2 === 0) ? COR.LINHA_PAR : COR.LINHA_IMPAR;
    bgs.push(Array(numCols).fill(bg));
    heights.push(22);
  }
  sheet.getRange(startRow, 1, numRows, numCols).setBackgrounds(bgs);
  // setRowHeights em lote
  sheet.setRowHeightsForced(startRow, numRows, 22);
}

function _bordaTabela(sheet, startRow, endRow, numCols) {
  sheet.getRange(startRow, 1, endRow - startRow + 1, numCols)
    .setBorder(true, true, true, true, true, true,
      COR.CINZA_BORDA, SpreadsheetApp.BorderStyle.SOLID);
}

function _configurarImpressao(sheet, orientation, fitToPage) {
  try {
    const ps = sheet.getPageSetup();
    ps.setPaperSize(SpreadsheetApp.PaperSize.A4);
    if (orientation === 'L') {
      ps.setOrientation(SpreadsheetApp.PageOrientation.LANDSCAPE);
    } else {
      ps.setOrientation(SpreadsheetApp.PageOrientation.PORTRAIT);
    }
    if (fitToPage !== false) ps.setFitToPage(true);
    ps.setPrintGridlines(false);
    ps.setPageOrder(SpreadsheetApp.PageOrder.DOWN_THEN_OVER);
    ps.setCenterOnPage(true, false);
    sheet.setPageSetup(ps);
  } catch(e) {
    // pageSetup não disponível neste contexto — ignora silenciosamente
  }
}

// ============================================================
//  ABA: 🗺️ NAVEGAÇÃO
// ============================================================
function _criarNavegacao(sheet, ss) {
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).breakApart();
  sheet.clear();
  sheet.setColumnWidth(1, 320);
  sheet.setColumnWidth(2, 460);
  sheet.setColumnWidth(3, 180);
  for (let c = 4; c <= 10; c++) sheet.setColumnWidth(c, 80);

  // Header (sem botão voltar na própria Navegação)
  _cabecalhoPrincipal(sheet, '🗺️ NAVEGAÇÃO DO SISTEMA',
    'Clique no nome do módulo para abrir a aba correspondente', 3, false);

  // Linha 4 – espaço
  sheet.setRowHeight(4, 16);

  // Título seção
  sheet.getRange(5,1).setValue('MÓDULOS DO SISTEMA');
  sheet.getRange(5,1,1,3).merge()
    .setBackground(COR.AZUL_ESCURO)
    .setFontColor(COR.BRANCO)
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontFamily('Google Sans')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(5, 30);

  // Cabeçalho da tabela
  _cabecalhoTabela(sheet, 6, ['Aba / Módulo', 'Descrição', 'Tipo']);

  const modulos = [
    ['🏠 Início',       'Tela de boas-vindas, versão e guia rápido',              'Informativo'],
    ['👥 Alunos',       'Cadastro completo dos alunos (até 1.000 registros)',     'Dados'],
    ['✅ Frequência',   'Registros diários enviados pelo App (QR Code)',          'Dados (App)'],
    ['🔲 QR Codes',     'Links para impressão e distribuição dos cartões QR',     'Impressão'],
    ['📋 Status',       'Status diário de presença por aluno',                    'Consulta'],
    ['📊 Dashboard',    'Painel de controle com totais e métricas por turma',     'Relatório'],
    ['⚠️ Alertas',     'Alunos com frequência abaixo de 75% (LDB Art. 24)',     'Alerta'],
    ['📅 Rel. Mensal',  'Relatório mensal de frequência por aluno',               'Relatório'],
    ['📚 Rel. por UD',  'Relatório por Unidade Didática / Bimestre',              'Relatório'],
    ['⚙️ Config',      'Configurações da escola, integração e turnos',           'Sistema'],
  ];

  const bgTipos = {
    'Informativo': COR.AZUL_CLARO,
    'Dados':       COR.VERDE_FUNDO,
    'Dados (App)': COR.VERDE_FUNDO,
    'Impressão':   COR.LARANJA_FUNDO,
    'Consulta':    COR.AZUL_CLARO,
    'Relatório':   COR.ROXO_FUNDO,
    'Alerta':      COR.VERMELHO_FUNDO,
    'Sistema':     COR.AMARELO_FUNDO,
  };

  modulos.forEach(([nome, desc, tipo], i) => {
    const r = 7 + i;
    const bg = (i % 2 === 0) ? COR.LINHA_IMPAR : COR.LINHA_PAR;

    // Coluna 1: nome do módulo como hyperlink para a respectiva aba
    const destino = ss.getSheetByName(nome);
    if (destino) {
      const url = '#gid=' + destino.getSheetId();
      const rt = SpreadsheetApp.newRichTextValue()
        .setText(nome)
        .setLinkUrl(url)
        .setTextStyle(SpreadsheetApp.newTextStyle()
          .setFontFamily('Google Sans').setFontSize(11)
          .setBold(true).setUnderline(true)
          .setForegroundColor(COR.AZUL_PRIMARIO).build())
        .build();
      sheet.getRange(r, 1).setRichTextValue(rt);
    } else {
      sheet.getRange(r, 1).setValue(nome)
        .setFontWeight('bold').setFontColor(COR.AZUL_PRIMARIO)
        .setFontSize(11).setFontFamily('Google Sans');
    }

    sheet.getRange(r, 2).setValue(desc)
      .setFontColor(COR.CINZA_ESCURO).setFontSize(10);
    sheet.getRange(r, 3).setValue(tipo)
      .setBackground(bgTipos[tipo] || bg)
      .setFontSize(9).setHorizontalAlignment('center')
      .setFontWeight('bold');
    sheet.getRange(r, 1, 1, 2).setBackground(bg);
    sheet.setRowHeight(r, 28);
  });

  _bordaTabela(sheet, 6, 6 + modulos.length, 3);

  // Nota de rodapé
  const nr = 8 + modulos.length;
  sheet.getRange(nr,1).setValue(
    '⚠️  NÃO edite diretamente as abas Frequência e Status — use o App Web para registrar presenças.'
  );
  sheet.getRange(nr,1,1,3).merge()
    .setBackground(COR.AMARELO_FUNDO)
    .setFontColor(COR.LARANJA)
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(nr, 30);

  // URL do App
  const nr2 = nr + 1;
  sheet.getRange(nr2,1).setValue('🌐 App Web (Portaria QR):');
  sheet.getRange(nr2,2).setValue('https://divinapastorajunqueiro-ship-it.github.io/Divina_Pastora/');
  sheet.getRange(nr2,2,1,2).merge()
    .setFontColor(COR.AZUL_PRIMARIO)
    .setFontWeight('bold');
  sheet.getRange(nr2,1,1,3).setBackground(COR.AZUL_CLARO);
  sheet.setRowHeight(nr2, 26);

  _configurarImpressao(sheet, 'P', true);
}

// ============================================================
//  ABA: 🏠 INÍCIO
// ============================================================
function _criarInicio(sheet) {
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).breakApart();
  sheet.clear();
  for (let c = 1; c <= 8; c++) sheet.setColumnWidth(c, c===1?60:c===2?220:c===3?180:120);

  _cabecalhoPrincipal(sheet,
    '📚 SISTEMA DE CONTROLE DE FREQUÊNCIA ESCOLAR',
    `com QR Code + Google Sheets  ·  Versão 3.0  ·  ${ANO}  ·  Para escolas públicas  ·  Uso gratuito`,
    4
  );

  sheet.setRowHeight(4, 14);

  // Box versão
  const boxData = [
    ['📌 SOBRE O SISTEMA', ''],
    ['Versão:', '3.0 – 2026'],
    ['Escola:', _getNomeEscola()],
    ['Total de Alunos:', `Até ${MAX_ALUNOS.toLocaleString('pt-BR')} registros`],
    ['Dias Letivos:', `${DIAS_LETIVOS} dias (configurável)`],
    ['Frequência Mínima:', `${(FREQ_MINIMA*100)}% (LDB Art. 24)`],
    ['App Web:', 'https://divinapastorajunqueiro-ship-it.github.io/Divina_Pastora/'],
  ];

  boxData.forEach(([label, val], i) => {
    const r = 5 + i;
    if (i === 0) {
      sheet.getRange(r,1).setValue(label);
      sheet.getRange(r,1,1,4).merge()
        .setBackground(COR.AZUL_PRIMARIO).setFontColor(COR.BRANCO)
        .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
      sheet.setRowHeight(r, 30);
    } else {
      sheet.getRange(r,1).setValue(label)
        .setFontWeight('bold').setFontColor(COR.CINZA_ESCURO)
        .setBackground(COR.CINZA_CLARO);
      sheet.getRange(r,2).setValue(val)
        .setFontColor(i===7 ? COR.AZUL_PRIMARIO : COR.CINZA_ESCURO);
      // (largura de coluna definida no início)
      sheet.setRowHeight(r, 24);
    }
  });

  _bordaTabela(sheet, 5, 5 + boxData.length - 1, 4);
  sheet.setRowHeight(5 + boxData.length, 14);

  // Guia de módulos
  const r0 = 5 + boxData.length + 1;
  sheet.getRange(r0,1).setValue('📋 GUIA RÁPIDO DAS ABAS');
  sheet.getRange(r0,1,1,4).merge()
    .setBackground(COR.VERDE_MEDIO).setFontColor(COR.BRANCO)
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  sheet.setRowHeight(r0, 30);

  const guia = [
    ['🗺️ Navegação',   'Links de acesso rápido a todos os módulos'],
    ['👥 Alunos',      'Cadastro e gestão dos alunos'],
    ['✅ Frequência',  'Registros diários (preenchido pelo App)'],
    ['🔲 QR Codes',    'Links para impressão dos cartões'],
    ['📋 Status',      'Status de presença diário'],
    ['📊 Dashboard',   'Painel de controle com totais automáticos'],
    ['⚠️ Alertas',    'Alunos abaixo de 75% de frequência'],
    ['📅 Rel. Mensal', 'Relatório mensal por aluno'],
    ['📚 Rel. por UD', 'Relatório por Unidade Didática (Bimestre)'],
    ['⚙️ Config',     'Configurações da escola e integração'],
  ];

  guia.forEach(([aba, desc], i) => {
    const r = r0 + 1 + i;
    const bg = i % 2 === 0 ? COR.LINHA_IMPAR : COR.LINHA_PAR;
    sheet.getRange(r,1).setValue(aba).setFontWeight('bold')
      .setFontColor(COR.AZUL_PRIMARIO).setBackground(bg);
    sheet.getRange(r,2).setValue(desc)
      .setFontColor(COR.CINZA_ESCURO).setBackground(bg);
    sheet.getRange(r,1,1,2).setFontSize(10);
    sheet.setRowHeight(r, 24);
  });

  _bordaTabela(sheet, r0, r0 + guia.length, 4);

  // Aviso
  const rv = r0 + guia.length + 2;
  sheet.getRange(rv,1).setValue(
    '⚠️  NÃO edite diretamente as abas Frequência e Status — use o App Web.'
  );
  sheet.getRange(rv,1,1,4).merge()
    .setBackground(COR.VERMELHO_FUNDO).setFontColor(COR.VERMELHO)
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.setRowHeight(rv, 30);

  _configurarImpressao(sheet, 'P', true);
}

// ============================================================
//  ABA: 👥 ALUNOS
// ============================================================
function _criarAlunos(sheet) {
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).breakApart();
  sheet.clear();

  const larguras = [30, 200, 130, 110, 80, 100, 200, 130, 200];
  larguras.forEach((w,i) => sheet.setColumnWidth(i+1, w));

  _cabecalhoPrincipal(sheet, '👥 CADASTRO DE ALUNOS',
    `Capacidade: ${MAX_ALUNOS.toLocaleString('pt-BR')} alunos  ·  Use validação de dados para Turma e Situação`, 10);

  // Contador de alunos ativos (linha 3 – extra info)
  sheet.getRange(4,1).setValue('Total de alunos ativos:');
  // CORRIGIDO: dados iniciam em B6 (B5 é cabeçalho). Contava cabeçalho e perdia última linha.
  sheet.getRange(4,2).setValue(_lf('=COUNTA(B6:B1005)'));
  sheet.getRange(4,1,1,9).setBackground(COR.AZUL_CLARO).setFontSize(10);
  sheet.getRange(4,1).setFontWeight('bold');
  sheet.getRange(4,2).setFontWeight('bold').setFontColor(COR.AZUL_PRIMARIO);
  sheet.setRowHeight(4, 26);

  const cols = ['#','Matrícula','Nome Completo','Turma','Situação','Turno','Data Nasc.','Responsável','Telefone','Observação'];
  _cabecalhoTabela(sheet, 5, cols);

  // Validação de turma removida — preencha livremente na coluna Turma

  const sitRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Ativo','Transferido','Evadido','Falecido'], true)
    .setAllowInvalid(false).build();
  sheet.getRange(6, 5, MAX_ALUNOS, 1).setDataValidation(sitRule);

  const turnoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Matutino','Vespertino','Noturno','Integral'], true)
    .setAllowInvalid(false).build();
  sheet.getRange(6, 6, MAX_ALUNOS, 1).setDataValidation(turnoRule);

  // Numeração automática – OTIMIZADO: um único setValues com array
  const nums = Array.from({length: MAX_ALUNOS}, (_, i) => [i + 1]);
  sheet.getRange(6, 1, MAX_ALUNOS, 1).setValues(nums);
  sheet.getRange(6, 1, MAX_ALUNOS, 1)
    .setFontColor(COR.CINZA_MEDIO).setFontSize(9).setHorizontalAlignment('center');

  _zebraStripes(sheet, 6, 5 + MAX_ALUNOS, cols.length);
  _bordaTabela(sheet, 5, 5 + MAX_ALUNOS, cols.length);

  // Congelar cabeçalho
  sheet.setFrozenRows(5);

  // Formatação de data
  sheet.getRange(6, 7, MAX_ALUNOS, 1).setNumberFormat('dd/mm/yyyy');

  _configurarImpressao(sheet, 'L', false);
}

// ============================================================
//  ABA: ✅ FREQUÊNCIA
// ============================================================
function _criarFrequencia(sheet) {
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).breakApart();
  sheet.clear();

  const larguras = [130, 200, 120, 100, 80, 100, 100, 130, 180];
  larguras.forEach((w,i) => sheet.setColumnWidth(i+1, w));

  _cabecalhoPrincipal(sheet, '✅ REGISTROS DE FREQUÊNCIA',
    '⚠️ Aba preenchida automaticamente pelo App Web — NÃO editar manualmente', 9);

  // Contador de registros hoje
  sheet.getRange(4,1).setValue('Registros hoje:');
  sheet.getRange(4,2).setValue(_lf('=IFERROR(COUNTIFS(D6:D' + (5+MAX_ALUNOS*5) + ',">="&INT(TODAY()),D6:D' + (5+MAX_ALUNOS*5) + ',"<"&INT(TODAY())+1),0)'));
  sheet.getRange(4,1,1,9).setBackground(COR.VERDE_FUNDO).setFontSize(10);
  sheet.getRange(4,1).setFontWeight('bold');
  sheet.getRange(4,2).setFontWeight('bold').setFontColor(COR.VERDE_OK);
  sheet.setRowHeight(4, 26);

  const cols = ['Matrícula','Nome','Turma','Data','Hora','Turno','Status','Registrado Via','Timestamp'];
  _cabecalhoTabela(sheet, 5, cols);

  _zebraStripes(sheet, 6, 5 + MAX_ALUNOS * 5, cols.length);
  _bordaTabela(sheet, 5, 5 + MAX_ALUNOS * 5, cols.length);

  // Formatação
  sheet.getRange(6, 4, MAX_ALUNOS * 5, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(6, 5, MAX_ALUNOS * 5, 1).setNumberFormat('hh:mm:ss');
  sheet.getRange(6, 9, MAX_ALUNOS * 5, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');

  sheet.setFrozenRows(5);

  // Proteção
  const protection = sheet.protect().setDescription('Aba protegida – use o App Web');
  protection.setWarningOnly(true);

  _configurarImpressao(sheet, 'L', false);
}

// ============================================================
//  ABA: 🔲 QR CODES
// ============================================================
function _criarQRCodes(sheet) {
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).breakApart();
  sheet.clear();

  const larguras = [130, 220, 120, 400, 120];
  larguras.forEach((w,i) => sheet.setColumnWidth(i+1, w));

  _cabecalhoPrincipal(sheet, '🔲 QR CODES – IMPRESSÃO E DISTRIBUIÇÃO',
    'Imprima e distribua os cartões QR Code para cada aluno. URL gerada automaticamente via fórmula.', 5);

  const cols = ['Matrícula','Nome Completo','Turma','URL do QR Code','Status'];
  _cabecalhoTabela(sheet, 4, cols);

  // Fórmulas de lookup na aba Alunos + geração URL QR – OTIMIZADO: batch por coluna
  // Nota: ENCODEURL substituído por SUBSTITUTE para compatibilidade com locale PT-BR
  const BASE_URL = "https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=";
  const fMatricula = [], fNome = [], fTurma = [], fQR = [], fStatus = [];
  for (let i = 0; i < MAX_ALUNOS; i++) {
    const r = 5 + i;
    const alunoRow = 6 + i;
    fMatricula.push([`='👥 Alunos'!B${alunoRow}`]);
    fNome.push([`='👥 Alunos'!C${alunoRow}`]);
    fTurma.push([`='👥 Alunos'!D${alunoRow}`]);
    // Usa SUBSTITUTE para codificar espaços (%20) e pipe (%7C) sem depender de ENCODEURL
    fQR.push([`=IF(A${r}="","",HYPERLINK("${BASE_URL}"&SUBSTITUTE(SUBSTITUTE(A${r}&"%7C"&B${r}&"%7C"&C${r}," ","%20"),"&","%26"),"🔗 Abrir QR Code"))`]);
    fStatus.push([`=IF(A${r}="","",IF('👥 Alunos'!E${alunoRow}="Ativo","✅ Ativo","⚫ Inativo"))`]);
  }
  sheet.getRange(5, 1, MAX_ALUNOS, 1).setValues(_lf2(fMatricula));
  sheet.getRange(5, 2, MAX_ALUNOS, 1).setValues(_lf2(fNome));
  sheet.getRange(5, 3, MAX_ALUNOS, 1).setValues(_lf2(fTurma));
  sheet.getRange(5, 4, MAX_ALUNOS, 1).setValues(_lf2(fQR));
  sheet.getRange(5, 5, MAX_ALUNOS, 1).setValues(_lf2(fStatus));

  _zebraStripes(sheet, 5, 4 + MAX_ALUNOS, cols.length);
  _bordaTabela(sheet, 4, 4 + MAX_ALUNOS, cols.length);

  // Formatação condicional – Status Ativo
  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Ativo')
    .setBackground(COR.VERDE_FUNDO)
    .setFontColor(COR.VERDE_OK)
    .setRanges([sheet.getRange(5, 5, MAX_ALUNOS, 1)])
    .build());
  sheet.setConditionalFormatRules(rules);

  sheet.setFrozenRows(4);
  _configurarImpressao(sheet, 'L', false);
}

// ============================================================
//  ABA: 📋 STATUS DIÁRIO
// ============================================================
function _criarStatus(sheet) {
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).breakApart();
  sheet.clear();

  const larguras = [130, 200, 120, 100, 110, 110, 120];
  larguras.forEach((w,i) => sheet.setColumnWidth(i+1, w));

  _cabecalhoPrincipal(sheet, '📋 STATUS DIÁRIO DE PRESENÇA',
    'Consulta automática do status de cada aluno no dia de hoje', 7);

  // Data de consulta
  sheet.getRange(4,1).setValue('Data de consulta:');
  sheet.getRange(4,2).setValue('=TODAY()').setNumberFormat('dddd, dd/mm/yyyy');
  sheet.getRange(4,1,1,7).setBackground(COR.AZUL_CLARO);
  sheet.getRange(4,1).setFontWeight('bold');
  sheet.getRange(4,2).setFontWeight('bold').setFontColor(COR.AZUL_PRIMARIO);
  sheet.setRowHeight(4, 26);

  const cols = ['Matrícula','Nome','Turma','Turno','Hora Limite','Hora da Leitura','Status'];
  _cabecalhoTabela(sheet, 5, cols);

  // OTIMIZADO: construir arrays de fórmulas e aplicar em batch por coluna
  const sF1=[], sF2=[], sF3=[], sF4=[], sF5=[], sF6=[], sF7=[];
  for (let i = 0; i < MAX_ALUNOS; i++) {
    const r = 6 + i;
    const aRow = 6 + i;
    sF1.push([`='👥 Alunos'!B${aRow}`]);
    sF2.push([`='👥 Alunos'!C${aRow}`]);
    sF3.push([`='👥 Alunos'!D${aRow}`]);
    sF4.push([`='👥 Alunos'!F${aRow}`]);
    sF5.push([`=IF(D${r}="Matutino",TIME(12,0,0),IF(D${r}="Vespertino",TIME(18,0,0),IF(D${r}="Noturno",TIME(23,59,0),TIME(12,0,0))))`]);
    // CORRIGIDO: MAXIFS retorna 0 (não erro) quando não há registro — 0 != "" causava falso "✅ No Horário".
    // Usamos COUNTIFS para verificar a existência do registro antes de retornar a hora.
    // Se não houver registro → retorna "" → Status marca "❌ Ausente" corretamente.
    sF6.push([`=IF(COUNTIFS('✅ Frequência'!A:A,A${r},'✅ Frequência'!D:D,">="&INT(TODAY()),'✅ Frequência'!D:D,"<"&INT(TODAY())+1)=0,"",IFERROR(MAXIFS('✅ Frequência'!E:E,'✅ Frequência'!A:A,A${r},'✅ Frequência'!D:D,">="&INT(TODAY()),'✅ Frequência'!D:D,"<"&INT(TODAY())+1),""))`]);
    sF7.push([`=IF(A${r}="","",IF(F${r}="","❌ Ausente",IF(F${r}<=E${r},"✅ No Horário","⚠️ Atrasado")))`]);
  }
  sheet.getRange(6,1,MAX_ALUNOS,1).setValues(_lf2(sF1));
  sheet.getRange(6,2,MAX_ALUNOS,1).setValues(_lf2(sF2));
  sheet.getRange(6,3,MAX_ALUNOS,1).setValues(_lf2(sF3));
  sheet.getRange(6,4,MAX_ALUNOS,1).setValues(_lf2(sF4));
  sheet.getRange(6,5,MAX_ALUNOS,1).setValues(_lf2(sF5));
  sheet.getRange(6,6,MAX_ALUNOS,1).setValues(_lf2(sF6));
  sheet.getRange(6,7,MAX_ALUNOS,1).setValues(_lf2(sF7));

  sheet.getRange(6, 5, MAX_ALUNOS, 2).setNumberFormat('hh:mm:ss');

  // Formatação condicional
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('No Horário').setBackground(COR.VERDE_FUNDO).setFontColor(COR.VERDE_OK)
      .setRanges([sheet.getRange(6, 7, MAX_ALUNOS, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Ausente').setBackground(COR.VERMELHO_FUNDO).setFontColor(COR.VERMELHO)
      .setRanges([sheet.getRange(6, 7, MAX_ALUNOS, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Atrasado').setBackground(COR.AMARELO_FUNDO).setFontColor(COR.LARANJA)
      .setRanges([sheet.getRange(6, 7, MAX_ALUNOS, 1)]).build(),
  ];
  sheet.setConditionalFormatRules(rules);

  _zebraStripes(sheet, 6, 5 + MAX_ALUNOS, cols.length);
  _bordaTabela(sheet, 5, 5 + MAX_ALUNOS, cols.length);
  sheet.setFrozenRows(5);
  _configurarImpressao(sheet, 'L', false);
}

// ============================================================
//  ABA: 📊 DASHBOARD
// ============================================================
function _criarDashboard(sheet) {
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).breakApart();
  sheet.clear();

  const larguras = [160, 140, 140, 140, 140, 140, 120];
  larguras.forEach((w,i) => sheet.setColumnWidth(i+1, w));

  _cabecalhoPrincipal(sheet, '📊 PAINEL DE CONTROLE – FREQUÊNCIA ESCOLAR',
    'Atualiza automaticamente a cada acesso · Dados calculados em tempo real', 8);

  // ---- KPIs (linha 4-6) ----
  const kpis = [
    { label: '👥 TOTAL DE ALUNOS',   formula: `=IFERROR(COUNTA('👥 Alunos'!B6:B1005),0)`,                                                                    col: 1, cor: COR.AZUL_PRIMARIO  },
    { label: '✅ PRESENTES HOJE',    formula: `=IFERROR(COUNTIFS('✅ Frequência'!D:D,">="&INT(TODAY()),'✅ Frequência'!D:D,"<"&INT(TODAY())+1,'✅ Frequência'!G:G,"Presente"),0)`, col: 3, cor: COR.VERDE_MEDIO    },
    { label: '❌ AUSENTES HOJE',     formula: '=IFERROR(A5-C5,0)',                                                                                              col: 5, cor: COR.VERMELHO       },
    { label: '📈 % PRESENÇA HOJE',   formula: '=IFERROR(C5/A5,0)',                                                                                             col: 7, cor: COR.LARANJA        },
  ];

  kpis.forEach(kpi => {
    sheet.getRange(4, kpi.col).setValue(kpi.label)
      .setBackground(kpi.cor).setFontColor(COR.BRANCO)
      .setFontWeight('bold').setFontSize(10).setFontFamily('Google Sans')
      .setHorizontalAlignment('center');
    sheet.getRange(4, kpi.col, 1, 2).merge().setBackground(kpi.cor)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');

    sheet.getRange(5, kpi.col).setValue(_lf(kpi.formula))
      .setFontSize(24).setFontWeight('bold').setFontColor(kpi.cor)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    if (kpi.label.includes('%')) {
      sheet.getRange(5, kpi.col).setNumberFormat('0.0%');
    }
    sheet.getRange(5, kpi.col, 1, 2).merge().setVerticalAlignment('middle');
    sheet.setRowHeight(4, 28);
    sheet.setRowHeight(5, 52);
  });

  // Info datas
  sheet.getRange(5, 9).setValue('Data:');
  sheet.getRange(5, 10).setValue('=TODAY()').setNumberFormat('dd/mm/yyyy');
  sheet.getRange(5, 11).setValue('Atualização:');
  sheet.getRange(5, 12).setValue('=NOW()').setNumberFormat('dd/mm/yyyy hh:mm');

  // Espaço
  sheet.setRowHeight(6, 14);

  // ---- Tabela por Turma ----
  _cabecalhoTabela(sheet, 7,
    ['Turma','Total Alunos','Presentes Hoje','Ausentes','% Presença','Situação']);

  // Lê turmas dinamicamente da aba Alunos (coluna D, a partir da linha 6)
  const abaAlunos = sheet.getParent().getSheetByName('👥 Alunos');
  let turmasUnicas = [];
  if (abaAlunos) {
    const vals = abaAlunos.getRange(6, 4, MAX_ALUNOS, 1).getValues();
    const setTurmas = new Set();
    vals.forEach(row => {
      const t = String(row[0]).trim();
      if (t && t !== '' && t !== 'undefined') setTurmas.add(t);
    });
    turmasUnicas = Array.from(setTurmas).sort();
  }

  const listaTurmas = turmasUnicas.length > 0 ? turmasUnicas : [];
  let rt = 8; // linha de total — se não houver turmas, fica direto na 8

  if (listaTurmas.length > 0) {
    listaTurmas.forEach((turma, i) => {
      const r = 8 + i;
      const bg = i % 2 === 0 ? COR.LINHA_IMPAR : COR.LINHA_PAR;
      sheet.getRange(r, 1).setValue(turma).setFontWeight('bold').setBackground(bg);
      sheet.getRange(r, 2).setValue(_lf(`=IFERROR(COUNTIFS('👥 Alunos'!D:D,"${turma}",'👥 Alunos'!E:E,"Ativo"),0)`)).setBackground(bg);
      sheet.getRange(r, 3).setValue(_lf(
        `=IFERROR(COUNTIFS('✅ Frequência'!C:C,"${turma}",'✅ Frequência'!D:D,">="&INT(TODAY()),'✅ Frequência'!D:D,"<"&INT(TODAY())+1,'✅ Frequência'!G:G,"Presente"),0)`
      )).setBackground(bg);
      sheet.getRange(r, 4).setValue(_lf(`=IFERROR(B${r}-C${r},0)`)).setBackground(bg);
      sheet.getRange(r, 5).setValue(_lf(`=IFERROR(C${r}/B${r},0)`))
        .setNumberFormat('0.0%').setBackground(bg);
      sheet.getRange(r, 6).setValue(_lf(
        `=IF(B${r}=0,"—",IF(E${r}>=0.85,"🟢 Normal",IF(E${r}>=0.75,"🟡 Atenção","🔴 Crítico")))`
      )).setBackground(bg);
      sheet.setRowHeight(r, 24);
    });
    rt = 8 + listaTurmas.length;
  }

  // Linha de TOTAL GERAL
  sheet.getRange(rt, 1).setValue('TOTAL GERAL').setFontWeight('bold')
    .setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
  if (listaTurmas.length > 0) {
    sheet.getRange(rt, 2).setValue(_lf(`=SUM(B8:B${rt-1})`)).setFontWeight('bold')
      .setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
    sheet.getRange(rt, 3).setValue(_lf(`=SUM(C8:C${rt-1})`)).setFontWeight('bold')
      .setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
    sheet.getRange(rt, 4).setValue(_lf(`=SUM(D8:D${rt-1})`)).setFontWeight('bold')
      .setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
    sheet.getRange(rt, 5).setValue(_lf(`=IFERROR(C${rt}/B${rt},0)`))
      .setNumberFormat('0.0%').setFontWeight('bold')
      .setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
  } else {
    // Sem turmas — totais direto das abas-fonte
    sheet.getRange(rt, 2).setValue(_lf(`=IFERROR(COUNTIFS('👥 Alunos'!E:E,"Ativo"),0)`))
      .setFontWeight('bold').setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
    sheet.getRange(rt, 3).setValue(_lf(
      `=IFERROR(COUNTIFS('✅ Frequência'!D:D,">="&INT(TODAY()),'✅ Frequência'!D:D,"<"&INT(TODAY())+1,'✅ Frequência'!G:G,"Presente"),0)`
    )).setFontWeight('bold').setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
    sheet.getRange(rt, 4).setValue(_lf(`=IFERROR(B${rt}-C${rt},0)`))
      .setFontWeight('bold').setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
    sheet.getRange(rt, 5).setValue(_lf(`=IFERROR(C${rt}/B${rt},0)`))
      .setNumberFormat('0.0%').setFontWeight('bold')
      .setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
  }
  sheet.getRange(rt, 6).setValue('').setBackground(COR.AZUL_ESCURO);
  sheet.setRowHeight(rt, 28);

  _bordaTabela(sheet, 7, rt, 6);

  // Formatação condicional % presença
  const numLinhasFC = Math.max(listaTurmas.length, 1);
  const rulesFreq = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0.85).setBackground(COR.VERDE_FUNDO).setFontColor(COR.VERDE_OK)
      .setRanges([sheet.getRange(8, 5, numLinhasFC, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.75, 0.849).setBackground(COR.AMARELO_FUNDO).setFontColor(COR.LARANJA)
      .setRanges([sheet.getRange(8, 5, numLinhasFC, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.75).setBackground(COR.VERMELHO_FUNDO).setFontColor(COR.VERMELHO)
      .setRanges([sheet.getRange(8, 5, numLinhasFC, 1)]).build(),
  ];
  sheet.setConditionalFormatRules(rulesFreq);

  sheet.setFrozenRows(7);
  _configurarImpressao(sheet, 'L', true);
}

// ============================================================
//  ABA: ⚠️ ALERTAS
// ============================================================
function _criarAlertas(sheet) {
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).breakApart();
  sheet.clear();

  const larguras = [130, 200, 120, 130, 140, 120, 130, 280];
  larguras.forEach((w,i) => sheet.setColumnWidth(i+1, w));

  _cabecalhoPrincipal(sheet, '⚠️ ALERTAS DE FREQUÊNCIA – ALUNOS EM RISCO',
    `Critério: abaixo de ${FREQ_MINIMA*100}% de frequência  ·  LDB Art. 24`, 8);

  // Legenda
  sheet.getRange(4,1).setValue('🔴 < 75% (crítico)');
  sheet.getRange(4,3).setValue('🟡 75–84% (atenção)');
  sheet.getRange(4,5).setValue('🟢 ≥ 85% (OK)');
  sheet.getRange(4,1,1,8).setBackground(COR.VERMELHO_FUNDO);
  sheet.getRange(4,1).setFontWeight('bold').setFontColor(COR.VERMELHO);
  sheet.getRange(4,3).setFontWeight('bold').setFontColor(COR.LARANJA);
  sheet.getRange(4,5).setFontWeight('bold').setFontColor(COR.VERDE_OK);
  sheet.setRowHeight(4, 26);

  const cols = ['Matrícula','Nome Completo','Turma','Total Dias Letivos','Dias Presentes','% Frequência','Status','Ação Necessária'];
  _cabecalhoTabela(sheet, 5, cols);

  // OTIMIZADO: arrays de fórmulas em batch por coluna
  const aF1=[], aF2=[], aF3=[], aF4=[], aF5=[], aF6=[], aF7=[], aF8=[];
  for (let i = 0; i < MAX_ALUNOS; i++) {
    const r = 6 + i;
    const aRow = 6 + i;
    aF1.push([`='👥 Alunos'!B${aRow}`]);
    aF2.push([`='👥 Alunos'!C${aRow}`]);
    aF3.push([`='👥 Alunos'!D${aRow}`]);
    // CORRIGIDO: $B$4 era o título da seção. "Total de Dias Letivos" está em B7.
    aF4.push([`=IF(A${r}="","",'⚙️ Config'!$B$7)`]);
    aF5.push([`=IF(A${r}="","",IFERROR(COUNTIFS('✅ Frequência'!A:A,A${r},'✅ Frequência'!G:G,"Presente"),0))`]);
    aF6.push([`=IF(A${r}="","",IFERROR(E${r}/D${r},0))`]);
    aF7.push([`=IF(A${r}="","",IF(F${r}>=85%,"🟢 Regular",IF(F${r}>=75%,"🟡 Atenção","🔴 Crítico")))`]);
    aF8.push([`=IF(A${r}="","",IF(F${r}<75%,"Acionar Conselho Tutelar – LDB Art.24",IF(F${r}<85%,"Contatar Responsável","OK")))`]);
  }
  sheet.getRange(6,1,MAX_ALUNOS,1).setValues(_lf2(aF1));
  sheet.getRange(6,2,MAX_ALUNOS,1).setValues(_lf2(aF2));
  sheet.getRange(6,3,MAX_ALUNOS,1).setValues(_lf2(aF3));
  sheet.getRange(6,4,MAX_ALUNOS,1).setValues(_lf2(aF4));
  sheet.getRange(6,5,MAX_ALUNOS,1).setValues(_lf2(aF5));
  sheet.getRange(6,6,MAX_ALUNOS,1).setValues(_lf2(aF6)).setNumberFormat('0.0%');
  sheet.getRange(6,7,MAX_ALUNOS,1).setValues(_lf2(aF7));
  sheet.getRange(6,8,MAX_ALUNOS,1).setValues(_lf2(aF8));

  // Formatação condicional
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Crítico').setBackground(COR.VERMELHO_FUNDO).setFontColor(COR.VERMELHO)
      .setRanges([sheet.getRange(6,7,MAX_ALUNOS,1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Atenção').setBackground(COR.AMARELO_FUNDO).setFontColor(COR.LARANJA)
      .setRanges([sheet.getRange(6,7,MAX_ALUNOS,1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Regular').setBackground(COR.VERDE_FUNDO).setFontColor(COR.VERDE_OK)
      .setRanges([sheet.getRange(6,7,MAX_ALUNOS,1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Conselho Tutelar').setBackground(COR.VERMELHO_FUNDO).setFontColor(COR.VERMELHO)
      .setRanges([sheet.getRange(6,8,MAX_ALUNOS,1)]).build(),
  ];
  sheet.setConditionalFormatRules(rules);

  _zebraStripes(sheet, 6, 5+MAX_ALUNOS, cols.length);
  _bordaTabela(sheet, 5, 5+MAX_ALUNOS, cols.length);
  sheet.setFrozenRows(5);
  _configurarImpressao(sheet, 'L', false);
}

// ============================================================
//  ABA: 📅 RELATÓRIO MENSAL
// ============================================================
function _criarRelatorioMensal(sheet) {
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).breakApart();
  sheet.clear();

  const larguras = [200, 130, 130, 130, 130, 130, 140];
  larguras.forEach((w,i) => sheet.setColumnWidth(i+1, w));

  _cabecalhoPrincipal(sheet, '📅 RELATÓRIO MENSAL DE FREQUÊNCIA',
    'Selecione Mês e Ano nas células destacadas para filtrar o relatório', 7);

  // Controles
  const meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                  'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const mesRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(meses, true).setAllowInvalid(false).build();

  sheet.getRange(4,1).setValue('Mês:');
  sheet.getRange(4,2).setValue('Maio').setDataValidation(mesRule)
    .setBackground(COR.AMARELO_FUNDO).setFontWeight('bold').setFontColor(COR.LARANJA);
  sheet.getRange(4,3).setValue('Ano:');
  sheet.getRange(4,4).setValue(ANO)
    .setBackground(COR.AMARELO_FUNDO).setFontWeight('bold').setFontColor(COR.LARANJA)
    .setNumberFormat('0000');
  sheet.getRange(4,5).setValue('Dias Letivos:');
  sheet.getRange(4,6).setValue(20)
    .setBackground(COR.AMARELO_FUNDO).setFontWeight('bold');
  // Mês numérico
  // CORRIGIDO: MATCH com array literal é frágil em locale PT-BR (separadores , vs \).
  // SWITCH é portável e mais legível.
  sheet.getRange(4,7).setValue(_lf('=IFERROR(SWITCH(B4,"Janeiro",1,"Fevereiro",2,"Março",3,"Abril",4,"Maio",5,"Junho",6,"Julho",7,"Agosto",8,"Setembro",9,"Outubro",10,"Novembro",11,"Dezembro",12),0)'))
    .setFontColor(COR.CINZA_BORDA);
  sheet.getRange(4,1,1,7).setBackground(COR.AZUL_CLARO);
  sheet.getRange(4,1).setFontWeight('bold');
  sheet.getRange(4,3).setFontWeight('bold');
  sheet.getRange(4,5).setFontWeight('bold');
  sheet.setRowHeight(4, 28);

  // Subtítulo dinâmico
  sheet.getRange(5,1).setValue(_lf(`="${MAX_ALUNOS} aluno(s) listado(s) para o período de "&B4&"/"&D4`));
  sheet.getRange(5,1,1,7).merge()
    .setBackground(COR.AZUL_CLARO).setFontStyle("italic").setFontColor(COR.CINZA_ESCURO)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(5, 22);

  const cols = ['Aluno','Matrícula','Turma','Presenças','Faltas','% Frequência','Situação'];
  _cabecalhoTabela(sheet, 6, cols);

  // OTIMIZADO: arrays de fórmulas em batch por coluna
  const mF1=[], mF2=[], mF3=[], mF4=[], mF5=[], mF6=[], mF7=[];
  for (let i = 0; i < MAX_ALUNOS; i++) {
    const r = 7 + i;
    const aRow = 6 + i;
    mF1.push([`='👥 Alunos'!C${aRow}`]);
    mF2.push([`='👥 Alunos'!B${aRow}`]);
    mF3.push([`='👥 Alunos'!D${aRow}`]);
    // CORRIGIDO: COUNTIFS não aceita MONTH()/YEAR() aplicados a intervalo como critério.
    // Substituído por SUMPRODUCT, que opera elemento a elemento.
    mF4.push([`=IF(B${r}="","",IFERROR(SUMPRODUCT(('✅ Frequência'!$A$6:$A$10000=B${r})*('✅ Frequência'!$G$6:$G$10000="Presente")*(IFERROR(MONTH('✅ Frequência'!$D$6:$D$10000),0)=$G$4)*(IFERROR(YEAR('✅ Frequência'!$D$6:$D$10000),0)=$D$4)),0))`]);
    mF5.push([`=IF(B${r}="","",$F$4-D${r})`]);
    mF6.push([`=IF(B${r}="","",IFERROR(D${r}/$F$4,0))`]);
    mF7.push([`=IF(B${r}="","",IF(F${r}>=85%,"🟢 Regular",IF(F${r}>=75%,"🟡 Atenção","🔴 Crítico")))`]);
  }
  sheet.getRange(7,1,MAX_ALUNOS,1).setValues(_lf2(mF1));
  sheet.getRange(7,2,MAX_ALUNOS,1).setValues(_lf2(mF2));
  sheet.getRange(7,3,MAX_ALUNOS,1).setValues(_lf2(mF3));
  sheet.getRange(7,4,MAX_ALUNOS,1).setValues(_lf2(mF4));
  sheet.getRange(7,5,MAX_ALUNOS,1).setValues(_lf2(mF5));
  sheet.getRange(7,6,MAX_ALUNOS,1).setValues(_lf2(mF6)).setNumberFormat('0.0%');
  sheet.getRange(7,7,MAX_ALUNOS,1).setValues(_lf2(mF7));
  sheet.setRowHeightsForced(7, MAX_ALUNOS, 22);

  // Formatação condicional
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Crítico').setBackground(COR.VERMELHO_FUNDO).setFontColor(COR.VERMELHO)
      .setRanges([sheet.getRange(7,7,MAX_ALUNOS,1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Atenção').setBackground(COR.AMARELO_FUNDO).setFontColor(COR.LARANJA)
      .setRanges([sheet.getRange(7,7,MAX_ALUNOS,1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Regular').setBackground(COR.VERDE_FUNDO).setFontColor(COR.VERDE_OK)
      .setRanges([sheet.getRange(7,7,MAX_ALUNOS,1)]).build(),
  ];
  sheet.setConditionalFormatRules(rules);

  _zebraStripes(sheet, 7, 6+MAX_ALUNOS, cols.length);
  _bordaTabela(sheet, 6, 6+MAX_ALUNOS, cols.length);
  sheet.setFrozenRows(6);
  _configurarImpressao(sheet, 'L', false);
}

// ============================================================
//  ABA: 📚 RELATÓRIO POR UNIDADE DIDÁTICA
// ============================================================
function _criarRelatorioPorUD(sheet) {
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).breakApart();
  sheet.clear();

  const larguras = [200, 130, 130, 130, 130, 130, 140];
  larguras.forEach((w,i) => sheet.setColumnWidth(i+1, w));

  _cabecalhoPrincipal(sheet, '📚 RELATÓRIO POR UNIDADE DIDÁTICA (BIMESTRE)',
    'Configure as datas de cada UD nas células destacadas', 7);

  // Tabela de UDs
  sheet.getRange(4,1).setValue('Unidade');
  sheet.getRange(4,2).setValue('Início (dd/mm/aaaa)');
  sheet.getRange(4,3).setValue('Fim (dd/mm/aaaa)');
  sheet.getRange(4,4).setValue('Dias Letivos');
  sheet.getRange(4,1,1,4).setBackground(COR.AZUL_ESCURO)
    .setFontColor(COR.BRANCO).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(4, 28);

  const uds = [
    ['UD1', new Date(2026,1,5),  new Date(2026,3,27), 50],
    ['UD2', new Date(2026,3,28), new Date(2026,6,3),  50],
    ['UD3', new Date(2026,6,21), new Date(2026,9,5),  48],
    ['UD4', new Date(2026,9,7),  new Date(2026,11,22),48],
  ];

  uds.forEach(([nome, ini, fim, dias], i) => {
    const r = 5 + i;
    sheet.getRange(r,1).setValue(nome).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(r,2).setValue(ini).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(r,3).setValue(fim).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(r,4).setValue(dias).setHorizontalAlignment('center');
    const bg = i % 2 === 0 ? COR.LINHA_IMPAR : COR.LINHA_PAR;
    sheet.getRange(r,1,1,4).setBackground(bg);
    sheet.setRowHeight(r, 24);
  });
  _bordaTabela(sheet, 4, 8, 4);

  // Seletor de UD
  const udRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['UD1','UD2','UD3','UD4'], true).setAllowInvalid(false).build();
  sheet.getRange(10,1).setValue('Unidade Didática selecionada:');
  sheet.getRange(10,2).setValue('UD2').setDataValidation(udRule)
    .setBackground(COR.AMARELO_FUNDO).setFontWeight('bold').setFontColor(COR.LARANJA);
  sheet.getRange(10,3).setValue('Início:');
  sheet.getRange(10,4).setValue(_lf('=VLOOKUP(B10,A5:D8,2,0)')).setNumberFormat('dd/mm/yyyy')
    .setFontColor(COR.AZUL_PRIMARIO).setFontWeight('bold');
  sheet.getRange(10,5).setValue('Fim:');
  sheet.getRange(10,6).setValue(_lf('=VLOOKUP(B10,A5:D8,3,0)')).setNumberFormat('dd/mm/yyyy')
    .setFontColor(COR.AZUL_PRIMARIO).setFontWeight('bold');
  sheet.getRange(10,7).setValue('Dias:');
  sheet.getRange(10,8).setValue(_lf('=VLOOKUP(B10,A5:D8,4,0)')).setFontWeight('bold');
  sheet.getRange(10,1,1,8).setBackground(COR.AZUL_CLARO);
  sheet.getRange(10,1).setFontWeight('bold');
  sheet.getRange(10,3).setFontWeight('bold');
  sheet.getRange(10,5).setFontWeight('bold');
  sheet.getRange(10,7).setFontWeight('bold');
  sheet.setRowHeight(10, 28);

  // Subtítulo dinâmico
  sheet.getRange(11,1).setValue(_lf(`="${MAX_ALUNOS} aluno(s) para "&B10`));
  sheet.getRange(11,1,1,7).merge()
    .setBackground(COR.AZUL_CLARO).setFontStyle("italic").setHorizontalAlignment('center');
  sheet.setRowHeight(11, 22);

  const cols = ['Aluno','Matrícula','Turma','Presenças','Faltas','% Frequência','Situação'];
  _cabecalhoTabela(sheet, 12, cols);

  // OTIMIZADO: arrays de fórmulas em batch por coluna
  const uF1=[], uF2=[], uF3=[], uF4=[], uF5=[], uF6=[], uF7=[];
  for (let i = 0; i < MAX_ALUNOS; i++) {
    const r = 13 + i;
    const aRow = 6 + i;
    uF1.push([`='👥 Alunos'!C${aRow}`]);
    uF2.push([`='👥 Alunos'!B${aRow}`]);
    uF3.push([`='👥 Alunos'!D${aRow}`]);
    uF4.push([`=IF(B${r}="","",IFERROR(COUNTIFS('✅ Frequência'!A:A,B${r},'✅ Frequência'!G:G,"Presente",'✅ Frequência'!D:D,">="&$D$10,'✅ Frequência'!D:D,"<="&$F$10),0))`]);
    // CORRIGIDO: vazio era " " (espaço) — quebrava subtração nas linhas vazias.
    uF5.push([`=IF(B${r}="","",$H$10-D${r})`]);
    uF6.push([`=IF(B${r}="","",IFERROR(D${r}/$H$10,0))`]);
    uF7.push([`=IF(B${r}="","",IF(F${r}>=85%,"🟢 Regular",IF(F${r}>=75%,"🟡 Atenção","🔴 Crítico")))`]);
  }
  sheet.getRange(13,1,MAX_ALUNOS,1).setValues(_lf2(uF1));
  sheet.getRange(13,2,MAX_ALUNOS,1).setValues(_lf2(uF2));
  sheet.getRange(13,3,MAX_ALUNOS,1).setValues(_lf2(uF3));
  sheet.getRange(13,4,MAX_ALUNOS,1).setValues(_lf2(uF4));
  sheet.getRange(13,5,MAX_ALUNOS,1).setValues(_lf2(uF5));
  sheet.getRange(13,6,MAX_ALUNOS,1).setValues(_lf2(uF6)).setNumberFormat('0.0%');
  sheet.getRange(13,7,MAX_ALUNOS,1).setValues(_lf2(uF7));
  sheet.setRowHeightsForced(13, MAX_ALUNOS, 22);

  _zebraStripes(sheet, 13, 12+MAX_ALUNOS, cols.length);
  _bordaTabela(sheet, 12, 12+MAX_ALUNOS, cols.length);
  sheet.setFrozenRows(12);
  _configurarImpressao(sheet, 'L', false);
}

// ============================================================
//  ABA: ⚙️ CONFIG
// ============================================================
function _criarConfig(sheet) {
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).breakApart();
  sheet.clear();

  const larguras = [230, 380, 320];
  larguras.forEach((w,i) => sheet.setColumnWidth(i+1, w));

  _cabecalhoPrincipal(sheet, '⚙️ CONFIGURAÇÕES DO SISTEMA',
    'Altere os campos destacados em amarelo conforme necessário', 3);

  const secoes = [
    {
      titulo: '🏫 CONFIGURAÇÃO GERAL',
      cor: COR.AZUL_PRIMARIO,
      rows: [
        ['Nome da Escola',          ESCOLA,                                     'Altere para o nome real da escola'],
        ['Total de Dias Letivos',   DIAS_LETIVOS,                               'Base para cálculo de frequência (padrão: 200)'],
        ['Ano Letivo',              ANO,                                        ''],
        ['Frequência Mínima (%)',   FREQ_MINIMA,                               'LDB Art. 24 – mínimo 75%'],
      ]
    },
    {
      titulo: '🔗 INTEGRAÇÃO GOOGLE APPS SCRIPT',
      cor: COR.VERDE_MEDIO,
      rows: [
        ['URL do Web App',          'https://script.google.com/macros/s/SEU_ID_AQUI/exec', 'Cole aqui a URL do Apps Script publicado'],
        ['Status da Integração',    '⏳ Aguardando configuração',                'Altere manualmente após configurar'],
      ]
    },
    {
      titulo: '🕐 TURNOS E HORÁRIOS',
      cor: COR.LARANJA,
      rows: [
        ['Entrada Matutino',   '07:30', ''],
        ['Saída Matutino',     '12:00', ''],
        ['Entrada Vespertino',   '13:00', ''],
        ['Saída Vespertino',     '17:00', ''],
        ['Entrada Noturno',   '19:00', ''],
        ['Saída Noturno',     '22:00', ''],
      ]
    },
    {
      titulo: '👤 RESPONSÁVEIS DO SISTEMA',
      cor: COR.ROXO,
      rows: [
        ['Coordenador(a) Pedagógico(a)', '[Nome]',   'Responsável pelas análises de frequência'],
        ['Responsável de TI / Sistema',  '[Nome]',   'Responsável pela manutenção técnica'],
        ['Porteiro(a) / Operador(a)',    '[Nome]',   'Usuário do app de portaria'],
      ]
    },
    {
      titulo: '📱 APP WEB (QR Code)',
      cor: COR.AZUL_ESCURO,
      rows: [
        ['URL do App',               'https://divinapastorajunqueiro-ship-it.github.io/Divina_Pastora/', 'Link do app de leitura QR Code'],
        ['URL QR Alunos (Script 2)', 'https://script.google.com/macros/s/AKfycbxIRxkJRMGR3vGpvUcCf6uipSxiJmmvKhHYSH_qacDjGKcz8JOl4_6rV02OZnsXfsJK/exec', 'Script para geração de QR dos alunos'],
      ]
    },
  ];

  let currentRow = 4;

  secoes.forEach(secao => {
    // Título da seção
    sheet.getRange(currentRow,1).setValue(secao.titulo);
    sheet.getRange(currentRow,1,1,3).merge()
      .setBackground(secao.cor).setFontColor(COR.BRANCO)
      .setFontWeight('bold').setFontSize(11).setFontFamily('Google Sans')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 32);
    currentRow++;

    // Cabeçalho
    sheet.getRange(currentRow,1,1,3)
      .setValues([['Configuração','Valor','Descrição']])
      .setBackground(COR.CINZA_CLARO).setFontWeight('bold')
      .setFontColor(COR.CINZA_ESCURO).setFontSize(10)
      .setHorizontalAlignment('center');
    sheet.setRowHeight(currentRow, 26);
    currentRow++;

    // Linhas de configuração
    secao.rows.forEach((row, i) => {
      const bg = i % 2 === 0 ? COR.LINHA_IMPAR : COR.LINHA_PAR;
      sheet.getRange(currentRow,1).setValue(row[0])
        .setFontWeight('bold').setFontColor(COR.CINZA_ESCURO).setBackground(bg);
      sheet.getRange(currentRow,2).setValue(row[1])
        .setBackground(COR.AMARELO_FUNDO).setFontColor(COR.CINZA_ESCURO);
      sheet.getRange(currentRow,3).setValue(row[2])
        .setFontColor(COR.CINZA_MEDIO).setFontStyle("italic").setBackground(bg);
      sheet.setRowHeight(currentRow, 26);
      currentRow++;
    });

    _bordaTabela(sheet, currentRow - secao.rows.length - 1, currentRow - 1, 3);
    currentRow++; // espaço entre seções
  });

  // Rodapé com instrução
  const rf = currentRow + 1;
  sheet.getRange(rf,1).setValue('📋 INSTRUÇÕES DE IMPLANTAÇÃO DO WEB APP:');
  sheet.getRange(rf,1,1,3).merge()
    .setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO)
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  sheet.setRowHeight(rf, 30);

  const instrucoes = [
    '1. Neste editor do Apps Script, clique em "Implantar" > "Novo Implantação"',
    '2. Tipo: App da Web | Executar como: Eu | Quem tem acesso: Qualquer pessoa',
    '3. Clique em "Implantar" e copie a URL gerada',
    '4. Cole a URL na célula B da linha "URL do Web App" acima',
    '5. Cole a mesma URL no campo de configuração do App Web (ícone ⚙️ no app)',
    '6. Clique em "Carregar Alunos" no App para verificar a conexão',
  ];

  instrucoes.forEach((inst, i) => {
    sheet.getRange(rf+1+i, 1).setValue(inst);
    sheet.getRange(rf+1+i, 1, 1, 3).merge()
      .setBackground(i%2===0 ? COR.AZUL_CLARO : COR.LINHA_IMPAR)
      .setFontColor(COR.CINZA_ESCURO).setFontSize(10);
    sheet.setRowHeight(rf+1+i, 26);
  });

  _configurarImpressao(sheet, 'P', true);
}

// ============================================================
//  WEB APP – BACKEND PARA O APP DE PORTARIA
// ============================================================

/**
 * Trata requisições GET do App Web
 * Actions disponíveis:
 *   ?action=ping              → teste de conexão
 *   ?action=listarAlunos      → retorna todos os alunos ativos
 */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'ping';
  let result;

  if (action === 'ping') {
    result = { status: 'ok', versao: 'v4', message: 'Conexão OK', timestamp: new Date().toISOString() };
  } else if (action === 'listarAlunos') {
    result = _listarAlunos();
  } else if (action === 'registrarPresenca') {
    result = _registrarFrequencia(e.parameter);
  } else {
    result = { status: 'error', message: 'Action desconhecida: ' + action };
  }

  const cb = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : null;
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Trata requisições POST do App Web
 * Body JSON esperado:
 * {
 *   matricula, nome, turma, data, hora, turno, status, via, timestamp
 * }
 */
function doPost(e) {
  let result;
  try {
    let dados;
    if (e.postData && e.postData.contents) {
      dados = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      dados = e.parameter;
    } else {
      throw new Error('Nenhum dado recebido');
    }
    result = _registrarFrequencia(dados);
  } catch (err) {
    result = { status: 'error', message: 'Erro ao processar: ' + err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function _listarAlunos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('👥 Alunos');
  if (!sheet) return { status: 'error', message: 'Aba Alunos não encontrada' };

  const data = sheet.getRange(6, 2, MAX_ALUNOS, 5).getValues(); // B:F
  const alunos = [];

  data.forEach(row => {
    const matricula = String(row[0]).trim();
    const nome      = String(row[1]).trim();
    const turma     = String(row[2]).trim();
    const situacao  = String(row[3]).trim();
    const turno     = String(row[4]).trim();

    if (matricula && matricula !== '' && situacao.toLowerCase() === 'ativo') {
      alunos.push({ matricula, nome, turma, turno });
    }
  });

  return { status: 'ok', total: alunos.length, alunos };
}

function _registrarFrequencia(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('✅ Frequência');
  if (!sheet) return { status: 'error', message: 'Aba Frequência não encontrada' };

  const agora = new Date();
  const dataHoje = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const horaAgora = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'HH:mm:ss');

  // Verificar duplicata (mesma matrícula, mesmo dia)
  const ultimaLinha = _ultimaLinhaPreenchida(sheet, 1);
  if (ultimaLinha >= 6) {
    const registros = sheet.getRange(6, 1, ultimaLinha - 5, 4).getValues();
    for (const reg of registros) {
      const mat = String(reg[0]).trim();
      const dt  = reg[3] instanceof Date
        ? Utilities.formatDate(reg[3], Session.getScriptTimeZone(), 'dd/MM/yyyy')
        : String(reg[3]);
      if (mat === String(dados.matricula).trim() && dt === dataHoje) {
        return { status: 'duplicado', message: 'Já registrado hoje', hora: String(reg[4]) };
      }
    }
  }

  // Inserir nova linha
  const novaLinha = [
    dados.matricula || '',
    dados.nome      || '',
    dados.turma     || '',
    agora,
    dados.hora      || horaAgora,
    dados.turno     || '',
    dados.status    || 'Presente',
    dados.via       || 'App Web',
    agora,
  ];

  const proxLinha = _ultimaLinhaPreenchida(sheet, 1) + 1;
  sheet.getRange(Math.max(proxLinha, 6), 1, 1, novaLinha.length).setValues([novaLinha]);

  // Formatação da linha inserida
  const r = Math.max(proxLinha, 6);
  sheet.getRange(r, 4).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(r, 5).setNumberFormat('hh:mm:ss');
  sheet.getRange(r, 9).setNumberFormat('dd/mm/yyyy hh:mm:ss');

  return {
    status: 'ok',
    message: 'Frequência registrada com sucesso',
    matricula: dados.matricula,
    nome: dados.nome,
    turma: dados.turma,
    hora: dados.hora || horaAgora,
    timestamp: agora.toISOString()
  };
}

function _ultimaLinhaPreenchida(sheet, col) {
  const vals = sheet.getRange(6, col, sheet.getMaxRows() - 5, 1).getValues();
  let ultima = 5;
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== '') { ultima = 6 + i; break; }
  }
  return ultima;
}

// ============================================================
//  GATILHO onEdit – PROPAGAÇÃO DO NOME DA ESCOLA
// ============================================================
/**
 * Disparado automaticamente pelo Google Sheets quando qualquer
 * célula é editada. Quando o usuário altera o nome da escola
 * na aba ⚙️ Config (célula B6), esta função atualiza o
 * cabeçalho (linha 1) de TODAS as outras abas automaticamente.
 *
 * IMPORTANTE: Para funcionar em planilhas compartilhadas e em
 * triggers instalados, registre também um trigger "onEdit"
 * instalável em Extensões > Apps Script > Gatilhos.
 */
function onEdit(e) {
  if (!e) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // Só age quando a edição for na aba Config, coluna B (valor), linha 6 (Nome da Escola)
  if (sheetName === '⚙️ Config' && e.range.getColumn() === 2 && e.range.getRow() === 6) {
    atualizarNomeEscolaEmTodasAsAbas();
  }
}

/**
 * Atualiza o cabeçalho (linha 1, Rich Text) de todas as abas
 * do sistema com o nome da escola lido de ⚙️ Config!B6.
 * Pode ser chamada manualmente pelo menu ou acionada pelo onEdit.
 */
function atualizarNomeEscolaEmTodasAsAbas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nomeEscola = _getNomeEscola();
  const sepBotao   = '     🗺️ ← Voltar à Navegação';
  const nav         = ss.getSheetByName('🗺️ Navegação');

  // Abas que têm botão "Voltar à Navegação" (todas exceto Navegação)
  const abasComBotao = [
    '🏠 Início','👥 Alunos','✅ Frequência','🔲 QR Codes',
    '📋 Status','📊 Dashboard','⚠️ Alertas',
    '📅 Rel. Mensal','📚 Rel. por UD','⚙️ Config',
  ];
  // Aba Navegação – sem botão
  const abasSemBotao = ['🗺️ Navegação'];

  const processarAba = (sheet, comBotao) => {
    if (!sheet) return;
    try {
      // Quantas colunas mescladas tem a linha 1? Detectar pelo merge existente
      // Usamos o número de colunas do merge atual ou fallback de 9
      let numCols = 9;
      try {
        const merges = sheet.getRange(1,1,1,20).getMergedRanges();
        if (merges.length > 0) numCols = merges[0].getNumColumns();
      } catch(_) {}

      // Desfaz merge anterior para recriar sem conflito
      sheet.getRange(1,1,1,numCols).breakApart();

      const textoL1 = comBotao ? (nomeEscola + sepBotao) : nomeEscola;
      const rtBuilder = SpreadsheetApp.newRichTextValue()
        .setText(textoL1)
        .setTextStyle(0, nomeEscola.length,
          SpreadsheetApp.newTextStyle()
            .setFontFamily('Google Sans').setFontSize(11)
            .setForegroundColor(COR.BRANCO).setBold(true).build());

      if (comBotao && nav) {
        const url = '#gid=' + nav.getSheetId();
        rtBuilder.setTextStyle(nomeEscola.length, textoL1.length,
          SpreadsheetApp.newTextStyle()
            .setFontFamily('Google Sans').setFontSize(11)
            .setForegroundColor('#FFEB3B').setBold(true)
            .setUnderline(true).build());
        rtBuilder.setLinkUrl(nomeEscola.length, textoL1.length, url);
      }

      sheet.getRange(1,1).setRichTextValue(rtBuilder.build());
      sheet.getRange(1,1,1,numCols).merge()
        .setBackground(COR.HEADER_GRAD)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
    } catch(err) {
      // Aba pode estar protegida ou com outro estado – ignora silenciosamente
      Logger.log('Erro ao atualizar aba ' + sheet.getName() + ': ' + err.message);
    }
  };

  abasComBotao.forEach(nome => processarAba(ss.getSheetByName(nome), true));
  abasSemBotao.forEach(nome => processarAba(ss.getSheetByName(nome), false));

  SpreadsheetApp.flush();
}

// ============================================================
//  MENU PERSONALIZADO
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Sistema de Frequência')
    .addItem('🔄 Recriar Planilha Completa', 'criarPlanilhaCompleta')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔧 Criação em Etapas (se der timeout)')
      .addItem('📌 Etapa 1 – Estrutura, Alunos e Frequência', 'criarEtapa1_EstruturaeAlunos')
      .addItem('📌 Etapa 2 – QR Codes, Status e Dashboard',   'criarEtapa2_QRCodesStatusDashboard')
      .addItem('📌 Etapa 3 – Alertas, Relatórios e Config',   'criarEtapa3_AlertasRelatoriosConfig'))
    .addSeparator()
    .addItem('📋 Atualizar Dashboard', 'atualizarDashboard')
    .addItem('🔔 Verificar Alertas', 'verificarAlertas')
    .addItem('🏫 Atualizar Nome da Escola em Todas as Abas', 'atualizarNomeEscolaEmTodasAsAbas')
    .addSeparator()
    .addItem('📥 Exportar Relatório do Mês (PDF)', 'exportarRelatorioPDF')
    .addItem('🗑️ Limpar Registros de Frequência', 'limparFrequencia')
    .addToUi();
}

function atualizarDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📊 Dashboard');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Aba Dashboard não encontrada.');
    return;
  }

  // Lê turmas atuais da aba Alunos
  const abaAlunos = ss.getSheetByName('👥 Alunos');
  let turmasUnicas = [];
  if (abaAlunos) {
    const vals = abaAlunos.getRange(6, 4, MAX_ALUNOS, 1).getValues();
    const setTurmas = new Set();
    vals.forEach(row => {
      const t = String(row[0]).trim();
      if (t && t !== '' && t !== 'undefined') setTurmas.add(t);
    });
    turmasUnicas = Array.from(setTurmas).sort();
  }

  // Limpar área da tabela de turmas (linhas 8 em diante, 6 colunas)
  const maxLinhas = sheet.getMaxRows();
  if (maxLinhas >= 8) {
    sheet.getRange(8, 1, maxLinhas - 7, 6).clearContent().clearFormat();
  }

  let rt = 8;

  if (turmasUnicas.length > 0) {
    turmasUnicas.forEach((turma, i) => {
      const r = 8 + i;
      const bg = i % 2 === 0 ? COR.LINHA_IMPAR : COR.LINHA_PAR;
      sheet.getRange(r, 1).setValue(turma).setFontWeight('bold').setBackground(bg);
      sheet.getRange(r, 2).setValue(_lf(`=IFERROR(COUNTIFS('👥 Alunos'!D:D,"${turma}",'👥 Alunos'!E:E,"Ativo"),0)`)).setBackground(bg);
      sheet.getRange(r, 3).setValue(_lf(
        `=IFERROR(COUNTIFS('✅ Frequência'!C:C,"${turma}",'✅ Frequência'!D:D,">="&INT(TODAY()),'✅ Frequência'!D:D,"<"&INT(TODAY())+1,'✅ Frequência'!G:G,"Presente"),0)`
      )).setBackground(bg);
      sheet.getRange(r, 4).setValue(_lf(`=IFERROR(B${r}-C${r},0)`)).setBackground(bg);
      sheet.getRange(r, 5).setValue(_lf(`=IFERROR(C${r}/B${r},0)`))
        .setNumberFormat('0.0%').setBackground(bg);
      sheet.getRange(r, 6).setValue(_lf(
        `=IF(B${r}=0,"—",IF(E${r}>=0.85,"🟢 Normal",IF(E${r}>=0.75,"🟡 Atenção","🔴 Crítico")))`
      )).setBackground(bg);
      sheet.setRowHeight(r, 24);
    });
    rt = 8 + turmasUnicas.length;
  }

  // Linha TOTAL GERAL
  sheet.getRange(rt, 1).setValue('TOTAL GERAL').setFontWeight('bold')
    .setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
  if (turmasUnicas.length > 0) {
    sheet.getRange(rt, 2).setValue(_lf(`=SUM(B8:B${rt-1})`))
      .setFontWeight('bold').setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
    sheet.getRange(rt, 3).setValue(_lf(`=SUM(C8:C${rt-1})`))
      .setFontWeight('bold').setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
    sheet.getRange(rt, 4).setValue(_lf(`=SUM(D8:D${rt-1})`))
      .setFontWeight('bold').setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
  } else {
    sheet.getRange(rt, 2).setValue(_lf(`=IFERROR(COUNTIFS('👥 Alunos'!E:E,"Ativo"),0)`))
      .setFontWeight('bold').setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
    sheet.getRange(rt, 3).setValue(_lf(
      `=IFERROR(COUNTIFS('✅ Frequência'!D:D,">="&INT(TODAY()),'✅ Frequência'!D:D,"<"&INT(TODAY())+1,'✅ Frequência'!G:G,"Presente"),0)`
    )).setFontWeight('bold').setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
    sheet.getRange(rt, 4).setValue(_lf(`=IFERROR(B${rt}-C${rt},0)`))
      .setFontWeight('bold').setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
  }
  sheet.getRange(rt, 5).setValue(_lf(`=IFERROR(C${rt}/B${rt},0)`))
    .setNumberFormat('0.0%').setFontWeight('bold')
    .setBackground(COR.AZUL_ESCURO).setFontColor(COR.BRANCO);
  sheet.getRange(rt, 6).setValue('').setBackground(COR.AZUL_ESCURO);
  sheet.setRowHeight(rt, 28);

  _bordaTabela(sheet, 7, rt, 6);

  // Formatação condicional
  const numLinhasFC = Math.max(turmasUnicas.length, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0.85).setBackground(COR.VERDE_FUNDO).setFontColor(COR.VERDE_OK)
      .setRanges([sheet.getRange(8, 5, numLinhasFC, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.75, 0.849).setBackground(COR.AMARELO_FUNDO).setFontColor(COR.LARANJA)
      .setRanges([sheet.getRange(8, 5, numLinhasFC, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.75).setBackground(COR.VERMELHO_FUNDO).setFontColor(COR.VERMELHO)
      .setRanges([sheet.getRange(8, 5, numLinhasFC, 1)]).build(),
  ]);

  SpreadsheetApp.flush();
  sheet.activate();

  const msg = turmasUnicas.length > 0
    ? `✅ Dashboard atualizado com ${turmasUnicas.length} turma(s): ${turmasUnicas.join(', ')}`
    : '✅ Dashboard atualizado! (Nenhuma turma cadastrada ainda na aba Alunos)';
  SpreadsheetApp.getUi().alert(msg);
}

function verificarAlertas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const alertSheet = ss.getSheetByName('⚠️ Alertas');
  if (alertSheet) alertSheet.activate();
  SpreadsheetApp.getUi().alert(
    '⚠️ Verifique a aba Alertas para alunos abaixo de 75% de frequência.'
  );
}

function exportarRelatorioPDF() {
  SpreadsheetApp.getUi().alert(
    'Para exportar como PDF:\n' +
    'Acesse a aba 📅 Rel. Mensal > Arquivo > Download > PDF'
  );
}

function limparFrequencia() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '⚠️ ATENÇÃO',
    'Isso apagará TODOS os registros de frequência. Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('✅ Frequência');
  if (sheet) {
    sheet.getRange(6, 1, sheet.getMaxRows() - 5, 9).clearContent();
    ui.alert('✅ Registros de frequência removidos.');
  }
}
