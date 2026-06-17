/**
 * SC FIRE - Coleta de Opiniões para Treinamento de Segurança contra Incêndio
 * Backend em Google Apps Script.
 *
 * ESTRUTURA DA PLANILHA (criada automaticamente por setupInicial):
 *  - Aba "Parâmetros": coluna A = lista de condomínios (a partir da linha 2)
 *  - Aba "Respostas": cabeçalho fixo + 1 linha por voto recebido
 *
 * COMO IMPLANTAR (passo a passo):
 *  1) Crie uma Planilha Google nova (ou abra a que deseja usar como banco de dados).
 *  2) Extensões > Apps Script. Cole este arquivo como "Código.gs".
 *  3) Crie um arquivo HTML chamado "index" e cole o conteúdo de index.html.
 *  4) Na barra de funções do editor, selecione "setupInicial" e clique em Executar
 *     (autorize as permissões solicitadas). Isso cria e formata as abas
 *     "Parâmetros" e "Respostas" automaticamente.
 *  5) (Opcional) Edite a aba "Parâmetros" para colocar os nomes reais dos
 *     30 condomínios, substituindo os exemplos "Condomínio 01"..."Condomínio 30".
 *  6) Clique em Implantar > Nova implantação.
 *     - Tipo: "App da Web" (Web app)
 *     - Executar como: "Eu" (sua conta)
 *     - Quem pode acessar: "Qualquer pessoa" (para moradores sem login Google)
 *  7) Copie a URL gerada (.../exec) — esse é o link a ser distribuído aos moradores.
 *  8) Cada nova publicação de alterações no código exige uma NOVA implantação
 *     (ou "Gerenciar implantações" > editar > Nova versão) para valer no link já enviado.
 */

// Nomes das abas usadas pela aplicação.
var SHEET_RESPOSTAS = 'Respostas';
var SHEET_PARAMETROS = 'Parâmetros';

// Temas fixos do treinamento (na ordem em que aparecem no formulário e na planilha).
var TEMAS = [
  'Sistemas e medidas preventivas contra incêndio e pânico',
  'Doenças e casos clínicos envolvendo a terceira idade',
  'Uso e manuseio de extintores (Teoria)',
  'Uso e manuseio de extintores (Prática)',
  'Teoria do incêndio e Classes de incêndio',
  'Procedimentos de Evacuação',
  'Atividade de Brigada de Incêndio',
  'Normas do Corpo de Bombeiros',
  'Segurança e alimentação de veículos elétricos em condomínios'
];

/**
 * Serve a página principal do Web App.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SC FIRE - Pesquisa de Treinamento')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Configuração inicial: cria e formata as abas "Parâmetros" e "Respostas"
 * caso ainda não existam. Execute esta função uma única vez manualmente
 * pelo editor do Apps Script antes de implantar o Web App.
 */
function setupInicial() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Aba "Parâmetros" ---
  var paramSheet = ss.getSheetByName(SHEET_PARAMETROS);
  if (!paramSheet) {
    paramSheet = ss.insertSheet(SHEET_PARAMETROS);
    paramSheet.getRange(1, 1).setValue('Condomínio').setFontWeight('bold');

    var nomesExemplo = [];
    for (var i = 1; i <= 30; i++) {
      var numero = (i < 10 ? '0' + i : '' + i);
      nomesExemplo.push(['Condomínio ' + numero]);
    }
    paramSheet.getRange(2, 1, nomesExemplo.length, 1).setValues(nomesExemplo);
    paramSheet.autoResizeColumn(1);
  }

  // --- Aba "Respostas" ---
  var respSheet = ss.getSheetByName(SHEET_RESPOSTAS);
  if (!respSheet) {
    respSheet = ss.insertSheet(SHEET_RESPOSTAS);

    var cabecalho = ['Data/Hora', 'Condomínio'].concat(TEMAS).concat(['Sugestão (Outros)']);
    respSheet.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]);

    var headerRange = respSheet.getRange(1, 1, 1, cabecalho.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#d9d9d9');
    respSheet.setFrozenRows(1);
    respSheet.autoResizeColumns(1, cabecalho.length);
  }

  SpreadsheetApp.getUi().alert('Setup concluído! Abas "Parâmetros" e "Respostas" prontas.');
}

/**
 * Retorna a lista de condomínios cadastrados na aba "Parâmetros",
 * em formato de array simples, para popular o <select> do front-end.
 */
function getCondominios() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PARAMETROS);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var valores = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var lista = [];
  for (var i = 0; i < valores.length; i++) {
    var nome = (valores[i][0] || '').toString().trim();
    if (nome !== '') lista.push(nome);
  }
  return lista;
}

/**
 * Recebe os dados do formulário (enviados via google.script.run) e grava
 * uma nova linha na aba "Respostas". A pesquisa é 100% anônima: nenhum
 * dado pessoal (nome, e-mail, CPF, etc.) é aceito ou armazenado.
 *
 * @param {Object} dados - objeto vindo do front-end, no formato:
 *   {
 *     condominio: "Condomínio 01",
 *     notas: [int, int, int, int, int, int, int, int, int], // 9 notas, 1 a 5
 *     sugestao: "texto livre opcional"
 *   }
 */
function doPost(e) {
  try {
    var dados = JSON.parse(e.postData.contents);
    
    // Se for uma requisição de ação da API externa (ex: GitHub Pages)
    if (dados && dados.action === 'getCondominios') {
      return ContentService.createTextOutput(
        JSON.stringify({ sucesso: true, lista: getCondominios() })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Caso contrário, é o registro de resposta normal do formulário
    return ContentService.createTextOutput(
      JSON.stringify(registrarResposta(dados))
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ sucesso: false, erro: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Função chamável diretamente do front-end via google.script.run
 * (caminho preferencial quando o front-end é servido pelo próprio
 * HtmlService, evitando problemas de CORS do doPost).
 */
function registrarResposta(dados) {
  var condominio = (dados && dados.condominio || '').toString().trim();
  if (condominio === '') {
    return { sucesso: false, erro: 'Condomínio não informado.' };
  }

  var notas = (dados && dados.notas) || [];
  if (!Array.isArray(notas) || notas.length !== TEMAS.length) {
    return { sucesso: false, erro: 'Notas inválidas ou incompletas.' };
  }

  var notasValidadas = [];
  for (var i = 0; i < TEMAS.length; i++) {
    var nota = parseInt(notas[i], 10);
    if (isNaN(nota) || nota < 1 || nota > 5) {
      return { sucesso: false, erro: 'Nota inválida para o tema: ' + TEMAS[i] };
    }
    notasValidadas.push(nota);
  }

  var sugestao = (dados && dados.sugestao || '').toString().trim();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_RESPOSTAS);
  if (!sheet) {
    return { sucesso: false, erro: 'Aba "Respostas" não encontrada. Execute setupInicial().' };
  }

  var linha = [new Date(), condominio].concat(notasValidadas).concat([sugestao]);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, linha.length).setValues([linha]);

  return { sucesso: true };
}
