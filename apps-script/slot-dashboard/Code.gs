/**
 * Painel "Chegada ao Slot" no Apps Script (doGet + HtmlService).
 *
 * Mesma interface do dashboard-pages: Styles.html, Core.html e App.html são GERADOS a partir
 * de dashboard-pages/ por tools/Build-AppsScriptDashboard.ps1 (não edite à mão).
 *
 * Desempenho (a 1ª abertura caiu de ~18 s para ~2 s com estas três medidas):
 *   1. Os dados vão embutidos no HTML na 1ª carga: a página não espera uma 2ª requisição.
 *   2. JSON compactado (dicionário + linhas) em vez de lista de objetos.
 *   3. Cache em pedaços (CacheService, 100 KB por chave) com a versão da base.
 * Depois, o navegador faz só a consulta leve de versão (obterVersao) a cada 90 s.
 *
 * Propriedades do script: PLANILHA_ID (planilha com as abas Base e Controle).
 * O código só LÊ a planilha. O escopo no manifesto é o completo (spreadsheets) porque
 * SpreadsheetApp.openById não funciona com spreadsheets.readonly ("permissions are not sufficient").
 * Acesso da implantação: ANYONE_ANONYMOUS ("Qualquer pessoa", sem login), necessário para o
 * painel público e para incorporar no Google Sites. ("ANYONE" exige conta Google.)
 */

var REGRAS = {
  janelaMinutos: 120,
  toleranciaMinutos: { 'Instalação': 0, 'Reparo': 30 },
  meta: 80,
  margemPerto: 5
};

function doGet(e) {
  var acao = (e && e.parameter && e.parameter.acao) || '';
  if (acao === 'dados') return json_(obterDados());
  if (acao === 'base') return json_(JSON.stringify({ versao: obterVersao() }));

  var modelo = HtmlService.createTemplateFromFile('Index');
  modelo.dadosIniciais = paraScript_(obterDados());
  return modelo.evaluate()
    .setTitle('Chegada ao Slot')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    // Permite incorporar no Google Sites (iframe).
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Usado no Index.html: <?!= include('Styles'); ?> */
function include(nome) {
  return HtmlService.createHtmlOutputFromFile(nome).getContent();
}

/** Consulta leve chamada pelo navegador (google.script.run) a cada 90 s. */
function obterVersao() {
  return planilha_().getSheetByName('Controle').getRange('B1').getDisplayValue();
}

/** JSON compactado (texto), do cache quando a versão não mudou. */
function obterDados() {
  var versao = obterVersao();
  var cache = CacheService.getScriptCache();
  var texto = lerCacheEmPedacos(cache, 'dados', versao);
  if (texto) return texto;

  var tabela = planilha_().getSheetByName('Base').getDataRange().getDisplayValues();
  var agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm");
  texto = JSON.stringify(compactarOrdens(tabela, { versao: versao, agora: agora, regras: REGRAS }));
  gravarCacheEmPedacos(cache, 'dados', versao, texto);
  return texto;
}

function planilha_() {
  var id = PropertiesService.getScriptProperties().getProperty('PLANILHA_ID');
  if (!id) throw new Error('Defina PLANILHA_ID em Configurações do projeto > Propriedades do script.');
  return SpreadsheetApp.openById(id);
}

function json_(texto) {
  return ContentService.createTextOutput(texto).setMimeType(ContentService.MimeType.JSON);
}

/** Torna o JSON seguro dentro de <script>: "</script>" no meio dos dados fecharia a tag. */
function paraScript_(texto) {
  return texto.replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
}
