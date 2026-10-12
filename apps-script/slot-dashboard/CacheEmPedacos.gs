/**
 * Cache em pedaços.
 *
 * O CacheService guarda no máximo 100 KB por chave. O JSON do painel passa disso, então
 * ele é dividido em pedaços (dados_0, dados_1, ...) mais uma chave de controle
 * (dados_meta) com a versão e a quantidade de pedaços. Na leitura, se a versão mudou
 * ou faltar QUALQUER pedaço (o cache pode expulsar chaves isoladas), descarta tudo e
 * remonta a partir da planilha.
 *
 * Funções PURAS: recebem o objeto de cache por parâmetro (CacheService.getScriptCache()
 * no Apps Script; um objeto falso nos testes do navegador).
 */

// 45 mil caracteres: mesmo que todos fossem acentuados (2 bytes em UTF-8), ficam ~90 KB.
var TAMANHO_PEDACO = 45000;
var CACHE_TTL_SEGUNDOS = 21600;   // 6 h, o máximo do CacheService

function dividirEmPedacos(texto, tamanho) {
  var t = tamanho || TAMANHO_PEDACO;
  var pedacos = [];
  for (var i = 0; i < texto.length; i += t) pedacos.push(texto.slice(i, i + t));
  return pedacos.length ? pedacos : [''];
}

function gravarCacheEmPedacos(cache, prefixo, versao, texto, tamanho) {
  var pedacos = dividirEmPedacos(texto, tamanho);
  var valores = {};
  pedacos.forEach(function (p, i) { valores[prefixo + '_' + i] = p; });
  valores[prefixo + '_meta'] = JSON.stringify({ versao: versao, partes: pedacos.length });
  cache.putAll(valores, CACHE_TTL_SEGUNDOS);
  return pedacos.length;
}

/** Devolve o texto completo, ou null se não houver cache válido para esta versão. */
function lerCacheEmPedacos(cache, prefixo, versao) {
  var metaTexto = cache.get(prefixo + '_meta');
  if (!metaTexto) return null;
  var meta;
  try { meta = JSON.parse(metaTexto); } catch (e) { return null; }
  if (!meta || meta.versao !== versao || !(meta.partes > 0)) return null;

  var chaves = [];
  for (var i = 0; i < meta.partes; i++) chaves.push(prefixo + '_' + i);
  var achados = cache.getAll(chaves);
  var partes = [];
  for (var j = 0; j < chaves.length; j++) {
    if (typeof achados[chaves[j]] !== 'string') return null;   // pedaço expulso do cache
    partes.push(achados[chaves[j]]);
  }
  return partes.join('');
}
