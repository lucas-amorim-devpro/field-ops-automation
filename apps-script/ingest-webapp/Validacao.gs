/**
 * Validação dos envios recebidos pelo doPost.
 *
 * Funções PURAS (sem SpreadsheetApp, PropertiesService etc.): rodam iguais no Apps Script
 * e no navegador, onde são testadas por apps-script/tests/tests.html.
 */

var COLUNAS_INDICADORES = ['grupo', 'chave', 'apuradas', 'no_prazo', 'fora', 'pendentes',
  'percentual', 'selo', 'percentual_d1', 'variacao_pp'];

var COLUNAS_ORDENS = ['id_os', 'data', 'regional', 'cidade', 'gerente', 'coordenacao',
  'supervisao', 'tecnico', 'tipo', 'slot', 'chegada', 'status'];

var RELATORIOS = {
  'chegada-slot': { colunas: COLUNAS_INDICADORES, maxLinhas: 500 },
  'ordens': { colunas: COLUNAS_ORDENS, maxLinhas: 50000 }
};

var SELOS_VALIDOS = ['Na meta', 'Perto', 'Abaixo', 'Sem dados'];

/**
 * Compara dois textos em tempo constante (não "vaza" quantos caracteres acertou).
 */
function tokenIgual(recebido, esperado) {
  if (typeof recebido !== 'string' || typeof esperado !== 'string' || !esperado) return false;
  var diferenca = recebido.length ^ esperado.length;
  for (var i = 0; i < esperado.length; i++) {
    diferenca |= (recebido.charCodeAt(i % (recebido.length || 1)) || 0) ^ esperado.charCodeAt(i);
  }
  return diferenca === 0;
}

function ehDataIso(texto) {
  if (typeof texto !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(texto)) return false;
  var p = texto.split('-').map(Number);
  var d = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  return d.getUTCFullYear() === p[0] && d.getUTCMonth() === p[1] - 1 && d.getUTCDate() === p[2];
}

function ehNumeroOuNulo(v) {
  return v === null || (typeof v === 'number' && isFinite(v));
}

/**
 * Valida o corpo já convertido de JSON.
 * @return {{ok: boolean, erro?: string}}
 */
function validarPayload(payload, tokenEsperado) {
  if (!payload || typeof payload !== 'object') return { ok: false, erro: 'Corpo vazio ou não é JSON.' };
  if (!tokenIgual(payload.token, tokenEsperado)) return { ok: false, erro: 'Token inválido.' };

  var regra = RELATORIOS[payload.relatorio];
  if (!regra) return { ok: false, erro: 'Relatório desconhecido: ' + payload.relatorio };
  if (!ehDataIso(payload.dataReferencia)) return { ok: false, erro: 'dataReferencia deve ser AAAA-MM-DD.' };

  if (!Array.isArray(payload.colunas) || payload.colunas.join('|') !== regra.colunas.join('|')) {
    return { ok: false, erro: 'Colunas diferentes do esperado: ' + regra.colunas.join(', ') };
  }
  if (!Array.isArray(payload.linhas) || payload.linhas.length === 0) return { ok: false, erro: 'Nenhuma linha enviada.' };
  if (payload.linhas.length > regra.maxLinhas) return { ok: false, erro: 'Linhas demais (máx. ' + regra.maxLinhas + ').' };

  for (var i = 0; i < payload.linhas.length; i++) {
    var linha = payload.linhas[i];
    if (!Array.isArray(linha) || linha.length !== regra.colunas.length) {
      return { ok: false, erro: 'Linha ' + (i + 1) + ' com quantidade de colunas errada.' };
    }
    if (payload.relatorio === 'chegada-slot') {
      var erro = validarLinhaIndicador(linha);
      if (erro) return { ok: false, erro: 'Linha ' + (i + 1) + ': ' + erro };
    }
  }
  return { ok: true };
}

function validarLinhaIndicador(l) {
  if (typeof l[0] !== 'string' || typeof l[1] !== 'string') return 'grupo e chave devem ser texto.';
  for (var c = 2; c <= 5; c++) {
    if (typeof l[c] !== 'number' || l[c] < 0 || Math.floor(l[c]) !== l[c]) return COLUNAS_INDICADORES[c] + ' deve ser inteiro >= 0.';
  }
  if (!ehNumeroOuNulo(l[6]) || (l[6] !== null && (l[6] < 0 || l[6] > 100))) return 'percentual fora de 0..100.';
  if (SELOS_VALIDOS.indexOf(l[7]) < 0) return 'selo desconhecido: ' + l[7];
  if (!ehNumeroOuNulo(l[8]) || !ehNumeroOuNulo(l[9])) return 'percentual_d1 e variacao_pp devem ser número ou null.';
  return null;
}

/**
 * Linhas para a aba Historico: data de referência + "agora" + colunas do indicador.
 */
function montarLinhasHistorico(payload) {
  return payload.linhas.map(function (l) {
    return [payload.dataReferencia, payload.agora || '', payload.relatorio].concat(l.map(function (v) {
      return v === null ? '' : v;
    }));
  });
}
