/**
 * Monta o JSON compactado (dicionário + linhas) a partir da aba Base.
 * Mesmo formato gerado por data-sample/Export-DashboardData.ps1 e lido por core.js.
 *
 * Funções PURAS: testadas no navegador por apps-script/tests/tests.html.
 *
 * A aba Base é lida com getDisplayValues() (texto), não getValues(): horas guardadas no
 * Sheets voltam como Date de 30/12/1899, e o fuso histórico dessa data (LMT) desloca
 * minutos e segundos. Texto "14:03" não tem esse problema.
 */

/** "12/10/2026", "2026-10-12" ou "12/10/26" -> "2026-10-12" (ou null). */
function normalizarData(texto) {
  var s = String(texto || '').trim();
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(s);
  if (!m) return null;
  var ano = m[3].length === 2 ? '20' + m[3] : m[3];
  return ano + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
}

/** "8:05", "08:05" ou "08:05:00" -> minutos desde 00:00 (ou null). */
function normalizarHora(texto) {
  var m = /^(\d{1,2}):(\d{2})(?::\d{2})?$/.exec(String(texto || '').trim());
  if (!m) return null;
  var h = Number(m[1]), min = Number(m[2]);
  return h < 24 && min < 60 ? h * 60 + min : null;
}

function minutosParaTexto(minutos) {
  return ('0' + Math.floor(minutos / 60)).slice(-2) + ':' + ('0' + (minutos % 60)).slice(-2);
}

/**
 * @param {Array<Array<string>>} tabela  Valores da aba Base (1ª linha = cabeçalho).
 * @param {{versao: string, agora: string, regras: Object, diasHistorico?: number}} opcoes
 */
function compactarOrdens(tabela, opcoes) {
  var cab = tabela[0].map(function (c) { return String(c).trim().toLowerCase(); });
  var col = {};
  COLUNAS_BASE.forEach(function (nome) {
    var i = cab.indexOf(nome);
    if (i < 0) throw new Error('Coluna ausente na aba Base: ' + nome);
    col[nome] = i;
  });

  var dic = { datas: [], regional: [], gerente: [], coordenacao: [], supervisao: [], tecnico: [], cidade: [], tipo: [], slot: [], status: [] };
  var indices = {};
  Object.keys(dic).forEach(function (k) { indices[k] = {}; });
  function indice(nome, chave, entrada) {
    if (!(chave in indices[nome])) {
      indices[nome][chave] = dic[nome].length;
      dic[nome].push(entrada);
    }
    return indices[nome][chave];
  }

  // Só os últimos N dias até "agora" (o painel mostra hoje, D-1 e a tendência).
  var dias = opcoes.diasHistorico || 7;
  var hoje = opcoes.agora.slice(0, 10);
  var limite = new Date(Date.UTC(+hoje.slice(0, 4), +hoje.slice(5, 7) - 1, +hoje.slice(8, 10) - (dias - 1)));
  var primeiroDia = limite.toISOString().slice(0, 10);

  var validas = [];
  for (var r = 1; r < tabela.length; r++) {
    var l = tabela[r];
    var data = normalizarData(l[col.data]);
    var slot = normalizarHora(l[col.slot]);
    if (!data || slot === null || data < primeiroDia || data > hoje) continue;
    validas.push({ l: l, data: data, slot: minutosParaTexto(slot) });
  }
  validas.sort(function (a, b) { return a.data < b.data ? -1 : a.data > b.data ? 1 : a.slot < b.slot ? -1 : a.slot > b.slot ? 1 : 0; });

  var linhas = validas.map(function (v) {
    var l = v.l;
    var txt = function (nome) { return String(l[col[nome]] || '').trim(); };
    var iReg = indice('regional', txt('regional'), txt('regional'));
    var iGer = indice('gerente', txt('gerente'), [txt('gerente'), iReg]);
    var iCoo = indice('coordenacao', txt('coordenacao'), [txt('coordenacao'), iGer]);
    var iSup = indice('supervisao', txt('supervisao'), [txt('supervisao'), iCoo]);
    var iTec = indice('tecnico', txt('tecnico'), [txt('tecnico'), iSup]);
    return [
      Number(txt('id_os').replace(/\D/g, '')),
      indice('datas', v.data, v.data),
      indice('cidade', txt('cidade'), txt('cidade')),
      iTec,
      indice('tipo', txt('tipo'), txt('tipo')),
      indice('slot', v.slot, v.slot),
      normalizarHora(l[col.chegada]),
      indice('status', txt('status'), txt('status'))
    ];
  });

  return {
    versao: opcoes.versao,
    agora: opcoes.agora,
    regras: opcoes.regras,
    dic: dic,
    colunas: ['id', 'data', 'cidade', 'tecnico', 'tipo', 'slot', 'chegadaMin', 'status'],
    linhas: linhas
  };
}

var COLUNAS_BASE = ['id_os', 'data', 'regional', 'cidade', 'gerente', 'coordenacao',
  'supervisao', 'tecnico', 'tipo', 'slot', 'chegada', 'status'];
