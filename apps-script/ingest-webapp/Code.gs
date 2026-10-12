/**
 * Web App de ingestão: recebe os envios do PowerShell (doPost) e grava no Google Sheets.
 *
 * Relatórios aceitos (campo "relatorio" do JSON):
 *   - "chegada-slot": indicadores do dia -> aba Historico (substitui o mesmo dia) e aba Indicadores (retrato atual)
 *   - "ordens":       base de ordens tratadas -> aba Base (lida pelo painel slot-dashboard)
 *
 * Segurança:
 *   - O token fica em Propriedades do script (INGEST_TOKEN), nunca no código.
 *     O doPost não tem acesso aos cabeçalhos HTTP, por isso o token vem no corpo JSON.
 *   - O ID da planilha fica em PLANILHA_ID (Propriedades do script).
 *   - LockService impede que dois envios simultâneos gravem ao mesmo tempo.
 *
 * O ContentService sempre responde HTTP 200; o resultado vai em {"ok": true|false}.
 */

var ABAS = {
  historico: 'Historico',
  indicadores: 'Indicadores',
  base: 'Base',
  controle: 'Controle'
};

var CABECALHO_HISTORICO = ['data_referencia', 'agora', 'relatorio'].concat(COLUNAS_INDICADORES);

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var payload = JSON.parse((e && e.postData && e.postData.contents) || 'null');
    var props = PropertiesService.getScriptProperties();
    var validacao = validarPayload(payload, props.getProperty('INGEST_TOKEN'));
    if (!validacao.ok) return responder({ ok: false, erro: validacao.erro });

    if (!lock.tryLock(30000)) return responder({ ok: false, erro: 'Outro envio em andamento; tente de novo.' });

    var planilha = SpreadsheetApp.openById(props.getProperty('PLANILHA_ID'));
    var mensagem = payload.relatorio === 'ordens'
      ? gravarBaseOrdens(planilha, payload)
      : gravarIndicadores(planilha, payload);

    registrarVersao(planilha, payload);
    return responder({ ok: true, mensagem: mensagem, linhas: payload.linhas.length });
  } catch (err) {
    console.error(err);
    return responder({ ok: false, erro: 'Falha ao processar: ' + err.message });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/** Upsert do dia na aba Historico + retrato atual na aba Indicadores. */
function gravarIndicadores(planilha, payload) {
  var historico = obterAba(planilha, ABAS.historico, CABECALHO_HISTORICO);
  removerLinhasDoDia(historico, payload.dataReferencia, payload.relatorio);

  var novas = montarLinhasHistorico(payload);
  var inicio = historico.getLastRow() + 1;
  // Datas e chave como texto: senão o Sheets converte "2026-10-12" em data e "08:30" em hora,
  // exatamente como o Excel faz. Números continuam números (fórmulas e gráficos funcionam).
  historico.getRange(inicio, 1, novas.length, 5).setNumberFormat('@');
  historico.getRange(inicio, 1, novas.length, CABECALHO_HISTORICO.length).setValues(novas);
  historico.getRange(inicio, 10, novas.length, 1).setNumberFormat('0.00');
  historico.getRange(inicio, 12, novas.length, 2).setNumberFormat('0.00');

  var indicadores = obterAba(planilha, ABAS.indicadores, null);
  indicadores.clear();
  var tabela = [COLUNAS_INDICADORES].concat(payload.linhas.map(function (l) {
    return l.map(function (v) { return v === null ? '' : v; });
  }));
  indicadores.getRange(1, 2, tabela.length, 1).setNumberFormat('@');
  indicadores.getRange(1, 1, tabela.length, COLUNAS_INDICADORES.length).setValues(tabela);
  indicadores.getRange(2, 7, tabela.length - 1, 1).setNumberFormat('0.00');
  indicadores.getRange(2, 9, tabela.length - 1, 2).setNumberFormat('0.00');
  indicadores.getRange(1, 1, 1, COLUNAS_INDICADORES.length).setFontWeight('bold');
  indicadores.getRange(1, 12).setValue('Atualizado em: ' + (payload.agora || payload.dataReferencia));
  indicadores.setFrozenRows(1);

  return 'Indicadores de ' + payload.dataReferencia + ' gravados (' + payload.linhas.length + ' linhas).';
}

/** Substitui a aba Base inteira em uma única escrita (setValues), sempre como texto. */
function gravarBaseOrdens(planilha, payload) {
  var base = obterAba(planilha, ABAS.base, null);
  base.clearContents();
  var tabela = [COLUNAS_ORDENS].concat(payload.linhas.map(function (l) {
    return l.map(function (v) { return v === null ? '' : String(v); });
  }));
  base.getRange(1, 1, tabela.length, COLUNAS_ORDENS.length).setNumberFormat('@').setValues(tabela);
  base.setFrozenRows(1);
  return 'Base de ordens substituída (' + payload.linhas.length + ' ordens).';
}

/** A aba Controle guarda a versão que o painel consulta (?acao=base) a cada 90 s. */
function registrarVersao(planilha, payload) {
  var controle = obterAba(planilha, ABAS.controle, null);
  var agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  controle.getRange('A1:B3').setNumberFormat('@').setValues([
    ['versao', agora],
    ['ultimo_relatorio', payload.relatorio],
    ['data_referencia', payload.dataReferencia]
  ]);
}

function removerLinhasDoDia(aba, dataReferencia, relatorio) {
  var ultima = aba.getLastRow();
  if (ultima < 2) return;
  var valores = aba.getRange(2, 1, ultima - 1, 3).getDisplayValues();
  // De baixo para cima, para os índices não mudarem ao apagar.
  for (var i = valores.length - 1; i >= 0; i--) {
    if (valores[i][0] === dataReferencia && valores[i][2] === relatorio) aba.deleteRow(i + 2);
  }
}

function obterAba(planilha, nome, cabecalho) {
  var aba = planilha.getSheetByName(nome) || planilha.insertSheet(nome);
  if (cabecalho && aba.getLastRow() === 0) {
    aba.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]).setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  return aba;
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// Utilitários para rodar à mão no editor (menu Executar)
// ---------------------------------------------------------------------------

/** Gera um token aleatório, salva em INGEST_TOKEN e mostra no log (copie para o config.json). */
function gerarToken() {
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('INGEST_TOKEN', token);
  console.log('Token salvo em Propriedades do script. Copie para o config.json: ' + token);
}

/** Simula um envio sem sair do editor, para testar permissões e gravação. */
function testarDoPostLocal() {
  var token = PropertiesService.getScriptProperties().getProperty('INGEST_TOKEN');
  var corpo = {
    token: token, relatorio: 'chegada-slot', dataReferencia: '2026-10-12', agora: '2026-10-12T14:35',
    colunas: COLUNAS_INDICADORES,
    linhas: [['Total', 'Total', 159, 140, 19, 64, 88.05, 'Na meta', 80.67, 7.38]]
  };
  var resposta = doPost({ postData: { contents: JSON.stringify(corpo) } });
  console.log(resposta.getContent());
}
