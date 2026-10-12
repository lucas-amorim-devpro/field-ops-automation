# Prompt para o Claude (web): implantar e testar a parte Google Apps Script

> **Como usar:** abra uma conversa nova em claude.ai, **anexe os 11 arquivos listados abaixo**
> e cole o texto da seção "Prompt" inteiro. Não cole tokens nem IDs reais na conversa.

## Arquivos para anexar

De `apps-script/ingest-webapp/`:
1. `Code.gs`
2. `Validacao.gs`
3. `appsscript.json`

De `apps-script/slot-dashboard/`:
4. `Code.gs`
5. `Compactacao.gs`
6. `CacheEmPedacos.gs`
7. `appsscript.json`
8. `Index.html`
9. `Styles.html`, `Core.html` e `App.html` (gerados; pode anexar os três)

De `excel-pipeline/`:
10. `config.example.json`
11. `Invoke-SlotPipeline.ps1` (para contexto do envio)

---

## Prompt

```text
Contexto
--------
Estou montando um repositório de portfólio ("field-ops-automation") com dados 100% fictícios.
Um pipeline PowerShell (Windows, PowerShell 5.1) abre uma matriz Excel via COM, atualiza
Power Query, roda uma macro VBA e envia os resultados em JSON para um Web App do Google Apps
Script (doPost). Um segundo Web App (doGet) serve um painel HTML que lê a planilha.

A parte local já foi construída e testada com o Claude Code no meu PC:
- 92 testes Pester passando (incluindo o pipeline completo com o Excel real em modo -DryRun);
- as funções PURAS dos .gs (Validacao.gs, Compactacao.gs, CacheEmPedacos.gs) passaram em 36
  testes no navegador, inclusive validando os payloads reais gerados pelo PowerShell;
- o Index.html gerado foi testado no navegador com google.script.run simulado.

O que NÃO pôde ser testado localmente, e é o que preciso de você:
criar, configurar, implantar e testar de ponta a ponta no Google Apps Script/Sheets de verdade.
Os arquivos estão anexados.

Arquitetura
-----------
1) Planilha Google (nova, só para a demo) com as abas: Base, Historico, Indicadores, Controle.
   As abas são criadas automaticamente pelo ingest-webapp na primeira gravação.
2) Projeto Apps Script "ingest-webapp" (Code.gs + Validacao.gs + appsscript.json):
   - doPost recebe {token, relatorio, dataReferencia, agora, colunas, linhas}.
   - relatorio "chegada-slot" -> Historico (substitui o mesmo dia) + Indicadores (retrato).
   - relatorio "ordens" -> Base (substitui tudo; ~1.700 linhas, ~265 KB de JSON).
   - Grava a versão em Controle!B1 (o painel consulta essa célula a cada 90 s).
   - Propriedades do script: INGEST_TOKEN (gerado por gerarToken()) e PLANILHA_ID.
   - Implantação: executar como "Eu", acesso "Qualquer pessoa" (o PowerShell não faz login;
     a proteção é o token no corpo, porque o doPost não enxerga cabeçalhos HTTP).
3) Projeto Apps Script "slot-dashboard" (Code.gs, Compactacao.gs, CacheEmPedacos.gs,
   Index.html, Styles.html, Core.html, App.html, appsscript.json):
   - doGet entrega o HTML com os dados embutidos na 1ª carga; depois o navegador usa
     google.script.run.obterVersao() a cada 90 s e obterDados() quando a versão muda.
   - Cache em pedaços (CacheService, 100 KB por chave, pedaços de 45 mil caracteres).
   - Propriedade do script: PLANILHA_ID. Escopo só leitura (spreadsheets.readonly).
   - Deve poder ser incorporado no Google Sites (XFrameOptionsMode.ALLOWALL).

O que eu quero que você faça
----------------------------
Me guie passo a passo, UMA etapa por vez, esperando eu confirmar ("feito" / colar o erro)
antes de seguir. Em cada etapa diga exatamente onde clicar (a interface do Apps Script está
em português) e o que devo ver se deu certo.

Etapa 0 - Revisão de código (antes de eu colar qualquer coisa)
  Revise os .gs e .html anexados procurando problemas que só aparecem no Apps Script real:
  - compatibilidade V8; ordem de carregamento entre arquivos .gs (variáveis globais usadas
    antes de definidas, ex.: COLUNAS_INDICADORES definida em Validacao.gs e usada em Code.gs);
  - LockService (tryLock/hasLock/releaseLock), CacheService.putAll/getAll e seus limites;
  - escopos em appsscript.json (spreadsheets / spreadsheets.readonly) são suficientes para
    tudo que o código chama? (Session.getScriptTimeZone, Utilities, ContentService, HtmlService)
  - HtmlService: template com <?!= include('...'); ?> e <?!= dadosIniciais ?>; o JSON vai
    embutido num <script> com '<' trocado por < - isso é seguro?
  - "access": "ANYONE" vs "ANYONE_ANONYMOUS" nos manifestos: qual usar em uma conta Gmail
    pessoal para (a) o doPost chamado pelo PowerShell sem login e (b) o painel público.
  Liste só problemas concretos, com a correção exata (trecho antes/depois). Se não achar
  nada, diga isso.

Etapa 1 - Criar a planilha de demonstração e anotar o ID (sem me pedir para colar o ID aqui).
Etapa 2 - Criar o projeto ingest-webapp: arquivos, manifesto (mostrar appsscript.json em
          Configurações do projeto), Propriedades do script (PLANILHA_ID), rodar gerarToken()
          e autorizar. Como copiar o token do log para o meu config.json local sem expor.
Etapa 3 - Rodar testarDoPostLocal() no editor e conferir as abas Historico, Indicadores e
          Controle na planilha.
Etapa 4 - Implantar como App da Web e testar a URL /exec com o PowerShell:
            1) copiar excel-pipeline\config.example.json para config.json e preencher url/token;
            2) .\excel-pipeline\Send-OrdersBase.ps1          (envia a aba Base)
            3) .\excel-pipeline\Invoke-SlotPipeline.ps1      (envia os indicadores)
          Explique o redirecionamento 302 de script.google.com para script.googleusercontent.com
          e o que fazer se vier HTML de login em vez de JSON.
Etapa 5 - Criar o projeto slot-dashboard (mesma planilha, só leitura), implantar e abrir a URL.
          Conferir: o painel mostra os mesmos números do PowerShell (no dia de referência dos
          dados de exemplo: 88,05% no geral, "Na meta", +7,38 p.p.)*. Testar a troca de versão
          reenviando a base (Send-OrdersBase.ps1) com o painel aberto: em até 90 s ele deve
          atualizar sozinho.
          *Observação: fora da demo o painel usa o relógio real. Se a data de hoje não for a
           data dos dados, os números de "hoje" ficam vazios. Me ajude a gerar dados para a
           data atual com: .\data-sample\New-FakeServiceOrders.ps1 (ReferenceDate = hoje).
Etapa 6 - Medir: tempo da 1ª abertura (sem cache) e da 2ª (com cache), e quantos pedaços o
          cache usou (sugira um console.log temporário em obterDados).
Etapa 7 - (Opcional) Incorporar o painel em uma página do Google Sites e verificar se o
          tooltip e o tema funcionam dentro do iframe.

Regras
------
- Nunca me peça para colar na conversa: token, ID da planilha, URL /exec ou e-mail.
  Use marcadores como <SEU_ID>.
- Se algo exigir mudar código, me dê o arquivo/trecho completo corrigido e diga se a mudança
  precisa ir também para o repositório (os arquivos Styles/Core/App/Index.html do
  slot-dashboard são GERADOS a partir de dashboard-pages/ por tools/Build-AppsScriptDashboard.ps1;
  correções de interface devem ser feitas em dashboard-pages/ e não nos gerados).
- Lembre que toda alteração de código no Apps Script só vale na URL /exec depois de
  "Gerenciar implantações > Editar > Nova versão".

Ao final
--------
Escreva um RESUMO PARA O CLAUDE CODE em um bloco de código, com:
- quais etapas funcionaram e quais não;
- todas as mudanças de código feitas (arquivo + trecho antes/depois);
- os tempos medidos na etapa 6;
- a URL pública do painel SOMENTE se eu disser que quero publicá-la no README.
Eu vou colar esse resumo de volta no Claude Code para atualizar o repositório e o README.
```

---

## Resultado (12/10/2026)

O teste real foi feito seguindo este prompt. Etapas 0 a 6 concluídas; a 7 (Google Sites) ficou
para depois. A revisão e a implantação encontraram 4 problemas, já corrigidos no repositório:

1. `ingest-webapp/Code.gs`: variável global dependia da ordem de carregamento dos `.gs` → virou função `cabecalhoHistorico_()`.
2. `slot-dashboard/appsscript.json`: `"access": "ANYONE"` exigia login → `ANYONE_ANONYMOUS`.
3. `slot-dashboard/appsscript.json`: `SpreadsheetApp.openById` não aceita `spreadsheets.readonly` → escopo `spreadsheets` (o código só lê).
4. `slot-dashboard/Code.gs`: U+2028/U+2029 literais na regex → escapes como texto (e um teste que impede a volta).

Tempos medidos e demais resultados: seção "Teste de ponta a ponta no Google" do README.

Atualização: o token do `ingest-webapp` foi trocado depois do teste (o antigo passou por um chat).
`gerarToken()` deixou de imprimir o token no log de execução; o valor é copiado em
Configurações do projeto > Propriedades do script. A etapa 7 (Google Sites) foi pulada: a demo
pública é a do GitHub Pages, e o Google Sites é só onde o painel original do trabalho é publicado.
