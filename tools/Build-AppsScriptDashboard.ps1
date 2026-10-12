<#
.SYNOPSIS
    Gera os arquivos HTML do painel no Apps Script a partir de dashboard-pages/.

.DESCRIPTION
    O HtmlService não serve arquivos .css/.js separados: tudo precisa estar em arquivos .html
    incluídos no template (<?!= include('Styles'); ?>). Este script empacota:
      dashboard-pages/css/styles.css -> apps-script/slot-dashboard/Styles.html
      dashboard-pages/js/core.js     -> apps-script/slot-dashboard/Core.html
      dashboard-pages/js/app.js      -> apps-script/slot-dashboard/App.html
      dashboard-pages/index.html     -> apps-script/slot-dashboard/Index.html (com a fonte de dados do Apps Script)
    Assim as duas versões do painel nunca divergem: edite só em dashboard-pages/ e rode este script.

.EXAMPLE
    .\tools\Build-AppsScriptDashboard.ps1
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path -Parent $PSScriptRoot
$src = Join-Path $repo 'dashboard-pages'
$dst = Join-Path $repo 'apps-script\slot-dashboard'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$banner = '<!-- GERADO por tools/Build-AppsScriptDashboard.ps1 a partir de dashboard-pages/. Não edite à mão. -->'

function Read-Source([string]$relative) {
    return [IO.File]::ReadAllText((Join-Path $src $relative), $utf8).Replace("`r`n", "`n").TrimEnd()
}

function Write-Generated([string]$name, [string]$content) {
    $path = Join-Path $dst $name
    [IO.File]::WriteAllText($path, ($content.Replace("`r`n", "`n") + "`n"), $utf8)
    Write-Host ('Gerado: apps-script\slot-dashboard\{0} ({1:N0} KB)' -f $name, ((Get-Item $path).Length / 1KB))
}

# Troca um trecho exato; falha alto se o trecho não existir (index.html mudou e o build precisa de ajuste).
function Set-Exact([string]$text, [string]$old, [string]$new) {
    if (-not $text.Contains($old)) { throw "Trecho não encontrado em index.html: $old" }
    return $text.Replace($old, $new)
}

Write-Generated 'Styles.html' ("$banner`n<style>`n" + (Read-Source 'css\styles.css') + "`n</style>")
Write-Generated 'Core.html'   ("$banner`n<script>`n" + (Read-Source 'js\core.js') + "`n</script>")
Write-Generated 'App.html'    ("$banner`n<script>`n" + (Read-Source 'js\app.js') + "`n</script>")

# Fonte de dados do Apps Script: 1ª carga embutida no HTML (dadosIniciais vem do doGet);
# depois, google.script.run.obterDados()/obterVersao(), sem CORS e sem expor a URL.
$dataSource = @'
<script>
    (function () {
      var inicial = <?!= dadosIniciais ?>;
      function chamar(funcao) {
        return new Promise(function (ok, falha) {
          google.script.run.withSuccessHandler(ok).withFailureHandler(falha)[funcao]();
        });
      }
      window.SLOT_DATA_SOURCE = {
        demo: false,
        load: function () {
          if (inicial) {
            var dados = inicial;
            inicial = null;
            return Promise.resolve({ json: dados, bytes: JSON.stringify(dados).length });
          }
          return chamar('obterDados').then(function (texto) { return { json: JSON.parse(texto), bytes: texto.length }; });
        },
        version: function () { return chamar('obterVersao'); }
      };
    })();
  </script>
  <?!= include('App'); ?>
'@.Trim()

$index = Read-Source 'index.html'
$index = Set-Exact $index '<link rel="stylesheet" href="css/styles.css">' "<?!= include('Styles'); ?>"
$index = Set-Exact $index '<script src="js/core.js"></script>' "<?!= include('Core'); ?>"
$index = Set-Exact $index '<script src="js/app.js"></script>' $dataSource
$index = $index.Replace('<!doctype html>', "<!doctype html>`n$banner")
Write-Generated 'Index.html' $index
