<#
.SYNOPSIS
    Relatório "Chegada ao Slot": CSV -> matriz Excel (Power Query + VBA) -> Google Sheets (Apps Script).

.DESCRIPTION
    Etapas (cada uma registrada no log com duração e erro):
      1. Ler configuração (config.json; se não existir, config.example.json em modo -DryRun)
      2. Abrir a matriz via COM, numa instância isolada e invisível do Excel
      3. Apontar a consulta Power Query para o CSV do dia e gravar "agora" nos parâmetros
      4. Atualizar o Power Query e esperar terminar
      5. Rodar a macro mdlSlot.CalcularIndicadores (Function que devolve "OK|..." ou "ERRO|...")
      6. Ler a aba Indicadores de uma vez (Value2)
      7. Conferir: recalcula o indicador em PowerShell e compara com o Excel
      8. Enviar ao Apps Script (doPost) ou, em -DryRun, só gravar o JSON em out\
    A matriz é fechada SEM salvar e o Excel é encerrado mesmo se alguma etapa falhar.

.PARAMETER ConfigPath
    Caminho do config.json (padrão: excel-pipeline\config.json).

.PARAMETER DryRun
    Executa tudo, menos o envio: grava o payload em out\payload_<data>.json.

.PARAMETER Visible
    Mostra a janela do Excel (útil para depurar).

.EXAMPLE
    .\excel-pipeline\Invoke-SlotPipeline.ps1 -DryRun
.EXAMPLE
    .\excel-pipeline\Invoke-SlotPipeline.ps1 -ConfigPath D:\config\slot.json
#>
[CmdletBinding()]
param(
    [string]$ConfigPath,
    [switch]$DryRun,
    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
Import-Module (Join-Path $root 'lib\FieldOps.psm1') -Force

# ---------------------------------------------------------------------------
# Funções do pipeline
# ---------------------------------------------------------------------------

function Read-PipelineConfig {
    param([string]$Path)
    if (-not $Path) { $Path = Join-Path $root 'config.json' }
    if (-not (Test-Path $Path)) {
        $example = Join-Path $root 'config.example.json'
        if (-not $script:DryRunMode) {
            throw "Configuração não encontrada: $Path. Copie config.example.json para config.json e preencha url e token."
        }
        Write-StepLog -Step 'Configuração' -Level AVISO -Message 'config.json ausente; usando config.example.json (modo -DryRun).'
        $Path = $example
    }
    $config = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) | ConvertFrom-Json

    # Caminhos relativos são relativos à pasta do pipeline, não à pasta de onde o script foi chamado.
    $resolve = { param($p) if ([IO.Path]::IsPathRooted($p)) { $p } else { [IO.Path]::GetFullPath((Join-Path $root $p)) } }
    $config.matriz = & $resolve $config.matriz
    $config.csv = & $resolve $config.csv
    $config.pastaLogs = & $resolve $config.pastaLogs
    $config.pastaSaida = & $resolve $config.pastaSaida
    foreach ($file in $config.matriz, $config.csv) {
        if (-not (Test-Path $file)) { throw "Arquivo não encontrado: $file" }
    }
    return $config
}

function Get-ReferenceNow {
    param([object]$Config)
    if ($Config.agora) {
        $value = ConvertTo-DateSafe ($Config.agora -replace 'T', ' ')
        if (-not $value) { throw "Valor inválido em 'agora': $($Config.agora)" }
        return $value
    }
    return Get-Date
}

function Set-WorkbookInputs {
    param($Workbook, [object]$Config, [datetime]$Now)
    $query = $Workbook.Queries.Item($Config.consulta)
    $query.Formula = Set-MTextParameter -Formula $query.Formula -Name 'CaminhoCsv' -Value $Config.csv
    Set-CellValue2 -Range $Workbook.Worksheets.Item('Parametros').Range('prmAgora') -Value $Now.ToOADate()
}

function Read-IndicatorSheet {
    param($Workbook)
    $range = $Workbook.Worksheets.Item('Indicadores').Range('A1').CurrentRegion
    return @(ConvertFrom-Value2Table (Get-CellValue2 -Range $range))
}

# Recalcula em PowerShell e compara com o Excel: os dois caminhos precisam dar o mesmo número.
function Test-IndicatorConsistency {
    param([object[]]$ExcelRows, [object]$Config, [datetime]$Now)
    $orders = (Import-ServiceOrder -Path $Config.csv).Orders | Where-Object { $_.Data -eq $Now.Date }
    $orders = @($orders)
    $checks = @(
        @{ Grupo = 'Total'; Chave = 'Total'; Lista = $orders }
    )
    foreach ($g in $orders | Group-Object Regional) { $checks += @{ Grupo = 'Regional'; Chave = $g.Name; Lista = $g.Group } }
    foreach ($g in $orders | Group-Object Slot)     { $checks += @{ Grupo = 'Slot'; Chave = $g.Name; Lista = $g.Group } }
    foreach ($g in $orders | Group-Object Tipo)     { $checks += @{ Grupo = 'Tipo'; Chave = $g.Name; Lista = $g.Group } }

    $divergent = @()
    foreach ($c in $checks) {
        $ps = Get-PunctualitySummary -Orders @($c.Lista) -Now $Now
        $xl = $ExcelRows | Where-Object { $_.grupo -eq $c.Grupo -and [string]$_.chave -eq $c.Chave } | Select-Object -First 1
        $xlPct = if ($xl -and $null -ne $xl.percentual) { Get-RoundedHalfUp $xl.percentual 2 } else { $null }
        $samePct = ($null -eq $xlPct -and $null -eq $ps.Percentual) -or ($null -ne $xlPct -and $xlPct -eq $ps.Percentual)
        if (-not $xl -or -not $samePct -or $xl.apuradas -ne $ps.Apuradas -or $xl.pendentes -ne $ps.Pendentes) {
            $divergent += '{0} {1}: Excel {2}% ({3} apuradas) x PowerShell {4}% ({5} apuradas)' -f
                $c.Grupo, $c.Chave, $xlPct, $(if ($xl) { $xl.apuradas } else { '-' }), $ps.Percentual, $ps.Apuradas
        }
    }
    if ($divergent) { throw ("Excel e PowerShell divergem:`n  " + ($divergent -join "`n  ")) }
    return $checks.Count
}

function New-Payload {
    param([object[]]$Rows, [object]$Config, [datetime]$Now)
    $inv = [Globalization.CultureInfo]::InvariantCulture
    return [ordered]@{
        token          = $Config.appsScript.token
        relatorio      = $Config.relatorio
        dataReferencia = $Now.ToString('yyyy-MM-dd', $inv)
        agora          = $Now.ToString('yyyy-MM-ddTHH:mm', $inv)
        geradoEm       = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss', $inv)
        colunas        = @('grupo', 'chave', 'apuradas', 'no_prazo', 'fora', 'pendentes', 'percentual', 'selo', 'percentual_d1', 'variacao_pp')
        linhas         = @($Rows | ForEach-Object {
            $r = $_
            , @($r.grupo, [string]$r.chave, [int]$r.apuradas, [int]$r.no_prazo, [int]$r.fora, [int]$r.pendentes,
                $r.percentual, $r.selo, $r.percentual_d1, $r.variacao_pp)
        })
    }
}

# ---------------------------------------------------------------------------
# Execução
# ---------------------------------------------------------------------------

$script:DryRunMode = [bool]$DryRun
[void](Disable-QuickEdit)

$config = Read-PipelineConfig -Path $ConfigPath
[void](Start-StepLog -Directory $config.pastaLogs -Name 'slot-pipeline')
$now = Get-ReferenceNow -Config $config
Write-StepLog -Step 'Configuração' -Message ('matriz={0} | csv={1} | agora={2:dd/MM/yyyy HH:mm}{3}' -f
    (Split-Path -Leaf $config.matriz), (Split-Path -Leaf $config.csv), $now, $(if ($DryRun) { ' | DRY-RUN' } else { '' }))

$excel = $null; $workbook = $null
$watch = [Diagnostics.Stopwatch]::StartNew()
try {
    $excel = Invoke-Step 'Abrir Excel' { New-ExcelApplication -Visible:$Visible }
    $workbook = Invoke-Step 'Abrir matriz' { $excel.Workbooks.Open($config.matriz, 0, $false) }
    Invoke-Step 'Definir entradas' { Set-WorkbookInputs -Workbook $workbook -Config $config -Now $now }
    Invoke-Step 'Atualizar Power Query' { Update-ExcelQuery -Excel $excel -Workbook $workbook }

    $status = Invoke-Step 'Rodar macro' { Invoke-ExcelMacro -Excel $excel -Workbook $workbook -Macro $config.macro }
    if ($status -notlike 'OK|*') { throw "Macro retornou: $status" }
    $parts = $status -split '\|'
    Write-StepLog -Step 'Rodar macro' -Message ("{0} ordens lidas, {1} grupos calculados em {2} s" -f $parts[1], $parts[2], $parts[3])

    $rows = Invoke-Step 'Ler indicadores' { Read-IndicatorSheet -Workbook $workbook }
    $total = $rows | Where-Object grupo -eq 'Total' | Select-Object -First 1
    Write-StepLog -Step 'Ler indicadores' -Message ('{0} linhas | Total: {1:N2}% ({2}) | variação {3:+0.00;-0.00;0.00} p.p.' -f
        $rows.Count, $total.percentual, $total.selo, $total.variacao_pp)

    $checked = Invoke-Step 'Conferir Excel x PowerShell' { Test-IndicatorConsistency -ExcelRows $rows -Config $config -Now $now }
    Write-StepLog -Step 'Conferir Excel x PowerShell' -Message "$checked grupos idênticos"

    $payload = New-Payload -Rows $rows -Config $config -Now $now
    if ($DryRun) {
        Invoke-Step 'Gravar payload (dry-run)' {
            if (-not (Test-Path $config.pastaSaida)) { New-Item -ItemType Directory -Path $config.pastaSaida | Out-Null }
            $safe = [ordered]@{}; foreach ($k in $payload.Keys) { $safe[$k] = $payload[$k] }
            $safe.token = '***'   # nunca gravar o token em arquivo
            $file = Join-Path $config.pastaSaida ('payload_{0:yyyy-MM-dd}.json' -f $now)
            [IO.File]::WriteAllText($file, (ConvertTo-CompactJson $safe), (New-Object Text.UTF8Encoding($false)))
            Write-StepLog -Step 'Gravar payload (dry-run)' -Message "Envio simulado; payload em $file"
        }
    }
    else {
        $answer = Invoke-Step 'Enviar ao Apps Script' {
            Send-AppsScriptPayload -Url $config.appsScript.url -Payload $payload `
                -TimeoutSec $config.appsScript.timeoutSegundos -Attempts $config.appsScript.tentativas
        }
        Write-StepLog -Step 'Enviar ao Apps Script' -Message "Resposta: $($answer.mensagem)"
    }

    Write-StepLog -Step 'Fim' -Level OK -Message ('Relatório concluído em {0:N1} s' -f $watch.Elapsed.TotalSeconds)
}
catch {
    Write-StepLog -Step 'Fim' -Level ERRO -Message ('Relatório interrompido: {0}' -f (Get-ErrorDetail $_))
    throw
}
finally {
    # Fecha sem salvar: a matriz é um modelo; o resultado vai para o Sheets.
    Close-ExcelSafely -Excel $excel -Workbook $workbook
}
