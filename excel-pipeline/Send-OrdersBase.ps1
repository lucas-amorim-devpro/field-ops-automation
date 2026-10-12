<#
.SYNOPSIS
    Envia a base de ordens tratadas para a aba "Base" do Google Sheets (lida pelo painel slot-dashboard).

.DESCRIPTION
    Lê o CSV, limpa e remove duplicadas (Import-ServiceOrder) e envia ao mesmo Web App de
    ingestão, com relatorio = "ordens". Valores vão como TEXTO (datas dd/MM/yyyy, horas HH:mm),
    e o Web App grava a aba com formato texto, para o Sheets não converter "08:30" em hora.

.PARAMETER ConfigPath
    Caminho do config.json (padrão: excel-pipeline\config.json).

.PARAMETER DryRun
    Não envia: grava o payload em out\ordens_<data>.json (sem o token).

.EXAMPLE
    .\excel-pipeline\Send-OrdersBase.ps1 -DryRun
#>
[CmdletBinding()]
param(
    [string]$ConfigPath,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
Import-Module (Join-Path $root 'lib\FieldOps.psm1') -Force

if (-not $ConfigPath) { $ConfigPath = Join-Path $root 'config.json' }
if (-not (Test-Path $ConfigPath)) {
    if (-not $DryRun) { throw "Configuração não encontrada: $ConfigPath. Copie config.example.json para config.json." }
    $ConfigPath = Join-Path $root 'config.example.json'
}
$config = [IO.File]::ReadAllText($ConfigPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
$resolve = { param($p) if ([IO.Path]::IsPathRooted($p)) { $p } else { [IO.Path]::GetFullPath((Join-Path $root $p)) } }
$csv = & $resolve $config.csv
$outDir = & $resolve $config.pastaSaida
[void](Start-StepLog -Directory (& $resolve $config.pastaLogs) -Name 'ordens-base')

$inv = [Globalization.CultureInfo]::InvariantCulture
$import = Invoke-Step 'Ler e tratar CSV' { Import-ServiceOrder -Path $csv }
Write-StepLog -Step 'Ler e tratar CSV' -Message ('{0} ordens válidas ({1} duplicadas removidas, {2} rejeitadas)' -f
    $import.Stats.Validas, $import.Stats.Duplicadas, $import.Stats.Rejeitadas)

$orders = @($import.Orders)
$reference = ($orders | Sort-Object Data | Select-Object -Last 1).Data
$payload = [ordered]@{
    token          = $config.appsScript.token
    relatorio      = 'ordens'
    dataReferencia = $reference.ToString('yyyy-MM-dd', $inv)
    agora          = (Get-Date).ToString('yyyy-MM-ddTHH:mm', $inv)
    colunas        = @('id_os', 'data', 'regional', 'cidade', 'gerente', 'coordenacao', 'supervisao', 'tecnico', 'tipo', 'slot', 'chegada', 'status')
    linhas         = @($orders | ForEach-Object {
        , @($_.IdOs, $_.Data.ToString('dd/MM/yyyy', $inv), $_.Regional, $_.Cidade, $_.Gerente, $_.Coordenacao,
            $_.Supervisao, $_.Tecnico, $_.Tipo, $_.Slot,
            $(if ($_.Chegada) { $_.Chegada.ToString('HH:mm', $inv) } else { '' }), $_.Status)
    })
}

if ($DryRun) {
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
    $payload.token = '***'
    $file = Join-Path $outDir ('ordens_{0:yyyy-MM-dd}.json' -f $reference)
    [IO.File]::WriteAllText($file, (ConvertTo-CompactJson $payload), (New-Object Text.UTF8Encoding($false)))
    Write-StepLog -Step 'Fim' -Level OK -Message ('Envio simulado: {0} ordens em {1} ({2:N0} KB)' -f $orders.Count, $file, ((Get-Item $file).Length / 1KB))
}
else {
    $answer = Invoke-Step 'Enviar ao Apps Script' {
        Send-AppsScriptPayload -Url $config.appsScript.url -Payload $payload `
            -TimeoutSec $config.appsScript.timeoutSegundos -Attempts $config.appsScript.tentativas
    }
    Write-StepLog -Step 'Fim' -Level OK -Message "Resposta: $($answer.mensagem)"
}
