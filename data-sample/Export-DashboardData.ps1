<#
.SYNOPSIS
    Converte o CSV de ordens em JSON compactado (dicionário + linhas) para o painel.

.DESCRIPTION
    Um JSON "ingênuo" (lista de objetos) repete os nomes das colunas e os textos longos
    (gerente, supervisor, cidade) em todas as linhas. No formato compactado:
      - cada texto aparece UMA vez em "dic" e as linhas guardam só o índice;
      - a hierarquia (técnico > supervisão > coordenação > gerente > regional) fica no
        dicionário, então a linha só precisa do índice do técnico;
      - horas viram minutos desde 00:00 (número inteiro em vez de "HH:mm").
    O script mostra o tamanho das duas versões para comparação.

    Também grava "versao.json", um arquivo minúsculo que o painel consulta periodicamente
    para saber se precisa baixar a base inteira de novo.

.PARAMETER CsvPath
    CSV gerado por New-FakeServiceOrders.ps1.

.PARAMETER OutputDirectory
    Pasta de saída (padrão: dashboard-pages\data).

.PARAMETER Now
    "Agora" da simulação. Padrão: última data do CSV às -ReferenceTime.

.PARAMETER ReferenceTime
    Hora usada quando -Now não é informado. Use a mesma do gerador (padrão 14:35).

.EXAMPLE
    .\Export-DashboardData.ps1
#>
[CmdletBinding()]
param(
    [string]$CsvPath,
    [string]$OutputDirectory,
    [Nullable[datetime]]$Now,
    [ValidatePattern('^\d{1,2}:\d{2}$')][string]$ReferenceTime = '14:35'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# Padrões calculados aqui, e não no param: no PowerShell 5.1 com -File, $PSScriptRoot chega vazio no bloco param.
if (-not $CsvPath) { $CsvPath = Join-Path $PSScriptRoot 'orders-sample.csv' }
if (-not $OutputDirectory) { $OutputDirectory = Join-Path (Split-Path -Parent $PSScriptRoot) 'dashboard-pages\data' }
Import-Module (Join-Path $PSScriptRoot '..\excel-pipeline\lib\FieldOps.psm1') -Force

# Devolve o índice de $Key no dicionário, criando a entrada se for nova.
function Get-DictionaryIndex {
    param(
        [Parameter(Mandatory)][hashtable]$Index,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$List,
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)][object]$Entry
    )
    if (-not $Index.ContainsKey($Key)) {
        $Index[$Key] = $List.Count
        $List.Add($Entry)
    }
    return $Index[$Key]
}

function New-Dictionary {
    return @{ Index = @{}; List = New-Object 'System.Collections.Generic.List[object]' }
}

function ConvertTo-MinuteOfDay {
    param([AllowNull()][object]$Moment)
    if ($null -eq $Moment) { return $null }
    return [int]([datetime]$Moment).TimeOfDay.TotalMinutes
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$inv = [Globalization.CultureInfo]::InvariantCulture

$import = Invoke-Step 'Ler e tratar CSV' { Import-ServiceOrder -Path $CsvPath }
$orders = @($import.Orders | Sort-Object Data, Slot, IdOs)
Write-StepLog -Step 'Ler e tratar CSV' -Message ('{0} válidas, {1} duplicadas removidas, {2} rejeitadas' -f
    $import.Stats.Validas, $import.Stats.Duplicadas, $import.Stats.Rejeitadas)

if (-not $Now) { $Now = $orders[-1].Data.Add([timespan]::Parse($ReferenceTime)) }   # $orders já está ordenado por data
$rule = Get-DefaultIndicatorRule

$compact = Invoke-Step 'Montar dicionário + linhas' {
    $dic = @{}
    foreach ($name in 'datas', 'regional', 'gerente', 'coordenacao', 'supervisao', 'tecnico', 'cidade', 'tipo', 'slot', 'status') {
        $dic[$name] = New-Dictionary
    }
    $rows = New-Object 'System.Collections.Generic.List[object]'

    foreach ($o in $orders) {
        $iReg  = Get-DictionaryIndex $dic.regional.Index $dic.regional.List $o.Regional $o.Regional
        $iGer  = Get-DictionaryIndex $dic.gerente.Index $dic.gerente.List $o.Gerente @($o.Gerente, $iReg)
        $iCoo  = Get-DictionaryIndex $dic.coordenacao.Index $dic.coordenacao.List $o.Coordenacao @($o.Coordenacao, $iGer)
        $iSup  = Get-DictionaryIndex $dic.supervisao.Index $dic.supervisao.List $o.Supervisao @($o.Supervisao, $iCoo)
        $iTec  = Get-DictionaryIndex $dic.tecnico.Index $dic.tecnico.List $o.Tecnico @($o.Tecnico, $iSup)
        $iData = Get-DictionaryIndex $dic.datas.Index $dic.datas.List ($o.Data.ToString('yyyy-MM-dd', $inv)) ($o.Data.ToString('yyyy-MM-dd', $inv))
        $iCid  = Get-DictionaryIndex $dic.cidade.Index $dic.cidade.List $o.Cidade $o.Cidade
        $iTipo = Get-DictionaryIndex $dic.tipo.Index $dic.tipo.List $o.Tipo $o.Tipo
        $iSlot = Get-DictionaryIndex $dic.slot.Index $dic.slot.List $o.Slot $o.Slot
        $iStat = Get-DictionaryIndex $dic.status.Index $dic.status.List $o.Status $o.Status

        $id = [int]($o.IdOs -replace '\D', '')
        $rows.Add(@($id, $iData, $iCid, $iTec, $iTipo, $iSlot, (ConvertTo-MinuteOfDay $o.Chegada), $iStat))
    }

    $dictOut = [ordered]@{}
    foreach ($name in 'datas', 'regional', 'gerente', 'coordenacao', 'supervisao', 'tecnico', 'cidade', 'tipo', 'slot', 'status') {
        $dictOut[$name] = $dic[$name].List
    }

    [ordered]@{
        versao   = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss', $inv)
        agora    = ([datetime]$Now).ToString('yyyy-MM-ddTHH:mm', $inv)
        regras   = [ordered]@{
            janelaMinutos     = $rule.JanelaMinutos
            toleranciaMinutos = [ordered]@{ 'Instalação' = $rule.ToleranciaMinutos['Instalação']; 'Reparo' = $rule.ToleranciaMinutos['Reparo'] }
            meta              = $rule.Meta
            margemPerto       = $rule.MargemPerto
        }
        dic      = $dictOut
        colunas  = @('id', 'data', 'cidade', 'tecnico', 'tipo', 'slot', 'chegadaMin', 'status')
        linhas   = $rows
    }
}

$json = Invoke-Step 'Serializar JSON' { ConvertTo-CompactJson $compact }

# Versão "ingênua" só para comparar tamanhos (não é gravada).
$naiveBytes = Invoke-Step 'Medir JSON ingênuo' {
    $naive = foreach ($o in $orders) {
        [ordered]@{
            id_os = $o.IdOs; data = $o.Data.ToString('yyyy-MM-dd', $inv); regional = $o.Regional; cidade = $o.Cidade
            gerente = $o.Gerente; coordenacao = $o.Coordenacao; supervisao = $o.Supervisao; tecnico = $o.Tecnico
            tipo = $o.Tipo; slot = $o.Slot
            chegada = if ($o.Chegada) { $o.Chegada.ToString('HH:mm', $inv) } else { $null }
            status = $o.Status
        }
    }
    $utf8NoBom.GetByteCount((ConvertTo-CompactJson @($naive)))
}

if (-not (Test-Path $OutputDirectory)) { New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null }
$dataPath = Join-Path $OutputDirectory 'slot-data.json'
$versionPath = Join-Path $OutputDirectory 'versao.json'
[IO.File]::WriteAllText($dataPath, $json, $utf8NoBom)
[IO.File]::WriteAllText($versionPath, (ConvertTo-CompactJson ([ordered]@{ versao = $compact.versao })), $utf8NoBom)

$compactBytes = $utf8NoBom.GetByteCount($json)
Write-StepLog -Step 'Resultado' -Level OK -Message ('{0} ordens | ingênuo {1:N0} KB -> compactado {2:N0} KB ({3:P0} menor)' -f
    $orders.Count, ($naiveBytes / 1KB), ($compactBytes / 1KB), (1 - $compactBytes / $naiveBytes))
Write-StepLog -Step 'Resultado' -Message "Gravado: $dataPath"
