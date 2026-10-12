<#
.SYNOPSIS
    Monta a matriz Excel de demonstração (slot-matriz.xlsm) a partir dos arquivos-fonte do repositório.

.DESCRIPTION
    A matriz não é versionada (binário com macro); ela é reconstruída a partir de:
      - power-query\tbOrdens.pq   -> consulta Power Query carregada na aba "Base"
      - vba\mdlSlot.bas           -> módulo VBA importado
      - aba "Parametros"          -> células nomeadas lidas pela macro (prmAgora, prmMeta...)
      - aba "Indicadores"         -> saída da macro, lida pelo pipeline

    Requer, no Excel: Arquivo > Opções > Central de Confiabilidade > Configurações de Macro >
    "Confiar no acesso ao modelo de objeto do projeto do VBA" (só para importar o .bas).

.PARAMETER OutputPath
    Caminho da matriz gerada (padrão: excel-pipeline\matriz\slot-matriz.xlsm).

.PARAMETER CsvPath
    CSV usado na primeira carga (padrão: data-sample\orders-sample.csv).

.EXAMPLE
    .\excel-pipeline\Build-DemoWorkbook.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputPath,
    [string]$CsvPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# Padrões calculados aqui, e não no param: no PowerShell 5.1 com -File, $PSScriptRoot chega vazio no bloco param.
if (-not $OutputPath) { $OutputPath = Join-Path $PSScriptRoot 'matriz\slot-matriz.xlsm' }
if (-not $CsvPath) { $CsvPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'data-sample\orders-sample.csv' }
Import-Module (Join-Path $PSScriptRoot 'lib\FieldOps.psm1') -Force

$xlSrcExternal = 0
$xlYes = 1
$xlCmdSql = 2
$xlOpenXMLWorkbookMacroEnabled = 52

function Test-VbaProjectAccess {
    $key = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
    $value = (Get-ItemProperty -Path $key -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
    return ($value -eq 1)
}

# O editor VBA só importa .bas em ANSI (Windows-1252); o repositório guarda em UTF-8.
function ConvertTo-AnsiCopy {
    param([Parameter(Mandatory)][string]$Path)
    $temp = Join-Path ([IO.Path]::GetTempPath()) ([IO.Path]::GetFileName($Path))
    $text = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    [IO.File]::WriteAllText($temp, $text, [Text.Encoding]::GetEncoding(1252))
    return $temp
}

function Add-Sheets {
    param($Workbook)
    $first = $Workbook.Worksheets.Item(1)
    $first.Name = 'Parametros'
    foreach ($name in 'Base', 'Indicadores') {
        $sheet = $Workbook.Worksheets.Add([Type]::Missing, $Workbook.Worksheets.Item($Workbook.Worksheets.Count))
        $sheet.Name = $name
    }
}

function Set-Parameters {
    param($Workbook, [object]$Rule)
    $ws = $Workbook.Worksheets.Item('Parametros')
    $rows = @(
        @('prmAgora',         'Agora (data e hora de referência)', ([datetime]'2026-10-12 14:35').ToOADate(), 'dd/mm/yyyy hh:mm'),
        @('prmJanela',        'Janela do slot (min)',               $Rule.JanelaMinutos, '0'),
        @('prmTolInstalacao', 'Tolerância Instalação (min)',        $Rule.ToleranciaMinutos['Instalação'], '0'),
        @('prmTolReparo',     'Tolerância Reparo (min)',            $Rule.ToleranciaMinutos['Reparo'], '0'),
        @('prmMeta',          'Meta (%)',                           $Rule.Meta, '0.00'),
        @('prmMargem',        'Margem "Perto" (p.p.)',              $Rule.MargemPerto, '0.00')
    )
    $ws.Range('A1').Value2 = 'Parâmetro'
    $ws.Range('B1').Value2 = 'Valor'
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $r = $i + 2
        $ws.Range("A$r").Value2 = $rows[$i][1]
        Set-CellValue2 -Range $ws.Range("B$r") -Value $rows[$i][2]
        $ws.Range("B$r").NumberFormat = $rows[$i][3]
        [void]$Workbook.Names.Add($rows[$i][0], "=Parametros!`$B`$$r")
    }
    $ws.Range('A1:B1').Font.Bold = $true
    [void]$ws.Columns.Item('A:B').AutoFit()
}

function Add-OrdersQuery {
    param($Workbook, [string]$Csv)
    $formula = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'power-query\tbOrdens.pq'), [Text.Encoding]::UTF8)
    $formula = Set-MTextParameter -Formula $formula -Name 'CaminhoCsv' -Value $Csv
    [void]$Workbook.Queries.Add('tbOrdens', $formula, 'Ordens de serviço tratadas')

    # Carrega a consulta como tabela na aba Base (mesma conexão que o Excel cria pela interface).
    $ws = $Workbook.Worksheets.Item('Base')
    $connection = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=tbOrdens;Extended Properties=""'
    $table = $ws.ListObjects.Add($xlSrcExternal, $connection, [Type]::Missing, $xlYes, $ws.Range('A1'))
    $table.Name = 'tbOrdens'
    $table.QueryTable.CommandType = $xlCmdSql
    $table.QueryTable.CommandText = 'SELECT * FROM [tbOrdens]'
    $table.QueryTable.BackgroundQuery = $false
    [void]$table.QueryTable.Refresh($false)
    return $table.ListRows.Count
}

# ---------------------------------------------------------------------------

if (-not (Test-VbaProjectAccess)) {
    throw ('O Excel não permite importar módulos VBA por script. Ative em: Arquivo > Opções > ' +
           'Central de Confiabilidade > Configurações > Configurações de Macro > ' +
           '"Confiar no acesso ao modelo de objeto do projeto do VBA". Ou importe vba\mdlSlot.bas à mão (Alt+F11).')
}
if (-not (Test-Path $CsvPath)) { throw "CSV não encontrado: $CsvPath. Rode data-sample\New-FakeServiceOrders.ps1." }

$outDir = Split-Path -Parent $OutputPath
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
if (Test-Path $OutputPath) { Remove-Item $OutputPath }

$excel = $null; $workbook = $null
try {
    $excel = Invoke-Step 'Abrir Excel' { New-ExcelApplication }
    $workbook = Invoke-Step 'Criar pasta' {
        $wb = $excel.Workbooks.Add()
        Add-Sheets $wb
        $wb
    }
    Invoke-Step 'Parâmetros' { Set-Parameters -Workbook $workbook -Rule (Get-DefaultIndicatorRule) }
    $rows = Invoke-Step 'Power Query' { Add-OrdersQuery -Workbook $workbook -Csv (Resolve-Path $CsvPath).Path }
    Write-StepLog -Step 'Power Query' -Message "$rows ordens carregadas na aba Base"

    Invoke-Step 'Importar VBA' {
        $ansi = ConvertTo-AnsiCopy (Join-Path $PSScriptRoot 'vba\mdlSlot.bas')
        try { [void]$workbook.VBProject.VBComponents.Import($ansi) } finally { Remove-Item $ansi -ErrorAction SilentlyContinue }
    }

    $status = Invoke-Step 'Testar macro' { Invoke-ExcelMacro -Excel $excel -Workbook $workbook -Macro 'mdlSlot.CalcularIndicadores' }
    if ($status -notlike 'OK|*') { throw "Macro retornou: $status" }
    Write-StepLog -Step 'Testar macro' -Message $status

    Invoke-Step 'Salvar' {
        $workbook.Worksheets.Item('Indicadores').Activate()
        $workbook.SaveAs($OutputPath, $xlOpenXMLWorkbookMacroEnabled)
    }
    Write-StepLog -Step 'Concluído' -Level OK -Message "Matriz criada: $OutputPath"
}
finally {
    Close-ExcelSafely -Excel $excel -Workbook $workbook
}
