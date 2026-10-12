# ---------------------------------------------------------------------------
# Excel via COM
#
# Aprendizados que viraram funções:
#   - Atribuir $range.Value2 = 123 depois de ter gravado texto na mesma variável COM pode
#     falhar com InvalidCastException: o binder do PowerShell "lembra" o tipo anterior.
#     InvokeMember('Value2', SetProperty) passa pelo IDispatch direto e não sofre disso.
#   - Macros chamadas de fora com Application.Run precisam do nome qualificado
#     ('Pasta.xlsm'!Modulo.Macro), senão o Excel procura em outra pasta aberta.
#   - Sem ReleaseComObject + GC, o processo EXCEL.EXE fica preso em segundo plano
#     e a próxima execução encontra o arquivo "em uso".
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Grava um valor em Range.Value2 via IDispatch (InvokeMember), evitando InvalidCast.
.EXAMPLE
    Set-CellValue2 -Range $ws.Range('B2') -Value 87.5
#>
function Set-CellValue2 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Range,
        [AllowNull()][object]$Value
    )
    # Parâmetros de função chegam embrulhados em PSObject; o IDispatch do Excel não entende
    # esse tipo e responde 0x800A03EC. Por isso o valor é desembrulhado antes da chamada.
    $arguments = New-Object 'object[]' 1
    $arguments[0] = if ($Value -is [psobject]) { $Value.psobject.BaseObject } else { $Value }

    [void]$Range.GetType().InvokeMember(
        'Value2', [Reflection.BindingFlags]::SetProperty, $null, $Range, $arguments)
}

<#
.SYNOPSIS
    Lê Range.Value2 via IDispatch. Para várias células, devolve object[,] (base 1).
#>
function Get-CellValue2 {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Range)
    $value = $Range.GetType().InvokeMember('Value2', [Reflection.BindingFlags]::GetProperty, $null, $Range, $null)
    # A vírgula impede o PowerShell de "desenrolar" o object[,] em itens soltos no pipeline.
    return , $value
}

<#
.SYNOPSIS
    Converte o object[,] de um Value2 com cabeçalho em uma lista de objetos.
.DESCRIPTION
    Ler a faixa inteira de uma vez (uma chamada COM) é muito mais rápido do que
    ler célula a célula (uma chamada COM por célula).
#>
function ConvertFrom-Value2Table {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Values)

    if ($Values -isnot [array] -or $Values.Rank -ne 2) { throw 'Esperado object[,] vindo de Range.Value2 com várias células.' }
    $rows = $Values.GetLength(0)
    $cols = $Values.GetLength(1)
    $headers = for ($c = 1; $c -le $cols; $c++) { [string]$Values[1, $c] }

    for ($r = 2; $r -le $rows; $r++) {
        $item = [ordered]@{}
        for ($c = 1; $c -le $cols; $c++) { $item[$headers[$c - 1]] = $Values[$r, $c] }
        [pscustomobject]$item
    }
}

<#
.SYNOPSIS
    Abre uma instância isolada do Excel, invisível e sem alertas.
#>
function New-ExcelApplication {
    [CmdletBinding()]
    param([switch]$Visible)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$Visible
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $true
    return $excel
}

<#
.SYNOPSIS
    Atualiza todas as consultas Power Query e espera terminar.
.DESCRIPTION
    RefreshAll volta antes de as consultas acabarem quando há conexões em segundo plano.
    CalculateUntilAsyncQueriesDone bloqueia até tudo estar pronto.
#>
function Update-ExcelQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Excel,
        [Parameter(Mandatory)][object]$Workbook
    )
    foreach ($connection in $Workbook.Connections) {
        try {
            if ($connection.Type -eq 1) { $connection.OLEDBConnection.BackgroundQuery = $false }   # 1 = xlConnectionTypeOLEDB
        }
        catch { Write-Verbose "Conexão sem BackgroundQuery: $($connection.Name)" }
    }
    $Workbook.RefreshAll()
    $Excel.CalculateUntilAsyncQueriesDone()
    $Excel.CalculateFull()
}

<#
.SYNOPSIS
    Executa uma macro VBA da pasta indicada e devolve o retorno (se for Function).
.EXAMPLE
    Invoke-ExcelMacro -Excel $excel -Workbook $wb -Macro 'mdlSlot.CalcularIndicadores'
#>
function Invoke-ExcelMacro {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Excel,
        [Parameter(Mandatory)][object]$Workbook,
        [Parameter(Mandatory)][string]$Macro,
        [object[]]$Arguments = @()
    )
    $qualified = "'{0}'!{1}" -f $Workbook.Name, $Macro
    # Desembrulha PSObject pelo mesmo motivo de Set-CellValue2.
    $callArgs = @($qualified) + @($Arguments | ForEach-Object {
        if ($_ -is [psobject]) { $_.psobject.BaseObject } else { $_ }
    })
    return $Excel.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $Excel, [object[]]$callArgs)
}

<#
.SYNOPSIS
    PID do processo EXCEL.EXE por trás de um objeto Excel.Application (via janela principal).
#>
function Get-ExcelProcessId {
    param([Parameter(Mandatory)][object]$Excel)
    if (-not ('FieldOps.Win32Window' -as [type])) {
        Add-Type -Namespace FieldOps -Name Win32Window -MemberDefinition @'
[DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
'@
    }
    $processId = [uint32]0
    [void][FieldOps.Win32Window]::GetWindowThreadProcessId([IntPtr]$Excel.Hwnd, [ref]$processId)
    return [int]$processId
}

<#
.SYNOPSIS
    Fecha a pasta e o Excel, libera os objetos COM e garante que o EXCEL.EXE terminou.
.DESCRIPTION
    Qualquer objeto COM ainda referenciado (uma planilha, um Range) mantém o EXCEL.EXE vivo
    depois do Quit: o "Excel fantasma" que segura o arquivo e trava a próxima execução.
    Estratégia: anotar o PID desta instância, liberar e coletar o lixo duas vezes (a 1ª
    coleta solta os wrappers, a 2ª finaliza o que eles referenciavam), esperar o processo
    sair e, só se ele não sair, encerrar ESSE PID. Instâncias que o usuário abriu não são tocadas.
#>
function Close-ExcelSafely {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$Excel,
        [AllowNull()][object]$Workbook,
        [switch]$Save,
        [int]$WaitSeconds = 10
    )
    $processId = 0
    if ($Excel) {
        try { $processId = Get-ExcelProcessId -Excel $Excel } catch { Write-Verbose "Sem PID: $($_.Exception.Message)" }
    }

    if ($Workbook) {
        try { $Workbook.Close([bool]$Save) } catch { Write-Verbose "Close: $($_.Exception.Message)" }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Workbook)
    }
    if ($Excel) {
        try { $Excel.Quit() } catch { Write-Verbose "Quit: $($_.Exception.Message)" }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Excel)
    }
    for ($i = 0; $i -lt 2; $i++) {
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    if ($processId -gt 0) {
        $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
        if ($process -and -not $process.WaitForExit($WaitSeconds * 1000)) {
            Write-Warning "EXCEL.EXE (PID $processId) não fechou em $WaitSeconds s; encerrando esta instância."
            Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
        }
    }
}
