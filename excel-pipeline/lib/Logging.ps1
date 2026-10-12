# ---------------------------------------------------------------------------
# Log por etapa
#
# Cada relatório é uma sequência de etapas ("Abrir matriz", "Atualizar consultas"...).
# Registrar início, duração e erro de cada uma mostra ONDE uma execução de madrugada
# falhou, sem precisar repetir tudo com o console aberto.
# ---------------------------------------------------------------------------

$script:LogFile = $null

<#
.SYNOPSIS
    Define o arquivo de log da execução (cria a pasta se preciso).
.EXAMPLE
    Start-StepLog -Directory .\logs -Name 'slot-pipeline'
    # -> .\logs\slot-pipeline_2026-10-12_131005.log
#>
function Start-StepLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Directory,
        [Parameter(Mandatory)][string]$Name
    )
    if (-not (Test-Path $Directory)) { New-Item -ItemType Directory -Path $Directory -Force | Out-Null }
    $script:LogFile = Join-Path $Directory ('{0}_{1:yyyy-MM-dd_HHmmss}.log' -f $Name, (Get-Date))
    Write-StepLog -Step 'Log' -Message "Arquivo de log: $script:LogFile"
    return $script:LogFile
}

<#
.SYNOPSIS
    Escreve uma linha "data hora | nível | etapa | mensagem" no console e no arquivo de log.
#>
function Write-StepLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Step,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [ValidateSet('INFO', 'OK', 'AVISO', 'ERRO')][string]$Level = 'INFO'
    )
    $line = '{0:yyyy-MM-dd HH:mm:ss} | {1,-5} | {2,-24} | {3}' -f (Get-Date), $Level, $Step, $Message

    $color = switch ($Level) { 'OK' { 'Green' } 'AVISO' { 'Yellow' } 'ERRO' { 'Red' } default { 'Gray' } }
    Write-Host $line -ForegroundColor $color

    if ($script:LogFile) {
        # UTF-8 explícito: Add-Content no 5.1 usaria ANSI e quebraria os acentos.
        [IO.File]::AppendAllText($script:LogFile, $line + [Environment]::NewLine, [Text.Encoding]::UTF8)
    }
}

<#
.SYNOPSIS
    Executa um bloco como uma etapa nomeada: registra início, duração e erro.
.DESCRIPTION
    Em caso de erro, registra a mensagem (com o HResult, se for erro COM) e relança,
    para o chamador decidir se aborta. Retorna o que o bloco retornar.
.EXAMPLE
    $wb = Invoke-Step 'Abrir matriz' { $excel.Workbooks.Open($path) }
#>
function Invoke-Step {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)][string]$Name,
        [Parameter(Mandatory, Position = 1)][scriptblock]$ScriptBlock
    )
    Write-StepLog -Step $Name -Message 'iniciando'
    $watch = [Diagnostics.Stopwatch]::StartNew()
    try {
        $result = & $ScriptBlock
        Write-StepLog -Step $Name -Level OK -Message ('concluída em {0:N1} s' -f $watch.Elapsed.TotalSeconds)
        return $result
    }
    catch {
        Write-StepLog -Step $Name -Level ERRO -Message ('falhou após {0:N1} s: {1}' -f $watch.Elapsed.TotalSeconds, (Get-ErrorDetail $_))
        throw
    }
}

<#
.SYNOPSIS
    Resume um erro, incluindo o HResult real de exceções COM.
.DESCRIPTION
    Erros do Excel via COM chegam embrulhados (TargetInvocationException,
    MethodInvocationException). O código útil (ex. 0x800A03EC) fica na exceção mais interna.
#>
function Get-ErrorDetail {
    param([Parameter(Mandatory)][object]$ErrorRecord)

    $ex = if ($ErrorRecord -is [Management.Automation.ErrorRecord]) { $ErrorRecord.Exception } else { $ErrorRecord }
    while ($ex.InnerException) { $ex = $ex.InnerException }

    $detail = $ex.Message.Trim()
    if ($ex -is [Runtime.InteropServices.COMException]) {
        $detail += ' (HResult 0x{0:X8})' -f $ex.HResult
    }
    return $detail
}
