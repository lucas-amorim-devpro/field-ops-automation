# ---------------------------------------------------------------------------
# Console
#
# Com o "Modo de Edição Rápida" (QuickEdit) ligado, um clique acidental na janela do
# console seleciona texto e PAUSA o script até alguém apertar Enter. Num relatório de
# 10 minutos, isso parecia travamento do Excel. Os scripts longos desligam o QuickEdit
# no início, só para o próprio processo.
# ---------------------------------------------------------------------------

$script:ConsoleApiLoaded = $false

function Initialize-ConsoleApi {
    if ($script:ConsoleApiLoaded) { return }
    if (-not ('FieldOps.ConsoleMode' -as [type])) {
        Add-Type -Namespace FieldOps -Name ConsoleMode -MemberDefinition @'
[DllImport("kernel32.dll", SetLastError = true)]
public static extern IntPtr GetStdHandle(int nStdHandle);
[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);
'@
    }
    $script:ConsoleApiLoaded = $true
}

<#
.SYNOPSIS
    Desliga o QuickEdit do console atual. Retorna $true se alterou, $false se não havia console.
.DESCRIPTION
    Não lança erro quando não existe console interativo (ISE, VS Code, tarefa agendada,
    saída redirecionada): nesses casos não há o que desligar.
#>
function Disable-QuickEdit {
    [CmdletBinding()]
    param()

    $STD_INPUT_HANDLE      = -10
    $ENABLE_QUICK_EDIT     = 0x0040
    $ENABLE_EXTENDED_FLAGS = 0x0080   # exigido para o Windows respeitar a mudança

    try {
        Initialize-ConsoleApi
        $handle = [FieldOps.ConsoleMode]::GetStdHandle($STD_INPUT_HANDLE)
        $mode = [uint32]0
        if (-not [FieldOps.ConsoleMode]::GetConsoleMode($handle, [ref]$mode)) {
            Write-Verbose 'Sem console interativo; QuickEdit não se aplica.'
            return $false
        }
        $newMode = ($mode -band (-bnot $ENABLE_QUICK_EDIT)) -bor $ENABLE_EXTENDED_FLAGS
        return [FieldOps.ConsoleMode]::SetConsoleMode($handle, [uint32]$newMode)
    }
    catch {
        Write-Verbose "Não foi possível alterar o QuickEdit: $($_.Exception.Message)"
        return $false
    }
}
