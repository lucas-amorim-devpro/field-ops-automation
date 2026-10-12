# Módulo com as funções comuns dos relatórios.
# Uso: Import-Module "$PSScriptRoot\lib\FieldOps.psm1" -Force

$ErrorActionPreference = 'Stop'

foreach ($file in 'Logging', 'Console', 'Dates', 'Cleaning', 'Indicators', 'Excel', 'Json', 'AppsScript') {
    . (Join-Path $PSScriptRoot "$file.ps1")
}

Export-ModuleMember -Function @(
    # Log por etapa
    'Start-StepLog', 'Write-StepLog', 'Invoke-Step', 'Get-ErrorDetail',
    # Console
    'Disable-QuickEdit',
    # Datas
    'ConvertTo-DateSafe', 'ConvertTo-TimeOfDay',
    # Tratamento
    'Repair-Mojibake', 'Format-CleanText', 'Get-ComparableText', 'Format-ProperName',
    'ConvertTo-OrderType', 'ConvertTo-OrderStatus', 'ConvertTo-SlotText',
    'ConvertTo-CleanOrder', 'Remove-DuplicateOrder', 'Import-ServiceOrder',
    # Indicador
    'Get-DefaultIndicatorRule', 'Get-RoundedHalfUp', 'Get-SlotDeadline', 'Get-PunctualityStatus',
    'Get-GoalBadge', 'Get-PunctualitySummary', 'Get-VariationPp',
    # Excel COM
    'Set-CellValue2', 'Get-CellValue2', 'ConvertFrom-Value2Table', 'New-ExcelApplication',
    'Update-ExcelQuery', 'Invoke-ExcelMacro', 'Close-ExcelSafely',
    # JSON
    'ConvertTo-CompactJson',
    # Apps Script e Power Query
    'Set-MTextParameter', 'Send-AppsScriptPayload'
)
