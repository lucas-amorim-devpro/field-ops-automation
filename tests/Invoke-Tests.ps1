<#
.SYNOPSIS
    Roda todos os testes Pester do repositório.

.PARAMETER SkipExcel
    Pula os testes de integração que abrem o Excel.

.EXAMPLE
    .\tests\Invoke-Tests.ps1
.EXAMPLE
    .\tests\Invoke-Tests.ps1 -SkipExcel
#>
[CmdletBinding()]
param([switch]$SkipExcel)

$ErrorActionPreference = 'Stop'

$pester = Get-Module -ListAvailable Pester | Sort-Object Version -Descending | Select-Object -First 1
if (-not $pester -or $pester.Version.Major -lt 5) {
    throw 'É preciso o Pester 5 ou superior: Install-Module Pester -Scope CurrentUser -Force -SkipPublisherCheck'
}
Import-Module Pester -MinimumVersion 5.0

$config = New-PesterConfiguration
$config.Run.Path = $PSScriptRoot
$config.Run.Exit = $true
$config.Output.Verbosity = 'Detailed'
if ($SkipExcel) { $config.Filter.ExcludeTag = @('Excel') }

Invoke-Pester -Configuration $config
