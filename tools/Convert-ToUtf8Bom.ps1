<#
.SYNOPSIS
    Regrava scripts PowerShell em UTF-8 com BOM.

.DESCRIPTION
    O Windows PowerShell 5.1 interpreta arquivos sem BOM como ANSI (Windows-1252),
    então "Concluída" vira "ConcluÃ­da". Este script procura .ps1/.psm1/.psd1 sem BOM
    e os regrava com BOM, sem alterar o conteúdo.

.PARAMETER Path
    Pasta raiz da varredura (padrão: raiz do repositório).

.PARAMETER WhatIf
    Só lista os arquivos que seriam convertidos.

.EXAMPLE
    .\tools\Convert-ToUtf8Bom.ps1
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# Padrão calculado aqui, e não no param: no PowerShell 5.1 com -File, $PSScriptRoot chega vazio no bloco param.
if (-not $Path) { $Path = Split-Path -Parent $PSScriptRoot }

# Retorna $true se o arquivo começa com os bytes EF BB BF.
function Test-Utf8Bom {
    param([Parameter(Mandatory)][string]$FilePath)
    $bytes = [IO.File]::ReadAllBytes($FilePath)
    return ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
}

$utf8Bom    = New-Object System.Text.UTF8Encoding($true)
$utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)   # lança erro se o arquivo não for UTF-8 válido

$files = Get-ChildItem -Path $Path -Recurse -File -Include *.ps1, *.psm1, *.psd1 |
    Where-Object { $_.FullName -notmatch '\\\.git\\' }

$converted = 0
foreach ($file in $files) {
    if (Test-Utf8Bom -FilePath $file.FullName) { continue }

    try {
        $text = $utf8Strict.GetString([IO.File]::ReadAllBytes($file.FullName))
    }
    catch {
        Write-Warning "Ignorado (não é UTF-8 válido, revise à mão): $($file.FullName)"
        continue
    }

    if ($PSCmdlet.ShouldProcess($file.FullName, 'Regravar em UTF-8 com BOM')) {
        [IO.File]::WriteAllText($file.FullName, $text, $utf8Bom)
        $converted++
        Write-Host "BOM adicionado: $($file.FullName.Substring($Path.Length).TrimStart('\'))"
    }
}

Write-Host "Concluído: $converted arquivo(s) convertido(s) de $($files.Count) verificados."
