# ---------------------------------------------------------------------------
# JSON compacto
#
# O ConvertTo-Json do PowerShell 5.1 é lento em milhares de linhas, tem limite de
# profundidade e "desenrola" listas de um item só (["a"] vira "a"). Para o formato
# dicionário + linhas do painel, um serializador pequeno e previsível é mais seguro.
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Serializa hashtable/ordered, listas, textos, números, booleanos, datas e $null em JSON sem espaços.
.EXAMPLE
    ConvertTo-CompactJson ([ordered]@{ a = 1; b = @('x', $null) })   # {"a":1,"b":["x",null]}
#>
function ConvertTo-CompactJson {
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][object]$InputObject)

    $sb = New-Object System.Text.StringBuilder
    Write-JsonValue -Builder $sb -Value $InputObject
    return $sb.ToString()
}

function Write-JsonValue {
    param(
        [Parameter(Mandatory)][System.Text.StringBuilder]$Builder,
        [AllowNull()][object]$Value
    )
    $inv = [Globalization.CultureInfo]::InvariantCulture
    if ($Value -is [psobject] -and $Value -isnot [Management.Automation.PSCustomObject]) { $Value = $Value.psobject.BaseObject }

    if ($null -eq $Value)        { [void]$Builder.Append('null'); return }
    if ($Value -is [bool])       { [void]$Builder.Append($(if ($Value) { 'true' } else { 'false' })); return }
    if ($Value -is [string] -or $Value -is [char]) { Write-JsonString -Builder $Builder -Text ([string]$Value); return }
    if ($Value -is [datetime])   { Write-JsonString -Builder $Builder -Text $Value.ToString('yyyy-MM-ddTHH:mm:ss', $inv); return }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal] -or $Value -is [single]) {
        [void]$Builder.Append(([IFormattable]$Value).ToString($null, $inv)); return
    }

    if ($Value -is [Collections.IDictionary]) {
        [void]$Builder.Append('{')
        $first = $true
        foreach ($key in $Value.Keys) {
            if (-not $first) { [void]$Builder.Append(',') }
            Write-JsonString -Builder $Builder -Text ([string]$key)
            [void]$Builder.Append(':')
            Write-JsonValue -Builder $Builder -Value $Value[$key]
            $first = $false
        }
        [void]$Builder.Append('}')
        return
    }

    if ($Value -is [Management.Automation.PSCustomObject]) {
        $dict = [ordered]@{}
        foreach ($p in $Value.psobject.Properties) { $dict[$p.Name] = $p.Value }
        Write-JsonValue -Builder $Builder -Value $dict
        return
    }

    if ($Value -is [Collections.IEnumerable]) {
        [void]$Builder.Append('[')
        $first = $true
        foreach ($item in $Value) {
            if (-not $first) { [void]$Builder.Append(',') }
            Write-JsonValue -Builder $Builder -Value $item
            $first = $false
        }
        [void]$Builder.Append(']')
        return
    }

    Write-JsonString -Builder $Builder -Text ([string]$Value)
}

function Write-JsonString {
    param([System.Text.StringBuilder]$Builder, [string]$Text)
    [void]$Builder.Append('"')
    foreach ($ch in $Text.ToCharArray()) {
        switch ($ch) {
            '"'  { [void]$Builder.Append('\"'); continue }
            '\'  { [void]$Builder.Append('\\'); continue }
            "`n" { [void]$Builder.Append('\n'); continue }
            "`r" { [void]$Builder.Append('\r'); continue }
            "`t" { [void]$Builder.Append('\t'); continue }
            default {
                if ([int]$ch -lt 0x20) { [void]$Builder.AppendFormat('\u{0:x4}', [int]$ch) }
                else { [void]$Builder.Append($ch) }
            }
        }
    }
    [void]$Builder.Append('"')
}
