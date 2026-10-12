# ---------------------------------------------------------------------------
# Datas e horas
#
# O mesmo relatório rodava em PCs com formatos regionais diferentes (d/M/yy num,
# dd/MM/yyyy em outro), e as exportações misturam ISO, texto e número serial do Excel.
# Regra: NUNCA usar [datetime]$texto nem Get-Date -Date $texto (dependem da cultura do
# Windows). Sempre ParseExact com a lista de formatos conhecidos e cultura invariante.
# ---------------------------------------------------------------------------

$script:DateFormats = [string[]]@(
    'dd/MM/yyyy', 'd/M/yyyy', 'dd/MM/yy', 'd/M/yy',
    'yyyy-MM-dd', 'yyyy/MM/dd', 'dd-MM-yyyy', 'd-M-yyyy', 'dd.MM.yyyy',
    'dd/MM/yyyy HH:mm', 'dd/MM/yyyy HH:mm:ss', 'd/M/yyyy H:mm', 'd/M/yyyy H:mm:ss',
    'yyyy-MM-dd HH:mm', 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-ddTHH:mm:ss'
)

$script:TimeFormats = [string[]]@('H:mm', 'HH:mm', 'H:mm:ss', 'HH:mm:ss', 'Hmm', 'HHmm')

# Faixa de seriais aceitos como data: 01/01/1955 a 31/12/2099. Fora disso é só um número.
$script:ExcelSerialMin = 20090
$script:ExcelSerialMax = 73050

<#
.SYNOPSIS
    Converte texto, número serial do Excel ou [datetime] em [datetime]. Retorna $null se não reconhecer.
.EXAMPLE
    ConvertTo-DateSafe '12/10/26'   # 12/10/2026
    ConvertTo-DateSafe 46307        # 12/10/2026 (serial do Excel)
#>
function ConvertTo-DateSafe {
    [CmdletBinding()]
    [OutputType([datetime])]
    param([Parameter(ValueFromPipeline)][AllowNull()][AllowEmptyString()][object]$Value)

    process {
        if ($null -eq $Value) { return $null }
        if ($Value -is [datetime]) { return $Value }

        # Value2 do Excel devolve datas como double (serial OLE Automation).
        if ($Value -is [double] -or $Value -is [int] -or $Value -is [long] -or $Value -is [decimal]) {
            return (ConvertFrom-ExcelSerial ([double]$Value))
        }

        $text = ([string]$Value).Trim()
        if ($text -eq '') { return $null }

        $inv = [Globalization.CultureInfo]::InvariantCulture
        $serial = 0.0
        if ($text -match '^\d+([.,]\d+)?$' -and
            [double]::TryParse($text.Replace(',', '.'), [Globalization.NumberStyles]::Float, $inv, [ref]$serial)) {
            return (ConvertFrom-ExcelSerial $serial)
        }

        $parsed = [datetime]::MinValue
        if ([datetime]::TryParseExact($text, $script:DateFormats, $inv, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            return $parsed
        }
        return $null
    }
}

function ConvertFrom-ExcelSerial {
    param([double]$Serial)
    if ($Serial -lt $script:ExcelSerialMin -or $Serial -gt $script:ExcelSerialMax) { return $null }
    return [datetime]::FromOADate($Serial)
}

<#
.SYNOPSIS
    Converte "8:05", "08:05:00", "0805", fração de dia do Excel (0,3368) ou [datetime] em [timespan].
    Retorna $null se não reconhecer.
#>
function ConvertTo-TimeOfDay {
    [CmdletBinding()]
    [OutputType([timespan])]
    param([Parameter(ValueFromPipeline)][AllowNull()][AllowEmptyString()][object]$Value)

    process {
        if ($null -eq $Value) { return $null }
        if ($Value -is [timespan]) { return $Value }
        if ($Value -is [datetime]) { return $Value.TimeOfDay }
        if ($Value -is [double] -or $Value -is [decimal]) {
            $fraction = [double]$Value
            if ($fraction -ge 0 -and $fraction -lt 1) { return [timespan]::FromMinutes([math]::Round($fraction * 1440)) }
            return $null
        }

        $text = ([string]$Value).Trim()
        if ($text -eq '') { return $null }

        $inv = [Globalization.CultureInfo]::InvariantCulture
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParseExact($text, $script:TimeFormats, $inv, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            return $parsed.TimeOfDay
        }
        return $null
    }
}
