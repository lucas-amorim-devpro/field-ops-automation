# ---------------------------------------------------------------------------
# Tratamento das exportações
#
# As bases chegam dos portais com espaços sobrando, caixa variada, acentuação quebrada
# (UTF-8 aberto como ANSI) e linhas repetidas. Tudo é normalizado aqui, antes de qualquer
# cálculo, para que "  Vale do Cedro", "VALE DO CEDRO" e "Vale do Cedro" sejam a mesma cidade.
# ---------------------------------------------------------------------------

$script:Utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
$script:Ansi1252   = [Text.Encoding]::GetEncoding(1252)

<#
.SYNOPSIS
    Desfaz acentuação quebrada do tipo "SÃ£o" -> "São" e "ConcluÃ­da" -> "Concluída".
.DESCRIPTION
    Só tenta reparar se o texto tiver os marcadores típicos (Ã, Â) e se os bytes
    resultantes formarem UTF-8 válido. Texto correto passa intacto.
#>
function Repair-Mojibake {
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(ValueFromPipeline)][AllowNull()][AllowEmptyString()][string]$Text)

    process {
        if ([string]::IsNullOrEmpty($Text) -or $Text -cnotmatch '[ÃÂ]') { return $Text }
        try {
            return $script:Utf8Strict.GetString($script:Ansi1252.GetBytes($Text))
        }
        catch {
            return $Text   # não era mojibake de verdade (ex. nome com "Ã" legítimo)
        }
    }
}

<#
.SYNOPSIS
    Remove espaços nas pontas, colapsa espaços duplos e repara a acentuação.
#>
function Format-CleanText {
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(ValueFromPipeline)][AllowNull()][AllowEmptyString()][string]$Text)

    process {
        if ($null -eq $Text) { return '' }
        $clean = Repair-Mojibake $Text
        return ($clean -replace '\s+', ' ').Trim()
    }
}

<#
.SYNOPSIS
    Remove acentos e passa para minúsculas. Usado só para COMPARAR textos.
#>
function Get-ComparableText {
    param([AllowNull()][AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $decomposed = (Format-CleanText $Text).Normalize([Text.NormalizationForm]::FormD)
    $chars = $decomposed.ToCharArray() | Where-Object {
        [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark
    }
    return (-join $chars).ToLowerInvariant()
}

<#
.SYNOPSIS
    Coloca nomes próprios em "Título": "  MARINA  COUTINHO " -> "Marina Coutinho".
    Preposições (da, de, do, dos, das, e) ficam minúsculas.
#>
function Format-ProperName {
    param([AllowNull()][AllowEmptyString()][string]$Text)
    $clean = Format-CleanText $Text
    if ($clean -eq '') { return '' }
    $culture = [Globalization.CultureInfo]::GetCultureInfo('pt-BR')
    $words = $culture.TextInfo.ToTitleCase($clean.ToLower($culture)) -split ' '
    $small = @('Da', 'De', 'Do', 'Das', 'Dos', 'E')
    for ($i = 1; $i -lt $words.Count; $i++) {
        if ($small -ccontains $words[$i]) { $words[$i] = $words[$i].ToLower($culture) }
    }
    return ($words -join ' ')
}

<#
.SYNOPSIS
    Padroniza o tipo da ordem para 'Instalação' ou 'Reparo'. Retorna $null se desconhecido.
#>
function ConvertTo-OrderType {
    param([AllowNull()][AllowEmptyString()][string]$Text)
    $key = Get-ComparableText $Text
    if ($key -like 'instal*') { return 'Instalação' }
    if ($key -like 'repar*')  { return 'Reparo' }
    return $null
}

<#
.SYNOPSIS
    Padroniza o status da ordem para um dos valores conhecidos. Retorna $null se desconhecido.
#>
function ConvertTo-OrderStatus {
    param([AllowNull()][AllowEmptyString()][string]$Text)
    $map = @{
        'concluida'       = 'Concluída'
        'em execucao'     = 'Em execução'
        'em deslocamento' = 'Em deslocamento'
        'agendada'        = 'Agendada'
        'nao realizada'   = 'Não realizada'
        'cancelada'       = 'Cancelada'
    }
    $key = Get-ComparableText $Text
    if ($map.ContainsKey($key)) { return $map[$key] }
    return $null
}

<#
.SYNOPSIS
    Normaliza o slot para "HH:mm" ("8:30" -> "08:30"). Retorna $null se inválido.
#>
function ConvertTo-SlotText {
    param([AllowNull()][object]$Value)
    $time = ConvertTo-TimeOfDay $Value
    if ($null -eq $time) { return $null }
    return ('{0:00}:{1:00}' -f $time.Hours, $time.Minutes)
}

<#
.SYNOPSIS
    Converte uma linha bruta do CSV em uma ordem tipada e limpa.
.OUTPUTS
    [pscustomobject] com Data [datetime], Chegada/Inicio/Termino [datetime] ou $null,
    e a propriedade Problemas (lista de textos) quando algum campo não pôde ser lido.
#>
function ConvertTo-CleanOrder {
    [CmdletBinding()]
    param([Parameter(Mandatory, ValueFromPipeline)][object]$Row)

    process {
        $problems = New-Object System.Collections.Generic.List[string]

        $date = ConvertTo-DateSafe $Row.data
        if ($null -eq $date) { $problems.Add("data inválida '$($Row.data)'") }

        $type = ConvertTo-OrderType $Row.tipo
        if ($null -eq $type) { $problems.Add("tipo desconhecido '$($Row.tipo)'") }

        $status = ConvertTo-OrderStatus $Row.status
        if ($null -eq $status) { $problems.Add("status desconhecido '$($Row.status)'") }

        $slot = ConvertTo-SlotText $Row.slot
        if ($null -eq $slot) { $problems.Add("slot inválido '$($Row.slot)'") }

        # Hora do dia + data da ordem = momento completo (ou $null se vazio).
        $toMoment = {
            param($raw)
            $t = ConvertTo-TimeOfDay $raw
            if ($null -eq $t -or $null -eq $date) { return $null }
            return $date.Date.Add($t)
        }

        [pscustomobject]@{
            IdOs        = (Format-CleanText $Row.id_os).ToUpperInvariant()
            Data        = if ($date) { $date.Date } else { $null }
            Regional    = (Format-CleanText $Row.regional).ToUpperInvariant()
            Cidade      = Format-ProperName $Row.cidade
            Gerente     = Format-ProperName $Row.gerente
            Coordenacao = Format-ProperName $Row.coordenacao
            Supervisao  = Format-ProperName $Row.supervisao
            Tecnico     = Format-ProperName $Row.tecnico
            Tipo        = $type
            Slot        = $slot
            Chegada     = & $toMoment $Row.chegada
            Inicio      = & $toMoment $Row.inicio
            Termino     = & $toMoment $Row.termino
            Status      = $status
            Problemas   = $problems.ToArray()
        }
    }
}

<#
.SYNOPSIS
    Remove ordens repetidas pelo IdOs, mantendo a primeira ocorrência.
#>
function Remove-DuplicateOrder {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Orders)
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($order in $Orders) {
        if ($seen.Add([string]$order.IdOs)) { $order }
    }
}

<#
.SYNOPSIS
    Lê o CSV exportado, limpa, tipa e remove duplicados.
.OUTPUTS
    [pscustomobject] com Orders (válidas), Rejected (com problemas) e Stats.
#>
function Import-ServiceOrder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Delimiter = ';'
    )
    if (-not (Test-Path $Path)) { throw "Arquivo não encontrado: $Path" }

    $raw = @(Import-Csv -Path $Path -Delimiter $Delimiter -Encoding UTF8)
    $clean = @($raw | ConvertTo-CleanOrder)
    $valid = @($clean | Where-Object { $_.Problemas.Count -eq 0 })
    $rejected = @($clean | Where-Object { $_.Problemas.Count -gt 0 })
    $unique = @(Remove-DuplicateOrder -Orders $valid)

    [pscustomobject]@{
        Orders   = $unique
        Rejected = $rejected
        Stats    = [pscustomobject]@{
            LinhasLidas = $raw.Count
            Rejeitadas  = $rejected.Count
            Duplicadas  = $valid.Count - $unique.Count
            Validas     = $unique.Count
        }
    }
}
