# ---------------------------------------------------------------------------
# Indicador "Chegada ao Slot"
#
# Mede se o técnico chegou ao cliente dentro da janela agendada.
#   - Janela: começa no horário do slot e dura JanelaMinutos (padrão 120).
#   - Tolerância por tipo: Reparo ganha minutos extras (regra de negócio diferente).
#   - Chegou até o limite            -> NoPrazo
#   - Chegou depois do limite         -> ForaDoPrazo
#   - Não chegou e o limite já passou -> ForaDoPrazo (estourou)
#   - Não chegou e ainda há tempo     -> Pendente (não entra no percentual)
#   - Cancelada                       -> Excluida
#
# % no prazo = NoPrazo / (NoPrazo + ForaDoPrazo), com 2 casas e arredondamento
# "meio para cima" (half away from zero), igual à função ARRED/ROUND do Google Sheets,
# para que o painel bata com a planilha que a equipe usa para conferir.
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Regras padrão do indicador. O pipeline sobrescreve com o que estiver no config.json.
#>
function Get-DefaultIndicatorRule {
    return [pscustomobject]@{
        JanelaMinutos     = 120
        ToleranciaMinutos = @{ 'Instalação' = 0; 'Reparo' = 30 }
        Meta              = 80
        MargemPerto       = 5
    }
}

<#
.SYNOPSIS
    Arredonda com "meio para cima" (0,125 -> 0,13; -0,125 -> -0,13), como o Google Sheets.
.DESCRIPTION
    [math]::Round usa "arredondamento bancário" por padrão (0,125 -> 0,12) e, em double,
    80,145 vira 80,14499999... Converter para [decimal] (15 dígitos significativos, a mesma
    precisão exibida pelo Sheets) e usar AwayFromZero resolve os dois problemas.
#>
function Get-RoundedHalfUp {
    [CmdletBinding()]
    [OutputType([decimal])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)][object]$Value,
        [ValidateRange(0, 10)][int]$Digits = 2
    )
    process {
        return [math]::Round([decimal]$Value, $Digits, [MidpointRounding]::AwayFromZero)
    }
}

<#
.SYNOPSIS
    Momento limite de chegada da ordem: data + slot + janela + tolerância do tipo.
#>
function Get-SlotDeadline {
    [CmdletBinding()]
    [OutputType([datetime])]
    param(
        [Parameter(Mandatory)][object]$Order,
        [object]$Rule = (Get-DefaultIndicatorRule)
    )
    $tolerance = 0
    if ($Rule.ToleranciaMinutos.ContainsKey($Order.Tipo)) { $tolerance = $Rule.ToleranciaMinutos[$Order.Tipo] }
    $slotStart = $Order.Data.Date.Add([timespan]::Parse($Order.Slot))
    return $slotStart.AddMinutes($Rule.JanelaMinutos + $tolerance)
}

<#
.SYNOPSIS
    Classifica uma ordem em NoPrazo, ForaDoPrazo, Pendente ou Excluida.
#>
function Get-PunctualityStatus {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][object]$Order,
        [Parameter(Mandatory)][datetime]$Now,
        [object]$Rule = (Get-DefaultIndicatorRule)
    )
    if ($Order.Status -eq 'Cancelada') { return 'Excluida' }

    $deadline = Get-SlotDeadline -Order $Order -Rule $Rule
    if ($null -ne $Order.Chegada) {
        if ($Order.Chegada -le $deadline) { return 'NoPrazo' }
        return 'ForaDoPrazo'
    }
    if ($Now -gt $deadline) { return 'ForaDoPrazo' }
    return 'Pendente'
}

<#
.SYNOPSIS
    Selo da meta: 'Na meta' (>= meta), 'Perto' (até MargemPerto p.p. abaixo) ou 'Abaixo'.
#>
function Get-GoalBadge {
    [OutputType([string])]
    param(
        [AllowNull()][object]$Percent,
        [object]$Rule = (Get-DefaultIndicatorRule)
    )
    if ($null -eq $Percent) { return 'Sem dados' }
    if ($Percent -ge $Rule.Meta) { return 'Na meta' }
    if ($Percent -ge ($Rule.Meta - $Rule.MargemPerto)) { return 'Perto' }
    return 'Abaixo'
}

<#
.SYNOPSIS
    Calcula o % no prazo de um conjunto de ordens, opcionalmente agrupado por uma propriedade.
.EXAMPLE
    Get-PunctualitySummary -Orders $hoje -Now (Get-Date) -GroupBy Regional
#>
function Get-PunctualitySummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Orders,
        [Parameter(Mandatory)][datetime]$Now,
        [string]$GroupBy,
        [object]$Rule = (Get-DefaultIndicatorRule)
    )

    $groups = if ($GroupBy) {
        $Orders | Group-Object -Property $GroupBy
    } else {
        , [pscustomobject]@{ Name = 'Total'; Group = $Orders }
    }

    foreach ($group in $groups) {
        $counts = @{ NoPrazo = 0; ForaDoPrazo = 0; Pendente = 0; Excluida = 0 }
        foreach ($order in $group.Group) {
            $counts[(Get-PunctualityStatus -Order $order -Now $Now -Rule $Rule)]++
        }

        $measured = $counts.NoPrazo + $counts.ForaDoPrazo
        $percent = $null
        if ($measured -gt 0) {
            # Divisão em decimal para não carregar erro de ponto flutuante até o arredondamento.
            $percent = Get-RoundedHalfUp ([decimal]$counts.NoPrazo / [decimal]$measured * 100) 2
        }

        [pscustomobject]@{
            Grupo      = $group.Name
            Apuradas   = $measured
            NoPrazo    = $counts.NoPrazo
            Fora       = $counts.ForaDoPrazo
            Pendentes  = $counts.Pendente
            Percentual = $percent
            Selo       = Get-GoalBadge -Percent $percent -Rule $Rule
        }
    }
}

<#
.SYNOPSIS
    Variação em pontos percentuais entre dois percentuais, arredondada como o Sheets.
#>
function Get-VariationPp {
    param([AllowNull()][object]$Current, [AllowNull()][object]$Previous)
    if ($null -eq $Current -or $null -eq $Previous) { return $null }
    return Get-RoundedHalfUp ([decimal]$Current - [decimal]$Previous) 2
}
