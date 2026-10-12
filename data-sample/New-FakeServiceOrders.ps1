<#
.SYNOPSIS
    Gera ordens de serviço FICTÍCIAS de instalação e reparo de fibra, com "sujeira" realista.

.DESCRIPTION
    Cria um CSV (separador ";" e UTF-8 com BOM, como o Excel em pt-BR espera) com N dias de
    ordens de serviço, simulando a exportação de um portal de despacho.

    Tudo é inventado: cidades, pessoas e IDs. A mesma semente (-Seed) gera sempre a mesma
    estrutura, o que deixa testes e demonstrações reproduzíveis.

    Sujeira incluída de propósito, para exercitar a etapa de tratamento:
      - espaços sobrando no início/fim e espaços duplos;
      - datas em formatos mistos (ISO, dd/MM/yyyy, dd/MM/yy, dd-MM-yyyy e número serial do Excel);
      - horas como "8:05", "08:05" ou "08:05:00";
      - tipo em caixa variada ("REPARO", "instalação");
      - acentuação quebrada (UTF-8 lido como Windows-1252: "SÃ£o", "ConcluÃ­da");
      - linhas duplicadas.

.PARAMETER Days
    Quantidade de dias gerados. O último dia é "hoje" (parcial, até -ReferenceTime).

.PARAMETER OrdersPerDay
    Média de ordens por dia (varia ±10%).

.PARAMETER Seed
    Semente do gerador aleatório.

.PARAMETER ReferenceDate
    Data do último dia ("hoje" da simulação). Padrão: data atual.

.PARAMETER ReferenceTime
    Horário de "agora" no último dia. Ordens depois dele ficam sem chegada/início/término.

.PARAMETER DirtRate
    Probabilidade de cada campo de uma linha receber algum tipo de sujeira.

.PARAMETER DuplicateRate
    Fração de linhas duplicadas no arquivo.

.PARAMETER OutputPath
    Caminho do CSV gerado.

.EXAMPLE
    .\New-FakeServiceOrders.ps1
    Gera 7 dias com cerca de 240 ordens por dia em data-sample\orders-sample.csv.

.EXAMPLE
    .\New-FakeServiceOrders.ps1 -Days 30 -OrdersPerDay 400 -Seed 7 -OutputPath .\mes.csv
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 365)][int]$Days = 7,
    [ValidateRange(10, 5000)][int]$OrdersPerDay = 240,
    [int]$Seed = 42,
    [datetime]$ReferenceDate = (Get-Date).Date,
    [ValidatePattern('^\d{1,2}:\d{2}$')][string]$ReferenceTime = '14:35',   # meio da tarde: slot 14:00 em andamento, limites vencendo
    [ValidateRange(0.0, 1.0)][double]$DirtRate = 0.08,
    [ValidateRange(0.0, 0.5)][double]$DuplicateRate = 0.02,
    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# Padrão calculado aqui, e não no param: no PowerShell 5.1 com -File, $PSScriptRoot chega vazio no bloco param.
if (-not $OutputPath) { $OutputPath = Join-Path $PSScriptRoot 'orders-sample.csv' }

$script:Rng = New-Object System.Random($Seed)

# ---------------------------------------------------------------------------
# Catálogos fictícios
# ---------------------------------------------------------------------------

# Cidades inventadas por regional. Qualquer semelhança com nomes reais é coincidência.
$CitiesByRegion = [ordered]@{
    'A' = @('Vale do Cedro', 'Porto Lúmen', 'São Bento da Névoa')
    'B' = @('Campo Nimbo', 'Lagoa Serena', 'Conceição do Alto Vidal')
    'C' = @('Monte Albor', 'Santa Íris do Mar', 'Água Vidrada')
}

$FirstNames = @(
    'Ana', 'Bruno', 'Carla', 'Diego', 'Elisa', 'Fábio', 'Gabriela', 'Heitor', 'Íris', 'João',
    'Karina', 'Leandro', 'Marina', 'Nícolas', 'Otávio', 'Paula', 'Rafael', 'Sabrina', 'Tiago',
    'Úrsula', 'Vítor', 'Wesley', 'Yasmin', 'Caio', 'Débora', 'Renan', 'Lívia', 'Mateus',
    'Tânia', 'Igor', 'Luana', 'Sérgio', 'Bianca', 'André', 'Cecília', 'Davi'
)
$LastNames = @(
    'Albuquerque', 'Bastos', 'Cardoso', 'Damasceno', 'Esteves', 'Falcão', 'Guimarães', 'Holanda',
    'Ibiapina', 'Jordão', 'Leitão', 'Macedo', 'Nogueira', 'Ornelas', 'Pimentel', 'Quintela',
    'Rangel', 'Siqueira', 'Toledo', 'Valença', 'Xavier', 'Zanetti', 'Brandão', 'Coutinho'
)

# Slots agendados: início da janela e peso na distribuição das ordens do dia.
$Slots = @(
    @{ Hora = '08:30'; Peso = 35 },
    @{ Hora = '10:30'; Peso = 25 },
    @{ Hora = '14:00'; Peso = 25 },
    @{ Hora = '16:00'; Peso = 15 }
)
$WindowMinutes = 120   # a janela de cada slot dura 2 h

$Header = @('id_os', 'data', 'regional', 'cidade', 'gerente', 'coordenacao', 'supervisao',
            'tecnico', 'tipo', 'slot', 'chegada', 'inicio', 'termino', 'status')

# ---------------------------------------------------------------------------
# Funções auxiliares de sorteio
# ---------------------------------------------------------------------------

function Get-RandomItem {
    param([Parameter(Mandatory)][object[]]$Items)
    return $Items[$script:Rng.Next(0, $Items.Count)]
}

function Test-Chance {
    param([Parameter(Mandatory)][double]$Probability)
    return ($script:Rng.NextDouble() -lt $Probability)
}

function Get-RandomInt {
    param([int]$Min, [int]$Max)   # intervalo fechado [Min, Max]
    return $script:Rng.Next($Min, $Max + 1)
}

function Get-WeightedSlot {
    $total = ($Slots | ForEach-Object { $_.Peso } | Measure-Object -Sum).Sum
    $roll = $script:Rng.Next(0, $total)
    foreach ($slot in $Slots) {
        if ($roll -lt $slot.Peso) { return $slot.Hora }
        $roll -= $slot.Peso
    }
    return $Slots[-1].Hora
}

# Gera um nome completo que ainda não foi usado.
function New-UniquePersonName {
    param([Parameter(Mandatory)][hashtable]$Used)
    for ($i = 0; $i -lt 500; $i++) {
        $name = '{0} {1}' -f (Get-RandomItem $FirstNames), (Get-RandomItem $LastNames)
        if (-not $Used.ContainsKey($name)) {
            $Used[$name] = $true
            return $name
        }
    }
    throw 'Não foi possível gerar um nome único. Aumente as listas de nomes.'
}

# ---------------------------------------------------------------------------
# Estrutura da operação: regional > gerente > coordenação > supervisão > técnico
# ---------------------------------------------------------------------------

function New-Organization {
    $used = @{}
    $technicians = New-Object System.Collections.Generic.List[object]

    foreach ($region in $CitiesByRegion.Keys) {
        $manager = New-UniquePersonName -Used $used
        foreach ($c in 1..2) {
            $coordinator = New-UniquePersonName -Used $used
            foreach ($s in 1..2) {
                $supervisor = New-UniquePersonName -Used $used
                foreach ($t in 1..4) {
                    $technicians.Add([pscustomobject]@{
                        Regional    = $region
                        Cidade      = Get-RandomItem $CitiesByRegion[$region]
                        Gerente     = $manager
                        Coordenacao = $coordinator
                        Supervisao  = $supervisor
                        Tecnico     = New-UniquePersonName -Used $used
                        # Chance individual de chegar dentro da janela (uns são mais pontuais que outros).
                        Pontualidade = 0.62 + ($script:Rng.NextDouble() * 0.33)
                    })
                }
            }
        }
    }
    return $technicians
}

# ---------------------------------------------------------------------------
# Geração das ordens
# ---------------------------------------------------------------------------

# Minutos de atraso/adiantamento da chegada em relação ao início do slot.
function Get-ArrivalOffset {
    param([double]$Punctuality)
    if (Test-Chance $Punctuality) {
        return Get-RandomInt -Min -20 -Max ($WindowMinutes - 5)        # dentro da janela
    }
    return Get-RandomInt -Min ($WindowMinutes + 5) -Max ($WindowMinutes + 150)   # fora
}

# Monta uma ordem limpa (sem sujeira). $Now é o "agora" da simulação.
function New-Order {
    param(
        [Parameter(Mandatory)][int]$Number,
        [Parameter(Mandatory)][datetime]$Date,
        [Parameter(Mandatory)][object]$Technician,
        [Parameter(Mandatory)][datetime]$Now
    )

    $type = if (Test-Chance 0.55) { 'Instalação' } else { 'Reparo' }
    $slot = Get-WeightedSlot
    $slotStart = $Date.Add([timespan]::Parse($slot))

    $order = [ordered]@{
        id_os       = 'OS-{0:D7}' -f $Number
        data        = $Date
        regional    = $Technician.Regional
        cidade      = $Technician.Cidade
        gerente     = $Technician.Gerente
        coordenacao = $Technician.Coordenacao
        supervisao  = $Technician.Supervisao
        tecnico     = $Technician.Tecnico
        tipo        = $type
        slot        = $slot
        chegada     = $null
        inicio      = $null
        termino     = $null
        status      = 'Agendada'
    }

    if (Test-Chance 0.04) {
        $order.status = 'Cancelada'
        return $order
    }

    $arrival = $slotStart.AddMinutes((Get-ArrivalOffset -Punctuality $Technician.Pontualidade))
    $start   = $arrival.AddMinutes((Get-RandomInt -Min 3 -Max 25))
    $minutes = if ($type -eq 'Instalação') { Get-RandomInt 60 150 } else { Get-RandomInt 30 120 }
    $end     = $start.AddMinutes($minutes)
    $notDone = Test-Chance 0.06   # cliente ausente: chegou, mas não executou

    # Ainda não chegou no "agora" da simulação.
    if ($arrival -gt $Now) {
        $order.status = if ($slotStart.AddMinutes(-60) -le $Now) { 'Em deslocamento' } else { 'Agendada' }
        return $order
    }

    $order.chegada = $arrival
    if ($notDone) {
        $order.status = 'Não realizada'
        return $order
    }

    if ($start -gt $Now) {
        $order.status = 'Em execução'
        return $order
    }
    $order.inicio = $start

    if ($end -gt $Now) {
        $order.status = 'Em execução'
        return $order
    }
    $order.termino = $end
    $order.status  = 'Concluída'
    return $order
}

# ---------------------------------------------------------------------------
# Formatação e sujeira
# ---------------------------------------------------------------------------

# Simula texto UTF-8 aberto como Windows-1252 ("São" -> "SÃ£o").
function ConvertTo-Mojibake {
    param([string]$Text)
    $bytes = [Text.Encoding]::UTF8.GetBytes($Text)
    return [Text.Encoding]::GetEncoding(1252).GetString($bytes)
}

# Número serial do Excel (dias desde 30/12/1899), como aparece quando a célula perde o formato.
function ConvertTo-ExcelSerial {
    param([datetime]$Date)
    return [int]($Date.Date - [datetime]'1899-12-30').TotalDays
}

function Format-DirtyDate {
    param([datetime]$Date, [bool]$Dirty)
    $inv = [Globalization.CultureInfo]::InvariantCulture
    if (-not $Dirty) { return $Date.ToString('dd/MM/yyyy', $inv) }
    switch ($script:Rng.Next(0, 5)) {
        0 { return $Date.ToString('yyyy-MM-dd', $inv) }
        1 { return $Date.ToString('dd/MM/yy', $inv) }
        2 { return $Date.ToString('dd-MM-yyyy', $inv) }
        3 { return $Date.ToString('d/M/yyyy', $inv) }
        default { return [string](ConvertTo-ExcelSerial $Date) }
    }
}

function Format-DirtyTime {
    param($Value, [bool]$Dirty)
    if ($null -eq $Value) { return '' }
    $inv = [Globalization.CultureInfo]::InvariantCulture
    if (-not $Dirty) { return ([datetime]$Value).ToString('HH:mm', $inv) }
    if (Test-Chance 0.5) { return ([datetime]$Value).ToString('H:mm', $inv) }
    return ([datetime]$Value).ToString('HH:mm:ss', $inv)
}

# Aplica um tipo de sujeira de texto: espaços, caixa ou acento quebrado.
function Add-TextDirt {
    param([string]$Text)
    switch ($script:Rng.Next(0, 4)) {
        0 { return "  $Text " }
        1 { return ($Text -replace ' ', '  ') }
        2 { return (ConvertTo-Mojibake $Text) }
        default { return $Text.ToUpperInvariant() }
    }
}

# Converte a ordem limpa em uma linha de texto (já com sujeira aleatória).
function ConvertTo-DirtyRow {
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Order)

    $d = { Test-Chance $DirtRate }
    $row = [ordered]@{}
    $row.id_os       = if (& $d) { " $($Order.id_os)" } else { $Order.id_os }
    $row.data        = Format-DirtyDate -Date $Order.data -Dirty (& $d)
    $row.regional    = if (& $d) { $Order.regional.ToLowerInvariant() } else { $Order.regional }
    foreach ($field in 'cidade', 'gerente', 'coordenacao', 'supervisao', 'tecnico', 'tipo', 'status') {
        $row[$field] = if (& $d) { Add-TextDirt $Order[$field] } else { $Order[$field] }
    }
    $row.slot    = if (& $d) { $Order.slot.TrimStart('0') } else { $Order.slot }
    $row.chegada = Format-DirtyTime -Value $Order.chegada -Dirty (& $d)
    $row.inicio  = Format-DirtyTime -Value $Order.inicio  -Dirty (& $d)
    $row.termino = Format-DirtyTime -Value $Order.termino -Dirty (& $d)

    return ($Header | ForEach-Object { ConvertTo-CsvField $row[$_] }) -join ';'
}

function ConvertTo-CsvField {
    param($Value)
    $text = [string]$Value
    if ($text -match '[;"\r\n]') { return '"' + $text.Replace('"', '""') + '"' }
    return $text
}

# ---------------------------------------------------------------------------
# Execução
# ---------------------------------------------------------------------------

$now = $ReferenceDate.Date.Add([timespan]::Parse($ReferenceTime))
$organization = New-Organization
Write-Verbose ("Estrutura: {0} técnicos em {1} regionais" -f $organization.Count, $CitiesByRegion.Count)

$lines = New-Object System.Collections.Generic.List[string]
$number = 4100000 + $script:Rng.Next(0, 90000)
$clean = 0

for ($offset = $Days - 1; $offset -ge 0; $offset--) {
    $date  = $ReferenceDate.Date.AddDays(-$offset)
    $count = [int]($OrdersPerDay * (0.9 + $script:Rng.NextDouble() * 0.2))
    for ($i = 0; $i -lt $count; $i++) {
        $number++
        $order = New-Order -Number $number -Date $date -Technician (Get-RandomItem $organization) -Now $now
        $line = ConvertTo-DirtyRow -Order $order
        $lines.Add($line)
        $clean++

        # Duplica algumas linhas logo em seguida (como uma exportação repetida).
        if (Test-Chance $DuplicateRate) { $lines.Add($line) }
    }
}

$outDir = Split-Path -Parent $OutputPath
if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

$content = @($Header -join ';') + $lines
[IO.File]::WriteAllLines($OutputPath, [string[]]$content, (New-Object System.Text.UTF8Encoding($true)))

Write-Host ("Gerado: {0}" -f $OutputPath)
Write-Host ("  Ordens únicas: {0} | linhas no arquivo: {1} | duplicadas: {2}" -f $clean, $lines.Count, ($lines.Count - $clean))
Write-Host ("  Período: {0:dd/MM/yyyy} a {1:dd/MM/yyyy} | 'agora' da simulação: {2:dd/MM/yyyy HH:mm}" -f $ReferenceDate.AddDays(1 - $Days), $ReferenceDate, $now)
