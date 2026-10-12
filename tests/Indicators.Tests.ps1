BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\excel-pipeline\lib\FieldOps.psm1') -Force

    # Monta uma ordem já limpa para os testes.
    function New-TestOrder {
        param(
            [string]$Tipo = 'Instalação',
            [string]$Slot = '08:30',
            [string]$Chegada,
            [string]$Status = 'Concluída',
            [string]$Regional = 'A'
        )
        [pscustomobject]@{
            Data     = [datetime]'2026-10-12'
            Regional = $Regional
            Tipo     = $Tipo
            Slot     = $Slot
            Chegada  = if ($Chegada) { [datetime]"2026-10-12 $Chegada" } else { $null }
            Status   = $Status
        }
    }
}

Describe 'Get-RoundedHalfUp (igual ao ARRED do Google Sheets)' {

    It '<Valor> com 2 casas vira <Esperado>' -TestCases @(
        @{ Valor = 0.125;  Esperado = 0.13 }    # [math]::Round daria 0,12 (bancário)
        @{ Valor = 2.675;  Esperado = 2.68 }    # em double é 2,67499999...
        @{ Valor = 80.145; Esperado = 80.15 }
        @{ Valor = 1.005;  Esperado = 1.01 }
        @{ Valor = -0.125; Esperado = -0.13 }   # meio "para longe do zero", como o Sheets
        @{ Valor = 80;     Esperado = 80 }
    ) {
        Get-RoundedHalfUp $Valor 2 | Should -Be ([decimal]$Esperado)
    }
}

Describe 'Get-SlotDeadline' {

    It 'Instalação: limite = slot + 2 h' {
        Get-SlotDeadline (New-TestOrder -Tipo 'Instalação' -Slot '08:30') | Should -Be ([datetime]'2026-10-12 10:30')
    }

    It 'Reparo: limite = slot + 2 h + 30 min de tolerância' {
        Get-SlotDeadline (New-TestOrder -Tipo 'Reparo' -Slot '14:00') | Should -Be ([datetime]'2026-10-12 16:30')
    }
}

Describe 'Get-PunctualityStatus' {

    BeforeAll { $now = [datetime]'2026-10-12 13:10' }

    It 'chegada antes do limite conta no prazo (inclusive adiantada)' {
        Get-PunctualityStatus (New-TestOrder -Chegada '08:10') -Now $now | Should -Be 'NoPrazo'
    }

    It 'chegada exatamente no limite conta no prazo' {
        Get-PunctualityStatus (New-TestOrder -Chegada '10:30') -Now $now | Should -Be 'NoPrazo'
    }

    It 'chegada depois do limite conta fora do prazo' {
        Get-PunctualityStatus (New-TestOrder -Chegada '10:31') -Now $now | Should -Be 'ForaDoPrazo'
    }

    It 'sem chegada e limite vencido conta fora do prazo (estourou)' {
        Get-PunctualityStatus (New-TestOrder -Slot '10:30' -Status 'Em deslocamento') -Now $now | Should -Be 'ForaDoPrazo'
    }

    It 'sem chegada e limite ainda no futuro fica pendente' {
        Get-PunctualityStatus (New-TestOrder -Slot '14:00' -Status 'Em deslocamento') -Now $now | Should -Be 'Pendente'
    }

    It 'cancelada é excluída do indicador' {
        Get-PunctualityStatus (New-TestOrder -Chegada '12:00' -Status 'Cancelada') -Now $now | Should -Be 'Excluida'
    }
}

Describe 'Get-GoalBadge (meta 80%, margem 5 p.p.)' {

    It '<Percentual> -> <Selo>' -TestCases @(
        @{ Percentual = 92;    Selo = 'Na meta' }
        @{ Percentual = 80;    Selo = 'Na meta' }
        @{ Percentual = 79.99; Selo = 'Perto' }
        @{ Percentual = 75;    Selo = 'Perto' }
        @{ Percentual = 74.99; Selo = 'Abaixo' }
    ) {
        Get-GoalBadge -Percent $Percentual | Should -Be $Selo
    }

    It 'sem dados quando não há ordens apuradas' {
        Get-GoalBadge -Percent $null | Should -Be 'Sem dados'
    }
}

Describe 'Get-PunctualitySummary' {

    BeforeAll {
        $now = [datetime]'2026-10-12 13:10'
        $orders = @(
            New-TestOrder -Regional 'A' -Chegada '08:40'                          # no prazo
            New-TestOrder -Regional 'A' -Chegada '09:50'                          # no prazo
            New-TestOrder -Regional 'A' -Chegada '11:15'                          # fora
            New-TestOrder -Regional 'B' -Chegada '09:00'                          # no prazo
            New-TestOrder -Regional 'B' -Slot '16:00' -Status 'Agendada'          # pendente
            New-TestOrder -Regional 'B' -Chegada '09:00' -Status 'Cancelada'      # excluída
        )
    }

    It 'calcula o total geral' {
        $total = Get-PunctualitySummary -Orders $orders -Now $now
        $total.Apuradas   | Should -Be 4
        $total.NoPrazo    | Should -Be 3
        $total.Fora       | Should -Be 1
        $total.Pendentes  | Should -Be 1
        $total.Percentual | Should -Be 75
        $total.Selo       | Should -Be 'Perto'
    }

    It 'agrupa por regional com 2 casas (2/3 = 66,67%)' {
        $byRegion = Get-PunctualitySummary -Orders $orders -Now $now -GroupBy Regional
        ($byRegion | Where-Object Grupo -eq 'A').Percentual | Should -Be ([decimal]66.67)
        ($byRegion | Where-Object Grupo -eq 'B').Percentual | Should -Be 100
    }

    It 'devolve percentual nulo quando nada foi apurado' {
        $pending = @(New-TestOrder -Slot '16:00' -Status 'Agendada')
        $result = Get-PunctualitySummary -Orders $pending -Now $now
        $result.Percentual | Should -BeNullOrEmpty
        $result.Selo | Should -Be 'Sem dados'
    }
}

Describe 'Get-VariationPp' {

    It 'calcula a diferença em pontos percentuais' {
        Get-VariationPp -Current 82.5 -Previous 79.25 | Should -Be ([decimal]3.25)
        Get-VariationPp -Current 70 -Previous 75.555 | Should -Be ([decimal]-5.56)
    }

    It 'devolve $null se faltar um dos lados' {
        Get-VariationPp -Current 80 -Previous $null | Should -BeNullOrEmpty
    }
}
