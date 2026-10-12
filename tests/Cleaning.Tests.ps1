BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\excel-pipeline\lib\FieldOps.psm1') -Force

    # Gera mojibake real (UTF-8 lido como Windows-1252), igual ao das exportações.
    function Break-Accents([string]$Text) {
        [Text.Encoding]::GetEncoding(1252).GetString([Text.Encoding]::UTF8.GetBytes($Text))
    }
}

Describe 'Repair-Mojibake' {

    It 'repara "<Original>"' -TestCases @(
        @{ Original = 'São Bento da Névoa' }
        @{ Original = 'Concluída' }
        @{ Original = 'Conceição do Alto Vidal' }
        @{ Original = 'Santa Íris do Mar' }      # Í gera um byte sem caractere no 1252
        @{ Original = 'Instalação' }
    ) {
        Repair-Mojibake (Break-Accents $Original) | Should -BeExactly $Original
    }

    It 'não altera texto que já está correto' {
        Repair-Mojibake 'Água Vidrada' | Should -BeExactly 'Água Vidrada'
        Repair-Mojibake 'Reparo' | Should -BeExactly 'Reparo'
    }
}

Describe 'Format-CleanText e Format-ProperName' {

    It 'remove espaços nas pontas e espaços duplos' {
        Format-CleanText '  Marina   Coutinho ' | Should -BeExactly 'Marina Coutinho'
    }

    It 'padroniza nome próprio com preposições minúsculas' {
        Format-ProperName '  SÃO BENTO DA NÉVOA ' | Should -BeExactly 'São Bento da Névoa'
    }

    It 'repara acento e caixa ao mesmo tempo' {
        Format-ProperName (Break-Accents 'CONCEIÇÃO DO ALTO VIDAL') | Should -BeExactly 'Conceição do Alto Vidal'
    }
}

Describe 'Padronização de tipo, status e slot' {

    It 'tipo "<Texto>" vira "<Esperado>"' -TestCases @(
        @{ Texto = 'INSTALAÇÃO';    Esperado = 'Instalação' }
        @{ Texto = ' instalacao ';  Esperado = 'Instalação' }
        @{ Texto = 'InstalaÃ§Ã£o';  Esperado = 'Instalação' }
        @{ Texto = 'REPARO';        Esperado = 'Reparo' }
    ) {
        ConvertTo-OrderType $Texto | Should -BeExactly $Esperado
    }

    It 'tipo desconhecido vira $null' {
        ConvertTo-OrderType 'Vistoria' | Should -BeNullOrEmpty
    }

    It 'status "<Texto>" vira "<Esperado>"' -TestCases @(
        @{ Texto = 'Não  realizada';  Esperado = 'Não realizada' }
        @{ Texto = 'NÃ£o realizada';  Esperado = 'Não realizada' }
        @{ Texto = 'CONCLUÍDA';       Esperado = 'Concluída' }
        @{ Texto = ' em execução ';   Esperado = 'Em execução' }
    ) {
        ConvertTo-OrderStatus $Texto | Should -BeExactly $Esperado
    }

    It 'slot "8:30" vira "08:30"' {
        ConvertTo-SlotText '8:30' | Should -BeExactly '08:30'
    }
}

Describe 'ConvertTo-CleanOrder' {

    It 'limpa e tipa uma linha suja' {
        $row = [pscustomobject]@{
            id_os = ' os-4100001'; data = '46307'; regional = 'b'; cidade = '  CAMPO  NIMBO '
            gerente = 'Ana Bastos'; coordenacao = 'Caio Rangel'; supervisao = 'Íris Toledo'; tecnico = 'Davi Xavier'
            tipo = 'REPARO'; slot = '8:30'; chegada = '9:47:00'; inicio = '10:01'; termino = ''; status = 'Em  execução'
        }
        $order = ConvertTo-CleanOrder $row

        $order.IdOs      | Should -BeExactly 'OS-4100001'
        $order.Data      | Should -Be ([datetime]'2026-10-12')
        $order.Regional  | Should -BeExactly 'B'
        $order.Cidade    | Should -BeExactly 'Campo Nimbo'
        $order.Tipo      | Should -BeExactly 'Reparo'
        $order.Slot      | Should -BeExactly '08:30'
        $order.Chegada   | Should -Be ([datetime]'2026-10-12 09:47')
        $order.Termino   | Should -BeNullOrEmpty
        $order.Status    | Should -BeExactly 'Em execução'
        $order.Problemas | Should -HaveCount 0
    }

    It 'registra problemas quando a data é inválida' {
        $row = [pscustomobject]@{
            id_os = 'OS-1'; data = 'ontem'; regional = 'A'; cidade = 'X'; gerente = ''; coordenacao = ''
            supervisao = ''; tecnico = ''; tipo = 'Reparo'; slot = '08:30'; chegada = ''; inicio = ''; termino = ''; status = 'Agendada'
        }
        (ConvertTo-CleanOrder $row).Problemas | Should -Contain "data inválida 'ontem'"
    }
}

Describe 'Remove-DuplicateOrder' {

    It 'mantém só a primeira ocorrência de cada IdOs' {
        $orders = @(
            [pscustomobject]@{ IdOs = 'OS-1'; N = 1 }
            [pscustomobject]@{ IdOs = 'OS-2'; N = 2 }
            [pscustomobject]@{ IdOs = 'OS-1'; N = 3 }
        )
        $result = @(Remove-DuplicateOrder -Orders $orders)
        $result | Should -HaveCount 2
        ($result | Where-Object IdOs -eq 'OS-1').N | Should -Be 1
    }
}

Describe 'Import-ServiceOrder com o gerador de dados fictícios' {

    BeforeAll {
        $csv = Join-Path $TestDrive 'ordens.csv'
        $generator = Join-Path $PSScriptRoot '..\data-sample\New-FakeServiceOrders.ps1'
        & $generator -Days 3 -OrdersPerDay 150 -Seed 7 -ReferenceDate '2026-10-12' -DirtRate 0.15 -DuplicateRate 0.05 -OutputPath $csv 6>$null
        $result = Import-ServiceOrder -Path $csv
    }

    It 'não rejeita nenhuma linha (toda sujeira gerada é tratável)' {
        $result.Stats.Rejeitadas | Should -Be 0
    }

    It 'remove as duplicadas' {
        $result.Stats.Duplicadas | Should -BeGreaterThan 0
        @($result.Orders.IdOs | Select-Object -Unique).Count | Should -Be $result.Orders.Count
    }

    It 'normaliza regionais, tipos e cidades' {
        @($result.Orders.Regional | Sort-Object -Unique) | Should -Be @('A', 'B', 'C')
        @($result.Orders.Tipo | Sort-Object -Unique) | Should -Be @('Instalação', 'Reparo')
        @($result.Orders.Cidade | Sort-Object -Unique).Count | Should -Be 9
    }

    It 'não deixa acentuação quebrada nem espaços duplos' {
        $all = ($result.Orders | ForEach-Object { $_.Cidade, $_.Tecnico, $_.Supervisao, $_.Status }) -join '|'
        $all | Should -Not -MatchExactly 'Ã'
        $all | Should -Not -Match '  '
    }
}
