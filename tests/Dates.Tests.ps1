BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\excel-pipeline\lib\FieldOps.psm1') -Force
}

Describe 'ConvertTo-DateSafe' {

    It 'reconhece "<Texto>" como 12/10/2026' -TestCases @(
        @{ Texto = '12/10/2026' }
        @{ Texto = '12/10/26' }
        @{ Texto = '2026-10-12' }
        @{ Texto = '12-10-2026' }
        @{ Texto = '2026/10/12' }
        @{ Texto = '  12/10/2026 ' }
        @{ Texto = '46307' }          # serial do Excel
    ) {
        (ConvertTo-DateSafe $Texto).Date | Should -Be ([datetime]'2026-10-12')
    }

    It 'aceita dia e mês sem zero à esquerda' {
        ConvertTo-DateSafe '2/9/2026' | Should -Be ([datetime]'2026-09-02')
    }

    It 'aceita o serial numérico vindo de Value2 (double com hora)' {
        ConvertTo-DateSafe 46307.5 | Should -Be ([datetime]'2026-10-12 12:00')
    }

    It 'não depende da cultura do Windows (en-US não troca dia e mês)' {
        $original = [Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'
            ConvertTo-DateSafe '05/10/2026' | Should -Be ([datetime]'2026-10-05')
        }
        finally {
            [Threading.Thread]::CurrentThread.CurrentCulture = $original
        }
    }

    It 'devolve $null para "<Texto>"' -TestCases @(
        @{ Texto = '' }
        @{ Texto = '   ' }
        @{ Texto = 'ontem' }
        @{ Texto = '31/02/2026' }
        @{ Texto = '12' }             # número pequeno demais para ser data
    ) {
        ConvertTo-DateSafe $Texto | Should -BeNullOrEmpty
    }

    It 'devolve $null para $null' {
        ConvertTo-DateSafe $null | Should -BeNullOrEmpty
    }
}

Describe 'ConvertTo-TimeOfDay' {

    It 'converte "<Texto>" em 08:05' -TestCases @(
        @{ Texto = '8:05' }
        @{ Texto = '08:05' }
        @{ Texto = '08:05:00' }
        @{ Texto = '0805' }
    ) {
        ConvertTo-TimeOfDay $Texto | Should -Be ([timespan]'08:05')
    }

    It 'converte fração de dia do Excel (0,5 = 12:00)' {
        ConvertTo-TimeOfDay 0.5 | Should -Be ([timespan]'12:00')
    }

    It 'devolve $null para texto inválido' {
        ConvertTo-TimeOfDay 'meio-dia' | Should -BeNullOrEmpty
        ConvertTo-TimeOfDay '' | Should -BeNullOrEmpty
    }
}
