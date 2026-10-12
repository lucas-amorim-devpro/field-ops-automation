BeforeDiscovery {
    $script:CanBuildWorkbook = ($null -ne [Type]::GetTypeFromProgID('Excel.Application')) -and
        ((Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security' -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM -eq 1)
}

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\excel-pipeline\lib\FieldOps.psm1') -Force
}

Describe 'Set-MTextParameter' {

    BeforeAll {
        $m = "let`r`n    CaminhoCsv = ""C:\antigo\ordens.csv"",`r`n    Fonte = Csv.Document(File.Contents(CaminhoCsv))`r`nin`r`n    Fonte"
    }

    It 'troca só o valor do parâmetro' {
        $new = Set-MTextParameter -Formula $m -Name 'CaminhoCsv' -Value 'D:\novo\base.csv'
        $new | Should -Match 'CaminhoCsv = "D:\\novo\\base\.csv",'
        $new | Should -Match 'File\.Contents\(CaminhoCsv\)'
    }

    It 'escapa aspas dobrando (regra do M)' {
        Set-MTextParameter -Formula $m -Name 'CaminhoCsv' -Value 'a"b' | Should -Match 'CaminhoCsv = "a""b"'
    }

    It 'falha se o parâmetro não existir' {
        { Set-MTextParameter -Formula $m -Name 'Inexistente' -Value 'x' } | Should -Throw '*não encontrado*'
    }
}

Describe 'ConvertTo-CompactJson' {

    It 'serializa estruturas aninhadas sem desenrolar listas de um item' {
        $obj = [ordered]@{ a = 1; b = @(, @('x')); c = $null; d = $true }
        ConvertTo-CompactJson $obj | Should -BeExactly '{"a":1,"b":[["x"]],"c":null,"d":true}'
    }

    It 'escapa aspas, barras e quebras de linha' {
        ConvertTo-CompactJson "a""b\c`nd" | Should -BeExactly '"a\"b\\c\nd"'
    }

    It 'usa ponto decimal mesmo com o Windows em pt-BR' {
        $original = [Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [Threading.Thread]::CurrentThread.CurrentCulture = 'pt-BR'
            ConvertTo-CompactJson @(88.05, [decimal]7.38) | Should -BeExactly '[88.05,7.38]'
        }
        finally { [Threading.Thread]::CurrentThread.CurrentCulture = $original }
    }

    It 'mantém acentos' {
        ConvertTo-CompactJson 'Instalação' | Should -BeExactly '"Instalação"'
    }
}

Describe 'Send-AppsScriptPayload' {

    It 'recusa URL que não é de Web App do Apps Script (sem fazer requisição)' {
        { Send-AppsScriptPayload -Url 'https://exemplo.com/api' -Payload @{ a = 1 } } | Should -Throw '*URL do Apps Script inválida*'
    }
}

Describe 'Pipeline de ponta a ponta (matriz real, dry-run)' -Tag 'Excel' -Skip:(-not $script:CanBuildWorkbook) {

    BeforeAll {
        $repo = Split-Path -Parent $PSScriptRoot
        $excelBefore = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)   # instâncias do usuário
        $matrix = Join-Path $TestDrive 'matriz.xlsm'
        & (Join-Path $repo 'excel-pipeline\Build-DemoWorkbook.ps1') -OutputPath $matrix 6>$null | Out-Null

        $config = [IO.File]::ReadAllText((Join-Path $repo 'excel-pipeline\config.example.json')) | ConvertFrom-Json
        $config.matriz = $matrix
        $config.csv = Join-Path $repo 'data-sample\orders-sample.csv'
        $config.pastaLogs = Join-Path $TestDrive 'logs'
        $config.pastaSaida = Join-Path $TestDrive 'out'
        $configPath = Join-Path $TestDrive 'config.json'
        [IO.File]::WriteAllText($configPath, ($config | ConvertTo-Json -Depth 5))

        & (Join-Path $repo 'excel-pipeline\Invoke-SlotPipeline.ps1') -ConfigPath $configPath -DryRun 6>$null | Out-Null
        $payloadFile = Get-ChildItem (Join-Path $TestDrive 'out') -Filter 'payload_*.json' | Select-Object -First 1
        $payload = [IO.File]::ReadAllText($payloadFile.FullName) | ConvertFrom-Json
    }

    It 'gera o payload com o total esperado e sem o token' {
        $payload.token | Should -Be '***'
        $total = $payload.linhas | Where-Object { $_[0] -eq 'Total' }
        $total[6] | Should -Be 88.05
        $total[7] | Should -Be 'Na meta'
    }

    It 'calcula todas as regionais, slots e tipos' {
        @($payload.linhas | Where-Object { $_[0] -eq 'Regional' }).Count | Should -Be 3
        @($payload.linhas | Where-Object { $_[0] -eq 'Slot' }).Count | Should -Be 4
        @($payload.linhas | Where-Object { $_[0] -eq 'Tipo' }).Count | Should -Be 2
    }

    It 'não deixa EXCEL.EXE fantasma (nenhum processo novo continua aberto)' {
        $leftovers = Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $excelBefore -notcontains $_.Id }
        $leftovers | Should -BeNullOrEmpty
    }
}
