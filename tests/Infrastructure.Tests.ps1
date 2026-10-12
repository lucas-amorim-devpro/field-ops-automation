BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\excel-pipeline\lib\FieldOps.psm1') -Force
    $repoRoot = Split-Path -Parent $PSScriptRoot
}

Describe 'Codificação dos scripts' {

    It 'todo .ps1/.psm1/.psd1 está em UTF-8 com BOM (exigência do PowerShell 5.1)' {
        $withoutBom = Get-ChildItem $repoRoot -Recurse -File -Include *.ps1, *.psm1, *.psd1 |
            Where-Object {
                $b = [IO.File]::ReadAllBytes($_.FullName)
                -not ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)
            } |
            ForEach-Object { $_.FullName.Substring($repoRoot.Length + 1) }

        $withoutBom | Should -BeNullOrEmpty -Because 'rode tools\Convert-ToUtf8Bom.ps1'
    }

    It 'nenhum arquivo de texto tem U+2028/U+2029 literal (o editor do Apps Script os remove e quebra o código)' {
        $extensions = '.ps1', '.psm1', '.gs', '.js', '.html', '.css', '.json', '.md', '.pq', '.bas'
        $withSeparators = Get-ChildItem $repoRoot -Recurse -File |
            Where-Object { $extensions -contains $_.Extension -and $_.FullName -notmatch '\\(\.git|out|logs|para-claude-web)\\' } |
            Where-Object { [IO.File]::ReadAllText($_.FullName) -match "[\u2028\u2029]" } |
            ForEach-Object { $_.FullName.Substring($repoRoot.Length + 1) }

        $withSeparators | Should -BeNullOrEmpty -Because 'escreva os escapes \u2028/\u2029 como texto'
    }
}

Describe 'Log por etapa' {

    BeforeAll { Start-StepLog -Directory $TestDrive -Name 'teste' 6>$null | Out-Null }

    It 'Invoke-Step devolve o resultado do bloco' {
        Invoke-Step 'Somar' { 1 + 1 } 6>$null | Should -Be 2
    }

    It 'Invoke-Step relança o erro e registra no arquivo' {
        { Invoke-Step 'Etapa quebrada' { throw 'falha simulada' } 6>$null } | Should -Throw '*falha simulada*'
        $log = Get-ChildItem $TestDrive -Filter 'teste_*.log' | Select-Object -First 1
        Get-Content $log.FullName -Encoding UTF8 -Raw | Should -Match 'ERRO\s+\| Etapa quebrada'
    }

    It 'o log preserva acentos' {
        Write-StepLog -Step 'Acentuação' -Message 'Conceição, Névoa' 6>$null
        $log = Get-ChildItem $TestDrive -Filter 'teste_*.log' | Select-Object -First 1
        Get-Content $log.FullName -Encoding UTF8 -Raw | Should -Match 'Conceição, Névoa'
    }
}

Describe 'Disable-QuickEdit' {

    It 'não lança erro mesmo sem console interativo' {
        { Disable-QuickEdit } | Should -Not -Throw
        Disable-QuickEdit | Should -BeOfType [bool]
    }
}
