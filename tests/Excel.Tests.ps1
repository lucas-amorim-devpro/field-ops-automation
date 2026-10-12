# Testes de integração com o Excel real (tag "Excel").
# São pulados automaticamente em máquinas sem Excel (ex. CI).

BeforeDiscovery {
    $script:HasExcel = $null -ne [Type]::GetTypeFromProgID('Excel.Application')
}

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\excel-pipeline\lib\FieldOps.psm1') -Force
}

Describe 'Excel via COM' -Tag 'Excel' -Skip:(-not $script:HasExcel) {

    BeforeAll {
        $excel = New-ExcelApplication
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
    }

    AfterAll {
        Close-ExcelSafely -Excel $excel -Workbook $workbook
    }

    It 'Set-CellValue2 grava texto e depois número na mesma variável COM sem InvalidCast' {
        $cell = $sheet.Range('A1')
        Set-CellValue2 -Range $cell -Value 'texto'
        Set-CellValue2 -Range $cell -Value 87.5
        Get-CellValue2 -Range $cell | Should -Be 87.5
        Get-CellValue2 -Range $cell | Should -BeOfType [double]
    }

    It 'Value2 devolve datas como serial, que ConvertTo-DateSafe entende' {
        $cell = $sheet.Range('A2')
        $cell.Formula = '=DATE(2026,10,12)'
        ConvertTo-DateSafe (Get-CellValue2 -Range $cell) | Should -Be ([datetime]'2026-10-12')
    }

    It 'ConvertFrom-Value2Table lê uma faixa inteira em uma chamada' {
        $sheet.Range('C1').Value2 = 'regional'
        $sheet.Range('D1').Value2 = 'percentual'
        $sheet.Range('C2').Value2 = 'A'
        Set-CellValue2 -Range $sheet.Range('D2') -Value 81.25
        $sheet.Range('C3').Value2 = 'B'
        Set-CellValue2 -Range $sheet.Range('D3') -Value 77

        $rows = @(ConvertFrom-Value2Table (Get-CellValue2 -Range $sheet.Range('C1:D3')))
        $rows | Should -HaveCount 2
        $rows[0].regional | Should -Be 'A'
        $rows[1].percentual | Should -Be 77
    }
}
