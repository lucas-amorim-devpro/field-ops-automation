# ---------------------------------------------------------------------------
# Envio ao Google Apps Script (Web App com doPost) e parâmetros do Power Query
#
#   - O doPost do Apps Script NÃO enxerga cabeçalhos HTTP (e.postData só traz o corpo).
#     Por isso o token de autenticação vai dentro do JSON, nunca em um header.
#   - A URL /exec responde 302 para script.googleusercontent.com; o Invoke-WebRequest
#     segue o redirecionamento e devolve a resposta final do script.
#   - O corpo é enviado como bytes UTF-8: passar string no 5.1 pode sair em ANSI.
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Troca o valor de um parâmetro de texto no código M de uma consulta: Nome = "valor".
.DESCRIPTION
    Função pura (só texto), usada para apontar a consulta para o CSV certo em cada PC,
    no lugar de ler o caminho de uma célula (que dispara Formula.Firewall).
.EXAMPLE
    Set-MTextParameter -Formula $q.Formula -Name 'CaminhoCsv' -Value 'D:\dados\ordens.csv'
#>
function Set-MTextParameter {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][string]$Formula,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Value
    )
    $pattern = '(?m)^(\s*' + [regex]::Escape($Name) + '\s*=\s*)"(?:[^"]|"")*"'
    if ($Formula -notmatch $pattern) { throw "Parâmetro '$Name' não encontrado no código M." }
    $escaped = $Value.Replace('"', '""')   # aspas em M são escapadas dobrando
    $regex = New-Object System.Text.RegularExpressions.Regex($pattern)
    $evaluator = [System.Text.RegularExpressions.MatchEvaluator] { param($m) $m.Groups[1].Value + '"' + $escaped + '"' }
    return $regex.Replace($Formula, $evaluator, 1)
}

<#
.SYNOPSIS
    Envia um objeto como JSON ao Web App do Apps Script, com novas tentativas e espera crescente.
.OUTPUTS
    Objeto da resposta (o doPost devolve JSON {"ok":true,...}).
#>
function Send-AppsScriptPayload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][object]$Payload,
        [int]$TimeoutSec = 30,
        [ValidateRange(1, 10)][int]$Attempts = 3
    )
    if ($Url -notmatch '^https://script\.google\.com/macros/s/[^/]+/exec$') {
        throw "URL do Apps Script inválida. Esperado: https://script.google.com/macros/s/<ID>/exec"
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $body = [Text.Encoding]::UTF8.GetBytes((ConvertTo-CompactJson $Payload))

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            $response = Invoke-WebRequest -Uri $Url -Method Post -Body $body -ContentType 'application/json; charset=utf-8' `
                -TimeoutSec $TimeoutSec -UseBasicParsing
            $result = $response.Content | ConvertFrom-Json
            if (-not $result.ok) { throw "Apps Script recusou: $($result.erro)" }
            return $result
        }
        catch {
            $detail = Get-ErrorDetail $_
            if ($attempt -eq $Attempts) { throw "Envio falhou após $Attempts tentativa(s): $detail" }
            $wait = [math]::Pow(2, $attempt)
            Write-StepLog -Step 'Enviar ao Apps Script' -Level AVISO -Message "tentativa $attempt falhou ($detail); nova tentativa em $wait s"
            Start-Sleep -Seconds $wait
        }
    }
}
