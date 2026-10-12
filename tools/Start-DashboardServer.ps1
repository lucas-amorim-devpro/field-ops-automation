<#
.SYNOPSIS
    Servidor HTTP local mínimo para testar o painel (sem Node, sem Python, sem IIS).

.DESCRIPTION
    O painel usa fetch() para ler o JSON, e o navegador bloqueia fetch em páginas abertas
    direto do disco (file://). Este script serve a pasta dashboard-pages em
    http://localhost:<Port>/ usando System.Net.HttpListener, que já vem no Windows.
    Só escuta em localhost; Ctrl+C encerra.

.EXAMPLE
    .\tools\Start-DashboardServer.ps1
    .\tools\Start-DashboardServer.ps1 -Port 8090 -Root .\dashboard-pages
#>
[CmdletBinding()]
param(
    [int]$Port = 8080,
    [string]$Root
)

$ErrorActionPreference = 'Stop'
# Padrão calculado aqui, e não no param: no PowerShell 5.1 com -File, $PSScriptRoot chega vazio no bloco param.
if (-not $Root) { $Root = Join-Path (Split-Path -Parent $PSScriptRoot) 'dashboard-pages' }
$Root = (Resolve-Path $Root).Path

$mimeTypes = @{
    '.html' = 'text/html; charset=utf-8'
    '.css'  = 'text/css; charset=utf-8'
    '.js'   = 'text/javascript; charset=utf-8'
    '.gs'   = 'text/javascript; charset=utf-8'   # Apps Script é JavaScript: permite testar no navegador
    '.csv'  = 'text/csv; charset=utf-8'
    '.json' = 'application/json; charset=utf-8'
    '.svg'  = 'image/svg+xml'
    '.png'  = 'image/png'
    '.gif'  = 'image/gif'
    '.ico'  = 'image/x-icon'
}

# Resolve o caminho pedido dentro de $Root, recusando "..\" para fora da pasta.
function Resolve-RequestPath {
    param([string]$UrlPath)
    $relative = [Uri]::UnescapeDataString($UrlPath.TrimStart('/')).Replace('/', '\')
    if ($relative -eq '' -or $relative.EndsWith('\')) { $relative += 'index.html' }
    $full = [IO.Path]::GetFullPath((Join-Path $Root $relative))
    if (-not $full.StartsWith($Root, [StringComparison]::OrdinalIgnoreCase)) { return $null }
    return $full
}

function Send-Response {
    param($Context, [int]$Status, [byte[]]$Body, [string]$ContentType = 'text/plain; charset=utf-8')
    $response = $Context.Response
    $response.StatusCode = $Status
    $response.ContentType = $ContentType
    $response.Headers['Cache-Control'] = 'no-store'
    $response.ContentLength64 = $Body.Length
    # HEAD só recebe os cabeçalhos; escrever o corpo violaria o protocolo.
    if ($Context.Request.HttpMethod -ne 'HEAD') { $response.OutputStream.Write($Body, 0, $Body.Length) }
    $response.OutputStream.Close()
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()
Write-Host "Servindo $Root em http://localhost:$Port/  (Ctrl+C para parar)"

try {
    while ($listener.IsListening) {
        # GetContextAsync + Wait(500) mantém o Ctrl+C responsivo no console.
        $task = $listener.GetContextAsync()
        while (-not $task.Wait(500)) { }
        $context = $task.Result
        try {
            $path = Resolve-RequestPath $context.Request.Url.AbsolutePath
            if ($path -and (Test-Path $path -PathType Leaf)) {
                $ext = [IO.Path]::GetExtension($path).ToLowerInvariant()
                $type = if ($mimeTypes.ContainsKey($ext)) { $mimeTypes[$ext] } else { 'application/octet-stream' }
                Send-Response $context 200 ([IO.File]::ReadAllBytes($path)) $type
                $status = 200
            }
            else {
                Send-Response $context 404 ([Text.Encoding]::UTF8.GetBytes('Não encontrado'))
                $status = 404
            }
            Write-Host ('{0:HH:mm:ss} {1} {2} {3}' -f (Get-Date), $status, $context.Request.HttpMethod, $context.Request.Url.AbsolutePath)
        }
        catch {
            # Uma requisição com problema (ex. navegador cancelou) não derruba o servidor.
            Write-Warning ('{0} {1}: {2}' -f $context.Request.HttpMethod, $context.Request.Url.AbsolutePath, $_.Exception.Message)
            try { $context.Response.Abort() } catch { }
        }
    }
}
finally {
    $listener.Stop()
    $listener.Close()
}
