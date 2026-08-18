#requires -Version 7.2
<#
.SYNOPSIS
Instala e configura o M365 License Assessment.
.DESCRIPTION
Instala o modulo Graph necessario, copia os componentes publicados, grava a
configuracao local e, opcionalmente, executa a primeira avaliacao interativa.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$InstallPath = "$env:ProgramData\M365LicenseAssessment",
    [string]$RepositoryRawUrl = 'https://raw.githubusercontent.com/JulioVicente/o365-telemetria-uso-licencas/main',
    [switch]$SkipModuleInstall
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$requiredFiles = @('Invoke-M365LicenseAssessment.ps1', 'config/license-catalog.pt-BR.json', 'README.md', 'QUICK_START.md')

function Write-Step([string]$Message) { Write-Host "`n==> $Message" -ForegroundColor Cyan }
function Read-Default([string]$Prompt, [string]$Default) {
    $suffix = if ($Default) { " [$Default]" } else { '' }
    $answer = Read-Host "$Prompt$suffix"
    if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
    return $answer
}
function Read-EmailAddress([string]$Prompt) {
    do {
        $value = Read-Default $Prompt ''
        $parsed = $null
        $valid = [System.Net.Mail.MailAddress]::TryCreate($value, [ref]$parsed) -and $parsed.Address -eq $value
        if (-not $valid) { Write-Warning 'Informe um endereco de email valido.' }
    } while (-not $valid)
    return $parsed.Address
}
function Test-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
function Assert-Environment {
    if ($env:OS -ne 'Windows_NT') { throw 'Este instalador e exclusivo para Windows.' }
    if (-not (Test-Administrator)) { throw 'Abra o PowerShell 7 como Administrador e execute novamente.' }
    if ($PSVersionTable.PSVersion -lt [version]'7.2') { throw 'PowerShell 7.2 ou superior e necessario.' }
}
function Ensure-GraphModule {
    Write-Step 'Validando Microsoft.Graph.Authentication'
    if (Get-Module -ListAvailable Microsoft.Graph.Authentication) { return }
    if ($SkipModuleInstall) { throw 'Microsoft.Graph.Authentication nao esta instalado.' }
    if ($PSCmdlet.ShouldProcess('Microsoft.Graph.Authentication', 'Instalar modulo para todos os usuarios')) {
        Install-Module Microsoft.Graph.Authentication -Scope AllUsers -Repository PSGallery -Force -AllowClobber
    }
}
function Ensure-7Zip {
    Write-Step 'Validando 7-Zip para o arquivo criptografado'
    $paths = @('C:\Program Files\7-Zip\7z.exe','C:\Program Files (x86)\7-Zip\7z.exe')
    if ($paths | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1) { return }
    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $winget) { throw '7-Zip nao encontrado e winget indisponivel. Instale o 7-Zip e execute novamente.' }
    if ($PSCmdlet.ShouldProcess('7-Zip', 'Instalar via winget para gerar ZIP AES-256')) {
        & $winget.Source install --id 7zip.7zip --exact --silent --accept-package-agreements --accept-source-agreements
        if ($LASTEXITCODE -ne 0) { throw "Falha ao instalar o 7-Zip via winget (codigo $LASTEXITCODE)." }
    }
}
function Copy-ProjectFiles([string]$Destination) {
    Write-Step 'Obtendo os componentes da solucao'
    foreach ($relativePath in $requiredFiles) {
        $relativeWindows = $relativePath -replace '/', '\'
        $target = Join-Path $Destination $relativeWindows
        New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
        $local = if ($PSScriptRoot) { Join-Path $PSScriptRoot $relativeWindows } else { $null }
        if ($local -and (Test-Path -LiteralPath $local -PathType Leaf)) {
            Copy-Item -LiteralPath $local -Destination $target -Force
        } else {
        $cacheKey = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
        $uri = "$($RepositoryRawUrl.TrimEnd('/'))/$relativePath`?nocache=$cacheKey"
            try { Invoke-WebRequest -UseBasicParsing -Uri $uri -OutFile $target }
            catch {
                Remove-Item -LiteralPath $target -Force -ErrorAction SilentlyContinue
                throw "Componente obrigatorio indisponivel: $uri. $($_.Exception.Message)"
            }
        }
    }
}

function Test-InstalledComponents([string]$Destination) {
    Write-Step 'Validando componentes e pacotes locais'
    foreach ($relativePath in $requiredFiles) {
        $path = Join-Path $Destination ($relativePath -replace '/', '\')
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Componente ausente apos instalacao: $relativePath" }
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    foreach ($command in 'Connect-MgGraph','Get-MgContext','Invoke-MgGraphRequest','Disconnect-MgGraph') {
        if (-not (Get-Command $command -ErrorAction SilentlyContinue)) { throw "Comando obrigatorio indisponivel: $command" }
    }
    $catalogPath = Join-Path $Destination 'config\license-catalog.pt-BR.json'
    $catalog = Get-Content -LiteralPath $catalogPath -Raw | ConvertFrom-Json
    if (-not $catalog.plans -or @($catalog.plans).Count -eq 0) { throw 'Catalogo de licencas vazio ou invalido.' }
    $probePath = Join-Path $Destination 'output\.write-test'
    Set-Content -LiteralPath $probePath -Value 'ok' -Encoding ascii
    Remove-Item -LiteralPath $probePath -Force
    Write-Host "Componentes locais validados; $(@($catalog.plans).Count) planos no catalogo." -ForegroundColor Green
}

$existed = Test-Path -LiteralPath $InstallPath
$rollback = Join-Path ([IO.Path]::GetTempPath()) ('m365-license-install-' + [guid]::NewGuid().ToString('N'))
try {
    Write-Host 'M365 License Assessment - Instalacao' -ForegroundColor Green
    Assert-Environment
    Ensure-GraphModule
    Ensure-7Zip
    if ($PSCmdlet.ShouldProcess($InstallPath, 'Criar diretorio de instalacao')) {
        if ($existed) {
            New-Item -ItemType Directory -Path $rollback -Force | Out-Null
            Copy-Item -LiteralPath $InstallPath -Destination (Join-Path $rollback 'previous') -Recurse -Force
        }
        foreach ($folder in '', 'config', 'output') { New-Item -ItemType Directory -Path (Join-Path $InstallPath $folder) -Force | Out-Null }
        Copy-ProjectFiles $InstallPath
        Test-InstalledComponents $InstallPath
    }

    Write-Step 'Configuracao interativa'
    Write-Host 'Informe o usuario que concedera as permissoes e recebera o relatorio.' -ForegroundColor Yellow
    $accountEmail = Read-EmailAddress 'Email do usuario'
    $configuration = [ordered]@{ SchemaVersion=1; AccountEmail=$accountEmail; BccAddress='suporte@bestsoft.com.br'; TelemetryPeriodDays=180 }
    $configPath = Join-Path $InstallPath 'config\settings.json'
    if ($PSCmdlet.ShouldProcess($configPath, 'Gravar configuracao')) {
        $configuration | ConvertTo-Json | Set-Content -LiteralPath $configPath -Encoding utf8
    }

    $runner = Join-Path $InstallPath 'Run-Assessment.ps1'
    $runnerContent = @'
#requires -Version 7.2
[CmdletBinding()] param()
$settings = Get-Content (Join-Path $PSScriptRoot 'config\settings.json') -Raw | ConvertFrom-Json
$arguments = @{ OutputPath=(Join-Path $PSScriptRoot 'output'); TelemetryPeriodDays=[int]$settings.TelemetryPeriodDays; ExpectedAccount=$settings.AccountEmail; EmailTo=$settings.AccountEmail; BccAddress=$settings.BccAddress; SendEmail=$true }
& (Join-Path $PSScriptRoot 'Invoke-M365LicenseAssessment.ps1') @arguments
'@
    if ($PSCmdlet.ShouldProcess($runner, 'Criar comando simplificado de execucao')) {
        Set-Content -LiteralPath $runner -Value $runnerContent -Encoding utf8
    }

    Write-Host "`nInstalacao concluida em: $InstallPath" -ForegroundColor Green
    Write-Host "Executar: pwsh -NoProfile -File `"$runner`""
    if (-not $WhatIfPreference) {
        Write-Step 'Validando login, coleta e envio do relatorio'
        Write-Host "Autentique-se como $accountEmail. A instalacao so termina se a coleta e o envio funcionarem." -ForegroundColor Yellow
        & $runner
        Write-Host 'Validacao concluida: coleta executada e email aceito pelo Microsoft Graph.' -ForegroundColor Green
    }
} catch {
    $message = $_.Exception.Message
    Write-Warning 'Falha detectada; revertendo a instalacao local.'
    if ($existed -and (Test-Path (Join-Path $rollback 'previous'))) {
        Remove-Item -LiteralPath $InstallPath -Recurse -Force -ErrorAction SilentlyContinue
        Copy-Item -LiteralPath (Join-Path $rollback 'previous') -Destination $InstallPath -Recurse -Force
    } elseif (-not $existed -and (Test-Path $InstallPath)) {
        Remove-Item -LiteralPath $InstallPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    throw "Instalacao interrompida e alteracoes locais revertidas: $message"
} finally {
    Remove-Item -LiteralPath $rollback -Recurse -Force -ErrorAction SilentlyContinue
}
