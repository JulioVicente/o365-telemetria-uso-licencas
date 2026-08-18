#requires -Version 7.2
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'output'),
    [string]$PriceCatalogPath = (Join-Path $PSScriptRoot 'config\license-catalog.pt-BR.json'),
    [ValidateSet(30, 90, 180)] [int]$TelemetryPeriodDays = 180,
    [string]$EmailTo,
    [string]$BccAddress = 'suporte@bestsoft.com.br',
    [string]$ExpectedAccount,
    [switch]$SendEmail = $true,
    [string]$ArchivePassword = 'bestsoft',
    [switch]$IncludeGuests
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$solutionVersion = '1.4.0'

function Write-ExecutionStatus {
    param([int]$Percent, [string]$Message)
    Write-Progress -Activity 'Avaliacao de licencas Microsoft 365' -Status $Message -PercentComplete $Percent
    Write-Host ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message) -ForegroundColor Cyan
}

if (-not ('M365LicenseAssessment.ConsoleSpinner' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Threading;
namespace M365LicenseAssessment {
    public sealed class ConsoleSpinner : IDisposable {
        private readonly string text;
        private readonly DateTime started = DateTime.UtcNow;
        private readonly Timer timer;
        private readonly char[] frames = new[] { '|', '/', '-', '\\' };
        private int index;
        private readonly bool enabled;
        public ConsoleSpinner(string text) {
            this.text = text;
            enabled = !Console.IsOutputRedirected;
            if (enabled) timer = new Timer(Tick, null, 0, 120);
        }
        private void Tick(object state) {
            try {
                var elapsed = DateTime.UtcNow - started;
                Console.Write("\r  {0} {1} ({2:mm\\:ss})", frames[index++ % frames.Length], text, elapsed);
            } catch { }
        }
        public void Dispose() {
            if (!enabled) return;
            timer.Dispose();
            try { Console.Write("\r" + new string(' ', Math.Min(Console.BufferWidth - 1, text.Length + 24)) + "\r"); } catch { }
        }
    }
}
'@
}

function Invoke-WithSpinner {
    param([Parameter(Mandatory)][string]$Message, [Parameter(Mandatory)][scriptblock]$Operation)
    $spinner = [M365LicenseAssessment.ConsoleSpinner]::new($Message)
    try { & $Operation } finally { $spinner.Dispose() }
}

function Invoke-GraphRequestWithRetry {
    param(
        [Parameter(Mandatory)][ValidateSet('GET','POST')][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [object]$Body,
        [string]$OutputFilePath,
        [string]$ContentType,
        [int]$MaxAttempts = 4
    )
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $parameters = @{ Method=$Method; Uri=$Uri }
            if ($PSBoundParameters.ContainsKey('Body')) { $parameters.Body = $Body }
            if ($OutputFilePath) { $parameters.OutputFilePath = $OutputFilePath }
            else { $parameters.OutputType = 'PSObject' }
            if ($ContentType) { $parameters.ContentType = $ContentType }
            return Invoke-MgGraphRequest @parameters
        } catch {
            $message = $_.Exception.Message
            $isTransient = $message -match '(429|TooManyRequests|throttl|500|502|503|504|InternalServerError|BadGateway|ServiceUnavailable|GatewayTimeout|temporar)'
            if (-not $isTransient -or $attempt -eq $MaxAttempts) { throw }
            $delay = [math]::Min(30, [math]::Pow(2, $attempt))
            Write-Warning "Falha transitoria do Microsoft Graph (tentativa $attempt/$MaxAttempts). Nova tentativa em $delay segundos."
            Start-Sleep -Seconds $delay
        }
    }
}

function Import-RequiredModule {
    param([Parameter(Mandatory)] [string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Modulo '$Name' nao encontrado. Execute: Install-Module $Name -Scope CurrentUser"
    }
    Import-Module $Name -ErrorAction Stop
}

function Get-GraphCollection {
    param([Parameter(Mandatory)] [string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    while ($Uri) {
        $response = Invoke-GraphRequestWithRetry -Method GET -Uri $Uri
        if ($null -ne $response.value) { foreach ($item in $response.value) { $items.Add($item) } }
        else { $items.Add($response); break }
        $nextLinkProperty = $response.PSObject.Properties['@odata.nextLink']
        $Uri = if ($nextLinkProperty) { [string]$nextLinkProperty.Value } else { $null }
    }
    return $items.ToArray()
}

function Get-ReportCsv {
    param([Parameter(Mandatory)] [string]$ReportName, [Parameter(Mandatory)] [int]$PeriodDays)
    $tempFile = Join-Path ([IO.Path]::GetTempPath()) ("m365-{0}-{1}.csv" -f $ReportName, [guid]::NewGuid())
    try {
        $uri = "https://graph.microsoft.com/v1.0/reports/$ReportName(period='D$PeriodDays')"
        Invoke-GraphRequestWithRetry -Method GET -Uri $uri -OutputFilePath $tempFile | Out-Null
        if ((Test-Path $tempFile) -and (Get-Item $tempFile).Length -gt 0) {
            return @(Import-Csv -LiteralPath $tempFile)
        }
        return @()
    }
    finally { Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue }
}

function Get-CopilotReportCsv {
    param([Parameter(Mandatory)] [int]$PeriodDays)
    $tempFile = Join-Path ([IO.Path]::GetTempPath()) ("m365-copilot-{0}.csv" -f [guid]::NewGuid())
    try {
        $uri = "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='D$PeriodDays',version='v1')"
        Invoke-GraphRequestWithRetry -Method GET -Uri $uri -OutputFilePath $tempFile | Out-Null
        if ((Test-Path $tempFile) -and (Get-Item $tempFile).Length -gt 0) { return @(Import-Csv -LiteralPath $tempFile) }
        return @()
    }
    finally { Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue }
}

function Find-PropertyValue {
    param([object]$Object, [string[]]$Names)
    if ($null -eq $Object) { return $null }
    foreach ($name in $Names) {
        $property = $Object.PSObject.Properties | Where-Object Name -EQ $name | Select-Object -First 1
        if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) { return $property.Value }
    }
    return $null
}

function Convert-ToDateOrNull {
    param([object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) { return $parsed.ToUniversalTime() }
    return $null
}

function Get-DaysSince {
    param([Nullable[datetime]]$Date, [datetime]$Now)
    if ($null -eq $Date) { return $null }
    return [math]::Floor(($Now - [datetime]$Date).TotalDays)
}

function Get-UsageState {
    param([Nullable[int]]$Days)
    if ($null -eq $Days) { return 'Sem atividade observada' }
    if ([int]$Days -gt 90) { return 'Inativo >90 dias' }
    if ([int]$Days -gt 30) { return 'Inativo >30 dias' }
    return 'Ativo <=30 dias'
}

function Get-DecimalSum {
    param([object[]]$InputObject, [Parameter(Mandatory)][string]$Property)
    [decimal]$total = 0
    foreach ($item in @($InputObject)) {
        if ($null -eq $item) { continue }
        $valueProperty = $item.PSObject.Properties[$Property]
        if ($valueProperty -and $null -ne $valueProperty.Value) { $total += [decimal]$valueProperty.Value }
    }
    return $total
}

function Test-MatchesAnyPattern {
    param([string]$Value, [string[]]$Patterns)
    foreach ($pattern in @($Patterns)) { if ($Value -like $pattern) { return $true } }
    return $false
}

function Get-7ZipPath {
    $command = Get-Command 7z.exe -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    foreach ($path in 'C:\Program Files\7-Zip\7z.exe','C:\Program Files (x86)\7-Zip\7z.exe') {
        if (Test-Path -LiteralPath $path -PathType Leaf) { return $path }
    }
    throw '7-Zip nao encontrado. Execute novamente o instalador para habilitar o ZIP criptografado.'
}

function Get-MinimumPlan {
    param([hashtable]$Needs, [object[]]$Catalog)
    if (-not ($Needs.Email -or $Needs.OneDrive -or $Needs.SharePoint -or $Needs.OfficeWeb -or $Needs.OfficeDesktop)) {
        return $null
    }
    $eligible = $Catalog | Where-Object {
        $_.recommendable -and
        (-not $Needs.Email -or $_.features.email) -and
        (-not $Needs.OneDrive -or $_.features.oneDrive) -and
        (-not $Needs.SharePoint -or $_.features.sharePoint) -and
        (-not $Needs.OfficeWeb -or $_.features.officeWeb) -and
        (-not $Needs.OfficeDesktop -or $_.features.officeDesktop)
    } | Sort-Object monthlyPriceBRL
    return $eligible | Select-Object -First 1
}

function Test-UserServiceEntitlement {
    param([object[]]$AssignedLicenses, [hashtable]$SkuIndex, [string[]]$ServicePlanPatterns)
    foreach ($assignedLicense in @($AssignedLicenses)) {
        $assignedSku = $SkuIndex[[string]$assignedLicense.skuId]
        if (-not $assignedSku) { continue }
        $disabledPlanIds = @($assignedLicense.disabledPlans | ForEach-Object { [string]$_ })
        foreach ($servicePlan in @($assignedSku.servicePlans)) {
            $matches = $false
            foreach ($pattern in $ServicePlanPatterns) {
                if ($servicePlan.servicePlanName -like $pattern) { $matches = $true; break }
            }
            if ($matches -and $servicePlan.provisioningStatus -ne 'Disabled' -and [string]$servicePlan.servicePlanId -notin $disabledPlanIds) { return $true }
        }
    }
    return $false
}

Import-RequiredModule Microsoft.Graph.Authentication
if (-not (Test-Path -LiteralPath $PriceCatalogPath)) { throw "Catalogo nao encontrado: $PriceCatalogPath" }
$catalogData = Get-Content -LiteralPath $PriceCatalogPath -Raw | ConvertFrom-Json
$catalog = @($catalogData.plans)
$nonCommercialSkuPatterns = @($catalogData.nonCommercialSkuPatterns)

$scopes = @('User.Read.All', 'AuditLog.Read.All', 'LicenseAssignment.Read.All', 'Organization.Read.All', 'Reports.Read.All')
if ($SendEmail) { $scopes += 'Mail.Send' }
Write-ExecutionStatus 5 'Aguardando autenticacao no Microsoft 365...'
Write-Host "M365 License Assessment v$solutionVersion" -ForegroundColor Green
Connect-MgGraph -Scopes $scopes -NoWelcome

$context = Get-MgContext
if ($ExpectedAccount -and $context.Account -ine $ExpectedAccount) {
    Disconnect-MgGraph | Out-Null
    throw "A conta autenticada '$($context.Account)' nao corresponde ao usuario informado '$ExpectedAccount'. Execute novamente e entre com o usuario informado na instalacao."
}
if ($SendEmail) {
    if ($EmailTo -and $EmailTo -ine $context.Account) {
        Write-Warning "O destinatario '$EmailTo' foi substituido pela conta autenticada '$($context.Account)'."
    }
    $EmailTo = $context.Account
}
Write-ExecutionStatus 8 'Autenticacao concluida. Executando validacao minima das APIs...'
try {
    Invoke-WithSpinner 'Validando identidade e perfil autenticado' {
        Invoke-GraphRequestWithRetry -Method GET -Uri 'https://graph.microsoft.com/v1.0/me?$select=id,userPrincipalName,mail' | Out-Null
    }
    Invoke-WithSpinner 'Validando leitura de usuarios e atividade de login' {
        Invoke-GraphRequestWithRetry -Method GET -Uri 'https://graph.microsoft.com/v1.0/users?$top=1&$select=id,signInActivity' | Out-Null
    }
    Invoke-WithSpinner 'Validando leitura de licencas e SKUs' {
        Invoke-GraphRequestWithRetry -Method GET -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus?$select=skuId,skuPartNumber' | Out-Null
    }
    Invoke-WithSpinner 'Validando dados de identificacao do tenant' {
        Invoke-GraphRequestWithRetry -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id,displayName,verifiedDomains,tenantType' | Out-Null
    }
    foreach ($probe in @(
        @{Name='atividade geral';Api='getOffice365ActiveUserDetail'},
        @{Name='aplicativos Office';Api='getM365AppUserDetail'},
        @{Name='email';Api='getEmailActivityUserDetail'},
        @{Name='OneDrive';Api='getOneDriveActivityUserDetail'},
        @{Name='SharePoint';Api='getSharePointActivityUserDetail'},
        @{Name='Teams';Api='getTeamsUserActivityUserDetail'}
    )) {
        Invoke-WithSpinner "Validando relatorio de $($probe.Name)" { Get-ReportCsv $probe.Api 7 | Out-Null }
    }
    if ($SendEmail) {
        $validationMessage = @{message=@{subject='Validacao tecnica - M365 License Assessment';body=@{contentType='Text';content='A autenticacao, a caixa Exchange Online e a permissao Mail.Send foram validadas. A analise completa sera iniciada em seguida.'};toRecipients=@(@{emailAddress=@{address=$EmailTo}})};saveToSentItems=$true} | ConvertTo-Json -Depth 8 -Compress
        Invoke-WithSpinner 'Validando caixa Exchange Online e permissao de envio' {
            Invoke-GraphRequestWithRetry -Method POST -Uri 'https://graph.microsoft.com/v1.0/me/sendMail' -Body $validationMessage -ContentType 'application/json' | Out-Null
        }
    }
} catch {
    throw "Pre-validacao interrompida; nenhuma analise foi iniciada. API ou permissao indisponivel: $($_.Exception.Message)"
}
Write-Host 'Pre-validacao concluida com sucesso em todas as APIs.' -ForegroundColor Green
Write-ExecutionStatus 10 'Iniciando coleta completa; esta etapa pode levar alguns minutos.'
$now = [datetime]::UtcNow
$meProfile = Invoke-GraphRequestWithRetry -Method GET -Uri 'https://graph.microsoft.com/v1.0/me?$select=id,displayName,userPrincipalName,mail'
$organization = @(Get-GraphCollection 'https://graph.microsoft.com/v1.0/organization?$select=id,displayName,verifiedDomains,tenantType') | Select-Object -First 1
$tenantName = Find-PropertyValue $organization @('displayName')
if (-not $tenantName) { $tenantName = $context.TenantId }
$verifiedDomainsProperty = $organization.PSObject.Properties['verifiedDomains']
$verifiedDomains = if ($verifiedDomainsProperty) { @($verifiedDomainsProperty.Value) } else { @() }
$defaultDomainObject = $verifiedDomains | Where-Object isDefault | Select-Object -First 1
if (-not $defaultDomainObject) { $defaultDomainObject = $verifiedDomains | Select-Object -First 1 }
$defaultDomain = if ($defaultDomainObject) { Find-PropertyValue $defaultDomainObject @('name') } else { 'Nao informado' }
$tenantType = Find-PropertyValue $organization @('tenantType')
if (-not $tenantType) { $tenantType = 'Nao informado' }
$operatorName = Find-PropertyValue $meProfile @('displayName')
if (-not $operatorName) { $operatorName = $context.Account }
$runId = $now.ToString('yyyyMMdd-HHmmss')
$runPath = Join-Path $OutputPath $runId
New-Item -ItemType Directory -Path $runPath -Force | Out-Null

$usersUri = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName,userType,accountEnabled,assignedLicenses,signInActivity&`$top=999"
Write-ExecutionStatus 18 'Coletando usuarios e atividade de login...'
$users = @(Invoke-WithSpinner 'Consultando diretorio' { Get-GraphCollection $usersUri })
if (-not $IncludeGuests) { $users = @($users | Where-Object userType -EQ 'Member') }
Write-ExecutionStatus 25 'Coletando assinaturas e SKUs do tenant...'
$skus = @(Invoke-WithSpinner 'Consultando assinaturas' { Get-GraphCollection 'https://graph.microsoft.com/v1.0/subscribedSkus' })
$companySubscriptionsAvailable = $true
try {
    $companySubscriptions = @(Invoke-WithSpinner 'Consultando ciclos das assinaturas' { Get-GraphCollection 'https://graph.microsoft.com/v1.0/directory/subscriptions' })
} catch {
    $companySubscriptionsAvailable = $false; $companySubscriptions = @()
    Write-Warning "Nao foi possivel consultar datas das assinaturas: $($_.Exception.Message)"
}

Write-ExecutionStatus 33 'Coletando atividade geral do Microsoft 365...'
$activeUsers = @(Invoke-WithSpinner 'Baixando atividade geral' { Get-ReportCsv 'getOffice365ActiveUserDetail' $TelemetryPeriodDays })
Write-ExecutionStatus 41 'Coletando atividade dos aplicativos Office...'
$appUsers = @(Invoke-WithSpinner 'Baixando atividade do Office' { Get-ReportCsv 'getM365AppUserDetail' $TelemetryPeriodDays })
Write-ExecutionStatus 49 'Coletando atividade de email...'
$emailUsers = @(Invoke-WithSpinner 'Baixando atividade de email' { Get-ReportCsv 'getEmailActivityUserDetail' $TelemetryPeriodDays })
Write-ExecutionStatus 57 'Coletando atividade do OneDrive...'
$oneDriveUsers = @(Invoke-WithSpinner 'Baixando atividade do OneDrive' { Get-ReportCsv 'getOneDriveActivityUserDetail' $TelemetryPeriodDays })
Write-ExecutionStatus 65 'Coletando atividade do SharePoint...'
$sharePointUsers = @(Invoke-WithSpinner 'Baixando atividade do SharePoint' { Get-ReportCsv 'getSharePointActivityUserDetail' $TelemetryPeriodDays })
Write-ExecutionStatus 69 'Coletando atividade do Microsoft Teams...'
$teamsUsers = @(Invoke-WithSpinner 'Baixando atividade do Teams' { Get-ReportCsv 'getTeamsUserActivityUserDetail' $TelemetryPeriodDays })
Write-ExecutionStatus 70 'Coletando uso do Microsoft 365 Copilot...'
$copilotReportAvailable = $true
try { $copilotUsers = @(Invoke-WithSpinner 'Baixando atividade do Copilot' { Get-CopilotReportCsv $TelemetryPeriodDays }) }
catch {
    $copilotReportAvailable = $false; $copilotUsers = @()
    Write-Warning "Nao foi possivel coletar o relatorio de uso do Copilot: $($_.Exception.Message)"
}

function New-ReportIndex { param([object[]]$Rows)
    $index = @{}
    foreach ($row in $Rows) {
        $upn = Find-PropertyValue $row @('User Principal Name', 'Owner Principal Name')
        if ($upn) { $index[[string]$upn.ToLowerInvariant()] = $row }
    }
    return $index
}
$activeIndex = New-ReportIndex $activeUsers
$appIndex = New-ReportIndex $appUsers
$emailIndex = New-ReportIndex $emailUsers
$oneDriveIndex = New-ReportIndex $oneDriveUsers
$sharePointIndex = New-ReportIndex $sharePointUsers
$teamsIndex = New-ReportIndex $teamsUsers
$copilotIndex = New-ReportIndex $copilotUsers
$skuById = @{}; foreach ($sku in $skus) { $skuById[[string]$sku.skuId] = $sku }
$priceBySku = @{}; foreach ($plan in $catalog) { foreach ($part in $plan.skuPartNumbers) { $priceBySku[$part] = $plan } }

$subscriptionRows = if ($companySubscriptionsAvailable) {
    foreach ($subscription in $companySubscriptions) {
        $sku = $skuById[[string]$subscription.skuId]
        $purchased = [int](Find-PropertyValue $subscription @('totalLicenses'))
        $used = if ($sku) { [int](Find-PropertyValue $sku @('consumedUnits')) } else { 0 }
        $enabled = if ($sku -and $sku.prepaidUnits) { [int](Find-PropertyValue $sku.prepaidUnits @('enabled')) } else { $purchased }
        [pscustomobject]@{
            Plano = Find-PropertyValue $subscription @('skuPartNumber')
            IDAssinatura = Find-PropertyValue $subscription @('commerceSubscriptionId')
            Status = Find-PropertyValue $subscription @('status')
            Avaliacao = [bool](Find-PropertyValue $subscription @('isTrial'))
            DataInicio = Convert-ToDateOrNull (Find-PropertyValue $subscription @('createdDateTime'))
            DataProximoCiclo = Convert-ToDateOrNull (Find-PropertyValue $subscription @('nextLifecycleDateTime'))
            QuantidadeAssinatura = $purchased
            QuantidadeHabilitadaNoSku = $enabled
            QuantidadeUsadaNoSku = $used
            QuantidadeDisponivelNoSku = [math]::Max(0, $enabled - $used)
            TermoFaturamento = 'Nao informado pelo Microsoft Graph; confirmar no Partner Center'
            RestricaoReducao = 'Planos anuais ou anuais pagos mensalmente podem nao permitir reducao durante a vigencia; validar na renovacao'
        }
    }
} else {
    foreach ($sku in $skus) {
        $enabled = if ($sku.prepaidUnits) { [int](Find-PropertyValue $sku.prepaidUnits @('enabled')) } else { 0 }
        $used = [int](Find-PropertyValue $sku @('consumedUnits'))
        [pscustomobject]@{
            Plano=$sku.skuPartNumber;IDAssinatura='Nao disponivel';Status=$sku.capabilityStatus;Avaliacao=$null
            DataInicio=$null;DataProximoCiclo=$null;QuantidadeAssinatura=$enabled;QuantidadeHabilitadaNoSku=$enabled
            QuantidadeUsadaNoSku=$used;QuantidadeDisponivelNoSku=[math]::Max(0,$enabled-$used)
            TermoFaturamento='Nao informado pelo Microsoft Graph; confirmar no Partner Center'
            RestricaoReducao='Planos anuais ou anuais pagos mensalmente podem nao permitir reducao durante a vigencia; validar na renovacao'
        }
    }
}
$lifecycleBySkuId = @{}
foreach ($subscription in $companySubscriptions) {
    $lifecycleDate = Convert-ToDateOrNull (Find-PropertyValue $subscription @('nextLifecycleDateTime'))
    if ($null -eq $lifecycleDate) { continue }
    $skuKey = [string]$subscription.skuId
    if (-not $lifecycleBySkuId.ContainsKey($skuKey)) { $lifecycleBySkuId[$skuKey] = [System.Collections.Generic.List[string]]::new() }
    $dateText = $lifecycleDate.ToString('dd/MM/yyyy')
    if (-not $lifecycleBySkuId[$skuKey].Contains($dateText)) { $lifecycleBySkuId[$skuKey].Add($dateText) }
}

Write-ExecutionStatus 72 'Consolidando telemetria e avaliando licencas...'
$results = foreach ($user in $users) {
    $key = [string]$user.userPrincipalName.ToLowerInvariant()
    $active = $activeIndex[$key]; $apps = $appIndex[$key]; $email = $emailIndex[$key]
    $drive = $oneDriveIndex[$key]; $site = $sharePointIndex[$key]; $teams = $teamsIndex[$key]; $copilot = $copilotIndex[$key]
    $signInActivityProperty = $user.PSObject.Properties['signInActivity']
    $signInActivity = if ($signInActivityProperty) { $signInActivityProperty.Value } else { $null }
    $signInDate = Convert-ToDateOrNull (Find-PropertyValue $signInActivity @('lastSuccessfulSignInDateTime'))
    if ($null -eq $signInDate) { $signInDate = Convert-ToDateOrNull (Find-PropertyValue $signInActivity @('lastSignInDateTime')) }
    $emailDate = Convert-ToDateOrNull (Find-PropertyValue $email @('Last Activity Date'))
    $driveDate = Convert-ToDateOrNull (Find-PropertyValue $drive @('Last Activity Date'))
    $siteDate = Convert-ToDateOrNull (Find-PropertyValue $site @('Last Activity Date'))
    $desktopDate = Convert-ToDateOrNull (Find-PropertyValue $apps @('Last Activity Date', 'Last Activated Date'))
    $webDate = Convert-ToDateOrNull (Find-PropertyValue $active @('Office 365 Last Activity Date', 'Microsoft 365 Apps Last Activity Date'))
    $teamsDate = Convert-ToDateOrNull (Find-PropertyValue $teams @('Last Activity Date'))
    $copilotDate = Convert-ToDateOrNull (Find-PropertyValue $copilot @('Last Activity Date'))
    $signInDays = Get-DaysSince $signInDate $now; $emailDays = Get-DaysSince $emailDate $now
    $driveDays = Get-DaysSince $driveDate $now; $siteDays = Get-DaysSince $siteDate $now
    $desktopDays = Get-DaysSince $desktopDate $now; $webDays = Get-DaysSince $webDate $now
    $teamsDays = Get-DaysSince $teamsDate $now
    $copilotDays = Get-DaysSince $copilotDate $now

    $assignedParts = @($user.assignedLicenses | ForEach-Object { $skuById[[string]$_.skuId].skuPartNumber } | Where-Object { $_ })
    $userLifecycleDates = @($user.assignedLicenses | ForEach-Object { $lifecycleBySkuId[[string]$_.skuId] } | ForEach-Object { $_ } | Sort-Object -Unique)
    $knownCurrentPlans = @($assignedParts | ForEach-Object { $priceBySku[$_] } | Where-Object { $_ } | Sort-Object name -Unique)
    $currentPrice = Get-DecimalSum $knownCurrentPlans 'monthlyPriceBRL'
    $unpriced = @($assignedParts | Where-Object {
        -not $priceBySku.ContainsKey($_) -and -not (Test-MatchesAnyPattern $_ $nonCommercialSkuPatterns)
    })
    $needs = @{
        Email = ($null -ne $emailDays -and $emailDays -le 90)
        OneDrive = ($null -ne $driveDays -and $driveDays -le 90)
        SharePoint = ($null -ne $siteDays -and $siteDays -le 90)
        OfficeWeb = ($null -ne $webDays -and $webDays -le 90)
        OfficeDesktop = ($null -ne $desktopDays -and $desktopDays -le 90)
    }
    $minimum = Get-MinimumPlan $needs $catalog
    $hasExchangeLicense = Test-UserServiceEntitlement @($user.assignedLicenses) $skuById @('EXCHANGE*')
    $hasTeamsLicense = Test-UserServiceEntitlement @($user.assignedLicenses) $skuById @('TEAMS*')
    $hasSharePointLicense = Test-UserServiceEntitlement @($user.assignedLicenses) $skuById @('SHAREPOINT*')
    $hasOneDriveLicense = Test-UserServiceEntitlement @($user.assignedLicenses) $skuById @('ONEDRIVE*','SHAREPOINT*')
    $hasOfficeDesktopLicense = Test-UserServiceEntitlement @($user.assignedLicenses) $skuById @('OFFICESUBSCRIPTION*','OFFICE_BUSINESS*','OFFICE_PRO_PLUS*')
    $hasOfficeWebLicense = Test-UserServiceEntitlement @($user.assignedLicenses) $skuById @('SHAREPOINTWAC*','WACONEDRIVE*','OFFICE_WEB*')
    $hasCopilotLicense = ($null -ne $copilot) -or
        (@($assignedParts | Where-Object { $_ -like 'Microsoft_365_Copilot*' }).Count -gt 0) -or
        (Test-UserServiceEntitlement @($user.assignedLicenses) $skuById @('M365_COPILOT_APPS','M365_COPILOT_BUSINESS_CHAT'))
    $teamsReviewCandidate = $hasTeamsLicense -and ($null -eq $teamsDays -or $teamsDays -gt 30)
    $inactiveForSharedMailbox = (-not [bool]$user.accountEnabled) -or ($null -ne $signInDays -and $signInDays -gt 90)
    $sharedMailboxCandidate = $hasExchangeLicense -and $inactiveForSharedMailbox
    $recommendation = if ($assignedParts.Count -eq 0) { 'Sem licenca atribuida' }
        elseif ($sharedMailboxCandidate) { 'Candidato a caixa compartilhada; validar tamanho, arquivo, retencao/hold, acesso delegado e bloquear login antes de remover licencas' }
        elseif ($null -eq $minimum) { 'Candidato a remocao completa; validar funcao, retencao, caixa compartilhada e requisitos de seguranca' }
        else { "Avaliar $($minimum.name)" }
    $retainedAddOns = @($knownCurrentPlans | Where-Object {
        $retainAlways = Find-PropertyValue $_ @('retainOnRecommendation')
        $retainWhenUsed = Find-PropertyValue $_ @('retainWhenUsed')
        $retainAlways -or ($retainWhenUsed -eq 'Copilot' -and $null -ne $copilotDate)
    })
    $retainedAddOnPrice = Get-DecimalSum $retainedAddOns 'monthlyPriceBRL'
    $recommendedPrice = if ($sharedMailboxCandidate) { [decimal]0 }
        elseif ($minimum) { [decimal]$minimum.monthlyPriceBRL + $retainedAddOnPrice }
        else { [decimal]0 }
    $saving = if ($unpriced.Count -eq 0) { [math]::Max(0, [decimal]$currentPrice - $recommendedPrice) } else { $null }

    [pscustomobject]@{
        Nome = $user.displayName; UPN = $user.userPrincipalName; ContaHabilitada = $user.accountEnabled
        LicencasAtuais = ($assignedParts -join '; '); LicencasSemPrecoPublico = ($unpriced -join '; ')
        ProximosCiclosLicencas = ($userLifecycleDates -join '; ')
        PrecoAtualMensalBRL = if ($unpriced.Count -eq 0) { [decimal]$currentPrice } else { $null }
        UltimoLogin = $signInDate; DiasSemLogin = $signInDays; StatusLogin = Get-UsageState $signInDays
        EmailLicenciado = $hasExchangeLicense; UltimoEmail = $emailDate; DiasSemEmail = $emailDays; StatusEmail = Get-UsageState $emailDays; CandidatoRevisaoEmail = ($hasExchangeLicense -and ($null -eq $emailDays -or $emailDays -gt 30))
        OneDriveLicenciado = $hasOneDriveLicense; UltimoOneDrive = $driveDate; DiasSemOneDrive = $driveDays; StatusOneDrive = Get-UsageState $driveDays; CandidatoRevisaoOneDrive = ($hasOneDriveLicense -and ($null -eq $driveDays -or $driveDays -gt 30))
        SharePointLicenciado = $hasSharePointLicense; UltimoSharePoint = $siteDate; DiasSemSharePoint = $siteDays; StatusSharePoint = Get-UsageState $siteDays; CandidatoRevisaoSharePoint = ($hasSharePointLicense -and ($null -eq $siteDays -or $siteDays -gt 30))
        OfficeDesktopLicenciado = $hasOfficeDesktopLicense; UltimoOfficeDesktop = $desktopDate; DiasSemOfficeDesktop = $desktopDays; StatusOfficeDesktop = Get-UsageState $desktopDays; CandidatoRevisaoOfficeDesktop = ($hasOfficeDesktopLicense -and ($null -eq $desktopDays -or $desktopDays -gt 30))
        OfficeWebLicenciado = $hasOfficeWebLicense; UltimoOfficeWeb = $webDate; DiasSemOfficeWeb = $webDays; StatusOfficeWeb = Get-UsageState $webDays; CandidatoRevisaoOfficeWeb = ($hasOfficeWebLicense -and ($null -eq $webDays -or $webDays -gt 30))
        TeamsLicenciado = $hasTeamsLicense; UltimoTeams = $teamsDate; DiasSemTeams = $teamsDays; StatusTeams = Get-UsageState $teamsDays
        CandidatoRevisaoTeams = $teamsReviewCandidate
        CopilotLicenciado = $hasCopilotLicense; CopilotUtilizado = ($null -ne $copilotDate)
        UltimoCopilot = $copilotDate; DiasSemCopilot = $copilotDays
        StatusCopilot = if (-not $copilotReportAvailable) { 'Coleta indisponivel' } else { Get-UsageState $copilotDays }
        CandidatoRevisaoCopilot = ($copilotReportAvailable -and $hasCopilotLicense -and ($null -eq $copilotDays -or $copilotDays -gt 30))
        ExchangeDetectado = $hasExchangeLicense
        CandidatoCaixaCompartilhada = $sharedMailboxCandidate
        Recomendacao = $recommendation; PlanoMinimoSugerido = if ($sharedMailboxCandidate) { 'Caixa compartilhada (sem licenca, se elegivel)' } elseif ($minimum) { $minimum.name } else { $null }
        PrecoSugeridoMensalBRL = $recommendedPrice; EconomiaMensalEstimadaBRL = $saving
    }
}

$csvPath = Join-Path $runPath 'analise-usuarios.csv'
Write-ExecutionStatus 82 'Gerando relatorio executivo e anexos...'
$results | Sort-Object UPN | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding utf8BOM
$subscriptionsPath = Join-Path $runPath 'planos-assinaturas.csv'
$subscriptionRows | Sort-Object Plano,DataProximoCiclo | Export-Csv -LiteralPath $subscriptionsPath -NoTypeInformation -Encoding utf8BOM
$summary = [pscustomobject]@{
    generatedAtUtc = $now.ToString('o'); solutionVersion = $solutionVersion
    tenantId = $context.TenantId; tenantName = $tenantName; defaultDomain = $defaultDomain
    generatedByName = $operatorName; generatedByAccount = $context.Account
    telemetryPeriodDays = $TelemetryPeriodDays; usersAnalyzed = @($results).Count
    licensedUsers = @($results | Where-Object LicencasAtuais).Count
    removalCandidates = @($results | Where-Object { $_.Recomendacao -like 'Candidato a remocao*' }).Count
    sharedMailboxCandidates = @($results | Where-Object CandidatoCaixaCompartilhada).Count
    copilotReportAvailable = $copilotReportAvailable
    copilotLicensedUsers = @($results | Where-Object CopilotLicenciado).Count
    copilotActiveUsers = @($results | Where-Object CopilotUtilizado).Count
    usersWithUnpricedLicenses = @($results | Where-Object LicencasSemPrecoPublico).Count
    estimatedMonthlySavingsBRL = Get-DecimalSum @($results) 'EconomiaMensalEstimadaBRL'
    priceCatalogAsOf = $catalogData.asOf; priceSource = $catalogData.source
    companySubscriptionsAvailable = $companySubscriptionsAvailable
    subscriptionsAnalyzed = @($subscriptionRows).Count
}
$summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $runPath 'resumo.json') -Encoding utf8

function Get-InactivityCount([string]$DaysProperty, [int]$Threshold) {
    return @($results | Where-Object { $null -ne $_.$DaysProperty -and $_.$DaysProperty -gt $Threshold }).Count
}
function Get-LicensedInactivityCount([string]$DaysProperty, [string]$LicensedProperty, [int]$Threshold) {
    return @($results | Where-Object { $_.$LicensedProperty -and ($null -eq $_.$DaysProperty -or [int]$_.$DaysProperty -gt $Threshold) }).Count
}
$inactivity = [ordered]@{
    Login30 = Get-InactivityCount 'DiasSemLogin' 30; Login90 = Get-InactivityCount 'DiasSemLogin' 90
    Email30 = Get-LicensedInactivityCount 'DiasSemEmail' 'EmailLicenciado' 30; Email90 = Get-LicensedInactivityCount 'DiasSemEmail' 'EmailLicenciado' 90
    OneDrive30 = Get-LicensedInactivityCount 'DiasSemOneDrive' 'OneDriveLicenciado' 30; OneDrive90 = Get-LicensedInactivityCount 'DiasSemOneDrive' 'OneDriveLicenciado' 90
    SharePoint30 = Get-LicensedInactivityCount 'DiasSemSharePoint' 'SharePointLicenciado' 30; SharePoint90 = Get-LicensedInactivityCount 'DiasSemSharePoint' 'SharePointLicenciado' 90
    OfficeDesktop30 = Get-LicensedInactivityCount 'DiasSemOfficeDesktop' 'OfficeDesktopLicenciado' 30; OfficeDesktop90 = Get-LicensedInactivityCount 'DiasSemOfficeDesktop' 'OfficeDesktopLicenciado' 90
    OfficeWeb30 = Get-LicensedInactivityCount 'DiasSemOfficeWeb' 'OfficeWebLicenciado' 30; OfficeWeb90 = Get-LicensedInactivityCount 'DiasSemOfficeWeb' 'OfficeWebLicenciado' 90
    Teams30 = Get-LicensedInactivityCount 'DiasSemTeams' 'TeamsLicenciado' 30; Teams90 = Get-LicensedInactivityCount 'DiasSemTeams' 'TeamsLicenciado' 90
    Copilot30 = if ($copilotReportAvailable) { Get-LicensedInactivityCount 'DiasSemCopilot' 'CopilotLicenciado' 30 } else { $null }
    Copilot90 = if ($copilotReportAvailable) { Get-LicensedInactivityCount 'DiasSemCopilot' 'CopilotLicenciado' 90 } else { $null }
}
$inactivityRows = foreach ($service in 'Login','Email','OneDrive','SharePoint','OfficeDesktop','OfficeWeb','Teams','Copilot') {
    [pscustomobject]@{ Servico=$service; 'Sem uso >30 dias'=$inactivity["${service}30"]; 'Sem uso >90 dias'=$inactivity["${service}90"] }
}
$inactivityHtml = ($inactivityRows | ConvertTo-Html -Fragment) -join "`n"
$teamsReviewRows = @($results | Where-Object CandidatoRevisaoTeams | Sort-Object @{Expression={if($null -eq $_.DiasSemTeams){[int]::MaxValue}else{$_.DiasSemTeams}};Descending=$true})
$teamsReviewHtml = if ($teamsReviewRows.Count -gt 0) {
    ($teamsReviewRows | ConvertTo-Html -Fragment -Property Nome,UPN,ContaHabilitada,LicencasAtuais,UltimoTeams,DiasSemTeams,StatusTeams) -join "`n"
} else { '<p>Nenhum usuario licenciado para Teams sem atividade acima de 30 dias.</p>' }
$serviceReviewRows = [System.Collections.Generic.List[object]]::new()
foreach ($result in $results) {
    foreach ($definition in @(
        @{Service='Email';Licensed='EmailLicenciado';Days='DiasSemEmail';Status='StatusEmail'},
        @{Service='OneDrive';Licensed='OneDriveLicenciado';Days='DiasSemOneDrive';Status='StatusOneDrive'},
        @{Service='SharePoint';Licensed='SharePointLicenciado';Days='DiasSemSharePoint';Status='StatusSharePoint'},
        @{Service='Office Desktop';Licensed='OfficeDesktopLicenciado';Days='DiasSemOfficeDesktop';Status='StatusOfficeDesktop'},
        @{Service='Office Web';Licensed='OfficeWebLicenciado';Days='DiasSemOfficeWeb';Status='StatusOfficeWeb'},
        @{Service='Teams';Licensed='TeamsLicenciado';Days='DiasSemTeams';Status='StatusTeams'},
        @{Service='Copilot';Licensed='CopilotLicenciado';Days='DiasSemCopilot';Status='StatusCopilot'}
    )) {
        if (-not $result.($definition.Licensed)) { continue }
        if ($definition.Service -eq 'Copilot' -and -not $copilotReportAvailable) { continue }
        $daysWithoutUse = $result.($definition.Days)
        if ($null -ne $daysWithoutUse -and [int]$daysWithoutUse -le 30) { continue }
        $serviceReviewRows.Add([pscustomobject]@{
            Servico=$definition.Service;Nome=$result.Nome;UPN=$result.UPN;ContaHabilitada=$result.ContaHabilitada
            LicencasAtuais=$result.LicencasAtuais;DiasSemUso=$daysWithoutUse;Status=$result.($definition.Status)
            JanelaOtimizacao=if($null -eq $daysWithoutUse -or [int]$daysWithoutUse -gt 90){'>90 dias / sem uso observado'}else{'>30 dias'}
        })
    }
}
$serviceReviewPath = Join-Path $runPath 'licencas-servicos-sem-uso.csv'
$serviceReviewRows | Sort-Object Servico,UPN | Export-Csv -LiteralPath $serviceReviewPath -NoTypeInformation -Encoding utf8BOM
$serviceReviewHtml = if ($serviceReviewRows.Count -gt 0) {
    ($serviceReviewRows | Sort-Object Servico,JanelaOtimizacao,UPN | ConvertTo-Html -Fragment -Property Servico,Nome,UPN,ContaHabilitada,LicencasAtuais,DiasSemUso,Status,JanelaOtimizacao) -join "`n"
} else { '<p>Nenhum direito de servico sem uso acima de 30 dias.</p>' }

$sharedCount = $summary.sharedMailboxCandidates
$unpricedCount = @($results | Where-Object LicencasSemPrecoPublico).Count
$planActions = @(
    [pscustomobject]@{ID=1;Pilar='Licenciamento';Acao='Extrair relatorio de usuarios e servicos ativos';Origem='Relatorios de uso';Prioridade='Alta';Responsavel='BestSoft + cliente';Prazo='';Status='Concluida pelo diagnostico';Evidencia="$($summary.usersAnalyzed) usuarios analisados; janela de $TelemetryPeriodDays dias";ProximoPasso='Validar a janela e a cobertura dos relatorios';Fonte='Microsoft Graph / Microsoft 365 admin center'},
    [pscustomobject]@{ID=2;Pilar='Licenciamento';Acao='Identificar contas inativas e licencas subutilizadas';Origem='Desperdicios';Prioridade='Alta';Responsavel='Gestores das areas';Prazo='';Status='Concluida pelo diagnostico';Evidencia="$($inactivity.Login30) sem login >30 dias; $($inactivity.Login90) >90 dias; $sharedCount candidatos a caixa compartilhada";ProximoPasso='Confirmar contexto com os gestores';Fonte='Telemetria e licencas sao sinais, nao sentenca'},
    [pscustomobject]@{ID=3;Pilar='Licenciamento';Acao='Mapear perfis: essencial, avancado e alta protecao';Origem='Perfis de usuario';Prioridade='Media';Responsavel='Cliente + BestSoft';Prazo='';Status='Em andamento';Evidencia='Perfil minimo de produtividade sugerido por uso observado';ProximoPasso='Adicionar necessidade de seguranca, conformidade e risco por area';Fonte='Validar por area'},
    [pscustomobject]@{ID=4;Pilar='Licenciamento';Acao='Revisar composicao Basic, Apps, Standard e Premium';Origem='Composicao Business';Prioridade='Alta';Responsavel='BestSoft';Prazo='';Status='Em andamento';Evidencia="Comparacao automatica de SKUs conhecidos; $unpricedCount usuarios com SKU sem preco publico completo";ProximoPasso='Comparar direitos e precos contratuais no Partner Center';Fonte='Limite agregado de 300 na familia Business'},
    [pscustomobject]@{ID=5;Pilar='Licenciamento';Acao='Avaliar Business x Enterprise para o proximo ciclo';Origem='Business x Enterprise';Prioridade='Media';Responsavel='BestSoft + cliente';Prazo='';Status='Em andamento';Evidencia="$($summary.licensedUsers) usuarios licenciados identificados";ProximoPasso='Considerar quantidade, seguranca, conformidade e crescimento';Fonte='Enterprise nao tem teto de 300'},
    [pscustomobject]@{ID=6;Pilar='Renovacao';Acao='Revisar impacto dos reajustes e novos pacotes de 2026';Origem='Novidades 2026';Prioridade='Alta';Responsavel='BestSoft';Prazo='';Status='Em andamento';Evidencia="Estimativa publica mensal de R$ $($summary.estimatedMonthlySavingsBRL.ToString('N2')); catalogo $($catalogData.asOf)";ProximoPasso='Conferir moeda, impostos, canal, vigencia e renovacao';Fonte='Microsoft Brasil e Partner Center Announcements'},
    [pscustomobject]@{ID=7;Pilar='Identidade';Acao='Verificar Security Defaults ou Acesso Condicional';Origem='Security Defaults';Prioridade='Alta';Responsavel='Administrador de identidade';Prazo='';Status='Nao avaliada neste diagnostico';Evidencia='Nao coletado pelo escopo de telemetria de licencas';ProximoPasso='Executar avaliacao de identidade e testar impacto antes de alterar';Fonte='Nao combinar Security Defaults e CA'},
    [pscustomobject]@{ID=8;Pilar='Identidade';Acao='Medir cobertura de MFA e revisar contas administrativas';Origem='Indicadores';Prioridade='Alta';Responsavel='Administrador de identidade';Prazo='';Status='Nao avaliada neste diagnostico';Evidencia='Nao coletado';ProximoPasso='Executar relatorio de MFA e separar contas administrativas';Fonte='Principio de menor privilegio'},
    [pscustomobject]@{ID=9;Pilar='Seguranca';Acao='Consultar Secure Score e separar acoes alcancaveis pela licenca atual';Origem='Indicadores';Prioridade='Media';Responsavel='Seguranca';Prazo='';Status='Nao avaliada neste diagnostico';Evidencia='Secure Score nao coletado';ProximoPasso='Executar avaliacao de seguranca';Fonte='Microsoft Secure Score'},
    [pscustomobject]@{ID=10;Pilar='Auditoria';Acao='Confirmar disponibilidade e retencao do Audit Standard';Origem='Indicadores';Prioridade='Media';Responsavel='Compliance';Prazo='';Status='Nao avaliada neste diagnostico';Evidencia='Retencao de auditoria nao coletada';ProximoPasso='Testar uma pesquisa de auditoria';Fonte='Varia por plano e retencao'},
    [pscustomobject]@{ID=11;Pilar='Colaboracao';Acao='Definir regra simples para Teams, SharePoint e OneDrive';Origem='Colaboracao';Prioridade='Media';Responsavel='Gestores das areas';Prazo='';Status='Em andamento';Evidencia="$($inactivity.SharePoint90) sem SharePoint >90 dias; $($inactivity.OneDrive90) sem OneDrive >90 dias; $($inactivity.Teams90) licenciados sem Teams >90 dias";ProximoPasso='Validar regras e necessidade de Teams com usuarios-chave';Fonte='Lugar, acesso e tempo'},
    [pscustomobject]@{ID=12;Pilar='Ciclo de vida';Acao='Padronizar onboarding com conta, licenca, grupo, dispositivo e MFA';Origem='Onboarding';Prioridade='Alta';Responsavel='RH + TI';Prazo='';Status='Nao avaliada neste diagnostico';Evidencia='Processo organizacional nao coletado';ProximoPasso='Criar checklist e aprovacao';Fonte='Revisar acesso apos entrada'},
    [pscustomobject]@{ID=13;Pilar='Ciclo de vida';Acao='Padronizar desligamento e transferencia de dados';Origem='Desligamento';Prioridade='Alta';Responsavel='RH + TI';Prazo='';Status='Em andamento';Evidencia="$sharedCount contas candidatas a caixa compartilhada por inatividade >90 dias";ProximoPasso='Validar preservacao, delegados, bloqueio de login e requisitos de licenca';Fonte='Acesso encerrado; dados preservados'},
    [pscustomobject]@{ID=14;Pilar='IA';Acao='Selecionar casos de uso de Copilot e revisar permissoes';Origem='IA responsavel';Prioridade='Media';Responsavel='Negocio + TI';Prazo='';Status='Em andamento';Evidencia="$($summary.copilotLicensedUsers) usuarios com licenca associada; $($summary.copilotActiveUsers) com uso observado na janela";ProximoPasso='Revisar licenciados sem uso e escolher metricas de adocao';Fonte='Microsoft Graph - Microsoft 365 Copilot usage'},
    [pscustomobject]@{ID=15;Pilar='Continuidade';Acao='Definir politica de backup e teste de recuperacao';Origem='Lacunas da licenca';Prioridade='Alta';Responsavel='TI + negocio';Prazo='';Status='Nao avaliada neste diagnostico';Evidencia='Backup nao e verificavel por relatorio de uso';ProximoPasso='Definir RPO, RTO, escopo e responsavel';Fonte='Retencao nao substitui backup dedicado'},
    [pscustomobject]@{ID=16;Pilar='Identidade';Acao='Manter menos de 5 Global Administrators e 2 contas de emergencia';Origem='Ganhos rapidos';Prioridade='Alta';Responsavel='Administrador de identidade';Prazo='';Status='Nao avaliada neste diagnostico';Evidencia='Funcoes administrativas nao coletadas';ProximoPasso='Inventariar administradores e testar contas break-glass';Fonte='Microsoft Entra - boas praticas de funcoes'}
)
$actionPlanPath = Join-Path $runPath 'plano-acoes.csv'
$planActions | Export-Csv -LiteralPath $actionPlanPath -NoTypeInformation -Encoding utf8BOM
$completedActions = @($planActions | Where-Object Status -Like 'Concluida*').Count
$inProgressActions = @($planActions | Where-Object Status -EQ 'Em andamento').Count
$notAssessedActions = @($planActions | Where-Object Status -Like 'Nao avaliada*').Count
$actionPlanHtml = ($planActions | ConvertTo-Html -Fragment -Property ID,Pilar,Acao,Prioridade,Responsavel,Status,Evidencia,ProximoPasso,Fonte) -join "`n"
$subscriptionsHtml = if (@($subscriptionRows).Count -gt 0) {
    ($subscriptionRows | Sort-Object Plano,DataProximoCiclo | ConvertTo-Html -Fragment -Property Plano,IDAssinatura,Status,Avaliacao,DataInicio,DataProximoCiclo,QuantidadeAssinatura,QuantidadeHabilitadaNoSku,QuantidadeUsadaNoSku,QuantidadeDisponivelNoSku,TermoFaturamento) -join "`n"
} else { '<p>Nenhuma assinatura retornada pelo Microsoft Graph.</p>' }

$rowsHtml = ($results | Sort-Object EconomiaMensalEstimadaBRL -Descending | Select-Object -First 100 |
    ConvertTo-Html -Fragment -Property Nome,UPN,LicencasAtuais,ProximosCiclosLicencas,CopilotLicenciado,CopilotUtilizado,UltimoCopilot,StatusCopilot,StatusLogin,StatusEmail,StatusOneDrive,StatusSharePoint,StatusOfficeDesktop,StatusOfficeWeb,Recomendacao,PrecoAtualMensalBRL,PrecoSugeridoMensalBRL,EconomiaMensalEstimadaBRL) -join "`n"
$html = @"
<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><style>
body{margin:0;background:#f3f6fa;color:#1f2937;font:14px 'Segoe UI',Arial,sans-serif}.page{max-width:1180px;margin:auto;background:#fff}.hero{padding:30px 34px;background:linear-gradient(135deg,#073b70,#0f6cbd);color:#fff}.hero h1{margin:0 0 8px;font-size:27px}.hero p{margin:3px 0;color:#dbeafe}.content{padding:26px 34px}.cards{display:flex;flex-wrap:wrap;gap:12px;margin:0 0 26px}.card{min-width:180px;flex:1;padding:17px;border:1px solid #dbe3ee;border-radius:10px;background:#fff;box-shadow:0 2px 8px #0000000d}.card .value{font-size:25px;font-weight:700;color:#0f6cbd}.card .label{color:#64748b;margin-top:5px}.saving .value{color:#107c10}h2{margin-top:28px;color:#073b70;border-bottom:2px solid #dbeafe;padding-bottom:8px}table{border-collapse:separate;border-spacing:0;width:100%;font-size:12px;border:1px solid #dbe3ee;border-radius:8px;overflow:hidden}th,td{padding:8px;border-bottom:1px solid #e5e7eb;text-align:left;vertical-align:top}th{background:#0f6cbd;color:#fff;white-space:nowrap}tr:nth-child(even) td{background:#f8fafc}.note{background:#fff8db;border-left:5px solid #f2c811;padding:14px 16px;border-radius:4px}.good{background:#e8f5e9;border-left:5px solid #107c10;padding:14px 16px}.cta{margin:18px 0 26px;padding:18px 20px;background:#eef6ff;border:1px solid #b7d7f5;border-radius:10px}.cta strong{display:block;color:#073b70;font-size:16px;margin-bottom:6px}.cta a{display:inline-block;margin:8px 10px 0 0;padding:8px 13px;background:#0f6cbd;color:#fff;text-decoration:none;border-radius:5px}.cta a.whatsapp{background:#107c10}.footer{margin-top:28px;padding-top:15px;border-top:1px solid #ddd;color:#64748b;font-size:12px}.table-wrap{overflow-x:auto}@media(max-width:700px){.content,.hero{padding:20px}.cards{display:block}.card{margin-bottom:10px}}
</style></head><body><div class="page"><div class="hero"><h1>Avaliacao de licencas Microsoft 365</h1><p>Resumo executivo de utilizacao e oportunidades de otimizacao</p><p>$tenantName &bull; $($now.ToString('dd/MM/yyyy HH:mm')) UTC &bull; janela de $TelemetryPeriodDays dias</p></div><div class="content">
<h2>Identificacao do relatorio</h2><table><tr><th>Organizacao</th><td>$tenantName</td><th>Dominio padrao</th><td>$defaultDomain</td></tr><tr><th>Tenant ID</th><td>$($context.TenantId)</td><th>Tipo de tenant</th><td>$tenantType</td></tr><tr><th>Gerado por</th><td>$operatorName</td><th>Conta autenticada</th><td>$($context.Account)</td></tr><tr><th>Data de geracao</th><td>$($now.ToString('dd/MM/yyyy HH:mm')) UTC</td><th>Ferramenta</th><td>M365 License Assessment v$solutionVersion</td></tr></table>
<div class="cta"><strong>Transforme os sinais deste relatorio em um plano seguro de otimizacao.</strong>Esta avaliacao automatizada e um ponto de partida. Para aprofundar licenciamento, seguranca, conformidade e economia com validacao do contexto de cada usuario, conte com a equipe BestSoft.<br><a href="https://www.bestsoft.com.br/">Conheca a BestSoft</a><a class="whatsapp" href="https://wa.me/555130265338">WhatsApp (51) 3026-5338</a></div>
<div class="cards"><div class="card"><div class="value">$(@($results).Count)</div><div class="label">Usuarios analisados</div></div><div class="card"><div class="value">$($summary.sharedMailboxCandidates)</div><div class="label">Candidatos a caixa compartilhada</div></div><div class="card"><div class="value">$($summary.removalCandidates)</div><div class="label">Candidatos a remocao</div></div><div class="card saving"><div class="value">R$ $($summary.estimatedMonthlySavingsBRL.ToString('N2'))</div><div class="label">Economia mensal estimada$(if($summary.usersWithUnpricedLicenses -gt 0){" (parcial; $($summary.usersWithUnpricedLicenses) usuarios pendentes de preco)"})</div></div></div>
<div class="good"><b>Objetivo:</b> priorizar oportunidades de economia sem executar qualquer alteracao automatica no tenant. O CSV anexo contem todos os usuarios e campos da analise.</div>
<div class="note"><b>Compromisso contratual:</b> planos anuais e planos anuais pagos mensalmente normalmente nao permitem reduzir quantidades durante a vigencia. As recomendacoes representam oportunidades para validacao e, quando aplicavel, para a proxima renovacao. Confirme termo de faturamento, data contratual e regras de reducao no Partner Center ou com o parceiro; o Microsoft Graph nao informa o termo de cobranca.</div>
<h2>Planos e assinaturas</h2><p>Quantidade comprada por assinatura e consumo agregado do SKU. A data do proximo ciclo e fornecida pelo Microsoft Graph e pode representar renovacao ou transicao de estado. Quando houver mais de uma assinatura do mesmo SKU, a quantidade usada e exibida no nivel agregado do SKU.</p><div class="table-wrap">$subscriptionsHtml</div>
<h2>Janela de inatividade por servico</h2><p>O quadro abaixo permite separar oportunidades recentes (mais de 30 dias) das mais consolidadas (mais de 90 dias).</p><div class="table-wrap">$inactivityHtml</div>
<h2>Licencas Teams sem uso observado</h2><p>Usuarios com um service plan ativo do Teams e sem atividade observada ha mais de 30 dias. Ausencia de atividade pode incluir usuarios nunca vistos na janela; confirme contexto antes de remover standalone, complemento ou migrar para uma suite sem Teams.</p><div class="table-wrap">$teamsReviewHtml</div>
<h2>Matriz licenciado versus utilizado</h2><p>Direitos ativos de email, OneDrive, SharePoint, Office Desktop, Office Web, Teams e Microsoft 365 Copilot sem uso observado acima de 30 dias. A janela acima de 90 dias indica maior prioridade de revisao.</p><div class="table-wrap">$serviceReviewHtml</div>
<h2>Recomendacoes por usuario</h2><p>Ordenadas pela maior economia mensal estimada. Caixas sem login por mais de 90 dias recebem destaque como candidatas a caixa compartilhada.</p><div class="table-wrap">$rowsHtml</div>
<h2>Plano de acoes — Microsoft 365 na pratica</h2><p><b>16</b> acoes mapeadas: <b>$completedActions</b> concluidas pelo diagnostico, <b>$inProgressActions</b> em andamento e <b>$notAssessedActions</b> ainda nao avaliadas neste escopo. Responsaveis e prazos devem ser confirmados com o cliente.</p><div class="table-wrap">$actionPlanHtml</div>
<h2>Como interpretar</h2><div class="note"><b>Revisao humana obrigatoria.</b> Ausencia de atividade nao comprova ausencia de necessidade. Antes de reduzir ou remover licencas, valide funcao do usuario, acesso delegado, tamanho da caixa, arquivo, retencao/hold, seguranca, compliance, dispositivos, Teams/telefonia e demais complementos. Valores sem preco publico nao entram na economia estimada.</div>
<div class="footer">Precos publicos de referencia em $($catalogData.asOf): $($catalogData.source). Valores podem variar por impostos, canal, contrato e promocao. Este documento e uma analise consultiva, nao uma ordem de alteracao.</div></div></div></body></html>
"@
$htmlPath = Join-Path $runPath 'relatorio.html'
Set-Content -LiteralPath $htmlPath -Value $html -Encoding utf8

Write-ExecutionStatus 90 'Gerando pacote criptografado do relatorio...'
$zipPath = Join-Path $runPath 'relatorio-m365-completo.zip'
$filesToArchive = @($csvPath, $subscriptionsPath, $actionPlanPath, $serviceReviewPath, $htmlPath, (Join-Path $runPath 'resumo.json'))
$sevenZip = Get-7ZipPath
$previousLocation = Get-Location
try {
    Set-Location -LiteralPath $runPath
    $archiveNames = @($filesToArchive | ForEach-Object { Split-Path $_ -Leaf })
    & $sevenZip a -tzip "-p$ArchivePassword" -mem=AES256 -- $zipPath @archiveNames | Out-Null
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $zipPath)) {
        throw "Falha ao gerar o ZIP criptografado (codigo $LASTEXITCODE)."
    }
} finally { Set-Location -LiteralPath $previousLocation }

if ($SendEmail) {
    Write-ExecutionStatus 94 'Enviando relatorio por email...'
    if ([string]::IsNullOrWhiteSpace($EmailTo)) { throw 'Nao foi possivel determinar o email da conta autenticada.' }
    try {
        $zipBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes($zipPath))
        $message = @{ message = @{ subject = "Avaliacao de licencas Microsoft 365 - $tenantName"; body = @{ contentType='HTML'; content=$html }
            toRecipients = @(@{emailAddress=@{address=$EmailTo}})
            bccRecipients = @(@{emailAddress=@{address=$BccAddress}})
            attachments = @(@{'@odata.type'='#microsoft.graph.fileAttachment';name='relatorio-m365-completo.zip';contentType='application/zip';contentBytes=$zipBytes})
        }; saveToSentItems = $true }
        $messageJson = $message | ConvertTo-Json -Depth 10 -Compress
        Write-Host ("  Pacote de envio: {0:N2} MB" -f ([Text.Encoding]::UTF8.GetByteCount($messageJson) / 1MB)) -ForegroundColor DarkGray
        Invoke-WithSpinner 'Enviando email' {
            Invoke-GraphRequestWithRetry -Method POST -Uri 'https://graph.microsoft.com/v1.0/me/sendMail' -Body $messageJson -ContentType 'application/json' | Out-Null
        }
    } catch {
        throw "A analise foi gerada, mas o envio falhou. Confirme que '$($context.Account)' possui caixa Exchange Online e consentimento Mail.Send. Detalhe: $($_.Exception.Message)"
    }
}

if ($SendEmail) {
    Write-Progress -Activity 'Avaliacao de licencas Microsoft 365' -Completed
    Write-Host 'Analise concluida e relatorio enviado por email.' -ForegroundColor Green
} else {
    Write-Progress -Activity 'Avaliacao de licencas Microsoft 365' -Completed
    Write-Host "Relatorio gerado em: $runPath"
}
