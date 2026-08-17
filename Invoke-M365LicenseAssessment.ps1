#requires -Version 7.2
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'output'),
    [string]$PriceCatalogPath = (Join-Path $PSScriptRoot 'config\license-catalog.pt-BR.json'),
    [ValidateSet(30, 90, 180)] [int]$TelemetryPeriodDays = 180,
    [string]$EmailTo,
    [string]$BccAddress = 'suprote@bestsoft.com.br',
    [string]$ExpectedAccount,
    [switch]$SendEmail,
    [switch]$IncludeGuests
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$solutionVersion = '1.0.4'

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
        $response = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
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
        Invoke-MgGraphRequest -Method GET -Uri $uri -OutputFilePath $tempFile | Out-Null
        if ((Test-Path $tempFile) -and (Get-Item $tempFile).Length -gt 0) {
            return @(Import-Csv -LiteralPath $tempFile)
        }
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
    return [math]::Floor(($Now - $Date.Value).TotalDays)
}

function Get-UsageState {
    param([Nullable[int]]$Days)
    if ($null -eq $Days) { return 'Sem atividade observada' }
    if ($Days.Value -gt 90) { return 'Inativo >90 dias' }
    if ($Days.Value -gt 30) { return 'Inativo >30 dias' }
    return 'Ativo <=30 dias'
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

Import-RequiredModule Microsoft.Graph.Authentication
if (-not (Test-Path -LiteralPath $PriceCatalogPath)) { throw "Catalogo nao encontrado: $PriceCatalogPath" }
$catalogData = Get-Content -LiteralPath $PriceCatalogPath -Raw | ConvertFrom-Json
$catalog = @($catalogData.plans)

$scopes = @('User.Read.All', 'AuditLog.Read.All', 'LicenseAssignment.Read.All', 'Reports.Read.All')
if ($SendEmail) { $scopes += 'Mail.Send' }
Write-ExecutionStatus 5 'Aguardando autenticacao no Microsoft 365...'
Write-Host "M365 License Assessment v$solutionVersion" -ForegroundColor Green
Connect-MgGraph -Scopes $scopes -NoWelcome

$context = Get-MgContext
if ($ExpectedAccount -and $context.Account -ine $ExpectedAccount) {
    Disconnect-MgGraph | Out-Null
    throw "A conta autenticada '$($context.Account)' nao corresponde ao usuario informado '$ExpectedAccount'. Execute novamente e entre com o usuario informado na instalacao."
}
if ($SendEmail -and [string]::IsNullOrWhiteSpace($EmailTo)) { $EmailTo = $context.Account }
Write-ExecutionStatus 10 'Autenticacao concluida. Iniciando coleta; esta etapa pode levar alguns minutos.'
$now = [datetime]::UtcNow
$runId = $now.ToString('yyyyMMdd-HHmmss')
$runPath = Join-Path $OutputPath $runId
New-Item -ItemType Directory -Path $runPath -Force | Out-Null

$usersUri = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName,userType,accountEnabled,assignedLicenses,signInActivity&`$top=999"
Write-ExecutionStatus 18 'Coletando usuarios e atividade de login...'
$users = @(Invoke-WithSpinner 'Consultando diretorio' { Get-GraphCollection $usersUri })
if (-not $IncludeGuests) { $users = @($users | Where-Object userType -EQ 'Member') }
Write-ExecutionStatus 25 'Coletando assinaturas e SKUs do tenant...'
$skus = @(Invoke-WithSpinner 'Consultando assinaturas' { Get-GraphCollection 'https://graph.microsoft.com/v1.0/subscribedSkus' })

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
$skuById = @{}; foreach ($sku in $skus) { $skuById[[string]$sku.skuId] = $sku }
$priceBySku = @{}; foreach ($plan in $catalog) { foreach ($part in $plan.skuPartNumbers) { $priceBySku[$part] = $plan } }

Write-ExecutionStatus 72 'Consolidando telemetria e avaliando licencas...'
$results = foreach ($user in $users) {
    $key = [string]$user.userPrincipalName.ToLowerInvariant()
    $active = $activeIndex[$key]; $apps = $appIndex[$key]; $email = $emailIndex[$key]
    $drive = $oneDriveIndex[$key]; $site = $sharePointIndex[$key]
    $signInDate = Convert-ToDateOrNull $user.signInActivity.lastSuccessfulSignInDateTime
    if ($null -eq $signInDate) { $signInDate = Convert-ToDateOrNull $user.signInActivity.lastSignInDateTime }
    $emailDate = Convert-ToDateOrNull (Find-PropertyValue $email @('Last Activity Date'))
    $driveDate = Convert-ToDateOrNull (Find-PropertyValue $drive @('Last Activity Date'))
    $siteDate = Convert-ToDateOrNull (Find-PropertyValue $site @('Last Activity Date'))
    $desktopDate = Convert-ToDateOrNull (Find-PropertyValue $apps @('Last Activity Date', 'Last Activated Date'))
    $webDate = Convert-ToDateOrNull (Find-PropertyValue $active @('Office 365 Last Activity Date', 'Microsoft 365 Apps Last Activity Date'))
    $signInDays = Get-DaysSince $signInDate $now; $emailDays = Get-DaysSince $emailDate $now
    $driveDays = Get-DaysSince $driveDate $now; $siteDays = Get-DaysSince $siteDate $now
    $desktopDays = Get-DaysSince $desktopDate $now; $webDays = Get-DaysSince $webDate $now

    $assignedParts = @($user.assignedLicenses | ForEach-Object { $skuById[[string]$_.skuId].skuPartNumber } | Where-Object { $_ })
    $knownCurrentPlans = @($assignedParts | ForEach-Object { $priceBySku[$_] } | Where-Object { $_ } | Sort-Object name -Unique)
    $currentPrice = ($knownCurrentPlans | Measure-Object monthlyPriceBRL -Sum).Sum
    $unpriced = @($assignedParts | Where-Object { -not $priceBySku.ContainsKey($_) })
    $needs = @{
        Email = ($null -ne $emailDays -and $emailDays -le 90)
        OneDrive = ($null -ne $driveDays -and $driveDays -le 90)
        SharePoint = ($null -ne $siteDays -and $siteDays -le 90)
        OfficeWeb = ($null -ne $webDays -and $webDays -le 90)
        OfficeDesktop = ($null -ne $desktopDays -and $desktopDays -le 90)
    }
    $minimum = Get-MinimumPlan $needs $catalog
    $hasExchangeLicense = @($knownCurrentPlans | Where-Object { $_.features.email }).Count -gt 0
    $sharedMailboxCandidate = ($hasExchangeLicense -and $null -ne $signInDays -and $signInDays -gt 90)
    $recommendation = if ($assignedParts.Count -eq 0) { 'Sem licenca atribuida' }
        elseif ($sharedMailboxCandidate) { 'Candidato a caixa compartilhada; validar tamanho, arquivo, retencao/hold, acesso delegado e bloquear login antes de remover licencas' }
        elseif ($null -eq $minimum) { 'Candidato a remocao completa; validar funcao, retencao, caixa compartilhada e requisitos de seguranca' }
        else { "Avaliar $($minimum.name)" }
    $recommendedPrice = if ($sharedMailboxCandidate) { [decimal]0 } elseif ($minimum) { [decimal]$minimum.monthlyPriceBRL } else { [decimal]0 }
    $saving = if ($unpriced.Count -eq 0) { [math]::Max(0, [decimal]$currentPrice - $recommendedPrice) } else { $null }

    [pscustomobject]@{
        Nome = $user.displayName; UPN = $user.userPrincipalName; ContaHabilitada = $user.accountEnabled
        LicencasAtuais = ($assignedParts -join '; '); LicencasSemPrecoPublico = ($unpriced -join '; ')
        PrecoAtualMensalBRL = if ($unpriced.Count -eq 0) { [decimal]$currentPrice } else { $null }
        UltimoLogin = $signInDate; DiasSemLogin = $signInDays; StatusLogin = Get-UsageState $signInDays
        UltimoEmail = $emailDate; DiasSemEmail = $emailDays; StatusEmail = Get-UsageState $emailDays
        UltimoOneDrive = $driveDate; DiasSemOneDrive = $driveDays; StatusOneDrive = Get-UsageState $driveDays
        UltimoSharePoint = $siteDate; DiasSemSharePoint = $siteDays; StatusSharePoint = Get-UsageState $siteDays
        UltimoOfficeDesktop = $desktopDate; DiasSemOfficeDesktop = $desktopDays; StatusOfficeDesktop = Get-UsageState $desktopDays
        UltimoOfficeWeb = $webDate; DiasSemOfficeWeb = $webDays; StatusOfficeWeb = Get-UsageState $webDays
        CandidatoCaixaCompartilhada = $sharedMailboxCandidate
        Recomendacao = $recommendation; PlanoMinimoSugerido = if ($sharedMailboxCandidate) { 'Caixa compartilhada (sem licenca, se elegivel)' } elseif ($minimum) { $minimum.name } else { $null }
        PrecoSugeridoMensalBRL = $recommendedPrice; EconomiaMensalEstimadaBRL = $saving
    }
}

$csvPath = Join-Path $runPath 'analise-usuarios.csv'
Write-ExecutionStatus 82 'Gerando relatorio executivo e anexos...'
$results | Sort-Object UPN | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding utf8BOM
$summary = [pscustomobject]@{
    generatedAtUtc = $now.ToString('o'); tenantId = $context.TenantId; account = $context.Account
    telemetryPeriodDays = $TelemetryPeriodDays; usersAnalyzed = @($results).Count
    licensedUsers = @($results | Where-Object LicencasAtuais).Count
    removalCandidates = @($results | Where-Object { $_.Recomendacao -like 'Candidato a remocao*' }).Count
    sharedMailboxCandidates = @($results | Where-Object CandidatoCaixaCompartilhada).Count
    estimatedMonthlySavingsBRL = [decimal](($results | Measure-Object EconomiaMensalEstimadaBRL -Sum).Sum)
    priceCatalogAsOf = $catalogData.asOf; priceSource = $catalogData.source
}
$summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $runPath 'resumo.json') -Encoding utf8

function Get-InactivityCount([string]$DaysProperty, [int]$Threshold) {
    return @($results | Where-Object { $null -ne $_.$DaysProperty -and $_.$DaysProperty -gt $Threshold }).Count
}
$inactivity = [ordered]@{
    Login30 = Get-InactivityCount 'DiasSemLogin' 30; Login90 = Get-InactivityCount 'DiasSemLogin' 90
    Email30 = Get-InactivityCount 'DiasSemEmail' 30; Email90 = Get-InactivityCount 'DiasSemEmail' 90
    OneDrive30 = Get-InactivityCount 'DiasSemOneDrive' 30; OneDrive90 = Get-InactivityCount 'DiasSemOneDrive' 90
    SharePoint30 = Get-InactivityCount 'DiasSemSharePoint' 30; SharePoint90 = Get-InactivityCount 'DiasSemSharePoint' 90
    OfficeDesktop30 = Get-InactivityCount 'DiasSemOfficeDesktop' 30; OfficeDesktop90 = Get-InactivityCount 'DiasSemOfficeDesktop' 90
    OfficeWeb30 = Get-InactivityCount 'DiasSemOfficeWeb' 30; OfficeWeb90 = Get-InactivityCount 'DiasSemOfficeWeb' 90
}
$inactivityRows = foreach ($service in 'Login','Email','OneDrive','SharePoint','OfficeDesktop','OfficeWeb') {
    [pscustomobject]@{ Servico=$service; 'Sem uso >30 dias'=$inactivity["${service}30"]; 'Sem uso >90 dias'=$inactivity["${service}90"] }
}
$inactivityHtml = ($inactivityRows | ConvertTo-Html -Fragment) -join "`n"

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
    [pscustomobject]@{ID=11;Pilar='Colaboracao';Acao='Definir regra simples para Teams, SharePoint e OneDrive';Origem='Colaboracao';Prioridade='Media';Responsavel='Gestores das areas';Prazo='';Status='Em andamento';Evidencia="$($inactivity.SharePoint90) sem SharePoint >90 dias; $($inactivity.OneDrive90) sem OneDrive >90 dias";ProximoPasso='Validar regras com usuarios-chave e incluir Teams em avaliacao complementar';Fonte='Lugar, acesso e tempo'},
    [pscustomobject]@{ID=12;Pilar='Ciclo de vida';Acao='Padronizar onboarding com conta, licenca, grupo, dispositivo e MFA';Origem='Onboarding';Prioridade='Alta';Responsavel='RH + TI';Prazo='';Status='Nao avaliada neste diagnostico';Evidencia='Processo organizacional nao coletado';ProximoPasso='Criar checklist e aprovacao';Fonte='Revisar acesso apos entrada'},
    [pscustomobject]@{ID=13;Pilar='Ciclo de vida';Acao='Padronizar desligamento e transferencia de dados';Origem='Desligamento';Prioridade='Alta';Responsavel='RH + TI';Prazo='';Status='Em andamento';Evidencia="$sharedCount contas candidatas a caixa compartilhada por inatividade >90 dias";ProximoPasso='Validar preservacao, delegados, bloqueio de login e requisitos de licenca';Fonte='Acesso encerrado; dados preservados'},
    [pscustomobject]@{ID=14;Pilar='IA';Acao='Selecionar casos de uso de Copilot e revisar permissoes';Origem='IA responsavel';Prioridade='Media';Responsavel='Negocio + TI';Prazo='';Status='Nao avaliada neste diagnostico';Evidencia='Uso e permissoes do Copilot nao coletados';ProximoPasso='Escolher grupo piloto e metricas';Fonte='IA respeita permissoes existentes'},
    [pscustomobject]@{ID=15;Pilar='Continuidade';Acao='Definir politica de backup e teste de recuperacao';Origem='Lacunas da licenca';Prioridade='Alta';Responsavel='TI + negocio';Prazo='';Status='Nao avaliada neste diagnostico';Evidencia='Backup nao e verificavel por relatorio de uso';ProximoPasso='Definir RPO, RTO, escopo e responsavel';Fonte='Retencao nao substitui backup dedicado'},
    [pscustomobject]@{ID=16;Pilar='Identidade';Acao='Manter menos de 5 Global Administrators e 2 contas de emergencia';Origem='Ganhos rapidos';Prioridade='Alta';Responsavel='Administrador de identidade';Prazo='';Status='Nao avaliada neste diagnostico';Evidencia='Funcoes administrativas nao coletadas';ProximoPasso='Inventariar administradores e testar contas break-glass';Fonte='Microsoft Entra - boas praticas de funcoes'}
)
$actionPlanPath = Join-Path $runPath 'plano-acoes.csv'
$planActions | Export-Csv -LiteralPath $actionPlanPath -NoTypeInformation -Encoding utf8BOM
$completedActions = @($planActions | Where-Object Status -Like 'Concluida*').Count
$inProgressActions = @($planActions | Where-Object Status -EQ 'Em andamento').Count
$notAssessedActions = @($planActions | Where-Object Status -Like 'Nao avaliada*').Count
$actionPlanHtml = ($planActions | ConvertTo-Html -Fragment -Property ID,Pilar,Acao,Prioridade,Responsavel,Status,Evidencia,ProximoPasso,Fonte) -join "`n"

$rowsHtml = ($results | Sort-Object EconomiaMensalEstimadaBRL -Descending | Select-Object -First 500 |
    ConvertTo-Html -Fragment -Property Nome,UPN,LicencasAtuais,StatusLogin,StatusEmail,StatusOneDrive,StatusSharePoint,StatusOfficeDesktop,StatusOfficeWeb,Recomendacao,PrecoAtualMensalBRL,PrecoSugeridoMensalBRL,EconomiaMensalEstimadaBRL) -join "`n"
$html = @"
<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><style>
body{margin:0;background:#f3f6fa;color:#1f2937;font:14px 'Segoe UI',Arial,sans-serif}.page{max-width:1180px;margin:auto;background:#fff}.hero{padding:30px 34px;background:linear-gradient(135deg,#073b70,#0f6cbd);color:#fff}.hero h1{margin:0 0 8px;font-size:27px}.hero p{margin:3px 0;color:#dbeafe}.content{padding:26px 34px}.cards{display:flex;flex-wrap:wrap;gap:12px;margin:0 0 26px}.card{min-width:180px;flex:1;padding:17px;border:1px solid #dbe3ee;border-radius:10px;background:#fff;box-shadow:0 2px 8px #0000000d}.card .value{font-size:25px;font-weight:700;color:#0f6cbd}.card .label{color:#64748b;margin-top:5px}.saving .value{color:#107c10}h2{margin-top:28px;color:#073b70;border-bottom:2px solid #dbeafe;padding-bottom:8px}table{border-collapse:separate;border-spacing:0;width:100%;font-size:12px;border:1px solid #dbe3ee;border-radius:8px;overflow:hidden}th,td{padding:8px;border-bottom:1px solid #e5e7eb;text-align:left;vertical-align:top}th{background:#0f6cbd;color:#fff;white-space:nowrap}tr:nth-child(even) td{background:#f8fafc}.note{background:#fff8db;border-left:5px solid #f2c811;padding:14px 16px;border-radius:4px}.good{background:#e8f5e9;border-left:5px solid #107c10;padding:14px 16px}.footer{margin-top:28px;padding-top:15px;border-top:1px solid #ddd;color:#64748b;font-size:12px}.table-wrap{overflow-x:auto}@media(max-width:700px){.content,.hero{padding:20px}.cards{display:block}.card{margin-bottom:10px}}
</style></head><body><div class="page"><div class="hero"><h1>Avaliacao de licencas Microsoft 365</h1><p>Resumo executivo de utilizacao e oportunidades de otimizacao</p><p>Tenant $($context.TenantId) &bull; $($now.ToString('dd/MM/yyyy HH:mm')) UTC &bull; janela de $TelemetryPeriodDays dias</p></div><div class="content">
<div class="cards"><div class="card"><div class="value">$(@($results).Count)</div><div class="label">Usuarios analisados</div></div><div class="card"><div class="value">$($summary.sharedMailboxCandidates)</div><div class="label">Candidatos a caixa compartilhada</div></div><div class="card"><div class="value">$($summary.removalCandidates)</div><div class="label">Candidatos a remocao</div></div><div class="card saving"><div class="value">R$ $($summary.estimatedMonthlySavingsBRL.ToString('N2'))</div><div class="label">Economia mensal estimada</div></div></div>
<div class="good"><b>Objetivo:</b> priorizar oportunidades de economia sem executar qualquer alteracao automatica no tenant. O CSV anexo contem todos os usuarios e campos da analise.</div>
<h2>Janela de inatividade por servico</h2><p>O quadro abaixo permite separar oportunidades recentes (mais de 30 dias) das mais consolidadas (mais de 90 dias).</p><div class="table-wrap">$inactivityHtml</div>
<h2>Recomendacoes por usuario</h2><p>Ordenadas pela maior economia mensal estimada. Caixas sem login por mais de 90 dias recebem destaque como candidatas a caixa compartilhada.</p><div class="table-wrap">$rowsHtml</div>
<h2>Plano de acoes — Microsoft 365 na pratica</h2><p><b>16</b> acoes mapeadas: <b>$completedActions</b> concluidas pelo diagnostico, <b>$inProgressActions</b> em andamento e <b>$notAssessedActions</b> ainda nao avaliadas neste escopo. Responsaveis e prazos devem ser confirmados com o cliente.</p><div class="table-wrap">$actionPlanHtml</div>
<h2>Como interpretar</h2><div class="note"><b>Revisao humana obrigatoria.</b> Ausencia de atividade nao comprova ausencia de necessidade. Antes de reduzir ou remover licencas, valide funcao do usuario, acesso delegado, tamanho da caixa, arquivo, retencao/hold, seguranca, compliance, dispositivos, Teams/telefonia e demais complementos. Valores sem preco publico nao entram na economia estimada.</div>
<div class="footer">Precos publicos de referencia em $($catalogData.asOf): $($catalogData.source). Valores podem variar por impostos, canal, contrato e promocao. Este documento e uma analise consultiva, nao uma ordem de alteracao.</div></div></div></body></html>
"@
$htmlPath = Join-Path $runPath 'relatorio.html'
Set-Content -LiteralPath $htmlPath -Value $html -Encoding utf8

if ($SendEmail) {
    Write-ExecutionStatus 94 'Enviando relatorio por email...'
    if ([string]::IsNullOrWhiteSpace($EmailTo)) { throw 'Nao foi possivel determinar o email da conta autenticada.' }
    $attachmentBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes($csvPath))
    $actionPlanBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes($actionPlanPath))
    $message = @{ message = @{ subject = 'Avaliacao de licencas Microsoft 365'; body = @{ contentType='HTML'; content=$html }
        toRecipients = @(@{emailAddress=@{address=$EmailTo}})
        bccRecipients = @(@{emailAddress=@{address=$BccAddress}})
        attachments = @(
            @{'@odata.type'='#microsoft.graph.fileAttachment';name='analise-usuarios.csv';contentType='text/csv';contentBytes=$attachmentBytes},
            @{'@odata.type'='#microsoft.graph.fileAttachment';name='plano-acoes.csv';contentType='text/csv';contentBytes=$actionPlanBytes}
        )
    }; saveToSentItems = $true }
    try {
        Invoke-WithSpinner 'Enviando email' {
            Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/me/sendMail' -Body ($message | ConvertTo-Json -Depth 10) -ContentType 'application/json' | Out-Null
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
