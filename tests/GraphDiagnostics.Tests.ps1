BeforeAll {
    $sourcePath = Join-Path $PSScriptRoot '../Invoke-M365LicenseAssessment.ps1'
    $tokens = $null; $parseErrors = $null
    $assessmentAst = [Management.Automation.Language.Parser]::ParseFile($sourcePath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw ($parseErrors.Message -join '; ') }
    foreach ($definition in $assessmentAst.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]}, $false)) {
        . ([scriptblock]::Create($definition.Extent.Text))
    }
    # Load functions without running login, collection, file generation or email.
    function Invoke-MgGraphRequest { param($Method,$Uri,$Body,$OutputType,$OutputFilePath,$ContentType) }
    $solutionVersion = '1.4.2'
    function New-TestGraphFailure {
        param([int]$Status = 403, [string]$Code = 'Authorization_RequestDenied', [string]$Message = 'Insufficient privileges to complete the operation.')
        $exception = [Net.Http.HttpRequestException]::new('Response status code does not indicate success: ' + $Status, $null, [Net.HttpStatusCode]$Status)
        $record = [Management.Automation.ErrorRecord]::new($exception,'GraphFailure',[Management.Automation.ErrorCategory]::PermissionDenied,$null)
        $record.ErrorDetails = [Management.Automation.ErrorDetails]::new((@{error=@{code=$Code;message=$Message;innerError=@{'request-id'='11111111-1111-1111-1111-111111111111'}}} | ConvertTo-Json -Depth 5))
        return $record
    }
}

Describe 'Graph failures preserve the denied API and recovery guidance' {
    BeforeEach { Mock Start-Sleep {} }
    It 'reports HTTP, endpoint, Graph code, correlation id and sign-in prerequisites' {
        Mock Invoke-MgGraphRequest { throw (New-TestGraphFailure) }
        $failure = $null
        try { Invoke-GraphRequestWithRetry GET 'https://graph.microsoft.com/v1.0/users?$select=id,signInActivity' } catch { $failure = $_ }
        $failure | Should -Not -BeNullOrEmpty
        $failure.Exception.Message | Should -BeLike '*HTTP: 403*Authorization_RequestDenied*11111111-1111-1111-1111-111111111111*Entra ID P1/P2*Reports Reader*'
        $failure.Exception.InnerException | Should -Not -BeNullOrEmpty
        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly
        Should -Invoke Start-Sleep -Times 0
    }
    It 'does not mistake a number in a forbidden error message for a transient status' {
        Mock Invoke-MgGraphRequest { throw (New-TestGraphFailure -Message 'Forbidden for object 500.') }
        { Invoke-GraphRequestWithRetry GET 'https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail' } | Should -Throw '*Reports.Read.All*Reports Reader*'
        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly
        Should -Invoke Start-Sleep -Times 0
    }
    It 'keeps the forbidden fallback when a provider omits structured status and body' {
        Mock Invoke-MgGraphRequest { throw 'Response status code does not indicate success: Forbidden (Forbidden).' }
        { Invoke-GraphRequestWithRetry GET 'https://graph.microsoft.com/v1.0/subscribedSkus' } | Should -Throw '*HTTP: 403*LicenseAssignment.Read.All*'
    }
    It 'retries throttling and returns a successful response unchanged' {
        $script:attempt = 0
        Mock Invoke-MgGraphRequest {
            $script:attempt++
            if ($script:attempt -eq 1) { throw (New-TestGraphFailure -Status 429 -Code TooManyRequests) }
            [pscustomobject]@{value=@('ok')}
        }
        (Invoke-GraphRequestWithRetry GET 'https://graph.microsoft.com/v1.0/organization').value | Should -Be 'ok'
        Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly
        Should -Invoke Start-Sleep -Times 1 -Exactly
    }
    It 'bounds transient retries and preserves the final API failure' {
        Mock Invoke-MgGraphRequest { throw (New-TestGraphFailure -Status 503 -Code ServiceUnavailable) }
        { Invoke-GraphRequestWithRetry GET 'https://graph.microsoft.com/v1.0/organization' -MaxAttempts 3 } | Should -Throw '*HTTP: 503*ServiceUnavailable*'
        Should -Invoke Invoke-MgGraphRequest -Times 3 -Exactly
        Should -Invoke Start-Sleep -Times 2 -Exactly
    }
    It 'explains authentication expiry separately from authorization' {
        $diagnostic = Get-GraphFailureDiagnostic (New-TestGraphFailure -Status 401 -Code InvalidAuthenticationToken) GET 'https://graph.microsoft.com/v1.0/me'
        $diagnostic.Action | Should -BeLike '*Autentique novamente*'
        $diagnostic.Action | Should -Not -BeLike '*Reports Reader*'
    }
    It 'saves failure details outside the installation without request body or headers' {
        Mock Invoke-MgGraphRequest { throw (New-TestGraphFailure) }
        try { Invoke-GraphRequestWithRetry POST 'https://graph.microsoft.com/v1.0/me/sendMail' -Body 'sensitive-body-must-not-be-logged' } catch { $failure = $_ }
        $path = Save-AssessmentFailure $failure -Directory $TestDrive
        $savedText = Get-Content -LiteralPath $path -Raw
        $saved = $savedText | ConvertFrom-Json
        $saved.Graph.Uri | Should -Be 'https://graph.microsoft.com/v1.0/me/sendMail'
        $saved.Graph.HttpStatus | Should -Be 403
        $saved.Graph.Action | Should -BeLike '*Mail.Send*Exchange Online*'
        $savedText | Should -Not -BeLike '*sensitive-body-must-not-be-logged*'
    }
    It 'does not replace the original failure when diagnostic storage fails' {
        Mock Set-Content { throw 'disk unavailable' }
        Save-AssessmentFailure (New-TestGraphFailure) -Directory $TestDrive | Should -BeLike 'indisponivel*'
    }
}

Describe 'Preflight fails before collection and email when access is denied' {
    BeforeAll {
        $preflight = $assessmentAst.FindAll({param($node)
            $node -is [Management.Automation.Language.TryStatementAst] -and $node.Extent.Text -match 'Pre-validacao interrompida'
        }, $false) | Select-Object -First 1
        $preflightCode = [scriptblock]::Create($preflight.Extent.Text)
    }
    BeforeEach {
        $SendEmail = $true; $EmailTo = 'test@example.com'
        Mock Invoke-WithSpinner { & $Operation }
        Mock Get-ReportCsv { @() }
        Mock Start-Sleep {}
        Mock Save-AssessmentFailure { Join-Path $TestDrive 'diagnostic.json' }
    }
    It 'stops on sign-in access denial and identifies that exact API' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -match 'signInActivity') { throw (New-TestGraphFailure) }
            [pscustomobject]@{value=@()}
        }
        { & $preflightCode } | Should -Throw '*Diagnostico preservado:*diagnostic.json*users?*signInActivity*HTTP: 403*'
        Should -Invoke Invoke-MgGraphRequest -Times 0 -ParameterFilter { $Method -eq 'POST' }
        Should -Invoke Get-ReportCsv -Times 0
        Should -Invoke Save-AssessmentFailure -Times 1 -Exactly
    }
    It 'checks all six reports and email after required reads succeed' {
        Mock Invoke-MgGraphRequest { [pscustomobject]@{value=@()} }
        { & $preflightCode } | Should -Not -Throw
        Should -Invoke Get-ReportCsv -Times 6 -Exactly
        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'POST' -and $Uri -like '*/me/sendMail' }
        Should -Invoke Save-AssessmentFailure -Times 0
    }
}
