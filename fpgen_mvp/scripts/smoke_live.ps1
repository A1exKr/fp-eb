#requires -Version 7.0
<#
.SYNOPSIS
    Live HTTP smoke suite for the FP-GEN MVP API deployed inside exaBase Studio.

.DESCRIPTION
    Exercises the public endpoints over the internet against a deployed instance
    (either the v2 SQLite topology or the PostgreSQL topology). Authentication is
    assumed DISABLED (the app returns a synthetic dev-admin identity), so the
    admin routes are reachable. Each check prints PASS/FAIL and the script exits
    non-zero if any check fails.

    Covers: health/capabilities/identity, master-data reads, the AI-backed parse
    and proposal-generation path, fee calculation, proposal get/update, JSX + INDD
    export, admin CRUD (personnel/rates/presets/reference-projects), and the CV +
    project-experience upload facility.

.PARAMETER BaseUrl
    Base URL of the deployed app, e.g.
    https://xxxx.circuits.examt001.studio.exabase.ai

.PARAMETER NoAiChecks
    Skip the assertions that expect the OpenAI-backed parse path (use when the
    LLM key is not configured and the deterministic fallback is expected).

.PARAMETER KeepData
    Do not delete the admin records/assets created during the run.

.EXAMPLE
    ./smoke_live.ps1 -BaseUrl https://xxxx.circuits.examt001.studio.exabase.ai

.EXAMPLE
    ./smoke_live.ps1 -BaseUrl https://xxxx.studio.exabase.ai -NoAiChecks -KeepData
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$BaseUrl,
    [switch]$NoAiChecks,
    [switch]$KeepData,
    [string]$ForwardedUser,
    [string]$ForwardedGroups = 'fpgen_admin,Finance'
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

if ($BaseUrl -notmatch '^https?://') { $BaseUrl = "https://$BaseUrl" }
$script:BaseUrl = $BaseUrl.TrimEnd('/')

# Optional forwarded identity for auth-enabled deployments whose proxy passes
# X-Forwarded-* through (the app trusts these headers when FPGEN_AUTH_ENABLED=true).
$script:AuthHeaders = @{}
if ($ForwardedUser) {
    $script:AuthHeaders['X-Forwarded-User'] = $ForwardedUser
    $script:AuthHeaders['X-Forwarded-Email'] = $ForwardedUser
    $script:AuthHeaders['X-Forwarded-Groups'] = $ForwardedGroups
}

$Root = Split-Path $PSScriptRoot -Parent
$SampleRfp = Join-Path $Root 'examples/sample_rfp.txt'
$GenReqPath = Join-Path $Root 'examples/generate_request.json'
$CvFile = Join-Path $Root 'data/asset_files/cv-test.txt'
$ExpFile = Join-Path $Root 'examples/sample_experience.txt'

# Exact deterministic (non-AI) methodology string emitted by the regex fallback
# parser. If parse returns something different, the AI path was exercised.
$FALLBACK_METHODOLOGY = "1. Kick-off`n2. Analysis`n3. Concept Development`n4. Refinement`n5. Finalization`n6. Quality Control"

foreach ($f in @($SampleRfp, $GenReqPath, $CvFile, $ExpFile)) {
    if (-not (Test-Path $f)) { throw "Required test-data file not found: $f" }
}

$script:Pass = 0
$script:Fail = 0
$script:Skipped = 0

function Write-Skip {
    param([string]$Name)
    $script:Skipped++
    Write-Host "  [SKIP] $Name" -ForegroundColor Yellow
}

function Assert-That {
    param([string]$Name, [bool]$Condition, [string]$Detail = '')
    if ($Condition) {
        $script:Pass++
        Write-Host "  [PASS] $Name" -ForegroundColor Green
    }
    else {
        $script:Fail++
        $suffix = if ($Detail) { " -> $Detail" } else { '' }
        Write-Host "  [FAIL] $Name$suffix" -ForegroundColor Red
    }
}

function Invoke-Api {
    param(
        [string]$Method,
        [string]$Path,
        [object]$JsonBody,
        [hashtable]$Form
    )
    $params = @{
        Method             = $Method
        Uri                = $script:BaseUrl + $Path
        SkipHttpErrorCheck = $true
        TimeoutSec         = 120
    }
    if ($script:AuthHeaders.Count -gt 0) { $params.Headers = $script:AuthHeaders }
    if ($PSBoundParameters.ContainsKey('JsonBody') -and $null -ne $JsonBody) {
        $params.Body = ($JsonBody | ConvertTo-Json -Depth 12)
        $params.ContentType = 'application/json'
    }
    if ($Form) { $params.Form = $Form }

    try {
        $resp = Invoke-WebRequest @params
    }
    catch {
        # Transport-level failure (DNS, TLS, connection refused): report as status 0.
        return [pscustomobject]@{ Status = 0; Content = "$($_.Exception.Message)"; Json = $null; Headers = @{} }
    }

    $json = $null
    if ($resp.Content) {
        try { $json = $resp.Content | ConvertFrom-Json } catch { $json = $null }
    }
    [pscustomobject]@{
        Status  = [int]$resp.StatusCode
        Content = $resp.Content
        Json    = $json
        Headers = $resp.Headers
    }
}

Write-Host "== FP-GEN live smoke suite ==" -ForegroundColor Cyan
Write-Host "Target: $script:BaseUrl`n"

# --------------------------------------------------------------------------- #
# Health & identity
# --------------------------------------------------------------------------- #
Write-Host "[Health & identity]" -ForegroundColor Cyan
$health = Invoke-Api GET '/health'
Assert-That 'GET /health -> 200' ($health.Status -eq 200) "status=$($health.Status)"
Assert-That 'health.status == ok' ($health.Status -eq 200 -and $health.Json.status -eq 'ok')
$inddCap = if ($health.Status -eq 200) { $health.Json.capabilities.indd_export } else { $null }
$inddEnabled = if ($inddCap -is [bool]) { [bool]$inddCap } elseif ($null -ne $inddCap) { [bool]$inddCap.enabled } else { $null }
$inddDetail = if ($null -ne $inddCap) { $inddCap | ConvertTo-Json -Compress } else { 'null' }
Assert-That 'health.capabilities.indd_export disabled (Linux)' ($inddEnabled -eq $false) "got=$inddDetail"

$cap = Invoke-Api GET '/v1/capabilities'
Assert-That 'GET /v1/capabilities -> 200' ($cap.Status -eq 200) "status=$($cap.Status)"

$me = Invoke-Api GET '/v1/me'
$authOpen = ($me.Status -eq 200)
if ($authOpen) {
    Assert-That 'GET /v1/me -> 200 (identity available)' $true
    Assert-That 'me.user is present' (-not [string]::IsNullOrWhiteSpace([string]$me.Json.user))
}
else {
    Write-Host "  [NOTE] /v1/me -> $($me.Status): app auth is ENABLED and no identity is forwarded." -ForegroundColor Yellow
    Write-Host "         Admin CRUD + CV/experience uploads will be SKIPPED. Set FPGEN_AUTH_ENABLED=false on" -ForegroundColor Yellow
    Write-Host "         the api pod (or pass -ForwardedUser if the proxy forwards identity) to exercise them." -ForegroundColor Yellow
}

# --------------------------------------------------------------------------- #
# Master-data reads
# --------------------------------------------------------------------------- #
Write-Host "`n[Master-data reads]" -ForegroundColor Cyan
$refProjects = Invoke-Api GET '/v1/reference-projects'
Assert-That 'GET /v1/reference-projects -> 200' ($refProjects.Status -eq 200) "status=$($refProjects.Status)"
$personnel = Invoke-Api GET '/v1/personnel'
Assert-That 'GET /v1/personnel -> 200' ($personnel.Status -eq 200) "status=$($personnel.Status)"
$presets = Invoke-Api GET '/v1/team-presets'
Assert-That 'GET /v1/team-presets -> 200' ($presets.Status -eq 200) "status=$($presets.Status)"
$assets = Invoke-Api GET '/v1/assets'
Assert-That 'GET /v1/assets -> 200 (template/cvs/experience)' ($assets.Status -eq 200 -and $null -ne $assets.Json.experience) "status=$($assets.Status)"

# --------------------------------------------------------------------------- #
# Parse (AI path)
# --------------------------------------------------------------------------- #
Write-Host "`n[Parse / AI path]" -ForegroundColor Cyan
$parseBody = @{
    rfp_text     = "Client requests a mixed-use master plan in Ho Chi Minh City with sustainability, mobility, and phased delivery over 16 weeks. Deliverables include concept options, a final master plan report, and presentation boards."
    project_hint = 'HCMC Mixed-Use Master Plan'
}
$parse = Invoke-Api POST '/v1/parse' -JsonBody $parseBody
Assert-That 'POST /v1/parse -> 200' ($parse.Status -eq 200) "status=$($parse.Status)"
Assert-That 'parse.parsed.project.name is set' ($parse.Status -eq 200 -and -not [string]::IsNullOrWhiteSpace([string]$parse.Json.parsed.project.name))
if (-not $NoAiChecks -and $parse.Status -eq 200) {
    $methodology = [string]$parse.Json.parsed.methodology.text
    Assert-That 'parse used the AI path (methodology != canned fallback)' ($methodology -and $methodology -ne $FALLBACK_METHODOLOGY) 'methodology matched the deterministic fallback string'
}

$parseFile = Invoke-Api POST '/v1/parse/file' -Form @{ rfp_file = Get-Item $SampleRfp; project_hint = 'HCMC Mixed-Use Master Plan' }
Assert-That 'POST /v1/parse/file -> 200' ($parseFile.Status -eq 200) "status=$($parseFile.Status)"
Assert-That 'parse/file returns extracted_text' ($parseFile.Status -eq 200 -and -not [string]::IsNullOrWhiteSpace([string]$parseFile.Json.extracted_text))

# --------------------------------------------------------------------------- #
# Fee & proposals
# --------------------------------------------------------------------------- #
Write-Host "`n[Fee & proposals]" -ForegroundColor Cyan
$genReq = Get-Content $GenReqPath -Raw | ConvertFrom-Json

$fee = Invoke-Api POST '/v1/fee/calculate' -JsonBody $genReq.fee_input
Assert-That 'POST /v1/fee/calculate -> 200' ($fee.Status -eq 200) "status=$($fee.Status)"
Assert-That 'fee.calculate returns a fee object' ($fee.Status -eq 200 -and $null -ne $fee.Json.fee)

$gen = Invoke-Api POST '/v1/proposals/generate' -JsonBody $genReq
Assert-That 'POST /v1/proposals/generate -> 200' ($gen.Status -eq 200) "status=$($gen.Status)"
$proposalId = if ($gen.Status -eq 200) { [string]$gen.Json.proposal_id } else { '' }
Assert-That 'generate returned a proposal_id' (-not [string]::IsNullOrWhiteSpace($proposalId))
Assert-That 'generate produced a cover_letter section' ($gen.Status -eq 200 -and -not [string]::IsNullOrWhiteSpace([string]$gen.Json.proposal.sections.cover_letter))
Assert-That 'generate produced markdown output' ($gen.Status -eq 200 -and -not [string]::IsNullOrWhiteSpace([string]$gen.Json.proposal.markdown))

if ($proposalId) {
    $get = Invoke-Api GET "/v1/proposals/$proposalId"
    Assert-That 'GET /v1/proposals/{id} -> 200' ($get.Status -eq 200) "status=$($get.Status)"
    Assert-That 'retrieved proposal matches id' ($get.Status -eq 200 -and $get.Json.proposal_id -eq $proposalId)

    $override = 'Smoke-test override cover letter.'
    $put = Invoke-Api PUT "/v1/proposals/$proposalId" -JsonBody @{ sections = @{ cover_letter = $override } }
    Assert-That 'PUT /v1/proposals/{id} -> 200' ($put.Status -eq 200) "status=$($put.Status)"
    Assert-That 'PUT applied the section override' ($put.Status -eq 200 -and [string]$put.Json.proposal.sections.cover_letter -eq $override)

    $jsx = Invoke-Api POST "/v1/proposals/$proposalId/export/jsx" -JsonBody @{ cv_assignments = @{}; experience_ids = @(); template_id = 'commercial' }
    $contentType = [string]$jsx.Headers['Content-Type']
    Assert-That 'POST /v1/proposals/{id}/export/jsx -> 200 zip' ($jsx.Status -eq 200 -and $contentType -match 'zip') "status=$($jsx.Status), type=$contentType"

    # INDD export is disabled on Linux (no COM) -> expect 500; 200 is fine on a Windows host.
    $indd = Invoke-Api POST "/v1/proposals/$proposalId/export/indd"
    Assert-That 'POST /v1/proposals/{id}/export/indd -> disabled(500) or ok(200)' ($indd.Status -eq 500 -or $indd.Status -eq 200) "status=$($indd.Status)"
}

# generate-from-file
$feeJson = $genReq.fee_input | ConvertTo-Json -Depth 12 -Compress
$overridesJson = @{
    project_name = 'HCMC Mixed-Use Master Plan'
    client_name  = 'Example Client'
    location     = 'Ho Chi Minh City, Vietnam'
    project_type = 'Master Plan'
} | ConvertTo-Json -Compress
$genFile = Invoke-Api POST '/v1/proposals/generate/file' -Form @{
    rfp_file                    = Get-Item $SampleRfp
    fee_input_json              = $feeJson
    selected_reference_ids_json = '[]'
    overrides_json              = $overridesJson
}
Assert-That 'POST /v1/proposals/generate/file -> 200' ($genFile.Status -eq 200) "status=$($genFile.Status)"
Assert-That 'generate/file returned a proposal_id' ($genFile.Status -eq 200 -and -not [string]::IsNullOrWhiteSpace([string]$genFile.Json.proposal_id))

# --------------------------------------------------------------------------- #
# Admin CRUD + uploads
# --------------------------------------------------------------------------- #
Write-Host "`n[Admin CRUD + uploads]" -ForegroundColor Cyan
$created = @{ personnel = @(); rates = @(); presets = @(); refs = @(); assets = @() }
if (-not $authOpen) {
    Write-Skip 'Admin CRUD + CV/experience uploads (app auth enabled; no forwarded identity)'
}
else {
    $personCreate = Invoke-Api POST '/v1/admin/personnel' -JsonBody @{ id = 'smoke-p-001'; name = 'Smoke Tester'; title = 'QA Engineer'; roles = @('Architect') }
    Assert-That 'POST /v1/admin/personnel -> 200' ($personCreate.Status -eq 200) "status=$($personCreate.Status)"
    if ($personCreate.Status -eq 200) { $created.personnel += 'smoke-p-001' }
    $personList = Invoke-Api GET '/v1/admin/personnel'
    Assert-That 'admin personnel list contains the created record' ($personList.Status -eq 200 -and [bool]($personList.Json | Where-Object { $_.id -eq 'smoke-p-001' }))

    $rateCreate = Invoke-Api POST '/v1/admin/rates' -JsonBody @{ role = 'SmokeRole'; rate = 123; currency = 'USD' }
    Assert-That 'POST /v1/admin/rates -> 200' ($rateCreate.Status -eq 200) "status=$($rateCreate.Status)"
    $rateList = Invoke-Api GET '/v1/admin/rates'
    $rateRow = if ($rateList.Status -eq 200) { $rateList.Json | Where-Object { $_.role -eq 'SmokeRole' } | Select-Object -First 1 } else { $null }
    Assert-That 'admin rates list contains the created rate' ([bool]$rateRow)
    if ($rateRow) { $created.rates += $rateRow.id }

    $presetCreate = Invoke-Api POST '/v1/admin/presets' -JsonBody @{ id = 'smoke-preset'; name = 'Smoke Preset'; types = @('Commercial'); assignments = @(@{ role = 'Architect' }) }
    Assert-That 'POST /v1/admin/presets -> 200' ($presetCreate.Status -eq 200) "status=$($presetCreate.Status)"
    if ($presetCreate.Status -eq 200) { $created.presets += 'smoke-preset' }

    $refCreate = Invoke-Api POST '/v1/admin/reference-projects' -JsonBody @{ id = 'smoke-rp'; name = 'Smoke Reference Project'; project_type = 'Commercial'; location = 'Tokyo'; keywords = @('smoke', 'test'); summary = 'A smoke-test reference project.' }
    Assert-That 'POST /v1/admin/reference-projects -> 200' ($refCreate.Status -eq 200) "status=$($refCreate.Status)"
    if ($refCreate.Status -eq 200) { $created.refs += 'smoke-rp' }

    # CV upload
    $cvUpload = Invoke-Api POST '/v1/admin/assets/upload' -Form @{ asset_id = 'smoke-cv'; kind = 'cv'; role = 'Architect'; file = Get-Item $CvFile }
    Assert-That 'POST /v1/admin/assets/upload (cv) -> 200' ($cvUpload.Status -eq 200) "status=$($cvUpload.Status)"
    Assert-That 'cv upload persisted size_bytes > 0' ($cvUpload.Status -eq 200 -and [int]$cvUpload.Json.size_bytes -gt 0)
    if ($cvUpload.Status -eq 200) { $created.assets += 'smoke-cv' }

    # Project-experience upload (the "upload project experience" facility)
    $expUpload = Invoke-Api POST '/v1/admin/assets/upload' -Form @{ asset_id = 'smoke-exp'; kind = 'experience'; reference_project_id = 'smoke-rp'; file = Get-Item $ExpFile }
    Assert-That 'POST /v1/admin/assets/upload (experience) -> 200' ($expUpload.Status -eq 200) "status=$($expUpload.Status)"
    Assert-That 'experience upload returns a storage_ref' ($expUpload.Status -eq 200 -and -not [string]::IsNullOrWhiteSpace([string]$expUpload.Json.storage_ref))
    if ($expUpload.Status -eq 200) { $created.assets += 'smoke-exp' }
}

# --------------------------------------------------------------------------- #
# Static pages
# --------------------------------------------------------------------------- #
Write-Host "`n[Static pages]" -ForegroundColor Cyan
$index = Invoke-Api GET '/'
Assert-That 'GET / -> 200' ($index.Status -eq 200) "status=$($index.Status)"
$adminPage = Invoke-Api GET '/admin/'
Assert-That 'GET /admin/ -> 200' ($adminPage.Status -eq 200) "status=$($adminPage.Status)"

# --------------------------------------------------------------------------- #
# Cleanup
# --------------------------------------------------------------------------- #
if (-not $KeepData) {
    Write-Host "`n[Cleanup]" -ForegroundColor Cyan
    foreach ($id in $created.assets) { Invoke-Api DELETE "/v1/admin/assets/$id" | Out-Null }
    foreach ($id in $created.presets) { Invoke-Api DELETE "/v1/admin/presets/$id" | Out-Null }
    foreach ($id in $created.refs) { Invoke-Api DELETE "/v1/admin/reference-projects/$id" | Out-Null }
    foreach ($id in $created.rates) { Invoke-Api DELETE "/v1/admin/rates/$id" | Out-Null }
    foreach ($id in $created.personnel) { Invoke-Api DELETE "/v1/admin/personnel/$id" | Out-Null }
    Write-Host "  Removed smoke-test records."
}
else {
    Write-Host "`n[Cleanup] skipped (-KeepData)." -ForegroundColor Yellow
}

# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
Write-Host "`n== Summary ==" -ForegroundColor Cyan
Write-Host ("  PASS: {0}" -f $script:Pass) -ForegroundColor Green
Write-Host ("  SKIP: {0}" -f $script:Skipped) -ForegroundColor Yellow
$failColor = if ($script:Fail -gt 0) { 'Red' } else { 'Green' }
Write-Host ("  FAIL: {0}" -f $script:Fail) -ForegroundColor $failColor

Write-Host "`nManual browser walkthrough still to do:" -ForegroundColor Yellow
Write-Host "  1. $script:BaseUrl/ui     -> upload examples/sample_rfp.txt, Parse, Generate, edit a section, Export JSX."
Write-Host "  2. $script:BaseUrl/admin/  -> add a person/rate/preset/reference project; upload a CV and an experience file."

exit ([int]($script:Fail -gt 0))
