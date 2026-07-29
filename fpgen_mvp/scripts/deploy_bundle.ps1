#requires -Version 7.0
<#
.SYNOPSIS
    Distribute an FP-GEN JSX proposal bundle with a timestamped name:
    archive the ZIP to the test-data folder and install the runnable JSX into
    InDesign's Scripts Panel.

.DESCRIPTION
    Everything is named  <name>_<yyyyMMdd_HHmm>  (to the minute). The script:
      1. (optional) generates a fresh bundle from a running FP-GEN API (-BaseUrl),
         otherwise uses an existing bundle .zip (-BundleZip).
      2. copies the ZIP to  <TestDataDir>\<name>_<stamp>.zip  and extracts it to
         <TestDataDir>\<name>_<stamp>\  (so its assets/ folder is available).
      3. writes a copy of assemble_proposal.jsx to
         <ScriptsPanelDir>\<name>_<stamp>.jsx  with two injected globals:
           var ASSET_ROOT = "<extracted bundle path>"  -> resolves bundled assets
           var OUT_NAME    = "<name>_<stamp>"           -> timestamps the output .indd
         so the panel copy runs correctly even though it lives away from its assets.

.PARAMETER BundleZip
    Path to an existing bundle .zip (from the app's "Export JSX"). Ignored if -BaseUrl is set.
.PARAMETER BaseUrl
    If set, generate a fresh proposal + JSX bundle from this FP-GEN API first.
.PARAMETER Name
    Base name for the artifacts. Default: the proposal/project name from the bundle.
.PARAMETER TestDataDir
    Folder that receives the ZIP + the extracted bundle.
.PARAMETER ScriptsPanelDir
    InDesign user Scripts Panel folder that receives the runnable .jsx.

.EXAMPLE
    ./deploy_bundle.ps1 -BundleZip ..\..\FP-GEN_sample_bundle.zip
.EXAMPLE
    ./deploy_bundle.ps1 -BaseUrl https://no3gv3wv0vvfepji.circuits.examt001.studio.exabase.ai
#>
[CmdletBinding()]
param(
    [string]$BundleZip,
    [string]$BaseUrl,
    [string]$Name,
    [string]$TestDataDir = '\\tsclient\C\Users\03669\Desktop\Trud\matls\created\FP-GEN\exaBase\test data',
    [string]$ScriptsPanelDir = '\\tsclient\C\Users\03669\AppData\Roaming\Adobe\InDesign\Version 17.0-J\ja_JP\Scripts\Scripts Panel',
    # RDP drive redirection: files are written from this session via the \\tsclient\C
    # UNC path, but on the outer PC (where InDesign runs) they live on its local C:.
    # ASSET_ROOT injected into the JSX must therefore use the outer PC's LOCAL path.
    [string]$ClientDrivePrefix = '\\tsclient\C',
    [string]$ClientLocalDrive = 'C:'
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$stamp = Get-Date -Format 'yyyyMMdd_HHmm'

function Expand-Fresh([string]$Zip, [string]$Dest) {
    if (Test-Path $Dest) { Remove-Item $Dest -Recurse -Force }
    Expand-Archive -Path $Zip -DestinationPath $Dest -Force
}

# 1. Obtain a bundle zip -------------------------------------------------------
if ($BaseUrl) {
    $BaseUrl = $BaseUrl.TrimEnd('/')
    $repoRoot = Split-Path $PSScriptRoot -Parent
    $body = Get-Content (Join-Path $repoRoot 'examples/generate_request.json') -Raw
    $gen = Invoke-RestMethod -Uri "$BaseUrl/v1/proposals/generate" -Method POST -Body $body -ContentType 'application/json' -TimeoutSec 180
    $proposalId = $gen.proposal_id
    if (-not $Name) { $Name = $gen.proposal.project.name }
    $BundleZip = Join-Path ([System.IO.Path]::GetTempPath()) "fpgen_bundle_$stamp.zip"
    Invoke-WebRequest -Uri "$BaseUrl/v1/proposals/$proposalId/export/jsx" -Method POST `
        -Body '{"cv_assignments":{},"experience_ids":[],"template_id":"commercial"}' `
        -ContentType 'application/json' -OutFile $BundleZip -TimeoutSec 180
}
if (-not $BundleZip -or -not (Test-Path $BundleZip)) {
    throw "No bundle: pass -BundleZip <path> or -BaseUrl <url>."
}

# 2. Base name (from the bundle's proposal_data.json when not supplied) --------
if (-not $Name) {
    try {
        $peek = Join-Path ([System.IO.Path]::GetTempPath()) "fpgen_peek_$stamp"
        Expand-Fresh $BundleZip $peek
        $Name = (Get-Content (Join-Path $peek 'proposal_data.json') -Raw | ConvertFrom-Json).project.name
        Remove-Item $peek -Recurse -Force
    }
    catch { $Name = [System.IO.Path]::GetFileNameWithoutExtension($BundleZip) }
}
$safe = ($Name -replace '[^A-Za-z0-9._\- ]', '_').Trim()
if (-not $safe) { $safe = 'proposal' }
$stamped = "${safe}_${stamp}"

# 3. Ensure destination folders ------------------------------------------------
foreach ($d in @($TestDataDir, $ScriptsPanelDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Force -Path $d | Out-Null }
}

# 4. Archive ZIP + extract into the test-data folder ---------------------------
$zipDest = Join-Path $TestDataDir "$stamped.zip"
Copy-Item $BundleZip $zipDest -Force
$extractDir = Join-Path $TestDataDir $stamped
Expand-Fresh $zipDest $extractDir

# 5. Install the JSX into the Scripts Panel with injected globals --------------
$jsxSrc = Join-Path $extractDir 'assemble_proposal.jsx'
if (-not (Test-Path $jsxSrc)) { throw "assemble_proposal.jsx not found in the bundle." }
# The JSX runs in InDesign on the OUTER PC, where the bundle lives on its local drive.
# Map the \\tsclient\C write path back to that local path for ASSET_ROOT.
$runtimeExtract = $extractDir
if ($ClientDrivePrefix -and $extractDir.StartsWith($ClientDrivePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
    $runtimeExtract = $ClientLocalDrive + $extractDir.Substring($ClientDrivePrefix.Length)
}
$assetRootJs = ($runtimeExtract -replace '\\', '/')
$lines = @(Get-Content -LiteralPath $jsxSrc)
$inject = @("var ASSET_ROOT = `"$assetRootJs`";", "var OUT_NAME = `"$stamped`";")
# Keep '#target indesign' as the first line; inject the globals right after it.
$out = @($lines[0]) + $inject + $lines[1..($lines.Count - 1)]
$jsxDest = Join-Path $ScriptsPanelDir "$stamped.jsx"
Set-Content -LiteralPath $jsxDest -Value $out -Encoding UTF8

Write-Host "Deployed ($stamp):" -ForegroundColor Green
Write-Host "  ZIP archive : $zipDest"
Write-Host "  Extracted   : $extractDir"
Write-Host "  JSX (panel) : $jsxDest"
Write-Host "  Asset root  : $assetRootJs  (path as seen by InDesign on the outer PC)"
Write-Host ""
Write-Host "In InDesign: Window > Utilities > Scripts > User > $stamped.jsx  (double-click to run)."
