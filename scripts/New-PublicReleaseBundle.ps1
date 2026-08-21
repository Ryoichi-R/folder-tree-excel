#Requires -Version 7.0
<#
.SYNOPSIS
    test candidateとrelease candidateを同一の共通入力からfresh buildし、cross-candidate検証、
    entrypoint 3-control gate、offline package検証を経て、release candidateだけから
    version付きasset・SHA256SUMS.txt・release-manifest.jsonを dist/<version>/<nonce>/ へ生成する。

.DESCRIPTION
    Phase 3.2 release bundle generator。既存artifactを直接Release名へrenameするだけの経路は設けない。
    release safety subsetは root外move拒否・reparse祖先拒否・destination collision拒否の
    3項目とし、-SkipReleaseSafetySubset を明示しない限り Invoke-ReleaseSafetySubset.ps1 で
    production UI経路を通して実際に検証する。

    Recycle Bin無効化を伴うfallback拒否の動的検証はOD-11（2026-08-20のowner判断）により
    実施しない。この安全特性は modRecycleBin.bas の NoRecycleFiles / NukeOnDelete ガード節の
    静的source reviewと、embedded VBAがcanonical sourceへhash一致することの確認だけに依拠し、
    manifestへ recycleBinUnavailableFallback = "static-source-review-only" として明示記録する。
#>
[CmdletBinding()]
param(
    [switch] $SkipReleaseSafetySubset
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd([IO.Path]::DirectorySeparatorChar)
if (-not (Test-Path -LiteralPath (Join-Path $root 'VERSION') -PathType Leaf) -or -not (Test-Path -LiteralPath (Join-Path $root 'src') -PathType Container)) {
    throw "Repository root identity check failed at $root (VERSION or src\ missing)."
}
$fullVersion = (Get-Content -LiteralPath (Join-Path $root 'VERSION') -Raw).Trim()
$displayVersion = "v$fullVersion"
$pwsh = (Get-Command pwsh -ErrorAction Stop).Source
$bundleNonce = [guid]::NewGuid().ToString('N')
$bundleRoot = Join-Path $root ".build-work\bundle-$bundleNonce"
New-Item -ItemType Directory -Path $bundleRoot -Force | Out-Null

function Get-Sha256([string] $path) { return (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash }

Write-Host "=== Step 1: common inputs ===" -ForegroundColor Cyan
$sourceNames = @('modFolderTree.bas', 'modOperationPlan.bas', 'modFileOperations.bas', 'modRecycleBin.bas', 'modWorkbookIo.bas')
$commonInputs = [ordered]@{
    sourceHashes = [ordered]@{}
    builderSha256 = Get-Sha256 (Join-Path $root 'scripts\build-xlsm.ps1')
    testScriptSha256 = Get-Sha256 (Join-Path $root 'scripts\test-xlsm.ps1')
    publicationPolicySha256 = Get-Sha256 (Join-Path $root 'release\publication-policy.json')
    entrypointInventoryPolicySha256 = Get-Sha256 (Join-Path $root 'release\vba-entrypoint-inventory-policy.json')
    packageDiffAllowlistSha256 = Get-Sha256 (Join-Path $root 'release\package-diff-allowlist.json')
    versionSha256 = Get-Sha256 (Join-Path $root 'VERSION')
}
foreach ($name in $sourceNames) { $commonInputs.sourceHashes[$name] = Get-Sha256 (Join-Path $root "src\$name") }
Write-Host "version=$fullVersion builder=$($commonInputs.builderSha256.Substring(0,12))..."

Write-Host "=== Step 2-3: test candidate build + full gate ===" -ForegroundColor Cyan
$testCandidatePath = Join-Path $bundleRoot "test-candidate.xlsm"
$testBuildReceiptPath = "$testCandidatePath.build-receipt.json"
$testGateReceiptPath = Join-Path $bundleRoot "test-candidate.gate.json"
$testPerformanceReceiptPath = Join-Path $bundleRoot "test-candidate.performance.json"
$testWorkRoot = Join-Path $root ".test-work\bundle-$bundleNonce"

& $pwsh -NoProfile -File (Join-Path $root 'scripts\build-xlsm.ps1') -CandidateOnly -Profile Test -OutputPath $testCandidatePath -CandidateBuildReceiptPath $testBuildReceiptPath
if ($LASTEXITCODE -ne 0) { throw "Test candidate build failed: exit $LASTEXITCODE" }

& $pwsh -NoProfile -File (Join-Path $root 'scripts\test-xlsm.ps1') -WorkbookPath $testCandidatePath -CandidateBuildReceiptPath $testBuildReceiptPath -TestWorkRoot $testWorkRoot -GateReceiptPath $testGateReceiptPath
$testGateExit = $LASTEXITCODE
if (Test-Path -LiteralPath $testWorkRoot) {
    $remaining = @(Get-ChildItem -LiteralPath $testWorkRoot -Force -ErrorAction SilentlyContinue)
    if ($remaining.Count -eq 0) { Remove-Item -LiteralPath $testWorkRoot -Force }
}
if ($testGateExit -ne 0) { throw "Test candidate required gate failed: exit $testGateExit" }

& $pwsh -NoProfile -File (Join-Path $root 'scripts\measure-performance.ps1') -WorkbookPath $testCandidatePath -ReceiptPath $testPerformanceReceiptPath
if ($LASTEXITCODE -ne 0) { throw "Test candidate performance gate failed: exit $LASTEXITCODE" }

$testGateReceipt = Get-Content -LiteralPath $testGateReceiptPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (@($testGateReceipt.requiredCaseIds).Count -ne 68 -or -not $testGateReceipt.allPassed) { throw 'Test candidate did not pass all 68 required cases.' }
Write-Host "test candidate: 68/68 required cases pass"

Write-Host "=== Step 4: release candidate build ===" -ForegroundColor Cyan
$releaseCandidatePath = Join-Path $bundleRoot "release-candidate.xlsm"
$releaseBuildReceiptPath = "$releaseCandidatePath.build-receipt.json"
& $pwsh -NoProfile -File (Join-Path $root 'scripts\build-xlsm.ps1') -CandidateOnly -Profile Release -OutputPath $releaseCandidatePath -CandidateBuildReceiptPath $releaseBuildReceiptPath
if ($LASTEXITCODE -ne 0) { throw "Release candidate build failed: exit $LASTEXITCODE" }
$releaseBuildReceipt = Get-Content -LiteralPath $releaseBuildReceiptPath -Raw -Encoding UTF8 | ConvertFrom-Json
foreach ($validationName in @('contentType', 'vbaProjectPart', 'vbaCompile', 'fileFormat', 'sheetOrder', 'validations', 'buttons', 'embeddedVba', 'noTestModule', 'defaultUiNoInjection', 'protection', 'reopened')) {
    $status = $releaseBuildReceipt.validation.$validationName.status
    if ($status -ne 'pass') { throw "Release candidate validation not pass: $validationName" }
}
Write-Host "release candidate: build+compile+embedded hash pass"

Write-Host "=== Step 5: cross-candidate part comparison ===" -ForegroundColor Cyan
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Get-PackagePartHashes([string] $path) {
    $zip = [IO.Compression.ZipFile]::OpenRead($path)
    $result = [ordered]@{}
    try {
        foreach ($entry in $zip.Entries) {
            $stream = $entry.Open(); $ms = New-Object IO.MemoryStream; $stream.CopyTo($ms)
            $bytes = $ms.ToArray()
            $result[$entry.FullName] = ([BitConverter]::ToString([Security.Cryptography.SHA256]::HashData($bytes))).Replace('-', '')
            $stream.Dispose(); $ms.Dispose()
        }
    } finally { $zip.Dispose() }
    return $result
}
$diffAllowlist = Get-Content -LiteralPath (Join-Path $root 'release\package-diff-allowlist.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$allowedVolatileParts = @($diffAllowlist.volatileXPaths | ForEach-Object { $_.part }) | Select-Object -Unique
$allowedProfileParts = @($diffAllowlist.profileSpecificParts | ForEach-Object { $_.part })
$testParts = Get-PackagePartHashes $testCandidatePath
$releaseParts = Get-PackagePartHashes $releaseCandidatePath
$allPartNames = @($testParts.Keys) + @($releaseParts.Keys) | Select-Object -Unique | Sort-Object
$unexpectedDiffs = [System.Collections.Generic.List[string]]::new()
foreach ($name in $allPartNames) {
    $inTest = $testParts.Contains($name); $inRelease = $releaseParts.Contains($name)
    if (-not $inTest -or -not $inRelease) { $unexpectedDiffs.Add("PART_PRESENCE_MISMATCH: $name"); continue }
    if ($testParts[$name] -ne $releaseParts[$name]) {
        if ($allowedVolatileParts -notcontains $name -and $allowedProfileParts -notcontains $name) {
            $unexpectedDiffs.Add("UNEXPECTED_PART_DIFF: $name")
        }
    }
}
if ($unexpectedDiffs.Count -gt 0) { throw "Cross-candidate comparison found unlisted differences:`n$($unexpectedDiffs -join "`n")" }
Write-Host "cross-candidate: all part diffs are within package-diff-allowlist.json"

Write-Host "=== Step 6: release candidate entrypoint 3-control gate + offline package ===" -ForegroundColor Cyan
& $pwsh -NoProfile -File (Join-Path $root 'scripts\build-xlsm.ps1') -ValidatePackageOffline -CandidatePath $releaseCandidatePath
if ($LASTEXITCODE -ne 0) { throw 'Release candidate offline package validation failed.' }

$inventoryPolicy = Get-Content -LiteralPath (Join-Path $root 'release\vba-entrypoint-inventory-policy.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$allTestWrapperNames = @()
foreach ($moduleName in $inventoryPolicy.modules.PSObject.Properties.Name) {
    $allTestWrapperNames += @($inventoryPolicy.modules.$moduleName.testWrappers | ForEach-Object { $_.name })
}
function Get-NormalizedMacroUndefinedSignature([string] $errorMessage, [string] $macroName) {
    # エラーメッセージ自体にマクロ名が埋め込まれるため、名前部分をplaceholderへ正規化してから比較する。
    return $errorMessage.Replace($macroName, '{MACRO_NAME}')
}
$excel = $null; $wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AutomationSecurity = 1
    $wb = $excel.Workbooks.Open($releaseCandidatePath)
    $positiveResult = $excel.Run('GetScanSnapshotCount')
    if ($null -eq $positiveResult -or [int64]$positiveResult -lt 0) { throw "Positive control entrypoint returned unexpected value: $positiveResult" }
    $dummyName = "NoSuchMacro_$([guid]::NewGuid().ToString('N'))"
    $dummySignature = $null
    try { $excel.Run($dummyName); throw "Dummy macro unexpectedly succeeded: $dummyName" }
    catch [System.Runtime.InteropServices.COMException] { $dummySignature = Get-NormalizedMacroUndefinedSignature $_.Exception.Message $dummyName }
    if ([string]::IsNullOrWhiteSpace($dummySignature)) { throw 'Dummy macro did not produce a COM exception signature.' }
    foreach ($wrapperName in $allTestWrapperNames) {
        $wrapperSignature = $null
        try { $excel.Run($wrapperName); throw "Test wrapper unexpectedly compiled into release candidate: $wrapperName" }
        catch [System.Runtime.InteropServices.COMException] { $wrapperSignature = Get-NormalizedMacroUndefinedSignature $_.Exception.Message $wrapperName }
        if ($wrapperSignature -ne $dummySignature) { throw "Test wrapper '$wrapperName' error signature differs from dummy macro (possible partial compile): actual=[$wrapperSignature] expected=[$dummySignature]" }
    }
    Write-Host ("entrypoint gate: positive control pass, {0} test wrappers confirmed non-compiled" -f $allTestWrapperNames.Count)
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {}; [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($null -ne $excel) { try { $excel.Quit() } catch {}; [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host "=== Step 6b: release safety subset (root外/reparse/collision) ===" -ForegroundColor Cyan
$releaseSafetySubset = [ordered]@{
    status = 'skipped-by-flag'
    checks = [ordered]@{
        rootBoundary = 'not-run'
        reparseBoundary = 'not-run'
        collision = 'not-run'
        recycleBinUnavailableFallback = 'static-source-review-only'
    }
}
if (-not $SkipReleaseSafetySubset) {
    $safetySubsetScript = Join-Path $root 'scripts\Invoke-ReleaseSafetySubset.ps1'
    $safetySubsetReceiptPath = Join-Path $root ".test-work\release-safety-subset-bundle-$bundleNonce.receipt.json"
    & $pwsh -NoProfile -File $safetySubsetScript -CandidatePath $releaseCandidatePath -ReceiptPath $safetySubsetReceiptPath
    $safetySubsetExit = $LASTEXITCODE
    if (-not (Test-Path -LiteralPath $safetySubsetReceiptPath -PathType Leaf)) { throw 'Release safety subset did not produce a receipt.' }
    $safetySubsetReceipt = Get-Content -LiteralPath $safetySubsetReceiptPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($safetySubsetReceipt.candidateSha256 -ne (Get-Sha256 $releaseCandidatePath)) { throw 'Release safety subset receipt candidate hash mismatch.' }
    $releaseSafetySubset = [ordered]@{
        status = $(if ($safetySubsetExit -eq 0) { 'pass' } else { 'fail' })
        checks = [ordered]@{
            rootBoundary = $safetySubsetReceipt.checks.rootBoundary.status
            reparseBoundary = $safetySubsetReceipt.checks.reparseBoundary.status
            collision = $safetySubsetReceipt.checks.collision.status
            recycleBinUnavailableFallback = 'static-source-review-only'
        }
        receiptSha256 = Get-Sha256 $safetySubsetReceiptPath
    }
    if ($safetySubsetExit -ne 0) { throw "Release safety subset failed: exit $safetySubsetExit (see $safetySubsetReceiptPath)" }
    Write-Host 'release safety subset: root外/reparse/collisionの3項目pass。Recycle Bin無効化fallbackはOD-11により動的検証の対象外（静的source reviewのみ）。' -ForegroundColor Green
} else {
    Write-Host 'release safety subset: -SkipReleaseSafetySubset のためskip。' -ForegroundColor Yellow
}

Write-Host "=== Step 7: signing ===" -ForegroundColor Cyan
Write-Host 'signing profile: unsigned (OD-3). No signing action taken.'

Write-Host "=== Step 8: release assets ===" -ForegroundColor Cyan
$distDir = Join-Path $root "dist\$fullVersion\$bundleNonce"
New-Item -ItemType Directory -Path $distDir -Force | Out-Null
$assetFileName = "folder-tree-excel-$displayVersion.xlsm"
$assetPath = Join-Path $distDir $assetFileName
Copy-Item -LiteralPath $releaseCandidatePath -Destination $assetPath -Force
$assetSha256 = Get-Sha256 $assetPath
$sha256SumsPath = Join-Path $distDir 'SHA256SUMS.txt'
"$assetSha256  $assetFileName" | Set-Content -LiteralPath $sha256SumsPath -Encoding utf8NoBOM
$performanceWaiverSha256 = 'CF3CC8B3D32D063FBCB3ED7BD594EE3C68F5429B34BF2A11BA96523529B494C8'
$knownLimitations = @(
    [ordered]@{
        summary = '実Recycle BinのShell ExtendedProperty latencyは少数項目の実測に限定されます。'
        waiverReceiptSha256 = $performanceWaiverSha256
    },
    [ordered]@{
        summary = '削除件数1,001〜10,000件の所要時間は「見積不能」として扱い、保証しません。'
        waiverReceiptSha256 = $performanceWaiverSha256
    }
)

$manifest = [ordered]@{
    schemaVersion = 1
    generatedAtUtc = [DateTime]::UtcNow.ToString('o')
    version = $fullVersion
    displayVersion = $displayVersion
    signing = [ordered]@{ status = 'unsigned' }
    asset = [ordered]@{ fileName = $assetFileName; sha256 = $assetSha256; bytes = (Get-Item -LiteralPath $assetPath).Length }
    sourceHashes = $commonInputs.sourceHashes
    builderSha256 = $commonInputs.builderSha256
    testCandidate = [ordered]@{
        candidateSha256 = Get-Sha256 $testCandidatePath
        buildReceiptSha256 = Get-Sha256 $testBuildReceiptPath
        gateReceiptSha256 = Get-Sha256 $testGateReceiptPath
        requiredCasesPassed = 68
    }
    releaseCandidate = [ordered]@{
        candidateSha256 = Get-Sha256 $releaseCandidatePath
        buildReceiptSha256 = Get-Sha256 $releaseBuildReceiptPath
        entrypointGatePassed = $true
        offlinePackagePassed = $true
    }
    releaseSafetySubset = $releaseSafetySubset
    knownLimitations = $knownLimitations
}
$manifestPath = Join-Path $distDir 'release-manifest.json'
$manifestJson = $manifest | ConvertTo-Json -Depth 10
$manifestSchemaPath = Join-Path $root 'release\release-manifest.schema.json'
if (-not (Test-Json -Json $manifestJson -SchemaFile $manifestSchemaPath)) {
    throw 'Generated release-manifest.json does not validate against release-manifest.schema.json.'
}
$manifestJson | Set-Content -LiteralPath $manifestPath -Encoding utf8NoBOM

Write-Host "=== Step 9: freeze ===" -ForegroundColor Cyan
Write-Host "dist bundle: $distDir"
Write-Host "  asset:    $assetFileName ($assetSha256)"
Write-Host "  manifest: release-manifest.json"
Write-Host "  test candidate frozen as evidence-only at: $testCandidatePath (not copied into dist/)"
Write-Host "  releaseSafetySubset: $($releaseSafetySubset.status) (recycleBinUnavailableFallback=static-source-review-only, OD-11)"
if ($releaseSafetySubset.status -ne 'pass') {
    Write-Host ("WARNING: releaseSafetySubset.status={0}. Re-run without -SkipReleaseSafetySubset before treating this bundle as gate-complete." -f $releaseSafetySubset.status) -ForegroundColor Yellow
}
Write-Output "distDir=$distDir"
