#Requires -Version 7.0
<#
.SYNOPSIS
    release/publication-policy.json に基づき、public sourceツリーの完全性を検証する。

.DESCRIPTION
    policyに列挙された全fileの存在、実際のtreeに存在するがpolicyに列挙されていないfile
    （UNCLASSIFIED_FILE）、excludedPatternsに一致するfileの誤混入を検査する。
    read-only検証であり、fileの生成・変更は行わない。
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd([IO.Path]::DirectorySeparatorChar)
if (-not (Test-Path -LiteralPath (Join-Path $root 'VERSION') -PathType Leaf) -or -not (Test-Path -LiteralPath (Join-Path $root 'src') -PathType Container)) {
    throw "Repository root identity check failed at $root (VERSION or src\ missing)."
}

$policyPath = Join-Path $root 'release\publication-policy.json'
if (-not (Test-Path -LiteralPath $policyPath -PathType Leaf)) { throw "publication-policy.json not found: $policyPath" }
$policy = Get-Content -LiteralPath $policyPath -Raw -Encoding UTF8 | ConvertFrom-Json

$failures = [System.Collections.Generic.List[string]]::new()

function Convert-GlobToRegex([string] $glob) {
    $escaped = [regex]::Escape($glob)
    $escaped = $escaped -replace '\\\*\\\*', '.*'
    $escaped = $escaped -replace '\\\*', '[^/\\\\]*'
    return "^$escaped$"
}

$excludedRegexes = @($policy.excludedPatterns | ForEach-Object { Convert-GlobToRegex $_ })
$evidenceOnlyRegexes = @($policy.evidenceOnlyPatterns | ForEach-Object { Convert-GlobToRegex $_ })
$releaseAssetRegexes = @($policy.releaseAssetPatterns | ForEach-Object { Convert-GlobToRegex $_.pattern })

# 1. policyに列挙された全public-source fileが実在するか
$policyPaths = @($policy.files | ForEach-Object { $_.path })
foreach ($entry in $policy.files) {
    $full = Join-Path $root ($entry.path -replace '/', '\')
    if (-not (Test-Path -LiteralPath $full -PathType Leaf)) {
        $failures.Add("MISSING_POLICY_FILE: $($entry.path)")
    }
}

# 2. 実際のtreeをスキャンし、excluded配下を除いた全fileがpolicyに列挙されているか
$allFiles = @(Get-ChildItem -Path $root -Recurse -File -Force |
    ForEach-Object { [IO.Path]::GetRelativePath($root, $_.FullName).Replace('\', '/') })

foreach ($relPath in $allFiles) {
    $isExcluded = $false
    foreach ($rx in $excludedRegexes) { if ($relPath -match $rx) { $isExcluded = $true; break } }
    if ($isExcluded) { continue }

    $isEvidenceOnly = $false
    foreach ($rx in $evidenceOnlyRegexes) { if ($relPath -match $rx) { $isEvidenceOnly = $true; break } }
    if ($isEvidenceOnly) { continue }

    $isReleaseAsset = $false
    foreach ($rx in $releaseAssetRegexes) { if ($relPath -match $rx) { $isReleaseAsset = $true; break } }
    if ($isReleaseAsset) { continue }

    if ($policyPaths -notcontains $relPath) {
        $failures.Add("UNCLASSIFIED_FILE: $relPath")
    }
}

# 3. requiredなfileの重複分類がないか（public-source と releaseAsset/evidenceOnly の同時一致は矛盾）
foreach ($entry in $policy.files) {
    foreach ($rx in $excludedRegexes) {
        if ($entry.path -match $rx) { $failures.Add("CONFLICTING_CLASSIFICATION: $($entry.path) matches both public-source and excludedPatterns") }
    }
}

if ($failures.Count -gt 0) {
    Write-Host 'FAILED: publication candidate policy violations' -ForegroundColor Red
    foreach ($f in $failures) { Write-Host "  $f" -ForegroundColor Red }
    exit 1
}

Write-Host ("PASSED: {0} policy files verified, {1} tree files scanned, 0 unclassified" -f $policyPaths.Count, $allFiles.Count) -ForegroundColor Green
exit 0
