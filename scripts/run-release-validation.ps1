#Requires -Version 7.0
<##
.SYNOPSIS
    リリース候補のcandidate build、必須gate、性能測定を順番に実行する。

.DESCRIPTION
    candidate-only経路だけを使用し、配布物への昇格は行わない。
    各段階のexit codeとhash bindingを検査し、全段階成功時だけsummary receiptを保存する。
#>
[CmdletBinding()]
param(
    [string] $CandidatePath,
    [string] $BuildReceiptPath,
    [string] $GateReceiptPath,
    [string] $PerformanceReceiptPath,
    [string] $SummaryReceiptPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd([IO.Path]::DirectorySeparatorChar)
if (-not (Test-Path -LiteralPath (Join-Path $root 'VERSION') -PathType Leaf) -or -not (Test-Path -LiteralPath (Join-Path $root 'src') -PathType Container)) {
    throw "Repository root identity check failed at $root (VERSION or src\ missing)."
}
$runNonce = [guid]::NewGuid().ToString('N')
if ([string]::IsNullOrWhiteSpace($CandidatePath)) { $CandidatePath = Join-Path $root ".build-work\release-validation-$runNonce.xlsm" }
if ([string]::IsNullOrWhiteSpace($BuildReceiptPath)) { $BuildReceiptPath = Join-Path $root ".build-work\release-validation-$runNonce.xlsm.build-receipt.json" }
if ([string]::IsNullOrWhiteSpace($GateReceiptPath)) { $GateReceiptPath = Join-Path $root ".build-work\release-validation-$runNonce.gate.json" }
if ([string]::IsNullOrWhiteSpace($PerformanceReceiptPath)) { $PerformanceReceiptPath = Join-Path $root ".build-work\release-validation-$runNonce.performance.json" }
if ([string]::IsNullOrWhiteSpace($SummaryReceiptPath)) { $SummaryReceiptPath = Join-Path $root ".build-work\release-validation-$runNonce.summary.json" }
$workspaceRoot = $root
$paths = @($CandidatePath, $BuildReceiptPath, $GateReceiptPath, $PerformanceReceiptPath, $SummaryReceiptPath)
foreach ($path in $paths) {
    $resolved = [IO.Path]::GetFullPath($path)
    if ($resolved -eq $workspaceRoot -or -not $resolved.StartsWith($workspaceRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Output path is outside repository root: $resolved"
    }
}

$pwsh = (Get-Command pwsh -ErrorAction Stop).Source
$buildScript = Join-Path $PSScriptRoot 'build-xlsm.ps1'
$testScript = Join-Path $PSScriptRoot 'test-xlsm.ps1'
$performanceScript = Join-Path $PSScriptRoot 'measure-performance.ps1'
$testWorkRoot = Join-Path $root ".test-work\release-validation-gate-$runNonce"

function Remove-EmptyOwnedDirectory {
    param([string] $Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $resolved = [IO.Path]::GetFullPath($Path)
    if ($resolved -eq $workspaceRoot -or -not $resolved.StartsWith($workspaceRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing cleanup outside workspace root: $resolved"
    }
    $item = Get-Item -LiteralPath $resolved -Force -ErrorAction Stop
    if (-not $item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Refusing cleanup of non-directory or reparse point: $resolved"
    }
    $children = @(Get-ChildItem -LiteralPath $resolved -Force -ErrorAction Stop)
    if ($children.Count -ne 0) { throw "Refusing cleanup of non-empty directory: $resolved" }
    Remove-Item -LiteralPath $resolved -Force -ErrorAction Stop
}

& $pwsh -NoProfile -File $buildScript `
    -CandidateOnly `
    -Profile Test `
    -OutputPath $CandidatePath `
    -CandidateBuildReceiptPath $BuildReceiptPath
if ($LASTEXITCODE -ne 0) { throw "Candidate build failed: exit $LASTEXITCODE" }

& $pwsh -NoProfile -File $testScript `
    -WorkbookPath $CandidatePath `
    -CandidateBuildReceiptPath $BuildReceiptPath `
    -TestWorkRoot $testWorkRoot `
    -GateReceiptPath $GateReceiptPath
$gateExit = $LASTEXITCODE
Remove-EmptyOwnedDirectory $testWorkRoot
if ($gateExit -ne 0) { throw "Required gate failed: exit $gateExit" }

& $pwsh -NoProfile -File $performanceScript `
    -WorkbookPath $CandidatePath `
    -ReceiptPath $PerformanceReceiptPath
if ($LASTEXITCODE -ne 0) { throw "Phase 7 performance failed: exit $LASTEXITCODE" }

$candidateHash = (Get-FileHash -LiteralPath $CandidatePath -Algorithm SHA256).Hash
$buildReceiptHash = (Get-FileHash -LiteralPath $BuildReceiptPath -Algorithm SHA256).Hash
$gateReceiptHash = (Get-FileHash -LiteralPath $GateReceiptPath -Algorithm SHA256).Hash
$performanceReceiptHash = (Get-FileHash -LiteralPath $PerformanceReceiptPath -Algorithm SHA256).Hash
$buildReceipt = Get-Content -LiteralPath $BuildReceiptPath -Raw -Encoding UTF8 | ConvertFrom-Json
$gateReceipt = Get-Content -LiteralPath $GateReceiptPath -Raw -Encoding UTF8 | ConvertFrom-Json
$performanceReceipt = Get-Content -LiteralPath $PerformanceReceiptPath -Raw -Encoding UTF8 | ConvertFrom-Json

if ($buildReceipt.candidateSha256 -ne $candidateHash) { throw 'Build receipt candidate hash mismatch.' }
if ($gateReceipt.candidateSha256 -ne $candidateHash -or -not $gateReceipt.allPassed) { throw 'Gate receipt is not green for this candidate.' }
if ($performanceReceipt.candidateSha256 -ne $candidateHash -or $performanceReceipt.status -ne 'PASS') { throw 'Performance receipt is not green for this candidate.' }

$summary = [ordered]@{
    schemaVersion = 1
    generatedAtUtc = [DateTime]::UtcNow.ToString('o')
    status = 'PASS'
    candidatePath = [IO.Path]::GetRelativePath($root, [IO.Path]::GetFullPath($CandidatePath))
    candidateSha256 = $candidateHash
    buildReceiptPath = [IO.Path]::GetRelativePath($root, [IO.Path]::GetFullPath($BuildReceiptPath))
    buildReceiptSha256 = $buildReceiptHash
    gateReceiptPath = [IO.Path]::GetRelativePath($root, [IO.Path]::GetFullPath($GateReceiptPath))
    gateReceiptSha256 = $gateReceiptHash
    performanceReceiptPath = [IO.Path]::GetRelativePath($root, [IO.Path]::GetFullPath($PerformanceReceiptPath))
    performanceReceiptSha256 = $performanceReceiptHash
    requiredCases = $gateReceipt.requiredCaseIds.Count
    performanceResults = $performanceReceipt.results.Count
    sequenceResults = $performanceReceipt.sequenceResults.Count
    overLimitPass = $performanceReceipt.overLimit.pass
    promotion = 'not_run'
}
$summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $SummaryReceiptPath -Encoding UTF8
Write-Host "summaryReceipt=$([IO.Path]::GetFullPath($SummaryReceiptPath))"
