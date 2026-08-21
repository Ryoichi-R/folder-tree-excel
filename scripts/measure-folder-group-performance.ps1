#Requires -Version 7.0
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $WorkbookPath,
    [string] $CandidateGateReceiptPath = ([IO.Path]::ChangeExtension($WorkbookPath, '.gate.json')),
    [string] $ReceiptPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$productRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd([IO.Path]::DirectorySeparatorChar)
if (-not (Test-Path -LiteralPath (Join-Path $productRoot 'VERSION') -PathType Leaf) -or -not (Test-Path -LiteralPath (Join-Path $productRoot 'src') -PathType Container)) {
    throw "Repository root identity check failed at $productRoot (VERSION or src\ missing)."
}
if ([string]::IsNullOrWhiteSpace($ReceiptPath)) { $ReceiptPath = Join-Path $productRoot '.build-work\folder-group-performance-receipt.json' }
$workspaceRoot = $productRoot
$testParent = [IO.Path]::GetFullPath((Join-Path $productRoot '.test-work')).TrimEnd('\')
$testRoot = Join-Path $testParent ("recycle-performance-folder-groups-{0}" -f [guid]::NewGuid().ToString('N').Substring(0, 8))

function Resolve-WorkspacePath {
    param([Parameter(Mandatory)][string] $Path)
    $full = [IO.Path]::GetFullPath($Path)
    if ($full -eq $workspaceRoot -or -not $full.StartsWith($workspaceRoot + '\', [StringComparison]::OrdinalIgnoreCase)) { throw "Path outside workspace: $full" }
    return $full
}

function Assert-SafeTestRoot {
    param([Parameter(Mandatory)][string] $Path)
    $full = Resolve-WorkspacePath $Path
    $prefix = [IO.Path]::GetFullPath((Join-Path $testParent 'recycle-performance-folder-groups-'))
    if (-not $full.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetDirectoryName($full) -ine $testParent) {
        throw "Folder-group test root rejected: $full"
    }
    foreach ($candidate in @($productRoot, $testParent, $full)) {
        if (Test-Path -LiteralPath $candidate) {
            $item = Get-Item -LiteralPath $candidate -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Reparse path rejected: $candidate" }
        }
    }
    return $full
}

function ConvertFrom-ProbeResult {
    param([Parameter(Mandatory)][string] $Value)
    $parts = @($Value -split '\|'); $result = [ordered]@{ raw = $Value; status = $parts[0] }
    foreach ($part in $parts | Select-Object -Skip 1) {
        $separator = $part.IndexOf('=')
        if ($separator -gt 0) { $result[$part.Substring(0, $separator)] = $part.Substring($separator + 1) }
    }
    return [pscustomobject]$result
}

function Normalize-FixturePath {
    param([Parameter(Mandatory)][string] $Path)
    return [IO.Path]::GetFullPath($Path).TrimEnd('\').ToUpperInvariant()
}

function Restore-RecycleFolderFixtures {
    param([Parameter(Mandatory)][string[]] $Sources)
    $missing = @($Sources | Where-Object { -not (Test-Path -LiteralPath $_) })
    if ($missing.Count -eq 0) { return 0 }
    $targets = @{}; foreach ($source in $missing) { $targets[(Normalize-FixturePath $source)] = $source }
    $shell = $null; $bin = $null; $items = @(); $matched = 0
    try {
        $shell = New-Object -ComObject Shell.Application; $bin = $shell.Namespace(10)
        if ($null -eq $bin) { throw 'Recycle Bin namespace unavailable.' }
        $items = @($bin.Items())
        foreach ($item in $items) {
            $deletedFrom = [string]$item.ExtendedProperty('System.Recycle.DeletedFrom')
            if ([string]::IsNullOrWhiteSpace($deletedFrom)) { continue }
            try { $candidate = Normalize-FixturePath $deletedFrom } catch { continue }
            if (-not $targets.ContainsKey($candidate)) {
                try { $candidate = Normalize-FixturePath (Join-Path $deletedFrom ([string]$item.Name)) } catch { continue }
            }
            if (-not $targets.ContainsKey($candidate)) { continue }
            $matched++
            $verb = @($item.Verbs() | Where-Object { $_.Name -match '(?i)restore|復元|元に戻す' } | Select-Object -First 1)
            if ($verb.Count -gt 0) { $verb[0].DoIt() } else { $item.InvokeVerb('RESTORE') }
        }
        if ($matched -ne $missing.Count) { throw "Restore match mismatch: expected=$($missing.Count) actual=$matched" }
        for ($attempt = 1; $attempt -le 60; $attempt++) {
            if (@($missing | Where-Object { -not (Test-Path -LiteralPath $_) }).Count -eq 0) { return $matched }
            Start-Sleep -Milliseconds 500
        }
        throw 'Folder fixture restore timed out.'
    } finally {
        foreach ($item in $items) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) } catch { } }
        if ($null -ne $bin) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($bin) } catch { } }
        if ($null -ne $shell) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell) } catch { } }
    }
}

function New-FolderDepthFixture {
    param([Parameter(Mandatory)][string] $Root, [Parameter(Mandatory)][int] $GroupCount)
    $targets = [Collections.Generic.List[string]]::new()
    for ($depth = 1; $depth -le $GroupCount; $depth++) {
        $current = $Root
        for ($level = 1; $level -lt $depth; $level++) {
            $current = Join-Path $current ("h{0}-{1}" -f $depth, $level)
            $null = New-Item -ItemType Directory -Path $current -Force
        }
        $target = Join-Path $current ("target-d{0}" -f $depth)
        $null = New-Item -ItemType Directory -Path $target
        $targets.Add($target)
    }
    return @($targets)
}

$WorkbookPath = Resolve-WorkspacePath $WorkbookPath
$CandidateGateReceiptPath = Resolve-WorkspacePath $CandidateGateReceiptPath
$ReceiptPath = Resolve-WorkspacePath $ReceiptPath
$testRoot = Assert-SafeTestRoot $testRoot
if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) { throw "Workbook missing: $WorkbookPath" }
$candidateHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $WorkbookPath).Hash
$gateHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $CandidateGateReceiptPath).Hash
$gate = Get-Content -Raw -LiteralPath $CandidateGateReceiptPath | ConvertFrom-Json
if (-not $gate.allPassed -or $gate.candidateSha256 -ne $candidateHash) { throw 'Candidate gate mismatch.' }

$excel = $null; $workbook = $null; $allTargets = [Collections.Generic.List[string]]::new(); $results = @(); $failure = $null
$cleanup = [ordered]@{ restored = $false; workRootRemoved = $false }
try {
    $null = New-Item -ItemType Directory -Path $testRoot
    $excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $true)
    $beforeSnapshot = ConvertFrom-ProbeResult ([string]$excel.Run('RecycleSnapshotPerformanceTest'))
    if ($beforeSnapshot.status -ne 'pass') { throw "Initial snapshot failed: $($beforeSnapshot.raw)" }
    foreach ($groupCount in @(1, 3, 8)) {
        $root = Join-Path $testRoot ("groups-{0}" -f $groupCount); $null = New-Item -ItemType Directory -Path $root
        $targets = @(New-FolderDepthFixture -Root $root -GroupCount $groupCount)
        foreach ($target in $targets) { $allTargets.Add($target) }
        $null = $excel.Run('RunScanTest', $root, 0, $true, $false, 0, 2, $false, $false, 0)
        if (-not [bool]$excel.Run('WorkbookIoCreateOperationDraftTest', $root)) { throw "Operation draft failed for groups=$groupCount" }
        $sheet = $workbook.Worksheets.Item('変更操作'); $sheet.Unprotect()
        $lastRow = [int]$sheet.Cells($sheet.Rows.Count, 1).End(-4162).Row
        foreach ($target in $targets) {
            $matchedRow = 0
            for ($row = 9; $row -le $lastRow; $row++) {
                if ([string]$sheet.Cells.Item($row, 17).Value2 -ieq $target) { $matchedRow = $row; break }
            }
            if ($matchedRow -eq 0) { throw "Target row not found: $target" }
            $sheet.Cells.Item($matchedRow, 2).Value2 = '空フォルダをゴミ箱へ'
        }
        $sheet.Protect()
        $preview = [string]$excel.Run('RunOperationPlanTest')
        if (-not $preview.StartsWith('pass|')) { throw "Folder plan rejected for groups=${groupCount}: $preview" }
        $null = $excel.Run('ResetRecyclePerformanceCounters')
        $timer = [Diagnostics.Stopwatch]::StartNew(); $execution = [string]$excel.Run('ExecuteOperationPlanTest'); $timer.Stop()
        if (-not $execution.StartsWith('success:')) { throw "Folder execution failed for groups=${groupCount}: $execution" }
        $counters = ConvertFrom-ProbeResult ("pass|" + [string]$excel.Run('GetRecyclePerformanceCounters'))
        if ([int]$counters.snapshotCalls -ne 2 * $groupCount) { throw "Unexpected snapshot calls for groups=${groupCount}: $($counters.raw)" }
        if (@($targets | Where-Object { Test-Path -LiteralPath $_ }).Count -ne 0) { throw "Folder delete left sources for groups=$groupCount" }
        $restoreTimer = [Diagnostics.Stopwatch]::StartNew(); $restored = Restore-RecycleFolderFixtures -Sources $targets; $restoreTimer.Stop()
        $results += [ordered]@{
            depthGroups = $groupCount
            targetFolders = $targets.Count
            deleteElapsedMs = [math]::Round($timer.Elapsed.TotalMilliseconds, 2)
            restoreElapsedMs = [math]::Round($restoreTimer.Elapsed.TotalMilliseconds, 2)
            snapshotCalls = [int]$counters.snapshotCalls
            expectedSnapshotCalls = 2 * $groupCount
            extendedPropertyCalls = [int]$counters.extendedPropertyCalls
            restoredFolders = $restored
            execution = $execution
        }
        Write-Host ("folder groups {0}: deleteMs={1} restoreMs={2} snapshotCalls={3}" -f $groupCount, [math]::Round($timer.Elapsed.TotalMilliseconds, 2), [math]::Round($restoreTimer.Elapsed.TotalMilliseconds, 2), $counters.snapshotCalls)
    }
    $afterSnapshot = ConvertFrom-ProbeResult ([string]$excel.Run('RecycleSnapshotPerformanceTest'))
    if ($afterSnapshot.status -ne 'pass' -or $afterSnapshot.fingerprint -ne $beforeSnapshot.fingerprint -or $afterSnapshot.items -ne $beforeSnapshot.items) {
        throw "Recycle Bin fingerprint changed: before=$($beforeSnapshot.raw) after=$($afterSnapshot.raw)"
    }
} catch {
    $failure = $_.Exception.Message
} finally {
    $missing = @($allTargets | Where-Object { -not (Test-Path -LiteralPath $_) })
    if ($missing.Count -gt 0) { try { $null = Restore-RecycleFolderFixtures -Sources $missing } catch { if ($null -eq $failure) { $failure = $_.Exception.Message } } }
    $cleanup.restored = (@($allTargets | Where-Object { -not (Test-Path -LiteralPath $_) }).Count -eq 0)
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch { }; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
    if ($null -ne $excel) { try { $excel.Quit() } catch { }; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($cleanup.restored -and (Test-Path -LiteralPath $testRoot)) { $testRoot = Assert-SafeTestRoot $testRoot; Remove-Item -LiteralPath $testRoot -Recurse -Force }
    $cleanup.workRootRemoved = -not (Test-Path -LiteralPath $testRoot)
}

$receipt = [ordered]@{
    schemaVersion = 1
    generatedAtUtc = [DateTime]::UtcNow.ToString('o')
    status = if ($null -eq $failure) { 'PASS' } else { 'FAIL' }
    candidatePath = [IO.Path]::GetRelativePath($productRoot, $WorkbookPath)
    candidateSha256 = $candidateHash
    candidateGateReceiptPath = [IO.Path]::GetRelativePath($productRoot, $CandidateGateReceiptPath)
    candidateGateReceiptSha256 = $gateHash
    results = $results
    assertions = [ordered]@{
        productionScanDraftPreviewExecutePath = $true
        snapshotCallsAreExactlyTwiceDepthGroups = ($null -eq $failure)
        escapeAndUnexecutedContracts = 'covered by required Cases 52 and 63 in the bound candidate gate'
        allWorkspaceFixturesRestored = $cleanup.restored
        recycleBinFingerprintRestored = ($null -eq $failure)
        workRootRemoved = $cleanup.workRootRemoved
    }
    cleanup = $cleanup
    failure = $failure
    promotion = 'not_run'
}
$receipt | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ReceiptPath -Encoding UTF8
Write-Host "receipt=$ReceiptPath"
if ($null -ne $failure) { throw $failure }
