#Requires -Version 7.0
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $WorkbookPath,
    [string] $CandidateBuildReceiptPath = "$WorkbookPath.build-receipt.json",
    [string] $CandidateGateReceiptPath = ([IO.Path]::ChangeExtension($WorkbookPath, '.gate.json')),
    [string] $ReceiptPath,
    [string] $TestWorkRoot = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$productRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd([IO.Path]::DirectorySeparatorChar)
if (-not (Test-Path -LiteralPath (Join-Path $productRoot 'VERSION') -PathType Leaf) -or -not (Test-Path -LiteralPath (Join-Path $productRoot 'src') -PathType Container)) {
    throw "Repository root identity check failed at $productRoot (VERSION or src\ missing)."
}
if ([string]::IsNullOrWhiteSpace($ReceiptPath)) { $ReceiptPath = Join-Path $productRoot '.build-work\recycle-performance-receipt.json' }
$workspaceRoot = $productRoot
$performanceParent = [IO.Path]::GetFullPath((Join-Path $productRoot '.test-work')).TrimEnd('\')
$requiredPrefix = [IO.Path]::GetFullPath((Join-Path $performanceParent 'recycle-performance-'))
if ([string]::IsNullOrWhiteSpace($TestWorkRoot)) {
    $TestWorkRoot = Join-Path $performanceParent ("recycle-performance-{0}" -f [guid]::NewGuid().ToString('N'))
}

function Resolve-WorkspacePath {
    param([Parameter(Mandatory)][string] $Path)
    $full = [IO.Path]::GetFullPath($Path)
    if ($full -eq $workspaceRoot -or -not $full.StartsWith($workspaceRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path is outside workspace: $full"
    }
    return $full
}

function Assert-PlainPathChain {
    param([Parameter(Mandatory)][string] $Path, [Parameter(Mandatory)][string] $StopAt)
    $current = [IO.Path]::GetFullPath($Path).TrimEnd('\')
    $stop = [IO.Path]::GetFullPath($StopAt).TrimEnd('\')
    while ($current.Length -ge $stop.Length) {
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Reparse path rejected: $current" }
        }
        if ($current -ieq $stop) { return }
        $next = [IO.Path]::GetDirectoryName($current)
        if ([string]::IsNullOrWhiteSpace($next) -or $next -eq $current) { break }
        $current = $next.TrimEnd('\')
    }
    throw "Path chain escaped expected root: $Path"
}

function ConvertFrom-ProbeResult {
    param([Parameter(Mandatory)][string] $Value)
    $parts = @($Value -split '\|')
    $result = [ordered]@{ raw = $Value; status = $parts[0] }
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

function Restore-RecycleFixtures {
    param([Parameter(Mandatory)][string[]] $Sources)
    $missing = @($Sources | Where-Object { -not (Test-Path -LiteralPath $_) })
    if ($missing.Count -eq 0) { return [pscustomobject]@{ matched = 0; restored = 0; elapsedMs = 0 } }
    $targets = @{}
    foreach ($source in $missing) { $targets[(Normalize-FixturePath $source)] = $source }
    $shell = $null; $bin = $null; $items = @(); $matched = 0
    $timer = [Diagnostics.Stopwatch]::StartNew()
    try {
        $shell = New-Object -ComObject Shell.Application
        $bin = $shell.Namespace(10)
        if ($null -eq $bin) { throw 'Recycle Bin namespace unavailable while restoring performance fixtures.' }
        $items = @($bin.Items())
        foreach ($item in $items) {
            $deletedFrom = [string]$item.ExtendedProperty('System.Recycle.DeletedFrom')
            if ([string]::IsNullOrWhiteSpace($deletedFrom)) { continue }
            $deletedNormalized = $null
            try { $deletedNormalized = Normalize-FixturePath $deletedFrom } catch { continue }
            $name = [string]$item.Name
            $candidate = $deletedNormalized
            if (-not $targets.ContainsKey($candidate)) {
                try { $candidate = Normalize-FixturePath (Join-Path $deletedFrom $name) } catch { continue }
            }
            if (-not $targets.ContainsKey($candidate)) { continue }
            $matched++
            $verb = @($item.Verbs() | Where-Object { $_.Name -match '(?i)restore|復元|元に戻す' } | Select-Object -First 1)
            if ($verb.Count -gt 0) { $verb[0].DoIt() } else { $item.InvokeVerb('RESTORE') }
        }
        if ($matched -ne $missing.Count) { throw "Recycle fixture restore match count mismatch: expected=$($missing.Count) actual=$matched" }
        for ($attempt = 1; $attempt -le 120; $attempt++) {
            $remaining = @($missing | Where-Object { -not (Test-Path -LiteralPath $_) })
            if ($remaining.Count -eq 0) { break }
            Start-Sleep -Milliseconds 500
        }
        $remaining = @($missing | Where-Object { -not (Test-Path -LiteralPath $_) })
        if ($remaining.Count -ne 0) { throw "Recycle fixture restore incomplete: $($remaining.Count) item(s) remain." }
        $timer.Stop()
        return [pscustomobject]@{ matched = $matched; restored = $missing.Count; elapsedMs = [math]::Round($timer.Elapsed.TotalMilliseconds, 2) }
    } finally {
        foreach ($item in $items) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) } catch { } }
        if ($null -ne $bin) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($bin) } catch { } }
        if ($null -ne $shell) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell) } catch { } }
    }
}

function Get-SafeEstimateSeconds {
    param([double] $ElapsedMs)
    $seconds = [math]::Max(1, $ElapsedMs / 1000 * 1.5)
    return [int]([math]::Ceiling($seconds / 5) * 5)
}

$WorkbookPath = Resolve-WorkspacePath $WorkbookPath
$CandidateBuildReceiptPath = Resolve-WorkspacePath $CandidateBuildReceiptPath
$CandidateGateReceiptPath = Resolve-WorkspacePath $CandidateGateReceiptPath
$ReceiptPath = Resolve-WorkspacePath $ReceiptPath
$TestWorkRoot = Resolve-WorkspacePath $TestWorkRoot
if (-not $TestWorkRoot.StartsWith($requiredPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw "TestWorkRoot prefix rejected: $TestWorkRoot" }
if ([IO.Path]::GetDirectoryName($TestWorkRoot) -ine $performanceParent) { throw "TestWorkRoot must be a direct child of $performanceParent" }
Assert-PlainPathChain -Path $performanceParent -StopAt $productRoot
if (Test-Path -LiteralPath $TestWorkRoot) { throw "TestWorkRoot already exists: $TestWorkRoot" }
if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) { throw "Workbook not found: $WorkbookPath" }
if (-not (Test-Path -LiteralPath $CandidateBuildReceiptPath -PathType Leaf)) { throw "Build receipt not found: $CandidateBuildReceiptPath" }
if (-not (Test-Path -LiteralPath $CandidateGateReceiptPath -PathType Leaf)) { throw "Gate receipt not found: $CandidateGateReceiptPath" }

$candidateHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $WorkbookPath).Hash
$buildReceiptHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $CandidateBuildReceiptPath).Hash
$gateReceiptHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $CandidateGateReceiptPath).Hash
$buildReceipt = Get-Content -Raw -LiteralPath $CandidateBuildReceiptPath | ConvertFrom-Json
$gateReceipt = Get-Content -Raw -LiteralPath $CandidateGateReceiptPath | ConvertFrom-Json
if ($buildReceipt.candidateSha256 -ne $candidateHash) { throw 'Build receipt candidate hash mismatch.' }
if ($gateReceipt.candidateSha256 -ne $candidateHash -or -not $gateReceipt.allPassed) { throw 'Gate receipt is not green for this candidate.' }

$excel = $null; $workbook = $null; $allFixtureSources = [Collections.Generic.List[string]]::new()
$deleteResults = @(); $syntheticResults = @(); $actualSnapshot = $null; $finalSnapshot = $null; $capacity = $null
$excelVersion = $null; $excelOperatingSystem = $null
$failure = $null; $cleanup = [ordered]@{ attempted = $false; restored = $false; workRootRemoved = $false }
try {
    $null = New-Item -ItemType Directory -Path $TestWorkRoot -Force
    Assert-PlainPathChain -Path $TestWorkRoot -StopAt $performanceParent
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excelVersion = [string]$excel.Version
    $excelOperatingSystem = [string]$excel.OperatingSystem
    $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $true)

    $excelProcess = Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Sort-Object StartTime -Descending | Select-Object -First 1
    $memoryBefore = if ($null -ne $excelProcess) { [int64]$excelProcess.WorkingSet64 } else { $null }
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $actualSnapshot = ConvertFrom-ProbeResult ([string]$excel.Run('RecycleSnapshotPerformanceTest'))
    $timer.Stop()
    if ($actualSnapshot.status -ne 'pass') { throw "Actual Recycle Bin snapshot failed: $($actualSnapshot.raw)" }
    $actualSnapshot | Add-Member -NotePropertyName elapsedMs -NotePropertyValue ([math]::Round($timer.Elapsed.TotalMilliseconds, 2))
    $actualSnapshot | Add-Member -NotePropertyName excelWorkingSetBeforeBytes -NotePropertyValue $memoryBefore

    foreach ($count in @(10000, 50000)) {
        if ($null -ne $excelProcess) { $excelProcess.Refresh(); $before = [int64]$excelProcess.WorkingSet64 } else { $before = $null }
        $timer = [Diagnostics.Stopwatch]::StartNew()
        $probe = ConvertFrom-ProbeResult ([string]$excel.Run('RecycleSyntheticSnapshotPerformanceTest', $count))
        $timer.Stop()
        if ($probe.status -ne 'pass') { throw "Synthetic snapshot probe failed for $count items: $($probe.raw)" }
        if ($null -ne $excelProcess) { $excelProcess.Refresh(); $after = [int64]$excelProcess.WorkingSet64 } else { $after = $null }
        $syntheticResults += [ordered]@{
            items = $count
            elapsedMs = [math]::Round($timer.Elapsed.TotalMilliseconds, 2)
            excelWorkingSetBeforeBytes = $before
            excelWorkingSetAfterBytes = $after
            excelWorkingSetDeltaBytes = if ($null -ne $before -and $null -ne $after) { $after - $before } else { $null }
            result = $probe.raw
            workloadType = 'synthetic canonicalization and production dictionary-delta core; no Shell ExtendedProperty calls'
        }
        Write-Host ("synthetic snapshot {0}: elapsedMs={1}" -f $count, [math]::Round($timer.Elapsed.TotalMilliseconds, 2))
    }

    $capacityFolder = Join-Path $TestWorkRoot 'capacity'
    $null = New-Item -ItemType Directory -Path $capacityFolder
    $capacitySource = Join-Path $capacityFolder 'probe.dat'
    [IO.File]::WriteAllBytes($capacitySource, [byte[]](1))
    $capacity = ConvertFrom-ProbeResult ([string]$excel.Run('RecycleCapacityBoundaryTest', $capacitySource))
    if ($capacity.status -ne 'pass') { throw "Capacity boundary probe failed: $($capacity.raw)" }

    foreach ($count in @(1, 100, 1000)) {
        $batchFolder = Join-Path $TestWorkRoot ("files-{0}" -f $count)
        $null = New-Item -ItemType Directory -Path $batchFolder
        $sources = [Collections.Generic.List[string]]::new()
        for ($i = 1; $i -le $count; $i++) {
            $source = Join-Path $batchFolder ("item-{0:D4}.dat" -f $i)
            [IO.File]::WriteAllBytes($source, [byte[]](1))
            $sources.Add($source); $allFixtureSources.Add($source)
        }
        $timer = [Diagnostics.Stopwatch]::StartNew()
        $probe = ConvertFrom-ProbeResult ([string]$excel.Run('RecycleDirectoryFilesPerformanceTest', $batchFolder))
        $timer.Stop()
        if ($probe.status -ne 'pass') { throw "Delete phase failed for $count items: $($probe.raw)" }
        if ([int]$probe.snapshotCalls -ne 2) { throw "Delete phase snapshot call count must be 2 for $count items: $($probe.raw)" }
        $remainingSources = @($sources | Where-Object { Test-Path -LiteralPath $_ })
        if ($remainingSources.Count -ne 0) { throw "Delete phase left $($remainingSources.Count) source item(s) for count=$count" }
        $restore = Restore-RecycleFixtures -Sources @($sources)
        $deleteResults += [ordered]@{
            items = $count
            deleteElapsedMs = [math]::Round($timer.Elapsed.TotalMilliseconds, 2)
            restoreElapsedMs = $restore.elapsedMs
            snapshotCalls = [int]$probe.snapshotCalls
            extendedPropertyCalls = [int]$probe.extendedPropertyCalls
            successItems = [int]$probe.success
            safetyState = $probe.safety
            safeEstimateSeconds = Get-SafeEstimateSeconds $timer.Elapsed.TotalMilliseconds
            sourceItemsRestored = $restore.restored
            result = $probe.raw
        }
        Write-Host ("delete phase {0}: deleteMs={1} restoreMs={2}" -f $count, [math]::Round($timer.Elapsed.TotalMilliseconds, 2), $restore.elapsedMs)
    }

    $finalSnapshot = ConvertFrom-ProbeResult ([string]$excel.Run('RecycleSnapshotPerformanceTest'))
    if ($finalSnapshot.status -ne 'pass') { throw "Final Recycle Bin snapshot failed: $($finalSnapshot.raw)" }
    if ($finalSnapshot.items -ne $actualSnapshot.items -or $finalSnapshot.fingerprint -ne $actualSnapshot.fingerprint) {
        throw "Recycle Bin changed outside the restored fixtures during measurement: before=$($actualSnapshot.raw) after=$($finalSnapshot.raw)"
    }
} catch {
    $failure = $_.Exception.Message
} finally {
    $cleanup.attempted = $true
    $missing = @($allFixtureSources | Where-Object { -not (Test-Path -LiteralPath $_) })
    if ($missing.Count -gt 0) {
        try { $null = Restore-RecycleFixtures -Sources $missing } catch { if ($null -eq $failure) { $failure = $_.Exception.Message } }
    }
    $cleanup.restored = (@($allFixtureSources | Where-Object { -not (Test-Path -LiteralPath $_) }).Count -eq 0)
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch { }; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
    if ($null -ne $excel) { try { $excel.Quit() } catch { }; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($cleanup.restored -and (Test-Path -LiteralPath $TestWorkRoot)) {
        Assert-PlainPathChain -Path $TestWorkRoot -StopAt $performanceParent
        Remove-Item -LiteralPath $TestWorkRoot -Recurse -Force
    }
    $cleanup.workRootRemoved = -not (Test-Path -LiteralPath $TestWorkRoot)
}

$receipt = [ordered]@{
    schemaVersion = 1
    generatedAtUtc = [DateTime]::UtcNow.ToString('o')
    status = if ($null -eq $failure) { 'PASS' } else { 'FAIL' }
    candidatePath = [IO.Path]::GetRelativePath($productRoot, $WorkbookPath)
    candidateSha256 = $candidateHash
    candidateBuildReceiptPath = [IO.Path]::GetRelativePath($productRoot, $CandidateBuildReceiptPath)
    candidateBuildReceiptSha256 = $buildReceiptHash
    candidateGateReceiptPath = [IO.Path]::GetRelativePath($productRoot, $CandidateGateReceiptPath)
    candidateGateReceiptSha256 = $gateReceiptHash
    environment = [ordered]@{
        os = [Environment]::OSVersion.VersionString
        powershell = $PSVersionTable.PSVersion.ToString()
        excelVersion = $excelVersion
        excelOperatingSystem = $excelOperatingSystem
    }
    actualRecycleBinSnapshot = $actualSnapshot
    syntheticSnapshotResults = $syntheticResults
    capacityBoundary = $capacity
    fileDeletePhaseResults = $deleteResults
    deleteEstimateRule = 'ceil(max(1 second, measured delete time x 1.5) to the next 5 seconds); only measured counts are estimable'
    deleteCountBands = [ordered]@{
        '1' = if ($deleteResults.Count -ge 1) { "$($deleteResults[0].safeEstimateSeconds) seconds" } else { '見積不能' }
        '100' = if ($deleteResults.Count -ge 2) { "$($deleteResults[1].safeEstimateSeconds) seconds" } else { '見積不能' }
        '1000' = if ($deleteResults.Count -ge 3) { "$($deleteResults[2].safeEstimateSeconds) seconds" } else { '見積不能' }
        '1001-10000' = '見積不能'
    }
    finalRecycleBinSnapshot = $finalSnapshot
    cleanup = $cleanup
    assertions = [ordered]@{
        realShellSnapshotMeasuredAtCurrentItemCount = $true
        synthetic10000And50000AreNotClaimedAsRealRecycleBinItems = $true
        oneSnapshotBeforeAndAfterEachDeletePhase = $true
        noPerTargetRecycleBinReenumeration = $true
        capacityBelowAndEqualAcceptedAboveRejected = $true
        aggregateCapacityWarning = $true
        allWorkspaceFixturesRestored = $cleanup.restored
        workRootRemoved = $cleanup.workRootRemoved
    }
    notMeasured = @(
        'real Recycle Bin with 10000 or 50000 pre-existing items',
        'empty-folder depth-group phases',
        'file-delete counts 1001 through 10000'
    )
    failure = $failure
    promotion = 'not_run'
}
$receipt | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $ReceiptPath -Encoding UTF8
Write-Host "receipt=$ReceiptPath"
if ($null -ne $failure) { throw $failure }
