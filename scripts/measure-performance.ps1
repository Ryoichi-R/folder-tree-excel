#Requires -Version 7.0
<##
.SYNOPSIS
    candidateの走査・変更一覧生成をリポジトリ内fixtureで測定する。

.DESCRIPTION
    実ファイル操作は行わず、RunScan、変更一覧生成、連番previewを測定する。
    10,000行、100,000行、100,001行相当のfixtureを作成し、同一candidate・同一production coreで
    行数、経過時間、Excel working set、上限超過時の部分一覧不生成をreceiptへ保存する。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $WorkbookPath,
    [string] $WorkRoot,
    [string] $ReceiptPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd([IO.Path]::DirectorySeparatorChar)
if (-not (Test-Path -LiteralPath (Join-Path $root 'VERSION') -PathType Leaf) -or -not (Test-Path -LiteralPath (Join-Path $root 'src') -PathType Container)) {
    throw "Repository root identity check failed at $root (VERSION or src\ missing)."
}
if ([string]::IsNullOrWhiteSpace($WorkRoot)) { $WorkRoot = Join-Path $root ('.test-work\performance-' + [guid]::NewGuid().ToString('N')) }
if ([string]::IsNullOrWhiteSpace($ReceiptPath)) { $ReceiptPath = Join-Path $root '.build-work\performance-receipt.json' }
$workspaceRoot = $root
$workFull = [IO.Path]::GetFullPath($WorkRoot)
$receiptFull = [IO.Path]::GetFullPath($ReceiptPath)
foreach ($path in @($workFull, $receiptFull)) {
    if ($path -eq $workspaceRoot -or -not $path.StartsWith($workspaceRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path is outside repository root: $path"
    }
}
if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) { throw "Workbook not found: $WorkbookPath" }

function Get-ExcelMemoryBytes {
    $processes = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
    if ($processes.Count -eq 0) { return 0L }
    return [int64](($processes | Measure-Object -Property WorkingSet64 -Sum).Sum)
}

function New-EmptyFileFixture {
    param([string] $Root, [int] $FileCount, [int] $DirectoryCount)
    $timer = [Diagnostics.Stopwatch]::StartNew()
    New-Item -ItemType Directory -Path $Root -Force | Out-Null
    $filesPerDirectory = [math]::Floor($FileCount / $DirectoryCount)
    $remaining = $FileCount
    for ($d = 0; $d -lt $DirectoryCount; $d++) {
        $directory = Join-Path $Root ('d{0:D4}' -f $d)
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
        $count = [math]::Min($filesPerDirectory, $remaining)
        if ($d -eq $DirectoryCount - 1) { $count = $remaining }
        for ($f = 0; $f -lt $count; $f++) {
            $file = Join-Path $directory ('f{0:D5}.txt' -f $f)
            $stream = [IO.File]::Create($file)
            $stream.Dispose()
        }
        $remaining -= $count
    }
    $timer.Stop()
    [pscustomobject]@{ fileCount = $FileCount; directoryCount = $DirectoryCount; creationMs = $timer.Elapsed.TotalMilliseconds }
}

function Get-TreeRowCount {
    param($Workbook)
    $sheet = $Workbook.Worksheets.Item('ツリー')
    $lastRow = $sheet.Cells.Item($sheet.Rows.Count, 1).End(-4162).Row
    return [int][math]::Max(0, $lastRow - 4 + 1)
}

function Get-OperationDataRowCount {
    param($Workbook)
    $sheet = $Workbook.Worksheets.Item('変更操作')
    $lastRow = $sheet.Cells.Item($sheet.Rows.Count, 1).End(-4162).Row
    return [int][math]::Max(0, $lastRow - 9 + 1)
}

$excel = $null
$workbook = $null
$results = @()
$sequenceResults = @()
$overLimitResult = $null
$fixtureRoots = @()
$status = 'PASS'
$failure = $null
try {
    New-Item -ItemType Directory -Path $workFull -Force | Out-Null
    $smallRoot = Join-Path $workFull 'rows-10000'
    $largeRoot = Join-Path $workFull 'rows-100000'
    $fixtureRoots += $smallRoot
    $fixtureRoots += $largeRoot
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open([IO.Path]::GetFullPath($WorkbookPath), 0, $true)
    $smallFixture = New-EmptyFileFixture $smallRoot 9900 99
    $largeFixture = New-EmptyFileFixture $largeRoot 99899 100

    foreach ($fixture in @(
            [pscustomobject]@{ name = '10000'; root = $smallRoot; expectedRows = 10000; create = $smallFixture },
            [pscustomobject]@{ name = '100000'; root = $largeRoot; expectedRows = 100000; create = $largeFixture }
        )) {
        $beforeMemory = Get-ExcelMemoryBytes
        $scanTimer = [Diagnostics.Stopwatch]::StartNew()
        $scanSummary = [string]$excel.Run('RunScan', $fixture.root, 0, $true, $false)
        $scanTimer.Stop()
        $rows = Get-TreeRowCount $workbook
        $draftTimer = [Diagnostics.Stopwatch]::StartNew()
        $draftCreated = [bool]$excel.Run('WorkbookIoCreateOperationDraft', $fixture.root)
        $draftTimer.Stop()
        $afterMemory = Get-ExcelMemoryBytes
        if ($rows -ne $fixture.expectedRows) { throw "Unexpected row count for $($fixture.name): $rows" }
        if (-not $draftCreated) { throw "Operation draft was not created for $($fixture.name) rows." }
        $results += [ordered]@{
            fixture = $fixture.name
            expectedRows = $fixture.expectedRows
            actualRows = $rows
            fixtureCreationMs = [math]::Round($fixture.create.creationMs, 2)
            scanMs = [math]::Round($scanTimer.Elapsed.TotalMilliseconds, 2)
            draftMs = [math]::Round($draftTimer.Elapsed.TotalMilliseconds, 2)
            excelWorkingSetBeforeBytes = $beforeMemory
            excelWorkingSetAfterBytes = $afterMemory
            excelWorkingSetDeltaBytes = ($afterMemory - $beforeMemory)
            scanSummary = $scanSummary
            draftCreated = $draftCreated
        }
        Write-Host ("{0}: rows={1} scanMs={2} draftMs={3} memoryDeltaBytes={4}" -f $fixture.name, $rows, [math]::Round($scanTimer.Elapsed.TotalMilliseconds, 2), [math]::Round($draftTimer.Elapsed.TotalMilliseconds, 2), ($afterMemory - $beforeMemory))

        if ($fixture.name -eq '10000') {
            $operationSheet = $workbook.Worksheets.Item('変更操作')
            $firstDataRow = 9
            $lastDataRow = $firstDataRow + 10000 - 1
            foreach ($selectionCount in @(100, 1000, 10000)) {
                Write-Host ("sequence {0}: preparing worksheet" -f $selectionCount)
                $sequenceDraftCreated = [bool]$excel.Run('WorkbookIoCreateOperationDraftTest', $smallRoot)
                if (-not $sequenceDraftCreated) { throw "Sequence draft creation failed for $selectionCount rows." }
                $operationSheet.Unprotect()
                if ($selectionCount -lt 10000) {
                    $clearStart = $firstDataRow + $selectionCount
                    $operationSheet.Range("A${clearStart}:V${lastDataRow}").ClearContents()
                }
                $selectionLastRow = $firstDataRow + $selectionCount - 1
                $operationSheet.Range("P${firstDataRow}:P${selectionLastRow}").Value2 = 'ファイル'
                $sequenceTimer = [Diagnostics.Stopwatch]::StartNew()
                $sequenceResult = [string]$excel.Run('WorkbookIoSetSequencePreviewTest', '先頭', 0)
                $sequenceTimer.Stop()
                if ($sequenceResult -ne "pass|selected=$selectionCount") {
                    throw "Unexpected sequence result for $selectionCount rows: $sequenceResult"
                }
                $sequenceResults += [ordered]@{
                    selectedRows = $selectionCount
                    elapsedMs = [math]::Round($sequenceTimer.Elapsed.TotalMilliseconds, 2)
                    result = $sequenceResult
                }
                Write-Host ("sequence {0}: elapsedMs={1}" -f $selectionCount, [math]::Round($sequenceTimer.Elapsed.TotalMilliseconds, 2))
            }
        }
    }

    $overflowFile = Join-Path $largeRoot 'd0000\overflow.txt'
    $overflowStream = [IO.File]::Create($overflowFile)
    $overflowStream.Dispose()
    $overflowScanTimer = [Diagnostics.Stopwatch]::StartNew()
    $overflowSummary = [string]$excel.Run('RunScan', $largeRoot, 0, $true, $false)
    $overflowScanTimer.Stop()
    $overflowRows = Get-TreeRowCount $workbook
    $overflowDraftTimer = [Diagnostics.Stopwatch]::StartNew()
    $overflowDraftCreated = [bool]$excel.Run('WorkbookIoCreateOperationDraftTest', $largeRoot)
    $overflowDraftTimer.Stop()
    $overflowDataRows = Get-OperationDataRowCount $workbook
    $overflowState = [string]$workbook.Worksheets.Item('変更操作').Range('B5').Value2
    if ($overflowRows -ne 100001) { throw "Unexpected over-limit row count: $overflowRows" }
    if ($overflowDraftCreated) { throw 'Over-limit operation draft unexpectedly succeeded.' }
    if ($overflowDataRows -ne 0) { throw "Over-limit operation draft left partial rows: $overflowDataRows" }
    if ($overflowState -ne '一覧上限超過') { throw "Unexpected over-limit state: $overflowState" }
    $overLimitResult = [ordered]@{
        expectedRows = 100001
        actualRows = $overflowRows
        scanMs = [math]::Round($overflowScanTimer.Elapsed.TotalMilliseconds, 2)
        draftRejectMs = [math]::Round($overflowDraftTimer.Elapsed.TotalMilliseconds, 2)
        draftCreated = $overflowDraftCreated
        operationDataRowsAfterReject = $overflowDataRows
        operationState = $overflowState
        scanSummary = $overflowSummary
        pass = $true
    }
    Write-Host ("over-limit: rows={0} draftCreated={1} operationDataRows={2}" -f $overflowRows, $overflowDraftCreated, $overflowDataRows)
}
catch {
    $failure = $_.Exception.Message
    $status = if ($failure -like '*80070520*') { 'FAILED_ENVIRONMENT' } else { 'FAILED' }
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        $workbook = $null
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        $excel = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    foreach ($fixtureRoot in $fixtureRoots) {
        if (Test-Path -LiteralPath $fixtureRoot) {
            $resolvedFixture = [IO.Path]::GetFullPath($fixtureRoot)
            if ($resolvedFixture -eq $workFull -or -not $resolvedFixture.StartsWith($workFull + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Refusing cleanup outside performance work root: $resolvedFixture"
            }
            $fixtureItem = Get-Item -LiteralPath $resolvedFixture -Force -ErrorAction Stop
            if (-not $fixtureItem.PSIsContainer -or ($fixtureItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "Refusing cleanup of non-directory or reparse point: $resolvedFixture"
            }
            $reparseEntries = @(Get-ChildItem -LiteralPath $resolvedFixture -Recurse -Force -ErrorAction Stop | Where-Object { ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 })
            if ($reparseEntries.Count -ne 0) { throw "Refusing cleanup because reparse points exist: $resolvedFixture" }
            Remove-Item -LiteralPath $resolvedFixture -Recurse -Force -ErrorAction Stop
        }
    }
    if (Test-Path -LiteralPath $workFull) {
        $workItem = Get-Item -LiteralPath $workFull -Force -ErrorAction Stop
        if (-not $workItem.PSIsContainer -or ($workItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Refusing cleanup of performance work root: $workFull"
        }
        $remainingEntries = @(Get-ChildItem -LiteralPath $workFull -Force -ErrorAction Stop)
        if ($remainingEntries.Count -ne 0) { throw "Performance work root is not empty: $workFull" }
        Remove-Item -LiteralPath $workFull -Force -ErrorAction Stop
    }
}

$candidateHash = (Get-FileHash -LiteralPath $WorkbookPath -Algorithm SHA256).Hash
$receipt = [ordered]@{
    schemaVersion = 1
    generatedAtUtc = [DateTime]::UtcNow.ToString('o')
    candidatePath = [IO.Path]::GetRelativePath($root, [IO.Path]::GetFullPath($WorkbookPath))
    candidateSha256 = $candidateHash
    test = 'RunScan + operation draft + sequence preview + over-limit rejection; no file operations'
    results = $results
    sequenceResults = $sequenceResults
    overLimit = $overLimitResult
    status = $status
    failure = $failure
}
New-Item -ItemType Directory -Path (Split-Path -Parent $receiptFull) -Force | Out-Null
$receipt | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $receiptFull -Encoding UTF8
Write-Host "receipt=$receiptFull"
if ($null -ne $failure) { throw $failure }
