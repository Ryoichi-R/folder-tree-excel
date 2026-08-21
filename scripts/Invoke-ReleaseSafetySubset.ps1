#Requires -Version 7.0
<#
.SYNOPSIS
    release candidateに対し、production UI経路だけを通じて root外move拒否・
    reparse祖先拒否・destination collision拒否の3項目を検証する。

.DESCRIPTION
    Phase 1.2 / 3.2 の release safety subset を構成する3項目を検証する。
    Recycle Bin無効化を伴うfallback拒否の動的検証はOD-11（2026-08-20のowner判断）により
    subsetの対象外であり、この安全特性は modRecycleBin.bas の NoRecycleFiles /
    NukeOnDelete ガード節の静的source reviewだけに依拠する。
    TEST_BUILDのwrapper（RunOperationPlanTest等）は一切使用せず、production の
    PreviewOperationPlan() だけを呼び出す。3項目とも VBA の BuildOperationPlan が
    計画構築段階（事前確認）で拒否するため、実ファイル削除や Recycle Bin操作は
    一切発生しない。

    PreviewOperationPlan() は結果を VBA の MsgBox で通知する。この MsgBox を
    このscriptが起動した専用Excel process ID にscopeしたWin32 window handle
    だけへ限定して閉じる。他のExcel windowや無関係なdialogには一切干渉しない。
    想定外のtitleのdialogは押さずに失敗として報告する。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $CandidatePath,
    [string] $ReleaseSmokeRoot,
    [string] $ReceiptPath,
    [int] $DialogTimeoutSeconds = 25
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd([IO.Path]::DirectorySeparatorChar)
if (-not (Test-Path -LiteralPath (Join-Path $root 'VERSION') -PathType Leaf) -or -not (Test-Path -LiteralPath (Join-Path $root 'src') -PathType Container)) {
    throw "Repository root identity check failed at $root (VERSION or src\ missing)."
}
if (-not (Test-Path -LiteralPath $CandidatePath -PathType Leaf)) { throw "Candidate not found: $CandidatePath" }
$workspaceFull = $root

if ([string]::IsNullOrWhiteSpace($ReleaseSmokeRoot)) { $ReleaseSmokeRoot = Join-Path $root '.test-work' }
$smokeParent = [IO.Path]::GetFullPath($ReleaseSmokeRoot)
if ($smokeParent -eq $workspaceFull -or -not $smokeParent.StartsWith($workspaceFull + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "ReleaseSmokeRoot must be a child of repository root: $smokeParent"
}
$runNonce = [guid]::NewGuid().ToString('N')
$runRoot = Join-Path $smokeParent "release-smoke-$runNonce"
New-Item -ItemType Directory -Path $runRoot -Force | Out-Null
$runItem = Get-Item -LiteralPath $runRoot -Force
if (($runItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Release smoke root cannot be a reparse point: $runRoot" }

if ([string]::IsNullOrWhiteSpace($ReceiptPath)) { $ReceiptPath = Join-Path $smokeParent "release-safety-subset-$runNonce.receipt.json" }
$receiptFull = [IO.Path]::GetFullPath($ReceiptPath)
if (-not $receiptFull.StartsWith($workspaceFull + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw "Receipt path is outside workspace: $receiptFull" }
if ($receiptFull.StartsWith($runRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "ReceiptPath must not be inside the nonce-scoped run root that gets fully removed during cleanup: $receiptFull"
}

$candidateSha256 = (Get-FileHash -LiteralPath $CandidatePath -Algorithm SHA256).Hash
$smokeWorkbookPath = Join-Path $runRoot 'release-smoke-under-test.xlsm'
Copy-Item -LiteralPath $CandidatePath -Destination $smokeWorkbookPath -Force

# ---- Win32 dialog automation: PID-scoped, exact-title-only, never a blind click ----
$nativeSource = @'
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;

namespace ReleaseSafetySubset
{
    public static class NativeDialog
    {
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDlgItem(IntPtr hDlg, int nIDDlgItem);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        private const uint BM_CLICK = 0x00F5;
        private const uint GW_CHILD = 5;
        private const uint GW_HWNDNEXT = 2;

        public static uint GetProcessId(IntPtr hWnd)
        {
            uint pid;
            GetWindowThreadProcessId(hWnd, out pid);
            return pid;
        }

        public static string GetTitle(IntPtr hWnd)
        {
            int len = GetWindowTextLength(hWnd);
            if (len <= 0) return string.Empty;
            var sb = new StringBuilder(len + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        public static string GetClass(IntPtr hWnd)
        {
            var sb = new StringBuilder(256);
            GetClassName(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        // 指定PIDが所有するtop-level・可視・class "#32770"（標準ダイアログ）だけを対象にする。
        public static List<IntPtr> FindTopLevelDialogsByProcess(uint targetPid)
        {
            var found = new List<IntPtr>();
            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
            {
                if (GetProcessId(hWnd) == targetPid && IsWindowVisible(hWnd) && GetClass(hWnd) == "#32770")
                {
                    found.Add(hWnd);
                }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        // 標準MessageBoxのcontrol ID（IDOK=1, IDYES=6）だけをtargetとする。
        // 見つからない場合は最初のButton子windowへfallbackするが、押す対象のhwnd自体は
        // 呼び出し元が既にtitle照合済みのdialog windowに限定されている。
        public static bool ClickButton(IntPtr dialogHwnd, int controlId)
        {
            IntPtr button = GetDlgItem(dialogHwnd, controlId);
            if (button == IntPtr.Zero)
            {
                button = GetWindow(dialogHwnd, GW_CHILD);
                while (button != IntPtr.Zero && GetClass(button) != "Button")
                {
                    button = GetWindow(button, GW_HWNDNEXT);
                }
            }
            if (button == IntPtr.Zero) return false;
            SendMessage(button, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
            return true;
        }
    }
}
'@
Add-Type -TypeDefinition $nativeSource -Language CSharp

# 期待されるdialog sequence（title → 押すcontrol ID）。未知titleは押さない。
$IDOK = 1
$watcherScript = {
    param([uint32] $TargetPid, [System.Collections.Hashtable[]] $ExpectedSequence, [int] $TimeoutSeconds)
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $observed = New-Object System.Collections.Generic.List[string]
    $stepIndex = 0
    while ((Get-Date) -lt $deadline) {
        if ($stepIndex -ge $ExpectedSequence.Count) {
            return [pscustomobject]@{ completed = $true; observedTitles = @($observed); anomalyTitle = $null; timedOut = $false }
        }
        $expected = $ExpectedSequence[$stepIndex]
        $hwnds = [ReleaseSafetySubset.NativeDialog]::FindTopLevelDialogsByProcess($TargetPid)
        foreach ($h in $hwnds) {
            $title = [ReleaseSafetySubset.NativeDialog]::GetTitle($h)
            if ([string]::IsNullOrWhiteSpace($title)) { continue }
            if (-not $observed.Contains($title)) { $observed.Add($title) }
            if ($title -eq $expected.Title) {
                [ReleaseSafetySubset.NativeDialog]::ClickButton($h, [int] $expected.ButtonId) | Out-Null
                $stepIndex++
                Start-Sleep -Milliseconds 200
                break
            } elseif ($title -ne $expected.Title) {
                # 想定外title: 押さずに異常として打ち切る（ハング回避のためprocessは強制終了する）。
                Stop-Process -Id $TargetPid -Force -ErrorAction SilentlyContinue
                return [pscustomobject]@{ completed = $false; observedTitles = @($observed); anomalyTitle = $title; timedOut = $false }
            }
        }
        Start-Sleep -Milliseconds 150
    }
    Stop-Process -Id $TargetPid -Force -ErrorAction SilentlyContinue
    return [pscustomobject]@{ completed = $false; observedTitles = @($observed); anomalyTitle = $null; timedOut = $true }
}

function Get-TreeFingerprint {
    # reparse point配下（例: root逸脱fixtureのjunction先）へは追跡しない。
    # 生成物側の「表示するが追跡しない」方針と一致させ、junction先（%TEMP%等）の
    # 無関係な変化がfingerprintへ混入しないようにする。
    #
    # timestampは含めない。検索indexerやウイルス対策等、application外の要因で
    # 新規作成directoryのLastWriteTimeが実測でサブミリ秒単位で揺れることを確認済みで
    # （内容は不変）、safety検証としてはpath集合とfile内容hashの不変性だけが意味を持つ。
    param([string] $Root)
    $fullRoot = [IO.Path]::GetFullPath($Root)
    $lines = [System.Collections.Generic.List[string]]::new()
    function Walk-TreeFingerprint([string] $Dir, [string] $FullRoot, [System.Collections.Generic.List[string]] $Lines) {
        $entries = @(Get-ChildItem -LiteralPath $Dir -Force | Sort-Object Name)
        foreach ($e in $entries) {
            $relative = [IO.Path]::GetRelativePath($FullRoot, $e.FullName).Replace('\', '/')
            $isReparse = (($e.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)
            if ($e.PSIsContainer) {
                $Lines.Add(('{0}|dir|reparse={1}' -f $relative, $isReparse))
                if (-not $isReparse) { Walk-TreeFingerprint $e.FullName $FullRoot $Lines }
            } else {
                $hash = (Get-FileHash -LiteralPath $e.FullName -Algorithm SHA256).Hash
                $Lines.Add(('{0}|file|{1}|{2}' -f $relative, $e.Length, $hash))
            }
        }
    }
    Walk-TreeFingerprint $fullRoot $fullRoot $lines
    return @($lines | Sort-Object)
}

function New-SafetyFixture {
    param([string] $Parent, [string] $Name)
    $fixtureRoot = Join-Path $Parent $Name
    $null = New-Item -ItemType Directory -Path (Join-Path $fixtureRoot 'src') -Force
    $null = New-Item -ItemType Directory -Path (Join-Path $fixtureRoot 'dst') -Force
    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'src\a.txt'), 'a')
    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'src\b.txt'), 'b')
    return $fixtureRoot
}

$expectRejectTitle = '事前確認で停止しました'

$excel = $null
$wb = $null
$results = [System.Collections.Generic.List[object]]::new()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $wb = $excel.Workbooks.Open($smokeWorkbookPath)
    $excelPid = [ReleaseSafetySubset.NativeDialog]::GetProcessId([IntPtr] $excel.Hwnd)
    if ($excelPid -eq 0) { throw 'Failed to resolve the dedicated Excel process id.' }

    function Invoke-RejectionScenario {
        param(
            [string] $Id,
            [string] $Name,
            [string] $FixtureName,
            [string] $SourceRelative,
            [string] $NewName,
            [string] $MoveRelative,
            [string] $JunctionName
        )
        Write-Host ("=== [{0}] {1} ===" -f $Id, $Name)
        try {
            # fixture作成もこのtry配下に置く。junction作成はDeveloper Mode無効かつ非管理者の
            # 環境では失敗し得るため、外側で作るとreceiptを1件も残さずscriptごと異常終了する。
            $FixtureRoot = New-SafetyFixture $runRoot $FixtureName
            if (-not [string]::IsNullOrWhiteSpace($JunctionName)) {
                $null = New-Item -ItemType Junction -Path (Join-Path $FixtureRoot $JunctionName) -Target ([IO.Path]::GetTempPath()) -Force
            }
            $sourceFull = Join-Path $FixtureRoot $SourceRelative
            $beforeFingerprint = Get-TreeFingerprint $FixtureRoot

            $null = $excel.Run('RunScan', $FixtureRoot, 0, $true, $false)
            if (-not [bool]$excel.Run('WorkbookIoCreateOperationDraft', $FixtureRoot)) {
                throw "[$Id] operation draft creation failed for $FixtureRoot"
            }
            $null = $excel.Run('OperationDraftCreated', $FixtureRoot)

            $opsSheet = $wb.Worksheets.Item('変更操作')
            $last = $opsSheet.Cells.Item($opsSheet.Rows.Count, 1).End(-4162).Row
            $targetRow = -1
            for ($r = 9; $r -le $last; $r++) {
                if ([string]$opsSheet.Cells.Item($r, 17).Value2 -ieq $sourceFull) { $targetRow = $r; break }
            }
            if ($targetRow -lt 0) { throw "[$Id] operation row not found: $sourceFull" }

            $opsSheet.Unprotect()
            $opsSheet.Cells.Item($targetRow, 2).Value2 = '名前変更/移動'
            $opsSheet.Cells.Item($targetRow, 10).Value2 = $NewName
            $opsSheet.Cells.Item($targetRow, 11).Value2 = $MoveRelative
            $opsSheet.Cells.Item($targetRow, 12).Value2 = ''
            $opsSheet.Protect()

            $sequence = @(@{ Title = $expectRejectTitle; ButtonId = $IDOK })
            $watchJob = Start-ThreadJob -ScriptBlock $watcherScript -ArgumentList $excelPid, $sequence, $DialogTimeoutSeconds

            $runError = $null
            try {
                $null = $excel.Run('PreviewOperationPlan')
            } catch {
                $runError = $_
            }

            $watchResult = Receive-Job -Job $watchJob -Wait -AutoRemoveJob

            $stateText = [string]$opsSheet.Range('B5').Value2
            $afterFingerprint = Get-TreeFingerprint $FixtureRoot
            $fingerprintDiff = @(Compare-Object -ReferenceObject $beforeFingerprint -DifferenceObject $afterFingerprint)
            $fingerprintUnchanged = ($fingerprintDiff.Count -eq 0)

            $notes = [System.Collections.Generic.List[string]]::new()
            $pass = $true
            if ($null -ne $runError) { $pass = $false; $notes.Add("PreviewOperationPlan threw: $($runError.Exception.Message)") }
            if (-not $watchResult.completed) {
                $pass = $false
                if ($watchResult.timedOut) { $notes.Add("dialog not observed within ${DialogTimeoutSeconds}s (observed: $($watchResult.observedTitles -join ', '))") }
                elseif ($null -ne $watchResult.anomalyTitle) { $notes.Add("unexpected dialog title observed and left unclicked: $($watchResult.anomalyTitle)") }
                else { $notes.Add('dialog watcher did not complete for an unknown reason') }
            }
            if ($stateText -ne '検査失敗') { $pass = $false; $notes.Add("operation state was '$stateText', expected '検査失敗'") }
            if (-not $fingerprintUnchanged) {
                $pass = $false
                $diffText = ($fingerprintDiff | ForEach-Object { "$($_.SideIndicator) $($_.InputObject)" }) -join ' ; '
                $notes.Add("fixture tree fingerprint changed even though the plan was expected to be rejected: $diffText")
            }

            $result = [pscustomobject]@{
                id = $Id
                name = $Name
                status = $(if ($pass) { 'pass' } else { 'fail' })
                dialogTitleObserved = $(if ($watchResult.completed) { $expectRejectTitle } else { $null })
                operationState = $stateText
                fingerprintUnchanged = $fingerprintUnchanged
                notes = @($notes)
            }
        } catch {
            $result = [pscustomobject]@{
                id = $Id
                name = $Name
                status = 'fail'
                dialogTitleObserved = $null
                operationState = $null
                fingerprintUnchanged = $null
                notes = @("harness error: $($_.Exception.Message)")
            }
        }
        Write-Host ("  -> {0}" -f $result.status) -ForegroundColor $(if ($result.status -eq 'pass') { 'Green' } else { 'Red' })
        foreach ($n in $result.notes) { Write-Host "     $n" -ForegroundColor Yellow }
        return $result
    }

    $results.Add((Invoke-RejectionScenario -Id 'RSS-1' -Name 'destination collisionの拒否' `
        -FixtureName 'collision' -SourceRelative 'src\a.txt' -NewName 'b.txt' -MoveRelative ''))

    $results.Add((Invoke-RejectionScenario -Id 'RSS-2' -Name 'root外destinationの拒否' `
        -FixtureName 'root-escape' -SourceRelative 'src\a.txt' -NewName 'a.txt' -MoveRelative '..\outside'))

    $results.Add((Invoke-RejectionScenario -Id 'RSS-3' -Name 'destination祖先のjunctionによるroot逸脱拒否' `
        -FixtureName 'reparse-escape' -SourceRelative 'src\a.txt' -NewName 'a.txt' -MoveRelative 'link' -JunctionName 'link'))
}
finally {
    if ($null -ne $wb) {
        try { $wb.Close($false) } catch { }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
        $wb = $null
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        $excel = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
}

# 期待するscenario IDを明示し、「1件も実行されなかった」状態をpassと誤判定しないようにする。
# `-not @()` はPowerShellでは $true になるため、失敗件数が0であることだけを条件にすると
# $results が空のときにfail-openする。test-xlsm.ps1のgate receiptと同じく、未実行を
# 明示的な 'not-run' として扱い、期待IDが全て揃っていることを合格条件へ含める。
$expectedScenarioIds = @('RSS-1', 'RSS-2', 'RSS-3')
function Get-ScenarioResult([string] $ScenarioId) {
    $match = @($results | Where-Object { $_.id -eq $ScenarioId })
    if ($match.Count -eq 0) {
        return [pscustomobject]@{
            id = $ScenarioId
            name = '(not executed)'
            status = 'not-run'
            dialogTitleObserved = $null
            operationState = $null
            fingerprintUnchanged = $null
            notes = @('scenario did not execute; the run aborted before reaching it')
        }
    }
    return $match[0]
}
$scenarioResults = @($expectedScenarioIds | ForEach-Object { Get-ScenarioResult $_ })
$notPassed = @($scenarioResults | Where-Object { $_.status -ne 'pass' })
$allPassed = ($notPassed.Count -eq 0)

$receipt = [ordered]@{
    schemaVersion = 1
    generatedAtUtc = [DateTime]::UtcNow.ToString('o')
    candidatePath = [IO.Path]::GetRelativePath($root, [IO.Path]::GetFullPath($CandidatePath))
    candidateSha256 = $candidateSha256
    dialogTimeoutSeconds = $DialogTimeoutSeconds
    expectedScenarioIds = $expectedScenarioIds
    executedScenarioCount = $results.Count
    checks = [ordered]@{
        rootBoundary = (Get-ScenarioResult 'RSS-2')
        reparseBoundary = (Get-ScenarioResult 'RSS-3')
        collision = (Get-ScenarioResult 'RSS-1')
        recycleBinUnavailableFallback = [ordered]@{ status = 'static-source-review-only'; note = 'OD-11 (2026-08-20): dynamic verification is intentionally out of scope. Backed by static source review of the NoRecycleFiles / NukeOnDelete guards in modRecycleBin.bas plus embedded-source hash binding.' }
    }
    allRequestedChecksPassed = [bool]$allPassed
}
$receipt | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ReceiptPath -Encoding utf8NoBOM

# ---- cleanup: reparse-safe recursive delete of the owned run root ----
$resolvedRunRoot = [IO.Path]::GetFullPath($runRoot)
if ($resolvedRunRoot -ne $runRoot -or $resolvedRunRoot -eq $workspaceFull -or -not $resolvedRunRoot.StartsWith($workspaceFull + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing cleanup outside the created release smoke root: $resolvedRunRoot"
}
$cleanupItem = Get-Item -LiteralPath $resolvedRunRoot -Force
if (($cleanupItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Refusing cleanup of reparse point: $resolvedRunRoot" }
$cleanupEntries = @(Get-ChildItem -LiteralPath $resolvedRunRoot -Recurse -Force)
$cleanupReparse = @($cleanupEntries | Where-Object { ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 })
foreach ($entry in $cleanupReparse) {
    $resolvedEntry = [IO.Path]::GetFullPath($entry.FullName)
    if ($resolvedEntry -eq $resolvedRunRoot -or -not $resolvedEntry.StartsWith($resolvedRunRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing cleanup of reparse point outside the created release smoke root: $resolvedEntry"
    }
    Remove-Item -LiteralPath $resolvedEntry -Force
    if (Test-Path -LiteralPath $resolvedEntry) { throw "Release smoke reparse cleanup failed: $resolvedEntry" }
}
$cleanupEntries = @(Get-ChildItem -LiteralPath $resolvedRunRoot -Recurse -Force)
if (@($cleanupEntries | Where-Object { ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 }).Count -gt 0) {
    throw "Refusing cleanup because reparse points remain: $resolvedRunRoot"
}
Remove-Item -LiteralPath $resolvedRunRoot -Recurse -Force

Write-Host ''
Write-Host ("release safety subset (3 dynamic checks; recycle-bin fallback is static-review-only per OD-11): {0}" -f $(if ($allPassed) { 'PASS' } else { 'FAIL' })) -ForegroundColor $(if ($allPassed) { 'Green' } else { 'Red' })
Write-Output "receiptPath=$([IO.Path]::GetFullPath($ReceiptPath))"

if (-not $allPassed) { exit 1 }
exit 0
