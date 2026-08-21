#Requires -Version 7.0
<#
.SYNOPSIS
    対象 xlsm の RunScan マクロを既知のフォルダ構成に対して実行し、出力を検証する。

.DESCRIPTION
    一時フォルダへテスト用の階層を作成し、Excel COM 経由でマクロを実行して
    行順序・階層・集計値・アウトライン／インデント設定を突き合わせる。
    ブックは保存せずに閉じるため、成果物には影響しない。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $WorkbookPath,
    [string] $TestWorkRoot,
    [string] $GateReceiptPath,
    [string] $CandidateBuildReceiptPath,
    [string] $SourcePath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd([IO.Path]::DirectorySeparatorChar)
if (-not (Test-Path -LiteralPath (Join-Path $root 'VERSION') -PathType Leaf) -or -not (Test-Path -LiteralPath (Join-Path $root 'src') -PathType Container)) {
    throw "Repository root identity check failed at $root (VERSION or src\ missing)."
}
if ([string]::IsNullOrWhiteSpace($TestWorkRoot)) { $TestWorkRoot = Join-Path $root '.test-work' }
if ([string]::IsNullOrWhiteSpace($SourcePath)) { $SourcePath = Join-Path $root 'src\modFolderTree.bas' }

$script:Failures = @()
$script:CurrentCase = 'setup'
$script:CaseFailures = @{}
$script:CaseNotRun = @{}
$script:CaseExecuted = @{}
$requiredCaseIds = @(1..68 | ForEach-Object { [string]$_ })
$optionalCaseIds = @(69 | ForEach-Object { [string]$_ })
$workspaceFull = $root
$testParent = [IO.Path]::GetFullPath($TestWorkRoot)
if ($testParent -eq $workspaceFull -or -not $testParent.StartsWith($workspaceFull + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "TestWorkRoot must be a child of repository root: $testParent"
}
$runRoot = Join-Path $testParent ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $runRoot -Force | Out-Null
$runItem = Get-Item -LiteralPath $runRoot -Force
if (($runItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Test work root cannot be a reparse point: $runRoot" }
$gateReceiptFull = $null
$candidateBuildReceiptFull = $null
if (-not [string]::IsNullOrWhiteSpace($GateReceiptPath)) {
    $gateReceiptFull = [IO.Path]::GetFullPath($GateReceiptPath)
    if (-not $gateReceiptFull.StartsWith($workspaceFull + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw "Gate receipt is outside workspace: $gateReceiptFull" }
    if ([string]::IsNullOrWhiteSpace($CandidateBuildReceiptPath)) { throw 'Gate receipt generation requires CandidateBuildReceiptPath.' }
    $candidateBuildReceiptFull = [IO.Path]::GetFullPath($CandidateBuildReceiptPath)
    if (-not $candidateBuildReceiptFull.StartsWith($workspaceFull + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw "Candidate build receipt is outside workspace: $candidateBuildReceiptFull" }
    if (-not (Test-Path -LiteralPath $candidateBuildReceiptFull -PathType Leaf)) { throw "Candidate build receipt not found: $candidateBuildReceiptFull" }
}
$testWorkbookPath = Join-Path $runRoot 'under-test.xlsm'
Copy-Item -LiteralPath $WorkbookPath -Destination $testWorkbookPath -Force
$workbookSha256BeforeOpen = (Get-FileHash -LiteralPath $WorkbookPath -Algorithm SHA256).Hash
$sourceSha256BeforeOpen = (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash
$allSourceHashesBeforeOpen = [ordered]@{}
foreach ($srcName in @('modFolderTree.bas','modOperationPlan.bas','modFileOperations.bas','modRecycleBin.bas','modWorkbookIo.bas')) {
    $allSourceHashesBeforeOpen[$srcName] = (Get-FileHash -LiteralPath (Join-Path $root "src\$srcName") -Algorithm SHA256).Hash
}
$buildScript = Join-Path $PSScriptRoot 'build-xlsm.ps1'
$builderSha256BeforeOpen = (Get-FileHash -LiteralPath $buildScript -Algorithm SHA256).Hash
$candidateBuildReceiptSha256 = $null
if ($null -ne $candidateBuildReceiptFull) {
    $candidateBuild = Get-Content -LiteralPath $candidateBuildReceiptFull -Raw -Encoding UTF8 | ConvertFrom-Json
    $candidateBuildReceiptSha256 = (Get-FileHash -LiteralPath $candidateBuildReceiptFull -Algorithm SHA256).Hash
    if ($candidateBuild.schemaVersion -lt 2 -or
        $candidateBuild.candidateSha256 -ne $workbookSha256BeforeOpen -or
        $candidateBuild.sourceSha256 -ne $sourceSha256BeforeOpen -or
        $candidateBuild.builderSha256 -ne $builderSha256BeforeOpen -or
        [string]$candidateBuild.candidatePath -ne [IO.Path]::GetRelativePath($root, [IO.Path]::GetFullPath($WorkbookPath)) -or
        ([string]$candidateBuild.sourcePath -replace '\\','/') -ne 'src/modFolderTree.bas') {
        throw 'Candidate build receipt does not bind the workbook and source under test.'
    }
    $requiredBuildValidations = @(
        'contentType', 'vbaProjectPart', 'vbaCompile', 'fileFormat', 'sheetOrder', 'validations',
        'buttons', 'embeddedVba', 'noTestModule', 'defaultUiNoInjection', 'reopened'
    )
    $invalidBuildValidations = @($requiredBuildValidations | Where-Object {
            $name = $_
            $property = $candidateBuild.validation.PSObject.Properties | Where-Object { $_.Name -eq $name }
            $null -eq $property -or [string]$candidateBuild.validation.$name.status -ne 'pass'
        })
    if ($invalidBuildValidations.Count -gt 0) {
        throw "Candidate build receipt is not green: $($invalidBuildValidations -join ', ')"
    }
}

function Assert-Equal {
    param([string] $Name, $Expected, $Actual)
    if ($Expected -eq $Actual) {
        Write-Host ("  [OK]   {0}: {1}" -f $Name, $Actual)
    } else {
        $msg = "  [FAIL] {0}: expected='{1}' actual='{2}'" -f $Name, $Expected, $Actual
        Write-Host $msg -ForegroundColor Red
        $script:Failures += $msg
        if (-not $script:CaseFailures.ContainsKey($script:CurrentCase)) {
            $script:CaseFailures[$script:CurrentCase] = @()
        }
        $script:CaseFailures[$script:CurrentCase] += $msg
    }
}

function Start-Case {
    param([string] $Id, [string] $Title)
    $script:CurrentCase = $Id
    $script:CaseExecuted[$Id] = $true
    Write-Host ''
    Write-Host ("=== ケース{0}: {1} ===" -f $Id, $Title)
}

function Mark-CaseNotRun {
    param([string] $Id, [string] $Reason)
    $script:CaseNotRun[$Id] = $Reason
    Write-Host ("  [SKIP] ケース{0}: {1}" -f $Id, $Reason) -ForegroundColor Yellow
}

# ---- テスト用フォルダ構成 ---------------------------------------------------
#   root/
#     A/            <- 100 + 300 byte
#       A1/         <- 300 byte
#         deep.txt      (300 byte)
#       a.txt           (100 byte)
#     B/            <- 空フォルダ
#     .secret/      <- 隠し属性フォルダ
#       hidden.txt      (50 byte)
#     root.txt          (10 byte)
$fixture = Join-Path $runRoot 'fixture'
$null = New-Item -ItemType Directory -Path (Join-Path $fixture 'A\A1') -Force
$null = New-Item -ItemType Directory -Path (Join-Path $fixture 'B') -Force
$null = New-Item -ItemType Directory -Path (Join-Path $fixture '.secret') -Force

[IO.File]::WriteAllBytes((Join-Path $fixture 'root.txt'), (New-Object byte[] 10))
[IO.File]::WriteAllBytes((Join-Path $fixture 'A\a.txt'), (New-Object byte[] 100))
[IO.File]::WriteAllBytes((Join-Path $fixture 'A\A1\deep.txt'), (New-Object byte[] 300))
[IO.File]::WriteAllBytes((Join-Path $fixture '.secret\hidden.txt'), (New-Object byte[] 50))
(Get-Item -LiteralPath (Join-Path $fixture '.secret') -Force).Attributes = 'Directory, Hidden'

$excel = $null
$wb = $null

function New-OperationFixture {
    param([string] $Parent, [string] $Name)
    $root = Join-Path $Parent $Name
    $null = New-Item -ItemType Directory -Path (Join-Path $root 'src') -Force
    $null = New-Item -ItemType Directory -Path (Join-Path $root 'dst') -Force
    [IO.File]::WriteAllText((Join-Path $root 'src\a.txt'), 'a')
    [IO.File]::WriteAllText((Join-Path $root 'src\b.txt'), 'b')
    return $root
}

function Get-TreeFingerprint {
    param([string] $Root)
    $fullRoot = [IO.Path]::GetFullPath($Root)
    $entries = @(Get-ChildItem -LiteralPath $fullRoot -Recurse -Force | ForEach-Object {
        $relative = [IO.Path]::GetRelativePath($fullRoot, $_.FullName).Replace('\', '/')
        if ($_.PSIsContainer) {
            '{0}|dir|{1}' -f $relative, $_.LastWriteTimeUtc.Ticks
        } else {
            $hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
            '{0}|file|{1}|{2}|{3}' -f $relative, $_.Length, $_.LastWriteTimeUtc.Ticks, $hash
        }
    } | Sort-Object)
    return ($entries -join "`n")
}

function Get-StableTreeFingerprint {
    param([string] $Root)
    $previous = $null
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        $current = Get-TreeFingerprint $Root
        if ($null -ne $previous -and $current -eq $previous) { return $current }
        $previous = $current
        Start-Sleep -Milliseconds 100
    }
    return $previous
}

function Prepare-OperationDraft {
    param($Excel, $Workbook, [string] $Root)
    $null = $Excel.Run('RunScan', $Root, 0, $true, $false)
    if (-not [bool]$Excel.Run('WorkbookIoCreateOperationDraft', $Root)) { throw "operation draft creation failed: $Root" }
    $null = $Excel.Run('OperationDraftCreated', $Root)
    return $Workbook.Worksheets.Item('変更操作')
}

function Get-OperationRow {
    param($Sheet, [string] $FullPath)
    $last = $Sheet.Cells.Item($Sheet.Rows.Count, 1).End(-4162).Row
    for ($row = 9; $row -le $last; $row++) {
        if ([string]$Sheet.Cells.Item($row, 17).Value2 -ieq $FullPath) { return $row }
    }
    throw "operation row not found: $FullPath"
}

function Set-OperationRow {
    param($Sheet, [int] $Row, [string] $Kind, [string] $NewName = '', [string] $MoveRelative = '', [string] $FolderRelative = '')
    $Sheet.Unprotect()
    $Sheet.Cells.Item($Row, 2).Value2 = $Kind
    $Sheet.Cells.Item($Row, 10).Value2 = $NewName
    $Sheet.Cells.Item($Row, 11).Value2 = $MoveRelative
    $Sheet.Cells.Item($Row, 12).Value2 = $FolderRelative
    $Sheet.Protect()
}

function Normalize-TestPath {
    param([string] $Path)
    $full = [IO.Path]::GetFullPath($Path)
    if ($full.Length -gt 3) { $full = $full.TrimEnd('\') }
    return $full.ToUpperInvariant()
}

function Get-RecycleItemInfo {
    param([string] $Source)
    $target = Normalize-TestPath $Source
    $parent = Normalize-TestPath ([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($Source)))
    $name = [IO.Path]::GetFileName([IO.Path]::GetFullPath($Source))
    $shell = $null
    $bin = $null
    try {
        $shell = New-Object -ComObject Shell.Application
        $bin = $shell.Namespace(10)
        if ($null -eq $bin) { return $null }
        foreach ($item in @($bin.Items())) {
            try {
                $deletedFrom = [string]$item.ExtendedProperty('System.Recycle.DeletedFrom')
                $deletedFromNormalized = $null
                if (-not [string]::IsNullOrWhiteSpace($deletedFrom)) {
                    try { $deletedFromNormalized = Normalize-TestPath $deletedFrom } catch { $deletedFromNormalized = $null }
                }
                $itemName = [string]$item.Name
                $matchesExact = $deletedFromNormalized -eq $target
                $matchesParentAndName = $deletedFromNormalized -eq $parent -and $itemName -ieq $name
                if ($matchesExact -or $matchesParentAndName) {
                    return [pscustomobject]@{
                        SourceProperty = $deletedFrom
                        ItemName = $itemName
                        Size = [string]$item.ExtendedProperty('System.Size')
                        DateModified = [string]$item.ExtendedProperty('System.DateModified')
                    }
                }
            } catch { }
        }
        return $null
    } finally {
        if ($null -ne $bin) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($bin) } catch { } }
        if ($null -ne $shell) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell) } catch { } }
    }
}

function Restore-RecycleItemBySource {
    param([string] $Source)
    $target = Normalize-TestPath $Source
    $parent = Normalize-TestPath ([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($Source)))
    $name = [IO.Path]::GetFileName([IO.Path]::GetFullPath($Source))
    $shell = $null
    $bin = $null
    $matched = $false
    try {
        $shell = New-Object -ComObject Shell.Application
        $bin = $shell.Namespace(10)
        if ($null -eq $bin) { throw 'Recycle Bin namespace unavailable while restoring a test fixture.' }
        foreach ($item in @($bin.Items())) {
            $deletedFrom = [string]$item.ExtendedProperty('System.Recycle.DeletedFrom')
            $deletedFromNormalized = $null
            if (-not [string]::IsNullOrWhiteSpace($deletedFrom)) {
                try { $deletedFromNormalized = Normalize-TestPath $deletedFrom } catch { $deletedFromNormalized = $null }
            }
            $itemName = [string]$item.Name
            if ($deletedFromNormalized -ne $target -and -not ($deletedFromNormalized -eq $parent -and $itemName -ieq $name)) { continue }
            $matched = $true
            $restoreVerb = @($item.Verbs() | Where-Object { $_.Name -match '(?i)restore|復元|元に戻す' } | Select-Object -First 1)
            if ($restoreVerb.Count -gt 0) {
                $restoreVerb[0].DoIt()
            } else {
                $item.InvokeVerb('RESTORE')
            }
            break
        }
        if (-not $matched) { throw "Recycle Bin test fixture not found for restore: $Source" }
        for ($attempt = 1; $attempt -le 20; $attempt++) {
            if (Test-Path -LiteralPath $Source) { return $true }
            Start-Sleep -Milliseconds 250
        }
        throw "Recycle Bin restore did not recreate the fixture: $Source"
    } finally {
        if ($null -ne $bin) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($bin) } catch { } }
        if ($null -ne $shell) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell) } catch { } }
    }
}

function Assert-RecycleRoundTrip {
    param([string] $CaseId, [string] $OperationParent, [string] $FixtureName, [string] $RelativeSource, [string] $Kind)
    $opRoot = New-OperationFixture $OperationParent $FixtureName
    $source = Join-Path $opRoot $RelativeSource
    try {
        $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $row = Get-OperationRow $ops $source
        Set-OperationRow $ops $row $Kind
        Assert-OperationPlanPass $script:V1Excel | Out-Null
        $result = [string]$script:V1Excel.Run('ExecuteOperationPlanTest')
        Assert-Equal ("case{0} recycle execution" -f $CaseId) $true $result.StartsWith('success:')
        Assert-Equal ("case{0} source removed" -f $CaseId) $false (Test-Path -LiteralPath $source)
        $info = Get-RecycleItemInfo $source
        Assert-Equal ("case{0} Recycle Bin item matched" -f $CaseId) $true ($null -ne $info)
    } finally {
        if (-not (Test-Path -LiteralPath $source)) {
            $info = Get-RecycleItemInfo $source
            if ($null -ne $info) { [void](Restore-RecycleItemBySource $source) }
        }
        Assert-Equal ("case{0} source restored" -f $CaseId) $true (Test-Path -LiteralPath $source)
    }
}

function Assert-OperationPlanPass {
    param($Excel)
    $result = [string]$Excel.Run('RunOperationPlanTest')
    if (-not $result.StartsWith('pass|')) { Write-Host ("  plan detail: {0}" -f $result) -ForegroundColor Red }
    Assert-Equal 'operation plan pass' $true $result.StartsWith('pass|')
    return $result
}

function Assert-OperationPlanReject {
    param($Excel)
    $result = [string]$Excel.Run('RunOperationPlanTest')
    Assert-Equal 'operation plan reject' $true $result.StartsWith('fail|')
    return $result
}

function New-SequenceFixture {
    param([string] $Parent, [string] $Name, [string[]] $FileNames)
    $root = Join-Path $Parent $Name
    $source = Join-Path $root 'src'
    $null = New-Item -ItemType Directory -Path $source -Force
    foreach ($fileName in $FileNames) {
        [IO.File]::WriteAllText((Join-Path $source $fileName), $fileName)
    }
    return $root
}

function Invoke-OperationCase {
    param([string] $Id, [string] $Title, [scriptblock] $Body)
    Start-Case $Id $Title
    try { & $Body }
    catch {
        $message = "ケース$Id 実行失敗: $($_.Exception.Message)"
        $script:CaseFailures[$Id] = @($message)
        $script:Failures += $message
        Write-Host $message -ForegroundColor Red
    }
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1   # msoAutomationSecurityLow: マクロを有効にして開く

    $wb = $excel.Workbooks.Open($testWorkbookPath)
    $script:V1Excel = $excel
    $script:V1Workbook = $wb

    Start-Case '1' '既定（ファイル表示あり / 隠し除外 / 階層無制限）'
    $summary = $excel.Run('RunScan', $fixture, 0, $true, $false)
    Write-Host ($summary -replace "`r`n", ' | ')

    $tree = $wb.Worksheets.Item('ツリー')
    $last = $tree.Cells.Item($tree.Rows.Count, 1).End(-4162).Row   # xlUp

    # 期待される出力（深さ優先: フォルダ -> 直下ファイル -> サブフォルダ）
    $expected = @(
        @{ Name = Split-Path $fixture -Leaf; Type = 'フォルダ'; Depth = 0; Items = 3;  Size = 410 },
        @{ Name = 'root.txt';                Type = 'ファイル'; Depth = 1; Items = ''; Size = 10 },
        @{ Name = 'A';                       Type = 'フォルダ'; Depth = 1; Items = 2;  Size = 400 },
        @{ Name = 'a.txt';                   Type = 'ファイル'; Depth = 2; Items = ''; Size = 100 },
        @{ Name = 'A1';                      Type = 'フォルダ'; Depth = 2; Items = 1;  Size = 300 },
        @{ Name = 'deep.txt';                Type = 'ファイル'; Depth = 3; Items = ''; Size = 300 },
        @{ Name = 'B';                       Type = 'フォルダ'; Depth = 1; Items = 0;  Size = 0 }
    )

    Assert-Equal '出力行数' (3 + $expected.Count) $last

    for ($i = 0; $i -lt $expected.Count; $i++) {
        $row = 4 + $i
        $e = $expected[$i]
        Assert-Equal ("行{0} 名前" -f $row)   $e.Name  $tree.Cells.Item($row, 1).Value2
        Assert-Equal ("行{0} 種別" -f $row)   $e.Type  $tree.Cells.Item($row, 2).Value2
        Assert-Equal ("行{0} 階層" -f $row)   $e.Depth $tree.Cells.Item($row, 3).Value2
        Assert-Equal ("行{0} サイズ" -f $row) $e.Size  $tree.Cells.Item($row, 5).Value2
        if ($e.Items -ne '') {
            Assert-Equal ("行{0} 配下ファイル数" -f $row) $e.Items $tree.Cells.Item($row, 4).Value2
        }
        # インデント = 階層、アウトライン = 階層 + 1
        Assert-Equal ("行{0} インデント" -f $row)   $e.Depth       $tree.Cells.Item($row, 1).IndentLevel
        Assert-Equal ("行{0} アウトライン" -f $row) ($e.Depth + 1) $tree.Rows.Item($row).OutlineLevel
    }

    Assert-Equal '隠しフォルダを除外' $false ($summary -like '*.secret*')
    Assert-Equal 'アウトライン集計行が上' 0 $tree.Outline.SummaryRow   # xlSummaryAbove = 0
    Assert-Equal 'オートフィルタ有効'     $true $tree.AutoFilterMode
    Start-Case '4' '子フォルダ更新日時'
    $expectedAWrite = (Get-Item -LiteralPath (Join-Path $fixture 'A') -Force).LastWriteTime.ToOADate()
    $actualAWrite = [double]$tree.Cells.Item(6, 6).Value2
    Assert-Equal '子フォルダ更新日時が空欄でない' $true ($actualAWrite -gt 0)
    Assert-Equal '子フォルダ更新日時が一致' $true ([Math]::Abs($expectedAWrite - $actualAWrite) -lt (2.0 / 86400.0))

    Start-Case '2' 'フォルダのみ / 最大階層 1'
    $null = $excel.Run('RunScan', $fixture, 1, $false, $false)
    $last2 = $tree.Cells.Item($tree.Rows.Count, 1).End(-4162).Row
    # root + A + B の 3 行（ファイル非表示、A1 は階層 1 で打ち切り）
    Assert-Equal '出力行数(最大階層1)' 6 $last2
    Assert-Equal '行4 名前' (Split-Path $fixture -Leaf) $tree.Cells.Item(4, 1).Value2
    Assert-Equal '行5 名前' 'A' $tree.Cells.Item(5, 1).Value2
    Assert-Equal '行6 名前' 'B' $tree.Cells.Item(6, 1).Value2
    # 打ち切り済みなので A の集計は直下の a.txt のみ
    Assert-Equal '行5 サイズ(打ち切り)' 100 $tree.Cells.Item(5, 5).Value2

    Start-Case '3' '隠し属性を含める'
    $null = $excel.Run('RunScan', $fixture, 0, $true, $true)
    Assert-Equal 'ルート合計サイズ(隠し込み)'   460 $tree.Cells.Item(4, 5).Value2
    Assert-Equal 'ルートファイル数(隠し込み)'     4 $tree.Cells.Item(4, 4).Value2
    $names = @(4..($tree.Cells.Item($tree.Rows.Count, 1).End(-4162).Row) | ForEach-Object { $tree.Cells.Item($_, 1).Value2 })
    Assert-Equal '隠しフォルダが出力される' $true ($names -contains '.secret')

    Start-Case '5' '空フォルダと0 byteファイル'
    $emptyFixture = Join-Path $runRoot 'empty-case'
    $null = New-Item -ItemType Directory -Path $emptyFixture -Force
    [IO.File]::WriteAllBytes((Join-Path $emptyFixture 'zero.bin'), (New-Object byte[] 0))
    $null = $excel.Run('RunScan', $emptyFixture, 0, $true, $false)
    $emptyLast = $tree.Cells.Item($tree.Rows.Count, 1).End(-4162).Row
    Assert-Equal '空/0byte 出力行数' 5 $emptyLast
    Assert-Equal '空/0byte ルート件数' 1 $tree.Cells.Item(4, 4).Value2
    Assert-Equal '0 byte サイズ' 0 $tree.Cells.Item(5, 5).Value2

    Start-Case '7' '260文字超の長いパス'
    $longFixture = Join-Path $runRoot 'long-case'
    $longPath = $longFixture
    try {
        $segments = 1..14 | ForEach-Object { ('segment{0:00}_abcdefghijklmnop' -f $_) }
        foreach ($segment in $segments) {
            $longPath = Join-Path $longPath $segment
            $null = New-Item -ItemType Directory -Path $longPath -Force
        }
        $longFile = Join-Path $longPath 'long.txt'
        [IO.File]::WriteAllText($longFile, 'long')
        Assert-Equal '長いパス長' $true (([IO.Path]::GetFullPath($longFile)).Length -gt 260)
        $null = $excel.Run('RunScan', $longFixture, 0, $true, $false)
        $longNames = @(4..($tree.Cells.Item($tree.Rows.Count, 1).End(-4162).Row) | ForEach-Object { [string]$tree.Cells.Item($_, 1).Value2 })
        Assert-Equal '長いパス末端ファイルを列挙' $true ($longNames -contains 'long.txt')
    } catch {
        $message = "長いパスfixture作成または実行失敗: $($_.Exception.Message)"
        $script:CaseFailures['7'] = @($message)
        $script:Failures += $message
    }

    Start-Case '8' 'reparse pointを表示するが追跡しない'
    $reparseFixture = Join-Path $runRoot 'reparse-case'
    $reparseTarget = Join-Path $reparseFixture 'target'
    $reparseLink = Join-Path $reparseFixture 'link'
    $null = New-Item -ItemType Directory -Path $reparseTarget -Force
    [IO.File]::WriteAllText((Join-Path $reparseTarget 'inside.txt'), 'inside')
    $linkCreated = $false
    try {
        $null = New-Item -ItemType Junction -Path $reparseLink -Target $reparseTarget -Force
        $linkCreated = $true
    } catch {
        Mark-CaseNotRun '8' "junction作成が許可されていません: $($_.Exception.Message)"
    }
    if ($linkCreated) {
        $null = $excel.Run('RunScan', $reparseFixture, 0, $true, $false)
        $reparseTypes = @(4..($tree.Cells.Item($tree.Rows.Count, 1).End(-4162).Row) | ForEach-Object { [string]$tree.Cells.Item($_, 2).Value2 })
        $reparsePaths = @(4..($tree.Cells.Item($tree.Rows.Count, 1).End(-4162).Row) | ForEach-Object { [string]$tree.Cells.Item($_, 7).Value2 })
        Assert-Equal 'reparseリンクを表示' $true ($reparseTypes -contains 'リンク')
        Assert-Equal 'リンク先を追跡しない' $false ($reparsePaths -contains (Join-Path $reparseLink 'inside.txt'))
        Remove-Item -LiteralPath $reparseLink -Force
    }

    Start-Case '9' 'FindFirstFileW障害の分類'
    $findFirstSummary = $excel.Run('RunScanTest', $fixture, 0, $true, $false, 0, 2, $false, $false, 0, 5, 0)
    Assert-Equal 'FindFirst access deniedの表示' $true ($findFirstSummary -like '*アクセスできなかったフォルダが*')
    Assert-Equal 'ClassifyFindResult access-denied' 'access-denied' ([string]$excel.Run('ClassifyFindResult', $false, 5))
    Assert-Equal 'ClassifyFindResult findfirst-error' 'findfirst-error' ([string]$excel.Run('ClassifyFindResult', $false, 999))
    Assert-Equal 'ClassifyFindResult complete' 'complete' ([string]$excel.Run('ClassifyFindResult', $true, 18))
    Assert-Equal 'ClassifyFindResult incomplete' 'incomplete' ([string]$excel.Run('ClassifyFindResult', $true, 123))

    Start-Case '10' 'FindNextFileW途中エラー'
    $findNextSummary = $excel.Run('RunScanTest', $fixture, 0, $true, $false, 0, 2, $false, $false, 0, 0, 123)
    Assert-Equal 'FindNext途中エラーの表示' $true ($findNextSummary -like '*列挙不完全またはI/Oエラー*')

    Start-Case '11' '内部512階層安全弁'
    $depthFixture = Join-Path $runRoot 'depth-case'
    $depthPath = $depthFixture
    try {
        1..513 | ForEach-Object {
            $depthPath = Join-Path $depthPath ('d{0:000}' -f $_)
            $null = New-Item -ItemType Directory -Path $depthPath -Force
        }
        $depthSummary = $excel.Run('RunScan', $depthFixture, 0, $false, $false)
        Assert-Equal '512階層安全弁の通知' $true ($depthSummary -like '*内部安全弁 512 階層*')
    } catch {
        $message = "512階層fixture作成または実行失敗: $($_.Exception.Message)"
        $script:CaseFailures['11'] = @($message)
        $script:Failures += $message
    }

    Start-Case '20' 'embedded VBAとsource本文のhash一致'
    $module = $wb.VBProject.VBComponents.Item('modFolderTree')
    $embedded = [regex]::Replace([regex]::Replace([regex]::Replace([string]$module.CodeModule.Lines(1, $module.CodeModule.CountOfLines), '(?m)^Attribute [^\r\n]*(?:\r?\n|$)', ''), '(?m)^#Const TEST_BUILD = (True|False)[ \t]*(?:\r?\n|$)', ''), '\r\n?|\n', [string][char]10).TrimEnd([char]10)
    $source = [regex]::Replace([regex]::Replace((Get-Content -LiteralPath $SourcePath -Raw -Encoding UTF8), '(?m)^Attribute [^\r\n]*(?:\r?\n|$)', ''), '\r\n?|\n', [string][char]10).TrimEnd([char]10)
    $sha = [Security.Cryptography.SHA256]::Create()
    $embeddedHash = ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($embedded.ToLowerInvariant())))).Replace('-', '')
    $sourceHash = ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($source.ToLowerInvariant())))).Replace('-', '')
    Assert-Equal 'embedded/source本文hash' $sourceHash $embeddedHash

    Start-Case '18' '破損ZIP partのoffline構造検出と無関係fileへの書き込みなし'
    $promotionFinal = Join-Path $runRoot 'sentinel-unwritten.xlsm'
    Copy-Item -LiteralPath $WorkbookPath -Destination $promotionFinal
    $promotionBefore = (Get-FileHash -LiteralPath $promotionFinal -Algorithm SHA256).Hash
    $corruptCandidate = Join-Path $runRoot 'corrupt-no-vba.xlsm'
    Copy-Item -LiteralPath $WorkbookPath -Destination $corruptCandidate
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $corruptZip = [IO.Compression.ZipFile]::Open($corruptCandidate, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $vbaPart = $corruptZip.GetEntry('xl/vbaProject.bin')
        if ($null -eq $vbaPart) { throw 'Known-good candidate has no xl/vbaProject.bin before corruption.' }
        $vbaPart.Delete()
    } finally {
        $corruptZip.Dispose()
    }
    $pwsh = (Get-Command pwsh -ErrorAction Stop).Source
    & $pwsh -NoProfile -File $buildScript -ValidatePackageOffline -CandidatePath $corruptCandidate *> $null
    $offlineValidationExit = $LASTEXITCODE
    $promotionAfter = (Get-FileHash -LiteralPath $promotionFinal -Algorithm SHA256).Hash
    Assert-Equal 'vbaProject.bin欠落のoffline検出' $true ($offlineValidationExit -ne 0)
    Assert-Equal '破損fixture検証時の無関係fileのhash不変' $promotionBefore $promotionAfter

    Start-Case '19' '配置後検証失敗時のrollback'
    $rollbackFinal = Join-Path $runRoot 'rollback-final.xlsm'
    Copy-Item -LiteralPath $WorkbookPath -Destination $rollbackFinal
    $rollbackCandidate = Join-Path $runRoot 'rollback-candidate.xlsm'
    Copy-Item -LiteralPath $WorkbookPath -Destination $rollbackCandidate
    $rollbackCandidateHash = (Get-FileHash -LiteralPath $rollbackCandidate -Algorithm SHA256).Hash
    $rollbackBefore = (Get-FileHash -LiteralPath $rollbackFinal -Algorithm SHA256).Hash
    $rollbackBuildReceipt = Join-Path $runRoot 'rollback-build-receipt.json'
    $pass = { param([string]$Evidence) [ordered]@{ status = 'pass'; evidence = $Evidence } }
    $rollbackValidation = [ordered]@{}
    foreach ($name in @('contentType','vbaProjectPart','vbaCompile','fileFormat','sheetOrder','validations','buttons','embeddedVba','noTestModule','defaultUiNoInjection','reopened')) {
        $rollbackValidation[$name] = & $pass 'test fixture copied from the green candidate'
    }
    $rollbackSourceHashes = [ordered]@{}
    foreach ($name in @('modFolderTree.bas','modOperationPlan.bas','modFileOperations.bas','modRecycleBin.bas','modWorkbookIo.bas')) {
        $rollbackSourceHashes[$name] = (Get-FileHash -LiteralPath (Join-Path $root "src\$name") -Algorithm SHA256).Hash
    }
    [ordered]@{
        schemaVersion = 2
        generatedAtUtc = [DateTime]::UtcNow.ToString('o')
        builderPath = [IO.Path]::GetRelativePath($root, $buildScript)
        builderSha256 = (Get-FileHash -LiteralPath $buildScript -Algorithm SHA256).Hash
        sourcePath = 'src/modFolderTree.bas'
        sourceSha256 = $sourceSha256BeforeOpen
        sourceBodySha256 = $sourceHash
        candidatePath = [IO.Path]::GetRelativePath($root, $rollbackCandidate)
        candidateSha256 = $rollbackCandidateHash
        sourceHashes = $rollbackSourceHashes
        embeddedVbaBodySha256 = $embeddedHash
        excelVersion = [string]$excel.Version
        excelBitness = if ([string]$excel.OperatingSystem -match '64-bit') { 'x64' } else { 'not-recorded' }
        excelOperatingSystem = [string]$excel.OperatingSystem
        vbaSignature = 'absent'
        validation = $rollbackValidation
    } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $rollbackBuildReceipt -Encoding UTF8
    $rollbackBuildReceiptHash = (Get-FileHash -LiteralPath $rollbackBuildReceipt -Algorithm SHA256).Hash
    $greenGate = Join-Path $runRoot 'green-gate.json'
    $greenCases = [ordered]@{}
    $requiredCaseIds | ForEach-Object { $greenCases[$_] = 'pass' }
    [ordered]@{
        schemaVersion = 1
        candidatePath = [IO.Path]::GetRelativePath($root, $rollbackCandidate)
        candidateSha256 = $rollbackCandidateHash
        sourcePath = [IO.Path]::GetRelativePath($root, [IO.Path]::GetFullPath($SourcePath))
        sourceSha256 = $sourceSha256BeforeOpen
        candidateBuildReceiptPath = [IO.Path]::GetRelativePath($root, $rollbackBuildReceipt)
        candidateBuildReceiptSha256 = $rollbackBuildReceiptHash
        allPassed = $true
        requiredCaseIds = $requiredCaseIds
        cases = $greenCases
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $greenGate -Encoding UTF8
    $greenReadiness = Join-Path $runRoot 'green-readiness.json'
    [ordered]@{
        schemaVersion = 1
        status = 'READY_FOR_PROMOTION_APPROVAL'
        promotionEligible = $true
        candidate = [ordered]@{ sha256 = $rollbackCandidateHash }
    } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $greenReadiness -Encoding UTF8
    $rollbackOutput = @(& $pwsh -NoProfile -File $buildScript -PromoteCandidate -SimulatePromotionFailure `
            -OutputPath $rollbackFinal -CandidatePath $rollbackCandidate `
            -CandidateBuildReceiptPath $rollbackBuildReceipt -CandidateGateReceiptPath $greenGate `
            -CandidateReadinessReceiptPath $greenReadiness 2>&1)
    $rollbackExit = $LASTEXITCODE
    $rollbackAfter = (Get-FileHash -LiteralPath $rollbackFinal -Algorithm SHA256).Hash
    $rollbackBackups = @(Get-ChildItem -LiteralPath $runRoot -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'rollback-final.xlsm.backup.*' })
    Write-Host ("  rollback subprocess: {0}" -f (($rollbackOutput | ForEach-Object { [string]$_ }) -join ' | '))
    Assert-Equal '配置後失敗の検出' $true ($rollbackExit -ne 0)
    Assert-Equal 'rollback後のhash復元' $rollbackBefore $rollbackAfter
    Assert-Equal 'rollback backup保全' $true ($rollbackBackups.Count -eq 1)
    $firstPromotionFinal = Join-Path $runRoot 'first-promotion-final.xlsm'
    $firstRollbackOutput = @(& $pwsh -NoProfile -File $buildScript -PromoteCandidate -SimulatePromotionFailure `
            -OutputPath $firstPromotionFinal -CandidatePath $rollbackCandidate `
            -CandidateBuildReceiptPath $rollbackBuildReceipt -CandidateGateReceiptPath $greenGate `
            -CandidateReadinessReceiptPath $greenReadiness 2>&1)
    $firstRollbackExit = $LASTEXITCODE
    Write-Host ("  first promotion rollback subprocess: {0}" -f (($firstRollbackOutput | ForEach-Object { [string]$_ }) -join ' | '))
    Assert-Equal '初回配置後失敗の検出' $true ($firstRollbackExit -ne 0)
    Assert-Equal '初回配置後失敗で新規target除去' $false (Test-Path -LiteralPath $firstPromotionFinal)

    Start-Case '13' '行上限の直前・一致・超過'
    foreach ($limit in 1, 2, 3) {
        $limitSummary = $excel.Run('RunScanTest', $fixture, 0, $true, $false, $limit, 2, $false, $false, 0)
        Assert-Equal ("行上限{0}の通知" -f $limit) $true ($limitSummary -like "*上限 $limit*")
    }
    Assert-Equal '行上限後の既存最終行を維持' $true ([string]$tree.Cells.Item(4, 1).Value2 -ne '')

    Start-Case '12' 'Esc中断と部分集計'
    $cancelSummary = $excel.Run('RunScanTest', $fixture, 0, $true, $false, 0, 2, $false, $false, 2)
    Assert-Equal '中断の通知' $true ($cancelSummary -like '*Esc キーで中断しました*')

    Start-Case '14' 'バッファ確保失敗'
    $allocationSummary = $excel.Run('RunScanTest', $fixture, 0, $true, $false, 0, 2, $true, $false, 0)
    Assert-Equal '確保失敗の通知' $true ($allocationSummary -like '*メモリを確保できなくなった*')

    Start-Case '15' 'チャンク境界'
    $chunkSummary = $excel.Run('RunScanTest', $fixture, 0, $true, $false, 0, 2, $false, $false, 0)
    Assert-Equal 'チャンク境界の通常完了' $true ($chunkSummary -like '*ツリーを作成しました*')

    Start-Case '16' 'Application状態の完全復元'
    $savedScreenUpdating = $excel.ScreenUpdating
    $savedEnableEvents = $excel.EnableEvents
    $savedCalculation = $excel.Calculation
    $savedCursor = $excel.Cursor
    $savedCancel = $excel.EnableCancelKey
    $savedStatusBar = $excel.StatusBar
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $excel.Calculation = -4135
    $excel.Cursor = 2
    $excel.EnableCancelKey = 1
    $excel.StatusBar = 'folder-tree-test-sentinel'
    $null = $excel.Run('RunScanTest', $fixture, 0, $true, $false, 0, 2, $false, $false, 0)
    Assert-Equal 'ScreenUpdating復元' $false $excel.ScreenUpdating
    Assert-Equal 'EnableEvents復元' $false $excel.EnableEvents
    Assert-Equal 'Calculation復元' -4135 $excel.Calculation
    Assert-Equal 'Cursor復元' 2 $excel.Cursor
    Assert-Equal 'EnableCancelKey復元' 1 $excel.EnableCancelKey
    Assert-Equal 'StatusBar復元' 'folder-tree-test-sentinel' $excel.StatusBar
    $excel.ScreenUpdating = $savedScreenUpdating
    $excel.EnableEvents = $savedEnableEvents
    $excel.Calculation = $savedCalculation
    $excel.Cursor = $savedCursor
    $excel.EnableCancelKey = $savedCancel
    $excel.StatusBar = $savedStatusBar

    Start-Case '17' 'アウトライン／インデント設定失敗'
    $outlineSummary = $excel.Run('RunScanTest', $fixture, 0, $true, $false, 0, 2, $false, $true, 0)
    Assert-Equal 'outline失敗の通知' $true ($outlineSummary -like '*アウトライン／インデント設定に*')

    Start-Case '21' '再入拒否と進行中状態'
    $reentrySummary = $excel.Run('ProbeReentry', $fixture)
    Assert-Equal '再入拒否' $true ($reentrySummary -like '*処理中です*')

    Start-Case '6' '外部由来文字列の数式化防止'
    $formulaRoot = Join-Path $runRoot '=root'
    $null = New-Item -ItemType Directory -Path $formulaRoot -Force
    foreach ($name in @('=eq', '+plus', '-minus', '@at')) {
        $dir = Join-Path $formulaRoot $name
        $null = New-Item -ItemType Directory -Path $dir -Force
        [IO.File]::WriteAllText((Join-Path $dir 'value.txt'), 'x')
    }
    $null = $excel.Run('RunScan', $formulaRoot, 0, $true, $false)
    $formulaNames = @(4..($tree.Cells.Item($tree.Rows.Count, 1).End(-4162).Row) | ForEach-Object { [string]$tree.Cells.Item($_, 1).Value2 })
    foreach ($name in @('=eq', '+plus', '-minus', '@at')) {
        Assert-Equal ("文字列保持 {0}" -f $name) $true ($formulaNames -contains $name)
    }
    $formulaCells = 0
    try { $formulaCells = $tree.UsedRange.SpecialCells(-4123).Count } catch { $formulaCells = 0 }
    Assert-Equal 'ツリー数式セル数' 0 $formulaCells
    $pipeValue = $excel.Run('RunStringSafetyProbe', 'pipe|name')
    Assert-Equal '禁止文字を含む注入値の保持' 'pipe|name' $pipeValue

    $operationParent = Join-Path $runRoot 'operations'
    $null = New-Item -ItemType Directory -Path $operationParent -Force

    Invoke-OperationCase '22' '変更一覧生成前後のfixture不変性' {
        $opRoot = New-OperationFixture $operationParent 'case22'
        $before = Get-StableTreeFingerprint $opRoot
        $null = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $after = Get-StableTreeFingerprint $opRoot
        Assert-Equal 'fixture fingerprint不変' $before $after
    }

    Invoke-OperationCase '23' '変更操作シートの構造・保護・ボタン' {
        $ops = $script:V1Workbook.Worksheets.Item('変更操作')
        $log = $script:V1Workbook.Worksheets.Item('実行ログ')
        Assert-Equal '変更操作保護' $true $ops.ProtectContents
        Assert-Equal '実行ログ保護' $true $log.ProtectContents
        foreach ($name in @('btn_SetOperationSequencePreview','btn_PreviewOperationPlan','btn_ExecuteOperationPlan','btn_RecreateOperationSheet')) { $null = $ops.Shapes.Item($name); Assert-Equal ("button $name") $true $true }
    }

    Invoke-OperationCase '24' '数式化された編集列の拒否' {
        $operationSource = Get-Content -LiteralPath (Join-Path $root 'src\modOperationPlan.bas') -Raw -Encoding UTF8
        Assert-Equal '編集列HasFormula拒否' $true ($operationSource -match 'OP_COL_NEW_NAME\)\.HasFormula')
        Assert-Equal '編集列数式メッセージ' $true ($operationSource -match '編集可能列に数式があります')
        $opRoot = New-OperationFixture $operationParent 'case24'
        $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $row = Get-OperationRow $ops (Join-Path $opRoot 'src\a.txt')
        $ops.Unprotect(); $ops.Cells.Item($row, 10).Formula = '=1'; $ops.Protect()
        $result = [string]$script:V1Excel.Run('RunOperationPlanTest')
        Assert-Equal '数式編集列の実行時拒否' $true $result.StartsWith('fail|')
    }

    Invoke-OperationCase '25' '一覧rootと設定rootの不一致' {
        $opRoot = New-OperationFixture $operationParent 'case25'; $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $ops.Unprotect(); $ops.Range('B2').Value2 = (Join-Path $operationParent 'other-root'); $ops.Protect()
        Assert-OperationPlanReject $script:V1Excel | Out-Null
    }

    Invoke-OperationCase '26' '単一ファイルの名前変更' {
        $opRoot = New-OperationFixture $operationParent 'case26'; $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $row = Get-OperationRow $ops (Join-Path $opRoot 'src\a.txt'); Set-OperationRow $ops $row '名前変更/移動' 'renamed.txt'
        Assert-OperationPlanPass $script:V1Excel | Out-Null
        $result = [string]$script:V1Excel.Run('ExecuteOperationPlanTest')
        Assert-Equal 'rename実行結果' $true $result.StartsWith('success:')
        Assert-Equal 'rename source消失' $false (Test-Path -LiteralPath (Join-Path $opRoot 'src\a.txt'))
        Assert-Equal 'rename destination存在' $true (Test-Path -LiteralPath (Join-Path $opRoot 'src\renamed.txt'))
    }

    Invoke-OperationCase '27' '先頭連番previewの命名規則' {
        $opRoot = New-SequenceFixture $operationParent 'case27' @('report.pdf','archive.tar.gz')
        $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $result = [string]$script:V1Excel.Run('WorkbookIoSetSequencePreviewTest', '先頭', 0)
        Assert-Equal 'production preview pass' $true $result.StartsWith('pass|selected=2')
        $reportPreview = [string]$ops.Cells((Get-OperationRow $ops (Join-Path $opRoot 'src\report.pdf')), 10).Value2
        $archivePreview = [string]$ops.Cells((Get-OperationRow $ops (Join-Path $opRoot 'src\archive.tar.gz')), 10).Value2
        Assert-Equal '先頭連番' $true ($reportPreview -match '^0[01]_report\.pdf$')
        Assert-Equal '先頭複合拡張子' $true ($archivePreview -match '^0[01]_archive\.tar\.gz$')
        Assert-Equal '00と01を一度ずつ使用' $true ($reportPreview.Substring(0, 2) -ne $archivePreview.Substring(0, 2))
    }

    Invoke-OperationCase '28' '末尾連番previewの命名規則' {
        $opRoot = New-SequenceFixture $operationParent 'case28' @('report.pdf','archive.tar.gz')
        $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $result = [string]$script:V1Excel.Run('WorkbookIoSetSequencePreviewTest', '末尾', 0)
        Assert-Equal 'production preview pass' $true $result.StartsWith('pass|selected=2')
        $reportPreview = [string]$ops.Cells((Get-OperationRow $ops (Join-Path $opRoot 'src\report.pdf')), 10).Value2
        $archivePreview = [string]$ops.Cells((Get-OperationRow $ops (Join-Path $opRoot 'src\archive.tar.gz')), 10).Value2
        Assert-Equal '末尾連番' $true ($reportPreview -match '^report_0[01]\.pdf$')
        Assert-Equal '末尾複合拡張子' $true ($archivePreview -match '^archive\.tar_0[01]\.gz$')
    }

    Invoke-OperationCase '29' '100件の自動3桁化' {
        $fileNames = @(0..99 | ForEach-Object { 'file-{0:D3}.txt' -f $_ })
        $opRoot = New-SequenceFixture $operationParent 'case29' $fileNames
        $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $result = [string]$script:V1Excel.Run('WorkbookIoSetSequencePreviewTest', '先頭', 0)
        Assert-Equal 'production 100件preview' $true $result.StartsWith('pass|selected=100')
        $last = $ops.Cells($ops.Rows.Count, 1).End(-4162).Row
        $previews = @($ops.Range("J9:J$last").Value2 | ForEach-Object { [string]$_ })
        Assert-Equal '100件すべて3桁' 100 @($previews | Where-Object { $_ -match '^\d{3}_' }).Count
        Assert-Equal '100件先頭' $true ($previews -contains '000_file-000.txt')
        Assert-Equal '100件末尾' $true ($previews -contains '099_file-099.txt')
    }

    Invoke-OperationCase '30' '開始番号変更と4桁境界' {
        $opRoot = New-SequenceFixture $operationParent 'case30' @('a.txt','b.txt')
        $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $result = [string]$script:V1Excel.Run('WorkbookIoSetSequencePreviewTest', '先頭', 999)
        Assert-Equal 'production 4桁境界preview' $true $result.StartsWith('pass|selected=2')
        $values = @([string]$ops.Cells((Get-OperationRow $ops (Join-Path $opRoot 'src\a.txt')), 10).Value2,
                    [string]$ops.Cells((Get-OperationRow $ops (Join-Path $opRoot 'src\b.txt')), 10).Value2)
        Assert-Equal '0999を含む' $true ($values -contains '0999_a.txt' -or $values -contains '0999_b.txt')
        Assert-Equal '1000を含む' $true ($values -contains '1000_a.txt' -or $values -contains '1000_b.txt')
        $thousandNames = @(0..999 | ForEach-Object { 'item-{0:D4}.txt' -f $_ })
        $thousandRoot = New-SequenceFixture $operationParent 'case30-thousand' $thousandNames
        $thousandOps = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $thousandRoot
        $thousandResult = [string]$script:V1Excel.Run('WorkbookIoSetSequencePreviewTest', '先頭', 0)
        Assert-Equal 'production 1000件preview' $true $thousandResult.StartsWith('pass|selected=1000')
        $thousandLast = $thousandOps.Cells($thousandOps.Rows.Count, 1).End(-4162).Row
        $thousandPreviews = @($thousandOps.Range("J9:J$thousandLast").Value2 | ForEach-Object { [string]$_ })
        Assert-Equal '1000件すべて4桁' 1000 @($thousandPreviews | Where-Object { $_ -match '^\d{4}_' }).Count
        Assert-Equal '1000件先頭' $true ($thousandPreviews -contains '0000_item-0000.txt')
        Assert-Equal '1000件末尾' $true ($thousandPreviews -contains '0999_item-0999.txt')
    }

    Invoke-OperationCase '31' '複合拡張子とdotfile' {
        $opRoot = New-SequenceFixture $operationParent 'case31' @('.gitignore','archive.tar.gz')
        $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $result = [string]$script:V1Excel.Run('WorkbookIoSetSequencePreviewTest', '末尾', 0)
        Assert-Equal 'production extension preview' $true $result.StartsWith('pass|selected=2')
        $dotPreview = [string]$ops.Cells((Get-OperationRow $ops (Join-Path $opRoot 'src\.gitignore')), 10).Value2
        $archivePreview = [string]$ops.Cells((Get-OperationRow $ops (Join-Path $opRoot 'src\archive.tar.gz')), 10).Value2
        Assert-Equal 'dotfile末尾' $true ($dotPreview -match '^\.gitignore_\d{2}$')
        Assert-Equal '複合拡張子末尾' $true ($archivePreview -match '^archive\.tar_\d{2}\.gz$')
    }

    Invoke-OperationCase '32' '禁止文字・予約名・末尾space/dotの拒否' {
        $opRoot = New-OperationFixture $operationParent 'case32'; $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $row = Get-OperationRow $ops (Join-Path $opRoot 'src\a.txt')
        foreach ($bad in @('bad:name.txt','CON.txt','bad.')) { Set-OperationRow $ops $row '名前変更/移動' $bad; Assert-OperationPlanReject $script:V1Excel | Out-Null }
    }

    Invoke-OperationCase '33' 'destination collisionの拒否' {
        $opRoot = New-OperationFixture $operationParent 'case33'; $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $row = Get-OperationRow $ops (Join-Path $opRoot 'src\a.txt'); Set-OperationRow $ops $row '名前変更/移動' 'b.txt'
        Assert-OperationPlanReject $script:V1Excel | Out-Null
    }

    Invoke-OperationCase '34' 'case-only renameとA/B交換' {
        $opRoot = New-OperationFixture $operationParent 'case34'; $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $rowA = Get-OperationRow $ops (Join-Path $opRoot 'src\a.txt'); $rowB = Get-OperationRow $ops (Join-Path $opRoot 'src\b.txt')
        Set-OperationRow $ops $rowA '名前変更/移動' 'b.txt'; Set-OperationRow $ops $rowB '名前変更/移動' 'a.txt'
        Assert-OperationPlanPass $script:V1Excel | Out-Null
        $result = [string]$script:V1Excel.Run('ExecuteOperationPlanTest')
        Assert-Equal 'swap実行結果' $true $result.StartsWith('success:')
        Assert-Equal 'swap a存在' $true (Test-Path -LiteralPath (Join-Path $opRoot 'src\a.txt'))
        Assert-Equal 'swap b存在' $true (Test-Path -LiteralPath (Join-Path $opRoot 'src\b.txt'))
    }

    Invoke-OperationCase '35' 'root内ファイル移動' {
        $opRoot = New-OperationFixture $operationParent 'case35'; $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $row = Get-OperationRow $ops (Join-Path $opRoot 'src\a.txt'); Set-OperationRow $ops $row '名前変更/移動' 'a.txt' 'dst'
        Assert-OperationPlanPass $script:V1Excel | Out-Null; $result = [string]$script:V1Excel.Run('ExecuteOperationPlanTest')
        Assert-Equal 'move実行結果' $true $result.StartsWith('success:')
        Assert-Equal 'move destination存在' $true (Test-Path -LiteralPath (Join-Path $opRoot 'dst\a.txt'))
    }

    Invoke-OperationCase '36' '相対path・absolute path・root外destinationの拒否' {
        $opRoot = New-OperationFixture $operationParent 'case36'; $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $row = Get-OperationRow $ops (Join-Path $opRoot 'src\a.txt'); Set-OperationRow $ops $row '名前変更/移動' 'a.txt' '..\outside'; Assert-OperationPlanReject $script:V1Excel | Out-Null
        Set-OperationRow $ops $row '名前変更/移動' 'a.txt' ([IO.Path]::GetTempPath()); Assert-OperationPlanReject $script:V1Excel | Out-Null
    }

    Invoke-OperationCase '37' 'destination祖先のjunctionによるroot逸脱拒否' {
        $opRoot = New-OperationFixture $operationParent 'case37'; $link = Join-Path $opRoot 'link'; $null = New-Item -ItemType Junction -Path $link -Target ([IO.Path]::GetTempPath()) -Force
        $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot; $row = Get-OperationRow $ops (Join-Path $opRoot 'src\a.txt'); Set-OperationRow $ops $row '名前変更/移動' 'a.txt' 'link'; Assert-OperationPlanReject $script:V1Excel | Out-Null
    }

    Invoke-OperationCase '38' 'フォルダの名前変更・移動拒否' {
        $opRoot = New-OperationFixture $operationParent 'case38'; $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot; $row = Get-OperationRow $ops (Join-Path $opRoot 'src'); Set-OperationRow $ops $row '名前変更/移動' 'renamed'; Assert-OperationPlanReject $script:V1Excel | Out-Null
    }

    Invoke-OperationCase '39' '単一・nested新規フォルダ作成' {
        $opRoot = New-OperationFixture $operationParent 'case39'; $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot; $ops.Unprotect(); $newRow = $ops.Cells($ops.Rows.Count, 1).End(-4162).Row + 1; $ops.Cells($newRow, 1).Value2 = 'new'; $ops.Cells($newRow, 2).Value2 = 'フォルダ作成'; $ops.Cells($newRow, 12).Value2 = 'new\nested'; $ops.Protect(); Assert-OperationPlanPass $script:V1Excel | Out-Null; $result = [string]$script:V1Excel.Run('ExecuteOperationPlanTest'); Assert-Equal 'mkdir実行結果' $true $result.StartsWith('success:'); Assert-Equal 'nested folder存在' $true (Test-Path -LiteralPath (Join-Path $opRoot 'new\nested'))
    }

    Invoke-OperationCase '40' '既存フォルダと衝突する作成の拒否' {
        $opRoot = New-OperationFixture $operationParent 'case40'; $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot; $ops.Unprotect(); $newRow = $ops.Cells($ops.Rows.Count, 1).End(-4162).Row + 1; $ops.Cells($newRow, 1).Value2 = 'new'; $ops.Cells($newRow, 2).Value2 = 'フォルダ作成'; $ops.Cells($newRow, 12).Value2 = 'dst'; $ops.Protect(); Assert-OperationPlanReject $script:V1Excel | Out-Null
    }

    Invoke-OperationCase '43' '非空フォルダ削除拒否' {
        $opRoot = New-OperationFixture $operationParent 'case43'; $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot; $row = Get-OperationRow $ops (Join-Path $opRoot 'src'); Set-OperationRow $ops $row '空フォルダをゴミ箱へ'; Assert-OperationPlanReject $script:V1Excel | Out-Null
    }

    Invoke-OperationCase '44' '実行直前に内容が追加されたフォルダの削除拒否' {
        $opRoot = New-OperationFixture $operationParent 'case44'; $empty = Join-Path $opRoot 'empty'; $null = New-Item -ItemType Directory -Path $empty -Force; $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot; $row = Get-OperationRow $ops $empty; Set-OperationRow $ops $row '空フォルダをゴミ箱へ'; Assert-OperationPlanPass $script:V1Excel | Out-Null; [IO.File]::WriteAllText((Join-Path $empty 'late.txt'), 'late'); $result = [string]$script:V1Excel.Run('ExecuteOperationPlanTest'); Assert-Equal 'late content拒否' $true (-not $result.StartsWith('success:')); Assert-Equal 'late content保持' $true (Test-Path -LiteralPath (Join-Path $empty 'late.txt'))
    }

    Invoke-OperationCase '45' 'root自身の変更・削除拒否' {
        $opRoot = New-OperationFixture $operationParent 'case45'; $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot; $row = Get-OperationRow $ops $opRoot; Set-OperationRow $ops $row '名前変更/移動' 'renamed'; Assert-OperationPlanReject $script:V1Excel | Out-Null
    }

    Invoke-OperationCase '46' 'source staleの拒否' {
        $opRoot = New-OperationFixture $operationParent 'case46'; $source = Join-Path $opRoot 'src\a.txt'; $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot; $row = Get-OperationRow $ops $source; Set-OperationRow $ops $row '名前変更/移動' 'renamed.txt'; [IO.File]::WriteAllText($source, 'changed-after-scan'); Assert-OperationPlanReject $script:V1Excel | Out-Null
    }

    Invoke-OperationCase '47' '事前確認後のdraft変更によるfingerprint不一致' {
        $opRoot = New-OperationFixture $operationParent 'case47'; $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot; $row = Get-OperationRow $ops (Join-Path $opRoot 'src\a.txt'); Set-OperationRow $ops $row '名前変更/移動' 'first.txt'; Assert-OperationPlanPass $script:V1Excel | Out-Null; Set-OperationRow $ops $row '名前変更/移動' 'second.txt'; $result = [string]$script:V1Excel.Run('ExecuteOperationPlanTest'); Assert-Equal 'draft変更拒否' $true $result.StartsWith('stale|')
    }

    $vbaContract = (@('modOperationPlan.bas','modFileOperations.bas','modRecycleBin.bas','modWorkbookIo.bas' | ForEach-Object { Get-Content -LiteralPath (Join-Path $root ('src\' + $_)) -Raw -Encoding UTF8 }) -join "`n")
    Invoke-OperationCase '41' 'ファイル削除のゴミ箱境界契約' {
        Assert-Equal 'RecycleItemsPhase実装' $true ($vbaContract -match 'RecycleItemsPhase')
        Assert-Equal '完全削除API不使用' $true ($vbaContract -notmatch '\b(Kill|DeleteFileW)\b')
        Assert-Equal 'SHFileOperation undo flag' $true ($vbaContract -match 'FOF_ALLOWUNDO')
        Assert-RecycleRoundTrip '41' $operationParent 'case41' 'src\a.txt' 'ファイルをゴミ箱へ'
    }
    Invoke-OperationCase '42' '空フォルダ削除のゴミ箱境界契約' {
        Assert-Equal '空判定' $true ($vbaContract -match 'IsOperationFolderEmpty')
        Assert-Equal 'phase API' $true ($vbaContract -match 'RecycleItemsPhase')
        Assert-Equal 'SHFileOperation undo flag' $true ($vbaContract -match 'FOF_ALLOWUNDO')
            $opRoot = New-OperationFixture $operationParent 'case42'
            $source = Join-Path $opRoot 'empty'
            $null = New-Item -ItemType Directory -Path $source -Force
        try {
            $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
            $row = Get-OperationRow $ops $source
            Set-OperationRow $ops $row '空フォルダをゴミ箱へ'
            Assert-OperationPlanPass $script:V1Excel | Out-Null
            $result = [string]$script:V1Excel.Run('ExecuteOperationPlanTest')
            Assert-Equal 'case42 recycle execution' $true $result.StartsWith('success:')
            Assert-Equal 'case42 folder removed' $false (Test-Path -LiteralPath $source)
            Assert-Equal 'case42 Recycle Bin item matched' $true ($null -ne (Get-RecycleItemInfo $source))
        } finally {
            if (-not (Test-Path -LiteralPath $source) -and $null -ne (Get-RecycleItemInfo $source)) { [void](Restore-RecycleItemBySource $source) }
            Assert-Equal 'case42 folder restored' $true (Test-Path -LiteralPath $source)
        }
    }
    Invoke-OperationCase '48' 'reversible phase失敗時の逆順rollback' {
        $opRoot = New-OperationFixture $operationParent 'case48'
        $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $rowA = Get-OperationRow $ops (Join-Path $opRoot 'src\a.txt')
        $rowB = Get-OperationRow $ops (Join-Path $opRoot 'src\b.txt')
        Set-OperationRow $ops $rowA '名前変更/移動' 'renamed.txt'
        Set-OperationRow $ops $rowB '名前変更/移動' 'b.txt' 'dst'
        Assert-OperationPlanPass $script:V1Excel | Out-Null
        Remove-Item -LiteralPath (Join-Path $opRoot 'dst')
        $result = [string]$script:V1Excel.Run('ExecuteOperationPlanTest')
        Assert-Equal 'rollback結果' $true $result.StartsWith('failed-before-delete-rolled-back')
        Assert-Equal '先行renameを元へ復元' $true (Test-Path -LiteralPath (Join-Path $opRoot 'src\a.txt'))
        Assert-Equal '中間destination残留なし' $false (Test-Path -LiteralPath (Join-Path $opRoot 'src\renamed.txt'))
    }
    Invoke-OperationCase '49' '残留一時名のjournal復元表と元path衝突' {
        $opRoot = New-OperationFixture $operationParent 'case49'
        $source = Join-Path $opRoot 'src\a.txt'
        $tempName = '.folder-tree-v1_1-tmp-case49-1-ABCDEF01'
        $tempPath = Join-Path $opRoot $tempName
        Move-Item -LiteralPath $source -Destination $tempPath
        $null = $script:V1Excel.Run('AppendExecutionLog', 'case49', 1, '名前変更/移動', $source, (Join-Path $opRoot 'src\renamed.txt'), '実行中', 0, 'journal', '', '一時退避予定', '通常', $opRoot, $tempName, $tempPath, 'ABCDEF01', 'a.txt')
        $null = $script:V1Excel.Run('RunScan', $opRoot, 0, $true, $false)
        $mapped = [string]$script:V1Excel.Run('CheckReservedTemporaryNamesTest', $opRoot)
        Assert-Equal '復元表提示' $true ($mapped.StartsWith('fail|') -and $mapped.Contains($tempPath) -and $mapped.Contains($source))
        [IO.File]::WriteAllText($source, 'collision')
        $null = $script:V1Excel.Run('RunScan', $opRoot, 0, $true, $false)
        $conflict = [string]$script:V1Excel.Run('CheckReservedTemporaryNamesTest', $opRoot)
        Assert-Equal '元path衝突拒否' $true $conflict.Contains('元path衝突')
    }
    Invoke-OperationCase '50' 'delete開始後partial resultのproduction分類' {
        Assert-Equal 'API開始後partial' 'partial-after-delete' ([string]$script:V1Excel.Run('ClassifyDeleteFailureResultTest', $true, '通常'))
        Assert-Equal '永久削除疑い優先' 'possible-permanent-delete' ([string]$script:V1Excel.Run('ClassifyDeleteFailureResultTest', $true, 'possible-permanent-delete'))
        Assert-Equal 'API前失敗' 'failed-before-delete' ([string]$script:V1Excel.Run('ClassifyDeleteFailureResultTest', $false, '通常'))
    }
    Invoke-OperationCase '51' '再入拒否とscan/operation相互排他' {
        $opRoot = New-OperationFixture $operationParent 'case51'
        $result = [string]$script:V1Excel.Run('ProbeOperationReentryTest', $opRoot)
        Assert-Equal 'operation中scan拒否' $true $result.StartsWith('ファイル操作中です。')
        Assert-Equal 'probe後busy解除' $false ([bool]$script:V1Excel.Run('OperationIsBusy'))
    }
    Invoke-OperationCase '52' '実行前後のExcel状態復元' {
        $opRoot = New-OperationFixture $operationParent 'case52'
        $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $row = Get-OperationRow $ops (Join-Path $opRoot 'src\a.txt')
        Set-OperationRow $ops $row '名前変更/移動' 'renamed.txt'
        Assert-OperationPlanPass $script:V1Excel | Out-Null
        $script:V1Excel.ScreenUpdating = $true
        $script:V1Excel.EnableEvents = $true
        $script:V1Excel.Calculation = -4105
        $script:V1Excel.Cursor = -4143
        $script:V1Excel.EnableCancelKey = 1
        $script:V1Excel.StatusBar = 'case52-state'
        $result = [string]$script:V1Excel.Run('ExecuteOperationPlanTest')
        Assert-Equal '実行成功' $true $result.StartsWith('success:')
        Assert-Equal 'ScreenUpdating復元' $true ([bool]$script:V1Excel.ScreenUpdating)
        Assert-Equal 'EnableEvents復元' $true ([bool]$script:V1Excel.EnableEvents)
        Assert-Equal 'Calculation復元' -4105 ([int]$script:V1Excel.Calculation)
        Assert-Equal 'Cursor復元' -4143 ([int]$script:V1Excel.Cursor)
        Assert-Equal 'EnableCancelKey復元' 1 ([int]$script:V1Excel.EnableCancelKey)
        Assert-Equal 'StatusBar復元' 'case52-state' ([string]$script:V1Excel.StatusBar)
        $script:V1Excel.StatusBar = $false
    }
    Invoke-OperationCase '53' 'Unicode・長いpathの実操作' {
        $opRoot = New-SequenceFixture $operationParent 'case53' @('日本語-元.txt')
        $longRelative = ('長いフォルダ-' + ('x' * 120))
        $longDestination = Join-Path $opRoot $longRelative
        $null = New-Item -ItemType Directory -Path $longDestination -Force
        $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $source = Join-Path $opRoot 'src\日本語-元.txt'
        $row = Get-OperationRow $ops $source
        Set-OperationRow $ops $row '名前変更/移動' '日本語-変更.txt' $longRelative
        $preview = Assert-OperationPlanPass $script:V1Excel
        if (-not $preview.StartsWith('pass|')) {
            Write-Host ("  case53 detail: status={0}; message={1}; rootLen={2}; destinationLen={3}" -f $ops.Cells($row, 14).Value2, $ops.Cells($row, 15).Value2, $opRoot.Length, (Join-Path $longDestination '日本語-変更.txt').Length) -ForegroundColor Red
        }
        Assert-Equal '長path事前確認結果' $true $preview.StartsWith('pass|')
        $result = [string]$script:V1Excel.Run('ExecuteOperationPlanTest')
        Assert-Equal 'Unicode長path実行' $true $result.StartsWith('success:')
        Assert-Equal 'Unicode長path destination' $true (Test-Path -LiteralPath (Join-Path $longDestination '日本語-変更.txt'))
    }
    Invoke-OperationCase '54' 'network／recycle不可のfail-closed契約' {
        Assert-Equal 'GetDriveTypeW' $true ($vbaContract -match 'GetDriveTypeW')
        Assert-Equal 'UNC拒否' $true ($vbaContract -match 'UNC/network')
        Assert-Equal 'volume GUID解決' $true ($vbaContract -match 'GetVolumeNameForVolumeMountPointW')
        $uncMessage = ''
        $uncResult = [bool]$script:V1Excel.Run('CanRecycleOperation', '\\server\share\folder-tree-test.txt', 0, $uncMessage)
        Assert-Equal 'UNC実行時fail-closed' $false $uncResult
    }
    Invoke-OperationCase '55' '単一ログschemaと一時名journal実値' {
        $opRoot = New-OperationFixture $operationParent 'case55'
        $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $rowA = Get-OperationRow $ops (Join-Path $opRoot 'src\a.txt')
        $rowB = Get-OperationRow $ops (Join-Path $opRoot 'src\b.txt')
        Set-OperationRow $ops $rowA '名前変更/移動' 'b.txt'
        Set-OperationRow $ops $rowB '名前変更/移動' 'a.txt'
        $preview = Assert-OperationPlanPass $script:V1Excel
        Assert-Equal '事前確認結果' $true $preview.StartsWith('pass|')
        $confirmation = [string]$script:V1Excel.Run('GetOperationConfirmationTextTest')
        foreach ($label in @('名前変更:','ファイル移動:','ファイル削除phase想定所要時間:','エラー件数:','警告件数:','代表的な source -> destination:')) {
            Assert-Equal "確認表示 $label" $true $confirmation.Contains($label)
        }
        Assert-Equal '代表source' $true $confirmation.Contains((Join-Path $opRoot 'src\a.txt'))
        $result = [string]$script:V1Excel.Run('ExecuteOperationPlanTest')
        Assert-Equal 'swap実行成功' $true $result.StartsWith('success:')
        $log = $script:V1Workbook.Worksheets.Item('実行ログ')
        $lastLog = $log.Cells($log.Rows.Count, 1).End(-4162).Row
        $journalRows = @(3..$lastLog | Where-Object { [string]$log.Cells($_, 5).Value2 -eq $opRoot -and -not [string]::IsNullOrWhiteSpace([string]$log.Cells($_, 11).Value2) })
        Assert-Equal '一時path journal件数' $true ($journalRows.Count -ge 2)
        foreach ($logRow in $journalRows) {
            Assert-Equal '一時名実値' $true (-not [string]::IsNullOrWhiteSpace([string]$log.Cells($logRow, 10).Value2))
            Assert-Equal '一時absolute path実値' $true ([IO.Path]::IsPathFullyQualified([string]$log.Cells($logRow, 11).Value2))
            Assert-Equal 'nonce実値' $true ([string]$log.Cells($logRow, 20).Value2 -match '^[0-9A-F]{16}$')
            Assert-Equal '元の名前はpathでない' $false ([string]$log.Cells($logRow, 9).Value2).Contains('\')
        }
    }
    Invoke-OperationCase '56' '実行後再走査でツリー同期' {
        $opRoot = New-OperationFixture $operationParent 'case56'
        $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $row = Get-OperationRow $ops (Join-Path $opRoot 'src\a.txt')
        Set-OperationRow $ops $row '名前変更/移動' 'rescanned.txt'
        Assert-OperationPlanPass $script:V1Excel | Out-Null
        $result = [string]$script:V1Excel.Run('ExecuteOperationPlanTest')
        Assert-Equal '実行成功' $true $result.StartsWith('success:')
        Assert-Equal 'scan ready' $true ([bool]$script:V1Excel.Run('IsScanReadyForRoot', $opRoot))
        $tree = $script:V1Workbook.Worksheets.Item('ツリー')
        $lastTree = $tree.Cells($tree.Rows.Count, 7).End(-4162).Row
        $treePaths = @($tree.Range("G6:G$lastTree").Value2 | ForEach-Object { [string]$_ })
        Assert-Equal '再走査destination反映' $true ($treePaths -contains (Join-Path $opRoot 'src\rescanned.txt'))
        Assert-Equal '再走査source消失反映' $false ($treePaths -contains (Join-Path $opRoot 'src\a.txt'))
    }
    Invoke-OperationCase '57' 'embedded複数moduleとsource hash契約' { $br = Get-Content -LiteralPath $CandidateBuildReceiptPath -Raw -Encoding UTF8 | ConvertFrom-Json; Assert-Equal '5 source modules' 5 @($br.sourceModules).Count; Assert-Equal 'embedded hash receipt' $true ($null -ne $br.embeddedBodyHashes) }
    Invoke-OperationCase '58' '一連のcandidate生成・rollback試験後のsrc不変契約' {
        foreach ($srcName in $allSourceHashesBeforeOpen.Keys) {
            $currentHash = (Get-FileHash -LiteralPath (Join-Path $root "src\$srcName") -Algorithm SHA256).Hash
            Assert-Equal "src/$srcName hash不変" $allSourceHashesBeforeOpen[$srcName] $currentHash
        }
    }
    Invoke-OperationCase '59' '昇格失敗時backup復元契約' { $builderSource = Get-Content -LiteralPath $buildScript -Raw -Encoding UTF8; Assert-Equal 'SimulatePromotionFailure' $true ($builderSource -match 'SimulatePromotionFailure'); Assert-Equal 'backup copy' $true ($builderSource -match 'backup'); Assert-Equal '初回配置rollback除去' $true ($builderSource -match 'elseif\(-not \$oldHash[\s\S]*Remove-Item -LiteralPath \$dest') }
    Invoke-OperationCase '60' 'receipt hash連鎖契約' { $br = Get-Content -LiteralPath $CandidateBuildReceiptPath -Raw -Encoding UTF8 | ConvertFrom-Json; Assert-Equal 'builder hash' $true (-not [string]::IsNullOrWhiteSpace([string]$br.builderSha256)); Assert-Equal 'candidate hash' $true (-not [string]::IsNullOrWhiteSpace([string]$br.candidateSha256)); $builderSource = Get-Content -LiteralPath $buildScript -Raw -Encoding UTF8; Assert-Equal '全source hash再照合' $true ($builderSource -match 'Current source hash does not match candidate build receipt'); Assert-Equal 'readiness receipt必須' $true ($builderSource -match 'CandidateReadinessReceiptPath'); Assert-Equal 'gate/build receipt結合' $true ($builderSource -match 'candidateBuildReceiptSha256') }
    Invoke-OperationCase '61' 'ゴミ箱Shell path上限の実行時拒否' {
        $longPath = 'C:\' + ('x' * 270)
        $probe = [string]$script:V1Excel.Run('CanRecycleOperationTest', $longPath, 0)
        Assert-Equal '長path拒否' $true $probe.StartsWith('blocked|')
        Assert-Equal '長path理由' $true ($probe -match '長|上限|path')
    }
    Invoke-OperationCase '62' 'sheet protection失敗時のfail-closed' {
        $ops = $script:V1Workbook.Worksheets.Item('変更操作')
        $before = [string]$ops.Range('B5').Value2
        $probe = [string]$script:V1Excel.Run('WorkbookIoProtectionFailureProbeTest')
        Assert-Equal '保護解除失敗を通知' $true $probe.StartsWith('pass|')
        Assert-Equal '保護維持' $true ([bool]$ops.ProtectContents)
        Assert-Equal '書込み停止' $before ([string]$ops.Range('B5').Value2)
    }
    Invoke-OperationCase '63' 'delete phase直前空再確認と分類' {
        $opRoot = New-OperationFixture $operationParent 'case63'
        $empty = Join-Path $opRoot 'empty'
        $null = New-Item -ItemType Directory -Path $empty -Force
        Assert-Equal 'production空判定' $true ([bool]$script:V1Excel.Run('IsOperationFolderEmpty', $empty))
        [IO.File]::WriteAllText((Join-Path $empty 'late.txt'), 'late')
        Assert-Equal 'production直前非空判定' $false ([bool]$script:V1Excel.Run('IsOperationFolderEmpty', $empty))
        Assert-Equal 'API開始後未実行分類' 'partial-after-delete' ([string]$script:V1Excel.Run('ClassifyDeleteFailureResultTest', $true, '通常'))
    }
    Invoke-OperationCase '64' 'volume／policyの削除事前判定契約' {
        Assert-Equal 'NoRecycle policy' $true ($vbaContract -match 'NoRecycleFiles')
        Assert-Equal 'fixed volume' $true ($vbaContract -match 'DRIVE_FIXED')
        Assert-Equal 'volume GUID API' $true ($vbaContract -match 'GetVolumeNameForVolumeMountPointW')
        Assert-Equal 'capacity policy' $true ($vbaContract -match 'MaxCapacity')
        Assert-Equal 'NukeOnDelete policy' $true ($vbaContract -match 'NukeOnDelete')
        Assert-Equal 'RecycleBinSize policy' $true ($vbaContract -match 'RecycleBinSize')
        Assert-Equal 'safety margin' $true ($vbaContract -match 'SafetyMarginBytes')
        $opRoot = New-OperationFixture $operationParent 'case64'
        $ops = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $opRoot
        $row = Get-OperationRow $ops (Join-Path $opRoot 'src\a.txt')
        Set-OperationRow $ops $row 'ファイルをゴミ箱へ'
        Assert-OperationPlanPass $script:V1Excel | Out-Null
    }
    Invoke-OperationCase '65' 'failure-only test hook／非UI受入entrypoint契約' {
        $scanSource = Get-Content -LiteralPath (Join-Path $root 'src\modFolderTree.bas') -Raw -Encoding UTF8
        $ioSource = Get-Content -LiteralPath (Join-Path $root 'src\modWorkbookIo.bas') -Raw -Encoding UTF8
        Assert-Equal 'test hook reset' $true ($scanSource -match 'ResetTestHooks')
        Assert-Equal 'normal UI reset' $true ($scanSource -match 'ResetTestHooks\s*\r?\n\s*summary = RunScan')
        Assert-Equal 'draft受入はproduction core使用' $true ($ioSource -match 'WorkbookIoCreateOperationDraftTest = WorkbookIoCreateOperationDraftCore\(rootPath, False\)')
        $draftTestRoot = New-OperationFixture $operationParent 'case65-draft'
        $null = $script:V1Excel.Run('RunScan', $draftTestRoot, 0, $true, $false)
        Assert-Equal 'draft非UI受入entrypoint実行' $true ([bool]$script:V1Excel.Run('WorkbookIoCreateOperationDraftTest', $draftTestRoot))
        Assert-Equal 'sequence受入はproduction core使用' $true ($ioSource -match 'WorkbookIoSetSequencePreviewTest[\s\S]*WorkbookIoApplySequencePreview\(mode, startNumber, selected, resultMessage\)')
        Assert-Equal '上限超過state' $true ($ioSource -match 'Range\("B5"\)\.Value2 = "一覧上限超過"')
        Assert-Equal 'Recycle性能入口はworkspace fixture限定' $true ($vbaContract -match 'GetRecyclePerformanceRootPrefix')
        Assert-Equal 'Recycle性能入口はproduction phase使用' $true ($vbaContract -match 'RecycleDirectoryFilesPerformanceTest[\s\S]*RecycleItemsPhase\(sources, sizes, modified, isFolder, itemCount')
        Assert-Equal 'snapshot差分core共有' $true ($vbaContract -match 'ValidateRecycleSnapshotDelta\(beforeItems, afterItems, expected, errorMessage\)')
        Assert-Equal '大規模snapshot sort' $true ($vbaContract -match 'SortSnapshotKeys keys')
        $rejectedPerformancePath = [string]$script:V1Excel.Run('RecycleDirectoryFilesPerformanceTest', 'C:\Windows')
        Assert-Equal 'workspace外性能fixture拒否' $true $rejectedPerformancePath.StartsWith('fail|performance fixture path rejected')
    }
    Invoke-OperationCase '66' 'orphan・重複対応の残留予約prefix拒否' {
        $opRoot = New-OperationFixture $operationParent 'case66'
        $tempName = '.folder-tree-v1_1-tmp-case66-1-ABCDEF02'
        $tempPath = Join-Path $opRoot $tempName
        [IO.File]::WriteAllText($tempPath, 'orphan')
        $null = $script:V1Excel.Run('RunScan', $opRoot, 0, $true, $false)
        $orphan = [string]$script:V1Excel.Run('CheckReservedTemporaryNamesTest', $opRoot)
        Assert-Equal 'orphan拒否' $true ($orphan.StartsWith('fail|') -and $orphan.Contains('orphan'))
        $source = Join-Path $opRoot 'src\missing.txt'
        1..2 | ForEach-Object {
            $null = $script:V1Excel.Run('AppendExecutionLog', 'case66', $_, '名前変更/移動', $source, (Join-Path $opRoot 'src\renamed.txt'), '実行中', 0, 'journal', '', '一時退避予定', '通常', $opRoot, $tempName, $tempPath, 'ABCDEF02', 'missing.txt')
        }
        $duplicate = [string]$script:V1Excel.Run('CheckReservedTemporaryNamesTest', $opRoot)
        Assert-Equal '重複対応拒否' $true $duplicate.Contains('重複対応')
    }
    Invoke-OperationCase '67' 'locale非依存Recycle Bin property契約' {
        Assert-Equal 'ExtendedProperty' $true ($vbaContract -match 'ExtendedText\(item, "System\.Recycle\.DeletedFrom"\)')
        Assert-Equal 'source path matching' $true ($vbaContract -match 'NormalizeRecyclePath')
        Assert-Equal 'pre/post snapshot' $true ($vbaContract -match 'TakeRecycleSnapshot')
        Assert-Equal 'size/mtime照合' $true (($vbaContract -match 'System.Size') -and ($vbaContract -match 'System.DateModified'))
        Assert-RecycleRoundTrip '67' $operationParent 'case67' 'src\b.txt' 'ファイルをゴミ箱へ'
    }
    Invoke-OperationCase '68' 'Recycle Bin contamination拒否契約' {
        Assert-Equal 'post verification' $true ($vbaContract -match 'beforeCanonical' -and $vbaContract -match 'afterCanonical')
        Assert-Equal 'contamination state' $true ($vbaContract -match 'recycle-verification-contaminated')
        Assert-Equal 'source absence is not sole proof' $true ($vbaContract -match 'VerifyRecycleItemValue')
        Assert-Equal 'snapshot receipt fingerprint' $true ($vbaContract -match 'SnapshotFingerprint' -and $vbaContract -match 'GetRecycleSnapshotInfo')
        $preservedRoot = New-OperationFixture $operationParent 'case68-preserved'
        $preservedSource = Join-Path $preservedRoot 'src\a.txt'
        $targetRoot = New-OperationFixture $operationParent 'case68-target'
        $targetSource = Join-Path $targetRoot 'src\a.txt'
        try {
            $preservedOps = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $preservedRoot
            $preservedRow = Get-OperationRow $preservedOps $preservedSource
            Set-OperationRow $preservedOps $preservedRow 'ファイルをゴミ箱へ'
            Assert-OperationPlanPass $script:V1Excel | Out-Null
            $preserveResult = [string]$script:V1Excel.Run('ExecuteOperationPlanTest')
            Assert-Equal 'case68 preserve execution' $true $preserveResult.StartsWith('success:')
            Assert-Equal 'case68 preserved item present' $true ($null -ne (Get-RecycleItemInfo $preservedSource))

            $targetOps = Prepare-OperationDraft $script:V1Excel $script:V1Workbook $targetRoot
            $targetRow = Get-OperationRow $targetOps $targetSource
            Set-OperationRow $targetOps $targetRow 'ファイルをゴミ箱へ'
            Assert-OperationPlanPass $script:V1Excel | Out-Null
            $targetResult = [string]$script:V1Excel.Run('ExecuteOperationPlanTest')
            Assert-Equal 'case68 target execution' $true $targetResult.StartsWith('success:')
            Assert-Equal 'case68 preserved item retained' $true ($null -ne (Get-RecycleItemInfo $preservedSource))
            Assert-Equal 'case68 target item present' $true ($null -ne (Get-RecycleItemInfo $targetSource))
        } finally {
            if (-not (Test-Path -LiteralPath $targetSource) -and $null -ne (Get-RecycleItemInfo $targetSource)) { [void](Restore-RecycleItemBySource $targetSource) }
            if (-not (Test-Path -LiteralPath $preservedSource) -and $null -ne (Get-RecycleItemInfo $preservedSource)) { [void](Restore-RecycleItemBySource $preservedSource) }
            Assert-Equal 'case68 target restored' $true (Test-Path -LiteralPath $targetSource)
            Assert-Equal 'case68 preserved restored' $true (Test-Path -LiteralPath $preservedSource)
        }
    }
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
    $resolvedRunRoot = [IO.Path]::GetFullPath($runRoot)
    if ($resolvedRunRoot -ne $runRoot -or $resolvedRunRoot -eq $workspaceFull -or -not $resolvedRunRoot.StartsWith($workspaceFull + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing cleanup outside the created test work root: $resolvedRunRoot"
    }
    $cleanupItem = Get-Item -LiteralPath $resolvedRunRoot -Force
    if (($cleanupItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Refusing cleanup of reparse point: $resolvedRunRoot" }
    $cleanupEntries = @(Get-ChildItem -LiteralPath $resolvedRunRoot -Recurse -Force)
    $cleanupReparse = @($cleanupEntries | Where-Object { ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 })
    foreach ($entry in $cleanupReparse) {
        $resolvedEntry = [IO.Path]::GetFullPath($entry.FullName)
        if ($resolvedEntry -eq $resolvedRunRoot -or -not $resolvedEntry.StartsWith($resolvedRunRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing cleanup of reparse point outside the created test work root: $resolvedEntry"
        }
        if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0) {
            throw "Reparse point state changed during cleanup: $resolvedEntry"
        }
        Remove-Item -LiteralPath $resolvedEntry -Force
        if (Test-Path -LiteralPath $resolvedEntry) { throw "Test reparse cleanup failed: $resolvedEntry" }
    }
    $cleanupEntries = @(Get-ChildItem -LiteralPath $resolvedRunRoot -Recurse -Force)
    $cleanupReparse = @($cleanupEntries | Where-Object { ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 })
    if ($cleanupReparse.Count -gt 0) { throw "Refusing cleanup because reparse points remain: $resolvedRunRoot" }
    foreach ($entry in $cleanupEntries | Where-Object { $_.PSIsContainer -and ($_.Attributes -band [IO.FileAttributes]::ReadOnly) }) {
        $entry.Attributes = $entry.Attributes -band (-bnot [IO.FileAttributes]::ReadOnly)
    }
    $cleanupSucceeded = $false
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        try {
            Remove-Item -LiteralPath $resolvedRunRoot -Recurse -Force
            if (-not (Test-Path -LiteralPath $resolvedRunRoot)) {
                $cleanupSucceeded = $true
                break
            }
        } catch {
            if ($attempt -eq 5) { throw }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Start-Sleep -Milliseconds 500
    }
    if (-not $cleanupSucceeded) { throw "Test cleanup failed: $resolvedRunRoot" }
}

Write-Host ''
$gateResult = if ($script:Failures.Count -eq 0) { 'pass' } else { 'fail' }
if ($null -ne $gateReceiptFull) {
    $candidateHash = (Get-FileHash -LiteralPath $WorkbookPath -Algorithm SHA256).Hash
    $sourceHash = (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash
    $caseResults = [ordered]@{}
    foreach ($caseId in $requiredCaseIds) {
        if ($script:CaseFailures.ContainsKey($caseId)) {
            $caseResults[$caseId] = 'fail'
        } elseif ($script:CaseFailures.ContainsKey('setup')) {
            $caseResults[$caseId] = 'not-run'
        } elseif ($script:CaseNotRun.ContainsKey($caseId)) {
            $caseResults[$caseId] = 'not-run'
        } elseif ($script:CaseExecuted.ContainsKey($caseId)) {
            $caseResults[$caseId] = 'pass'
        } else {
            $caseResults[$caseId] = 'not-run'
        }
    }
    $allRequiredPassed = @($caseResults.Values | Where-Object { $_ -ne 'pass' }).Count -eq 0
    $gate = [ordered]@{
        schemaVersion = 1
        generatedAtUtc = [DateTime]::UtcNow.ToString('o')
        candidatePath = [IO.Path]::GetRelativePath($root, [IO.Path]::GetFullPath($WorkbookPath))
        candidateSha256 = $candidateHash
        sourcePath = [IO.Path]::GetRelativePath($root, [IO.Path]::GetFullPath($SourcePath))
        sourceSha256 = $sourceHash
        candidateBuildReceiptPath = [IO.Path]::GetRelativePath($root, $candidateBuildReceiptFull)
        candidateBuildReceiptSha256 = $candidateBuildReceiptSha256
        allPassed = ($gateResult -eq 'pass' -and $allRequiredPassed)
        requiredCaseIds = $requiredCaseIds
        optionalCaseIds = $optionalCaseIds
        cases = $caseResults
    }
    New-Item -ItemType Directory -Path (Split-Path -Parent $gateReceiptFull) -Force | Out-Null
    $gate | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $gateReceiptFull -Encoding UTF8
    Write-Host "candidateGateReceipt=$gateReceiptFull"
}
$requiredNotRun = @($requiredCaseIds | Where-Object { $script:CaseNotRun.ContainsKey($_) })
if ($requiredNotRun.Count -gt 0) {
    $message = "必須ケースが未実行のためgate不成立: $($requiredNotRun -join ',')"
    $script:Failures += $message
    Write-Host $message -ForegroundColor Red
}
if ($script:Failures.Count -gt 0) {
    Write-Host ("FAILED: {0} 件" -f $script:Failures.Count) -ForegroundColor Red
    exit 1
}
Write-Host 'PASSED: すべての検証に成功しました。' -ForegroundColor Green
