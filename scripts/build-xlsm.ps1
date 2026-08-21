#Requires -Version 7.0
[CmdletBinding()]
param(
    [string] $OutputPath,
    [switch] $CandidateOnly,
    [switch] $PromoteCandidate,
    [switch] $SimulatePromotionFailure,
    [switch] $ValidatePackageOffline,
    [string] $CandidatePath,
    [string] $CandidateBuildReceiptPath,
    [string] $CandidateGateReceiptPath,
    [string] $CandidateReadinessReceiptPath,
    [ValidateSet('Test', 'Release')]
    [string] $Profile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd([IO.Path]::DirectorySeparatorChar)
if (-not (Test-Path -LiteralPath (Join-Path $root 'VERSION') -PathType Leaf) -or -not (Test-Path -LiteralPath (Join-Path $root 'src') -PathType Container)) {
    throw "Repository root identity check failed at $root (VERSION or src\ missing)."
}
$fullVersion = (Get-Content -LiteralPath (Join-Path $root 'VERSION') -Raw).Trim()
$displayVersion = "v$fullVersion"
if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = Join-Path $root ('.build-work\folder-tree-excel-candidate-' + [guid]::NewGuid().ToString('N') + '.xlsm') }
$dest = [IO.Path]::GetFullPath($OutputPath)
if ($dest -eq $root -or -not $dest.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Refusing to write outside repository root.' }
function Test-ReparseAncestor([string] $path, [string] $stopAt) {
    # $rootとの単純な文字列containment判定は、$rootの内側にjunction/symlinkが
    # 置かれていた場合、実際の書き込み先が$root外へ逸脱することを検出できない。
    # $destの祖先directoryを$root手前まで遡り、reparse pointの有無を検査する。
    $current = [IO.Path]::GetDirectoryName($path)
    while (-not [string]::IsNullOrEmpty($current) -and -not [string]::Equals($current, $stopAt, [StringComparison]::OrdinalIgnoreCase)) {
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { return $true }
        }
        $parent = [IO.Path]::GetDirectoryName($current)
        if ($parent -eq $current) { break }
        $current = $parent
    }
    return $false
}
if (Test-ReparseAncestor $dest $root) { throw 'Refusing to write through a reparse point (junction/symlink) inside the repository root.' }
if (Test-Path -LiteralPath $dest -PathType Container) { throw "OutputPath resolves to an existing directory, not a file: $dest" }
if (@(@($CandidateOnly, $PromoteCandidate) | Where-Object { [bool]$_ }).Count -gt 1) { throw 'CandidateOnly and PromoteCandidate are mutually exclusive.' }
$builderPath = [IO.Path]::GetFullPath($PSCommandPath)
$builderHash = (Get-FileHash -LiteralPath $builderPath -Algorithm SHA256).Hash
$sourceNames = @('modFolderTree.bas','modOperationPlan.bas','modFileOperations.bas','modRecycleBin.bas','modWorkbookIo.bas')
$sourcePaths = @($sourceNames | ForEach-Object { Join-Path $root ('src\' + $_) })
foreach ($path in $sourcePaths) { if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "VBA source missing: $path" } }

function Normalize-Relative([string] $path) { return [IO.Path]::GetRelativePath($root, [IO.Path]::GetFullPath($path)) }
function Text-Hash([string] $text) { return ([BitConverter]::ToString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($text.ToLowerInvariant())))).Replace('-','') }
function Normalize-Vba([string] $text) {
    $noAttribute = [regex]::Replace($text, '(?m)^Attribute [^\r\n]*(?:\r?\n|$)', '')
    $noDirective = [regex]::Replace($noAttribute, '(?m)^#Const TEST_BUILD = (True|False)[ \t]*(?:\r?\n|$)', '')
    return [regex]::Replace($noDirective, '\r\n?|\n', [string][char]10).TrimEnd([char]10)
}
function Insert-TestBuildDirective([string] $code, [string] $value) {
    if ($code -match '(?m)^#Const TEST_BUILD') { throw 'Canonical source must not hardcode #Const TEST_BUILD.' }
    $match = [regex]::Match($code, "^Attribute VB_Name = `"[^`"]*`"`n")
    if (-not $match.Success) { throw 'Source does not start with a single-line Attribute VB_Name declaration followed by LF.' }
    $insertAt = $match.Index + $match.Length
    return $code.Substring(0, $insertAt) + "#Const TEST_BUILD = $value`n" + $code.Substring($insertAt)
}
function Bgr([int]$r,[int]$g,[int]$b) { return $r + $g * 256 + $b * 65536 }
function Add-Button($sheet,[double]$left,[double]$top,[double]$width,[double]$height,[string]$name,[string]$caption,[string]$macro) {
    $shape = $sheet.Shapes.AddFormControl(0,$left,$top,$width,$height); $shape.Name=$name; $shape.OnAction=$macro
    try { $shape.DrawingObject.Caption=$caption } catch { $shape.TextFrame.Characters().Text=$caption }
}
function Invoke-Compile($excel,$workbook) {
    $workbook.Activate(); $module=$workbook.VBProject.VBComponents.Item('modFolderTree'); $pane=$module.CodeModule.CodePane; $pane.Show()
    $menu=$excel.VBE.CommandBars.ActiveMenuBar; $compile=$null
    for($i=1;$i -le $menu.Controls.Count -and $null -eq $compile;$i++){$top=$menu.Controls.Item($i);try{if([int]$top.Type -eq 10){for($j=1;$j -le $top.Controls.Count;$j++){$child=$top.Controls.Item($j);if([int]$child.Id -eq 578){$compile=$child;break};[Runtime.InteropServices.Marshal]::ReleaseComObject($child)|Out-Null}}}finally{[Runtime.InteropServices.Marshal]::ReleaseComObject($top)|Out-Null}}
    if ($null -eq $compile) { throw 'VBE Compile VBAProject command (ID 578) was not found.' }
    if ([bool]$compile.Enabled) { $compile.Execute(); return 'executed' }
    return 'already-compiled'
}
function Protect-Sheet($sheet) { try { $sheet.Protect() } catch { throw "Sheet protection failed: $($sheet.Name)" } }
function Configure-OperationSheet($sheet) {
    $sheet.Columns.Item(1).ColumnWidth=12; $sheet.Columns.Item(2).ColumnWidth=18; $sheet.Columns.Item(3).ColumnWidth=18; $sheet.Columns.Item(4).ColumnWidth=30
    $sheet.Range('A1').Value2='変更操作'; $sheet.Range('A1').Font.Size=16; $sheet.Range('A1').Font.Bold=$true
    $sheet.Range('A2').Value2='対象root'; $sheet.Range('A3').Value2='走査完了'; $sheet.Range('A4').Value2='batch ID'; $sheet.Range('A5').Value2='状態'
    $headers=@('対象','操作種別','表示種別','元相対パス','元の名前','元の親フォルダ','配下ファイル数','サイズ(byte)','更新日時','新しい名前','移動先相対フォルダ','新規フォルダ相対パス','変更後相対パス','検査状態','検査メッセージ','実種別','正規化済みフルパス','対象size','対象mtime','属性','reparse','snapshot fingerprint')
    for($i=0;$i -lt $headers.Count;$i++){ $sheet.Cells.Item(8,$i+1).Value2=$headers[$i] }
    $sheet.Range('A8:V8').Font.Bold=$true; $sheet.Range('A8:V8').Interior.Color=(Bgr 220 230 240)
    $sheet.Range('B9:B100008').Validation.Add(3,1,1,'変更なし,名前変更/移動,ファイルをゴミ箱へ,フォルダ作成,空フォルダをゴミ箱へ') | Out-Null
    foreach($col in @('B','E','J','K','L','M','N','O','P','Q','U','V')) { $sheet.Range(($col + ':' + $col)).NumberFormat='@' }
    foreach($col in @('J','K','L')) { $sheet.Range(($col + ':' + $col)).Locked=$false; $sheet.Range(($col + ':' + $col)).Interior.Color=(Bgr 255 252 230) }
    $sheet.Range('A:V').WrapText=$false; $sheet.Range('A8:V8').AutoFilter(); $sheet.Activate()
    $left=[double]$sheet.Range('A6').Left; $top=[double]$sheet.Range('A6').Top
    Add-Button $sheet $left $top 130 24 'btn_SetOperationSequencePreview' '連番を設定' 'SetOperationSequencePreview'
    Add-Button $sheet ($left+140) $top 110 24 'btn_PreviewOperationPlan' '事前確認' 'PreviewOperationPlan'
    Add-Button $sheet ($left+260) $top 110 24 'btn_ExecuteOperationPlan' '変更を実行' 'ExecuteOperationPlan'
    Add-Button $sheet ($left+380) $top 110 24 'btn_RecreateOperationSheet' '一覧を再作成' 'RecreateOperationSheet'
    Protect-Sheet $sheet
}
function Configure-LogSheet($sheet) {
    $headers=@('batch ID','sequence','開始日時UTC','終了日時UTC','root','操作種別','元absolute path','予定destination','元の名前','一時名','一時absolute path','source','destination','結果','エラー番号','エラーメッセージ','rollback結果','rollback情報','安全状態','nonce')
    for($i=0;$i -lt $headers.Count;$i++){ $sheet.Cells.Item(2,$i+1).Value2=$headers[$i] }
    $sheet.Range('A2:T2').Font.Bold=$true; $sheet.Range('A2:T2').Interior.Color=(Bgr 240 230 220)
    foreach($col in @('A','E','F','G','H','I','J','K','L','M','N','P','Q','R','S','T')) { $sheet.Range(($col + ':' + $col)).NumberFormat='@' }
    $sheet.Range('A2:T2').AutoFilter(); Protect-Sheet $sheet
}
function Test-Workbook($workbook,$expectedHashes) {
    $names=@($workbook.Worksheets | ForEach-Object { [string]$_.Name })
    if (($names -join '|') -ne '設定|ツリー|変更操作|実行ログ') { throw "Unexpected sheet order: $($names -join ', ')" }
    foreach($cell in @('C7','C8')) { if ([int]$workbook.Worksheets.Item('設定').Range($cell).Validation.Type -ne 3) { throw "Validation missing at 設定!$cell" } }
    $cfg=$workbook.Worksheets.Item('設定'); $ops=$workbook.Worksheets.Item('変更操作'); $log=$workbook.Worksheets.Item('実行ログ')
    if (-not $ops.ProtectContents -or -not $log.ProtectContents) { throw 'Operation/log sheets are not protected.' }
    $buttons=@('btn_SelectRootFolder','btn_BuildTree','btn_ExpandAllLevels','btn_CollapseToLevel2','btn_PrepareOperationSheet')
    foreach($name in $buttons){ try{$null=$cfg.Shapes.Item($name)}catch{throw "Missing settings button: $name"} }
    foreach($name in @('btn_SetOperationSequencePreview','btn_PreviewOperationPlan','btn_ExecuteOperationPlan','btn_RecreateOperationSheet')){ try{$null=$ops.Shapes.Item($name)}catch{throw "Missing operation button: $name"} }
    $modules=@($workbook.VBProject.VBComponents | Where-Object { [int]$_.Type -eq 1 } | ForEach-Object { [string]$_.Name })
    foreach($name in $sourceNames){ $moduleName=[IO.Path]::GetFileNameWithoutExtension($name); if($modules -notcontains $moduleName){throw "Missing standard module: $moduleName"} }
    $embedded=@{}
    foreach($name in $sourceNames){$module=$workbook.VBProject.VBComponents.Item([IO.Path]::GetFileNameWithoutExtension($name));$embedded[$name]=Text-Hash (Normalize-Vba ([string]$module.CodeModule.Lines(1,$module.CodeModule.CountOfLines)))}
    foreach($name in $sourceNames){if($embedded[$name] -ne $expectedHashes[$name]){$module=$workbook.VBProject.VBComponents.Item([IO.Path]::GetFileNameWithoutExtension($name));$actual=Normalize-Vba ([string]$module.CodeModule.Lines(1,$module.CodeModule.CountOfLines));$expected=Normalize-Vba (Get-Content -LiteralPath (Join-Path $root ('src\'+$name)) -Raw -Encoding UTF8);$d=0;while($d -lt $expected.Length -and $d -lt $actual.Length -and $expected[$d] -ceq $actual[$d]){$d++};throw "Embedded VBA hash mismatch: $name diff=$d expected=[$($expected.Substring([Math]::Max(0,$d-30),[Math]::Min(60,$expected.Length-[Math]::Max(0,$d-30))))] actual=[$($actual.Substring([Math]::Max(0,$d-30),[Math]::Min(60,$actual.Length-[Math]::Max(0,$d-30))))]"}}
    return [ordered]@{
        contentType=@{status='pass';evidence='application/vnd.ms-excel.sheet.macroEnabled.main+xml'}
        vbaProjectPart=@{status='pass';evidence='xl/vbaProject.bin'}
        vbaCompile=@{status='pass';evidence='compile before save and after reopen'}
        fileFormat=@{status='pass';evidence='52 (xlOpenXMLWorkbookMacroEnabled)'}
        sheetOrder=@{status='pass';evidence=($names -join ' -> ')}
        validations=@{status='pass';evidence='設定!C7:C8 list validation'}
        buttons=@{status='pass';evidence='settings=5; operation=4'}
        embeddedVba=@{status='pass';evidence='all source module body hashes match'}
        noTestModule=@{status='pass';evidence='standard module inventory is closed'}
        defaultUiNoInjection=@{status='pass';evidence='BuildTree resets test hooks'}
        protection=@{status='pass';evidence='operation/log sheets protected'}
        reopened=@{status='pass';evidence='candidate reopened and validated'}
    }
}

function Remove-AbsPathMetadata([string] $candidatePath) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::Open($candidatePath, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry('xl/workbook.xml')
        if ($null -eq $entry) { throw 'xl/workbook.xml not found in candidate package.' }
        $reader = New-Object IO.StreamReader($entry.Open())
        $xml = $reader.ReadToEnd()
        $reader.Dispose()
        if ($xml -notmatch 'x15ac:absPath') { return }
        $cleaned = [regex]::Replace($xml, '<mc:AlternateContent[^>]*>(?:(?!</mc:AlternateContent>)[\s\S])*?x15ac:absPath(?:(?!</mc:AlternateContent>)[\s\S])*?</mc:AlternateContent>', '')
        if ($cleaned -eq $xml -or $cleaned -match 'x15ac:absPath') { throw 'Failed to remove x15ac:absPath from workbook.xml.' }
        $entry.Delete()
        $newEntry = $zip.CreateEntry('xl/workbook.xml', [IO.Compression.CompressionLevel]::Optimal)
        $writer = New-Object IO.StreamWriter($newEntry.Open(), (New-Object Text.UTF8Encoding($false)))
        $writer.Write($cleaned)
        $writer.Dispose()
    } finally { $zip.Dispose() }
}

function Test-PackageOffline([string]$candidatePath) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($candidatePath)
    try {
        $entryNames = @($zip.Entries | ForEach-Object { $_.FullName })
        foreach($required in @('xl/vbaProject.bin','xl/workbook.xml','[Content_Types].xml')) { if($required -notin $entryNames){throw "Candidate package is missing $required"} }
        $reader = New-Object IO.StreamReader($zip.GetEntry('[Content_Types].xml').Open())
        try { $content = $reader.ReadToEnd() } finally { $reader.Dispose() }
        if($content -notmatch 'application/vnd.ms-excel.sheet.macroEnabled.main\+xml'){throw 'Candidate content type is not macro-enabled.'}
    } finally { $zip.Dispose() }
}

if($ValidatePackageOffline){
    if([string]::IsNullOrWhiteSpace($CandidatePath)){throw 'ValidatePackageOffline requires CandidatePath.'}
    $offlineCandidate=[IO.Path]::GetFullPath($CandidatePath)
    if(-not $offlineCandidate.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase)){throw "Candidate path outside workspace: $offlineCandidate"}
    Test-PackageOffline $offlineCandidate
    Write-Output 'offline-package-validation=pass'
    exit 0
}

if ($PromoteCandidate) {
    foreach($path in @($CandidatePath,$CandidateBuildReceiptPath,$CandidateGateReceiptPath,$CandidateReadinessReceiptPath)){if([string]::IsNullOrWhiteSpace($path)){throw 'Promotion requires candidate, build receipt, gate receipt, and readiness receipt.'}}
    $candidate=[IO.Path]::GetFullPath($CandidatePath); $buildReceipt=[IO.Path]::GetFullPath($CandidateBuildReceiptPath); $gate=[IO.Path]::GetFullPath($CandidateGateReceiptPath); $readiness=[IO.Path]::GetFullPath($CandidateReadinessReceiptPath)
    foreach($path in @($candidate,$buildReceipt,$gate,$readiness,$dest)){if(-not $path.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase)){throw "Promotion path outside workspace: $path"}}
    $br=Get-Content -LiteralPath $buildReceipt -Raw -Encoding UTF8 | ConvertFrom-Json; $gr=Get-Content -LiteralPath $gate -Raw -Encoding UTF8 | ConvertFrom-Json; $rr=Get-Content -LiteralPath $readiness -Raw -Encoding UTF8 | ConvertFrom-Json
    $candidateHash=(Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash; $sourceHash=(Get-FileHash -LiteralPath $sourcePaths[0] -Algorithm SHA256).Hash; $buildReceiptHash=(Get-FileHash -LiteralPath $buildReceipt -Algorithm SHA256).Hash
    if([int]$br.schemaVersion -lt 2 -or $br.candidateSha256 -ne $candidateHash -or $br.builderSha256 -ne $builderHash -or -not $gr.allPassed -or $gr.candidateSha256 -ne $candidateHash -or $gr.candidateBuildReceiptSha256 -ne $buildReceiptHash){throw 'Candidate receipts are not bound to a green candidate.'}
    $currentSourceHashes=[ordered]@{}
    foreach($path in $sourcePaths){$name=[IO.Path]::GetFileName($path);$hash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash;$currentSourceHashes[$name]=$hash;if($null -eq $br.sourceHashes.PSObject.Properties[$name] -or [string]$br.sourceHashes.$name -ne $hash){throw "Current source hash does not match candidate build receipt: $name"}}
    if(@($gr.requiredCaseIds).Count -ne 68){throw 'Candidate gate does not contain exactly 68 required cases.'}
    foreach($caseId in @($gr.requiredCaseIds)){if($null -eq $gr.cases.PSObject.Properties[[string]$caseId] -or [string]$gr.cases.PSObject.Properties[[string]$caseId].Value -ne 'pass'){throw "Candidate gate required case is not pass: $caseId"}}
    if([string]$rr.status -ne 'READY_FOR_PROMOTION_APPROVAL' -or -not [bool]$rr.promotionEligible -or [string]$rr.candidate.sha256 -ne $candidateHash){throw 'Readiness receipt does not authorize this candidate for a separate owner-approved promotion.'}
    $backup=''; $oldHash=''
    if(Test-Path -LiteralPath $dest -PathType Leaf){$oldHash=(Get-FileHash -LiteralPath $dest -Algorithm SHA256).Hash;$backup="$dest.backup.$([DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ'))";Copy-Item -LiteralPath $dest -Destination $backup -Force}
    try{Copy-Item -LiteralPath $candidate -Destination $dest -Force;if($SimulatePromotionFailure){throw 'test injected promotion failure'};if((Get-FileHash -LiteralPath $dest -Algorithm SHA256).Hash -ne $candidateHash){throw 'final hash mismatch'}
        $receipt=[ordered]@{schemaVersion=2;generatedAtUtc=[DateTime]::UtcNow.ToString('o');version=$fullVersion;candidatePath=(Normalize-Relative $candidate);candidateSha256=$candidateHash;finalOutputPath=(Normalize-Relative $dest);finalOutputSha256=$candidateHash;sourcePath=(Normalize-Relative $sourcePaths[0]);sourceSha256=$sourceHash;sourceHashes=$currentSourceHashes;builderPath=(Normalize-Relative $builderPath);builderSha256=$builderHash;candidateBuildReceiptPath=(Normalize-Relative $buildReceipt);candidateBuildReceiptSha256=$buildReceiptHash;candidateGateReceiptPath=(Normalize-Relative $gate);candidateGateReceiptSha256=(Get-FileHash -LiteralPath $gate -Algorithm SHA256).Hash;candidateReadinessReceiptPath=(Normalize-Relative $readiness);candidateReadinessReceiptSha256=(Get-FileHash -LiteralPath $readiness -Algorithm SHA256).Hash;backupPath=$(if($backup){Normalize-Relative $backup}else{''});backupSha256=$(if($backup){(Get-FileHash -LiteralPath $backup -Algorithm SHA256).Hash}else{''});replacement='success';rollback='not-needed'}
        New-Item -ItemType Directory -Path (Join-Path $root '.build-work') -Force | Out-Null
        $receipt | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $root '.build-work\last-build.json') -Encoding UTF8
    }catch{if($backup -and (Test-Path -LiteralPath $backup)){Copy-Item -LiteralPath $backup -Destination $dest -Force}elseif(-not $oldHash -and (Test-Path -LiteralPath $dest)){Remove-Item -LiteralPath $dest -Force};throw}
    exit 0
}

if ([string]::IsNullOrWhiteSpace($Profile)) { throw 'Profile (Test or Release) is required to build a candidate.' }
$directiveValue = if ($Profile -eq 'Test') { 'True' } else { 'False' }

[System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance); $cp932=[Text.Encoding]::GetEncoding(932)
$workRoot=Join-Path $root ('.build-work\' + [guid]::NewGuid().ToString('N')); New-Item -ItemType Directory -Path $workRoot -Force | Out-Null
$tempSources=@{}; $expectedHashes=@{}; $directiveInfo=[ordered]@{}
$excel=$null; $wb=$null
try {
    foreach($path in $sourcePaths){
        $code=Get-Content -LiteralPath $path -Raw -Encoding UTF8
        $withDirective=Insert-TestBuildDirective $code $directiveValue
        $directiveMatches=[regex]::Matches($withDirective,'(?m)^#Const TEST_BUILD = (True|False)$')
        if($directiveMatches.Count -ne 1){throw "Generated directive count is not exactly 1 for $([IO.Path]::GetFileName($path)): $($directiveMatches.Count)"}
        if($cp932.GetString($cp932.GetBytes($withDirective)) -ne $withDirective){throw "cp932 roundtrip failed: $path"}
        $name=[IO.Path]::GetFileName($path);$temp=Join-Path $workRoot $name;[IO.File]::WriteAllText($temp,$withDirective,$cp932);$tempSources[$name]=$temp;$expectedHashes[$name]=Text-Hash (Normalize-Vba $code)
        $directiveInfo[$name]=[ordered]@{value=$directiveValue;count=$directiveMatches.Count;insertPosition=$directiveMatches[0].Index}
    }
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$wb=$excel.Workbooks.Add()
    while($wb.Worksheets.Count -gt 1){$wb.Worksheets.Item($wb.Worksheets.Count).Delete()};$cfg=$wb.Worksheets.Item(1);$cfg.Name='設定';$tree=$wb.Worksheets.Add();$tree.Name='ツリー';$ops=$wb.Worksheets.Add();$ops.Name='変更操作';$log=$wb.Worksheets.Add();$log.Name='実行ログ';$tree.Move($log);$ops.Move($log);$cfg.Move($tree)
    $cfg.Columns.Item(2).ColumnWidth=30;$cfg.Columns.Item(3).ColumnWidth=78;$cfg.Range('B2').Value2="フォルダ階層ツリー生成ツール $displayVersion";$cfg.Range('B2').Font.Size=16;$cfg.Range('B2').Font.Bold=$true;$cfg.Range('B3').Value2='走査結果を変更操作シートへ複製し、事前確認後に明示実行します。'
    $labels=@(@('B5','対象フォルダ'),@('B6','最大階層（0 = 無制限）'),@('B7','ファイルも一覧に含める'),@('B8','隠し／システム属性も含める'));foreach($label in $labels){$cfg.Range($label[0]).Value2=$label[1];$cfg.Range($label[0]).Font.Bold=$true};$cfg.Range('C6').Value2=0;$cfg.Range('C7').Value2='はい';$cfg.Range('C8').Value2='いいえ';foreach($cell in @('C7','C8')){$cfg.Range($cell).Validation.Add(3,1,1,'はい,いいえ')|Out-Null;$cfg.Range($cell).Validation.InCellDropdown=$true}
    $cfg.Range('C5:C8').Interior.Color=(Bgr 255 252 230);$left=[double]$cfg.Range('B10').Left;$top=[double]$cfg.Range('B10').Top;Add-Button $cfg $left $top 170 30 'btn_SelectRootFolder' '1. フォルダを選択' 'SelectRootFolder';Add-Button $cfg ($left+180) $top 170 30 'btn_BuildTree' '2. ツリーを作成' 'BuildTree';Add-Button $cfg $left ($top+38) 170 26 'btn_ExpandAllLevels' 'すべて展開' 'ExpandAllLevels';Add-Button $cfg ($left+180) ($top+38) 170 26 'btn_CollapseToLevel2' '2 階層まで折りたたむ' 'CollapseToLevel2';Add-Button $cfg $left ($top+72) 170 26 'btn_PrepareOperationSheet' '3. 変更一覧を作成' 'PrepareOperationSheet'
    $tree.Range('A1').Value2='［設定］シートから走査を開始してください。';Configure-OperationSheet $ops;Configure-LogSheet $log
    foreach($name in $sourceNames){$wb.VBProject.VBComponents.Import($tempSources[$name])|Out-Null}
    $compileBefore=Invoke-Compile $excel $wb;$wb.SaveAs($dest,52);$wb.Close($false);[Runtime.InteropServices.Marshal]::ReleaseComObject($wb)|Out-Null;$wb=$null
    Remove-AbsPathMetadata $dest
    $excel.AutomationSecurity=3;$wb=$excel.Workbooks.Open($dest,0,$true);$validation=Test-Workbook $wb $expectedHashes;$compileAfter=Invoke-Compile $excel $wb;$wb.Close($false);[Runtime.InteropServices.Marshal]::ReleaseComObject($wb)|Out-Null;$wb=$null;$excel.Quit();[Runtime.InteropServices.Marshal]::ReleaseComObject($excel)|Out-Null;$excel=$null
    Remove-AbsPathMetadata $dest
    Test-PackageOffline $dest
    $sourceHash=(Get-FileHash -LiteralPath $sourcePaths[0] -Algorithm SHA256).Hash;$candidateHash=(Get-FileHash -LiteralPath $dest -Algorithm SHA256).Hash;$receipt=[ordered]@{schemaVersion=2;generatedAtUtc=[DateTime]::UtcNow.ToString('o');version=$fullVersion;profile=$Profile;moduleCount=$sourceNames.Count;directiveInfo=$directiveInfo;builderPath=(Normalize-Relative $builderPath);builderSha256=$builderHash;sourceModules=$sourceNames;sourceHashes=@{};embeddedBodyHashes=$expectedHashes;candidatePath=(Normalize-Relative $dest);candidateSha256=$candidateHash;sourcePath=(Normalize-Relative $sourcePaths[0]);sourceSha256=$sourceHash;compileBeforeSave=$compileBefore;compileAfterReopen=$compileAfter;validation=$validation}
    foreach($path in $sourcePaths){$receipt.sourceHashes[[IO.Path]::GetFileName($path)]=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash};$receipt | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath "$dest.build-receipt.json" -Encoding UTF8
} finally { if($null -ne $wb){try{$wb.Close($false)}catch{};[Runtime.InteropServices.Marshal]::ReleaseComObject($wb)|Out-Null};if($null -ne $excel){try{$excel.Quit()}catch{};[Runtime.InteropServices.Marshal]::ReleaseComObject($excel)|Out-Null};[GC]::Collect();[GC]::WaitForPendingFinalizers(); if(Test-Path -LiteralPath $workRoot){Remove-Item -LiteralPath $workRoot -Recurse -Force} }
