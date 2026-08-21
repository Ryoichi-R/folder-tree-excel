Attribute VB_Name = "modRecycleBin"
Option Explicit

Private Const DRIVE_FIXED As Long = 3
Private Const FOF_SILENT As Long = &H4
Private Const FOF_NOCONFIRMATION As Long = &H10
Private Const FOF_ALLOWUNDO As Long = &H40
Private Const FOF_NOERRORUI As Long = &H400
Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_CANCELLED As Long = 1223
Private Const RECYCLE_MARGIN_MIN_BYTES As Double = 16777216#
Private Const RECYCLE_MARGIN_RATIO As Double = 0.05
Private Const RECYCLE_PATH_LIMIT As Long = 240

Private Type SHFILEOPSTRUCTW
    #If VBA7 Then
        hwnd As LongPtr
    #Else
        hwnd As Long
    #End If
    wFunc As Long
    #If VBA7 Then
        pFrom As LongPtr
        pTo As LongPtr
    #Else
        pFrom As Long
        pTo As Long
    #End If
    fFlags As Integer
    fAnyOperationsAborted As Long
    #If VBA7 Then
        hNameMappings As LongPtr
        lpszProgressTitle As LongPtr
    #Else
        hNameMappings As Long
        lpszProgressTitle As Long
    #End If
End Type

#If Win64 Then
Private Type SHQUERYRBINFO
    cbSize As Long
    reserved As Long
    i64Size As Currency
    i64NumItems As Currency
End Type
#Else
Private Type SHQUERYRBINFO
    cbSize As Long
    i64Size As Currency
    i64NumItems As Currency
End Type
#End If

Private Type RecycleVolumeConfig
    VolumePath As String
    VolumeGuid As String
    TotalBytes As Double
    UsedBytes As Double
    ItemCount As Double
    CapacityBytes As Double
    SafetyMarginBytes As Double
    Warning As String
End Type

#If VBA7 Then
    Private Declare PtrSafe Function SHFileOperationW Lib "shell32" (ByRef lpFileOp As SHFILEOPSTRUCTW) As Long
    Private Declare PtrSafe Function SHQueryRecycleBinW Lib "shell32" (ByVal pszRootPath As LongPtr, ByRef pSHQueryRBInfo As SHQUERYRBINFO) As Long
    Private Declare PtrSafe Function GetDriveTypeW Lib "kernel32" (ByVal lpRootPathName As LongPtr) As Long
    Private Declare PtrSafe Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As LongPtr) As Long
    Private Declare PtrSafe Function GetVolumePathNameW Lib "kernel32" (ByVal lpszFileName As LongPtr, ByVal lpszVolumePathName As LongPtr, ByVal cchBufferLength As Long) As Long
    Private Declare PtrSafe Function GetVolumeNameForVolumeMountPointW Lib "kernel32" (ByVal lpszVolumeMountPoint As LongPtr, ByVal lpszVolumeName As LongPtr, ByVal cchBufferLength As Long) As Long
    Private Declare PtrSafe Function GetDiskFreeSpaceExW Lib "kernel32" (ByVal lpDirectoryName As LongPtr, ByRef lpFreeBytesAvailableToCaller As Currency, ByRef lpTotalNumberOfBytes As Currency, ByRef lpTotalNumberOfFreeBytes As Currency) As Long
#Else
    Private Declare Function SHFileOperationW Lib "shell32" (ByRef lpFileOp As SHFILEOPSTRUCTW) As Long
    Private Declare Function SHQueryRecycleBinW Lib "shell32" (ByVal pszRootPath As Long, ByRef pSHQueryRBInfo As SHQUERYRBINFO) As Long
    Private Declare Function GetDriveTypeW Lib "kernel32" (ByVal lpRootPathName As Long) As Long
    Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long
    Private Declare Function GetVolumePathNameW Lib "kernel32" (ByVal lpszFileName As Long, ByVal lpszVolumePathName As Long, ByVal cchBufferLength As Long) As Long
    Private Declare Function GetVolumeNameForVolumeMountPointW Lib "kernel32" (ByVal lpszVolumeMountPoint As Long, ByVal lpszVolumeName As Long, ByVal cchBufferLength As Long) As Long
    Private Declare Function GetDiskFreeSpaceExW Lib "kernel32" (ByVal lpDirectoryName As Long, ByRef lpFreeBytesAvailableToCaller As Currency, ByRef lpTotalNumberOfBytes As Currency, ByRef lpTotalNumberOfFreeBytes As Currency) As Long
#End If

Private mRecycleLastWarning As String
Private mRecycleLastSnapshotInfo As String
Private mRecycleSnapshotCalls As Long
Private mRecycleExtendedPropertyCalls As Long

Public Function GetOperationAttributes(ByVal path As String) As Long
    Dim apiPath As String
    If Len(path) = 0 Then GetOperationAttributes = -1: Exit Function
    On Error GoTo Fail
    apiPath = ToExtendedOperationPath(path)
    GetOperationAttributes = GetFileAttributesW(StrPtr(apiPath))
    Exit Function
Fail:
    GetOperationAttributes = -1
End Function

Public Function ToExtendedOperationPath(ByVal path As String) As String
    If Left$(path, 4) = "\\?\" Then
        ToExtendedOperationPath = path
    ElseIf Left$(path, 2) = "\\" Then
        ToExtendedOperationPath = "\\?\UNC\" & Mid$(path, 3)
    ElseIf Len(path) >= 3 And Mid$(path, 2, 2) = ":\" Then
        ToExtendedOperationPath = "\\?\" & path
    Else
        ToExtendedOperationPath = path
    End If
End Function

Public Function VerifyOperationSnapshot(ByVal path As String, ByVal expectedAttributes As Long, _
                                        ByVal expectedSize As Double, ByVal expectedModified As Double, _
                                        ByRef message As String) As Boolean
    Dim fso As Object, entry As Object, actualAttributes As Long
    Dim actualSize As Double, actualModified As Double
    message = ""
    actualAttributes = GetOperationAttributes(path)
    If actualAttributes = -1 Then message = "sourceが一覧作成後に消失しました。": Exit Function
    If actualAttributes <> expectedAttributes Then message = "sourceの属性が一覧作成後に変化しました。": Exit Function
    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (actualAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0 Then
        Set entry = fso.GetFolder(path)
    Else
        Set entry = fso.GetFile(path)
        actualSize = CDbl(entry.Size)
    End If
    actualModified = CDbl(entry.DateLastModified)
    If (actualAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
        If Abs(actualSize - expectedSize) > 0.5 Then message = "sourceのsizeが一覧作成後に変化しました。": Exit Function
    End If
    If expectedModified > 0 And Abs(actualModified - expectedModified) > (2# / 86400#) Then
        message = "sourceの更新日時が一覧作成後に変化しました。": Exit Function
    End If
    VerifyOperationSnapshot = True
    Exit Function
Fail:
    message = "sourceのidentityを再照合できません。"
End Function

Public Function CanRecycleOperation(ByVal source As String, ByVal itemSize As Double, ByRef message As String) As Boolean
    Dim config As RecycleVolumeConfig
    message = ""
    If Len(source) > RECYCLE_PATH_LIMIT Then message = "ゴミ箱APIのpath長上限を超えています。": Exit Function
    If Not ResolveRecycleVolumeConfig(source, config, message) Then Exit Function
    If Len(config.Warning) > 0 Then mRecycleLastWarning = config.Warning
    If Not IsRecycleItemSizeAllowed(config, itemSize, message) Then Exit Function
    UpdateRecycleBatchCapacityWarning config, itemSize
    CanRecycleOperation = True
End Function

#If TEST_BUILD Then
Public Function CanRecycleOperationTest(ByVal source As String, ByVal itemSize As Double) As String
    Dim message As String, allowed As Boolean
    allowed = CanRecycleOperation(source, itemSize, message)
    CanRecycleOperationTest = IIf(allowed, "allowed|", "blocked|") & message
End Function
#End If

Public Function GetRecycleBatchWarning() As String
    GetRecycleBatchWarning = mRecycleLastWarning
End Function

Public Sub ResetRecycleBatchWarning()
    mRecycleLastWarning = ""
    mRecycleLastSnapshotInfo = ""
End Sub

Public Function GetRecycleSnapshotInfo() As String
    GetRecycleSnapshotInfo = mRecycleLastSnapshotInfo
End Function

#If TEST_BUILD Then
Public Sub ResetRecyclePerformanceCounters()
    mRecycleSnapshotCalls = 0
    mRecycleExtendedPropertyCalls = 0
End Sub

Public Function GetRecyclePerformanceCounters() As String
    GetRecyclePerformanceCounters = "snapshotCalls=" & CStr(mRecycleSnapshotCalls) & _
                                    "|extendedPropertyCalls=" & CStr(mRecycleExtendedPropertyCalls)
End Function

Public Function RecycleSnapshotPerformanceTest() As String
    Dim items As Object, canonical As String, message As String
    ResetRecyclePerformanceCounters
    If Not TakeRecycleSnapshot(items, canonical, message) Then
        RecycleSnapshotPerformanceTest = "fail|" & Replace$(message, "|", "/")
        Exit Function
    End If
    RecycleSnapshotPerformanceTest = "pass|items=" & CStr(items.Count) & _
                                     "|canonicalLength=" & CStr(Len(canonical)) & _
                                     "|fingerprint=" & SnapshotFingerprint(canonical) & _
                                     "|" & GetRecyclePerformanceCounters()
End Function

Public Function RecycleSyntheticSnapshotPerformanceTest(ByVal itemCount As Long) As String
    Dim beforeItems As Object, afterItems As Object, expected As Object
    Dim beforeCanonical As String, afterCanonical As String, message As String
    Dim i As Long, key As String, value As String
    If itemCount < 1 Or itemCount > 50000 Then
        RecycleSyntheticSnapshotPerformanceTest = "fail|itemCount must be 1..50000"
        Exit Function
    End If
    Set beforeItems = CreateObject("Scripting.Dictionary"): beforeItems.CompareMode = 1
    Set afterItems = CreateObject("Scripting.Dictionary"): afterItems.CompareMode = 1
    Set expected = CreateObject("Scripting.Dictionary"): expected.CompareMode = 1
    For i = 1 To itemCount
        key = "C:\synthetic-recycle\item-" & Right$("000000" & CStr(i), 6) & ".dat"
        value = CStr(i) & "|45000|False|item-" & CStr(i) & ".dat"
        beforeItems.Add key, value
        afterItems.Add key, value
    Next i
    key = "C:\synthetic-recycle\expected.dat"
    value = "1|45000|False|expected.dat"
    expected.Add key, value
    afterItems.Add key, value
    beforeCanonical = SnapshotCanonical(beforeItems)
    afterCanonical = SnapshotCanonical(afterItems)
    If Not ValidateRecycleSnapshotDelta(beforeItems, afterItems, expected, message) Then
        RecycleSyntheticSnapshotPerformanceTest = "fail|" & Replace$(message, "|", "/")
        Exit Function
    End If
    RecycleSyntheticSnapshotPerformanceTest = "pass|items=" & CStr(itemCount) & _
                                              "|beforeLength=" & CStr(Len(beforeCanonical)) & _
                                              "|afterLength=" & CStr(Len(afterCanonical)) & _
                                              "|beforeFingerprint=" & SnapshotFingerprint(beforeCanonical) & _
                                              "|afterFingerprint=" & SnapshotFingerprint(afterCanonical)
End Function

Public Function RecycleCapacityBoundaryTest(ByVal source As String) As String
    Dim config As RecycleVolumeConfig, message As String, threshold As Double, aggregateSize As Double
    Dim belowPass As Boolean, equalPass As Boolean, abovePass As Boolean, aggregateWarning As Boolean
    If Not ResolveRecycleVolumeConfig(source, config, message) Then
        RecycleCapacityBoundaryTest = "fail|" & Replace$(message, "|", "/")
        Exit Function
    End If
    threshold = config.CapacityBytes - config.SafetyMarginBytes
    belowPass = IsRecycleItemSizeAllowed(config, threshold - 1#, message)
    equalPass = IsRecycleItemSizeAllowed(config, threshold, message)
    abovePass = IsRecycleItemSizeAllowed(config, threshold + 1#, message)
    mRecycleLastWarning = ""
    aggregateSize = config.CapacityBytes - config.UsedBytes + 1#
    If aggregateSize < 1# Then aggregateSize = 1#
    UpdateRecycleBatchCapacityWarning config, aggregateSize
    aggregateWarning = (Len(mRecycleLastWarning) > 0)
    If Not belowPass Or Not equalPass Or abovePass Or Not aggregateWarning Then
        RecycleCapacityBoundaryTest = "fail|capacity boundary assertion failed"
        Exit Function
    End If
    RecycleCapacityBoundaryTest = "pass|totalBytes=" & CStr(config.TotalBytes) & _
                                  "|usedBytes=" & CStr(config.UsedBytes) & _
                                  "|capacityBytes=" & CStr(config.CapacityBytes) & _
                                  "|safetyMarginBytes=" & CStr(config.SafetyMarginBytes) & _
                                  "|thresholdBytes=" & CStr(threshold) & _
                                  "|below=True|equal=True|above=False|aggregateWarning=True"
End Function

Private Function GetRecyclePerformanceRootPrefix() As String
    ' test candidateは <repoRoot>\.build-work（ビルド直後）または
    ' <repoRoot>\.test-work\<nonce>（test-xlsm.ps1が一時copyを開いた状態）のいずれかで開かれる。
    ' どちらの既知マーカーもrepo root自身の名前には依存しないため、
    ' 任意folderへ緩めることなくrepo rootを動的に解決できる。
    Dim workbookDir As String, repoRoot As String, marker As Long
    workbookDir = ThisWorkbook.Path & "\"
    marker = InStr(1, workbookDir, "\.build-work\", vbTextCompare)
    If marker = 0 Then marker = InStr(1, workbookDir, "\.test-work\", vbTextCompare)
    If marker = 0 Then
        Err.Raise 5, "modRecycleBin.GetRecyclePerformanceRootPrefix", "test candidateの配置場所からrepo rootを解決できません。"
    End If
    repoRoot = Left$(workbookDir, marker - 1)
    GetRecyclePerformanceRootPrefix = repoRoot & "\.test-work\recycle-performance-"
End Function

Public Function RecycleDirectoryFilesPerformanceTest(ByVal folderPath As String) As String
    Dim fso As Object, fixtureFolder As Object, file As Object, itemCount As Long, i As Long
    Dim sources() As String, sizes() As Double, modified() As Double, isFolder() As Boolean, states() As String
    Dim errorNumber As Long, errorMessage As String, safetyState As String, apiStarted As Boolean, successCount As Long
    Dim normalizedFolder As String, performanceRoot As String, separatorAt As Long, allowedPrefix As String
    On Error GoTo Fail
    normalizedFolder = NormalizeRecyclePath(folderPath)
    allowedPrefix = GetRecyclePerformanceRootPrefix()
    If StrComp(Left$(normalizedFolder, Len(allowedPrefix)), allowedPrefix, vbTextCompare) <> 0 Then
        RecycleDirectoryFilesPerformanceTest = "fail|performance fixture path rejected"
        Exit Function
    End If
    separatorAt = InStr(Len(allowedPrefix) + 1, normalizedFolder, "\")
    If separatorAt = 0 Then
        RecycleDirectoryFilesPerformanceTest = "fail|performance fixture must be below a nonce root"
        Exit Function
    End If
    performanceRoot = Left$(normalizedFolder, separatorAt - 1)
    If GetOperationAttributes(normalizedFolder) = -1 Or _
       (GetOperationAttributes(normalizedFolder) And FILE_ATTRIBUTE_REPARSE_POINT) <> 0 Or _
       HasReparseAncestor(normalizedFolder, performanceRoot) Then
        RecycleDirectoryFilesPerformanceTest = "fail|performance fixture path is missing or reparse-backed"
        Exit Function
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fixtureFolder = fso.GetFolder(normalizedFolder)
    If fixtureFolder.SubFolders.Count <> 0 Then
        RecycleDirectoryFilesPerformanceTest = "fail|performance fixture must not contain subfolders"
        Exit Function
    End If
    itemCount = fixtureFolder.Files.Count
    If itemCount < 1 Or itemCount > 1000 Then
        RecycleDirectoryFilesPerformanceTest = "fail|file count must be 1..1000"
        Exit Function
    End If
    ReDim sources(1 To itemCount): ReDim sizes(1 To itemCount): ReDim modified(1 To itemCount): ReDim isFolder(1 To itemCount)
    For Each file In fixtureFolder.Files
        i = i + 1
        sources(i) = CStr(file.Path)
        sizes(i) = CDbl(file.Size)
        modified(i) = CDbl(file.DateLastModified)
    Next file
    ResetRecyclePerformanceCounters
    If Not RecycleItemsPhase(sources, sizes, modified, isFolder, itemCount, states, errorNumber, errorMessage, safetyState, apiStarted) Then
        RecycleDirectoryFilesPerformanceTest = "fail|count=" & CStr(itemCount) & _
                                               "|errorNumber=" & CStr(errorNumber) & _
                                               "|safety=" & safetyState & _
                                               "|message=" & Replace$(errorMessage, "|", "/") & _
                                               "|" & GetRecyclePerformanceCounters()
        Exit Function
    End If
    For i = 1 To itemCount
        If states(i) = "成功" Then successCount = successCount + 1
    Next i
    RecycleDirectoryFilesPerformanceTest = "pass|count=" & CStr(itemCount) & _
                                           "|success=" & CStr(successCount) & _
                                           "|safety=" & safetyState & _
                                           "|" & GetRecyclePerformanceCounters() & _
                                           "|snapshot=" & GetRecycleSnapshotInfo()
    Exit Function
Fail:
    RecycleDirectoryFilesPerformanceTest = "fail|" & Replace$(Err.Description, "|", "/")
End Function
#End If

Public Function RecyclePath(ByVal source As String, ByRef errorNumber As Long, ByRef errorMessage As String) As Boolean
    Dim sources(1 To 1) As String, sizes(1 To 1) As Double, modified(1 To 1) As Double
    Dim folders(1 To 1) As Boolean, states() As String, safety As String, apiStarted As Boolean
    Dim fso As Object, file As Object
    sources(1) = source
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (GetOperationAttributes(source) And FILE_ATTRIBUTE_DIRECTORY) <> 0 Then
        folders(1) = True
        modified(1) = CDbl(fso.GetFolder(source).DateLastModified)
    Else
        Set file = fso.GetFile(source)
        sizes(1) = CDbl(file.Size)
        modified(1) = CDbl(file.DateLastModified)
    End If
    RecyclePath = RecycleItemsPhase(sources, sizes, modified, folders, 1, states, errorNumber, errorMessage, safety, apiStarted)
End Function

Public Function HasReparseAncestor(ByVal path As String, ByVal rootPath As String) As Boolean
    Dim current As String, attrs As Long
    current = NormalizeRecyclePath(path)
    rootPath = NormalizeRecyclePath(rootPath)
    Do
        attrs = GetOperationAttributes(current)
        If attrs <> -1 Then
            If (attrs And FILE_ATTRIBUTE_REPARSE_POINT) <> 0 Then HasReparseAncestor = True: Exit Function
        End If
        If StrComp(current, rootPath, vbTextCompare) = 0 Then Exit Do
        current = ParentRecyclePath(current)
        If Len(current) = 0 Then Exit Do
    Loop
End Function

Public Function IsOperationFolderEmpty(ByVal path As String) As Boolean
    Dim fso As Object, folder As Object
    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(path)
    IsOperationFolderEmpty = (folder.Files.Count = 0 And folder.SubFolders.Count = 0)
    Exit Function
Fail:
    IsOperationFolderEmpty = False
End Function

Public Function RecycleItemsPhase(ByRef sources() As String, ByRef itemSizes() As Double, _
                                  ByRef itemModified() As Double, ByRef itemIsFolder() As Boolean, _
                                  ByVal itemCount As Long, ByRef itemStates() As String, _
                                  ByRef errorNumber As Long, ByRef errorMessage As String, _
                                  ByRef safetyState As String, ByRef apiStarted As Boolean) As Boolean
    Dim config As RecycleVolumeConfig, i As Long, totalSize As Double
    Dim beforeItems As Object, afterItems As Object, beforeCanonical As String, afterCanonical As String
    Dim fromBuffer As String, op As SHFILEOPSTRUCTW, result As Long
    Dim expected As Object, key As Variant
    Dim sourceExists As Boolean

    errorNumber = 0: errorMessage = "": safetyState = "通常": apiStarted = False
    mRecycleLastSnapshotInfo = ""
    If itemCount <= 0 Then errorMessage = "削除対象がありません。": Exit Function
    ReDim itemStates(1 To itemCount)
    For i = 1 To itemCount
        itemStates(i) = "未実行"
        totalSize = totalSize + itemSizes(i)
    Next i
    If Not ResolveRecycleVolumeConfig(sources(1), config, errorMessage) Then Exit Function
    UpdateRecycleBatchCapacityWarning config, totalSize
    For i = 1 To itemCount
        If Not SameVolumeGuid(sources(i), config.VolumeGuid, errorMessage) Then Exit Function
        If Len(sources(i)) > RECYCLE_PATH_LIMIT Then errorMessage = "ゴミ箱APIのpath長上限を超えています。": Exit Function
        If itemSizes(i) < 0 Then errorMessage = "対象sizeを解決できません。": Exit Function
        If Not itemIsFolder(i) And Not IsRecycleItemSizeAllowed(config, itemSizes(i), errorMessage) Then Exit Function
    Next i

    If Not TakeRecycleSnapshot(beforeItems, beforeCanonical, errorMessage) Then Exit Function
    mRecycleLastSnapshotInfo = "before=" & SnapshotFingerprint(beforeCanonical)
    Set expected = CreateObject("Scripting.Dictionary")
    expected.CompareMode = 1
    For i = 1 To itemCount
        key = NormalizeRecyclePath(sources(i))
        If beforeItems.Exists(key) Then errorMessage = "削除対象と同じ元pathの既存Recycle Bin項目があり、照合が曖昧です。": Exit Function
        expected.Add key, CStr(itemSizes(i)) & "|" & CStr(itemModified(i)) & "|" & CStr(itemIsFolder(i))
        fromBuffer = fromBuffer & sources(i) & vbNullChar
    Next i
    fromBuffer = fromBuffer & vbNullChar

    op.wFunc = 3
    op.pFrom = StrPtr(fromBuffer)
    op.fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION Or FOF_NOERRORUI Or FOF_SILENT
    On Error GoTo ApiFail
    apiStarted = True
    result = SHFileOperationW(op)
    If result <> ERROR_SUCCESS Then errorNumber = result: errorMessage = "SHFileOperationW failed.": GoTo VerifyPhase
    If op.fAnyOperationsAborted <> 0 Then errorNumber = ERROR_CANCELLED: errorMessage = "ゴミ箱移動が中断されました。": GoTo VerifyPhase

VerifyPhase:
    If Not TakeRecycleSnapshot(afterItems, afterCanonical, errorMessage) Then
        mRecycleLastSnapshotInfo = mRecycleLastSnapshotInfo & ";after=unavailable"
        safetyState = "recycle-verification-contaminated"
        For i = 1 To itemCount: itemStates(i) = safetyState: Next i
        Exit Function
    End If
    mRecycleLastSnapshotInfo = mRecycleLastSnapshotInfo & ";after=" & SnapshotFingerprint(afterCanonical)
    If Not ValidateRecycleSnapshotDelta(beforeItems, afterItems, expected, errorMessage) Then
        safetyState = "recycle-verification-contaminated"
        errorMessage = errorMessage & "; source=" & CStr(sources(1))
        For i = 1 To itemCount: itemStates(i) = safetyState: Next i
        Exit Function
    End If
    For i = 1 To itemCount
        key = NormalizeRecyclePath(sources(i))
        sourceExists = (GetOperationAttributes(sources(i)) <> -1)
        If Not afterItems.Exists(key) Then
            If sourceExists Then
                itemStates(i) = "失敗"
            Else
                itemStates(i) = "possible-permanent-delete"
                safetyState = "possible-permanent-delete"
            End If
        ElseIf Not VerifyRecycleItemValue(afterItems(key), itemSizes(i), itemModified(i), itemIsFolder(i)) Then
            itemStates(i) = "recycle-verification-contaminated"
            safetyState = "recycle-verification-contaminated"
            errorMessage = "Recycle Bin item value mismatch. key=" & CStr(key) & "; actual=" & CStr(afterItems(key)) & "; expected=" & CStr(expected(key))
        ElseIf sourceExists Then
            itemStates(i) = "失敗"
        Else
            itemStates(i) = "成功"
        End If
    Next i
    For i = 1 To itemCount
        If itemStates(i) <> "成功" Then
            If safetyState = "通常" Then safetyState = itemStates(i)
            If Len(errorMessage) = 0 Then errorMessage = "Recycle Bin phaseの事後照合に失敗しました。前後snapshot=" & Len(beforeCanonical) & "/" & Len(afterCanonical)
            Exit Function
        End If
    Next i
    RecycleItemsPhase = True
    Exit Function
ApiFail:
    errorNumber = Err.Number: errorMessage = Err.Description
    If apiStarted Then safetyState = "possible-permanent-delete"
    For i = 1 To itemCount: itemStates(i) = "未実行": Next i
End Function

Private Function ResolveRecycleVolumeConfig(ByVal source As String, ByRef config As RecycleVolumeConfig, ByRef message As String) As Boolean
    Dim volumePath As String, volumeGuid As String, totalBytes As Currency, freeBytes As Currency, freeTotal As Currency
    Dim query As SHQUERYRBINFO, driveType As Long, status As Long
    Dim shell As Object, volumeKey As String, policyKey As String, maxCapacity As Double, policySize As Double
    Dim hasVolumeCapacity As Boolean, hasPolicySize As Boolean, nuke As Double, noRecycle As Double
    message = "": config.VolumePath = "": config.VolumeGuid = ""
    If Len(source) = 0 Or Left$(source, 2) = "\\" Or InStr(source, "\?") = 1 Then message = "UNC/networkまたは拡張pathはゴミ箱対象外です。": Exit Function
    If Len(source) > RECYCLE_PATH_LIMIT Then message = "ゴミ箱APIのpath長上限を超えています。": Exit Function
    driveType = GetDriveTypeW(StrPtr(Left$(source, 3)))
    If driveType <> DRIVE_FIXED Then message = "固定local volume以外の削除は拒否します。": Exit Function
    volumePath = String$(260, vbNullChar)
    If GetVolumePathNameW(StrPtr(source), StrPtr(volumePath), Len(volumePath)) = 0 Then message = "volume pathを解決できません。": Exit Function
    volumePath = Left$(volumePath, InStr(volumePath, vbNullChar) - 1)
    volumeGuid = String$(260, vbNullChar)
    If GetVolumeNameForVolumeMountPointW(StrPtr(volumePath), StrPtr(volumeGuid), Len(volumeGuid)) = 0 Then message = "volume GUIDを解決できません。": Exit Function
    volumeGuid = Left$(volumeGuid, InStr(volumeGuid, vbNullChar) - 1)
    If GetDiskFreeSpaceExW(StrPtr(volumePath), freeBytes, totalBytes, freeTotal) = 0 Then message = "volume容量を解決できません。": Exit Function
    query.cbSize = LenB(query)
    If SHQueryRecycleBinW(StrPtr(volumePath), query) <> ERROR_SUCCESS Then message = "Recycle Bin容量snapshotを取得できません。": Exit Function
    If Not CurrencyToDouble(totalBytes, config.TotalBytes) Then message = "volume総容量を解決できません。": Exit Function
    If Not CurrencyToDouble(query.i64Size, config.UsedBytes) Then message = "Recycle Bin使用量を解決できません。": Exit Function
    config.ItemCount = CurrencyToDoubleValue(query.i64NumItems)
    config.VolumePath = volumePath
    config.VolumeGuid = volumeGuid

    Set shell = CreateObject("WScript.Shell")
    policyKey = "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    volumeKey = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\BitBucket\Volume\" & VolumeGuidKey(volumeGuid)
    status = ReadRegistryNumber(shell, policyKey, "NoRecycleFiles", noRecycle)
    If status < 0 Then message = "NoRecycleFiles policyの型を解決できません。": Exit Function
    If status > 0 And noRecycle <> 0 Then message = "NoRecycleFiles policyによりゴミ箱移動が無効です。": Exit Function
    status = ReadRegistryNumber(shell, volumeKey, "NukeOnDelete", nuke)
    If status < 0 Then message = "NukeOnDelete設定の型を解決できません。": Exit Function
    If status > 0 And nuke <> 0 Then message = "NukeOnDelete設定により完全削除の可能性があるため拒否します。": Exit Function
    status = ReadRegistryNumber(shell, policyKey, "RecycleBinSize", policySize)
    If status < 0 Then message = "RecycleBinSize policyの型を解決できません。": Exit Function
    If status > 0 Then hasPolicySize = True
    status = ReadRegistryNumber(shell, volumeKey, "MaxCapacity", maxCapacity)
    If status < 0 Then message = "MaxCapacity設定の型を解決できません。": Exit Function
    If status > 0 Then hasVolumeCapacity = True
    If hasPolicySize Then
        If policySize >= 1 And policySize <= 100 Then
            config.CapacityBytes = config.TotalBytes * policySize / 100#
        ElseIf policySize > 100 Then
            config.CapacityBytes = policySize * 1048576#
        Else
            message = "RecycleBinSize policyの値を安全に解釈できません。": Exit Function
        End If
    ElseIf hasVolumeCapacity And maxCapacity > 0 Then
        config.CapacityBytes = maxCapacity * 1048576#
    Else
        message = "volume別MaxCapacityまたは優先policyが解決できません。": Exit Function
    End If
    If config.CapacityBytes > config.TotalBytes Then config.CapacityBytes = config.TotalBytes
    config.SafetyMarginBytes = RECYCLE_MARGIN_MIN_BYTES
    If config.TotalBytes * RECYCLE_MARGIN_RATIO > config.SafetyMarginBytes Then config.SafetyMarginBytes = -Int(-(config.TotalBytes * RECYCLE_MARGIN_RATIO))
    If config.CapacityBytes <= config.SafetyMarginBytes Then message = "ゴミ箱容量が安全margin以下です。": Exit Function
    If config.UsedBytes + config.SafetyMarginBytes > config.CapacityBytes Then config.Warning = "既存使用量が容量上限に近く、自動整理の可能性があります。"
    ResolveRecycleVolumeConfig = True
End Function

Private Function IsRecycleItemSizeAllowed(ByRef config As RecycleVolumeConfig, ByVal itemSize As Double, ByRef message As String) As Boolean
    message = ""
    If itemSize < 0 Then message = "対象sizeを解決できません。": Exit Function
    If itemSize > config.CapacityBytes - config.SafetyMarginBytes Then
        message = "対象sizeがゴミ箱設定上限から安全marginを引いた値を超えています。": Exit Function
    End If
    IsRecycleItemSizeAllowed = True
End Function

Private Sub UpdateRecycleBatchCapacityWarning(ByRef config As RecycleVolumeConfig, ByVal batchSize As Double)
    If config.UsedBytes + batchSize > config.CapacityBytes Then
        mRecycleLastWarning = "現在使用量と削除batch合計がゴミ箱容量を超える可能性があります。既存項目の自動整理を事後照合で検出します。"
    End If
End Sub

Private Function TakeRecycleSnapshot(ByRef items As Object, ByRef canonical As String, ByRef message As String) As Boolean
    Dim shell As Object, registryShell As Object, bin As Object, item As Object, key As String, size As Double, modified As Double, deletedFrom As String
    Dim isFolder As Boolean, timezoneBias As Double, timezoneStatus As Long
    message = "": canonical = ""
    mRecycleSnapshotCalls = mRecycleSnapshotCalls + 1
    On Error GoTo Fail
    Set items = CreateObject("Scripting.Dictionary")
    items.CompareMode = 1
    Set shell = CreateObject("Shell.Application")
    Set bin = shell.Namespace(10)
    If bin Is Nothing Then message = "Recycle Bin namespaceを取得できません。": Exit Function
    Set registryShell = CreateObject("WScript.Shell")
    timezoneStatus = ReadRegistryNumber(registryShell, "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation", "ActiveTimeBias", timezoneBias)
    If timezoneStatus <= 0 Then message = "Recycle BinのDateModifiedをlocal timeへ変換するtimezone biasを取得できません。": Exit Function
    For Each item In bin.Items
        deletedFrom = ExtendedText(item, "System.Recycle.DeletedFrom")
        If Len(deletedFrom) = 0 Then message = "Recycle Bin項目の元path propertyを取得できません。": Exit Function
        If Not ExtendedNumber(item, "System.Size", size) Then message = "Recycle Bin項目のsize propertyを取得できません。": Exit Function
        If Not ExtendedNumber(item, "System.DateModified", modified) Then message = "Recycle Bin項目のmtime propertyを取得できません。": Exit Function
        ' System.DateModifiedはShell propertyではUTCシリアルで返るため、
        ' source側のFSO/FileTimeToLocalFileTimeと同じlocal timeへ揃える。
        modified = modified - timezoneBias / 1440#
        On Error Resume Next
        isFolder = CBool(item.IsFolder)
        If Err.Number <> 0 Then Err.Clear: On Error GoTo Fail: message = "Recycle Bin項目の種別を取得できません。": Exit Function
        On Error GoTo Fail
        ' System.Recycle.DeletedFrom はWindows環境では元の親フォルダを返す。
        ' 同じフォルダから複数項目を削除すると親pathだけでは衝突するため、
        ' item.Nameを結合して元の項目full pathを構成する。既にfull pathを
        ' 返すShell実装では重複結合しない。
        key = NormalizeRecyclePath(deletedFrom)
        If StrComp(Mid$(key, InStrRev(key, "\") + 1), CStr(item.Name), vbTextCompare) <> 0 Then
            key = NormalizeRecyclePath(key & "\" & CStr(item.Name))
        End If
        If Len(key) = 0 Or items.Exists(key) Then message = "Recycle Bin項目の元pathが一意に解決できません。": Exit Function
        items.Add key, CStr(size) & "|" & CStr(modified) & "|" & CStr(isFolder) & "|" & CStr(item.Name)
    Next item
    canonical = SnapshotCanonical(items)
    TakeRecycleSnapshot = True
    Exit Function
Fail:
    message = "Recycle Bin snapshotを取得できません。"
End Function

Private Function VerifyRecycleItemValue(ByVal value As String, ByVal expectedSize As Double, ByVal expectedModified As Double, ByVal isFolder As Boolean) As Boolean
    Dim parts() As String, actualSize As Double, actualModified As Double
    parts = Split(value, "|")
    If UBound(parts) < 2 Then Exit Function
    actualSize = CDbl(parts(0)): actualModified = CDbl(parts(1))
    If StrComp(parts(2), CStr(isFolder), vbTextCompare) <> 0 Then Exit Function
    If Not isFolder And Abs(actualSize - expectedSize) > 0.5 Then Exit Function
    If expectedModified > 0 And Abs(actualModified - expectedModified) > (2# / 86400#) Then Exit Function
    VerifyRecycleItemValue = True
End Function

Private Function ExtendedText(ByVal item As Object, ByVal propertyName As String) As String
    On Error GoTo Fail
    mRecycleExtendedPropertyCalls = mRecycleExtendedPropertyCalls + 1
    ExtendedText = CStr(item.ExtendedProperty(propertyName))
    Exit Function
Fail:
    ExtendedText = ""
End Function

Private Function ExtendedNumber(ByVal item As Object, ByVal propertyName As String, ByRef result As Double) As Boolean
    Dim value As Variant
    On Error GoTo Fail
    mRecycleExtendedPropertyCalls = mRecycleExtendedPropertyCalls + 1
    value = item.ExtendedProperty(propertyName)
    If IsDate(value) Then result = CDbl(CDate(value)) Else result = CDbl(value)
    ExtendedNumber = True
    Exit Function
Fail:
    ExtendedNumber = False
End Function

Private Function ReadRegistryNumber(ByVal shell As Object, ByVal key As String, ByVal name As String, ByRef result As Double) As Long
    Dim value As Variant, t As Long
    On Error GoTo Missing
    value = shell.RegRead(key & "\" & name)
    t = VarType(value)
    If t = vbByte Or t = vbInteger Or t = vbLong Or t = vbSingle Or t = vbDouble Or t = vbCurrency Or t = vbDecimal Then
        result = CDbl(value): ReadRegistryNumber = 1: Exit Function
    End If
    ReadRegistryNumber = -1
    Exit Function
Missing:
    ReadRegistryNumber = 0
End Function

Private Function CurrencyToDouble(ByVal value As Currency, ByRef result As Double) As Boolean
    On Error GoTo Fail
    result = CDbl(value) * 10000#
    CurrencyToDouble = True
    Exit Function
Fail:
    CurrencyToDouble = False
End Function

Private Function CurrencyToDoubleValue(ByVal value As Currency) As Double
    CurrencyToDoubleValue = CDbl(value) * 10000#
End Function

Private Function VolumeGuidKey(ByVal value As String) As String
    Dim p As Long, q As Long
    p = InStr(1, value, "{", vbTextCompare): q = InStrRev(value, "}")
    If p > 0 And q >= p Then VolumeGuidKey = Mid$(value, p, q - p + 1) Else VolumeGuidKey = value
End Function

Private Function SameVolumeGuid(ByVal source As String, ByVal expectedGuid As String, ByRef message As String) As Boolean
    Dim config As RecycleVolumeConfig
    If Not ResolveRecycleVolumeConfig(source, config, message) Then Exit Function
    If StrComp(config.VolumeGuid, expectedGuid, vbTextCompare) <> 0 Then message = "削除batchが複数volumeにまたがるため停止します。": Exit Function
    SameVolumeGuid = True
End Function

Private Function SnapshotCanonical(ByVal items As Object) As String
    Dim keys As Variant, lines() As String, i As Long
    If items.Count = 0 Then SnapshotCanonical = "0:": Exit Function
    keys = items.Keys
    SortSnapshotKeys keys, LBound(keys), UBound(keys)
    ReDim lines(LBound(keys) To UBound(keys))
    For i = LBound(keys) To UBound(keys): lines(i) = CStr(keys(i)) & "=" & CStr(items(keys(i))): Next i
    SnapshotCanonical = CStr(items.Count) & ":" & Join(lines, vbCrLf) & vbCrLf
End Function

Private Sub SortSnapshotKeys(ByRef keys As Variant, ByVal lower As Long, ByVal upper As Long)
    Dim i As Long, j As Long, pivot As String, temp As Variant
    i = lower: j = upper: pivot = CStr(keys((lower + upper) \ 2))
    Do While i <= j
        Do While StrComp(CStr(keys(i)), pivot, vbTextCompare) < 0: i = i + 1: Loop
        Do While StrComp(CStr(keys(j)), pivot, vbTextCompare) > 0: j = j - 1: Loop
        If i <= j Then
            temp = keys(i): keys(i) = keys(j): keys(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If lower < j Then SortSnapshotKeys keys, lower, j
    If i < upper Then SortSnapshotKeys keys, i, upper
End Sub

Private Function ValidateRecycleSnapshotDelta(ByVal beforeItems As Object, ByVal afterItems As Object, _
                                              ByVal expected As Object, ByRef message As String) As Boolean
    Dim key As Variant
    message = ""
    For Each key In beforeItems.Keys
        If Not afterItems.Exists(key) Then
            message = "削除前後のRecycle Bin差分で既存項目が消失しました。消失key=" & CStr(key)
            Exit Function
        End If
    Next key
    For Each key In afterItems.Keys
        If Not beforeItems.Exists(key) And Not expected.Exists(key) Then
            message = "削除前後のRecycle Bin差分に外部項目が混入しました。追加key=" & CStr(key)
            Exit Function
        End If
    Next key
    ValidateRecycleSnapshotDelta = True
End Function

Private Function SnapshotFingerprint(ByVal value As String) As String
    ' 31-bitの決定的fingerprint。暗号学的ハッシュではなく、前後snapshotの
    ' receipt相関用であり、照合そのものはdictionaryのpath/type/size/mtimeで行う。
    Dim i As Long, code As Long, hashValue As Double
    hashValue = 5381#
    For i = 1 To Len(value)
        code = AscW(Mid$(value, i, 1))
        If code < 0 Then code = code + 65536
        hashValue = hashValue * 33# + code
        ' VBAのModはDouble式でも内部でLongへ寄せるため、長いsnapshotでは
        ' 中間値がLong上限を超えてオーバーフローする。Doubleのまま剰余を求める。
        hashValue = hashValue - 2147483647# * Fix(hashValue / 2147483647#)
    Next i
    SnapshotFingerprint = "djb2-31:" & Right$("00000000" & Hex$(CLng(hashValue)), 8)
End Function

Private Function NormalizeRecyclePath(ByVal value As String) As String
    Dim fso As Object
    value = Trim$(Replace$(value, "/", "\"))
    If (Len(value) >= 3 And Mid$(value, 2, 2) = ":\" Or Left$(value, 2) = "\\") And _
       InStr(1, value, "\.\", vbBinaryCompare) = 0 And InStr(1, value, "\..\", vbBinaryCompare) = 0 Then
        NormalizeRecyclePath = value
    Else
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next
        NormalizeRecyclePath = fso.GetAbsolutePathName(value)
        On Error GoTo 0
    End If
    If Right$(NormalizeRecyclePath, 1) = "\" And Len(NormalizeRecyclePath) > 3 Then NormalizeRecyclePath = Left$(NormalizeRecyclePath, Len(NormalizeRecyclePath) - 1)
End Function

Private Function ParentRecyclePath(ByVal value As String) As String
    Dim p As Long
    p = InStrRev(value, "\")
    If p > 0 Then ParentRecyclePath = Left$(value, p - 1)
End Function
