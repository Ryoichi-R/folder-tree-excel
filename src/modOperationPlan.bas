Attribute VB_Name = "modOperationPlan"
Option Explicit

' v1.1 の操作案、検査、fingerprintを担当する。状態はVBAメモリだけに保持する。
Public Const OP_NONE As String = "変更なし"
Public Const OP_RENAME_MOVE As String = "名前変更/移動"
Public Const OP_RECYCLE_FILE As String = "ファイルをゴミ箱へ"
Public Const OP_MKDIR As String = "フォルダ作成"
Public Const OP_RECYCLE_FOLDER As String = "空フォルダをゴミ箱へ"

Public Type OperationPlanItem
    WorksheetRow As Long
    Sequence As Long
    Kind As String
    Source As String
    Destination As String
    OriginalName As String
    NewName As String
    RelativeDestination As String
    SourceType As String
    SourceSize As Double
    SourceModified As Double
    SourceAttributes As Long
    SourceReparse As Boolean
    SourceFingerprint As String
    TemporarySource As String
    TemporaryNonce As String
    IsDelete As Boolean
    IsFolder As Boolean
End Type

Private mPlan() As OperationPlanItem
Private mPlanCount As Long
Private mPlanRoot As String
Private mPlanBatchId As String
Private mPlanNonce As String
Private mPlanFingerprint As String
Private mDraftFingerprint As String
Private mPlanApproved As Boolean
Private mOperationBusy As Boolean
Private Const OP_MAX_BATCH_ITEMS As Long = 10000
Private Const OP_RESERVED_PREFIX As String = ".folder-tree-v1_1-tmp-"

Public Sub OperationDraftCreated(ByVal rootPath As String)
    mPlanCount = 0
    Erase mPlan
    mPlanRoot = NormalizeOperationPath(rootPath)
    mPlanBatchId = ""
    mPlanNonce = ""
    mPlanFingerprint = ""
    mDraftFingerprint = ""
    mPlanApproved = False
End Sub

Public Function OperationIsBusy() As Boolean
    OperationIsBusy = mOperationBusy
End Function

Public Sub PrepareOperationSheet()
    If mOperationBusy Then
        MsgBox "ファイル操作が進行中です。完了後に再実行してください。", vbExclamation, "変更操作"
        Exit Sub
    End If

    Dim rootPath As String
    rootPath = Trim$(CStr(ThisWorkbook.Worksheets("設定").Range("C5").Value2))
    If Not IsScanReadyForRoot(rootPath) Then
        MsgBox "正常に完了した同じ走査rootの結果がありません。先に [2. ツリーを作成] を実行してください。", vbExclamation, "変更操作"
        Exit Sub
    End If

    Dim recoveryMessage As String
    If Not CheckReservedTemporaryNames(rootPath, recoveryMessage) Then
        MsgBox recoveryMessage, vbCritical, "残留一時名を検出しました"
        Exit Sub
    End If

    If Not WorkbookIoCreateOperationDraft(rootPath) Then Exit Sub
    OperationDraftCreated rootPath
    MsgBox "変更操作シートを作成しました。操作種別と変更案を編集してから [事前確認] を実行してください。", vbInformation, "変更操作"
End Sub

Public Sub RecreateOperationSheet()
    If mPlanApproved Then
        If MsgBox("事前確認済みの操作案を破棄して一覧を再作成しますか？", vbYesNo + vbQuestion, "変更操作") <> vbYes Then Exit Sub
    End If
    PrepareOperationSheet
End Sub

Public Sub SetOperationSequencePreview()
    If mOperationBusy Then
        MsgBox "ファイル操作が進行中です。", vbExclamation, "変更操作"
        Exit Sub
    End If
    If Not OperationDraftExists() Then
        MsgBox "先に [変更一覧を作成] を実行してください。", vbExclamation, "変更操作"
        Exit Sub
    End If
    WorkbookIoSetSequencePreview
    mPlanApproved = False
    mPlanFingerprint = ""
End Sub

Public Sub PreviewOperationPlan()
    If mOperationBusy Then
        MsgBox "ファイル操作が進行中です。", vbExclamation, "変更操作"
        Exit Sub
    End If
    If Not OperationDraftExists() Then
        MsgBox "先に [変更一覧を作成] を実行してください。", vbExclamation, "変更操作"
        Exit Sub
    End If

    Dim message As String
    If Not BuildOperationPlan(message) Then
        WorkbookIoSetOperationState "検査失敗"
        MsgBox message, vbExclamation, "事前確認で停止しました"
        Exit Sub
    End If

    mPlanApproved = True
    mPlanFingerprint = mDraftFingerprint
    WorkbookIoSetOperationState "事前確認済み"
    MsgBox message & vbCrLf & vbCrLf & _
           "この内容を確認した後、明示的に [変更を実行] を押してください。", _
           vbInformation, "事前確認"
End Sub

#If TEST_BUILD Then
' UIを表示しない受入用entrypoint。成功側の判定を注入するhookではなく、
' 実際のworksheet値を同じBuildOperationPlanへ渡すだけに限定する。
Public Function RunOperationPlanTest() As String
    Dim message As String
    If BuildOperationPlan(message) Then
        mPlanApproved = True
        mPlanFingerprint = mDraftFingerprint
        RunOperationPlanTest = "pass|" & message
    Else
        mPlanApproved = False
        RunOperationPlanTest = "fail|" & message
    End If
End Function

Public Function GetOperationConfirmationTextTest() As String
    If Not mPlanApproved Or mPlanCount = 0 Then
        GetOperationConfirmationTextTest = "not-approved"
    Else
        GetOperationConfirmationTextTest = OperationConfirmationText()
    End If
End Function

Public Function ProbeOperationReentryTest(ByVal rootPath As String) As String
    On Error GoTo Fail
    mOperationBusy = True
    SetScanBlockedByOperation True
    ProbeOperationReentryTest = RunScan(rootPath, 0, True, False)
    SetScanBlockedByOperation False
    mOperationBusy = False
    Exit Function
Fail:
    SetScanBlockedByOperation False
    mOperationBusy = False
    Err.Raise Err.Number, "modOperationPlan.ProbeOperationReentryTest", Err.Description
End Function

Public Function ExecuteOperationPlanTest() As String
    Dim resultText As String
    If Not mPlanApproved Or mPlanCount = 0 Then
        ExecuteOperationPlanTest = "not-approved"
        Exit Function
    End If
    Dim recheckMessage As String
    If Not RecheckApprovedPlan(recheckMessage) Then
        ExecuteOperationPlanTest = "stale|" & recheckMessage
        Exit Function
    End If
    On Error GoTo Fail
    ExecuteApprovedPlanCore resultText
    mPlanApproved = False
    ExecuteOperationPlanTest = resultText
    Exit Function
Fail:
    mOperationBusy = False
    mPlanApproved = False
    ExecuteOperationPlanTest = "error|" & Err.Number & ":" & Err.Description
End Function
#End If

Public Sub ExecuteOperationPlan()
    If mOperationBusy Then
        MsgBox "ファイル操作が進行中です。", vbExclamation, "変更操作"
        Exit Sub
    End If
    If Not mPlanApproved Or mPlanCount = 0 Then
        MsgBox "先に [事前確認] を完了してください。", vbExclamation, "変更操作"
        Exit Sub
    End If
    If BuildDraftFingerprint() <> mPlanFingerprint Then
        mPlanApproved = False
        WorkbookIoSetOperationState "再確認が必要"
        MsgBox "事前確認後に操作案が変更されています。再度 [事前確認] を実行してください。", vbExclamation, "変更操作"
        Exit Sub
    End If
    Dim recheckMessage As String
    If Not RecheckApprovedPlan(recheckMessage) Then
        mPlanApproved = False
        WorkbookIoSetOperationState "再確認が必要"
        MsgBox recheckMessage, vbExclamation, "変更操作"
        Exit Sub
    End If

    Dim answer As VbMsgBoxResult
    answer = MsgBox(OperationConfirmationText(), vbYesNo + vbExclamation, "変更を実行しますか？")
    If answer <> vbYes Then Exit Sub

    On Error GoTo Fail
    WorkbookIoSetOperationState "実行中"
    Dim resultText As String
    ExecuteApprovedPlanCore resultText
    mPlanApproved = False
    WorkbookIoSetOperationState "実行完了"
    MsgBox resultText, vbInformation, "変更操作"
    Exit Sub
Fail:
    mOperationBusy = False
    mPlanApproved = False
    WorkbookIoSetOperationState "失敗"
    MsgBox "操作を停止しました。" & vbCrLf & Err.Number & ": " & Err.Description, vbCritical, "変更操作"
End Sub

Private Sub ExecuteApprovedPlanCore(ByRef resultText As String)
    Dim savedScreenUpdating As Boolean, savedEnableEvents As Boolean
    Dim savedCalculation As Long, savedCursor As Long, savedEnableCancelKey As Long
    Dim savedStatusBar As Variant
    Dim savedNumber As Long, savedDescription As String
    Dim refreshSummary As String

    savedScreenUpdating = Application.ScreenUpdating
    savedEnableEvents = Application.EnableEvents
    savedCalculation = Application.Calculation
    savedCursor = Application.Cursor
    savedEnableCancelKey = Application.EnableCancelKey
    savedStatusBar = Application.StatusBar
    mOperationBusy = True
    SetScanBlockedByOperation True
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.EnableCancelKey = xlErrorHandler
    Application.Cursor = xlWait
    Application.StatusBar = "ファイル操作を実行しています..."
    resultText = ExecuteFileOperations(mPlan, mPlanCount, mPlanBatchId, mPlanRoot)
    mOperationBusy = False
    SetScanBlockedByOperation False
    Application.StatusBar = "操作結果を再走査しています..."
    refreshSummary = RefreshTreeAfterOperations(mPlanRoot)
    RestoreOperationApplicationState savedScreenUpdating, savedEnableEvents, savedCalculation, _
                                     savedCursor, savedEnableCancelKey, savedStatusBar
    Exit Sub
Fail:
    savedNumber = Err.Number
    savedDescription = Err.Description
    mOperationBusy = False
    SetScanBlockedByOperation False
    On Error Resume Next
    RestoreOperationApplicationState savedScreenUpdating, savedEnableEvents, savedCalculation, _
                                     savedCursor, savedEnableCancelKey, savedStatusBar
    On Error GoTo 0
    Err.Raise savedNumber, "modOperationPlan.ExecuteApprovedPlanCore", savedDescription
End Sub

Private Sub RestoreOperationApplicationState(ByVal screenUpdating As Boolean, ByVal enableEvents As Boolean, _
                                             ByVal calculation As Long, ByVal cursor As Long, _
                                             ByVal enableCancelKey As Long, ByVal statusBar As Variant)
    Application.ScreenUpdating = screenUpdating
    Application.EnableEvents = enableEvents
    Application.Calculation = calculation
    Application.Cursor = cursor
    Application.EnableCancelKey = enableCancelKey
    Application.StatusBar = statusBar
End Sub

Private Function BuildOperationPlan(ByRef resultMessage As String) As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim kind As String, source As String, newName As String, moveRel As String, folderRel As String
    Dim itemType As String, errorCount As Long, warningCount As Long
    Dim item As OperationPlanItem
    Dim draft As String

    Set ws = ThisWorkbook.Worksheets("変更操作")
    ResetRecycleBatchWarning
    mPlanCount = 0
    Erase mPlan
    mPlanRoot = NormalizeOperationPath(CStr(ws.Range("B2").Value2))
    If Len(mPlanRoot) = 0 Or Not IsScanReadyForRoot(mPlanRoot) Then
        resultMessage = "一覧のrootが最新の正常走査rootと一致しません。"
        Exit Function
    End If

    draft = ""
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < OP_FIRST_DATA_ROW Then
        resultMessage = "操作対象がありません。"
        Exit Function
    End If

    For r = OP_FIRST_DATA_ROW To lastRow
        kind = Trim$(CStr(ws.Cells(r, OP_COL_KIND).Value2))
        newName = Trim$(CStr(ws.Cells(r, OP_COL_NEW_NAME).Value2))
        moveRel = Trim$(CStr(ws.Cells(r, OP_COL_MOVE_REL).Value2))
        folderRel = Trim$(CStr(ws.Cells(r, OP_COL_FOLDER_REL).Value2))
        draft = draft & CStr(r) & Chr$(30) & kind & Chr$(30) & newName & Chr$(30) & moveRel & Chr$(30) & folderRel & Chr$(31)

        If ws.Cells(r, OP_COL_KIND).HasFormula Or ws.Cells(r, OP_COL_NEW_NAME).HasFormula Or _
           ws.Cells(r, OP_COL_MOVE_REL).HasFormula Or ws.Cells(r, OP_COL_FOLDER_REL).HasFormula Then
            SetOperationRowResult ws, r, "拒否", "編集可能列に数式があります。値へ置き換えてから再確認してください.", ""
            errorCount = errorCount + 1
            GoTo NextRow
        End If
        If kind = "" Or kind = OP_NONE Then
            ' 空行／変更なし行は計画へ入らない。ここで毎行 Protect/Unprotect を
            ' 行うと、対象件数の多い一覧でExcelのCOM応答が停止し得るため、
            ' 実際に検査した行だけ状態を書き込む。
            GoTo NextRow
        End If

        source = CStr(ws.Cells(r, OP_COL_FULL_PATH).Value2)
        itemType = CStr(ws.Cells(r, OP_COL_INTERNAL_TYPE).Value2)
        If Len(source) = 0 And kind <> OP_MKDIR Then
            SetOperationRowResult ws, r, "拒否", "sourceがありません。", ""
            errorCount = errorCount + 1
            GoTo NextRow
        End If
        If Not ValidateOperationItem(ws, r, kind, source, itemType, newName, moveRel, folderRel, item, resultMessage) Then
            SetOperationRowResult ws, r, "拒否", resultMessage, ""
            errorCount = errorCount + 1
            GoTo NextRow
        End If
        AddOperationItem item
        SetOperationRowResult ws, r, "検査OK", "", item.Destination
NextRow:
    Next r

    mDraftFingerprint = DraftFingerprintFromText(draft)
    If errorCount > 0 Then
        resultMessage = "検査結果: エラー " & errorCount & " 件。安全のため実行できません。"
        BuildOperationPlan = False
        Exit Function
    End If
    If mPlanCount = 0 Then
        resultMessage = "操作対象がありません。"
        BuildOperationPlan = False
        Exit Function
    End If
    If mPlanCount > OP_MAX_BATCH_ITEMS Then
        resultMessage = "1 batchの操作件数上限 " & OP_MAX_BATCH_ITEMS & " 件を超えています。対象rootを分けてください。"
        BuildOperationPlan = False
        Exit Function
    End If
    mPlanBatchId = Format$(Now, "yyyymmddhhnnss") & "-" & Replace$(Mid$(CStr(Timer), 3), ".", "")
    mPlanNonce = NewOperationNonce()
    If Not ValidatePlanCollisions(resultMessage) Then
        BuildOperationPlan = False
        Exit Function
    End If
    If Not PrepareRenameTemporaryNames(resultMessage) Then
        BuildOperationPlan = False
        Exit Function
    End If
    If Len(GetRecycleBatchWarning()) > 0 Then warningCount = warningCount + 1
    resultMessage = BuildPlanSummary(errorCount, warningCount)
    If Len(GetRecycleBatchWarning()) > 0 Then resultMessage = resultMessage & vbCrLf & "警告: " & GetRecycleBatchWarning()
    BuildOperationPlan = True
End Function

Private Function ValidateOperationItem(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal kind As String, _
                                       ByVal source As String, ByVal itemType As String, ByVal newName As String, _
                                       ByVal moveRel As String, ByVal folderRel As String, _
                                       ByRef item As OperationPlanItem, ByRef errorMessage As String) As Boolean
    Dim attrs As Long, expectedAttrs As Long, target As String, parentPath As String
    Dim expectedSize As Double, expectedModified As Double
    source = NormalizeOperationPath(source)
    item.WorksheetRow = rowNumber
    item.Kind = kind
    item.Source = ""
    item.Destination = ""
    item.RelativeDestination = ""
    item.SourceType = ""
    item.SourceSize = 0
    item.SourceModified = 0
    item.SourceAttributes = 0
    item.SourceReparse = False
    item.SourceFingerprint = ""
    item.TemporarySource = ""
    item.TemporaryNonce = ""
    item.IsDelete = False
    item.IsFolder = False
    item.Source = source
    item.OriginalName = CStr(ws.Cells(rowNumber, OP_COL_ORIGINAL_NAME).Value2)
    item.NewName = newName
    item.SourceType = itemType
    item.SourceSize = CDbl(Val(CStr(ws.Cells(rowNumber, OP_COL_SIZE).Value2)))
    item.SourceModified = CDbl(Val(CStr(ws.Cells(rowNumber, OP_COL_DATE).Value2)))
    attrs = GetOperationAttributes(source)
    item.SourceAttributes = attrs
    item.SourceReparse = ((attrs And FILE_ATTRIBUTE_REPARSE_POINT) <> 0)
    item.IsFolder = ((attrs And FILE_ATTRIBUTE_DIRECTORY) <> 0)
    item.SourceFingerprint = BuildSourceFingerprint(source, attrs, item.SourceSize, item.SourceModified)
    expectedAttrs = CLng(Val(CStr(ws.Cells(rowNumber, OP_COL_ATTRIBUTES).Value2)))
    expectedSize = CDbl(Val(CStr(ws.Cells(rowNumber, OP_COL_SOURCE_SIZE).Value2)))
    expectedModified = CDbl(Val(CStr(ws.Cells(rowNumber, OP_COL_SOURCE_DATE).Value2)))

    If kind = OP_MKDIR Then
        If Len(folderRel) = 0 Then errorMessage = "新規フォルダ相対パスが空です。": Exit Function
        If Not IsSafeRelativePath(folderRel) Then errorMessage = "新規フォルダ相対パスが不正です。": Exit Function
        target = CombineOperationPath(mPlanRoot, folderRel)
        If Not IsInsideOperationRoot(target) Or PathExists(target) Then errorMessage = "新規フォルダのroot境界または衝突を検査できません。": Exit Function
        If HasReparseAncestor(target, mPlanRoot) Then errorMessage = "新規フォルダの祖先にreparse pointがあります。": Exit Function
        item.Destination = target
        item.RelativeDestination = folderRel
        item.IsFolder = True
        ValidateOperationItem = True
        Exit Function
    End If

    If Not PathExists(source) Then errorMessage = "sourceが存在しません。": Exit Function
    If Not IsInsideOperationRoot(source) Or source = mPlanRoot Then errorMessage = "sourceが走査root境界外です。": Exit Function
    If item.SourceReparse Then errorMessage = "reparse pointは操作対象にできません。": Exit Function
    If Not VerifyOperationSnapshot(source, expectedAttrs, expectedSize, expectedModified, errorMessage) Then Exit Function

    Select Case kind
        Case OP_RENAME_MOVE
            If item.IsFolder Then errorMessage = "フォルダの名前変更／移動はv1.1対象外です。": Exit Function
            If Len(newName) = 0 Then newName = item.OriginalName
            If Not IsSafeName(newName) Then errorMessage = "新しい名前がWindowsの禁止規則に違反しています。": Exit Function
            If Len(moveRel) > 0 And Not IsSafeRelativePath(moveRel) Then errorMessage = "移動先相対フォルダが不正です。": Exit Function
            parentPath = GetParentOperationPath(source)
            If Len(moveRel) > 0 Then parentPath = CombineOperationPath(mPlanRoot, moveRel)
            target = CombineOperationPath(parentPath, newName)
            If Not IsInsideOperationRoot(target) Or HasReparseAncestor(target, mPlanRoot) Then errorMessage = "destinationがroot境界またはreparse境界外です。": Exit Function
            If PathExists(target) And StrComp(target, source, vbTextCompare) <> 0 Then
                ' A/B交換は後段で予約一時名へ展開する。その他の既存destinationは拒否する。
                If Not IsRenameSourceInPlan(target) And Not IsDraftRenameSource(ws, target) Then errorMessage = "destinationが既存項目と衝突します。": Exit Function
            End If
            item.NewName = newName: item.Destination = target: item.IsDelete = False
        Case OP_RECYCLE_FILE
            If item.IsFolder Then errorMessage = "フォルダはファイル削除操作へ指定できません。": Exit Function
            If Not CanRecycleOperation(source, item.SourceSize, errorMessage) Then Exit Function
            item.Destination = "[ごみ箱]": item.IsDelete = True
        Case OP_RECYCLE_FOLDER
            If Not item.IsFolder Then errorMessage = "空フォルダ削除にはフォルダを指定してください。": Exit Function
            If Not IsOperationFolderEmpty(source) Then errorMessage = "フォルダが空ではありません。": Exit Function
            If Not CanRecycleOperation(source, 0, errorMessage) Then Exit Function
            item.Destination = "[ごみ箱]": item.IsDelete = True
        Case Else
            errorMessage = "不明な操作種別です。": Exit Function
    End Select
    ValidateOperationItem = True
End Function

Private Sub AddOperationItem(ByRef item As OperationPlanItem)
    mPlanCount = mPlanCount + 1
    ReDim Preserve mPlan(1 To mPlanCount)
    item.Sequence = mPlanCount
    mPlan(mPlanCount) = item
End Sub

Private Function OperationConfirmationText() As String
    Dim warningCount As Long
    If Len(GetRecycleBatchWarning()) > 0 Then warningCount = 1
    OperationConfirmationText = BuildPlanSummary(0, warningCount) & vbCrLf & vbCrLf & _
        "削除開始後は自動rollbackできません。ゴミ箱移動に失敗した場合も完全削除へ切り替えません。" & vbCrLf & _
        "実行してよろしいですか？"
End Function

Private Function OperationDraftExists() As Boolean
    OperationDraftExists = (SheetExists("変更操作") And Len(CStr(ThisWorkbook.Worksheets("変更操作").Range("B2").Value2)) > 0)
End Function

Private Function BuildDraftFingerprint() As String
    Dim ws As Worksheet, lastRow As Long, r As Long, s As String
    Set ws = ThisWorkbook.Worksheets("変更操作")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = OP_FIRST_DATA_ROW To lastRow
        s = s & CStr(r) & Chr$(30) & CStr(ws.Cells(r, OP_COL_KIND).Value2) & Chr$(30) & _
            CStr(ws.Cells(r, OP_COL_NEW_NAME).Value2) & Chr$(30) & CStr(ws.Cells(r, OP_COL_MOVE_REL).Value2) & Chr$(30) & _
            CStr(ws.Cells(r, OP_COL_FOLDER_REL).Value2) & Chr$(31)
    Next r
    BuildDraftFingerprint = DraftFingerprintFromText(s)
End Function

Private Function DraftFingerprintFromText(ByVal value As String) As String
    ' VBA標準だけではSHA-256を安全に共通化できないため、canonical文字列そのものを
    ' セッション内fingerprintとして保持する。セルには保存せず、比較時は完全一致させる。
    DraftFingerprintFromText = "draft:" & value
End Function

Private Function NormalizeOperationPath(ByVal path As String) As String
    Dim fso As Object
    path = Trim$(Replace$(path, "/", "\"))
    If Len(path) = 0 Then Exit Function
    If (Len(path) >= 3 And Mid$(path, 2, 2) = ":\" Or Left$(path, 2) = "\\") And _
       InStr(1, path, "\.\", vbBinaryCompare) = 0 And InStr(1, path, "\..\", vbBinaryCompare) = 0 Then
        NormalizeOperationPath = path
    Else
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next
        NormalizeOperationPath = fso.GetAbsolutePathName(path)
        On Error GoTo 0
    End If
    If Right$(NormalizeOperationPath, 1) = "\" And Len(NormalizeOperationPath) > 3 Then NormalizeOperationPath = Left$(NormalizeOperationPath, Len(NormalizeOperationPath) - 1)
End Function

Private Function CombineOperationPath(ByVal basePath As String, ByVal relativePath As String) As String
    If Len(relativePath) = 0 Or relativePath = "." Then
        CombineOperationPath = NormalizeOperationPath(basePath)
    Else
        CombineOperationPath = NormalizeOperationPath(basePath & "\" & Replace$(relativePath, "/", "\"))
    End If
End Function

Private Function IsInsideOperationRoot(ByVal path As String) As Boolean
    path = NormalizeOperationPath(path)
    IsInsideOperationRoot = (StrComp(path, mPlanRoot, vbTextCompare) = 0 Or _
        LCase$(Left$(path, Len(mPlanRoot) + 1)) = LCase$(mPlanRoot & "\"))
End Function

Private Function IsSafeRelativePath(ByVal path As String) As Boolean
    Dim normalized As String, parts() As String, i As Long
    normalized = Replace$(path, "/", "\")
    If Len(normalized) = 0 Or Left$(normalized, 1) = "\" Or Right$(normalized, 1) = "\" Then Exit Function
    If InStr(normalized, ":") > 0 Or Left$(normalized, 2) = "\\" Or Left$(normalized, 4) = "\\?\" Then Exit Function
    parts = Split(normalized, "\")
    For i = LBound(parts) To UBound(parts)
        If Not IsSafeName(parts(i)) Then Exit Function
    Next i
    IsSafeRelativePath = (Len(normalized) <= 240)
End Function

Private Function IsSafeName(ByVal value As String) As Boolean
    Dim reserved As String, i As Long, ch As String, stem As String, codeUnit As Long
    If Len(value) = 0 Or value = "." Or value = ".." Then Exit Function
    If Right$(value, 1) = " " Or Right$(value, 1) = "." Then Exit Function
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        codeUnit = AscW(ch)
        If codeUnit < 0 Then codeUnit = codeUnit + 65536
        If codeUnit < 32 Or InStr(1, "<>:" & Chr$(92) & "/|?*", ch, vbBinaryCompare) > 0 Then Exit Function
    Next i
    If Len(value) > 255 Then Exit Function
    stem = UCase$(Split(value, ".")(0))
    reserved = "CON PRN AUX NUL COM1 COM2 COM3 COM4 COM5 COM6 COM7 COM8 COM9 LPT1 LPT2 LPT3 LPT4 LPT5 LPT6 LPT7 LPT8 LPT9"
    If InStr(" " & reserved & " ", " " & stem & " ") > 0 Then Exit Function
    IsSafeName = True
End Function

Private Function BuildSourceFingerprint(ByVal source As String, ByVal attrs As Long, _
                                        ByVal itemSize As Double, ByVal itemModified As Double) As String
    BuildSourceFingerprint = LCase$(NormalizeOperationPath(source)) & "|" & CStr(attrs) & "|" & _
                              Format$(itemSize, "0.################") & "|" & Format$(itemModified, "0.################")
End Function

Private Function IsRenameSourceInPlan(ByVal path As String) As Boolean
    Dim i As Long
    For i = 1 To mPlanCount
        If mPlan(i).Kind = OP_RENAME_MOVE Then
            If StrComp(mPlan(i).Source, path, vbTextCompare) = 0 Then IsRenameSourceInPlan = True: Exit Function
        End If
    Next i
End Function

Private Function IsDraftRenameSource(ByVal ws As Worksheet, ByVal path As String) As Boolean
    Dim r As Long, lastRow As Long, kind As String, candidate As String
    lastRow = ws.Cells(ws.Rows.Count, OP_COL_TARGET).End(xlUp).Row
    For r = OP_FIRST_DATA_ROW To lastRow
        kind = Trim$(CStr(ws.Cells(r, OP_COL_KIND).Value2))
        If kind = OP_RENAME_MOVE Then
            candidate = NormalizeOperationPath(CStr(ws.Cells(r, OP_COL_FULL_PATH).Value2))
            If StrComp(candidate, path, vbTextCompare) = 0 Then IsDraftRenameSource = True: Exit Function
        End If
    Next r
End Function

Private Function ValidatePlanCollisions(ByRef errorMessage As String) As Boolean
    Dim i As Long, j As Long
    For i = 1 To mPlanCount
        If mPlan(i).Kind = OP_RENAME_MOVE Or mPlan(i).Kind = OP_MKDIR Then
            For j = i + 1 To mPlanCount
                If Len(mPlan(j).Destination) > 0 And StrComp(mPlan(i).Destination, mPlan(j).Destination, vbTextCompare) = 0 Then
                    errorMessage = "destinationが同じ操作で重複しています。"
                    SetOperationRowResult ThisWorkbook.Worksheets("変更操作"), mPlan(i).WorksheetRow, "拒否", errorMessage, mPlan(i).Destination
                    SetOperationRowResult ThisWorkbook.Worksheets("変更操作"), mPlan(j).WorksheetRow, "拒否", errorMessage, mPlan(j).Destination
                    Exit Function
                End If
            Next j
        End If
    Next i
    ValidatePlanCollisions = True
End Function

Private Function PrepareRenameTemporaryNames(ByRef errorMessage As String) As Boolean
    Dim i As Long, candidate As String, n As Long
    For i = 1 To mPlanCount
        If mPlan(i).Kind = OP_RENAME_MOVE Then
            If StrComp(mPlan(i).Source, mPlan(i).Destination, vbTextCompare) = 0 Or _
               IsRenameSourceInPlan(mPlan(i).Destination) Then
                n = n + 1
                candidate = CombineOperationPath(mPlanRoot, OP_RESERVED_PREFIX & mPlanBatchId & "-" & CStr(n) & "-" & mPlanNonce)
                Do While PathExists(candidate) Or IsRenameDestinationInPlan(candidate)
                    n = n + 1
                    candidate = CombineOperationPath(mPlanRoot, OP_RESERVED_PREFIX & mPlanBatchId & "-" & CStr(n) & "-" & mPlanNonce)
                    If n > mPlanCount + 100 Then errorMessage = "予約一時名を確保できません。": Exit Function
                Loop
                mPlan(i).TemporarySource = candidate
                mPlan(i).TemporaryNonce = mPlanNonce
            End If
        End If
    Next i
    PrepareRenameTemporaryNames = True
End Function

Private Function NewOperationNonce() As String
    Randomize
    NewOperationNonce = Right$("00000000" & Hex$(CLng(Timer * 1000) And &H7FFFFFFF), 8) & _
                        Right$("00000000" & Hex$(CLng(Rnd() * 2147483647#)), 8)
End Function

#If TEST_BUILD Then
Public Function CheckReservedTemporaryNamesTest(ByVal rootPath As String) As String
    Dim message As String
    If CheckReservedTemporaryNames(rootPath, message) Then
        CheckReservedTemporaryNamesTest = "pass|" & message
    Else
        CheckReservedTemporaryNamesTest = "fail|" & message
    End If
End Function
#End If

Private Function CheckReservedTemporaryNames(ByVal rootPath As String, ByRef message As String) As Boolean
    Dim count As Long, i As Long, itemName As String, itemType As String, fullPath As String
    Dim depth As Long, items As Variant, itemSize As Double, modified As Variant, attrs As Long, reparse As Boolean
    Dim source As String, batchId As String, nonce As String, sequence As Long, matches As Long
    Dim recovery As String, failures As String

    rootPath = NormalizeOperationPath(rootPath)
    If Not IsScanReadyForRoot(rootPath) Then message = "最新の正常走査結果がありません。": Exit Function
    count = GetScanSnapshotCount()
    For i = 1 To count
        If Not GetScanSnapshotRow(i, itemName, itemType, depth, items, itemSize, modified, fullPath, attrs, reparse) Then
            message = "予約一時名の走査結果を取得できません。": Exit Function
        End If
        If StrComp(Left$(itemName, Len(OP_RESERVED_PREFIX)), OP_RESERVED_PREFIX, vbTextCompare) = 0 Then
            FindTemporaryLogMapping fullPath, matches, source, batchId, sequence, nonce
            If matches = 0 Then
                failures = failures & vbCrLf & "orphan: " & fullPath
            ElseIf matches > 1 Then
                failures = failures & vbCrLf & "重複対応: " & fullPath
            ElseIf Len(source) = 0 Or PathExists(source) Then
                failures = failures & vbCrLf & "元path衝突: " & source & " <- " & fullPath
            Else
                recovery = recovery & vbCrLf & batchId & "/" & CStr(sequence) & "/" & nonce & ": " & fullPath & " -> " & source
            End If
        End If
    Next i
    If Len(failures) > 0 Then
        message = "残留予約prefixを安全に復元できません。新しいbatchを開始しません。" & failures
        Exit Function
    End If
    If Len(recovery) > 0 Then
        message = "残留予約prefixの復元が必要です。新しいbatchを開始しません。" & recovery
        Exit Function
    End If
    message = "残留予約prefixなし"
    CheckReservedTemporaryNames = True
End Function

Private Sub FindTemporaryLogMapping(ByVal temporaryPath As String, ByRef matches As Long, _
                                    ByRef source As String, ByRef batchId As String, _
                                    ByRef sequence As Long, ByRef nonce As String)
    Dim ws As Worksheet, lastRow As Long, r As Long
    matches = 0: source = "": batchId = "": sequence = 0: nonce = ""
    If Not SheetExists("実行ログ") Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("実行ログ")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = LOG_FIRST_DATA_ROW To lastRow
        If StrComp(NormalizeOperationPath(CStr(ws.Cells(r, 11).Value2)), NormalizeOperationPath(temporaryPath), vbTextCompare) = 0 Then
            matches = matches + 1
            source = NormalizeOperationPath(CStr(ws.Cells(r, 12).Value2))
            batchId = CStr(ws.Cells(r, 1).Value2)
            sequence = CLng(Val(CStr(ws.Cells(r, 2).Value2)))
            nonce = CStr(ws.Cells(r, 20).Value2)
        End If
    Next r
End Sub

Private Function BuildPlanSummary(ByVal errorCount As Long, ByVal warningCount As Long) As String
    Dim fileDeletes As Long, folderDeletes As Long, examples As String
    fileDeletes = CountPlanKind(OP_RECYCLE_FILE)
    folderDeletes = CountPlanKind(OP_RECYCLE_FOLDER)
    examples = RepresentativeOperations(3)
    BuildPlanSummary = "検査OK: " & mPlanCount & " 件" & vbCrLf & _
        "名前変更: " & CountRenameOrMove(False) & " 件" & vbCrLf & _
        "ファイル移動: " & CountRenameOrMove(True) & " 件" & vbCrLf & _
        "フォルダ作成: " & CountPlanKind(OP_MKDIR) & " 件" & vbCrLf & _
        "ファイル削除: " & fileDeletes & " 件" & vbCrLf & _
        "ファイル削除phase想定所要時間: " & FileDeleteEstimate(fileDeletes) & vbCrLf & _
        "空フォルダ削除: " & folderDeletes & " 件" & vbCrLf & _
        "エラー件数: " & errorCount & " 件" & vbCrLf & _
        "警告件数: " & warningCount & " 件" & vbCrLf & _
        "代表的な source -> destination:" & examples & vbCrLf & _
        "batch ID: " & mPlanBatchId
    If fileDeletes + folderDeletes > 0 Then
        BuildPlanSummary = BuildPlanSummary & vbCrLf & _
            "削除phaseは開始後にEscで中断できず、所要時間は対象数に比例します。"
    End If
End Function

Private Function CountRenameOrMove(ByVal moves As Boolean) As Long
    Dim i As Long, isMove As Boolean
    For i = 1 To mPlanCount
        If mPlan(i).Kind = OP_RENAME_MOVE Then
            isMove = (StrComp(GetParentOperationPath(mPlan(i).Source), GetParentOperationPath(mPlan(i).Destination), vbTextCompare) <> 0)
            If isMove = moves Then CountRenameOrMove = CountRenameOrMove + 1
        End If
    Next i
End Function

Private Function FileDeleteEstimate(ByVal count As Long) As String
    If count = 0 Then
        FileDeleteEstimate = "対象なし"
    ElseIf count <= 100 Then
        FileDeleteEstimate = "最大5秒"
    ElseIf count <= 1000 Then
        FileDeleteEstimate = "最大30秒"
    Else
        FileDeleteEstimate = "見積不能"
    End If
End Function

Private Function RepresentativeOperations(ByVal maximum As Long) As String
    Dim i As Long, shown As Long
    For i = 1 To mPlanCount
        If shown >= maximum Then Exit For
        RepresentativeOperations = RepresentativeOperations & vbCrLf & "  " & _
            IIf(Len(mPlan(i).Source) > 0, mPlan(i).Source, "[新規]") & " -> " & mPlan(i).Destination
        shown = shown + 1
    Next i
End Function

Private Function IsRenameDestinationInPlan(ByVal path As String) As Boolean
    Dim i As Long
    For i = 1 To mPlanCount
        If StrComp(mPlan(i).Destination, path, vbTextCompare) = 0 Then IsRenameDestinationInPlan = True: Exit Function
    Next i
End Function

Private Function GetParentOperationPath(ByVal path As String) As String
    Dim p As Long
    p = InStrRev(path, "\")
    If p > 0 Then GetParentOperationPath = Left$(path, p - 1)
End Function

Private Function PathExists(ByVal path As String) As Boolean
    PathExists = (GetOperationAttributes(path) <> -1)
End Function

Private Function SheetExists(ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function CountPlanKind(ByVal kind As String) As Long
    Dim i As Long
    For i = 1 To mPlanCount
        If mPlan(i).Kind = kind Then CountPlanKind = CountPlanKind + 1
    Next i
End Function

Public Function GetOperationPlanCount() As Long
    GetOperationPlanCount = mPlanCount
End Function

Public Function GetOperationPlanItem(ByVal index As Long) As OperationPlanItem
    GetOperationPlanItem = mPlan(index)
End Function

Public Function GetOperationPlanRoot() As String
    GetOperationPlanRoot = mPlanRoot
End Function

Public Function GetOperationPlanBatchId() As String
    GetOperationPlanBatchId = mPlanBatchId
End Function

Public Function RecheckApprovedPlan(ByRef message As String) As Boolean
    Dim i As Long
    message = ""
    If Not mPlanApproved Or mPlanCount = 0 Then message = "事前確認済みの操作案がありません。": Exit Function
    If BuildDraftFingerprint() <> mPlanFingerprint Then message = "事前確認後に操作案が変更されています。": Exit Function
    For i = 1 To mPlanCount
        If mPlan(i).Kind <> OP_MKDIR Then
            If Not VerifyOperationSnapshot(mPlan(i).Source, mPlan(i).SourceAttributes, mPlan(i).SourceSize, mPlan(i).SourceModified, message) Then
                message = "操作直前の再照合で停止しました。" & vbCrLf & message
                Exit Function
            End If
        End If
        If mPlan(i).Kind = OP_RENAME_MOVE And Len(mPlan(i).TemporarySource) = 0 Then
            If PathExists(mPlan(i).Destination) And StrComp(mPlan(i).Source, mPlan(i).Destination, vbTextCompare) <> 0 Then
                message = "操作直前にdestinationが衝突しました。"
                Exit Function
            End If
        End If
    Next i
    RecheckApprovedPlan = True
End Function

Public Sub MarkOperationPlanRow(ByVal sequence As Long, ByVal result As String, ByVal message As String, ByVal destination As String)
    Dim i As Long
    For i = 1 To mPlanCount
        If mPlan(i).Sequence = sequence Then
            SetOperationRowResult ThisWorkbook.Worksheets("変更操作"), mPlan(i).WorksheetRow, result, message, destination
            Exit Sub
        End If
    Next i
End Sub

Private Sub SetOperationRowResult(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal result As String, ByVal message As String, ByVal destination As String)
    WorkbookIoWriteOperationStatus rowNumber, result, message, destination
End Sub
