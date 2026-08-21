Attribute VB_Name = "modWorkbookIo"
Option Explicit

Public Const OP_HEADER_ROW As Long = 8
Public Const OP_FIRST_DATA_ROW As Long = 9
Public Const OP_COL_TARGET As Long = 1
Public Const OP_COL_KIND As Long = 2
Public Const OP_COL_DISPLAY_TYPE As Long = 3
Public Const OP_COL_RELATIVE As Long = 4
Public Const OP_COL_ORIGINAL_NAME As Long = 5
Public Const OP_COL_PARENT As Long = 6
Public Const OP_COL_ITEMS As Long = 7
Public Const OP_COL_SIZE As Long = 8
Public Const OP_COL_DATE As Long = 9
Public Const OP_COL_NEW_NAME As Long = 10
Public Const OP_COL_MOVE_REL As Long = 11
Public Const OP_COL_FOLDER_REL As Long = 12
Public Const OP_COL_PREVIEW As Long = 13
Public Const OP_COL_STATUS As Long = 14
Public Const OP_COL_MESSAGE As Long = 15
Public Const OP_COL_INTERNAL_TYPE As Long = 16
Public Const OP_COL_FULL_PATH As Long = 17
Public Const OP_COL_SOURCE_SIZE As Long = 18
Public Const OP_COL_SOURCE_DATE As Long = 19
Public Const OP_COL_ATTRIBUTES As Long = 20
Public Const OP_COL_REPARSE As Long = 21
Public Const OP_COL_FINGERPRINT As Long = 22

Public Const LOG_FIRST_DATA_ROW As Long = 3
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Public Const FILE_ATTRIBUTE_REPARSE_POINT As Long = &H400

Private Const OP_SHEET As String = "変更操作"
Private Const LOG_SHEET As String = "実行ログ"

Public Function WorkbookIoCreateOperationDraft(ByVal rootPath As String) As Boolean
    WorkbookIoCreateOperationDraft = WorkbookIoCreateOperationDraftCore(rootPath, True)
End Function

#If TEST_BUILD Then
' UIを表示しない性能・受入用entrypoint。実際のscan snapshotと同じcoreを使い、
' 成功結果や行数は注入しない。
Public Function WorkbookIoCreateOperationDraftTest(ByVal rootPath As String) As Boolean
    WorkbookIoCreateOperationDraftTest = WorkbookIoCreateOperationDraftCore(rootPath, False)
End Function
#End If

Private Function WorkbookIoCreateOperationDraftCore(ByVal rootPath As String, ByVal showMessages As Boolean) As Boolean
    Dim ws As Worksheet, r As Long, count As Long
    Dim errorNumber As Long, errorDescription As String
    Dim itemName As String, itemType As String, fullPath As String, modified As Variant
    Dim depth As Long, items As Variant, itemSize As Double, attrs As Long, reparse As Boolean
    Dim values() As Variant

    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(OP_SHEET)
    PrepareSheetForWrite ws
    ws.Range("B2").Value2 = CStr(rootPath)
    ws.Range("B3").Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Range("B4").Value2 = ""
    ws.Range("B5").Value2 = "一覧作成済み"
    ClearOperationRows ws
    count = GetScanSnapshotCount()
    If count > 100000 Then
        ws.Range("B5").Value2 = "一覧上限超過"
        ProtectSheetAfterWrite ws
        If showMessages Then MsgBox "走査結果が100,000行を超えています。一部だけの一覧は作成しません。対象サブフォルダを新しい走査rootとして再走査してください。", vbExclamation, "変更操作"
        Exit Function
    End If

    If count > 0 Then ReDim values(1 To count, 1 To OP_COL_FINGERPRINT)
    For r = 1 To count
        If Not GetScanSnapshotRow(r, itemName, itemType, depth, items, itemSize, modified, fullPath, attrs, reparse) Then GoTo Fail
        values(r, OP_COL_TARGET) = CStr(r)
        values(r, OP_COL_KIND) = "変更なし"
        values(r, OP_COL_DISPLAY_TYPE) = itemType
        values(r, OP_COL_RELATIVE) = RelativeOperationPath(rootPath, fullPath)
        values(r, OP_COL_ORIGINAL_NAME) = itemName
        values(r, OP_COL_PARENT) = RelativeOperationPath(rootPath, ParentPath(fullPath))
        values(r, OP_COL_ITEMS) = items
        values(r, OP_COL_SIZE) = itemSize
        values(r, OP_COL_DATE) = modified
        values(r, OP_COL_NEW_NAME) = ""
        values(r, OP_COL_MOVE_REL) = ""
        values(r, OP_COL_FOLDER_REL) = ""
        values(r, OP_COL_PREVIEW) = ""
        values(r, OP_COL_STATUS) = ""
        values(r, OP_COL_MESSAGE) = ""
        values(r, OP_COL_INTERNAL_TYPE) = IIf((attrs And FILE_ATTRIBUTE_DIRECTORY) <> 0, "フォルダ", "ファイル")
        values(r, OP_COL_FULL_PATH) = fullPath
        values(r, OP_COL_SOURCE_SIZE) = itemSize
        values(r, OP_COL_SOURCE_DATE) = modified
        values(r, OP_COL_ATTRIBUTES) = attrs
        values(r, OP_COL_REPARSE) = IIf(reparse, "True", "False")
        values(r, OP_COL_FINGERPRINT) = CStr(attrs) & "|" & CStr(itemSize) & "|" & CStr(modified) & "|" & fullPath
    Next r
    If count > 0 Then ws.Cells(OP_FIRST_DATA_ROW, 1).Resize(count, OP_COL_FINGERPRINT).Value2 = values
    ws.Range(ws.Cells(OP_FIRST_DATA_ROW, OP_COL_INTERNAL_TYPE), ws.Cells(OP_FIRST_DATA_ROW + count - 1, OP_COL_FINGERPRINT)).EntireColumn.Hidden = True
    ws.Range(ws.Cells(OP_HEADER_ROW, 1), ws.Cells(OP_HEADER_ROW + count, OP_COL_FINGERPRINT)).AutoFilter
    ProtectSheetAfterWrite ws
    WorkbookIoCreateOperationDraftCore = True
    Exit Function
Fail:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    ProtectSheetAfterWrite ws
    On Error GoTo 0
    If showMessages Then MsgBox "変更一覧を作成できませんでした。" & vbCrLf & errorNumber & ": " & errorDescription, vbCritical, "変更操作"
End Function

Public Sub WorkbookIoSetSequencePreview()
    Dim mode As String, startText As String, startNumber As Long
    Dim resultMessage As String, selected As Long
    mode = InputBox("先頭 または 末尾 を入力してください。", "連番の方式", "先頭")
    If mode <> "先頭" And mode <> "末尾" Then Exit Sub
    startText = InputBox("開始番号を入力してください。", "連番の開始番号", "0")
    If Not IsNumeric(startText) Then Exit Sub
    startNumber = CLng(startText)
    If Not WorkbookIoApplySequencePreview(mode, startNumber, selected, resultMessage) Then
        MsgBox resultMessage, vbExclamation, "連番"
        Exit Sub
    End If
    MsgBox resultMessage, vbInformation, "連番"
End Sub

#If TEST_BUILD Then
' UIを表示しない性能・受入用entrypoint。表示中の実worksheet行へ同じcoreを適用する。
Public Function WorkbookIoSetSequencePreviewTest(ByVal mode As String, ByVal startNumber As Long) As String
    Dim resultMessage As String, selected As Long
    If WorkbookIoApplySequencePreview(mode, startNumber, selected, resultMessage) Then
        WorkbookIoSetSequencePreviewTest = "pass|selected=" & CStr(selected)
    Else
        WorkbookIoSetSequencePreviewTest = "fail|" & resultMessage
    End If
End Function
#End If

Private Function WorkbookIoApplySequencePreview(ByVal mode As String, ByVal startNumber As Long, _
                                                  ByRef selected As Long, ByRef resultMessage As String) As Boolean
    Dim digits As Long, ws As Worksheet, lastRow As Long, r As Long, index As Long, name As String
    If mode <> "先頭" And mode <> "末尾" Then resultMessage = "連番方式が不正です。": Exit Function
    If startNumber < 0 Then resultMessage = "開始番号が不正です。": Exit Function
    Set ws = ThisWorkbook.Worksheets(OP_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, OP_COL_TARGET).End(xlUp).Row
    For r = OP_FIRST_DATA_ROW To lastRow
        If CStr(ws.Cells(r, OP_COL_INTERNAL_TYPE).Value2) = "ファイル" And _
           CStr(ws.Cells(r, OP_COL_KIND).Value2) = OP_NONE And Not ws.Rows(r).Hidden Then selected = selected + 1
    Next r
    If selected = 0 Then
        resultMessage = "表示中の通常ファイル行がありません。"
        Exit Function
    End If
    digits = Len(CStr(selected))
    If Len(CStr(selected - 1 + startNumber)) > digits Then digits = Len(CStr(selected - 1 + startNumber))
    If digits < 2 Then digits = 2
    PrepareSheetForWrite ws
    index = 0
    For r = OP_FIRST_DATA_ROW To lastRow
        If CStr(ws.Cells(r, OP_COL_INTERNAL_TYPE).Value2) = "ファイル" And _
           CStr(ws.Cells(r, OP_COL_KIND).Value2) = OP_NONE And Not ws.Rows(r).Hidden Then
            name = CStr(ws.Cells(r, OP_COL_ORIGINAL_NAME).Value2)
            ws.Cells(r, OP_COL_NEW_NAME).Value2 = SequenceName(name, mode, startNumber + index, digits)
            ws.Cells(r, OP_COL_STATUS).Value2 = "preview"
            index = index + 1
        End If
    Next r
    ProtectSheetAfterWrite ws
    resultMessage = "連番previewを作成しました。実ファイルは変更していません。"
    WorkbookIoApplySequencePreview = True
End Function

Private Function SequenceName(ByVal originalName As String, ByVal mode As String, ByVal number As Long, ByVal digits As Long) As String
    Dim prefix As String, p As Long, stem As String, ext As String
    prefix = Format$(number, String$(digits, "0")) & "_"
    If mode = "先頭" Then
        SequenceName = prefix & originalName
        Exit Function
    End If
    p = InStrRev(originalName, ".")
    If p <= 1 Then
        stem = originalName: ext = ""
    Else
        stem = Left$(originalName, p - 1): ext = Mid$(originalName, p)
    End If
    SequenceName = stem & "_" & Format$(number, String$(digits, "0")) & ext
End Function

Public Sub WorkbookIoSetOperationState(ByVal stateText As String)
    If Not SheetExistsIo(OP_SHEET) Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(OP_SHEET)
    PrepareSheetForWrite ws
    ws.Range("B5").Value2 = stateText
    ProtectSheetAfterWrite ws
End Sub

#If TEST_BUILD Then
Public Function WorkbookIoProtectionFailureProbeTest() As String
    Dim ws As Worksheet, beforeValue As String, probeSecret As String
    Dim caught As Boolean, caughtNumber As Long
    Set ws = ThisWorkbook.Worksheets(OP_SHEET)
    beforeValue = CStr(ws.Range("B5").Value2)
    probeSecret = Hex$(CLng(Timer * 1000)) & Hex$(CLng(Rnd() * 1000000))
    ws.Unprotect
    ws.Protect Password:=probeSecret
    On Error GoTo ExpectedFailure
    WorkbookIoSetOperationState "should-not-write"
    GoTo UnsafeCompletion
ExpectedFailure:
    caught = True
    caughtNumber = Err.Number
    Err.Clear
    On Error GoTo CleanupFailure
    ws.Unprotect Password:=probeSecret
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    probeSecret = ""
    If caught And caughtNumber <> 0 And ws.ProtectContents And CStr(ws.Range("B5").Value2) = beforeValue Then
        WorkbookIoProtectionFailureProbeTest = "pass|fail-closed"
    Else
        WorkbookIoProtectionFailureProbeTest = "fail|state-changed"
    End If
    Exit Function
UnsafeCompletion:
    On Error Resume Next
    ws.Unprotect Password:=probeSecret
    ws.Range("B5").Value2 = beforeValue
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    probeSecret = ""
    On Error GoTo 0
    WorkbookIoProtectionFailureProbeTest = "fail|write-completed"
    Exit Function
CleanupFailure:
    probeSecret = ""
    WorkbookIoProtectionFailureProbeTest = "fail|cleanup-error|" & CStr(Err.Number)
End Function
#End If

Public Sub WorkbookIoWriteOperationStatus(ByVal rowNumber As Long, ByVal resultText As String, ByVal message As String, ByVal destination As String)
    Dim ws As Worksheet
    If Not SheetExistsIo(OP_SHEET) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(OP_SHEET)
    PrepareSheetForWrite ws
    ws.Cells(rowNumber, OP_COL_PREVIEW).Value2 = destination
    ws.Cells(rowNumber, OP_COL_STATUS).Value2 = resultText
    ws.Cells(rowNumber, OP_COL_MESSAGE).Value2 = message
    ProtectSheetAfterWrite ws
End Sub

Public Sub AppendExecutionLog(ByVal batchId As String, ByVal sequence As Long, ByVal operationKind As String, _
                              ByVal source As String, ByVal destination As String, ByVal resultText As String, _
                              ByVal errorNumber As Long, ByVal errorMessage As String, ByVal rollbackResult As String, _
                              ByVal rollbackInfo As String, ByVal safetyState As String, ByVal operationRoot As String, _
                              Optional ByVal temporaryName As String = "", Optional ByVal temporaryPath As String = "", _
                              Optional ByVal nonce As String = "", Optional ByVal originalName As String = "")
    Dim ws As Worksheet, nextRow As Long, values(1 To 1, 1 To 20) As Variant
    If Not SheetExistsIo(LOG_SHEET) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    PrepareSheetForWrite ws
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    values(1, 1) = batchId: values(1, 2) = sequence: values(1, 3) = Format$(Now, "yyyy-mm-ddThh:nn:ss")
    values(1, 4) = Format$(Now, "yyyy-mm-ddThh:nn:ss"): values(1, 5) = operationRoot
    values(1, 6) = operationKind: values(1, 7) = source: values(1, 8) = destination
    If Len(originalName) = 0 Then originalName = FileNameFromOperationPath(source)
    values(1, 9) = originalName: values(1, 10) = temporaryName: values(1, 11) = temporaryPath: values(1, 12) = source
    values(1, 13) = destination: values(1, 14) = resultText: values(1, 15) = errorNumber
    values(1, 16) = errorMessage: values(1, 17) = rollbackResult: values(1, 18) = rollbackInfo
    values(1, 19) = safetyState: values(1, 20) = nonce
    ws.Cells(nextRow, 1).Resize(1, 20).Value2 = values
    ProtectSheetAfterWrite ws
End Sub

Private Sub ClearOperationRows(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= OP_FIRST_DATA_ROW Then ws.Range(ws.Cells(OP_FIRST_DATA_ROW, 1), ws.Cells(lastRow, OP_COL_FINGERPRINT)).ClearContents
End Sub

Private Sub PrepareSheetForWrite(ByVal ws As Worksheet)
    If ws Is Nothing Then Err.Raise 91, "modWorkbookIo.PrepareSheetForWrite", "対象sheetを取得できません。"
    If ws.ProtectContents Then ws.Unprotect Password:=""
    If ws.ProtectContents Then Err.Raise vbObjectError + 1101, "modWorkbookIo.PrepareSheetForWrite", "sheet保護を解除できません。"
End Sub

Private Sub ProtectSheetAfterWrite(ByVal ws As Worksheet)
    If ws Is Nothing Then Err.Raise 91, "modWorkbookIo.ProtectSheetAfterWrite", "対象sheetを取得できません。"
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    If Not ws.ProtectContents Then Err.Raise vbObjectError + 1102, "modWorkbookIo.ProtectSheetAfterWrite", "sheet保護を復元できません。"
End Sub

Private Function FileNameFromOperationPath(ByVal path As String) As String
    Dim p As Long
    p = InStrRev(path, "\")
    If p > 0 Then
        FileNameFromOperationPath = Mid$(path, p + 1)
    Else
        FileNameFromOperationPath = path
    End If
End Function

Private Function SheetExistsIo(ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    SheetExistsIo = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function ParentPath(ByVal path As String) As String
    Dim p As Long
    p = InStrRev(path, "\")
    If p > 0 Then ParentPath = Left$(path, p - 1)
End Function

Private Function RelativeOperationPath(ByVal rootPath As String, ByVal fullPath As String) As String
    If Len(fullPath) <= Len(rootPath) Then
        RelativeOperationPath = "."
    Else
        RelativeOperationPath = Mid$(fullPath, Len(rootPath) + 2)
    End If
End Function
