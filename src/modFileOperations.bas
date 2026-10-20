Attribute VB_Name = "modFileOperations"
Option Explicit

Private Const MOVEFILE_WRITE_THROUGH As Long = &H8
Private Const ERROR_ALREADY_EXISTS As Long = 183

Private Type UndoEntry
    Kind As String
    Source As String
    Destination As String
    Identity As String
    Attributes As Long
    Size As Double
    Modified As Double
    Sequence As Long
End Type

Private mUndo() As UndoEntry
Private mUndoCount As Long
Private mUndoCapacity As Long

#If VBA7 Then
    Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal existingName As LongPtr, ByVal newName As LongPtr, ByVal flags As Long) As Long
    Private Declare PtrSafe Function CreateDirectoryW Lib "kernel32" (ByVal pathName As LongPtr, ByVal securityAttributes As LongPtr) As Long
    Private Declare PtrSafe Function RemoveDirectoryW Lib "kernel32" (ByVal pathName As LongPtr) As Long
    Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal virtualKey As Long) As Integer
#Else
    Private Declare Function MoveFileExW Lib "kernel32" (ByVal existingName As Long, ByVal newName As Long, ByVal flags As Long) As Long
    Private Declare Function CreateDirectoryW Lib "kernel32" (ByVal pathName As Long, ByVal securityAttributes As Long) As Long
    Private Declare Function RemoveDirectoryW Lib "kernel32" (ByVal pathName As Long) As Long
    Private Declare Function GetLastError Lib "kernel32" () As Long
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal virtualKey As Long) As Integer
#End If

Public Function ExecuteFileOperations(ByRef items() As OperationPlanItem, ByVal itemCount As Long, _
                                      ByVal batchId As String, ByVal rootPath As String) As String
    Dim i As Long, errNo As Long, errText As String
    Dim moveSource As String
    Dim deleteStarted As Boolean, rollbackOk As Boolean
    Dim phaseResult As String, savedNumber As Long, savedDescription As String
    Erase mUndo
    mUndoCount = 0
    mUndoCapacity = 0
    On Error GoTo UnexpectedFailure

    ' 1. 新規フォルダ（浅い順）
    For i = 1 To itemCount
        If items(i).Kind = OP_MKDIR Then
            If Not CreateFolderSafe(items(i).Destination, items(i).Sequence, errNo, errText) Then
                items(i).ExecutionState = "失敗"
                MarkOperationPlanRow items(i).Sequence, "失敗", errText, items(i).Destination
                AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, "", items(i).Destination, "失敗", errNo, errText, "", "", "通常", rootPath
                rollbackOk = RollbackReversible(items, itemCount, batchId, rootPath)
                MarkUnexecuted items, 1, itemCount
                ExecuteFileOperations = "failed-before-delete-" & IIf(rollbackOk, "rolled-back", "rollback-incomplete")
                Exit Function
            End If
            items(i).ExecutionState = "成功"
            MarkOperationPlanRow items(i).Sequence, "成功", "新規フォルダを作成しました。", items(i).Destination
            AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, "", items(i).Destination, "成功", 0, "", "", "", "通常", rootPath
        End If
    Next i

    ' 2a. 交換・case-only renameは予約一時名へ退避する。
    For i = 1 To itemCount
        If items(i).Kind = OP_RENAME_MOVE And Len(items(i).TemporarySource) > 0 Then
            AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, items(i).Source, items(i).Destination, "実行中", 0, "一時退避前journal", "", "一時退避予定", "通常", rootPath, TemporaryFileOperationName(items(i).TemporarySource), items(i).TemporarySource, items(i).TemporaryNonce, items(i).OriginalName
            If Not MovePlannedFile(items(i), items(i).Source, items(i).TemporarySource, errNo, errText) Then
                items(i).ExecutionState = "失敗"
                MarkOperationPlanRow items(i).Sequence, "失敗", errText, items(i).Destination
                AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, items(i).Source, items(i).Destination, "失敗", errNo, errText, "", items(i).TemporarySource, "通常", rootPath, TemporaryFileOperationName(items(i).TemporarySource), items(i).TemporarySource, items(i).TemporaryNonce, items(i).OriginalName
                rollbackOk = RollbackReversible(items, itemCount, batchId, rootPath)
                MarkUnexecuted items, 1, itemCount
                ExecuteFileOperations = "failed-before-delete-" & IIf(rollbackOk, "rolled-back", "rollback-incomplete")
                Exit Function
            End If
            If items(i).ExecutionState = "未実行" Then items(i).ExecutionState = "実行中"
        End If
    Next i

    ' 2b. 名前変更／移動（ファイルだけ。衝突時は上書きしない）
    For i = 1 To itemCount
        If items(i).Kind = OP_RENAME_MOVE Then
            moveSource = items(i).Source
            If Len(items(i).TemporarySource) > 0 Then moveSource = items(i).TemporarySource
            If Not MovePlannedFile(items(i), moveSource, items(i).Destination, errNo, errText) Then
                items(i).ExecutionState = "失敗"
                MarkOperationPlanRow items(i).Sequence, "失敗", errText, items(i).Destination
                AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, items(i).Source, items(i).Destination, "失敗", errNo, errText, "", "", "通常", rootPath, TemporaryFileOperationName(items(i).TemporarySource), items(i).TemporarySource, items(i).TemporaryNonce, items(i).OriginalName
                rollbackOk = RollbackReversible(items, itemCount, batchId, rootPath)
                MarkUnexecuted items, 1, itemCount
                ExecuteFileOperations = "failed-before-delete-" & IIf(rollbackOk, "rolled-back", "rollback-incomplete")
                Exit Function
            End If
            items(i).ExecutionState = "成功"
            MarkOperationPlanRow items(i).Sequence, "成功", "名前変更／移動しました。", items(i).Destination
            AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, items(i).Source, items(i).Destination, "成功", 0, "", "", "", "通常", rootPath, TemporaryFileOperationName(items(i).TemporarySource), items(i).TemporarySource, items(i).TemporaryNonce, items(i).OriginalName
        End If
    Next i

    ' 3. ファイル削除phaseは一括、空フォルダは同一深さgroup単位。
    If Not ExecuteRecyclePhaseForKind(items, itemCount, OP_RECYCLE_FILE, batchId, rootPath, deleteStarted, phaseResult) Then
        If Not deleteStarted Then
            rollbackOk = RollbackReversible(items, itemCount, batchId, rootPath)
            ExecuteFileOperations = "failed-before-delete-" & IIf(rollbackOk, "rolled-back", "rollback-incomplete")
        Else
            ExecuteFileOperations = phaseResult
        End If
        Exit Function
    End If
    If Not ExecuteRecyclePhaseForKind(items, itemCount, OP_RECYCLE_FOLDER, batchId, rootPath, deleteStarted, phaseResult) Then
        If Not deleteStarted Then
            rollbackOk = RollbackReversible(items, itemCount, batchId, rootPath)
            ExecuteFileOperations = "failed-before-delete-" & IIf(rollbackOk, "rolled-back", "rollback-incomplete")
        Else
            ExecuteFileOperations = phaseResult
        End If
        Exit Function
    End If

    ExecuteFileOperations = "success: " & itemCount & " 件の操作が完了しました。"
    Exit Function
UnexpectedFailure:
    savedNumber = Err.Number: savedDescription = Err.Description
    ' Excel I/O errors must not bypass the same rollback/delete boundary.
    On Error Resume Next
    If Not deleteStarted Then
        rollbackOk = RollbackReversible(items, itemCount, batchId, rootPath)
        ExecuteFileOperations = "failed-before-delete-" & IIf(rollbackOk, "rolled-back", "rollback-incomplete")
    Else
        ExecuteFileOperations = "partial-after-delete"
    End If
    ExecuteFileOperations = ExecuteFileOperations & ": " & CStr(savedNumber) & " " & savedDescription
    On Error GoTo 0
End Function

Private Function ExecuteRecyclePhaseForKind(ByRef items() As OperationPlanItem, ByVal itemCount As Long, _
                                            ByVal operationKind As String, ByVal batchId As String, ByVal rootPath As String, _
                                            ByRef deleteStarted As Boolean, ByRef resultText As String) As Boolean
    Dim maxDepth As Long, currentDepth As Long, i As Long, k As Long, groupCount As Long
    Dim groupRows() As Long, sources() As String, sizes() As Double, modified() As Double, folders() As Boolean
    Dim states() As String, errNo As Long, errText As String, safetyState As String, apiStarted As Boolean
    Dim rollbackInfo As String, snapshotInfo As String
    Dim isFolderPhase As Boolean, phaseOk As Boolean, previousCancel As Long

    resultText = "partial-after-delete"
    isFolderPhase = (operationKind = OP_RECYCLE_FOLDER)
    If isFolderPhase Then
        For i = 1 To itemCount
            If items(i).Kind = operationKind Then
                If OperationDepth(rootPath, items(i).Source) > maxDepth Then maxDepth = OperationDepth(rootPath, items(i).Source)
            End If
        Next i
    End If

    If Not isFolderPhase Then maxDepth = 0
    For currentDepth = maxDepth To 0 Step -1
        If isFolderPhase And IsEscapePressed() Then
            MarkUnexecutedRecycleKind items, itemCount, operationKind, rootPath, currentDepth, batchId
            resultText = "partial-after-delete"
            Exit Function
        End If
        groupCount = 0
        For i = 1 To itemCount
            If items(i).Kind = operationKind Then
                If (Not isFolderPhase) Or OperationDepth(rootPath, items(i).Source) = currentDepth Then groupCount = groupCount + 1
            End If
        Next i
        If groupCount = 0 Then
            If Not isFolderPhase Then Exit For
            GoTo NextRecycleDepth
        End If

        ReDim groupRows(1 To groupCount)
        ReDim sources(1 To groupCount)
        ReDim sizes(1 To groupCount)
        ReDim modified(1 To groupCount)
        ReDim folders(1 To groupCount)
        k = 0
        For i = 1 To itemCount
            If items(i).Kind = operationKind Then
                If (Not isFolderPhase) Or OperationDepth(rootPath, items(i).Source) = currentDepth Then
                    k = k + 1
                    groupRows(k) = i
                    sources(k) = items(i).Source
                    sizes(k) = items(i).SourceSize
                    modified(k) = items(i).SourceModified
                    folders(k) = items(i).IsFolder
                    If Not VerifyOperationBoundary(items(i).Source, errText) Then GoTo SourceChanged
                    If GetOperationIdentity(items(i).Source) <> items(i).SourceIdentity Then
                        errText = "削除対象のidentityが変化しました。": GoTo SourceChanged
                    End If
                    If Not VerifyOperationSnapshot(items(i).Source, items(i).SourceAttributes, items(i).SourceSize, items(i).SourceModified, errText) Then GoTo SourceChanged
                    If isFolderPhase And Not IsOperationFolderEmpty(items(i).Source) Then
                        items(i).ExecutionState = "失敗"
                        MarkOperationPlanRow items(i).Sequence, "失敗", "直前の再確認で空ではありません。", "[ごみ箱]"
                        AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, items(i).Source, "[ごみ箱]", "失敗", 1, "直前の空判定に失敗", "不可", "削除開始後のため自動復元なし", "通常", rootPath
                        MarkUnexecutedRecycleKind items, itemCount, operationKind, rootPath, currentDepth, batchId
                        resultText = "partial-after-delete"
                        Exit Function
                    End If
                End If
            End If
        Next i

        previousCancel = Application.EnableCancelKey
        Application.EnableCancelKey = xlDisabled
        phaseOk = RecycleItemsPhase(sources, sizes, modified, folders, groupCount, states, errNo, errText, safetyState, apiStarted, rootPath)
        If apiStarted Then deleteStarted = True
        Application.EnableCancelKey = previousCancel
        If Not phaseOk Then
            snapshotInfo = GetRecycleSnapshotInfo()
            rollbackInfo = IIf(deleteStarted, "削除開始後のため自動復元なし", "API呼出し前の拒否")
            If Len(snapshotInfo) > 0 Then rollbackInfo = rollbackInfo & ";" & snapshotInfo
            For k = 1 To groupCount
                items(groupRows(k)).ExecutionState = states(k)
                If states(k) = "成功" Then
                    MarkOperationPlanRow items(groupRows(k)).Sequence, "成功", "ゴミ箱へ移動しました。", "[ごみ箱]"
                    AppendExecutionLog batchId, items(groupRows(k)).Sequence, items(groupRows(k)).Kind, items(groupRows(k)).Source, "[ごみ箱]", "成功", 0, "", "不可", "ゴミ箱からの自動復元なし;" & snapshotInfo, "通常", rootPath
                Else
                    MarkOperationPlanRow items(groupRows(k)).Sequence, states(k), errText, "[ごみ箱]"
                    AppendExecutionLog batchId, items(groupRows(k)).Sequence, items(groupRows(k)).Kind, items(groupRows(k)).Source, "[ごみ箱]", states(k), errNo, errText, "不可", rollbackInfo, IIf(Len(safetyState) > 0, safetyState, states(k)), rootPath
                End If
            Next k
            If isFolderPhase Then MarkUnexecutedRecycleKind items, itemCount, operationKind, rootPath, currentDepth - 1, batchId
            If Not isFolderPhase Then MarkUnexecutedRecycleKind items, itemCount, OP_RECYCLE_FOLDER, rootPath, 2147483647, batchId
            resultText = ClassifyDeleteFailureResult(deleteStarted, safetyState)
            Exit Function
        End If
        snapshotInfo = GetRecycleSnapshotInfo()
        For k = 1 To groupCount
            items(groupRows(k)).ExecutionState = "成功"
            MarkOperationPlanRow items(groupRows(k)).Sequence, "成功", IIf(isFolderPhase, "空フォルダをゴミ箱へ移動しました。", "ゴミ箱へ移動しました。"), "[ごみ箱]"
            AppendExecutionLog batchId, items(groupRows(k)).Sequence, items(groupRows(k)).Kind, items(groupRows(k)).Source, "[ごみ箱]", "成功", 0, "", "不可", "ゴミ箱からの自動復元なし;" & snapshotInfo, "通常", rootPath
        Next k
NextRecycleDepth:
    Next currentDepth
    ExecuteRecyclePhaseForKind = True
    Exit Function
SourceChanged:
    items(i).ExecutionState = "失敗"
    MarkOperationPlanRow items(i).Sequence, "失敗", errText, "[ごみ箱]"
    AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, items(i).Source, "[ごみ箱]", "失敗", 0, errText, "", "API呼出し前の拒否", "通常", rootPath
    MarkUnexecutedRecycleKind items, itemCount, operationKind, rootPath, currentDepth, batchId
    If Not isFolderPhase Then MarkUnexecutedRecycleKind items, itemCount, OP_RECYCLE_FOLDER, rootPath, 2147483647, batchId
    resultText = ClassifyDeleteFailureResult(deleteStarted, "通常")
End Function

#If TEST_BUILD Then
Public Sub MarkUnexecutedRecycleTest()
    Dim items() As OperationPlanItem, i As Long, count As Long
    count = GetOperationPlanCount()
    If count <> 3 Then Err.Raise 5, "MarkUnexecutedRecycleTest", "Exactly three planned file deletions required."
    ReDim items(1 To count)
    For i = 1 To count
        items(i) = GetOperationPlanItem(i)
        If items(i).Kind <> OP_RECYCLE_FILE Then Err.Raise 5
    Next i
    items(1).ExecutionState = "成功"
    items(2).ExecutionState = "possible-permanent-delete"
    items(3).ExecutionState = "未実行"
    For i = 1 To count
        MarkOperationPlanRow items(i).Sequence, items(i).ExecutionState, "state probe", "[ごみ箱]"
    Next i
    MarkUnexecutedRecycleKind items, count, OP_RECYCLE_FILE, GetOperationPlanRoot(), 0, GetOperationPlanBatchId()
End Sub

Public Function ClassifyDeleteFailureResultTest(ByVal apiStarted As Boolean, ByVal safetyState As String) As String
    ClassifyDeleteFailureResultTest = ClassifyDeleteFailureResult(apiStarted, safetyState)
End Function
#End If

Private Function ClassifyDeleteFailureResult(ByVal apiStarted As Boolean, ByVal safetyState As String) As String
    If safetyState = "possible-permanent-delete" Or safetyState = "recycle-verification-contaminated" Then
        ClassifyDeleteFailureResult = safetyState
    ElseIf apiStarted Then
        ClassifyDeleteFailureResult = "partial-after-delete"
    Else
        ClassifyDeleteFailureResult = "failed-before-delete"
    End If
End Function

Private Sub MarkUnexecutedRecycleKind(ByRef items() As OperationPlanItem, ByVal itemCount As Long, _
                                      ByVal operationKind As String, ByVal rootPath As String, _
                                      ByVal maxDepth As Long, ByVal batchId As String)
    Dim i As Long, depth As Long
    For i = 1 To itemCount
        If items(i).Kind = operationKind And items(i).ExecutionState = "未実行" Then
            depth = OperationDepth(rootPath, items(i).Source)
            If (operationKind = OP_RECYCLE_FILE) Or depth <= maxDepth Then
                MarkOperationPlanRow items(i).Sequence, "未実行", "前phaseの中断または失敗により実行していません。", "[ごみ箱]"
                AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, items(i).Source, "[ごみ箱]", "未実行", 0, "phase境界で停止", "不可", "削除開始後のため自動復元なし", "通常", rootPath
            End If
        End If
    Next i
End Sub

Private Function IsEscapePressed() As Boolean
    IsEscapePressed = (GetAsyncKeyState(27) < 0)
End Function

Private Function OperationDepth(ByVal rootPath As String, ByVal sourcePath As String) As Long
    Dim relative As String, i As Long
    relative = sourcePath
    If Len(sourcePath) > Len(rootPath) Then relative = Mid$(sourcePath, Len(rootPath) + 2)
    If Len(relative) = 0 Or relative = "." Then Exit Function
    OperationDepth = 1
    For i = 1 To Len(relative)
        If Mid$(relative, i, 1) = "\" Then OperationDepth = OperationDepth + 1
    Next i
End Function

Private Sub ReserveUndo(ByVal kind As String, ByVal source As String, ByVal destination As String, _
                        ByVal sequence As Long, ByVal identity As String)
    ' Allocate BEFORE mutation; increment only after a successful Win32 call.
    If mUndoCount = mUndoCapacity Then
        mUndoCapacity = mUndoCapacity + 1024
        ReDim Preserve mUndo(1 To mUndoCapacity)
    End If
    With mUndo(mUndoCount + 1)
        .Kind = kind: .Source = source: .Destination = destination
        .Sequence = sequence: .Identity = identity
    End With
End Sub

Private Function CreateFolderSafe(ByVal path As String, ByVal sequence As Long, ByRef errorNumber As Long, ByRef errorMessage As String) As Boolean
    Dim parent As String, apiPath As String
    errorNumber = 0: errorMessage = ""
    If Not VerifyOperationBoundary(path, errorMessage) Then Exit Function
    If GetOperationAttributes(path) <> -1 Then errorNumber = ERROR_ALREADY_EXISTS: errorMessage = "作成先が既に存在します。": Exit Function
    parent = ParentFileOperationPath(path)
    If Len(parent) > 0 And GetOperationAttributes(parent) = -1 Then
        If Not CreateFolderSafe(parent, sequence, errorNumber, errorMessage) Then Exit Function
    End If
    If Not VerifyOperationBoundary(path, errorMessage) Then Exit Function
    ReserveUndo OP_MKDIR, "", path, sequence, ""
    apiPath = ToExtendedOperationPath(path)
    If CreateDirectoryW(StrPtr(apiPath), 0) = 0 Then errorNumber = GetLastError(): errorMessage = "CreateDirectoryWに失敗しました。": Exit Function
    mUndoCount = mUndoCount + 1
    mUndo(mUndoCount).Identity = GetOperationIdentity(path)
    If Len(mUndo(mUndoCount).Identity) = 0 Then errorMessage = "作成したフォルダを識別できません。": Exit Function
    CreateFolderSafe = True
End Function

Private Function MovePlannedFile(ByRef item As OperationPlanItem, ByVal source As String, ByVal destination As String, _
                                 ByRef errorNumber As Long, ByRef errorMessage As String) As Boolean
    If Not VerifyOperationBoundary(source, errorMessage) Then Exit Function
    If GetOperationIdentity(source) <> item.SourceIdentity Then errorMessage = "sourceのfile identityが変化しました。": Exit Function
    If Not VerifyOperationSnapshot(source, item.SourceAttributes, item.SourceSize, item.SourceModified, errorMessage) Then Exit Function
    ReserveUndo OP_RENAME_MOVE, source, destination, item.Sequence, item.SourceIdentity
    With mUndo(mUndoCount + 1)
        .Attributes = item.SourceAttributes: .Size = item.SourceSize: .Modified = item.SourceModified
    End With
    MovePlannedFile = MoveFileSafe(source, destination, errorNumber, errorMessage, True)
End Function

Private Function MoveFileSafe(ByVal source As String, ByVal destination As String, ByRef errorNumber As Long, _
                              ByRef errorMessage As String, Optional ByVal recordUndo As Boolean = False) As Boolean
    Dim apiSource As String, apiDestination As String
    errorNumber = 0: errorMessage = ""
    If Not VerifyOperationBoundary(source, errorMessage) Then Exit Function
    If Not VerifyOperationBoundary(destination, errorMessage) Then Exit Function
    If GetOperationAttributes(destination) <> -1 Then errorNumber = 80: errorMessage = "destinationが既に存在します。": Exit Function
    apiSource = ToExtendedOperationPath(source)
    apiDestination = ToExtendedOperationPath(destination)
    If MoveFileExW(StrPtr(apiSource), StrPtr(apiDestination), MOVEFILE_WRITE_THROUGH) = 0 Then
        errorNumber = GetLastError(): errorMessage = "MoveFileExWに失敗しました。": Exit Function
    End If
    If recordUndo Then mUndoCount = mUndoCount + 1
    If GetOperationAttributes(source) <> -1 Or GetOperationAttributes(destination) = -1 Then errorNumber = 1: errorMessage = "移動後のsource/destination照合に失敗しました。": Exit Function
    MoveFileSafe = True
End Function

Private Function RollbackReversible(ByRef items() As OperationPlanItem, ByVal lastIndex As Long, ByVal batchId As String, ByVal rootPath As String) As Boolean
    Dim i As Long, errNo As Long, errText As String, apiPath As String, restored As Boolean, allRestored As Boolean
    allRestored = True
    On Error GoTo Failed
    For i = mUndoCount To 1 Step -1
        restored = False: errNo = 0: errText = ""
        With mUndo(i)
            If Not VerifyOperationBoundary(.Destination, errText) Then GoTo UndoResult
            If Len(.Identity) = 0 Or GetOperationIdentity(.Destination) <> .Identity Then
                errText = "復元対象のidentityが一致しません。": GoTo UndoResult
            End If
            If .Kind = OP_RENAME_MOVE Then
                If Not VerifyOperationSnapshot(.Destination, .Attributes, .Size, .Modified, errText) Then GoTo UndoResult
                restored = MoveFileSafe(.Destination, .Source, errNo, errText)
            ElseIf .Kind = OP_MKDIR Then
                If IsOperationFolderEmpty(.Destination) Then
                    apiPath = ToExtendedOperationPath(.Destination)
                    restored = (RemoveDirectoryW(StrPtr(apiPath)) <> 0)
                    If Not restored Then errNo = GetLastError()
                Else
                    errText = "作成フォルダが空ではありません。"
                End If
            End If
UndoResult:
            If Not restored Then allRestored = False
            AppendExecutionLog batchId, .Sequence, .Kind, .Destination, .Source, "rollback", errNo, errText, _
                               IIf(restored, "成功", "失敗"), "実行済みjournalの逆操作", "通常", rootPath
        End With
    Next i
    For i = 1 To lastIndex
        If items(i).ExecutionState = "成功" Or items(i).ExecutionState = "実行中" Then
            items(i).ExecutionState = IIf(allRestored, "復元済み", "復元要確認")
            MarkOperationPlanRow items(i).Sequence, items(i).ExecutionState, "実行ログのrollback結果を確認してください。", items(i).Destination
        End If
    Next i
    RollbackReversible = allRestored
    Exit Function
Failed:
    RollbackReversible = False
End Function

Private Sub MarkUnexecuted(ByRef items() As OperationPlanItem, ByVal firstIndex As Long, ByVal lastIndex As Long)
    Dim i As Long
    If firstIndex > lastIndex Then Exit Sub
    For i = firstIndex To lastIndex
        If i >= LBound(items) And i <= UBound(items) Then
            If Len(items(i).Kind) > 0 And items(i).ExecutionState = "未実行" Then
                MarkOperationPlanRow items(i).Sequence, "未実行", "前項目の失敗により停止しました。", items(i).Destination
            End If
        End If
    Next i
End Sub

Private Function ParentFileOperationPath(ByVal path As String) As String
    Dim p As Long
    p = InStrRev(path, "\")
    If p > 0 Then ParentFileOperationPath = Left$(path, p - 1)
End Function

Private Function TemporaryFileOperationName(ByVal path As String) As String
    Dim p As Long
    p = InStrRev(path, "\")
    If p > 0 Then
        TemporaryFileOperationName = Mid$(path, p + 1)
    Else
        TemporaryFileOperationName = path
    End If
End Function
