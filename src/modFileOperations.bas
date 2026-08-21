Attribute VB_Name = "modFileOperations"
Option Explicit

Private Const MOVEFILE_WRITE_THROUGH As Long = &H8
Private Const ERROR_ALREADY_EXISTS As Long = 183

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
    Dim i As Long, errNo As Long, errText As String, reversibleDone As Long
    Dim moveSource As String
    Dim deleteStarted As Boolean, rollbackOk As Boolean
    Dim phaseResult As String

    ' 1. 新規フォルダ（浅い順）
    For i = 1 To itemCount
        If items(i).Kind = OP_MKDIR Then
            If Not CreateFolderSafe(items(i).Destination, errNo, errText) Then
                MarkOperationPlanRow items(i).Sequence, "失敗", errText, items(i).Destination
                AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, "", items(i).Destination, "失敗", errNo, errText, "", "", "通常", rootPath
                rollbackOk = RollbackReversible(items, itemCount, batchId, rootPath)
                ExecuteFileOperations = "failed-before-delete-" & IIf(rollbackOk, "rolled-back", "rollback-incomplete")
                Exit Function
            End If
            items(i).Destination = items(i).Destination
            MarkOperationPlanRow items(i).Sequence, "成功", "新規フォルダを作成しました。", items(i).Destination
            AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, "", items(i).Destination, "成功", 0, "", "", "", "通常", rootPath
            reversibleDone = i
        End If
    Next i

    ' 2a. 交換・case-only renameは予約一時名へ退避する。
    For i = 1 To itemCount
        If items(i).Kind = OP_RENAME_MOVE And Len(items(i).TemporarySource) > 0 Then
            AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, items(i).Source, items(i).Destination, "実行中", 0, "一時退避前journal", "", "一時退避予定", "通常", rootPath, TemporaryFileOperationName(items(i).TemporarySource), items(i).TemporarySource, items(i).TemporaryNonce, items(i).OriginalName
            If Not MoveFileSafe(items(i).Source, items(i).TemporarySource, errNo, errText) Then
                MarkOperationPlanRow items(i).Sequence, "失敗", errText, items(i).Destination
                AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, items(i).Source, items(i).Destination, "失敗", errNo, errText, "", items(i).TemporarySource, "通常", rootPath, TemporaryFileOperationName(items(i).TemporarySource), items(i).TemporarySource, items(i).TemporaryNonce, items(i).OriginalName
                rollbackOk = RollbackReversible(items, itemCount, batchId, rootPath)
                MarkUnexecuted items, i + 1, itemCount
                ExecuteFileOperations = "failed-before-delete-" & IIf(rollbackOk, "rolled-back", "rollback-incomplete")
                Exit Function
            End If
            reversibleDone = i
        End If
    Next i

    ' 2b. 名前変更／移動（ファイルだけ。衝突時は上書きしない）
    For i = 1 To itemCount
        If items(i).Kind = OP_RENAME_MOVE Then
            moveSource = items(i).Source
            If Len(items(i).TemporarySource) > 0 Then moveSource = items(i).TemporarySource
            If Not MoveFileSafe(moveSource, items(i).Destination, errNo, errText) Then
                MarkOperationPlanRow items(i).Sequence, "失敗", errText, items(i).Destination
                AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, items(i).Source, items(i).Destination, "失敗", errNo, errText, "", "", "通常", rootPath, TemporaryFileOperationName(items(i).TemporarySource), items(i).TemporarySource, items(i).TemporaryNonce, items(i).OriginalName
                rollbackOk = RollbackReversible(items, itemCount, batchId, rootPath)
                MarkUnexecuted items, i + 1, itemCount
                ExecuteFileOperations = "failed-before-delete-" & IIf(rollbackOk, "rolled-back", "rollback-incomplete")
                Exit Function
            End If
            MarkOperationPlanRow items(i).Sequence, "成功", "名前変更／移動しました。", items(i).Destination
            AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, items(i).Source, items(i).Destination, "成功", 0, "", "", "", "通常", rootPath, TemporaryFileOperationName(items(i).TemporarySource), items(i).TemporarySource, items(i).TemporaryNonce, items(i).OriginalName
            reversibleDone = i
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
End Function

Private Function ExecuteRecyclePhaseForKind(ByRef items() As OperationPlanItem, ByVal itemCount As Long, _
                                            ByVal operationKind As String, ByVal batchId As String, ByVal rootPath As String, _
                                            ByRef deleteStarted As Boolean, ByRef resultText As String) As Boolean
    Dim maxDepth As Long, currentDepth As Long, i As Long, k As Long, groupCount As Long
    Dim groupRows() As Long, sources() As String, sizes() As Double, modified() As Double, folders() As Boolean
    Dim states() As String, errNo As Long, errText As String, safetyState As String, apiStarted As Boolean
    Dim rollbackInfo As String, snapshotInfo As String
    Dim isFolderPhase As Boolean

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
                    If isFolderPhase And Not IsOperationFolderEmpty(items(i).Source) Then
                        MarkOperationPlanRow items(i).Sequence, "失敗", "直前の再確認で空ではありません。", "[ごみ箱]"
                        AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, items(i).Source, "[ごみ箱]", "失敗", 1, "直前の空判定に失敗", "不可", "削除開始後のため自動復元なし", "通常", rootPath
                        MarkUnexecutedRecycleKind items, itemCount, operationKind, rootPath, currentDepth, batchId
                        resultText = "partial-after-delete"
                        Exit Function
                    End If
                End If
            End If
        Next i

        If Not RecycleItemsPhase(sources, sizes, modified, folders, groupCount, states, errNo, errText, safetyState, apiStarted) Then
            If apiStarted Then deleteStarted = True
            snapshotInfo = GetRecycleSnapshotInfo()
            rollbackInfo = IIf(apiStarted, "削除開始後のため自動復元なし", "API呼出し前の拒否")
            If Len(snapshotInfo) > 0 Then rollbackInfo = rollbackInfo & ";" & snapshotInfo
            For k = 1 To groupCount
                If states(k) = "成功" Then
                    MarkOperationPlanRow items(groupRows(k)).Sequence, "成功", "ゴミ箱へ移動しました。", "[ごみ箱]"
                    AppendExecutionLog batchId, items(groupRows(k)).Sequence, items(groupRows(k)).Kind, items(groupRows(k)).Source, "[ごみ箱]", "成功", 0, "", "不可", "ゴミ箱からの自動復元なし;" & snapshotInfo, "通常", rootPath
                Else
                    MarkOperationPlanRow items(groupRows(k)).Sequence, states(k), errText, "[ごみ箱]"
                    AppendExecutionLog batchId, items(groupRows(k)).Sequence, items(groupRows(k)).Kind, items(groupRows(k)).Source, "[ごみ箱]", states(k), errNo, errText, "不可", rollbackInfo, IIf(Len(safetyState) > 0, safetyState, states(k)), rootPath
                End If
            Next k
            MarkUnexecutedRecycleKind items, itemCount, operationKind, rootPath, currentDepth - 1, batchId
            resultText = ClassifyDeleteFailureResult(apiStarted, safetyState)
            Exit Function
        End If
        snapshotInfo = GetRecycleSnapshotInfo()
        For k = 1 To groupCount
            MarkOperationPlanRow items(groupRows(k)).Sequence, "成功", IIf(isFolderPhase, "空フォルダをゴミ箱へ移動しました。", "ゴミ箱へ移動しました。"), "[ごみ箱]"
            AppendExecutionLog batchId, items(groupRows(k)).Sequence, items(groupRows(k)).Kind, items(groupRows(k)).Source, "[ごみ箱]", "成功", 0, "", "不可", "ゴミ箱からの自動復元なし;" & snapshotInfo, "通常", rootPath
        Next k
NextRecycleDepth:
    Next currentDepth
    ExecuteRecyclePhaseForKind = True
End Function

#If TEST_BUILD Then
Public Function ClassifyDeleteFailureResultTest(ByVal apiStarted As Boolean, ByVal safetyState As String) As String
    ClassifyDeleteFailureResultTest = ClassifyDeleteFailureResult(apiStarted, safetyState)
End Function
#End If

Private Function ClassifyDeleteFailureResult(ByVal apiStarted As Boolean, ByVal safetyState As String) As String
    If Len(safetyState) > 0 And safetyState <> "通常" Then
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
        If items(i).Kind = operationKind Then
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

Private Function CreateFolderSafe(ByVal path As String, ByRef errorNumber As Long, ByRef errorMessage As String) As Boolean
    Dim parent As String, attrs As Long, ok As Long, apiPath As String
    errorNumber = 0: errorMessage = ""
    If GetOperationAttributes(path) <> -1 Then errorNumber = ERROR_ALREADY_EXISTS: errorMessage = "作成先が既に存在します。": Exit Function
    parent = ParentFileOperationPath(path)
    If Len(parent) > 0 And GetOperationAttributes(parent) = -1 Then
        If Not CreateFolderSafe(parent, errorNumber, errorMessage) Then Exit Function
    End If
    On Error GoTo Fail
    apiPath = ToExtendedOperationPath(path)
    ok = CreateDirectoryW(StrPtr(apiPath), 0)
    If ok = 0 Then errorNumber = GetLastError(): errorMessage = "CreateDirectoryWに失敗しました。": Exit Function
    CreateFolderSafe = True
    Exit Function
Fail:
    errorNumber = Err.Number: errorMessage = Err.Description
End Function

Private Function MoveFileSafe(ByVal source As String, ByVal destination As String, ByRef errorNumber As Long, ByRef errorMessage As String) As Boolean
    Dim ok As Long, apiSource As String, apiDestination As String
    errorNumber = 0: errorMessage = ""
    If GetOperationAttributes(destination) <> -1 Then errorNumber = 80: errorMessage = "destinationが既に存在します。": Exit Function
    On Error GoTo Fail
    apiSource = ToExtendedOperationPath(source)
    apiDestination = ToExtendedOperationPath(destination)
    ok = MoveFileExW(StrPtr(apiSource), StrPtr(apiDestination), MOVEFILE_WRITE_THROUGH)
    If ok = 0 Then errorNumber = GetLastError(): errorMessage = "MoveFileExWに失敗しました。": Exit Function
    If GetOperationAttributes(source) <> -1 Or GetOperationAttributes(destination) = -1 Then errorNumber = 1: errorMessage = "移動後のsource/destination照合に失敗しました。": Exit Function
    MoveFileSafe = True
    Exit Function
Fail:
    errorNumber = Err.Number: errorMessage = Err.Description
End Function

Private Function RollbackReversible(ByRef items() As OperationPlanItem, ByVal lastIndex As Long, ByVal batchId As String, ByVal rootPath As String) As Boolean
    Dim i As Long, errNo As Long, errText As String, ok As Boolean
    ok = True
    For i = lastIndex To 1 Step -1
        If items(i).Kind = OP_RENAME_MOVE Then
            If Len(items(i).TemporarySource) > 0 And GetOperationAttributes(items(i).TemporarySource) <> -1 And GetOperationAttributes(items(i).Source) = -1 Then
                If Not MoveFileSafe(items(i).TemporarySource, items(i).Source, errNo, errText) Then
                    ok = False
                    AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, items(i).Destination, items(i).Source, "未実行", errNo, errText, "失敗", "手動復旧が必要", "通常", rootPath, TemporaryFileOperationName(items(i).TemporarySource), items(i).TemporarySource, items(i).TemporaryNonce, items(i).OriginalName
                Else
                    AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, items(i).Destination, items(i).Source, "未実行", 0, "", "成功", "逆操作を実行", "通常", rootPath, TemporaryFileOperationName(items(i).TemporarySource), items(i).TemporarySource, items(i).TemporaryNonce, items(i).OriginalName
                End If
            ElseIf GetOperationAttributes(items(i).Destination) <> -1 And GetOperationAttributes(items(i).Source) = -1 Then
                If Not MoveFileSafe(items(i).Destination, items(i).Source, errNo, errText) Then
                    ok = False
                    AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, items(i).Destination, items(i).Source, "未実行", errNo, errText, "失敗", "手動復旧が必要", "通常", rootPath, TemporaryFileOperationName(items(i).TemporarySource), items(i).TemporarySource, items(i).TemporaryNonce, items(i).OriginalName
                Else
                    AppendExecutionLog batchId, items(i).Sequence, items(i).Kind, items(i).Destination, items(i).Source, "未実行", 0, "", "成功", "逆操作を実行", "通常", rootPath, TemporaryFileOperationName(items(i).TemporarySource), items(i).TemporarySource, items(i).TemporaryNonce, items(i).OriginalName
                End If
            End If
        ElseIf items(i).Kind = OP_MKDIR Then
            If GetOperationAttributes(items(i).Destination) <> -1 And IsOperationFolderEmpty(items(i).Destination) Then
                On Error Resume Next
                If RemoveDirectoryW(StrPtr(items(i).Destination)) = 0 Then ok = False
                On Error GoTo 0
            End If
        End If
    Next i
    RollbackReversible = ok
End Function

Private Sub MarkUnexecuted(ByRef items() As OperationPlanItem, ByVal firstIndex As Long, ByVal lastIndex As Long)
    Dim i As Long
    If firstIndex > lastIndex Then Exit Sub
    For i = firstIndex To lastIndex
        If i >= LBound(items) And i <= UBound(items) Then
            If Len(items(i).Kind) > 0 Then
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
