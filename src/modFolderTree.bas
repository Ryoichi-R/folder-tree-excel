Attribute VB_Name = "modFolderTree"
Option Explicit

'==============================================================
' フォルダ階層ツリー生成ツール
'
'   [設定] シートで対象フォルダとオプションを指定し、
'   [2. ツリーを作成] ボタンで [ツリー] シートへ階層を書き出す。
'   階層は Excel のアウトライン機能（左端の + / - ）で折りたためる。
'
'   実装方針:
'     - 列挙は Win32 の FindFirstFileW / FindNextFileW を直接使う。
'       FileSystemObject は 1 ファイルにつき Attributes / Size / Name /
'       Path / DateLastModified をそれぞれ別の COM 呼び出しで取得するため、
'       クラウド同期ドライブ（Google ドライブ等）ではドライバへの往復が
'       膨大になりストールする。Win32 API なら 1 エントリ 1 回の呼び出しで
'       属性・サイズ・更新日時をまとめて取得できる。
'     - 走査は再帰ではなく明示スタックによる反復処理。VBA のスタック領域に
'       依存しないため「実行時エラー 28: スタック領域が不足しています」が起きない。
'     - 長いパス（260 文字超）は \\?\ プレフィックスで扱う。
'     - Esc キーで中断でき、その時点までの結果を書き出す。
'
'   参照設定は不要。
'==============================================================

'--- Win32 API -----------------------------------------------
Private Const MAX_PATH_CHARS      As Long = 260
Private Const INVALID_FILE_ATTRS  As Long = -1
Private Const LONG_PATH_THRESHOLD As Long = 230
Private Const ERROR_NO_MORE_FILES As Long = 18
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_FILE_NOT_FOUND As Long = 2

Private Const FILE_ATTRIBUTE_DIRECTORY     As Long = &H10
Private Const FILE_ATTRIBUTE_HIDDEN        As Long = &H2
Private Const FILE_ATTRIBUTE_SYSTEM        As Long = &H4
Private Const FILE_ATTRIBUTE_REPARSE_POINT As Long = &H400

Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear         As Integer
    wMonth        As Integer
    wDayOfWeek    As Integer
    wDay          As Integer
    wHour         As Integer
    wMinute       As Integer
    wSecond       As Integer
    wMilliseconds As Integer
End Type

Private Type WIN32_FIND_DATAW
    dwFileAttributes   As Long
    ftCreationTime     As FILETIME
    ftLastAccessTime   As FILETIME
    ftLastWriteTime    As FILETIME
    nFileSizeHigh      As Long
    nFileSizeLow       As Long
    dwReserved0        As Long
    dwReserved1        As Long
    cFileName(0 To 259) As Integer
    cAlternate(0 To 13) As Integer
End Type

#If VBA7 Then
    Private Declare PtrSafe Function FindFirstFileW Lib "kernel32" ( _
        ByVal lpFileName As LongPtr, ByRef lpFindFileData As WIN32_FIND_DATAW) As LongPtr
    Private Declare PtrSafe Function FindNextFileW Lib "kernel32" ( _
        ByVal hFindFile As LongPtr, ByRef lpFindFileData As WIN32_FIND_DATAW) As Long
    Private Declare PtrSafe Function FindClose Lib "kernel32" ( _
        ByVal hFindFile As LongPtr) As Long
    Private Declare PtrSafe Function GetFileAttributesW Lib "kernel32" ( _
        ByVal lpFileName As LongPtr) As Long
    Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
    Private Declare PtrSafe Function FileTimeToLocalFileTime Lib "kernel32" ( _
        ByRef lpFileTime As FILETIME, ByRef lpLocalFileTime As FILETIME) As Long
    Private Declare PtrSafe Function FileTimeToSystemTime Lib "kernel32" ( _
        ByRef lpFileTime As FILETIME, ByRef lpSystemTime As SYSTEMTIME) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
#Else
    Private Declare Function FindFirstFileW Lib "kernel32" ( _
        ByVal lpFileName As Long, ByRef lpFindFileData As WIN32_FIND_DATAW) As Long
    Private Declare Function FindNextFileW Lib "kernel32" ( _
        ByVal hFindFile As Long, ByRef lpFindFileData As WIN32_FIND_DATAW) As Long
    Private Declare Function FindClose Lib "kernel32" ( _
        ByVal hFindFile As Long) As Long
    Private Declare Function GetFileAttributesW Lib "kernel32" ( _
        ByVal lpFileName As Long) As Long
    Private Declare Function GetLastError Lib "kernel32" () As Long
    Private Declare Function FileTimeToLocalFileTime Lib "kernel32" ( _
        ByRef lpFileTime As FILETIME, ByRef lpLocalFileTime As FILETIME) As Long
    Private Declare Function FileTimeToSystemTime Lib "kernel32" ( _
        ByRef lpFileTime As FILETIME, ByRef lpSystemTime As SYSTEMTIME) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If

'--- シート・セル定義 ----------------------------------------
Private Const SHEET_CONFIG As String = "設定"
Private Const SHEET_TREE   As String = "ツリー"

Private Const CELL_ROOT     As String = "C5"
Private Const CELL_MAXDEPTH As String = "C6"
Private Const CELL_FILES    As String = "C7"
Private Const CELL_HIDDEN   As String = "C8"

Private Const COL_NAME  As Long = 1
Private Const COL_TYPE  As Long = 2
Private Const COL_DEPTH As Long = 3
Private Const COL_ITEMS As Long = 4
Private Const COL_SIZE  As Long = 5
Private Const COL_DATE  As Long = 6
Private Const COL_PATH  As Long = 7
Private Const COL_MAX   As Long = 7

Private Const HEADER_ROW As Long = 3
Private Const FIRST_ROW  As Long = 4

Private Const MAX_OUTLINE_LEVEL As Long = 8
Private Const MAX_INDENT_LEVEL  As Long = 15
Private Const MAX_OUTPUT_ROWS   As Long = 1000000
Private Const MAX_SCAN_DEPTH    As Long = 512      ' 暴走構造に対する安全弁
Private Const MAX_ADDR_LEN      As Long = 200      ' Range アドレス文字列の上限（実際は 255）
Private Const OUTPUT_CHUNK_ROWS As Long = 20000

'--- 出力バッファ（「列 × 行」で保持し、書き込み直前に転置する）----
' VBA の ReDim Preserve は最終次元しか拡張できないため、行を最終次元に置く。
Private mBuf()    As Variant
Private mLevel()  As Long
Private mCount    As Long
Private mCapacity As Long

'--- 走査スタック（再帰を使わない）---------------------------
Private mStPath()  As String
Private mStDepth() As Long
Private mStRow()   As Long
Private mStSubs()  As Variant   ' 各要素はサブフォルダ名の String 配列
Private mStNum()   As Long      ' サブフォルダ数
Private mStNext()  As Long      ' 次に処理するサブフォルダの位置
Private mStSize()  As Double
Private mStFiles() As Double
Private mStAttrs() As Variant
Private mStDates() As Variant
Private mStTop     As Long
Private mStCap     As Long

'--- 実行時の状態 --------------------------------------------
Private mMaxDepth      As Long
Private mIncludeFiles  As Boolean
Private mIncludeHidden As Boolean
Private mFolderCount   As Double
Private mDeniedCount   As Long
Private mLinkCount     As Long
Private mCycleCount    As Long
Private mDepthReached  As Long
Private mTruncated     As Boolean
Private mAllocFailed   As Boolean
Private mCancelled     As Boolean
Private mTotalSize     As Double
Private mTotalFiles    As Double
Private mIoErrorCount  As Long
Private mDepthLimitHit As Boolean
Private mOutlineErrors As Long
Private mFirstOutlineError As String
Private mAllocErrorNumber As Long
Private mAllocErrorDescription As String
Private mRunning As Boolean
Private mOperationBlocksScan As Boolean
Private mPrevScreenUpdating As Boolean
Private mPrevEnableEvents As Boolean
Private mPrevCalculation As Long
Private mPrevCursor As Long
Private mPrevCancel As Long
Private mPrevStatusBar As Variant
Private mOutputLimit As Long
Private mChunkRows As Long
Private mTestRowLimit As Long
Private mTestChunkRows As Long
Private mTestFailAllocation As Boolean
Private mTestFailOutline As Boolean
Private mTestCancelAfterRows As Long
Private mTestFindFirstError As Long
Private mTestFindNextError As Long

' v1.1 operation snapshot. Display columns remain seven columns; these arrays
' retain the per-item attributes needed by the operation-plan verifier.
Private mSnapshotAttributes() As Long
Private mSnapshotReparse() As Boolean
Private mScanRoot As String
Private mScanReady As Boolean

'--------------------------------------------------------------
' 1. フォルダを選択
'--------------------------------------------------------------
Public Sub SelectRootFolder()
    Dim fd As Object
    Dim cfg As Worksheet
    Dim current As String

    Set cfg = ThisWorkbook.Worksheets(SHEET_CONFIG)
    Set fd = Application.FileDialog(4)   ' msoFileDialogFolderPicker

    fd.Title = "階層を読み取るフォルダを選択してください"
    current = Trim$(CStr(cfg.Range(CELL_ROOT).Value))
    If Len(current) > 0 Then
        On Error Resume Next
        fd.InitialFileName = current
        On Error GoTo 0
    End If

    If fd.Show = -1 Then
        cfg.Range(CELL_ROOT).Value = fd.SelectedItems(1)
    End If
End Sub

'--------------------------------------------------------------
' 2. ツリーを作成
'--------------------------------------------------------------
Public Sub BuildTree()
    Dim cfg As Worksheet
    Dim rootPath As String
    Dim maxDepth As Long
    Dim summary As String

    Set cfg = ThisWorkbook.Worksheets(SHEET_CONFIG)
    rootPath = Trim$(CStr(cfg.Range(CELL_ROOT).Value))

    If Len(rootPath) = 0 Then
        MsgBox "対象フォルダが指定されていません。" & vbCrLf & _
               "[1. フォルダを選択] から指定してください。", _
               vbExclamation, "フォルダ階層ツリー"
        Exit Sub
    End If

    If Not DirectoryExists(rootPath) Then
        MsgBox "フォルダが見つかりません:" & vbCrLf & rootPath, _
               vbExclamation, "フォルダ階層ツリー"
        Exit Sub
    End If

    maxDepth = 0
    If IsNumeric(cfg.Range(CELL_MAXDEPTH).Value) Then
        maxDepth = CLng(cfg.Range(CELL_MAXDEPTH).Value)
        If maxDepth < 0 Then maxDepth = 0
    End If

    On Error GoTo Fail
    ' UI経路はテスト注入状態を明示的に無効化してから通常走査を開始する。
    ResetTestHooks
    summary = RunScan(rootPath, maxDepth, _
                      ToBool(cfg.Range(CELL_FILES).Value), _
                      ToBool(cfg.Range(CELL_HIDDEN).Value))
    MsgBox summary, vbInformation, "フォルダ階層ツリー"
    Exit Sub

Fail:
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & vbCrLf & _
           "エラー " & Err.Number & ": " & Err.Description, _
           vbCritical, "フォルダ階層ツリー"
End Sub

'--------------------------------------------------------------
' スキャン本体（UI を伴わないので自動テストからも呼び出せる）
'   戻り値: 結果サマリーの文字列
'--------------------------------------------------------------
Public Function RunScan(ByVal rootPath As String, ByVal maxDepth As Long, _
                        ByVal includeFiles As Boolean, ByVal includeHidden As Boolean) As String
    Dim ws As Worksheet
    Dim startedAt As Double
    Dim elapsed As Double
    Dim msg As String
    Dim savedNumber As Long
    Dim savedDescription As String

    If mOperationBlocksScan Then
        RunScan = "ファイル操作中です。操作完了後に再走査してください。"
        Exit Function
    End If
    If mRunning Then
        RunScan = "処理中です。現在の走査が完了してから再実行してください。"
        Exit Function
    End If

    rootPath = NormalizeRoot(rootPath)
    If Not DirectoryExists(rootPath) Then
        Err.Raise 76, "modFolderTree.RunScan", "フォルダが見つかりません: " & rootPath
    End If

    mRunning = True
    mScanReady = False
    mScanRoot = rootPath

    mMaxDepth = maxDepth
    If mMaxDepth < 0 Then mMaxDepth = 0
    mIncludeFiles = includeFiles
    mIncludeHidden = includeHidden

    mCount = 0
    mCapacity = 0
    mFolderCount = 0
    mDeniedCount = 0
    mIoErrorCount = 0
    mLinkCount = 0
    mCycleCount = 0
    mDepthReached = 0
    mTruncated = False
    mAllocFailed = False
    mCancelled = False
    mDepthLimitHit = False
    mOutlineErrors = 0
    mFirstOutlineError = ""
    mAllocErrorNumber = 0
    mAllocErrorDescription = ""
    mTotalSize = 0
    mTotalFiles = 0
    mOutputLimit = MAX_OUTPUT_ROWS
    If mTestRowLimit > 0 And mTestRowLimit < mOutputLimit Then mOutputLimit = mTestRowLimit
    mChunkRows = OUTPUT_CHUNK_ROWS
    If mTestChunkRows > 0 Then mChunkRows = mTestChunkRows
    Erase mBuf
    Erase mLevel
    Erase mSnapshotAttributes
    Erase mSnapshotReparse

    startedAt = Timer
    mPrevScreenUpdating = Application.ScreenUpdating
    mPrevEnableEvents = Application.EnableEvents
    mPrevCalculation = Application.Calculation
    mPrevCursor = Application.Cursor
    mPrevCancel = Application.EnableCancelKey
    mPrevStatusBar = Application.StatusBar

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait
    Application.EnableCancelKey = xlErrorHandler

    ScanTree rootPath

    Application.EnableCancelKey = xlDisabled
    Set ws = EnsureTreeSheet()
    WriteToSheet ws, rootPath

    elapsed = Timer - startedAt
    If elapsed < 0 Then elapsed = elapsed + 86400   ' 日付をまたいだ場合

    Application.EnableCancelKey = xlDisabled
    RestoreApplicationState
    mRunning = False
    mScanReady = (Not mCancelled And Not mAllocFailed And Not mTruncated And mIoErrorCount = 0)

    msg = "ツリーを作成しました。" & vbCrLf & vbCrLf & _
          "対象      : " & rootPath & vbCrLf & _
          "フォルダ  : " & Format$(mFolderCount, "#,##0") & " 個" & vbCrLf & _
          "ファイル  : " & Format$(mTotalFiles, "#,##0") & " 個" & vbCrLf & _
          "合計サイズ: " & FormatSize(mTotalSize) & vbCrLf & _
          "最大階層  : " & mDepthReached & vbCrLf & _
          "出力行数  : " & Format$(mCount, "#,##0") & " 行" & vbCrLf & _
          "所要時間  : " & Format$(elapsed, "0.0") & " 秒"

    If mCancelled Then
        msg = "Esc キーで中断しました。中断時点までの結果を書き出しています。" & vbCrLf & vbCrLf & msg
    End If
    If mMaxDepth > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "※ 最大階層 " & mMaxDepth & " で打ち切っています。" & vbCrLf & _
              "   打ち切った先のファイル数・サイズは集計に含まれません。"
    End If
    If mDeniedCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "※ アクセスできなかったフォルダが " & Format$(mDeniedCount, "#,##0") & " 個あります。" & vbCrLf & _
              "   該当行の [種別] に「一部アクセス不可」と表示しています。"
    End If
    If mLinkCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "※ シンボリックリンク／ジャンクションを " & Format$(mLinkCount, "#,##0") & " 個検出しました。" & vbCrLf & _
              "   循環を避けるため、リンク先はたどっていません。"
    End If
    If mCycleCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "※ 循環する参照を " & Format$(mCycleCount, "#,##0") & " 個検出し、その先をたどっていません。"
    End If
    If mAllocFailed Then
        msg = msg & vbCrLf & vbCrLf & _
              "※ メモリを確保できなくなったため途中で打ち切りました。" & vbCrLf & _
              "   最大階層を小さくするか、ファイル表示をオフにしてください。"
        If mAllocErrorNumber <> 0 Then
            msg = msg & vbCrLf & "   確保エラー " & mAllocErrorNumber & ": " & mAllocErrorDescription
        End If
    ElseIf mTruncated Then
        msg = msg & vbCrLf & vbCrLf & _
               "※ 出力行数が上限 " & Format$(mOutputLimit, "#,##0") & " 行に達したため途中で打ち切りました。" & vbCrLf & _
              "   最大階層を小さくするか、ファイル表示をオフにしてください。"
    End If

    If mDepthLimitHit Then
        msg = msg & vbCrLf & vbCrLf & "※ ユーザー上限なしでも内部安全弁 " & MAX_SCAN_DEPTH & " 階層で打ち切りました。打ち切り先の集計は含まれません。"
    End If
    If mIoErrorCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & "※ 列挙不完全またはI/Oエラー: " & Format$(mIoErrorCount, "#,##0") & " 件。完全な全体集計ではありません。"
    End If
    If mOutlineErrors > 0 Then
        msg = msg & vbCrLf & vbCrLf & "※ アウトライン／インデント設定に " & mOutlineErrors & " 件失敗しました。"
        If Len(mFirstOutlineError) > 0 Then msg = msg & vbCrLf & "   最初のエラー: " & mFirstOutlineError
    End If
    RunScan = msg
    Exit Function

Fail:
    savedNumber = Err.Number
    savedDescription = Err.Description
    RestoreApplicationState
    mRunning = False
    mScanReady = False
    Err.Raise savedNumber, "modFolderTree.RunScan", savedDescription
End Function

Public Sub SetScanBlockedByOperation(ByVal blocked As Boolean)
    mOperationBlocksScan = blocked
End Sub

Public Function RefreshTreeAfterOperations(ByVal rootPath As String) As String
    Dim cfg As Worksheet, maxDepth As Long
    Set cfg = ThisWorkbook.Worksheets(SHEET_CONFIG)
    maxDepth = 0
    If IsNumeric(cfg.Range(CELL_MAXDEPTH).Value) Then
        maxDepth = CLng(cfg.Range(CELL_MAXDEPTH).Value)
        If maxDepth < 0 Then maxDepth = 0
    End If
    ResetTestHooks
    RefreshTreeAfterOperations = RunScan(rootPath, maxDepth, _
                                         ToBool(cfg.Range(CELL_FILES).Value), _
                                         ToBool(cfg.Range(CELL_HIDDEN).Value))
End Function

Public Function IsScanReadyForRoot(ByVal rootPath As String) As Boolean
    If Not mScanReady Then Exit Function
    On Error GoTo Fail
    IsScanReadyForRoot = (StrComp(NormalizeRoot(rootPath), mScanRoot, vbTextCompare) = 0)
    Exit Function
Fail:
    IsScanReadyForRoot = False
End Function

Public Function GetScanSnapshotCount() As Long
    If mScanReady Then GetScanSnapshotCount = mCount
End Function

Public Function GetScanSnapshotRow(ByVal rowIndex As Long, ByRef itemName As String, ByRef itemType As String, _
                                   ByRef depth As Long, ByRef items As Variant, ByRef itemSize As Double, _
                                   ByRef modified As Variant, ByRef fullPath As String, ByRef attrs As Long, _
                                   ByRef reparse As Boolean) As Boolean
    If Not mScanReady Then Exit Function
    If rowIndex < 1 Or rowIndex > mCount Then Exit Function
    itemName = CStr(mBuf(COL_NAME, rowIndex))
    itemType = CStr(mBuf(COL_TYPE, rowIndex))
    depth = CLng(mBuf(COL_DEPTH, rowIndex))
    items = mBuf(COL_ITEMS, rowIndex)
    If IsNumeric(mBuf(COL_SIZE, rowIndex)) Then itemSize = CDbl(mBuf(COL_SIZE, rowIndex)) Else itemSize = 0
    modified = mBuf(COL_DATE, rowIndex)
    fullPath = CStr(mBuf(COL_PATH, rowIndex))
    attrs = mSnapshotAttributes(rowIndex)
    reparse = mSnapshotReparse(rowIndex)
    GetScanSnapshotRow = True
End Function

#If TEST_BUILD Then
Public Function RunScanTest(ByVal rootPath As String, ByVal maxDepth As Long, _
                            ByVal includeFiles As Boolean, ByVal includeHidden As Boolean, _
                            Optional ByVal rowLimit As Long = 0, _
                            Optional ByVal chunkRows As Long = 0, _
                            Optional ByVal failAllocation As Boolean = False, _
                            Optional ByVal failOutline As Boolean = False, _
                            Optional ByVal cancelAfterRows As Long = 0, _
                            Optional ByVal findFirstError As Long = 0, _
                            Optional ByVal findNextError As Long = 0) As String
    Dim savedNumber As Long
    Dim savedDescription As String

    mTestRowLimit = rowLimit
    mTestChunkRows = chunkRows
    mTestFailAllocation = failAllocation
    mTestFailOutline = failOutline
    mTestCancelAfterRows = cancelAfterRows
    mTestFindFirstError = findFirstError
    mTestFindNextError = findNextError
    On Error GoTo Fail
    RunScanTest = RunScan(rootPath, maxDepth, includeFiles, includeHidden)
    ResetTestHooks
    Exit Function

Fail:
    savedNumber = Err.Number
    savedDescription = Err.Description
    ResetTestHooks
    Err.Raise savedNumber, "modFolderTree.RunScanTest", savedDescription
End Function
#End If

Private Sub ResetTestHooks()
    mTestRowLimit = 0
    mTestChunkRows = 0
    mTestFailAllocation = False
    mTestFailOutline = False
    mTestCancelAfterRows = 0
    mTestFindFirstError = 0
    mTestFindNextError = 0
End Sub

#If TEST_BUILD Then
Public Function ClassifyFindResult(ByVal handleValid As Boolean, ByVal lastError As Long) As String
    If Not handleValid Then
        If lastError = ERROR_ACCESS_DENIED Then
            ClassifyFindResult = "access-denied"
        Else
            ClassifyFindResult = "findfirst-error"
        End If
    ElseIf lastError = ERROR_NO_MORE_FILES Then
        ClassifyFindResult = "complete"
    Else
        ClassifyFindResult = "incomplete"
    End If
End Function

Public Function ProbeReentry(ByVal rootPath As String) As String
    On Error GoTo Fail
    mRunning = True
    ProbeReentry = RunScan(rootPath, 0, True, False)
    mRunning = False
    Exit Function

Fail:
    mRunning = False
    Err.Raise Err.Number, "modFolderTree.ProbeReentry", Err.Description
End Function

Public Function RunStringSafetyProbe(ByVal injectedName As String) As String
    Dim ws As Worksheet

    mCount = 0
    mCapacity = 0
    mOutputLimit = 1
    mChunkRows = 1
    mAllocFailed = False
    mTruncated = False
    Erase mBuf
    Erase mLevel
    If AddRow("ファイル", injectedName, 0, Empty, 0, Empty, injectedName) = 0 Then
        Err.Raise 7, "modFolderTree.RunStringSafetyProbe", "テスト行を追加できません。"
    End If
    Set ws = EnsureTreeSheet()
    WriteToSheet ws, injectedName
    RunStringSafetyProbe = CStr(ws.Cells(FIRST_ROW, COL_NAME).Value2)
End Function
#End If

Private Sub RestoreApplicationState()
    On Error Resume Next
    Application.ScreenUpdating = mPrevScreenUpdating
    Application.EnableEvents = mPrevEnableEvents
    Application.Calculation = mPrevCalculation
    Application.Cursor = mPrevCursor
    Application.EnableCancelKey = mPrevCancel
    Application.StatusBar = mPrevStatusBar
    On Error GoTo 0
End Sub

'==============================================================
' 走査（明示スタックによる反復処理。再帰は使わない）
'==============================================================
Private Sub ScanTree(ByVal rootPath As String)
    Dim idx As Long
    Dim childName As String
    Dim childPath As String
    Dim doneSize As Double
    Dim doneFiles As Double
    Dim subs As Variant
    Dim attrs As Variant
    Dim dates As Variant
    Dim childIndex As Long

    mStTop = 0
    mStCap = 0
    EnsureStackCapacity

    On Error GoTo Cancelled

    PushDirectory rootPath, 0, RootDisplayName(rootPath), RootAttributes(rootPath), RootModified(rootPath)

    Do While mStTop > 0
        idx = mStTop

        If mStNext(idx) > mStNum(idx) Or mTruncated Then
            ' このフォルダの処理が終わった → 集計をバッファへ書き戻し、親へ加算する
            If mStRow(idx) > 0 Then
                mBuf(COL_ITEMS, mStRow(idx)) = mStFiles(idx)
                mBuf(COL_SIZE, mStRow(idx)) = mStSize(idx)
            End If
            doneSize = mStSize(idx)
            doneFiles = mStFiles(idx)
            mStSubs(idx) = Empty
            mStAttrs(idx) = Empty
            mStDates(idx) = Empty
            mStTop = mStTop - 1
            If mStTop > 0 Then
                mStSize(mStTop) = mStSize(mStTop) + doneSize
                mStFiles(mStTop) = mStFiles(mStTop) + doneFiles
            Else
                mTotalSize = doneSize
                mTotalFiles = doneFiles
            End If
        Else
            subs = mStSubs(idx)
            attrs = mStAttrs(idx)
            dates = mStDates(idx)
            childIndex = mStNext(idx)
            childName = subs(childIndex)
            mStNext(idx) = mStNext(idx) + 1
            childPath = CombinePath(mStPath(idx), childName)
            PushDirectory childPath, mStDepth(idx) + 1, childName, _
                          CLng(attrs(childIndex)), dates(childIndex)
            If mCancelled Then Exit Do
        End If
    Loop
    Exit Sub

Cancelled:
    If Err.Number = 18 Then
        ' Esc による中断。積み残しのフォルダは集計を確定させずに抜ける
        mCancelled = True
        Exit Sub
    End If
    Err.Raise Err.Number, "modFolderTree.ScanTree", Err.Description
End Sub

'--------------------------------------------------------------
' 1 フォルダをスタックへ積む
'   ・自分の行をバッファへ追加
'   ・直下のファイルを列挙して集計（表示設定なら行も追加）
'   ・直下のサブフォルダ名を集めてフレームに保持
'--------------------------------------------------------------
Private Sub PushDirectory(ByVal dirPath As String, ByVal depth As Long, _
                          ByVal dispName As String, ByVal dirAttrs As Long, _
                          ByVal dirModified As Variant)
    Dim fd As WIN32_FIND_DATAW
    Dim subNames() As String
    Dim subAttrs() As Long
    Dim subDates() As Variant
    Dim subCount As Long
    Dim accSize As Double
    Dim accFiles As Double
    Dim itemType As String
    Dim myRow As Long
    Dim entryName As String
    Dim entryAttrs As Long
    Dim isDir As Boolean
    Dim denied As Boolean
    Dim collectSubs As Boolean
    Dim ok As Long
    Dim lastError As Long
    Dim findIncomplete As Boolean
    Dim closeOk As Long
    #If VBA7 Then
        Dim hFind As LongPtr
    #Else
        Dim hFind As Long
    #End If

    If mTruncated Then Exit Sub
    If depth > mDepthReached Then mDepthReached = depth

    itemType = "フォルダ"
    If (dirAttrs And FILE_ATTRIBUTE_REPARSE_POINT) <> 0 Then
        itemType = "リンク"
        mLinkCount = mLinkCount + 1
    End If

    If IsEmpty(dirModified) Then dirModified = Empty

    myRow = AddRow(itemType, dispName, depth, Empty, Empty, dirModified, dirPath)
    If myRow = 0 Then Exit Sub
    If mCancelled Then Exit Sub
    mFolderCount = mFolderCount + 1

    If (mFolderCount Mod 500) = 0 Then
        Application.StatusBar = "スキャン中... フォルダ " & Format$(mFolderCount, "#,##0") & _
                                " 個 / 出力 " & Format$(mCount, "#,##0") & " 行  (Esc で中断)"
        DoEvents
    End If

    ' reparse pointはリンク行だけを出力し、リンク先の直下エントリも列挙しない。
    If (dirAttrs And FILE_ATTRIBUTE_REPARSE_POINT) <> 0 Then Exit Sub

    ' リンク先・深さ上限・循環はたどらない
    collectSubs = True
    If (dirAttrs And FILE_ATTRIBUTE_REPARSE_POINT) <> 0 Then collectSubs = False
    If mMaxDepth > 0 And depth >= mMaxDepth Then collectSubs = False
    If depth >= MAX_SCAN_DEPTH Then
        collectSubs = False
        mDepthLimitHit = True
    End If
    If IsAncestorPath(dirPath) Then
        mCycleCount = mCycleCount + 1
        collectSubs = False
    End If

    subCount = 0
    ReDim subNames(1 To 16)
    ReDim subAttrs(1 To 16)
    ReDim subDates(1 To 16)

    hFind = FindFirstFileW(StrPtr(SearchPattern(dirPath)), fd)
    lastError = GetLastError()
    If mTestFindFirstError <> 0 Then
        hFind = -1
        lastError = mTestFindFirstError
        mTestFindFirstError = 0
    End If
    If hFind = -1 Then
        If lastError = ERROR_ACCESS_DENIED Then
            denied = True
        Else
            findIncomplete = True
            mIoErrorCount = mIoErrorCount + 1
        End If
    Else
        Do
            entryName = FindDataName(fd)
            If entryName <> "." And entryName <> ".." And Len(entryName) > 0 Then
                entryAttrs = fd.dwFileAttributes
                isDir = ((entryAttrs And FILE_ATTRIBUTE_DIRECTORY) <> 0)

                If mIncludeHidden Or _
                   ((entryAttrs And (FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_SYSTEM)) = 0) Then
                    If isDir Then
                        If collectSubs Then
                            subCount = subCount + 1
                            If subCount > UBound(subNames) Then
                                ReDim Preserve subNames(1 To UBound(subNames) * 2)
                                ReDim Preserve subAttrs(1 To UBound(subAttrs) * 2)
                                ReDim Preserve subDates(1 To UBound(subDates) * 2)
                            End If
                            subNames(subCount) = entryName
                            subAttrs(subCount) = entryAttrs
                            subDates(subCount) = FindDataTime(fd)
                        End If
                    Else
                        accFiles = accFiles + 1
                        accSize = accSize + FindDataSize(fd)
                        If mIncludeFiles Then
                            If AddRow("ファイル", entryName, depth + 1, Empty, FindDataSize(fd), _
                                      FindDataTime(fd), CombinePath(dirPath, entryName)) = 0 Then Exit Do
                            If mCancelled Then Exit Do
                            If mTruncated Then Exit Do
                        End If
                    End If
                End If
            End If
            ok = FindNextFileW(hFind, fd)
            lastError = GetLastError()
            If mTestFindNextError <> 0 Then
                ok = 0
                lastError = mTestFindNextError
                mTestFindNextError = 0
            End If
        Loop While ok <> 0
        If lastError <> ERROR_NO_MORE_FILES Then
            findIncomplete = True
            mIoErrorCount = mIoErrorCount + 1
        End If
        closeOk = FindClose(hFind)
        If closeOk = 0 Then mIoErrorCount = mIoErrorCount + 1
    End If

    If denied Then
        mDeniedCount = mDeniedCount + 1
        If myRow > 0 Then
            mBuf(COL_TYPE, myRow) = CStr(mBuf(COL_TYPE, myRow)) & "（一部アクセス不可）"
        End If
    End If
    If findIncomplete And myRow > 0 Then
        mBuf(COL_TYPE, myRow) = CStr(mBuf(COL_TYPE, myRow)) & "（列挙不完全: error " & lastError & "）"
    End If

    ' フレームを積む
    mStTop = mStTop + 1
    EnsureStackCapacity

    mStPath(mStTop) = dirPath
    mStDepth(mStTop) = depth
    mStRow(mStTop) = myRow
    mStNum(mStTop) = subCount
    mStNext(mStTop) = 1
    mStSize(mStTop) = accSize
    mStFiles(mStTop) = accFiles
    If subCount > 0 Then
        mStSubs(mStTop) = subNames
        mStAttrs(mStTop) = subAttrs
        mStDates(mStTop) = subDates
    Else
        mStSubs(mStTop) = Empty
        mStAttrs(mStTop) = Empty
        mStDates(mStTop) = Empty
    End If
End Sub

Private Sub EnsureStackCapacity()
    If mStTop < mStCap Then Exit Sub

    If mStCap = 0 Then
        mStCap = 64
        ReDim mStPath(1 To mStCap)
        ReDim mStDepth(1 To mStCap)
        ReDim mStRow(1 To mStCap)
        ReDim mStSubs(1 To mStCap)
        ReDim mStAttrs(1 To mStCap)
        ReDim mStDates(1 To mStCap)
        ReDim mStNum(1 To mStCap)
        ReDim mStNext(1 To mStCap)
        ReDim mStSize(1 To mStCap)
        ReDim mStFiles(1 To mStCap)
    Else
        mStCap = mStCap * 2
        ReDim Preserve mStPath(1 To mStCap)
        ReDim Preserve mStDepth(1 To mStCap)
        ReDim Preserve mStRow(1 To mStCap)
        ReDim Preserve mStSubs(1 To mStCap)
        ReDim Preserve mStAttrs(1 To mStCap)
        ReDim Preserve mStDates(1 To mStCap)
        ReDim Preserve mStNum(1 To mStCap)
        ReDim Preserve mStNext(1 To mStCap)
        ReDim Preserve mStSize(1 To mStCap)
        ReDim Preserve mStFiles(1 To mStCap)
    End If
End Sub

' 祖先に同じパスがあれば循環とみなす
Private Function IsAncestorPath(ByVal path As String) As Boolean
    Dim i As Long
    Dim target As String

    target = LCase$(path)
    For i = 1 To mStTop
        If LCase$(mStPath(i)) = target Then
            IsAncestorPath = True
            Exit Function
        End If
    Next i
End Function

'==============================================================
' Win32 ヘルパー
'==============================================================
Private Function SearchPattern(ByVal dirPath As String) As String
    SearchPattern = ExtendedPath(CombinePath(dirPath, "*"))
End Function

' 260 文字前後を超えるパスは \\?\ プレフィックスで扱う
Private Function ExtendedPath(ByVal path As String) As String
    If Len(path) < LONG_PATH_THRESHOLD Then
        ExtendedPath = path
    ElseIf Left$(path, 4) = "\\?\" Then
        ExtendedPath = path
    ElseIf Left$(path, 2) = "\\" Then
        ExtendedPath = "\\?\UNC\" & Mid$(path, 3)
    Else
        ExtendedPath = "\\?\" & path
    End If
End Function

Private Function CombinePath(ByVal basePath As String, ByVal leaf As String) As String
    If Len(basePath) = 0 Then
        CombinePath = leaf
    ElseIf Right$(basePath, 1) = "\" Then
        CombinePath = basePath & leaf
    Else
        CombinePath = basePath & "\" & leaf
    End If
End Function

Private Function FindDataName(ByRef fd As WIN32_FIND_DATAW) As String
    Dim s As String
    Dim p As Long

    s = Space$(MAX_PATH_CHARS)
    CopyMemory StrPtr(s), VarPtr(fd.cFileName(0)), MAX_PATH_CHARS * 2
    p = InStr(1, s, vbNullChar)
    If p > 0 Then
        FindDataName = Left$(s, p - 1)
    Else
        FindDataName = s
    End If
End Function

Private Function FindDataSize(ByRef fd As WIN32_FIND_DATAW) As Double
    Dim lo As Double
    Dim hi As Double

    lo = fd.nFileSizeLow
    If lo < 0 Then lo = lo + 4294967296#
    hi = fd.nFileSizeHigh
    If hi < 0 Then hi = hi + 4294967296#

    FindDataSize = hi * 4294967296# + lo
End Function

Private Function FindDataTime(ByRef fd As WIN32_FIND_DATAW) As Variant
    FindDataTime = FileTimeToDate(fd.ftLastWriteTime)
End Function

Private Function FileTimeToDate(ByRef ft As FILETIME) As Variant
    Dim lft As FILETIME
    Dim st As SYSTEMTIME

    FileTimeToDate = Empty
    If ft.dwLowDateTime = 0 And ft.dwHighDateTime = 0 Then Exit Function
    If FileTimeToLocalFileTime(ft, lft) = 0 Then Exit Function
    If FileTimeToSystemTime(lft, st) = 0 Then Exit Function
    If st.wYear < 1900 Or st.wYear > 9999 Then Exit Function

    On Error Resume Next
    FileTimeToDate = DateSerial(st.wYear, st.wMonth, st.wDay) + _
                     TimeSerial(st.wHour, st.wMinute, st.wSecond)
    On Error GoTo 0
End Function

Private Function GetAttributesOf(ByVal path As String) As Long
    Dim a As Long

    a = GetFileAttributesW(StrPtr(ExtendedPath(path)))
    If a = INVALID_FILE_ATTRS Then
        GetAttributesOf = FILE_ATTRIBUTE_DIRECTORY
    Else
        GetAttributesOf = a
    End If
End Function

Private Function DirectoryExists(ByVal path As String) As Boolean
    Dim a As Long

    If Len(Trim$(path)) = 0 Then Exit Function
    a = GetFileAttributesW(StrPtr(ExtendedPath(NormalizeRoot(path))))
    If a = INVALID_FILE_ATTRS Then Exit Function

    DirectoryExists = ((a And FILE_ATTRIBUTE_DIRECTORY) <> 0)
End Function

' 末尾の \ を落とす（ただし "C:\" のようなドライブ直下は残す）
Private Function NormalizeRoot(ByVal path As String) As String
    Dim p As String

    p = Trim$(path)
    Do While Len(p) > 3 And Right$(p, 1) = "\"
        p = Left$(p, Len(p) - 1)
    Loop
    NormalizeRoot = p
End Function

Private Function RootDisplayName(ByVal path As String) As String
    Dim p As Long

    If Len(path) <= 3 And Mid$(path, 2, 1) = ":" Then
        RootDisplayName = path
        Exit Function
    End If

    p = InStrRev(path, "\")
    If p > 0 Then
        RootDisplayName = Mid$(path, p + 1)
    Else
        RootDisplayName = path
    End If
End Function

Private Function RootAttributes(ByVal path As String) As Long
    RootAttributes = GetAttributesOf(path)
End Function

Private Function RootModified(ByVal path As String) As Variant
    Dim fd As WIN32_FIND_DATAW
    #If VBA7 Then
        Dim hFind As LongPtr
    #Else
        Dim hFind As Long
    #End If

    RootModified = Empty
    If Len(path) <= 3 Then Exit Function   ' ドライブ直下は FindFirstFileW で取得できない

    hFind = FindFirstFileW(StrPtr(ExtendedPath(path)), fd)
    If hFind = -1 Then Exit Function

    RootModified = FindDataTime(fd)
    FindClose hFind
End Function

'==============================================================
' 出力バッファ
'==============================================================
Private Function AddRow(ByVal itemType As String, ByVal itemName As String, ByVal depth As Long, _
                         ByVal items As Variant, ByVal itemSize As Variant, _
                         ByVal modified As Variant, ByVal fullPath As String) As Long
    If mCount >= mOutputLimit Then
        mTruncated = True
        AddRow = 0
        Exit Function
    End If

    EnsureCapacity
    If mCount >= mCapacity Then
        If Not mAllocFailed Then mTruncated = True
        AddRow = 0
        Exit Function
    End If

    mCount = mCount + 1

    mBuf(COL_NAME, mCount) = itemName
    mBuf(COL_TYPE, mCount) = itemType
    mBuf(COL_DEPTH, mCount) = depth
    mBuf(COL_ITEMS, mCount) = items
    mBuf(COL_SIZE, mCount) = itemSize
    mBuf(COL_DATE, mCount) = modified
    mBuf(COL_PATH, mCount) = fullPath
    mSnapshotAttributes(mCount) = GetAttributesOf(fullPath)
    mSnapshotReparse(mCount) = ((mSnapshotAttributes(mCount) And FILE_ATTRIBUTE_REPARSE_POINT) <> 0)
    mLevel(mCount) = depth
    If mTestCancelAfterRows > 0 And mCount >= mTestCancelAfterRows Then mCancelled = True
    AddRow = mCount
End Function

Private Sub EnsureCapacity()
    Dim newCap As Long
    Dim ok As Boolean

    If mCount < mCapacity Then Exit Sub

    If mCapacity = 0 Then
        newCap = 16384
    Else
        newCap = mCapacity * 2
    End If
    If newCap > mOutputLimit Then newCap = mOutputLimit

    If mTestFailAllocation Then
        mAllocFailed = True
        mTruncated = True
        mAllocErrorNumber = 7
        mAllocErrorDescription = "テスト注入: バッファ確保失敗"
        Exit Sub
    End If

    ok = False
    On Error Resume Next
    If mCapacity = 0 Then
        ReDim mBuf(1 To COL_MAX, 1 To newCap)
        ReDim mLevel(1 To newCap)
        ReDim mSnapshotAttributes(1 To newCap)
        ReDim mSnapshotReparse(1 To newCap)
    Else
        ReDim Preserve mBuf(1 To COL_MAX, 1 To newCap)
        ReDim Preserve mLevel(1 To newCap)
        ReDim Preserve mSnapshotAttributes(1 To newCap)
        ReDim Preserve mSnapshotReparse(1 To newCap)
    End If
    If Err.Number <> 0 Then
        mAllocErrorNumber = Err.Number
        mAllocErrorDescription = Err.Description
        Err.Clear
    End If
    ok = (mCapacity < newCap And UBound(mLevel) = newCap) Or (mCapacity = newCap)
    On Error GoTo 0

    If ok Then
        mCapacity = newCap
    Else
        mAllocFailed = True
        mTruncated = True
    End If
End Sub

'==============================================================
' シート出力
'==============================================================
Private Sub WriteToSheet(ByVal ws As Worksheet, ByVal rootPath As String)
    Dim dataRange As Range
    Dim lastRow As Long

    Application.StatusBar = "シートへ書き出し中..."

    On Error Resume Next
    ws.AutoFilterMode = False
    ws.Cells.ClearOutline
    On Error GoTo 0
    ws.Cells.Clear
    ws.Cells.FormatConditions.Delete

    ws.Range("A1").Value2 = "対象フォルダ"
    ws.Range("B1").NumberFormat = "@"
    ws.Range("B1").Value2 = rootPath
    ws.Range("A2").Value = "生成日時"
    ws.Range("B2").Value = Now
    ws.Range("B2").NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Range("A1:A2").Font.Bold = True

    ws.Cells(HEADER_ROW, COL_NAME).Value = "名前"
    ws.Cells(HEADER_ROW, COL_TYPE).Value = "種別"
    ws.Cells(HEADER_ROW, COL_DEPTH).Value = "階層"
    ws.Cells(HEADER_ROW, COL_ITEMS).Value = "配下ファイル数"
    ws.Cells(HEADER_ROW, COL_SIZE).Value = "サイズ(byte)"
    ws.Cells(HEADER_ROW, COL_DATE).Value = "更新日時"
    ws.Cells(HEADER_ROW, COL_PATH).Value = "フルパス"

    With ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, COL_MAX))
        .Font.Bold = True
        .Interior.Color = RGB(221, 231, 243)
        .HorizontalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With

    If mCount = 0 Then
        ws.Cells(FIRST_ROW, COL_NAME).Value = "（該当する項目がありません）"
        Exit Sub
    End If

    lastRow = FIRST_ROW + mCount - 1

    ' アウトラインとインデントは「値を書き込む前」に設定する。
    ' 同じ操作でも、データが入ったシートに対して行うと 1 回あたり約 9 倍遅くなる
    ' （実測 10.8 ms → 1.2 ms）。大きなツリーでは数分の差になる。
    ApplyIndentAndOutline ws

    Application.StatusBar = "シートへ書き出し中..."

    Set dataRange = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(lastRow, COL_MAX))
    ws.Range(ws.Cells(FIRST_ROW, COL_NAME), ws.Cells(lastRow, COL_NAME)).NumberFormat = "@"
    ws.Range(ws.Cells(FIRST_ROW, COL_TYPE), ws.Cells(lastRow, COL_TYPE)).NumberFormat = "@"
    ws.Range(ws.Cells(FIRST_ROW, COL_PATH), ws.Cells(lastRow, COL_PATH)).NumberFormat = "@"
    WriteOutputChunks ws

    ws.Range(ws.Cells(FIRST_ROW, COL_ITEMS), ws.Cells(lastRow, COL_ITEMS)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(FIRST_ROW, COL_SIZE), ws.Cells(lastRow, COL_SIZE)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(FIRST_ROW, COL_DATE), ws.Cells(lastRow, COL_DATE)).NumberFormat = "yyyy/mm/dd hh:mm"
    ws.Range(ws.Cells(FIRST_ROW, COL_DEPTH), ws.Cells(lastRow, COL_DEPTH)).HorizontalAlignment = xlCenter

    With dataRange.FormatConditions.Add(xlExpression, , _
            "=LEFT($B" & FIRST_ROW & ",4)=""フォルダ""")
        .Font.Bold = True
        .Interior.Color = RGB(242, 246, 252)
    End With
    With dataRange.FormatConditions.Add(xlExpression, , _
            "=OR(ISNUMBER(SEARCH(""アクセス不可"",$B" & FIRST_ROW & ")),ISNUMBER(SEARCH(""列挙不完全"",$B" & FIRST_ROW & ")))")
        .Font.Color = RGB(176, 0, 32)
    End With

    ws.Columns(COL_NAME).ColumnWidth = 44
    ws.Columns(COL_TYPE).ColumnWidth = 22
    ws.Columns(COL_DEPTH).ColumnWidth = 6
    ws.Columns(COL_ITEMS).ColumnWidth = 14
    ws.Columns(COL_SIZE).ColumnWidth = 15
    ws.Columns(COL_DATE).ColumnWidth = 18
    ws.Columns(COL_PATH).ColumnWidth = 80

    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lastRow, COL_MAX)).AutoFilter

    ws.Activate
    ws.Cells(FIRST_ROW, 1).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
End Sub

Private Sub WriteOutputChunks(ByVal ws As Worksheet)
    Dim out() As Variant
    Dim startIndex As Long
    Dim endIndex As Long
    Dim i As Long
    Dim c As Long
    Dim rowCount As Long
    Dim chunkRange As Range

    startIndex = 1
    Do While startIndex <= mCount
        endIndex = startIndex + mChunkRows - 1
        If endIndex > mCount Then endIndex = mCount
        rowCount = endIndex - startIndex + 1
        ReDim out(1 To rowCount, 1 To COL_MAX)
        For i = 1 To rowCount
            For c = 1 To COL_MAX
                out(i, c) = mBuf(c, startIndex + i - 1)
            Next c
        Next i
        Set chunkRange = ws.Range(ws.Cells(FIRST_ROW + startIndex - 1, 1), _
                                  ws.Cells(FIRST_ROW + endIndex - 1, COL_MAX))
        chunkRange.Value2 = out
        startIndex = endIndex + 1
        DoEvents
    Loop
End Sub

'--------------------------------------------------------------
' インデントとアウトライン
'
'   インデント: 階層ごとに複数範囲を 1 つのアドレス文字列へ束ねて
'               設定できる（COM 呼び出し回数を削減）。
'   アウトライン: OutlineLevel も Group も複数範囲をまとめた Range を
'               受け付けないため、単一範囲で呼び出すしかない。
'               そこで「ブロックごとに OutlineLevel を代入」ではなく
'               「階層 L 以上の行のかたまりを Group する」を L=1..7 で
'               繰り返す。行 d は d 回グループ化され、結果として
'               OutlineLevel = min(d + 1, 8) になる。
'               実測ではこの方式で呼び出し回数が約 1/8 になる
'               （63,249 行のツリーで 19,276 回 → 2,549 回）。
'--------------------------------------------------------------
Private Sub ApplyIndentAndOutline(ByVal ws As Worksheet)
    Dim indentAddr(0 To MAX_INDENT_LEVEL) As String
    Dim i As Long
    Dim blockStart As Long
    Dim curLevel As Long
    Dim isBoundary As Boolean
    Dim indent As Long

    ws.Outline.SummaryRow = xlSummaryAbove
    ws.Outline.SummaryColumn = xlSummaryOnLeft

    ' --- インデント（連続ブロックを階層ごとにまとめて設定）---
    blockStart = 1
    curLevel = mLevel(1)

    For i = 2 To mCount + 1
        isBoundary = True
        If i <= mCount Then
            If mLevel(i) = curLevel Then isBoundary = False
        End If

        If isBoundary Then
            indent = curLevel
            If indent > MAX_INDENT_LEVEL Then indent = MAX_INDENT_LEVEL
            AppendIndentAddress ws, indentAddr(indent), _
                "A" & (FIRST_ROW + blockStart - 1) & ":A" & (FIRST_ROW + i - 2), indent

            If i <= mCount Then
                blockStart = i
                curLevel = mLevel(i)
            End If
        End If
    Next i

    For i = 0 To MAX_INDENT_LEVEL
        FlushIndentAddress ws, indentAddr(i), i
    Next i

    ApplyOutlineByGrouping ws
End Sub

Private Sub ApplyOutlineByGrouping(ByVal ws As Worksheet)
    Dim lv As Long
    Dim i As Long
    Dim spanStart As Long
    Dim inSpan As Boolean

    For lv = 1 To MAX_OUTLINE_LEVEL - 1
        Application.StatusBar = "階層の折りたたみを設定中... " & lv & " / " & _
                                (MAX_OUTLINE_LEVEL - 1) & "  (Esc で中断)"
        DoEvents

        inSpan = False
        For i = 1 To mCount
            If mLevel(i) >= lv Then
                If Not inSpan Then
                    spanStart = i
                    inSpan = True
                End If
            ElseIf inSpan Then
                GroupRows ws, spanStart, i - 1
                inSpan = False
            End If
        Next i

        If inSpan Then GroupRows ws, spanStart, mCount
    Next lv
End Sub

Private Sub GroupRows(ByVal ws As Worksheet, ByVal fromIdx As Long, ByVal toIdx As Long)
    If mTestFailOutline Then
        mOutlineErrors = mOutlineErrors + 1
        If Len(mFirstOutlineError) = 0 Then mFirstOutlineError = "テスト注入: Group"
        Exit Sub
    End If
    On Error Resume Next
    ws.Rows((FIRST_ROW + fromIdx - 1) & ":" & (FIRST_ROW + toIdx - 1)).Group
    If Err.Number <> 0 Then
        mOutlineErrors = mOutlineErrors + 1
        If Len(mFirstOutlineError) = 0 Then mFirstOutlineError = Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub AppendIndentAddress(ByVal ws As Worksheet, ByRef buffer As String, _
                                ByVal addr As String, ByVal level As Long)
    If Len(buffer) = 0 Then
        buffer = addr
    ElseIf Len(buffer) + Len(addr) + 1 > MAX_ADDR_LEN Then
        FlushIndentAddress ws, buffer, level
        buffer = addr
    Else
        buffer = buffer & "," & addr
    End If
End Sub

Private Sub FlushIndentAddress(ByVal ws As Worksheet, ByRef buffer As String, ByVal level As Long)
    If Len(buffer) = 0 Then Exit Sub

    If mTestFailOutline Then
        mOutlineErrors = mOutlineErrors + 1
        If Len(mFirstOutlineError) = 0 Then mFirstOutlineError = "テスト注入: IndentLevel"
        buffer = ""
        Exit Sub
    End If

    On Error Resume Next
    ws.Range(buffer).IndentLevel = level
    If Err.Number <> 0 Then
        mOutlineErrors = mOutlineErrors + 1
        If Len(mFirstOutlineError) = 0 Then mFirstOutlineError = Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    buffer = ""
End Sub

'==============================================================
' 表示レベル操作
'==============================================================
Public Sub ExpandAllLevels()
    ShowOutlineLevels MAX_OUTLINE_LEVEL
End Sub

Public Sub CollapseToLevel2()
    ShowOutlineLevels 2
End Sub

Private Sub ShowOutlineLevels(ByVal levels As Long)
    Dim ws As Worksheet
    Dim errNumber As Long
    Dim errDescription As String

    Set ws = FindTreeSheet()
    If ws Is Nothing Then
        MsgBox "先に [2. ツリーを作成] を実行してください。", _
               vbExclamation, "フォルダ階層ツリー"
        Exit Sub
    End If

    On Error Resume Next
    ws.Activate
    ws.Outline.ShowLevels RowLevels:=levels
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber <> 0 Then
        MsgBox "アウトライン表示レベルの変更に失敗しました。" & vbCrLf & _
               "エラー " & errNumber & ": " & errDescription, _
               vbExclamation, "フォルダ階層ツリー"
    End If
End Sub

'==============================================================
' ヘルパー
'==============================================================
Private Function EnsureTreeSheet() As Worksheet
    Dim ws As Worksheet

    Set ws = FindTreeSheet()
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_TREE
    End If

    Set EnsureTreeSheet = ws
End Function

Private Function FindTreeSheet() As Worksheet
    On Error Resume Next
    Set FindTreeSheet = ThisWorkbook.Worksheets(SHEET_TREE)
    On Error GoTo 0
End Function

Private Function ToBool(ByVal v As Variant) As Boolean
    Dim s As String

    If VarType(v) = vbBoolean Then
        ToBool = CBool(v)
        Exit Function
    End If

    s = LCase$(Trim$(CStr(v)))
    Select Case s
        Case "はい", "true", "yes", "y", "1", "on", "含める"
            ToBool = True
        Case Else
            ToBool = False
    End Select
End Function

Private Function FormatSize(ByVal bytes As Double) As String
    Dim units As Variant
    Dim idx As Long
    Dim v As Double

    units = Array("B", "KB", "MB", "GB", "TB", "PB")
    v = bytes
    idx = 0

    Do While v >= 1024 And idx < UBound(units)
        v = v / 1024
        idx = idx + 1
    Loop

    If idx = 0 Then
        FormatSize = Format$(v, "#,##0") & " B"
    Else
        FormatSize = Format$(v, "#,##0.0") & " " & units(idx) & _
                     " (" & Format$(bytes, "#,##0") & " byte)"
    End If
End Function
