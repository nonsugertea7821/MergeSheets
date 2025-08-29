Option Explicit

'===== 定数群 =====

'===== 処理仕様  =====
Private Const RESULTS_SHEET As String = "Results"    ' 統合結果を記録するシート名
Private Const DATE_FMT As String = "yyyymmdd"        ' 日付フォルダ作成用フォーマット
Private Const TIME_FMT As String = "hhnnss"          ' ファイル名に付与する時間フォーマット
Private Const READ_ONLY As Boolean = True            ' 統合元は読み取り専用で開く
Private Const TMP_SUFFIX As String = "_tmp"          ' 上書き時に仮シート名に付与する接尾辞

'===== エラーメッセージ =====
Private Const ERR_LOCK As String =  "統合先ブックが開かれているか、ロックファイルが存在します。" & vbCrLf & _
                                    "対象ブックを閉じ、再実行してください。"
Private Const ERR_DUPLICATE As String = "統合中にシート名の重複が発生しました。" & vbCrLf & _
                                        "上書き設定を確認してください。"
Private Const ERR_NO_FILE As String =   "統合対象ブックが存在しません。" & vbCrLf & _
                                        "パスを確認してください。"
Private Const ERR_UNEXPECTED As String = "統合処理中に予期せぬエラーが発生しました。"

'===== 設定セル座標 =====
Private Const CELL_TARGET_BOOK As String = "B4"   ' 統合先ブック
Private Const CELL_ROOT_FOLDER As String = "B5"   ' 統合元ルートフォルダ
Private Const CELL_OVERWRITE As String = "B6"     ' 同名シート上書き許可
Private Const CELL_IGNORE_LOCK As String = "B7"   ' ロック無視

'===== エラー番号定義 =====
Private Enum MergeError
    ERR_NO_TARGET_FILE = vbObjectError + 100   ' 統合対象ブックが存在しない
    ERR_LOCK_FILE = vbObjectError + 101        ' 統合先ロック存在
    ERR_SHEET_DUPLICATE = vbObjectError + 102  ' シート重複（上書き不可）
End Enum

'===== 設定構造体 =====
Private Type Config
    TargetBook As String     ' 統合先ブックのフルパス
    RootFolder As String     ' 統合元ルートフォルダ
    Overwrite As Boolean     ' 同名シート上書き許可フラグ
    IgnoreLock As Boolean    ' ロック無視フラグ
End Type


'===== メイン処理 =====
Sub MergeSheets()
    ' 統合全体の制御
    Dim cfg As Config
    Dim wsResult As Worksheet
    Dim wbTarget As Workbook
    Dim mergedFilePath As String
    Dim fso As Object, dateFolder As String
    Dim userMsg As String
    
    ' Excel画面更新・計算停止で処理高速化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler
    
    ' 1. 設定取得と結果シート準備
    cfg = ReadConfig()                        ' 設定をシートから取得
    Set wsResult = PrepareResultSheet()       ' 統合結果シートを用意
    CheckLock cfg                             ' ロックファイルがあれば処理中断または削除
    
    ' ファイル操作用オブジェクト作成
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 2. 日付フォルダ作成（マージファイル格納用）
    dateFolder = fso.GetParentFolderName(cfg.TargetBook) & "\" & Format(Date, DATE_FMT)
    If Not fso.FolderExists(dateFolder) Then fso.CreateFolder dateFolder
    
    ' 3. 統合先ブックをコピーして新規マージファイル作成
    mergedFilePath = dateFolder & "\" & fso.GetBaseName(cfg.TargetBook) & "_merged_" & Format(Now, TIME_FMT) & ".xlsx"
    fso.CopyFile cfg.TargetBook, mergedFilePath, True
    Set wbTarget = Workbooks.Open(mergedFilePath)
    
    ' 4. 統合元フォルダを再帰的に検索してシート統合
    MergeFolder cfg.RootFolder, wbTarget, cfg, wsResult
    
    ' 5. マージ完了後、統合先を保存して閉じる
    wbTarget.Close SaveChanges:=True
    MsgBox "統合完了: " & mergedFilePath
    
Cleanup:
    ' Excelの設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrHandler:
    ' エラー番号に応じてユーザー向けメッセージを作成
    Select Case Err.Number
        Case MergeError.ERR_LOCK_FILE
            userMsg = ERR_LOCK & vbCrLf & "詳細: " & Err.Description
        Case MergeError.ERR_SHEET_DUPLICATE
            userMsg = ERR_DUPLICATE & vbCrLf & "シート: " & Err.Description
        Case MergeError.ERR_NO_TARGET_FILE
            userMsg = ERR_NO_FILE & vbCrLf & "詳細: " & Err.Description
        Case Else
            userMsg = ERR_UNEXPECTED & vbCrLf & "詳細: " & Err.Description
    End Select
    
    ' 開いている統合先ブックを閉じ、作成済みマージファイルを削除
    If Not wbTarget Is Nothing Then wbTarget.Close SaveChanges:=False
    If mergedFilePath <> "" Then
        On Error Resume Next
        If fso.FileExists(mergedFilePath) Then fso.DeleteFile mergedFilePath, True
        On Error GoTo 0
    End If
    
    MsgBox userMsg, vbCritical, "統合エラー"
    Resume Cleanup
End Sub

'===== 設定読み込み =====
Private Function ReadConfig() As Config
    ' シート「MergeSheets」から統合設定を読み込む
    Dim cfg As Config, ws As Worksheet
    Set ws = ThisWorkbook.Sheets("MergeSheets")
    
    cfg.TargetBook = CStr(ws.Range(CELL_TARGET_BOOK).Value)
    cfg.RootFolder = CStr(ws.Range(CELL_ROOT_FOLDER).Value)
    cfg.Overwrite = (UCase(CStr(ws.Range(CELL_OVERWRITE).Value)) = "TRUE")
    cfg.IgnoreLock = (UCase(CStr(ws.Range(CELL_IGNORE_LOCK).Value)) = "TRUE")
    
    ' フォルダパス末尾に "\" を付与
    If Right(cfg.RootFolder, 1) <> "\" Then cfg.RootFolder = cfg.RootFolder & "\"
    ReadConfig = cfg
End Function

'===== 結果シート準備 =====
Private Function PrepareResultSheet() As Worksheet
    ' 統合結果を記録するシートを取得・新規作成し初期化
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(RESULTS_SHEET)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = RESULTS_SHEET
    End If
    
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("シート", "状態", "統合元")
    
    Set PrepareResultSheet = ws
End Function

'===== ロック確認 =====
Private Sub CheckLock(cfg As Config)
    ' 統合先ブックのロックファイル確認
    Dim fso As Object, lockFile As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    lockFile = fso.GetParentFolderName(cfg.TargetBook) & "\~$" & fso.GetFileName(cfg.TargetBook)
    
    If fso.FileExists(lockFile) Then
        If cfg.IgnoreLock Then
            ' 無視可能なら削除
            fso.DeleteFile lockFile, True
        Else
            ' 無視不可ならエラー
            Err.Raise MergeError.ERR_LOCK_FILE, , lockFile
        End If
    End If
End Sub

'===== フォルダ統合（再帰） =====
Private Sub MergeFolder(folderPath As String, wbTarget As Workbook, cfg As Config, wsResult As Worksheet)
    ' 指定フォルダ内のすべてのExcelファイルを開き、シートを統合
    Dim fso As Object, f As Object, sf As Object
    Dim wbSrc As Workbook, ws As Worksheet
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' ファイル統合
    For Each f In fso.GetFolder(folderPath).Files
        If LCase(fso.GetExtensionName(f.name)) Like "xls*" Then
            Set wbSrc = Nothing
            On Error GoTo CloseSrc
            Set wbSrc = Workbooks.Open(f.Path, ReadOnly:=READ_ONLY)
            wbSrc.Windows(1).Visible = False  ' 統合中は非表示
            For Each ws In wbSrc.Worksheets
                MergeSheet ws, wbTarget, cfg, wsResult
            Next ws
            
CloseSrc:
            If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
            Set wbSrc = Nothing
            On Error GoTo 0
        End If
    Next f
    
    ' サブフォルダを再帰的に処理
    For Each sf In fso.GetFolder(folderPath).SubFolders
        MergeFolder sf.Path, wbTarget, cfg, wsResult
    Next sf
End Sub

'===== 単一シート統合 =====
Private Sub MergeSheet(wsSrc As Worksheet, wbTarget As Workbook, cfg As Config, wsResult As Worksheet)
    ' 1シート単位で統合先にコピー
    Dim tgt As Worksheet, tmp As Worksheet
    Dim state As String, sheetName As String, relPath As String
    
    sheetName = wsSrc.name
    
    ' 同名シートがあるか確認
    On Error Resume Next
    Set tgt = wbTarget.Sheets(sheetName)
    On Error GoTo 0
    
    If Not tgt Is Nothing Then
        If cfg.Overwrite Then
            ' 上書き時は仮名コピー後に削除して名前戻し
            wsSrc.Copy After:=wbTarget.Sheets(wbTarget.Sheets.Count)
            Set tmp = wbTarget.Sheets(wbTarget.Sheets.Count)
            tmp.name = sheetName & TMP_SUFFIX
            Application.DisplayAlerts = False
            tgt.Delete
            Application.DisplayAlerts = True
            tmp.name = sheetName
            state = "上書き"
        Else
            Err.Raise MergeError.ERR_SHEET_DUPLICATE, , sheetName
        End If
    Else
        ' 新規追加
        wsSrc.Copy After:=wbTarget.Sheets(wbTarget.Sheets.Count)
        wbTarget.Sheets(wbTarget.Sheets.Count).name = sheetName
        state = "追加"
    End If
    
    ' 統合元の相対パスを取得
    relPath = Replace(wsSrc.Parent.FullName, cfg.RootFolder, "")
    UpdateResult wsResult, sheetName, state, relPath
End Sub

'===== 結果更新 =====
Private Sub UpdateResult(ws As Worksheet, name As String, state As String, src As String)
    ' 統合結果シートに追加・更新
    Dim foundRow As Range, newRow As Long
    
    Set foundRow = ws.Columns(1).Find(name, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundRow Is Nothing Then
        ws.Cells(foundRow.Row, 2).Value = state
        ws.Cells(foundRow.Row, 3).Value = src
    Else
        newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(newRow, 1).Value = name
        ws.Cells(newRow, 2).Value = state
        ws.Cells(newRow, 3).Value = src
    End If
End Sub