Attribute VB_Name = "modCSVMerge"
Option Explicit

'==============================================================
' CSVMerge - 複数CSVの一括結合・出力マクロ
'
' 指定フォルダ内のCSVファイルをすべて読み込み、
' 1枚のシート（および1つのCSVファイル）に結合して出力します。
'
' 作者: 天電輝合同会社 (TENDENKO)
' ライセンス: MIT License
'
' ---- 使い方 ----
' 1. リボンまたはAlt+F8から「Main_結合実行」を実行
' 2. フォルダ選択ダイアログで、結合したいCSVが入っているフォルダを選ぶ
' 3. 「結合結果」シートに全件が結合され、同じフォルダに
'    「結合結果_yyyymmdd_hhmmss.csv」が出力される
'
' ---- カスタマイズポイント（下の設定セクションを書き換えるだけ） ----
' ・INPUT_ENCODING    : 読み込むCSVの文字コード
' ・OUTPUT_ENCODING   : 出力するCSVの文字コード
' ・HAS_HEADER        : 1行目がヘッダー行かどうか
' ・TARGET_EXTENSION  : 対象とする拡張子
' ・OUTPUT_SHEET_NAME : 結合結果を書き出すシート名
'
' ---- 安全対策 ----
' 1行目（ヘッダー）の列数を基準に、以降すべての行の列数をチェックします。
' 列数が基準と一致しない行があった場合、文字コードの設定間違いによる
' 文字化けの疑いとして「ログ」シートに警告を出力します。
'==============================================================

'---------------- 設定セクション（ここだけ書き換えればOK） ----------------

' 文字コード。"shift-jis" か "utf-8" を指定（大文字小文字どちらでも可）
Private Const INPUT_ENCODING As String = "shift-jis"
Private Const OUTPUT_ENCODING As String = "shift-jis"

' CSVの1行目がヘッダー（項目名）かどうか。
' True の場合、2ファイル目以降のヘッダー行は読み飛ばします。
Private Const HAS_HEADER As Boolean = True

' 結合対象とする拡張子（小文字で指定）
Private Const TARGET_EXTENSION As String = "csv"

' 区切り文字（通常は "," のままでOK。タブ区切りなら vbTab）
Private Const DELIMITER As String = ","

' 結合結果を書き出すシート名
Private Const OUTPUT_SHEET_NAME As String = "結合結果"

' ログ（エラー・スキップ・列数不一致の警告）を書き出すシート名
Private Const LOG_SHEET_NAME As String = "ログ"

'---------------------------------------------------------------------

'==============================================================
' Main_結合実行
'   エントリーポイント。ボタンやAlt+F8から呼び出す。
'==============================================================
Sub Main_結合実行()

    Dim targetFolder As String
    Dim fileList() As String
    Dim fileCount As Long

    On Error GoTo ErrHandler

    ' --- フォルダ選択 ---
    targetFolder = PickFolder()
    If targetFolder = "" Then
        MsgBox "フォルダが選択されませんでした。処理を中止します。", vbInformation
        Exit Sub
    End If

    ' --- 対象CSV一覧を取得 ---
    fileCount = GetCsvFileList(targetFolder, fileList)
    If fileCount = 0 Then
        MsgBox "指定フォルダに ." & TARGET_EXTENSION & " ファイルが見つかりませんでした。" & vbCrLf & _
               "フォルダ: " & targetFolder, vbExclamation
        Exit Sub
    End If

    If MsgBox(fileCount & " 件のCSVファイルが見つかりました。結合を開始しますか？", _
              vbYesNo + vbQuestion, "確認") <> vbYes Then
        Exit Sub
    End If

    ' --- 結合処理 ---
    Dim mergedData As Collection
    Dim warningCount As Long
    warningCount = 0
    Set mergedData = MergeCsvFiles(targetFolder, fileList, fileCount, warningCount)

    ' --- シートへ書き出し ---
    WriteToSheet mergedData

    ' --- CSVファイルとして出力 ---
    Dim outputPath As String
    outputPath = ExportMergedCsv(targetFolder, mergedData)

    Application.StatusBar = False

    ' --- 完了メッセージ（列数不一致があった場合は警告を併記） ---
    Dim msg As String
    msg = "結合が完了しました。" & vbCrLf & _
          "対象ファイル数: " & fileCount & vbCrLf & _
          "出力行数: " & (mergedData.Count - IIf(HAS_HEADER, 1, 0)) & vbCrLf & vbCrLf & _
          "結果は「" & OUTPUT_SHEET_NAME & "」シートと以下のファイルに出力しました。" & vbCrLf & _
          outputPath

    If warningCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "【要確認】" & warningCount & " 件、1行目と列数が異なる行を検出しました。" & vbCrLf & _
              "文字コードの設定（INPUT_ENCODING）がCSVの実際の文字コードと" & vbCrLf & _
              "合っていない可能性があります。詳細は「" & LOG_SHEET_NAME & "」シートをご確認ください。"
        MsgBox msg, vbExclamation, "完了（要確認あり）"
    Else
        MsgBox msg, vbInformation, "完了"
    End If

    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "予期しないエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "内容: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================
' Main_結果シート削除
'   「結合結果」「ログ」シートを削除する。
'   配布・再テスト前のクリーンアップ用のエントリーポイント。
'==============================================================
Sub Main_結果シート削除()

    Dim targets As Variant
    targets = Array(OUTPUT_SHEET_NAME, LOG_SHEET_NAME)

    Dim existCount As Long
    existCount = 0

    Dim i As Long
    For i = LBound(targets) To UBound(targets)
        If SheetExists(CStr(targets(i))) Then
            existCount = existCount + 1
        End If
    Next i

    If existCount = 0 Then
        MsgBox "「" & OUTPUT_SHEET_NAME & "」「" & LOG_SHEET_NAME & "」シートはどちらも存在しません。", _
               vbInformation, "削除不要"
        Exit Sub
    End If

    ' 削除すると全シートがなくなってしまう場合は中止（Excelは最低1シート必要）
    If ThisWorkbook.Worksheets.Count - existCount < 1 Then
        MsgBox "削除すると全てのシートがなくなってしまうため中止しました。" & vbCrLf & _
               "操作用のシートなどを別に1つ残してから実行してください。", vbCritical, "中止"
        Exit Sub
    End If

    If MsgBox("「" & OUTPUT_SHEET_NAME & "」「" & LOG_SHEET_NAME & "」シートを削除します。" & vbCrLf & _
              "元に戻せません。よろしいですか？", vbYesNo + vbExclamation, "確認") <> vbYes Then
        Exit Sub
    End If

    Application.DisplayAlerts = False

    For i = LBound(targets) To UBound(targets)
        If SheetExists(CStr(targets(i))) Then
            ThisWorkbook.Worksheets(CStr(targets(i))).Delete
        End If
    Next i

    Application.DisplayAlerts = True

    MsgBox "削除が完了しました。", vbInformation, "完了"

End Sub

'==============================================================
' 指定シートが存在するかどうかを判定する
'==============================================================
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

'==============================================================
' フォルダ選択ダイアログを表示し、選択されたパスを返す
'==============================================================
Private Function PickFolder() As String

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    fd.Title = "結合したいCSVが入っているフォルダを選択してください"
    fd.AllowMultiSelect = False

    ' 初期表示フォルダ：このブック（CSVMerge.xlsm）が置かれているフォルダを開く
    If ThisWorkbook.Path <> "" Then
        fd.InitialFileName = ThisWorkbook.Path & "\"
    End If

    If fd.Show = -1 Then
        PickFolder = fd.SelectedItems(1)
    Else
        PickFolder = ""
    End If

End Function

'==============================================================
' 指定フォルダ内の対象拡張子ファイル一覧を取得
' 戻り値: 件数（fileList() に絶対パスを格納）
'==============================================================
Private Function GetCsvFileList(ByVal folderPath As String, ByRef fileList() As String) As Long

    Dim fileName As String
    Dim n As Long
    Dim tempList() As String
    ReDim tempList(1000) ' 初期サイズ。足りなければ後で拡張

    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    fileName = Dir(folderPath & "*." & TARGET_EXTENSION)
    n = 0
    Do While fileName <> ""
        If n > UBound(tempList) Then
            ReDim Preserve tempList(UBound(tempList) + 1000)
        End If
        tempList(n) = folderPath & fileName
        n = n + 1
        fileName = Dir()
    Loop

    If n = 0 Then
        GetCsvFileList = 0
        Exit Function
    End If

    ReDim fileList(n - 1)
    Dim i As Long
    For i = 0 To n - 1
        fileList(i) = tempList(i)
    Next i

    GetCsvFileList = n

End Function

'==============================================================
' 複数CSVファイルを読み込み、1つのCollectionに結合する
' （Collectionの各要素 = 1行分のデータを表す配列）
'
' warningCount: 列数不一致（文字化けの疑い）を検出した件数を
'               呼び出し元に返す（ByRef）
'==============================================================
Private Function MergeCsvFiles(ByVal folderPath As String, ByRef fileList() As String, _
                                ByVal fileCount As Long, ByRef warningCount As Long) As Collection

    Dim result As New Collection
    Dim i As Long
    Dim headerWritten As Boolean
    headerWritten = False

    Dim logEntries As New Collection

    ' 列数チェックの基準値。-1 = まだ未確定（最初の行で確定する）
    Dim expectedColCount As Long
    expectedColCount = -1

    For i = 0 To fileCount - 1

        Application.StatusBar = "処理中... (" & (i + 1) & "/" & fileCount & ") " & fileList(i)
        DoEvents

        Dim lines() As String
        Dim lineCount As Long
        lineCount = ReadTextFileLines(fileList(i), INPUT_ENCODING, lines)

        If lineCount = 0 Then
            logEntries.Add "スキップ（空ファイルまたは読込エラー）: " & fileList(i)
            GoTo NextFile
        End If

        Dim startRow As Long
        startRow = 0

        If HAS_HEADER Then
            Dim headerFields() As String
            headerFields = Split(lines(0), DELIMITER)

            CheckColumnCount headerFields, expectedColCount, _
                              GetFileName(fileList(i)) & " の1行目(ヘッダー)", _
                              lines(0), logEntries, warningCount

            If Not headerWritten Then
                result.Add headerFields
                headerWritten = True
            End If
            startRow = 1
        End If

        Dim r As Long
        For r = startRow To lineCount - 1
            If Trim(lines(r)) <> "" Then
                Dim rowFields() As String
                rowFields = Split(lines(r), DELIMITER)

                CheckColumnCount rowFields, expectedColCount, _
                                  GetFileName(fileList(i)) & " " & (r + 1) & "行目", _
                                  lines(r), logEntries, warningCount

                result.Add rowFields
            End If
        Next r

NextFile:
    Next i

    ' ログ（スキップ・警告）をシートに書き出す
    If logEntries.Count > 0 Then
        WriteLogSheet logEntries, warningCount
    End If

    Set MergeCsvFiles = result

End Function

'==============================================================
' 行の列数が基準（expectedColCount）と一致するかチェックする。
' ・基準が未確定（-1）の場合は、この行の列数を基準として確定する
' ・基準と異なる場合は「文字化けの疑い」としてログに追加し、
'   warningCount をインクリメントする
'==============================================================
Private Sub CheckColumnCount(ByRef fields() As String, ByRef expectedColCount As Long, _
                              ByVal sourceLabel As String, ByVal rawLine As String, _
                              ByRef logEntries As Collection, ByRef warningCount As Long)

    Dim colCount As Long
    colCount = UBound(fields) - LBound(fields) + 1

    If expectedColCount = -1 Then
        expectedColCount = colCount
        Exit Sub
    End If

    If colCount <> expectedColCount Then
        logEntries.Add "【文字化けの疑い】列数不一致: " & sourceLabel & _
                        " - 検出列数 " & colCount & "（基準 " & expectedColCount & " 列との差）" & _
                        " / 元の行: " & rawLine
        warningCount = warningCount + 1
    End If

End Sub

'==============================================================
' フルパスからファイル名のみを取り出す
'==============================================================
Private Function GetFileName(ByVal fullPath As String) As String
    Dim pos As Long
    pos = InStrRev(fullPath, "\")
    If pos > 0 Then
        GetFileName = Mid(fullPath, pos + 1)
    Else
        GetFileName = fullPath
    End If
End Function

'==============================================================
' テキストファイルを指定エンコーディングで読み込み、行配列を返す
' 戻り値: 行数
'==============================================================
Private Function ReadTextFileLines(ByVal filePath As String, ByVal encoding As String, ByRef lines() As String) As Long

    On Error GoTo ErrHandler

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2 ' adTypeText
    stm.Charset = NormalizeCharset(encoding)
    stm.Open
    stm.LoadFromFile filePath

    Dim content As String
    content = stm.ReadText
    stm.Close

    ' 改行コードの統一（CRLF / CR / LF すべてに対応）
    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)

    ' 末尾の空行を除去してから分割
    Do While Len(content) > 0 And Right(content, 1) = vbLf
        content = Left(content, Len(content) - 1)
    Loop

    If Len(content) = 0 Then
        ReadTextFileLines = 0
        Exit Function
    End If

    lines = Split(content, vbLf)
    ReadTextFileLines = UBound(lines) - LBound(lines) + 1

    Exit Function

ErrHandler:
    ReadTextFileLines = 0
End Function

'==============================================================
' "shift-jis" / "utf-8" などの表記ゆれをADODBが認識する名称に変換
'==============================================================
Private Function NormalizeCharset(ByVal encoding As String) As String

    Dim e As String
    e = LCase(Trim(encoding))

    Select Case e
        Case "sjis", "shift-jis", "shift_jis", "shiftjis"
            NormalizeCharset = "shift-jis"
        Case "utf8", "utf-8"
            NormalizeCharset = "utf-8"
        Case "utf8bom", "utf-8-bom"
            NormalizeCharset = "utf-8"
        Case Else
            NormalizeCharset = e
    End Select

End Function

'==============================================================
' 結合データをシートに書き出す
' （Range書き込みは1回にまとめ、速度を確保する）
'==============================================================
Private Sub WriteToSheet(ByRef data As Collection)

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(OUTPUT_SHEET_NAME)
    ws.Cells.Clear

    If data.Count = 0 Then Exit Sub

    ' 列数を最初の行から取得
    Dim colCount As Long
    colCount = UBound(data(1)) - LBound(data(1)) + 1

    Dim outputArray() As Variant
    ReDim outputArray(1 To data.Count, 1 To colCount)

    Dim r As Long, c As Long
    Dim rowData As Variant

    For r = 1 To data.Count
        rowData = data(r)
        For c = 1 To colCount
            If c - 1 <= UBound(rowData) - LBound(rowData) Then
                outputArray(r, c) = rowData(LBound(rowData) + c - 1)
            Else
                outputArray(r, c) = ""
            End If
        Next c
    Next r

    ws.Range(ws.Cells(1, 1), ws.Cells(data.Count, colCount)).Value = outputArray

    If HAS_HEADER Then
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).Interior.Color = RGB(221, 235, 247)
    End If

    ws.Columns.AutoFit
    ws.Activate
    ws.Range("A1").Select

End Sub

'==============================================================
' 結合データを新しいCSVファイルとして出力する
' 戻り値: 出力したファイルのフルパス
'==============================================================
Private Function ExportMergedCsv(ByVal folderPath As String, ByRef data As Collection) As String

    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Dim outputPath As String
    outputPath = folderPath & "結合結果_" & Format(Now, "yyyymmdd_hhnnss") & "." & TARGET_EXTENSION

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' adTypeText
    stm.Charset = NormalizeCharset(OUTPUT_ENCODING)
    stm.Open

    Dim r As Long
    Dim rowData As Variant
    Dim lineText As String

    For r = 1 To data.Count
        rowData = data(r)
        lineText = Join(rowData, DELIMITER)
        stm.WriteText lineText, 1 ' adWriteLine
    Next r

    stm.SaveToFile outputPath, 2 ' adSaveCreateOverWrite
    stm.Close

    ExportMergedCsv = outputPath

End Function

'==============================================================
' ログシートへの書き出し（スキップ・エラー・列数不一致があった場合のみ作成）
'==============================================================
Private Sub WriteLogSheet(ByRef logEntries As Collection, ByVal warningCount As Long)

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(LOG_SHEET_NAME)
    ws.Cells.Clear

    ws.Cells(1, 1).Value = "実行日時"
    ws.Cells(1, 2).Value = Format(Now, "yyyy/mm/dd hh:nn:ss")
    ws.Cells(1, 1).Font.Bold = True

    ws.Cells(2, 1).Value = "列数不一致（文字化けの疑い）件数"
    ws.Cells(2, 2).Value = warningCount
    If warningCount > 0 Then
        ws.Cells(2, 1).Font.Color = RGB(192, 0, 0)
        ws.Cells(2, 2).Font.Color = RGB(192, 0, 0)
        ws.Cells(2, 2).Font.Bold = True
    End If

    ws.Cells(4, 1).Value = "詳細"
    ws.Cells(4, 1).Font.Bold = True

    Dim i As Long
    For i = 1 To logEntries.Count
        ws.Cells(4 + i, 1).Value = logEntries(i)
        If InStr(logEntries(i), "文字化けの疑い") > 0 Then
            ws.Cells(4 + i, 1).Font.Color = RGB(192, 0, 0)
        End If
    Next i

    ws.Columns("A").ColumnWidth = 120

End Sub

'==============================================================
' シートが存在すれば取得、なければ新規作成して返す
'==============================================================
Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set GetOrCreateSheet = ws

End Function
