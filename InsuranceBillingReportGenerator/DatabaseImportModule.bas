Attribute VB_Name = "DatabaseImportModule"
Option Explicit

' CSVファイルからデータベースにデータをインポートする関数
Public Function ImportFromCsvToDatabase(csv_file_path As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' データベースシートの確認
    Dim ws_database As Worksheet
    On Error Resume Next
    Set ws_database = ThisWorkbook.Worksheets("データベース")
    On Error GoTo ErrorHandler
    
    If ws_database Is Nothing Then
        ' データベースシートが存在しない場合は作成
        If Not DatabaseSheetModule.CreateDatabaseSheet(ThisWorkbook) Then
            MsgBox "データベースシートの作成に失敗しました。", vbCritical, "エラー"
            ImportFromCsvToDatabase = False
            Exit Function
        End If
        Set ws_database = ThisWorkbook.Worksheets("データベース")
    End If
    
    ' CSVファイル処理
    Dim file_system As Object, text_stream As Object
    Dim line_text As String, csv_data() As String
    Dim row_data As Variant
    Dim last_row As Long
    Dim i As Long, row_index As Long
    
    Set file_system = CreateObject("Scripting.FileSystemObject")
    
    ' ファイルが存在するか確認
    If Not file_system.FileExists(csv_file_path) Then
        MsgBox "CSVファイルが見つかりません。" & vbCrLf & csv_file_path, vbExclamation, "エラー"
        ImportFromCsvToDatabase = False
        Exit Function
    End If
    
    ' CSVファイルを開く
    Set text_stream = file_system.OpenTextFile(csv_file_path, 1, False, -1) ' 読み取り専用、UTF-8
    
    ' 最終行を取得
    last_row = ws_database.Cells(ws_database.Rows.Count, "A").End(xlUp).Row
    row_index = last_row + 1
    
    ' CSVファイルを行ごとに処理
    Do Until text_stream.AtEndOfStream
        line_text = text_stream.ReadLine
        
        ' 空行をスキップ
        If Trim(line_text) <> "" Then
            ' CSVデータを分割
            csv_data = Split(line_text, ",")
            
            ' ヘッダー行をスキップ
            If csv_data(0) <> "ID" And IsNumeric(csv_data(0)) Then
                ' データをデータベースに追加
                ws_database.Cells(row_index, 1).Value = row_index - 1 ' ID
                
                ' 残りのデータを追加（CSVの列順に応じて調整が必要）
                For i = 1 To UBound(csv_data)
                    If i < ws_database.Columns.Count Then
                        ws_database.Cells(row_index, i + 1).Value = csv_data(i)
                    End If
                Next i
                
                row_index = row_index + 1
            End If
        End If
    Loop
    
    ' ファイルを閉じる
    text_stream.Close
    
    ' データベースの書式を整える
    FormatDatabaseSheet ws_database
    
    MsgBox "CSVファイルからデータベースへのインポートが完了しました。", vbInformation, "完了"
    ImportFromCsvToDatabase = True
    
    Exit Function
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in ImportFromCsvToDatabase"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    MsgBox "CSVファイルのインポート中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
    ImportFromCsvToDatabase = False
End Function

' マクロ有効ファイルからデータベースにデータをインポートする関数
Public Function ImportFromExcelToDatabase(excel_file_path As String, sheet_name As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' データベースシートの確認
    Dim ws_database As Worksheet
    On Error Resume Next
    Set ws_database = ThisWorkbook.Worksheets("データベース")
    On Error GoTo ErrorHandler
    
    If ws_database Is Nothing Then
        ' データベースシートが存在しない場合は作成
        If Not DatabaseSheetModule.CreateDatabaseSheet(ThisWorkbook) Then
            MsgBox "データベースシートの作成に失敗しました。", vbCritical, "エラー"
            ImportFromExcelToDatabase = False
            Exit Function
        End If
        Set ws_database = ThisWorkbook.Worksheets("データベース")
    End If
    
    ' ソースワークブックを開く
    Dim source_wb As Workbook, source_ws As Worksheet
    Dim file_system As Object
    
    Set file_system = CreateObject("Scripting.FileSystemObject")
    
    ' ファイルが存在するか確認
    If Not file_system.FileExists(excel_file_path) Then
        MsgBox "Excelファイルが見つかりません。" & vbCrLf & excel_file_path, vbExclamation, "エラー"
        ImportFromExcelToDatabase = False
        Exit Function
    End If
    
    ' ワークブックを開く
    Set source_wb = Workbooks.Open(excel_file_path, ReadOnly:=True)
    
    ' シートを確認
    On Error Resume Next
    Set source_ws = source_wb.Worksheets(sheet_name)
    If source_ws Is Nothing Then
        ' 指定されたシートが存在しない場合は最初のシートを使用
        Set source_ws = source_wb.Worksheets(1)
    End If
    On Error GoTo ErrorHandler
    
    ' データの範囲を取得
    Dim src_last_row As Long, src_last_col As Long
    Dim db_last_row As Long
    Dim i As Long, j As Long, row_index As Long
    
    src_last_row = source_ws.Cells(source_ws.Rows.Count, "A").End(xlUp).Row
    src_last_col = source_ws.Cells(1, source_ws.Columns.Count).End(xlToLeft).Column
    
    ' データベースの最終行を取得
    db_last_row = ws_database.Cells(ws_database.Rows.Count, "A").End(xlUp).Row
    row_index = db_last_row + 1
    
    ' データをコピー（ヘッダー行をスキップ）
    For i = 2 To src_last_row ' 2行目からスタート（ヘッダーをスキップ）
        ' IDを設定
        ws_database.Cells(row_index, 1).Value = row_index - 1
        
        ' 残りのデータをコピー
        For j = 2 To src_last_col
            If j <= ws_database.Columns.Count Then
                ws_database.Cells(row_index, j).Value = source_ws.Cells(i, j - 1).Value
            End If
        Next j
        
        row_index = row_index + 1
    Next i
    
    ' ソースワークブックを閉じる
    source_wb.Close SaveChanges:=False
    
    ' データベースの書式を整える
    FormatDatabaseSheet ws_database
    
    MsgBox "Excelファイルからデータベースへのインポートが完了しました。", vbInformation, "完了"
    ImportFromExcelToDatabase = True
    
    Exit Function
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in ImportFromExcelToDatabase"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    ' ワークブックが開かれていたら閉じる
    On Error Resume Next
    If Not source_wb Is Nothing Then
        source_wb.Close SaveChanges:=False
    End If
    On Error GoTo 0
    
    MsgBox "Excelファイルのインポート中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
    ImportFromExcelToDatabase = False
End Function

' データベースシートの書式を整える関数
Private Sub FormatDatabaseSheet(ws As Worksheet)
    On Error Resume Next
    
    Dim last_row As Long, last_col As Long
    
    ' 最終行・列を取得
    last_row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    last_col = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' 列幅を自動調整
    ws.Columns("A:" & Chr(64 + last_col)).AutoFit
    
    ' セルの書式設定
    With ws.Range("A1:" & Chr(64 + last_col) & "1")
        .Font.Bold = True
        .Interior.ColorIndex = 15
        .Borders.LineStyle = xlContinuous
    End With
    
    ' データ範囲の罫線
    If last_row > 1 Then
        ws.Range("A2:" & Chr(64 + last_col) & last_row).Borders.LineStyle = xlContinuous
    End If
    
    ' 金額列の書式設定
    ws.Range("F2:G" & last_row).NumberFormat = "#,##0"
    ws.Range("I2:J" & last_row).NumberFormat = "#,##0"
    
    ' 日付列の書式設定
    ws.Range("D2:D" & last_row).NumberFormat = "yyyy/mm/dd"
    ws.Range("H2:H" & last_row).NumberFormat = "yyyy/mm/dd"
    ws.Range("L2:L" & last_row).NumberFormat = "yyyy/mm/dd"
    ws.Range("N2:N" & last_row).NumberFormat = "yyyy/mm/dd"
End Sub
