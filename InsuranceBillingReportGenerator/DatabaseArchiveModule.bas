Attribute VB_Name = "DatabaseArchiveModule"
Option Explicit

' 再請求済みデータをアーカイブする関数
Public Function ArchiveCompletedData() As Boolean
    On Error GoTo ErrorHandler

    ' 売掛管理表シートの確認
    Dim ws_database As Worksheet
    On Error Resume Next
    Set ws_database = ThisWorkbook.Worksheets("売掛管理表")
    On Error GoTo ErrorHandler

    If ws_database Is Nothing Then
        MsgBox "売掛管理表シートが見つかりません。", vbExclamation, "エラー"
        ArchiveCompletedData = False
        Exit Function
    End If

    ' 最終行を取得
    Dim last_row As Long
    last_row = ws_database.Cells(ws_database.Rows.Count, "A").End(xlUp).Row

    ' アーカイブ対象のデータを確認
    Dim i As Long, archive_count As Long
    Dim status_col As Long

    ' 再請求日の列インデックス（列K）
    Dim rebilling_date_col As Long
    rebilling_date_col = 11

    ' 備考（ステータス）列
    status_col = 16

    archive_count = 0

    ' 各行をチェック
    For i = 2 To last_row ' ヘッダー行をスキップ
        ' 再請求日がある（再請求済み）かつ ステータスが空欄のデータをアーカイブ
        If Not IsEmpty(ws_database.Cells(i, rebilling_date_col).Value) And _
           IsEmpty(ws_database.Cells(i, status_col).Value) Then

            ' ステータスを「完了」に設定
            ws_database.Cells(i, status_col).Value = "完了"

            ' セルの書式設定（背景色を変更）
            ws_database.Cells(i, status_col).Interior.ColorIndex = 35 ' 薄い緑色

            archive_count = archive_count + 1
        End If
    Next i

    If archive_count > 0 Then
        MsgBox archive_count & " 件のデータをアーカイブしました。", vbInformation, "完了"
    Else
        MsgBox "アーカイブ対象のデータがありませんでした。", vbInformation, "完了"
    End If

    ArchiveCompletedData = True
    Exit Function

ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in ArchiveCompletedData"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="

    MsgBox "データのアーカイブ中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
    ArchiveCompletedData = False
End Function
' 未完了データを次期売掛管理表ファイルに転記する関数
Public Function TransferIncompleteData(target_workbook_path As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' 売掛管理表シートの確認
    Dim ws_database As Worksheet
    On Error Resume Next
    Set ws_database = ThisWorkbook.Worksheets("売掛管理表")
    On Error GoTo ErrorHandler
    
    If ws_database Is Nothing Then
        MsgBox "売掛管理表シートが見つかりません。", vbExclamation, "エラー"
        TransferIncompleteData = False
        Exit Function
    End If
    
    ' 転記先のワークブックを開く
    Dim target_wb As Workbook
    Dim target_ws As Worksheet
    Dim file_system As Object
    Dim create_new As Boolean
    
    Set file_system = CreateObject("Scripting.FileSystemObject")
    
    create_new = False
    
    ' ファイルが存在するか確認
    If Not file_system.FileExists(target_workbook_path) Then
        ' 新しいワークブックを作成
        Set target_wb = Workbooks.Add
        create_new = True
    Else
        ' 既存のワークブックを開く
        Set target_wb = Workbooks.Open(target_workbook_path)
    End If
    
    ' 転記先の売掛管理表シートを確認
    On Error Resume Next
    Set target_ws = target_wb.Worksheets("売掛管理表")
    On Error GoTo ErrorHandler
    
    If target_ws Is Nothing Then
        ' 売掛管理表シートを作成
        If Not DatabaseSheetModule.CreateDatabaseSheet(target_wb) Then
            MsgBox "転記先ワークブックに売掛管理表シートを作成できませんでした。", vbCritical, "エラー"
            target_wb.Close SaveChanges:=False
            TransferIncompleteData = False
            Exit Function
        End If
        Set target_ws = target_wb.Worksheets("売掛管理表")
    End If
    
    ' 最終行を取得
    Dim src_last_row As Long, target_last_row As Long
    src_last_row = ws_database.Cells(ws_database.Rows.Count, "A").End(xlUp).Row
    target_last_row = target_ws.Cells(target_ws.Rows.Count, "A").End(xlUp).Row
    
    ' 未完了データを転記
    Dim i As Long, j As Long, transfer_count As Long
    Dim src_last_col As Long
    Dim status_col As Long
    
    src_last_col = ws_database.Cells(1, ws_database.Columns.Count).End(xlToLeft).Column
    status_col = 16 ' 備考（ステータス）列
    
    transfer_count = 0
    
    ' 各行をチェック
    For i = 2 To src_last_row ' ヘッダー行をスキップ
        ' ステータスが空欄または「完了」以外のデータを転記
        If IsEmpty(ws_database.Cells(i, status_col).Value) Or _
           ws_database.Cells(i, status_col).Value <> "完了" Then
           
            ' 転記先の行インデックス
            Dim target_row As Long
            target_row = target_last_row + 1 + transfer_count
            
            ' IDを設定
            target_ws.Cells(target_row, 1).Value = target_row - 1
            
            ' 残りのデータをコピー
            For j = 2 To src_last_col
                target_ws.Cells(target_row, j).Value = ws_database.Cells(i, j).Value
            Next j
            
            transfer_count = transfer_count + 1
        End If
    Next i
    
    ' 売掛管理表の書式を整える
    FormatTransferredDatabase target_ws
    
    ' ファイルを保存
    If create_new Then
        ' 新規ファイルの場合、名前を付けて保存
        target_wb.SaveAs Filename:=target_workbook_path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Else
        ' 既存ファイルの場合、上書き保存
        target_wb.Save
    End If
    
    ' ワークブックを閉じる
    target_wb.Close SaveChanges:=False
    
    If transfer_count > 0 Then
        MsgBox transfer_count & " 件の未完了データを次期売掛管理表ファイルに転記しました。" & vbCrLf & _
               "保存先: " & target_workbook_path, vbInformation, "完了"
    Else
        MsgBox "転記対象の未完了データがありませんでした。", vbInformation, "完了"
    End If
    
    TransferIncompleteData = True
    Exit Function
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in TransferIncompleteData"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    ' ワークブックが開かれていたら閉じる
    On Error Resume Next
    If Not target_wb Is Nothing Then
        target_wb.Close SaveChanges:=False
    End If
    On Error GoTo 0
    
    MsgBox "データの転記中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
    TransferIncompleteData = False
End Function

' 転記先売掛管理表の書式を整える関数
Private Sub FormatTransferredDatabase(ws As Worksheet)
    On Error Resume Next
    
    Dim last_row As Long, last_col As Long
    
    ' 最終行・列を取得
    last_row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    last_col = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' 列幅を自動調整
    ws.Columns("A:" & Chr(64 + last_col)).AutoFit
    
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

' 半期データ管理を行う関数
Public Function ManageHalfYearData() As Boolean
    On Error GoTo ErrorHandler
    
    ' 期間の選択
    Dim year_val As String, period_type As String
    Dim target_workbook_path As String
    
    year_val = InputBox("対象年度を入力してください（例：2025）", "半期データ管理", Year(Date))
    If year_val = "" Then Exit Function
    
    period_type = ""
    Do While period_type <> "上期" And period_type <> "下期"
        period_type = InputBox("対象期間を入力してください（上期/下期）", "半期データ管理", "上期")
        If period_type = "" Then Exit Function
    Loop
    
    ' 保存先の選択
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "次期売掛管理表ファイルの保存先を選択"
        .InitialFileName = "保険請求売掛管理表_" & year_val & "_" & period_type & ".xlsm"
        .FilterIndex = 2  ' Excelマクロ有効ブック
        If .Show = -1 Then
            target_workbook_path = .SelectedItems(1)
        Else
            Exit Function
        End If
    End With
    
    ' 完了済みデータをアーカイブ
    If Not ArchiveCompletedData() Then
        MsgBox "データのアーカイブに失敗しました。処理を中止します。", vbCritical, "エラー"
        ManageHalfYearData = False
        Exit Function
    End If
    
    ' 未完了データを次期売掛管理表ファイルに転記
    If Not TransferIncompleteData(target_workbook_path) Then
        MsgBox "データの転記に失敗しました。", vbCritical, "エラー"
        ManageHalfYearData = False
        Exit Function
    End If
    
    ManageHalfYearData = True
    Exit Function
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in ManageHalfYearData"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    MsgBox "半期データ管理中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
    ManageHalfYearData = False
End Function
