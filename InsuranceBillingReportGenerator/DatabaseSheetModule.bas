Attribute VB_Name = "DatabaseSheetModule"
Option Explicit

' データベース形式のシート3を作成・更新する関数
Public Function CreateDatabaseSheet(ByVal wb As Workbook) As Boolean
    On Error GoTo ErrorHandler
    
    ' データベースシートが存在するか確認し、存在しない場合は作成
    Dim ws_database As Worksheet
    On Error Resume Next
    Set ws_database = wb.Worksheets("データベース")
    On Error GoTo ErrorHandler
    
    If ws_database Is Nothing Then
        ' データベースシートを作成
        Set ws_database = wb.Worksheets.Add(After:=wb.Worksheets("まとめ"))
        ws_database.Name = "データベース"
        
        ' ヘッダーの設定
        With ws_database
            .Range("A1").Value = "ID"
            .Range("B1").Value = "請求先"
            .Range("C1").Value = "区分"
            .Range("D1").Value = "患者名"
            .Range("E1").Value = "調剤年月"
            .Range("F1").Value = "医療機関"
            .Range("G1").Value = "金額"
            .Range("H1").Value = "備考"
            
            ' ヘッダー行の書式設定
            .Range("A1:H1").Font.Bold = True
            .Range("A1:H1").Interior.ColorIndex = 15
            .Range("A1:H1").Borders.LineStyle = xlContinuous
            
            ' 列幅の自動調整
            .Columns("A:H").AutoFit
            
            ' フィルターの追加
            .Range("A1:H1").AutoFilter
        End With
    End If
    
    ' 詳細シートからデータを集計
    PopulateDatabaseFromDetails wb
    
    CreateDatabaseSheet = True
    Exit Function
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in CreateDatabaseSheet"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    MsgBox "データベースシートの作成中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
    CreateDatabaseSheet = False
End Function

' 詳細シートからデータを収集してデータベースシートに入力する関数
Private Sub PopulateDatabaseFromDetails(ByVal wb As Workbook)
    On Error GoTo ErrorHandler
    
    Dim ws_database As Worksheet
    Dim i As Long
    Dim current_row As Long
    Dim last_row As Long
    
    ' データベースシートを取得
    Set ws_database = wb.Worksheets("データベース")
    
    ' 既存のデータをクリア（ヘッダー行を除く）
    If ws_database.Range("A2").Value <> "" Then
        last_row = ws_database.Cells(ws_database.Rows.Count, "A").End(xlUp).Row
        If last_row > 1 Then
            ws_database.Range("A2:H" & last_row).Clear
        End If
    End If
    
    current_row = 2  ' データは2行目から開始
    
    ' 全ての月シートを処理
    Dim sheet_index As Long
    For sheet_index = 1 To wb.Worksheets.Count
        Dim ws As Worksheet
        Set ws = wb.Worksheets(sheet_index)
        
        ' 月のシート（丸数字のシート名）のみを処理
        If IsNumeric(ws.Index) Or ws.Name Like "??" Or InStr("①②③④⑤⑥⑦⑧⑨⑩⑪⑫", ws.Name) > 0 Then
            ' シートからデータを収集
            current_row = CollectDataFromSheet(ws, ws_database, current_row)
        End If
    Next sheet_index
    
    ' 列幅の自動調整
    ws_database.Columns("A:H").AutoFit
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in PopulateDatabaseFromDetails"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    MsgBox "データベースへのデータ入力中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
End Sub

' 指定されたシートからデータを収集する関数
Private Function CollectDataFromSheet(ByVal ws As Worksheet, ByVal ws_database As Worksheet, ByVal start_row As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim last_row As Long
    Dim i As Long
    Dim current_row As Long
    Dim all_category_rows As Object
    Dim cat_key As Variant
    Dim cat_key_2 As Variant
    Dim row_start As Long, row_end As Long
    
    current_row = start_row
    
    ' カテゴリーの開始行を取得
    Set all_category_rows = UtilityModule.GetMarkedCategoryRows(ws)
    
    ' 各カテゴリーのデータを処理
    For Each cat_key In all_category_rows.Keys
        row_start = all_category_rows(cat_key)
        
        ' カテゴリーの終了行を推定（次のカテゴリーの開始行の前まで）
        row_end = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        For Each cat_key_2 In all_category_rows.Keys
            If all_category_rows(cat_key_2) > row_start And all_category_rows(cat_key_2) < row_end Then
                row_end = all_category_rows(cat_key_2) - 1
            End If
        Next cat_key_2
        
        ' このカテゴリーの区分を決定
        Dim category As String
        If InStr(cat_key, "再請求") > 0 Then
            category = "再請求"
        ElseIf InStr(cat_key, "遅請") > 0 Then
            category = "遅請求"
        ElseIf InStr(cat_key, "返戻") > 0 Then
            category = "返戻"
        ElseIf InStr(cat_key, "減点") > 0 Or InStr(cat_key, "査定") > 0 Then
            category = "減点"
        ElseIf InStr(cat_key, "未請求") > 0 Then
            category = "未請求"
        Else
            category = "その他"
        End If
        
        ' このカテゴリーの請求先を決定
        Dim billing_destination As String
        If InStr(cat_key, "社保") > 0 Then
            billing_destination = "社保"
        ElseIf InStr(cat_key, "国保") > 0 Then
            billing_destination = "国保"
        Else
            billing_destination = "その他"
        End If
        
        ' このカテゴリーのデータを処理
        For i = row_start + 1 To row_end
            ' 空行はスキップ
            If ws.Cells(i, 1).Value <> "" Then
                ' IDを生成
                ws_database.Cells(current_row, 1).Value = current_row - 1  ' 単純な連番ID
                
                ' 請求先
                ws_database.Cells(current_row, 2).Value = billing_destination
                
                ' 区分
                ws_database.Cells(current_row, 3).Value = category
                
                ' 患者名
                ws_database.Cells(current_row, 4).Value = ws.Cells(i, 4).Value
                
                ' 調剤年月
                ws_database.Cells(current_row, 5).Value = ws.Cells(i, 5).Value
                
                ' 医療機関
                ws_database.Cells(current_row, 6).Value = ws.Cells(i, 6).Value
                
                ' 金額
                ws_database.Cells(current_row, 7).Value = ws.Cells(i, 10).Value
                If IsNumeric(ws_database.Cells(current_row, 7).Value) Then
                    ws_database.Cells(current_row, 7).NumberFormat = "#,##0"
                End If
                
                ' 備考
                ws_database.Cells(current_row, 8).Value = ""  ' 備考欄は空白として開始
                
                current_row = current_row + 1
            End If
        Next i
    Next cat_key
    
    CollectDataFromSheet = current_row
    Exit Function
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in CollectDataFromSheet"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "Sheet Name: " & ws.Name
    Debug.Print "=================================="
    
    MsgBox "シートからのデータ収集中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "シート名: " & ws.Name, _
           vbCritical, "エラー"
           
    CollectDataFromSheet = current_row
End Function

' データベースシートを更新する関数（既存のデータベースシートがある場合）
Public Sub UpdateDatabaseSheet()
    On Error GoTo ErrorHandler
    
    ' データベースシートが存在するか確認
    Dim ws_database As Worksheet
    On Error Resume Next
    Set ws_database = ThisWorkbook.Worksheets("データベース")
    On Error GoTo ErrorHandler
    
    If ws_database Is Nothing Then
        ' データベースシートが存在しない場合は新規作成
        If CreateDatabaseSheet(ThisWorkbook) Then
            MsgBox "データベースシートを新規作成しました。", vbInformation, "完了"
        Else
            MsgBox "データベースシートの作成に失敗しました。", vbCritical, "エラー"
        End If
    Else
        ' 既存のデータベースシートを更新
        PopulateDatabaseFromDetails ThisWorkbook
        MsgBox "データベースシートを更新しました。", vbInformation, "完了"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in UpdateDatabaseSheet"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    MsgBox "データベースシートの更新中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
End Sub

' CSVファイル処理後にデータベースシートを更新するサブルーチン
Public Sub UpdateDatabaseAfterProcessing()
    On Error Resume Next
    
    ' 保存先フォルダの取得
    Dim save_folder As String
    save_folder = ThisWorkbook.Sheets(1).Range("B3").Value
    
    If save_folder <> "" Then
        ' 現在のワークブックのデータベースシートを更新
        UpdateDatabaseSheet
    End If
End Sub
