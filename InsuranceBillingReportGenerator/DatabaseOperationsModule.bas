Attribute VB_Name = "DatabaseOperationsModule"
Option Explicit

' データベースシートを検索する関数
Public Sub SearchDatabase()
    On Error GoTo ErrorHandler
    
    ' データベースシートが存在するか確認
    Dim ws_database As Worksheet
    On Error Resume Next
    Set ws_database = ThisWorkbook.Worksheets("データベース")
    On Error GoTo ErrorHandler
    
    If ws_database Is Nothing Then
        MsgBox "データベースシートが見つかりません。先にデータベースを作成してください。", vbExclamation, "エラー"
        Exit Sub
    End If
    
    ' 検索フォームを表示
    Dim search_form As New DatabaseSearchForm
    search_form.Show
    
    ' キャンセルされた場合は終了
    If search_form.Cancelled Then
        Exit Sub
    End If
    
    ' 検索条件を取得
    Dim billing_destination As String, category As String
    Dim date_from As String, date_to As String
    Dim amount_from As String, amount_to As String
    Dim search_text As String
    Dim i As Long
    
    ' 新しい日付フィールド用の変数
    Dim billing_date_from As String, billing_date_to As String
    Dim processing_date_from As String, processing_date_to As String
    Dim return_date_from As String, return_date_to As String
    Dim rebilling_date_from As String, rebilling_date_to As String
    
    ' 新しい金額フィールド用の変数
    Dim primary_insurance_from As String, primary_insurance_to As String
    Dim public_insurance_from As String, public_insurance_to As String
    Dim primary_rebilling_from As String, primary_rebilling_to As String
    Dim public_rebilling_from As String, public_rebilling_to As String
    
    ' 機関フィールド用の変数
    Dim billing_institution As String, rebilling_institution As String
    
    billing_destination = search_form.SelectedBillingDestination
    category = search_form.SelectedCategory
    date_from = search_form.DateFrom
    date_to = search_form.DateTo
    amount_from = search_form.AmountFrom
    amount_to = search_form.AmountTo
    search_text = search_form.SearchText
    
    ' 新しい日付フィールドの値を取得
    billing_date_from = search_form.BillingDateFrom
    billing_date_to = search_form.BillingDateTo
    processing_date_from = search_form.ProcessingDateFrom
    processing_date_to = search_form.ProcessingDateTo
    return_date_from = search_form.ReturnDateFrom
    return_date_to = search_form.ReturnDateTo
    rebilling_date_from = search_form.RebillingDateFrom
    rebilling_date_to = search_form.RebillingDateTo
    
    ' 新しい金額フィールドの値を取得
    primary_insurance_from = search_form.PrimaryInsuranceFrom
    primary_insurance_to = search_form.PrimaryInsuranceTo
    public_insurance_from = search_form.PublicInsuranceFrom
    public_insurance_to = search_form.PublicInsuranceTo
    primary_rebilling_from = search_form.PrimaryRebillingFrom
    primary_rebilling_to = search_form.PrimaryRebillingTo
    public_rebilling_from = search_form.PublicRebillingFrom
    public_rebilling_to = search_form.PublicRebillingTo
    
    ' 機関フィールドの値を取得
    billing_institution = search_form.BillingInstitution
    rebilling_institution = search_form.RebillingInstitution
    
    ' フィルターをクリア
    ws_database.AutoFilterMode = False
    
    ' フィルター条件の配列を作成
    Dim criteria(1 To 16) As Variant
    For i = 1 To 16
        criteria(i) = ""
    Next i
    
    ' 検索条件が指定されている列にフィルター条件を設定
    If category <> "" Then
        criteria(2) = category
    End If
    
    ' 日付範囲の処理
    If date_from <> "" Or date_to <> "" Then
        If date_from <> "" And date_to <> "" Then
            criteria(4) = ">=" & date_from & " " & "<=" & date_to
        ElseIf date_from <> "" Then
            criteria(4) = ">=" & date_from
        ElseIf date_to <> "" Then
            criteria(4) = "<=" & date_to
        End If
    End If
    
    ' 患者名の処理
    If search_form.txtPatientName.Value <> "" Then
        criteria(3) = "*" & search_form.txtPatientName.Value & "*"
    End If
    
    ' 金額範囲の処理
    If amount_from <> "" Or amount_to <> "" Then
        If IsNumeric(amount_from) And IsNumeric(amount_to) Then
            criteria(6) = ">=" & amount_from & " " & "<=" & amount_to
        ElseIf IsNumeric(amount_from) Then
            criteria(6) = ">=" & amount_from
        ElseIf IsNumeric(amount_to) Then
            criteria(6) = "<=" & amount_to
        End If
    End If
    
    ' 請求日範囲の処理
    If billing_date_from <> "" Or billing_date_to <> "" Then
        If billing_date_from <> "" And billing_date_to <> "" Then
            criteria(7) = ">=" & billing_date_from & " " & "<=" & billing_date_to
        ElseIf billing_date_from <> "" Then
            criteria(7) = ">=" & billing_date_from
        ElseIf billing_date_to <> "" Then
            criteria(7) = "<=" & billing_date_to
        End If
    End If
    
    ' 処理日範囲の処理
    If processing_date_from <> "" Or processing_date_to <> "" Then
        If processing_date_from <> "" And processing_date_to <> "" Then
            criteria(8) = ">=" & processing_date_from & " " & "<=" & processing_date_to
        ElseIf processing_date_from <> "" Then
            criteria(8) = ">=" & processing_date_from
        ElseIf processing_date_to <> "" Then
            criteria(8) = "<=" & processing_date_to
        End If
    End If
    
    ' 返戻日範囲の処理
    If return_date_from <> "" Or return_date_to <> "" Then
        If return_date_from <> "" And return_date_to <> "" Then
            criteria(9) = ">=" & return_date_from & " " & "<=" & return_date_to
        ElseIf return_date_from <> "" Then
            criteria(9) = ">=" & return_date_from
        ElseIf return_date_to <> "" Then
            criteria(9) = "<=" & return_date_to
        End If
    End If
    
    ' 再請求日範囲の処理
    If rebilling_date_from <> "" Or rebilling_date_to <> "" Then
        If rebilling_date_from <> "" And rebilling_date_to <> "" Then
            criteria(10) = ">=" & rebilling_date_from & " " & "<=" & rebilling_date_to
        ElseIf rebilling_date_from <> "" Then
            criteria(10) = ">=" & rebilling_date_from
        ElseIf rebilling_date_to <> "" Then
            criteria(10) = "<=" & rebilling_date_to
        End If
    End If
    
    ' 主保険請求額範囲の処理
    If primary_insurance_from <> "" Or primary_insurance_to <> "" Then
        If IsNumeric(primary_insurance_from) And IsNumeric(primary_insurance_to) Then
            criteria(11) = ">=" & primary_insurance_from & " " & "<=" & primary_insurance_to
        ElseIf IsNumeric(primary_insurance_from) Then
            criteria(11) = ">=" & primary_insurance_from
        ElseIf IsNumeric(primary_insurance_to) Then
            criteria(11) = "<=" & primary_insurance_to
        End If
    End If
    
    ' 公費請求額範囲の処理
    If public_insurance_from <> "" Or public_insurance_to <> "" Then
        If IsNumeric(public_insurance_from) And IsNumeric(public_insurance_to) Then
            criteria(12) = ">=" & public_insurance_from & " " & "<=" & public_insurance_to
        ElseIf IsNumeric(public_insurance_from) Then
            criteria(12) = ">=" & public_insurance_from
        ElseIf IsNumeric(public_insurance_to) Then
            criteria(12) = "<=" & public_insurance_to
        End If
    End If
    
    ' 主保険再請求額範囲の処理
    If primary_rebilling_from <> "" Or primary_rebilling_to <> "" Then
        If IsNumeric(primary_rebilling_from) And IsNumeric(primary_rebilling_to) Then
            criteria(13) = ">=" & primary_rebilling_from & " " & "<=" & primary_rebilling_to
        ElseIf IsNumeric(primary_rebilling_from) Then
            criteria(13) = ">=" & primary_rebilling_from
        ElseIf IsNumeric(primary_rebilling_to) Then
            criteria(13) = "<=" & primary_rebilling_to
        End If
    End If
    
    ' 公費再請求額範囲の処理
    If public_rebilling_from <> "" Or public_rebilling_to <> "" Then
        If IsNumeric(public_rebilling_from) And IsNumeric(public_rebilling_to) Then
            criteria(14) = ">=" & public_rebilling_from & " " & "<=" & public_rebilling_to
        ElseIf IsNumeric(public_rebilling_from) Then
            criteria(14) = ">=" & public_rebilling_from
        ElseIf IsNumeric(public_rebilling_to) Then
            criteria(14) = "<=" & public_rebilling_to
        End If
    End If
    
    ' 請求先機関の処理
    If billing_institution <> "" Then
        criteria(15) = "*" & billing_institution & "*"
    End If
    
    ' 再請求先機関の処理
    If rebilling_institution <> "" Then
        criteria(16) = "*" & rebilling_institution & "*"
    End If
    
    ' テキスト検索の処理
    If search_text <> "" Then
        ' 複数列で検索（患者名、医療機関、請求先機関、再請求先機関）
        ws_database.Range("A1:P1").AutoFilter Field:=3, Criteria1:="*" & search_text & "*", Operator:=xlOr, _
            Criteria2:=Array("*" & search_text & "*", "*" & search_text & "*", "*" & search_text & "*")
    End If
    
    ' 各列にフィルターを適用
    For i = 1 To 16
        If criteria(i) <> "" Then
            ws_database.Range("A1:P1").AutoFilter Field:=i, Criteria1:=criteria(i)
        End If
    Next i
    
    ' フィルタリング結果があるかチェック
    Dim visible_count As Long
    visible_count = WorksheetFunction.Subtotal(3, ws_database.Range("A:A")) - 1  ' ヘッダー行を除く
    
    ' 結果のメッセージを表示
    MsgBox "検索結果: " & visible_count & " 件のレコードが見つかりました。", vbInformation, "検索完了"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in SearchDatabase"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    MsgBox "データベース検索中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
End Sub

' データベースの内容をCSVファイルにエクスポートする関数
Public Sub ExportDatabaseToCsv()
    On Error GoTo ErrorHandler
    
    ' データベースシートが存在するか確認
    Dim ws_database As Worksheet
    On Error Resume Next
    Set ws_database = ThisWorkbook.Worksheets("データベース")
    On Error GoTo ErrorHandler
    
    If ws_database Is Nothing Then
        MsgBox "データベースシートが見つかりません。先にデータベースを作成してください。", vbExclamation, "エラー"
        Exit Sub
    End If
    
    ' 保存先を選択
    Dim save_path As String
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "CSVファイルの保存先を選択"
        .InitialFileName = "保険請求データベース_" & Format(Date, "yyyymmdd") & ".csv"
        .FilterIndex = 3  ' CSVフィルター
        If .Show = -1 Then
            save_path = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ' データの範囲を取得
    Dim last_row As Long
    last_row = ws_database.Cells(ws_database.Rows.Count, "A").End(xlUp).Row
    
    ' エクスポート用のテンポラリーブックを作成
    Dim temp_wb As Workbook, temp_ws As Worksheet
    Set temp_wb = Workbooks.Add
    Set temp_ws = temp_wb.Sheets(1)
    
    ' データをコピー（フィルターされた表示データのみ）
    ws_database.Range("A1:P" & last_row).SpecialCells(xlCellTypeVisible).Copy
    temp_ws.Range("A1").PasteSpecial xlPasteValues
    
    ' CSV形式で保存
    Application.DisplayAlerts = False
    temp_wb.SaveAs Filename:=save_path, FileFormat:=xlCSV, Local:=True
    temp_wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    MsgBox "データベースを正常にCSVファイルにエクスポートしました。" & vbCrLf & _
           "ファイル: " & save_path, vbInformation, "エクスポート完了"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in ExportDatabaseToCsv"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    If Not temp_wb Is Nothing Then
        temp_wb.Close SaveChanges:=False
    End If
    
    Application.DisplayAlerts = True
    
    MsgBox "CSVエクスポート中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
End Sub

' データベースの集計レポートを作成する関数
Public Sub CreateDatabaseSummaryReport()
    On Error GoTo ErrorHandler
    
    ' データベースシートが存在するか確認
    Dim ws_database As Worksheet
    On Error Resume Next
    Set ws_database = ThisWorkbook.Worksheets("データベース")
    On Error GoTo ErrorHandler
    
    If ws_database Is Nothing Then
        MsgBox "データベースシートが見つかりません。先にデータベースを作成してください。", vbExclamation, "エラー"
        Exit Sub
    End If
    
    ' 集計レポートシートが存在するか確認し、存在しない場合は作成
    Dim ws_report As Worksheet
    On Error Resume Next
    Set ws_report = ThisWorkbook.Worksheets("集計レポート")
    On Error GoTo ErrorHandler
    
    If ws_report Is Nothing Then
        Set ws_report = ThisWorkbook.Worksheets.Add(After:=ws_database)
        ws_report.Name = "集計レポート"
    Else
        ws_report.Cells.Clear
    End If
    
    ' レポートのヘッダーを設定
    With ws_report
        .Range("A1").Value = "保険請求データベース集計レポート"
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "【請求先別集計】"
        .Range("A3").Font.Bold = True
        .Range("A4").Value = "請求先"
        .Range("B4").Value = "金額合計"
        .Range("C4").Value = "主保険請求額合計"
        .Range("D4").Value = "公費請求額合計"
        .Range("E4").Value = "主保険再請求額合計"
        .Range("F4").Value = "公費再請求額合計"
        .Range("G4").Value = "請求先機関"
        .Range("H4").Value = "再請求先機関"
        
        .Range("A8").Value = "【区分別集計】"
        .Range("A8").Font.Bold = True
        .Range("A9").Value = "区分"
        .Range("B9").Value = "金額合計"
        .Range("C9").Value = "主保険請求額合計"
        .Range("D9").Value = "公費請求額合計"
        .Range("E9").Value = "主保険再請求額合計"
        .Range("F9").Value = "公費再請求額合計"
        .Range("G9").Value = "請求先機関"
        .Range("H9").Value = "再請求先機関"
        
        .Range("A14").Value = "【月別集計】"
        .Range("A14").Font.Bold = True
        .Range("A15").Value = "調剤年月"
        .Range("B15").Value = "金額合計"
        .Range("C15").Value = "主保険請求額合計"
        .Range("D15").Value = "公費請求額合計"
        .Range("E15").Value = "主保険再請求額合計"
        .Range("F15").Value = "公費再請求額合計"
        .Range("G15").Value = "請求先機関"
        .Range("H15").Value = "再請求先機関"
    End With
    
    ' データベースシートからデータを集計
    
    ' 請求先別集計
    Dim billing_types As Object
    Set billing_types = CreateObject("Scripting.Dictionary")
    billing_types.Add "社保", Array(0, 0, 0, 0, 0)  ' 金額合計, 主保険請求額, 公費請求額, 主保険再請求額, 公費再請求額
    billing_types.Add "国保", Array(0, 0, 0, 0, 0)
    billing_types.Add "その他", Array(0, 0, 0, 0, 0)
    
    ' 区分別集計
    Dim categories As Object
    Set categories = CreateObject("Scripting.Dictionary")
    categories.Add "未請求", Array(0, 0, 0, 0, 0)  ' 金額合計, 主保険請求額, 公費請求額, 主保険再請求額, 公費再請求額
    categories.Add "返戻", Array(0, 0, 0, 0, 0)
    categories.Add "減点", Array(0, 0, 0, 0, 0)
    categories.Add "再請求", Array(0, 0, 0, 0, 0)
    categories.Add "遅請求", Array(0, 0, 0, 0, 0)
    categories.Add "その他", Array(0, 0, 0, 0, 0)
    
    ' 月別集計
    Dim months As Object
    Set months = CreateObject("Scripting.Dictionary")
    
    ' 請求日別集計
    Dim billing_dates As Object
    Set billing_dates = CreateObject("Scripting.Dictionary")
    
    ' データの範囲を取得
    Dim last_row As Long
    last_row = ws_database.Cells(ws_database.Rows.Count, "A").End(xlUp).Row
    
    ' データを集計
    Dim i As Long
    For i = 2 To last_row  ' ヘッダー行をスキップ
        ' 非表示行はスキップ（フィルターが適用されている場合）
        If ws_database.Rows(i).Hidden = False Then
            ' 請求先別集計
            Dim billing_type As String
            billing_type = ws_database.Cells(i, 2).Value
            If billing_type = "" Then billing_type = "その他"
            
            Dim billing_array As Variant
            billing_array = billing_types(billing_type)
            ' 金額が数値の場合のみ加算
            If IsNumeric(ws_database.Cells(i, 7).Value) Then
                billing_array(0) = billing_array(0) + ws_database.Cells(i, 7).Value  ' 金額を加算
            End If
            
            billing_types(billing_type) = billing_array
            
            ' 区分別集計
            Dim category As String
            category = ws_database.Cells(i, 3).Value
            If category = "" Then category = "その他"
            
            Dim category_array As Variant
            category_array = categories(category)
            ' 金額が数値の場合のみ加算
            If IsNumeric(ws_database.Cells(i, 7).Value) Then
                category_array(0) = category_array(0) + ws_database.Cells(i, 7).Value  ' 金額を加算
            End If
            
            categories(category) = category_array
            
            ' 月別集計
            Dim month_key As String
            month_key = ws_database.Cells(i, 5).Value
            If month_key = "" Then month_key = "不明"
            
            If Not months.Exists(month_key) Then
                months.Add month_key, Array(0, 0, 0, 0, 0)  ' 金額合計, 主保険請求額, 公費請求額, 主保険再請求額, 公費再請求額
            End If
            
            Dim month_array As Variant
            month_array = months(month_key)
            ' 金額が数値の場合のみ加算
            If IsNumeric(ws_database.Cells(i, 7).Value) Then
                month_array(0) = month_array(0) + ws_database.Cells(i, 7).Value  ' 金額を加算
            End If
            
            ' 主保険請求額が数値の場合のみ加算
            If IsNumeric(ws_database.Cells(i, 8).Value) Then
                month_array(1) = month_array(1) + ws_database.Cells(i, 8).Value  ' 主保険請求額を加算
            End If
            
            ' 公費請求額が数値の場合のみ加算
            If IsNumeric(ws_database.Cells(i, 9).Value) Then
                month_array(2) = month_array(2) + ws_database.Cells(i, 9).Value  ' 公費請求額を加算
            End If
            
            ' 主保険再請求額が数値の場合のみ加算
            If IsNumeric(ws_database.Cells(i, 10).Value) Then
                month_array(3) = month_array(3) + ws_database.Cells(i, 10).Value  ' 主保険再請求額を加算
            End If
            
            ' 公費再請求額が数値の場合のみ加算
            If IsNumeric(ws_database.Cells(i, 11).Value) Then
                month_array(4) = month_array(4) + ws_database.Cells(i, 11).Value  ' 公費再請求額を加算
            End If
            
            ' 主保険請求額が数値の場合のみ加算
            If IsNumeric(ws_database.Cells(i, 8).Value) Then
                month_array(2) = month_array(2) + ws_database.Cells(i, 8).Value
            End If
            
            ' 公費請求額が数値の場合のみ加算
            If IsNumeric(ws_database.Cells(i, 9).Value) Then
                month_array(3) = month_array(3) + ws_database.Cells(i, 9).Value
            End If
            
            ' 主保険再請求額が数値の場合のみ加算
            If IsNumeric(ws_database.Cells(i, 10).Value) Then
                month_array(4) = month_array(4) + ws_database.Cells(i, 10).Value
            End If
            
            ' 公費再請求額が数値の場合のみ加算
            If IsNumeric(ws_database.Cells(i, 11).Value) Then
                month_array(5) = month_array(5) + ws_database.Cells(i, 11).Value
            End If
            
            months(month_key) = month_array
            
            ' 請求日別集計
            Dim billing_date_key As String
            billing_date_key = ws_database.Cells(i, 12).Value
            If billing_date_key = "" Then billing_date_key = "不明"
            
            If Not billing_dates.Exists(billing_date_key) Then
                billing_dates.Add billing_date_key, Array(0, 0, 0)  ' 金額合計, 主保険請求額, 公費請求額
            End If
            
            Dim billing_date_array As Variant
            billing_date_array = billing_dates(billing_date_key)
            
            ' 金額が数値の場合のみ加算
            If IsNumeric(ws_database.Cells(i, 7).Value) Then
                billing_date_array(0) = billing_date_array(0) + ws_database.Cells(i, 7).Value  ' 金額を加算
            End If
            
            ' 主保険請求額が数値の場合のみ加算
            If IsNumeric(ws_database.Cells(i, 8).Value) Then
                billing_date_array(1) = billing_date_array(1) + ws_database.Cells(i, 8).Value
            End If
            
            ' 公費請求額が数値の場合のみ加算
            If IsNumeric(ws_database.Cells(i, 9).Value) Then
                billing_date_array(2) = billing_date_array(2) + ws_database.Cells(i, 9).Value
            End If
            
            billing_dates(billing_date_key) = billing_date_array
        End If
    Next i
    
    ' 請求先別集計をレポートに出力
    Dim row_index As Long
    row_index = 5
    
    Dim billing_key As Variant
    For Each billing_key In billing_types.Keys
        ws_report.Cells(row_index, 1).Value = billing_key
        ws_report.Cells(row_index, 2).Value = billing_types(billing_key)(0)
        ws_report.Cells(row_index, 2).NumberFormat = "#,##0"
        ws_report.Cells(row_index, 3).Value = billing_types(billing_key)(1)
        ws_report.Cells(row_index, 3).NumberFormat = "#,##0"
        ws_report.Cells(row_index, 4).Value = billing_types(billing_key)(2)
        ws_report.Cells(row_index, 4).NumberFormat = "#,##0"
        ws_report.Cells(row_index, 5).Value = billing_types(billing_key)(3)
        ws_report.Cells(row_index, 5).NumberFormat = "#,##0"
        ws_report.Cells(row_index, 6).Value = billing_types(billing_key)(4)
        ws_report.Cells(row_index, 6).NumberFormat = "#,##0"
        row_index = row_index + 1
    Next billing_key
    
    ' 合計行を追加
    ws_report.Cells(row_index, 1).Value = "合計"
    ws_report.Cells(row_index, 1).Font.Bold = True
    ws_report.Cells(row_index, 2).Formula = "=SUM(B5:B" & (row_index - 1) & ")"
    ws_report.Cells(row_index, 2).NumberFormat = "#,##0"
    ws_report.Cells(row_index, 3).Formula = "=SUM(C5:C" & (row_index - 1) & ")"
    ws_report.Cells(row_index, 3).NumberFormat = "#,##0"
    ws_report.Cells(row_index, 4).Formula = "=SUM(D5:D" & (row_index - 1) & ")"
    ws_report.Cells(row_index, 4).NumberFormat = "#,##0"
    ws_report.Cells(row_index, 5).Formula = "=SUM(E5:E" & (row_index - 1) & ")"
    ws_report.Cells(row_index, 5).NumberFormat = "#,##0"
    ws_report.Cells(row_index, 6).Formula = "=SUM(F5:F" & (row_index - 1) & ")"
    ws_report.Cells(row_index, 6).NumberFormat = "#,##0"
    
    ' 区分別集計をレポートに出力
    row_index = 10
    
    Dim category_key As Variant
    For Each category_key In categories.Keys
        ws_report.Cells(row_index, 1).Value = category_key
        ws_report.Cells(row_index, 2).Value = categories(category_key)(0)
        ws_report.Cells(row_index, 2).NumberFormat = "#,##0"
        ws_report.Cells(row_index, 3).Value = categories(category_key)(1)
        ws_report.Cells(row_index, 3).NumberFormat = "#,##0"
        ws_report.Cells(row_index, 4).Value = categories(category_key)(2)
        ws_report.Cells(row_index, 4).NumberFormat = "#,##0"
        ws_report.Cells(row_index, 5).Value = categories(category_key)(3)
        ws_report.Cells(row_index, 5).NumberFormat = "#,##0"
        ws_report.Cells(row_index, 6).Value = categories(category_key)(4)
        ws_report.Cells(row_index, 6).NumberFormat = "#,##0"
        row_index = row_index + 1
    Next category_key
    
    ' 合計行を追加
    ws_report.Cells(row_index, 1).Value = "合計"
    ws_report.Cells(row_index, 1).Font.Bold = True
    ws_report.Cells(row_index, 2).Formula = "=SUM(B10:B" & (row_index - 1) & ")"
    ws_report.Cells(row_index, 2).NumberFormat = "#,##0"
    ws_report.Cells(row_index, 3).Formula = "=SUM(C10:C" & (row_index - 1) & ")"
    ws_report.Cells(row_index, 3).NumberFormat = "#,##0"
    ws_report.Cells(row_index, 4).Formula = "=SUM(D10:D" & (row_index - 1) & ")"
    ws_report.Cells(row_index, 4).NumberFormat = "#,##0"
    ws_report.Cells(row_index, 5).Formula = "=SUM(E10:E" & (row_index - 1) & ")"
    ws_report.Cells(row_index, 5).NumberFormat = "#,##0"
    ws_report.Cells(row_index, 6).Formula = "=SUM(F10:F" & (row_index - 1) & ")"
    ws_report.Cells(row_index, 6).NumberFormat = "#,##0"
    
    ' 月別集計をレポートに出力
    row_index = 16
    
    ' 月キーを日付順にソート
    Dim month_keys As Variant
    month_keys = months.Keys
    
    ' 月別集計をレポートに出力
    Dim month_key_iter As Variant
    For Each month_key_iter In month_keys
        ws_report.Cells(row_index, 1).Value = month_key_iter
        ws_report.Cells(row_index, 2).Value = months(month_key_iter)(0)
        ws_report.Cells(row_index, 3).Value = months(month_key_iter)(1)
        ws_report.Cells(row_index, 3).NumberFormat = "#,##0"
        ws_report.Cells(row_index, 4).Value = months(month_key_iter)(2)
        ws_report.Cells(row_index, 4).NumberFormat = "#,##0"
        ws_report.Cells(row_index, 5).Value = months(month_key_iter)(3)
        ws_report.Cells(row_index, 5).NumberFormat = "#,##0"
        ws_report.Cells(row_index, 6).Value = months(month_key_iter)(4)
        ws_report.Cells(row_index, 6).NumberFormat = "#,##0"
        ws_report.Cells(row_index, 7).Value = months(month_key_iter)(5)
        ws_report.Cells(row_index, 7).NumberFormat = "#,##0"
        row_index = row_index + 1
    Next month_key_iter
    
    ' 合計行を追加
    ws_report.Cells(row_index, 1).Value = "合計"
    ws_report.Cells(row_index, 1).Font.Bold = True
    ws_report.Cells(row_index, 2).Formula = "=SUM(B16:B" & (row_index - 1) & ")"
    ws_report.Cells(row_index, 3).Formula = "=SUM(C16:C" & (row_index - 1) & ")"
    ws_report.Cells(row_index, 3).NumberFormat = "#,##0"
    ws_report.Cells(row_index, 4).Formula = "=SUM(D16:D" & (row_index - 1) & ")"
    ws_report.Cells(row_index, 4).NumberFormat = "#,##0"
    ws_report.Cells(row_index, 5).Formula = "=SUM(E16:E" & (row_index - 1) & ")"
    ws_report.Cells(row_index, 5).NumberFormat = "#,##0"
    ws_report.Cells(row_index, 6).Formula = "=SUM(F16:F" & (row_index - 1) & ")"
    ws_report.Cells(row_index, 6).NumberFormat = "#,##0"
    ws_report.Cells(row_index, 7).Formula = "=SUM(G16:G" & (row_index - 1) & ")"
    ws_report.Cells(row_index, 7).NumberFormat = "#,##0"
    
    ' 請求日別集計のヘッダーを設定
    ws_report.Range("A" & (row_index + 3)).Value = "【請求日別集計】"
    ws_report.Range("A" & (row_index + 3)).Font.Bold = True
    ws_report.Range("A" & (row_index + 4)).Value = "請求日"
    ws_report.Range("B" & (row_index + 4)).Value = "件数"
    ws_report.Range("C" & (row_index + 4)).Value = "金額合計"
    ws_report.Range("D" & (row_index + 4)).Value = "主保険請求額合計"
    ws_report.Range("E" & (row_index + 4)).Value = "公費請求額合計"
    
    ' 請求日別集計をレポートに出力
    Dim billing_date_row_index As Long
    billing_date_row_index = row_index + 5
    
    ' 請求日キーを日付順にソート
    Dim billing_date_keys As Variant
    billing_date_keys = billing_dates.Keys
    
    ' 請求日別集計をレポートに出力
    Dim billing_date_key As Variant
    For Each billing_date_key In billing_date_keys
        ws_report.Cells(billing_date_row_index, 1).Value = billing_date_key
        ws_report.Cells(billing_date_row_index, 2).Value = billing_dates(billing_date_key)(0)
        ws_report.Cells(billing_date_row_index, 3).Value = billing_dates(billing_date_key)(1)
        ws_report.Cells(billing_date_row_index, 3).NumberFormat = "#,##0"
        ws_report.Cells(billing_date_row_index, 4).Value = billing_dates(billing_date_key)(2)
        ws_report.Cells(billing_date_row_index, 4).NumberFormat = "#,##0"
        ws_report.Cells(billing_date_row_index, 5).Value = billing_dates(billing_date_key)(3)
        ws_report.Cells(billing_date_row_index, 5).NumberFormat = "#,##0"
        billing_date_row_index = billing_date_row_index + 1
    Next billing_date_key
    
    ' 合計行を追加
    Dim start_row As Long
    start_row = row_index + 5
    ws_report.Cells(billing_date_row_index, 1).Value = "合計"
    ws_report.Cells(billing_date_row_index, 1).Font.Bold = True
    ws_report.Cells(billing_date_row_index, 2).Formula = "=SUM(B" & start_row & ":B" & (billing_date_row_index - 1) & ")"
    ws_report.Cells(billing_date_row_index, 3).Formula = "=SUM(C" & start_row & ":C" & (billing_date_row_index - 1) & ")"
    ws_report.Cells(billing_date_row_index, 3).NumberFormat = "#,##0"
    ws_report.Cells(billing_date_row_index, 4).Formula = "=SUM(D" & start_row & ":D" & (billing_date_row_index - 1) & ")"
    ws_report.Cells(billing_date_row_index, 4).NumberFormat = "#,##0"
    ws_report.Cells(billing_date_row_index, 5).Formula = "=SUM(E" & start_row & ":E" & (billing_date_row_index - 1) & ")"
    ws_report.Cells(billing_date_row_index, 5).NumberFormat = "#,##0"
    
    ' レポートの書式設定
    ws_report.Columns("A:G").AutoFit
    
    ' レポートを表示
    ws_report.Activate
    
    MsgBox "集計レポートを作成しました。", vbInformation, "完了"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in CreateDatabaseSummaryReport"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    MsgBox "集計レポート作成中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
End Sub

' データベースメニューを表示する関数
Public Sub ShowDatabaseMenu()
    On Error GoTo ErrorHandler
    
    ' データベースメニューフォームを表示
    Dim result As VbMsgBoxResult
    result = MsgBox("データベース機能を選択してください：" & vbCrLf & vbCrLf & _
                    "「はい」：データベース検索・フィルタリング" & vbCrLf & _
                    "「いいえ」：データベースをCSVにエクスポート" & vbCrLf & _
                    "「キャンセル」：集計レポート作成", _
                    vbYesNoCancel + vbQuestion, "データベース機能")
    
    Select Case result
        Case vbYes
            ' データベース検索
            SearchDatabase
        Case vbNo
            ' CSVエクスポート
            ExportDatabaseToCsv
        Case vbCancel
            ' 集計レポート作成
            CreateDatabaseSummaryReport
    End Select
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in ShowDatabaseMenu"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    MsgBox "データベースメニュー表示中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
End Sub
