Attribute VB_Name = "SummarySheetFunctions"
Option Explicit

' シート3（まとめシート）を作成・更新する関数
Public Function CreateSummarySheet(ByVal wb As Workbook) As Boolean
    On Error GoTo ErrorHandler
    
    ' シート3が存在するか確認し、存在しない場合は作成
    Dim ws_summary As Worksheet
    On Error Resume Next
    Set ws_summary = wb.Worksheets("まとめ")
    On Error GoTo ErrorHandler
    
    If ws_summary Is Nothing Then
        ' シート3を作成
        Set ws_summary = wb.Worksheets.Add(After:=wb.Worksheets(2))
        ws_summary.Name = "まとめ"
        
        ' ヘッダーの設定
        With ws_summary
            .Range("A1").Value = "保険請求管理まとめ"
            .Range("A1").Font.Size = 14
            .Range("A1").Font.Bold = True
            
            ' 社保セクション
            .Range("A3").Value = "【社保】"
            .Range("A3").Font.Bold = True
            .Range("A4").Value = "区分"
            .Range("B4").Value = "件数"
            .Range("C4").Value = "金額"
            .Range("A5").Value = "未請求"
            .Range("A6").Value = "返戻"
            .Range("A7").Value = "減点"
            .Range("A8").Value = "合計"
            
            ' 国保セクション
            .Range("A10").Value = "【国保】"
            .Range("A10").Font.Bold = True
            .Range("A11").Value = "区分"
            .Range("B11").Value = "件数"
            .Range("C11").Value = "金額"
            .Range("A12").Value = "未請求"
            .Range("A13").Value = "返戻"
            .Range("A14").Value = "減点"
            .Range("A15").Value = "合計"
            
            ' 全体合計セクション
            .Range("A17").Value = "【総合計】"
            .Range("A17").Font.Bold = True
            .Range("A18").Value = "区分"
            .Range("B18").Value = "件数"
            .Range("C18").Value = "金額"
            .Range("A19").Value = "未請求"
            .Range("A20").Value = "返戻"
            .Range("A21").Value = "減点"
            .Range("A22").Value = "合計"
            
            ' 書式設定
            .Range("A3:C22").Borders.LineStyle = xlContinuous
            .Columns("A:C").AutoFit
        End With
    End If
    
    ' 詳細シートからデータを集計
    UpdateSummaryFromDetails wb
    
    CreateSummarySheet = True
    Exit Function
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in CreateSummarySheet"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    MsgBox "まとめシートの作成中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
    CreateSummarySheet = False
End Function

' 詳細シートからデータを集計してまとめシートを更新する関数
Private Sub UpdateSummaryFromDetails(ByVal wb As Workbook)
    On Error GoTo ErrorHandler
    
    Dim ws_details As Worksheet
    Dim ws_summary As Worksheet
    Dim i As Long, j As Long
    
    ' 詳細シートを取得
    On Error Resume Next
    Set ws_details = wb.Worksheets(2)
    Set ws_summary = wb.Worksheets("まとめ")
    On Error GoTo ErrorHandler
    
    If ws_details Is Nothing Or ws_summary Is Nothing Then
        Exit Sub
    End If
    
    ' カウンターと合計の初期化
    Dim shaho_unclaimed_count As Long, shaho_unclaimed_amount As Currency
    Dim shaho_return_count As Long, shaho_return_amount As Currency
    Dim shaho_adjust_count As Long, shaho_adjust_amount As Currency
    
    Dim kokuho_unclaimed_count As Long, kokuho_unclaimed_amount As Currency
    Dim kokuho_return_count As Long, kokuho_return_amount As Currency
    Dim kokuho_adjust_count As Long, kokuho_adjust_amount As Currency
    
    ' 詳細シートのデータを読み取り
    Dim last_row As Long
    last_row = ws_details.Cells(ws_details.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To last_row
        ' 社保/国保の判定
        Dim is_shaho As Boolean
        is_shaho = (ws_details.Cells(i, 8).Value = "社保" Or InStr(ws_details.Cells(i, 8).Value, "社保") > 0)
        
        ' 種別の判定 (未請求/返戻/減点)
        Dim category As String
        category = CStr(ws_details.Cells(i, 2).Value)
        
        ' 金額の取得
        Dim amount As Currency
        On Error Resume Next
        amount = CCur(ws_details.Cells(i, 10).Value)
        If Err.Number <> 0 Then amount = 0
        On Error GoTo ErrorHandler
        
        ' カウントと集計
        If is_shaho Then
            If InStr(category, "未請求") > 0 Then
                shaho_unclaimed_count = shaho_unclaimed_count + 1
                shaho_unclaimed_amount = shaho_unclaimed_amount + amount
            ElseIf InStr(category, "返戻") > 0 Then
                shaho_return_count = shaho_return_count + 1
                shaho_return_amount = shaho_return_amount + amount
            ElseIf InStr(category, "減点") > 0 Or InStr(category, "査定") > 0 Then
                shaho_adjust_count = shaho_adjust_count + 1
                shaho_adjust_amount = shaho_adjust_amount + amount
            End If
        Else
            If InStr(category, "未請求") > 0 Then
                kokuho_unclaimed_count = kokuho_unclaimed_count + 1
                kokuho_unclaimed_amount = kokuho_unclaimed_amount + amount
            ElseIf InStr(category, "返戻") > 0 Then
                kokuho_return_count = kokuho_return_count + 1
                kokuho_return_amount = kokuho_return_amount + amount
            ElseIf InStr(category, "減点") > 0 Or InStr(category, "査定") > 0 Then
                kokuho_adjust_count = kokuho_adjust_count + 1
                kokuho_adjust_amount = kokuho_adjust_amount + amount
            End If
        End If
    Next i
    
    ' まとめシートに結果を反映
    ' 社保
    ws_summary.Range("B5").Value = shaho_unclaimed_count
    ws_summary.Range("C5").Value = shaho_unclaimed_amount
    ws_summary.Range("B6").Value = shaho_return_count
    ws_summary.Range("C6").Value = shaho_return_amount
    ws_summary.Range("B7").Value = shaho_adjust_count
    ws_summary.Range("C7").Value = shaho_adjust_amount
    ws_summary.Range("B8").Value = shaho_unclaimed_count + shaho_return_count + shaho_adjust_count
    ws_summary.Range("C8").Value = shaho_unclaimed_amount + shaho_return_amount + shaho_adjust_amount
    
    ' 国保
    ws_summary.Range("B12").Value = kokuho_unclaimed_count
    ws_summary.Range("C12").Value = kokuho_unclaimed_amount
    ws_summary.Range("B13").Value = kokuho_return_count
    ws_summary.Range("C13").Value = kokuho_return_amount
    ws_summary.Range("B14").Value = kokuho_adjust_count
    ws_summary.Range("C14").Value = kokuho_adjust_amount
    ws_summary.Range("B15").Value = kokuho_unclaimed_count + kokuho_return_count + kokuho_adjust_count
    ws_summary.Range("C15").Value = kokuho_unclaimed_amount + kokuho_return_amount + kokuho_adjust_amount
    
    ' 全体合計
    ws_summary.Range("B19").Value = shaho_unclaimed_count + kokuho_unclaimed_count
    ws_summary.Range("C19").Value = shaho_unclaimed_amount + kokuho_unclaimed_amount
    ws_summary.Range("B20").Value = shaho_return_count + kokuho_return_count
    ws_summary.Range("C20").Value = shaho_return_amount + kokuho_return_amount
    ws_summary.Range("B21").Value = shaho_adjust_count + kokuho_adjust_count
    ws_summary.Range("C21").Value = shaho_adjust_amount + kokuho_adjust_amount
    ws_summary.Range("B22").Value = ws_summary.Range("B19").Value + ws_summary.Range("B20").Value + ws_summary.Range("B21").Value
    ws_summary.Range("C22").Value = ws_summary.Range("C19").Value + ws_summary.Range("C20").Value + ws_summary.Range("C21").Value
    
    ' 金額の書式設定
    ws_summary.Range("C5:C8").NumberFormat = "#,##0"
    ws_summary.Range("C12:C15").NumberFormat = "#,##0"
    ws_summary.Range("C19:C22").NumberFormat = "#,##0"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in UpdateSummaryFromDetails"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    MsgBox "まとめシートの更新中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
End Sub
