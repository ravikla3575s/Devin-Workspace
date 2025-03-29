Attribute VB_Name = "BillingDiscrepancyModule"
Option Explicit

' 請求誤差追求報告書を作成する関数
Public Function CreateBillingDiscrepancyReport() As Boolean
    On Error GoTo ErrorHandler
    
    Dim report_folder As String
    Dim file_system As Object
    Dim target_month As String
    Dim target_year As String
    
    ' 対象年月の選択
    target_year = InputBox("対象年を入力してください（例：2025）", "請求誤差追求報告書", Year(Date))
    If target_year = "" Then Exit Function
    
    target_month = InputBox("対象月を入力してください（例：4）", "請求誤差追求報告書", Month(Date))
    If target_month = "" Then Exit Function
    
    ' レポートフォルダの選択
    report_folder = SelectReportFolder()
    If report_folder = "" Then Exit Function
    
    ' FileSystemObjectの初期化
    Set file_system = CreateObject("Scripting.FileSystemObject")
    
    ' 対象月のレポートファイルを検索
    Dim report_file As Object
    Set report_file = FindReportFile(file_system, report_folder, target_year, target_month)
    
    If report_file Is Nothing Then
        MsgBox "対象月のレポートファイルが見つかりませんでした。", vbExclamation, "エラー"
        CreateBillingDiscrepancyReport = False
        Exit Function
    End If
    
    ' 請求誤差追求報告書の作成
    GenerateDiscrepancyReport file_system, report_file, target_year, target_month
    
    CreateBillingDiscrepancyReport = True
    Exit Function
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in CreateBillingDiscrepancyReport"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    MsgBox "請求誤差追求報告書の作成中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
    CreateBillingDiscrepancyReport = False
End Function

' レポートフォルダを選択する関数
Private Function SelectReportFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "レポートフォルダを選択してください"
        If .Show = -1 Then
            SelectReportFolder = .SelectedItems(1) & "\"
        Else
            SelectReportFolder = ""
        End If
    End With
End Function

' 対象月のレポートファイルを検索する関数
Private Function FindReportFile(file_system As Object, report_folder As String, _
                             target_year As String, target_month As String) As Object
    Dim file_obj As Object
    Dim formatted_month As String
    
    ' 月の書式を整える
    formatted_month = Format(CInt(target_month), "00")
    
    ' フォルダ内のXLSMファイルを検索
    For Each file_obj In file_system.GetFolder(report_folder).Files
        If LCase(file_system.GetExtensionName(file_obj.Name)) = "xlsm" Then
            ' ファイル名に年と月が含まれているか確認
            If InStr(file_obj.Name, target_year) > 0 And _
               (InStr(file_obj.Name, formatted_month & "月") > 0 Or _
                InStr(file_obj.Name, "月" & formatted_month) > 0) Then
                Set FindReportFile = file_obj
                Exit Function
            End If
        End If
    Next file_obj
    
    Set FindReportFile = Nothing
End Function

' 請求誤差追求報告書を生成する関数
Private Sub GenerateDiscrepancyReport(file_system As Object, report_file As Object, _
                                   target_year As String, target_month As String)
    Dim source_wb As Workbook
    Dim report_wb As Workbook
    Dim main_ws As Worksheet
    Dim details_ws As Worksheet
    Dim discrepancy_ws As Worksheet
    Dim last_row As Long
    Dim i As Long, row_index As Long
    
    ' ソースワークブックを開く
    Set source_wb = Workbooks.Open(report_file.Path, ReadOnly:=True)
    
    ' メインシートと詳細シートを取得
    On Error Resume Next
    Set main_ws = source_wb.Worksheets(1)
    Set details_ws = source_wb.Worksheets(2)
    On Error GoTo 0
    
    If main_ws Is Nothing Or details_ws Is Nothing Then
        MsgBox "レポートファイルに必要なシートが見つかりませんでした。", vbExclamation, "エラー"
        source_wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' 新規ワークブックの作成
    Set report_wb = Workbooks.Add
    Set discrepancy_ws = report_wb.Sheets(1)
    discrepancy_ws.Name = "請求誤差追求報告書"
    
    ' ヘッダーの設定
    With discrepancy_ws
        .Range("A1").Value = target_year & "年" & target_month & "月分 請求誤差追求報告書"
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "No."
        .Range("B3").Value = "患者名"
        .Range("C3").Value = "調剤日"
        .Range("D3").Value = "医療機関"
        .Range("E3").Value = "種別"
        .Range("F3").Value = "保険種別"
        .Range("G3").Value = "請求点数"
        .Range("H3").Value = "実績点数"
        .Range("I3").Value = "差異"
        .Range("J3").Value = "原因"
        .Range("K3").Value = "対策"
        
        .Range("A3:K3").Font.Bold = True
        .Range("A3:K3").Borders.LineStyle = xlContinuous
    End With
    
    ' 詳細シートからデータを抽出
    last_row = details_ws.Cells(details_ws.Rows.Count, "A").End(xlUp).Row
    row_index = 4
    
    For i = 2 To last_row
        ' 種別が「減点」または「査定」のデータを抽出
        Dim category As String
        category = CStr(details_ws.Cells(i, 2).Value)
        
        If InStr(category, "減点") > 0 Or InStr(category, "査定") > 0 Then
            ' データを転記
            discrepancy_ws.Cells(row_index, 1).Value = row_index - 3 ' No.
            discrepancy_ws.Cells(row_index, 2).Value = details_ws.Cells(i, 3).Value ' 患者名
            discrepancy_ws.Cells(row_index, 3).Value = details_ws.Cells(i, 4).Value ' 調剤日
            discrepancy_ws.Cells(row_index, 4).Value = details_ws.Cells(i, 5).Value ' 医療機関
            discrepancy_ws.Cells(row_index, 5).Value = details_ws.Cells(i, 2).Value ' 種別
            discrepancy_ws.Cells(row_index, 6).Value = details_ws.Cells(i, 8).Value ' 保険種別
            
            ' 請求点数と実績点数の計算
            Dim amount As Currency
            On Error Resume Next
            amount = CCur(details_ws.Cells(i, 10).Value)
            If Err.Number <> 0 Then amount = 0
            On Error GoTo 0
            
            Dim claimed_points As Long, actual_points As Long
            claimed_points = CLng(details_ws.Cells(i, 9).Value)
            actual_points = claimed_points - CLng(amount / 10)
            
            discrepancy_ws.Cells(row_index, 7).Value = claimed_points ' 請求点数
            discrepancy_ws.Cells(row_index, 8).Value = actual_points ' 実績点数
            discrepancy_ws.Cells(row_index, 9).Value = claimed_points - actual_points ' 差異
            
            ' 原因と対策は空欄（ユーザーが入力）
            
            row_index = row_index + 1
        End If
    Next i
    
    ' 書式設定
    With discrepancy_ws
        .Range("A3:K" & (row_index - 1)).Borders.LineStyle = xlContinuous
        .Range("G4:I" & (row_index - 1)).NumberFormat = "#,##0"
        .Columns("A:K").AutoFit
    End With
    
    ' ファイルの保存
    Dim save_path As String
    save_path = ThisWorkbook.Sheets(1).Range("B3").Value
    If save_path = "" Then
        save_path = GetDesktopPath()
    End If
    
    Dim report_file_name As String
    report_file_name = "請求誤差追求報告書_" & target_year & "_" & target_month & ".xlsx"
    report_wb.SaveAs Filename:=save_path & "\" & report_file_name, FileFormat:=xlOpenXMLWorkbook
    
    ' ソースワークブックを閉じる
    source_wb.Close SaveChanges:=False
    
    MsgBox "請求誤差追求報告書を作成しました。" & vbCrLf & _
           "保存先: " & save_path & "\" & report_file_name, vbInformation, "完了"
End Sub

' デスクトップパスを取得する関数
Private Function GetDesktopPath() As String
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    GetDesktopPath = wsh.SpecialFolders("Desktop")
End Function
