Attribute VB_Name = "HalfYearCalculationModule"
Option Explicit

' 半期ごとの決算書作成のための売掛金繰越額を計算する関数
Public Function CalculateAccountsReceivable() As Boolean
    On Error GoTo ErrorHandler
    
    Dim report_folder As String
    Dim file_system As Object
    Dim report_files As Collection
    Dim year_period As Integer
    Dim period_type As String
    
    ' 対象期間の選択
    year_period = InputBox("対象年度を入力してください（例：2025）", "売掛金繰越額計算", Year(Date))
    If year_period = 0 Then Exit Function
    
    period_type = ""
    Do While period_type <> "上期" And period_type <> "下期"
        period_type = InputBox("対象期間を入力してください（上期/下期）", "売掛金繰越額計算", "上期")
        If period_type = "" Then Exit Function
    Loop
    
    ' レポートフォルダの選択
    report_folder = SelectReportFolder()
    If report_folder = "" Then Exit Function
    
    ' FileSystemObjectの初期化
    Set file_system = CreateObject("Scripting.FileSystemObject")
    
    ' 対象期間のレポートファイルを収集
    Set report_files = GetReportFilesForPeriod(file_system, report_folder, year_period, period_type)
    
    If report_files.Count = 0 Then
        MsgBox "対象期間のレポートファイルが見つかりませんでした。", vbExclamation, "エラー"
        CalculateAccountsReceivable = False
        Exit Function
    End If
    
    ' 売掛金繰越額計算レポートの作成
    CreateAccountsReceivableReport file_system, report_files, year_period, period_type
    
    CalculateAccountsReceivable = True
    Exit Function
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in CalculateAccountsReceivable"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    MsgBox "売掛金繰越額計算中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
    CalculateAccountsReceivable = False
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

' 対象期間のレポートファイルを取得する関数
Private Function GetReportFilesForPeriod(file_system As Object, report_folder As String, _
                                      year_period As Integer, period_type As String) As Collection
    Dim report_files As New Collection
    Dim file_obj As Object
    Dim target_months As Variant
    
    ' 対象月を設定
    If period_type = "上期" Then
        target_months = Array("04", "05", "06", "07", "08", "09")
    Else
        target_months = Array("10", "11", "12", "01", "02", "03")
    End If
    
    ' フォルダ内のXLSMファイルを検索
    For Each file_obj In file_system.GetFolder(report_folder).Files
        If LCase(file_system.GetExtensionName(file_obj.Name)) = "xlsm" Then
            ' ファイル名に年度と月が含まれているか確認
            Dim file_name As String, is_target As Boolean
            file_name = file_obj.Name
            is_target = False
            
            ' 対象年度の確認
            If InStr(file_name, CStr(year_period)) > 0 Then
                ' 対象月の確認
                Dim i As Integer
                For i = LBound(target_months) To UBound(target_months)
                    If InStr(file_name, "月" & target_months(i) & "月") > 0 Or _
                       InStr(file_name, target_months(i) & "月") > 0 Then
                        is_target = True
                        Exit For
                    End If
                Next i
            End If
            
            If is_target Then
                report_files.Add file_obj
            End If
        End If
    Next file_obj
    
    Set GetReportFilesForPeriod = report_files
End Function

' 売掛金繰越額計算レポートを作成する関数
Private Sub CreateAccountsReceivableReport(file_system As Object, report_files As Collection, _
                                        year_period As Integer, period_type As String)
    Dim report_wb As Workbook
    Dim summary_ws As Worksheet
    Dim file_obj As Object
    Dim row_index As Long
    
    ' 新規ワークブックの作成
    Set report_wb = Workbooks.Add
    Set summary_ws = report_wb.Sheets(1)
    summary_ws.Name = "売掛金繰越額計算"
    
    ' ヘッダーの設定
    With summary_ws
        .Range("A1").Value = year_period & "年度 " & period_type & " 売掛金繰越額計算"
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "月次"
        .Range("B3").Value = "未請求"
        .Range("C3").Value = "返戻"
        .Range("D3").Value = "減点"
        .Range("E3").Value = "合計"
        
        .Range("A3:E3").Font.Bold = True
        .Range("A3:E3").Borders.LineStyle = xlContinuous
    End With
    
    ' 各レポートファイルから情報を集計
    row_index = 4
    For Each file_obj In report_files
        Dim source_wb As Workbook
        Dim source_summary As Worksheet
        Dim month_name As String
        
        ' レポートファイルを開く
        Set source_wb = Workbooks.Open(file_obj.Path, ReadOnly:=True)
        
        ' まとめシートを検索
        On Error Resume Next
        Set source_summary = source_wb.Worksheets("まとめ")
        On Error GoTo 0
        
        If Not source_summary Is Nothing Then
            ' 月次情報を取得
            month_name = ExtractMonthFromFileName(file_obj.Name)
            
            ' データを転記
            summary_ws.Cells(row_index, 1).Value = month_name
            
            ' 総合計の値を取得（行22のデータ）
            summary_ws.Cells(row_index, 2).Value = source_summary.Range("C19").Value ' 未請求合計
            summary_ws.Cells(row_index, 3).Value = source_summary.Range("C20").Value ' 返戻合計
            summary_ws.Cells(row_index, 4).Value = source_summary.Range("C21").Value ' 減点合計
            summary_ws.Cells(row_index, 5).Value = source_summary.Range("C22").Value ' 総合計
            
            row_index = row_index + 1
        End If
        
        ' ソースワークブックを閉じる
        source_wb.Close SaveChanges:=False
    Next file_obj
    
    ' 合計行の追加
    With summary_ws
        .Cells(row_index, 1).Value = "合計"
        .Cells(row_index, 1).Font.Bold = True
        
        .Cells(row_index, 2).Formula = "=SUM(B4:B" & (row_index - 1) & ")"
        .Cells(row_index, 3).Formula = "=SUM(C4:C" & (row_index - 1) & ")"
        .Cells(row_index, 4).Formula = "=SUM(D4:D" & (row_index - 1) & ")"
        .Cells(row_index, 5).Formula = "=SUM(E4:E" & (row_index - 1) & ")"
        
        .Range("A" & row_index & ":E" & row_index).Font.Bold = True
        .Range("A" & row_index & ":E" & row_index).Borders.LineStyle = xlContinuous
        
        ' 書式設定
        .Range("B4:E" & row_index).NumberFormat = "#,##0"
        .Range("A3:E" & row_index).Borders.LineStyle = xlContinuous
        .Columns("A:E").AutoFit
    End With
    
    ' ファイルの保存
    Dim save_path As String
    save_path = ThisWorkbook.Sheets(1).Range("B3").Value
    If save_path = "" Then
        save_path = GetDesktopPath()
    End If
    
    Dim report_file_name As String
    report_file_name = "売掛金繰越額計算_" & year_period & "_" & period_type & ".xlsx"
    report_wb.SaveAs Filename:=save_path & "\" & report_file_name, FileFormat:=xlOpenXMLWorkbook
    
    MsgBox "売掛金繰越額計算レポートを作成しました。" & vbCrLf & _
           "保存先: " & save_path & "\" & report_file_name, vbInformation, "完了"
End Sub

' ファイル名から月次情報を抽出する関数
Private Function ExtractMonthFromFileName(file_name As String) As String
    Dim month_match As Object
    Dim month_pattern As String
    Dim reg_ex As Object
    
    Set reg_ex = CreateObject("VBScript.RegExp")
    month_pattern = "(\d{1,2})月"
    
    With reg_ex
        .Global = False
        .Pattern = month_pattern
        .IgnoreCase = True
        
        Set month_match = .Execute(file_name)
        
        If month_match.Count > 0 Then
            ExtractMonthFromFileName = month_match(0).SubMatches(0) & "月"
        Else
            ExtractMonthFromFileName = "不明"
        End If
    End With
End Function

' デスクトップパスを取得する関数
Private Function GetDesktopPath() As String
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    GetDesktopPath = wsh.SpecialFolders("Desktop")
End Function
