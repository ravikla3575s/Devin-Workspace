Attribute VB_Name = "SummaryIntegrationModule"
Option Explicit

' メインモジュールから呼び出すためのサブルーチン
Public Sub UpdateAllReportSummaries()
    On Error GoTo ErrorHandler
    
    Dim save_folder As String
    Dim file_system As Object
    
    ' 保存先フォルダの取得
    save_folder = ThisWorkbook.Sheets(1).Range("B3").Value
    If save_folder = "" Then
        ' 保存先フォルダが設定されていない場合は選択させる
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "レポートフォルダを選択してください"
            If .Show = -1 Then
                save_folder = .SelectedItems(1)
            Else
                Exit Sub
            End If
        End With
    End If
    
    ' FileSystemObjectの初期化
    Set file_system = CreateObject("Scripting.FileSystemObject")
    
    ' すべてのレポートファイルのまとめシートを更新
    ProcessReportSummaries file_system, save_folder
    
    MsgBox "すべてのレポートファイルのまとめシートを更新しました。", vbInformation, "完了"
    Exit Sub
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error occurred in UpdateAllReportSummaries"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    MsgBox "レポートのまとめシート更新中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
End Sub

' レポートフォルダ内のすべてのレポートファイルのまとめシートを更新する関数
Private Sub ProcessReportSummaries(file_system As Object, report_folder As String)
    On Error GoTo ErrorHandler
    
    Dim file_obj As Object
    Dim report_wb As Workbook
    
    ' レポートフォルダ内のすべてのXLSMファイルを処理
    For Each file_obj In file_system.GetFolder(report_folder).Files
        If LCase(file_system.GetExtensionName(file_obj.Name)) = "xlsm" Then
            ' ワークブックを開く
            On Error Resume Next
            Set report_wb = Workbooks.Open(file_obj.Path, ReadOnly:=False, UpdateLinks:=False)
            If Err.Number <> 0 Then
                Debug.Print "ERROR: Failed to open workbook: " & file_obj.Path
                Debug.Print "Error number: " & Err.Number
                Debug.Print "Error description: " & Err.Description
                Err.Clear
                GoTo NextFile
            End If
            On Error GoTo ErrorHandler
            
            If Not report_wb Is Nothing Then
                ' まとめシートの作成・更新
                On Error Resume Next
                SummarySheetFunctions.CreateSummarySheet report_wb
                If Err.Number <> 0 Then
                    Debug.Print "ERROR in CreateSummarySheet: " & Err.Description
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
                
                ' ワークブックを保存して閉じる
                report_wb.Save
                report_wb.Close SaveChanges:=True
                Set report_wb = Nothing
            End If
        End If
NextFile:
    Next file_obj
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error occurred in ProcessReportSummaries"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    If Not report_wb Is Nothing Then
        report_wb.Close SaveChanges:=False
        Set report_wb = Nothing
    End If
End Sub
