Attribute VB_Name = "SummaryMenuModule"
Option Explicit

' メニューからまとめシート更新フォームを表示するサブルーチン
Public Sub ShowSummaryUpdateForm()
    SummaryUpdateForm.Show
End Sub

' CSVファイル処理後にまとめシートを更新するサブルーチン
Public Sub UpdateSummariesAfterProcessing()
    On Error Resume Next
    
    ' 保存先フォルダの取得
    Dim save_folder As String
    save_folder = ThisWorkbook.Sheets(1).Range("B3").Value
    
    If save_folder <> "" Then
        ' FileSystemObjectの初期化
        Dim file_system As Object
        Set file_system = CreateObject("Scripting.FileSystemObject")
        
        ' すべてのレポートファイルのまとめシートを更新
        SummaryIntegrationModule.ProcessReportSummaries file_system, save_folder
    End If
End Sub
