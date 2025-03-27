Attribute VB_Name = "ProcessFileBatch"
Option Explicit

' ファイルバッチを処理する（メモリ使用量を最適化）
Public Sub ProcessFileBatch(ByVal folderPath As String, ByVal fileShelfData() As FileShelfData, _
                          ByVal startIndex As Integer, ByVal endIndex As Integer)
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim filePath As String
    Dim originalStatusBar As String
    
    ' 元のステータスバーテキストを保存
    originalStatusBar = Application.StatusBar
    
    ' 指定範囲のファイルを処理
    For i = startIndex To endIndex
        ' 進捗状況表示
        Application.StatusBar = "CSVファイルを処理中... (" & i & "/" & UBound(fileShelfData) & ")"
        
        ' 現在のファイルのパスを作成
        filePath = folderPath & "\" & fileShelfData(i).FileName
        
        ' ファイルを処理
        ProcessSingleCSVFileWithArray filePath, fileShelfData(i)
        
        ' UIの応答性を維持
        DoEvents
        
        ' 定期的にメモリを解放
        If i Mod 10 = 0 Then
            CollectGarbage
        End If
    Next i
    
    ' ステータスバーを元に戻す
    Application.StatusBar = originalStatusBar
    
    Exit Sub
    
ErrorHandler:
    ' ステータスバーを元に戻す
    Application.StatusBar = originalStatusBar
    MsgBox "ファイルバッチの処理中にエラーが発生しました: " & Err.Description, vbCritical
End Sub
