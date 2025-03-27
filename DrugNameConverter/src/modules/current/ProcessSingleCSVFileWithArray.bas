Attribute VB_Name = "ProcessSingleCSVFileWithArray"
Option Explicit

' 配列を使用して単一のCSVファイルを処理する
Public Sub ProcessSingleCSVFileWithArray(ByVal filePath As String, ByVal fileData As FileShelfData)
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim line As String
    Dim row As Long
    Dim settingsSheet As Worksheet
    Dim invalidCodes As New Collection
    Dim i As Integer
    
    ' 設定シートを取得
    Set settingsSheet = ThisWorkbook.Sheets("設定")
    
    ' 設定シートの既存データをクリア（A7以降）
    Dim lastRow As Long
    lastRow = settingsSheet.Cells(settingsSheet.Rows.Count, "A").End(xlUp).row
    If lastRow >= 7 Then
        settingsSheet.Range("A7:B" & lastRow).ClearContents
    End If
    
    ' 開始行を設定
    row = 7
    
    ' CSVファイルに対応する棚名を設定シートのB1-B3に設定
    ' 注意: 入力されていない棚名は更新しない（そのままにする）
    For i = 1 To 3
        If fileData.ShelfNames(i) <> "" Then
            settingsSheet.Cells(i, 2).Value = fileData.ShelfNames(i)
        End If
    Next i
    
    ' CSVファイルを開く
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    ' ファイルの各行を読み込む
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        
        ' 空行をスキップ
        If Trim(line) <> "" Then
            ' GTINコードのバリデーション（14桁の数字かチェック）
            If IsValidGTIN14(line) Then
                ' 設定シートにGTINコードを書き込む (A列)
                settingsSheet.Cells(row, 1).Value = line
                
                ' 次の行へ
                row = row + 1
            Else
                ' 無効なGTINコードを記録
                On Error Resume Next
                invalidCodes.Add line
                On Error GoTo ErrorHandler
            End If
        End If
    Loop
    
    ' ファイルを閉じる
    Close #fileNum
    
    ' 医薬品コードに対応する医薬品名を取得して処理
    ProcessItems
    
    ' 無効なGTINコードがあれば報告
    ReportInvalidCodes invalidCodes
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    Close #fileNum
    MsgBox "CSVファイルの処理中にエラーが発生しました: " & Err.Description, vbCritical
End Sub
