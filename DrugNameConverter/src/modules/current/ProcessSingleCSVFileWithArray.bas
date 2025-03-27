Attribute VB_Name = "ProcessSingleCSVFileWithArray"
Option Explicit

' 配列を使用して単一のCSVファイルを処理する（最適化版）
Public Sub ProcessSingleCSVFileWithArray(ByVal filePath As String, ByVal fileData As FileShelfData)
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim line As String
    Dim lines() As String
    Dim lineCount As Long
    Dim currentLine As Long
    Dim row As Long
    Dim settingsSheet As Worksheet
    Dim invalidCodes As New Collection
    Dim i As Integer
    Dim validGtinCodes() As String
    Dim validCodesCount As Long
    Dim isScreenUpdatingEnabled As Boolean
    Dim isCalculationAutomatic As Boolean
    
    ' 画面更新と自動計算を一時的に無効化（パフォーマンス向上）
    isScreenUpdatingEnabled = Application.ScreenUpdating
    isCalculationAutomatic = Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
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
    
    ' CSVファイルに対応する棚名を設定シートのB1-B3に一括設定
    ' 棚名を配列に集約して一括設定
    Dim shelfNames(1 To 3, 1 To 2) As Variant  ' 値とセル位置
    Dim shelfCount As Integer
    shelfCount = 0
    
    For i = 1 To 3
        If fileData.ShelfNames(i) <> "" Then
            shelfCount = shelfCount + 1
            shelfNames(shelfCount, 1) = fileData.ShelfNames(i)
            shelfNames(shelfCount, 2) = i
        End If
    Next i
    
    ' 一括設定（セルのアクセス回数を削減）
    If shelfCount > 0 Then
        For i = 1 To shelfCount
            settingsSheet.Cells(CInt(shelfNames(i, 2)), 2).Value = shelfNames(i, 1)
        Next i
    End If
    
    ' CSVファイルを一括読み込み
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    ' 行数を事前カウント（配列のサイズ確保用）
    lineCount = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        lineCount = lineCount + 1
    Loop
    
    ' ファイルを閉じて再度開く
    Close #fileNum
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    ' 行データを配列に読み込み
    ReDim lines(1 To lineCount) As String
    currentLine = 1
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        lines(currentLine) = line
        currentLine = currentLine + 1
    Loop
    
    ' ファイルを閉じる
    Close #fileNum
    
    ' 有効なGTINコードを一括抽出
    ReDim validGtinCodes(1 To lineCount) As String
    validCodesCount = 0
    
    For i = 1 To lineCount
        ' 空行をスキップ
        If Trim(lines(i)) <> "" Then
            ' GTINコードのバリデーション（14桁の数字かチェック）
            If IsValidGTIN14(lines(i)) Then
                validCodesCount = validCodesCount + 1
                validGtinCodes(validCodesCount) = lines(i)
            Else
                ' 無効なGTINコードを記録
                On Error Resume Next
                invalidCodes.Add lines(i)
                On Error GoTo ErrorHandler
            End If
        End If
    Next i
    
    ' 不要になった配列の開放
    Erase lines
    
    ' 有効なGTINコードをシートに一括書き込み
    If validCodesCount > 0 Then
        ReDim Preserve validGtinCodes(1 To validCodesCount) As String
        
        ' 一括書き込み用の範囲を作成
        Dim writeRange As Range
        Set writeRange = settingsSheet.Range("A7").Resize(validCodesCount, 1)
        
        ' 2次元配列に変換（Range.Valueには2次元配列が必要）
        Dim writeData() As Variant
        ReDim writeData(1 To validCodesCount, 1 To 1)
        
        For i = 1 To validCodesCount
            writeData(i, 1) = validGtinCodes(i)
        Next i
        
        ' 一括書き込み
        writeRange.Value = writeData
    End If
    
    ' 不要になった配列の開放
    Erase validGtinCodes
    Erase writeData
    Set writeRange = Nothing
    
    ' 医薬品コードに対応する医薬品名を取得して処理
    ProcessItems
    
    ' 無効なGTINコードがあれば報告
    ReportInvalidCodes invalidCodes
    
    ' オブジェクト変数のクリーンアップ
    Set settingsSheet = Nothing
    Set invalidCodes = Nothing
    
    ' 画面更新と自動計算を元に戻す
    Application.ScreenUpdating = isScreenUpdatingEnabled
    Application.Calculation = IIf(isCalculationAutomatic, xlCalculationAutomatic, xlCalculationManual)
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    ' クリーンアップ
    Close #fileNum
    Erase lines
    Erase validGtinCodes
    Erase writeData
    Set writeRange = Nothing
    Set settingsSheet = Nothing
    Set invalidCodes = Nothing
    
    Application.ScreenUpdating = isScreenUpdatingEnabled
    Application.Calculation = IIf(isCalculationAutomatic, xlCalculationAutomatic, xlCalculationManual)
    MsgBox "CSVファイルの処理中にエラーが発生しました: " & Err.Description, vbCritical
End Sub
