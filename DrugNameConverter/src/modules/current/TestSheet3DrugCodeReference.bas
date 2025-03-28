Attribute VB_Name = "TestSheet3DrugCodeReference"
Option Explicit

' Sheet3を使用した医薬品コード参照機能のテスト
Public Sub TestSheet3DrugCodeReference()
    On Error GoTo ErrorHandler
    
    ' テスト開始メッセージ
    Debug.Print "Sheet3を使用した医薬品コード参照機能のテストを開始します..."
    
    ' テスト用のGTINコード（実際のテストでは既知のコードを使用）
    Dim testGtinCodes(1 To 3) As String
    testGtinCodes(1) = "12345678901234"  ' 存在するコード（テスト用）
    testGtinCodes(2) = "23456789012345"  ' 存在しないコード（テスト用）
    testGtinCodes(3) = "34567890123456"  ' 別のコード（テスト用）
    
    ' 各GTINコードでテスト
    Dim i As Integer
    For i = 1 To UBound(testGtinCodes)
        Dim drugInfo As DrugInfo
        drugInfo = GS1CodeProcessor.GetDrugInfoFromGS1Code(testGtinCodes(i))
        
        ' 結果を表示
        Debug.Print "GTIN: " & testGtinCodes(i)
        Debug.Print "  医薬品名: " & drugInfo.DrugName
        Debug.Print "  成分名: " & drugInfo.BaseName
        Debug.Print "  剤形: " & drugInfo.FormType
        Debug.Print "  規格: " & drugInfo.Strength
        Debug.Print "  メーカー: " & drugInfo.Maker
        Debug.Print "  包装形態: " & drugInfo.PackageForm
        Debug.Print "  包装規格: " & drugInfo.PackageSpec
        Debug.Print "-------------------"
    Next i
    
    ' 実際のSheet3のデータを使用したテスト
    Debug.Print "Sheet3の実データを使用したテスト:"
    
    ' Sheet3から最初の3つのGTINコードを取得
    Dim ws3 As Worksheet
    Dim lastRow As Long
    Dim actualGtinCodes(1 To 3) As String
    Dim codeCount As Integer
    
    Set ws3 = ThisWorkbook.Worksheets(3) ' 医薬品コードシート
    lastRow = ws3.Cells(ws3.Rows.Count, "F").End(xlUp).Row
    
    codeCount = 0
    For i = 2 To lastRow ' ヘッダー行をスキップ
        If codeCount < 3 Then
            If Len(ws3.Cells(i, "F").Value) > 0 Then
                codeCount = codeCount + 1
                actualGtinCodes(codeCount) = ws3.Cells(i, "F").Value
            End If
        Else
            Exit For
        End If
    Next i
    
    ' 実際のGTINコードでテスト
    For i = 1 To codeCount
        Dim actualDrugInfo As DrugInfo
        actualDrugInfo = GS1CodeProcessor.GetDrugInfoFromGS1Code(actualGtinCodes(i))
        
        ' 結果を表示
        Debug.Print "実際のGTIN: " & actualGtinCodes(i)
        Debug.Print "  医薬品名: " & actualDrugInfo.DrugName
        Debug.Print "  成分名: " & actualDrugInfo.BaseName
        Debug.Print "  剤形: " & actualDrugInfo.FormType
        Debug.Print "  規格: " & actualDrugInfo.Strength
        Debug.Print "  メーカー: " & actualDrugInfo.Maker
        Debug.Print "  包装形態: " & actualDrugInfo.PackageForm
        Debug.Print "  包装規格: " & actualDrugInfo.PackageSpec
        Debug.Print "-------------------"
    Next i
    
    ' ShelfManager.GetDrugName関数のテスト
    Debug.Print "ShelfManager.GetDrugName関数のテスト:"
    
    ' ShelfManagerモジュールのGetDrugName関数を直接呼び出せないため、
    ' 同等の処理をここで実装してテスト
    For i = 1 To codeCount
        Dim drugName As String
        drugName = GetDrugNameTest(actualGtinCodes(i))
        
        Debug.Print "GTIN: " & actualGtinCodes(i)
        Debug.Print "  医薬品名: " & drugName
        Debug.Print "-------------------"
    Next i
    
    ' テスト完了メッセージ
    Debug.Print "Sheet3を使用した医薬品コード参照機能のテストが完了しました。"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "テスト中にエラーが発生しました: " & Err.Description
End Sub

' ShelfManager.GetDrugName関数と同等の処理を実装したテスト用関数
Private Function GetDrugNameTest(gtin As String) As String
    On Error GoTo ErrorHandler
    
    ' GS1CodeProcessorを使用してGTIN-14コードから医薬品情報を取得
    Dim drugInfo As DrugInfo
    drugInfo = GS1CodeProcessor.GetDrugInfoFromGS1Code(gtin)
    
    ' 結果を返す
    GetDrugNameTest = drugInfo.DrugName
    
    Exit Function
    
ErrorHandler:
    GetDrugNameTest = ""
End Function

' 実際のSheet3のデータ構造を確認するテスト
Public Sub TestSheet3Structure()
    On Error GoTo ErrorHandler
    
    ' テスト開始メッセージ
    Debug.Print "Sheet3のデータ構造確認テストを開始します..."
    
    ' Sheet3を取得
    Dim ws3 As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    Set ws3 = ThisWorkbook.Worksheets(3) ' 医薬品コードシート
    
    ' 最終行と最終列を取得
    lastRow = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row
    lastCol = ws3.Cells(1, ws3.Columns.Count).End(xlToLeft).Column
    
    ' シート情報を表示
    Debug.Print "シート名: " & ws3.Name
    Debug.Print "データ行数: " & lastRow
    Debug.Print "データ列数: " & lastCol
    
    ' ヘッダー行を表示
    Dim col As Integer
    Dim headerRow As String
    
    headerRow = "ヘッダー行: "
    For col = 1 To lastCol
        headerRow = headerRow & ws3.Cells(1, col).Value & " | "
    Next col
    Debug.Print headerRow
    
    ' F列（GTINコード）とG列（医薬品名）の最初の5行を表示
    Debug.Print "GTINコードと医薬品名のサンプル:"
    Dim row As Integer
    For row = 2 To Application.WorksheetFunction.Min(6, lastRow)
        Debug.Print "  行" & row & ": " & ws3.Cells(row, "F").Value & " - " & ws3.Cells(row, "G").Value
    Next row
    
    ' テスト完了メッセージ
    Debug.Print "Sheet3のデータ構造確認テストが完了しました。"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "テスト中にエラーが発生しました: " & Err.Description
End Sub
