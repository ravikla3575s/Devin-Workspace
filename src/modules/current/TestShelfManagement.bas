Attribute VB_Name = "TestShelfManagement"
Option Explicit

' 棚番一括更新機能のテスト用モジュール

' すべてのテスト関数を実行
Public Sub RunAllShelfManagementTests()
    On Error GoTo ErrorHandler
    
    Debug.Print "===== 棚番一括更新機能の全テストを開始します ====="
    Debug.Print ""
    
    ' 各テスト関数を順番に実行
    TestCSVImport
    Debug.Print ""
    
    TestDrugNameMatching
    Debug.Print ""
    
    TestShelfNameUpdate
    Debug.Print ""
    
    TestUndoFunctionality
    Debug.Print ""
    
    Debug.Print "===== 全テストが完了しました ====="
    Exit Sub
    
ErrorHandler:
    Debug.Print "テスト実行中にエラーが発生しました: " & Err.Description
End Sub

' CSVファイル取り込み機能のテスト
Public Sub TestCSVImport()
    On Error GoTo ErrorHandler
    
    Debug.Print "CSVファイル取り込み機能のテストを開始します..."
    
    ' テスト用CSVファイルの作成
    Dim testFolderPath As String
    testFolderPath = CreateTestCSVFiles()
    
    ' CSVファイル取り込み
    ShelfManager_new.ImportCSVFiles testFolderPath
    
    ' 結果の検証 - 設定シートにデータが正しく取り込まれたか
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Sheets("設定")
    
    Dim lastRow As Long
    lastRow = settingsSheet.Cells(settingsSheet.Rows.Count, "A").End(xlUp).Row
    
    Debug.Print "取り込まれた行数: " & (lastRow - 6) & " 行"
    Debug.Print "最初のGTIN: " & settingsSheet.Cells(7, 1).Value
    
    ' テスト完了メッセージ
    Debug.Print "CSVファイル取り込み機能のテストが完了しました。"
    Exit Sub
    
ErrorHandler:
    Debug.Print "テスト中にエラーが発生しました: " & Err.Description
End Sub

' 医薬品名マッチング機能のテスト
Public Sub TestDrugNameMatching()
    On Error GoTo ErrorHandler
    
    Debug.Print "医薬品名マッチング機能のテストを開始します..."
    
    ' テスト用のGTINコード
    Dim testGTIN As String
    testGTIN = "14912345678901" ' 実際のテストでは有効なコードに置き換える
    
    ' 医薬品名を取得
    Dim drugName As String
    drugName = ShelfManager_new.GetDrugName(testGTIN)
    
    Debug.Print "GTIN: " & testGTIN
    Debug.Print "取得された医薬品名: " & drugName
    
    ' tmp_tanaでの検索テスト
    Dim matchRow As Long
    matchRow = FindMedicineRowByName(drugName)
    
    Debug.Print "tmp_tanaでのマッチ行: " & matchRow
    
    ' テスト完了メッセージ
    Debug.Print "医薬品名マッチング機能のテストが完了しました。"
    Exit Sub
    
ErrorHandler:
    Debug.Print "テスト中にエラーが発生しました: " & Err.Description
End Sub

' 棚名更新機能のテスト
Public Sub TestShelfNameUpdate()
    On Error GoTo ErrorHandler
    
    Debug.Print "棚名更新機能のテストを開始します..."
    
    ' テスト用の棚名を設定
    Dim testShelfNames(1 To 3) As String
    testShelfNames(1) = "A-01"
    testShelfNames(2) = "B-02"
    testShelfNames(3) = "C-03"
    
    ' 設定シートに棚名を設定
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Sheets("設定")
    
    settingsSheet.Cells(1, 2).Value = testShelfNames(1)
    settingsSheet.Cells(2, 2).Value = testShelfNames(2)
    settingsSheet.Cells(3, 2).Value = testShelfNames(3)
    
    ' 棚名が正しく設定されたか確認
    Debug.Print "設定された棚名1: " & settingsSheet.Cells(1, 2).Value
    Debug.Print "設定された棚名2: " & settingsSheet.Cells(2, 2).Value
    Debug.Print "設定された棚名3: " & settingsSheet.Cells(3, 2).Value
    
    ' 棚名更新機能のテスト
    ' ShelfManager_new.UpdateShelfNames testShelfNames
    
    ' テスト完了メッセージ
    Debug.Print "棚名更新機能のテストが完了しました。"
    Exit Sub
    
ErrorHandler:
    Debug.Print "テスト中にエラーが発生しました: " & Err.Description
End Sub

' 元に戻す機能のテスト
Public Sub TestUndoFunctionality()
    On Error GoTo ErrorHandler
    
    Debug.Print "元に戻す機能のテストを開始します..."
    
    ' テスト用のバックアップデータを作成
    Dim tmpTanaSheet As Worksheet
    Set tmpTanaSheet = ThisWorkbook.Sheets("tmp_tana")
    
    ' 元のデータをバックアップ
    Dim originalValue As String
    originalValue = tmpTanaSheet.Cells(2, 1).Value
    
    ' データを変更
    tmpTanaSheet.Cells(2, 1).Value = "テスト値"
    Debug.Print "変更後の値: " & tmpTanaSheet.Cells(2, 1).Value
    
    ' 元に戻す機能を実行
    ShelfManager_new.UndoChanges
    
    ' 結果を確認
    Debug.Print "元に戻した後の値: " & tmpTanaSheet.Cells(2, 1).Value
    
    ' 元の値に戻す
    tmpTanaSheet.Cells(2, 1).Value = originalValue
    
    ' テスト完了メッセージ
    Debug.Print "元に戻す機能のテストが完了しました。"
    Exit Sub
    
ErrorHandler:
    Debug.Print "テスト中にエラーが発生しました: " & Err.Description
    
    ' エラーが発生した場合も元の値に戻す
    tmpTanaSheet.Cells(2, 1).Value = originalValue
End Sub

' テスト用CSVファイルを作成する関数
Private Function CreateTestCSVFiles() As String
    On Error GoTo ErrorHandler
    
    ' テスト用フォルダパス
    Dim testFolderPath As String
    testFolderPath = ThisWorkbook.Path & "\test_csv"
    
    ' フォルダが存在しない場合は作成
    If Dir(testFolderPath, vbDirectory) = "" Then
        MkDir testFolderPath
    End If
    
    ' テスト用CSVファイルのパス
    Dim testFilePath As String
    testFilePath = testFolderPath & "\test_gtin.csv"
    
    ' CSVファイルを作成
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open testFilePath For Output As #fileNum
    
    ' ヘッダー行
    Print #fileNum, "GTIN,数量,備考"
    
    ' テストデータ行
    Print #fileNum, "14912345678901,10,テスト医薬品1"
    Print #fileNum, "14912345678902,5,テスト医薬品2"
    Print #fileNum, "14912345678903,3,テスト医薬品3"
    
    Close #fileNum
    
    Debug.Print "テスト用CSVファイルを作成しました: " & testFilePath
    
    CreateTestCSVFiles = testFolderPath
    Exit Function
    
ErrorHandler:
    Debug.Print "テスト用CSVファイル作成中にエラーが発生しました: " & Err.Description
    CreateTestCSVFiles = ""
End Function

' 医薬品名からtmp_tanaシートの行を検索する関数
Private Function FindMedicineRowByName(drugName As String) As Long
    On Error GoTo ErrorHandler
    
    Dim tmpTanaSheet As Worksheet
    Set tmpTanaSheet = ThisWorkbook.Sheets("tmp_tana")
    
    Dim lastRow As Long
    lastRow = tmpTanaSheet.Cells(tmpTanaSheet.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow ' ヘッダー行をスキップ
        If InStr(1, tmpTanaSheet.Cells(i, "A").Value, drugName, vbTextCompare) > 0 Then
            FindMedicineRowByName = i
            Exit Function
        End If
    Next i
    
    ' 見つからなかった場合
    FindMedicineRowByName = 0
    Exit Function
    
ErrorHandler:
    Debug.Print "行検索中にエラーが発生しました: " & Err.Description
    FindMedicineRowByName = -1
End Function
