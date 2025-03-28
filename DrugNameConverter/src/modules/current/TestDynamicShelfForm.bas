Attribute VB_Name = "TestDynamicShelfForm"
Option Explicit

' 動的棚名フォームのテスト（様々なファイル数でテスト）
Public Sub TestDynamicShelfForm()
    On Error GoTo ErrorHandler
    
    Dim fileCount As Integer
    Dim fileNames() As String
    
    ' テスト開始メッセージ
    Debug.Print "動的棚名入力フォームのテストを開始します..."
    
    ' 様々なファイル数でテスト
    For fileCount = 1 To 12  ' 最大値を超えるテストも含む
        Debug.Print "テスト: " & fileCount & "ファイル"
        
        ' テスト用のファイル名配列を作成
        ReDim fileNames(1 To fileCount)
        Dim i As Integer
        For i = 1 To fileCount
            fileNames(i) = "test_file_" & i & ".csv"
        Next i
        
        ' 動的ユーザーフォームを表示
        DynamicShelfNameForm.SetFileCount fileCount, fileNames
        DynamicShelfNameForm.Show
        
        ' キャンセルされた場合は次のテストへ
        If DynamicShelfNameForm.IsCancelled Then
            Debug.Print "  キャンセルされました"
        Else
            Debug.Print "  OKがクリックされました"
            
            ' 設定シートの値を確認
            Dim settingsSheet As Worksheet
            Set settingsSheet = ThisWorkbook.Sheets("設定")
            
            For i = 1 To fileCount
                If i <= 10 Then  ' 設定シートの制限を考慮
                    Debug.Print "  棚名" & i & ": " & settingsSheet.Cells(i, 2).Value
                End If
            Next i
        End If
        
        Debug.Print "-------------------"
    Next fileCount
    
    ' テスト完了メッセージ
    Debug.Print "動的棚名入力フォームのテストが完了しました。"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "テスト中にエラーが発生しました: " & Err.Description
End Sub

' 実際のCSVファイルを使用したテスト
Public Sub TestWithRealCSVFiles()
    On Error GoTo ErrorHandler
    
    Dim folderPath As String
    Dim fileCount As Integer
    Dim fileNames As Variant
    
    ' テスト開始メッセージ
    Debug.Print "実際のCSVファイルを使用したテストを開始します..."
    
    ' フォルダ選択ダイアログを表示
    folderPath = ShelfManager.GetFolderPath()
    If folderPath = "" Then
        Debug.Print "フォルダが選択されていないため、テストを中止します。"
        Exit Sub
    End If
    
    ' CSVファイル数をカウント
    fileCount = ShelfManager.CountCSVFiles(folderPath)
    
    ' ファイルがない場合はテスト中止
    If fileCount = 0 Then
        Debug.Print "指定フォルダにCSVファイルが見つかりません。"
        Exit Sub
    End If
    
    ' CSVファイル名を取得
    fileNames = GetCSVFileNames(folderPath, fileCount)
    
    ' ファイル情報を表示
    Debug.Print "フォルダ: " & folderPath
    Debug.Print "CSVファイル数: " & fileCount
    
    Dim i As Integer
    For i = 1 To UBound(fileNames)
        Debug.Print "ファイル" & i & ": " & fileNames(i)
    Next i
    
    ' 動的ユーザーフォームを表示
    DynamicShelfNameForm.SetFileCount fileCount, fileNames
    DynamicShelfNameForm.Show
    
    ' キャンセルされた場合はテスト中止
    If DynamicShelfNameForm.IsCancelled Then
        Debug.Print "キャンセルされました。"
        Exit Sub
    End If
    
    ' 設定シートの値を確認
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Sheets("設定")
    
    Debug.Print "設定シートに保存された棚名:"
    For i = 1 To fileCount
        If i <= 10 Then  ' 設定シートの制限を考慮
            Debug.Print "棚名" & i & ": " & settingsSheet.Cells(i, 2).Value
        End If
    Next i
    
    ' テスト完了メッセージ
    Debug.Print "実際のCSVファイルを使用したテストが完了しました。"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "テスト中にエラーが発生しました: " & Err.Description
End Sub

' CSVファイル名を取得する（ShelfManagerのプライベート関数を再実装）
Private Function GetCSVFileNames(ByVal folderPath As String, ByVal maxCount As Integer) As Variant
    On Error GoTo ErrorHandler
    
    Dim fileNames() As String
    Dim fileName As String
    Dim i As Integer
    
    ReDim fileNames(1 To maxCount)
    i = 1
    
    ' フォルダ内のCSVファイルを検索
    fileName = Dir(folderPath & "\*.csv")
    
    ' 各CSVファイル名を配列に格納
    Do While fileName <> "" And i <= maxCount
        fileNames(i) = fileName
        i = i + 1
        fileName = Dir
    Loop
    
    GetCSVFileNames = fileNames
    Exit Function
    
ErrorHandler:
    Debug.Print "CSVファイル名の取得中にエラーが発生しました: " & Err.Description
    GetCSVFileNames = Array()
End Function

' スクロール機能のテスト
Public Sub TestScrollFunctionality()
    On Error GoTo ErrorHandler
    
    Dim fileCount As Integer
    Dim fileNames() As String
    
    ' テスト開始メッセージ
    Debug.Print "スクロール機能のテストを開始します..."
    
    ' スクロールが必要になる数のファイルでテスト
    fileCount = 10
    
    ' テスト用のファイル名配列を作成
    ReDim fileNames(1 To fileCount)
    Dim i As Integer
    For i = 1 To fileCount
        fileNames(i) = "scroll_test_" & i & ".csv"
    Next i
    
    ' 動的ユーザーフォームを表示
    DynamicShelfNameForm.SetFileCount fileCount, fileNames
    DynamicShelfNameForm.Show
    
    ' テスト手順を表示
    Debug.Print "スクロールテスト手順:"
    Debug.Print "1. マウスホイールを使用してフォームをスクロールしてください"
    Debug.Print "2. 最下部までスクロールできることを確認してください"
    Debug.Print "3. 最上部に戻れることを確認してください"
    Debug.Print "4. 各テキストボックスに入力できることを確認してください"
    Debug.Print "5. OKボタンをクリックして設定が保存されることを確認してください"
    
    ' キャンセルされた場合はテスト中止
    If DynamicShelfNameForm.IsCancelled Then
        Debug.Print "キャンセルされました。"
        Exit Sub
    End If
    
    ' テスト完了メッセージ
    Debug.Print "スクロール機能のテストが完了しました。"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "テスト中にエラーが発生しました: " & Err.Description
End Sub
