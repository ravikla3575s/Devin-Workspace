Attribute VB_Name = "ShelfManager"
Option Explicit

' グローバル変数 - 元の棚名データを保持（Undo用）
Private originalShelfNames As Variant

' フォルダ内のCSVファイル数をカウントする
Public Function CountCSVFiles(folderPath As String) As Integer
    On Error GoTo ErrorHandler
    
    Dim fileName As String
    Dim count As Integer
    
    count = 0
    
    ' フォルダ内のCSVファイルを検索
    fileName = Dir(folderPath & "\*.csv")
    
    ' CSVファイルがない場合は0を返す
    If fileName = "" Then
        CountCSVFiles = 0
        Exit Function
    End If
    
    ' 各CSVファイルをカウント
    Do While fileName <> ""
        count = count + 1
        fileName = Dir
    Loop
    
    CountCSVFiles = count
    Exit Function
    
ErrorHandler:
    MsgBox "CSVファイル数のカウント中にエラーが発生しました: " & Err.Description, vbCritical
    CountCSVFiles = 0
End Function

' フォルダ内のCSVファイル名を取得する
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
    MsgBox "CSVファイル名の取得中にエラーが発生しました: " & Err.Description, vbCritical
    GetCSVFileNames = Array()
End Function

' メインエントリポイント - 棚番一括更新マクロ
Public Sub Main()
    On Error GoTo ErrorHandler
    
    Dim folderPath As String
    Dim outputPath As String
    Dim fileCount As Integer
    Dim fileNames As Variant
    
    ' フォルダ選択ダイアログを表示
    folderPath = GetFolderPath()
    If folderPath = "" Then
        MsgBox "フォルダが選択されていないため、処理を中止します。", vbExclamation
        Exit Sub
    End If
    
    ' CSVファイル数をカウント
    fileCount = CountCSVFiles(folderPath)
    
    ' ファイルがない場合は処理中止
    If fileCount = 0 Then
        MsgBox "指定フォルダにCSVファイルが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' CSVファイル名を取得
    fileNames = GetCSVFileNames(folderPath, fileCount)
    
    ' 動的ユーザーフォームを表示（棚名入力）
    DynamicShelfNameForm.SetFileCount fileCount, fileNames
    DynamicShelfNameForm.Show
    
    ' キャンセルされた場合は処理中止
    If DynamicShelfNameForm.IsCancelled Then
        Exit Sub
    End If
    
    ' 元の棚名データを保存（Undo用）
    SaveOriginalShelfNames
    
    ' CSVファイルを取り込み
    ImportCSVFiles folderPath
    
    ' 設定シート上のGTIN一覧を処理
    ProcessItems
    
    ' tmp_tanaシートをCSV保存
    ExportTemplateCSV
    
    ' 使用したファイルパスを取得
    outputPath = GetTemplateOutputPath()
    If outputPath = "" Then
        outputPath = ThisWorkbook.Path & "\update_tmp_tana.csv"
    End If
    
    ' 完了メッセージ
    MsgBox "処理が完了しました。" & vbCrLf & "ファイル: " & outputPath, vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

' フォルダ選択ダイアログを表示し、選択されたフォルダパスを返す
Private Function GetFolderPath() As String
    Dim folderDialog As FileDialog
    
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    With folderDialog
        .Title = "GTINコードCSVファイルが保存されているフォルダを選択してください"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            GetFolderPath = .SelectedItems(1)
        Else
            GetFolderPath = ""
        End If
    End With
End Function

' 指定フォルダ内のCSVファイルを読み込み、設定シートにGTINコードを展開する
Public Sub ImportCSVFiles(folderPath As String)
    On Error GoTo ErrorHandler
    
    Dim fileName As String
    Dim filePath As String
    Dim fileNum As Integer
    Dim line As String
    Dim row As Long
    Dim colIndex As Long
    Dim csvCount As Integer
    Dim invalidCodes As New Collection
    Dim maxFiles As Integer
    
    ' 設定シートを取得
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Sheets("設定")
    
    ' 設定シートの既存データをクリア（A7以降）
    Dim lastRow As Long
    lastRow = settingsSheet.Cells(settingsSheet.Rows.Count, "A").End(xlUp).row
    If lastRow >= 7 Then
        ' 列数を拡張（最大10ファイル対応）
        settingsSheet.Range("A7:M" & lastRow).ClearContents
    End If
    
    ' 開始行を設定
    row = 7
    csvCount = 0
    
    ' 最大ファイル数（設定シートの制限を考慮）
    maxFiles = 10
    
    ' フォルダ内のCSVファイルを検索
    fileName = Dir(folderPath & "\*.csv")
    
    ' CSVファイルが見つからない場合
    If fileName = "" Then
        MsgBox "指定フォルダにCSVファイルが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 進捗状況表示
    Application.StatusBar = "CSVファイルを読み込んでいます..."
    
    ' 各CSVファイルを処理
    Do While fileName <> ""
        ' CSVファイルのカウントを増やす
        csvCount = csvCount + 1
        
        ' 最大ファイル数を超えた場合は処理を中止
        If csvCount > maxFiles Then
            MsgBox "警告: " & maxFiles + 1 & "つ以上のCSVファイルが見つかりました。最初の" & maxFiles & "つのみ処理します。", vbExclamation
            Exit Do
        End If
        
        ' 対応する棚名列を決定（D=棚名1, E=棚名2, F=棚名3...）
        colIndex = 3 + csvCount  ' D=4, E=5, F=6...
        
        ' CSVファイルのフルパス
        filePath = folderPath & "\" & fileName
        
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
                    ' 設定シートにGTINコードを書き込む
                    settingsSheet.Cells(row, 1).Value = line
                    
                    ' 対応する棚名を書き込む（設定シートB1〜B10から取得）
                    settingsSheet.Cells(row, colIndex).Value = settingsSheet.Cells(csvCount, 2).Value
                    
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
        
        ' 次のCSVファイルを検索
        fileName = Dir
    Loop
    
    ' 無効なGTINコードがあれば報告
    If invalidCodes.Count > 0 Then
        Dim message As String
        Dim i As Integer
        
        message = "以下の" & invalidCodes.Count & "件のコードは14桁の数字ではないため、処理対象外としました:" & vbCrLf & vbCrLf
        
        For i = 1 To invalidCodes.Count
            If i <= 10 Then
                message = message & invalidCodes(i) & vbCrLf
            Else
                message = message & "... 他 " & (invalidCodes.Count - 10) & " 件"
                Exit For
            End If
        Next i
        
        MsgBox message, vbExclamation
    End If
    
    ' 進捗状況表示をクリア
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    Close #fileNum
    Application.StatusBar = False
    MsgBox "CSVファイルの読み込み中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' GTIN14コードが有効かチェック（14桁の数字であること）
Private Function IsValidGTIN14(code As String) As Boolean
    ' 14桁の数字かどうかをチェック
    IsValidGTIN14 = (Len(code) = 14) And IsNumeric(code)
End Function

' 設定シート上のGTIN一覧を処理し、tmp_tanaシートを更新する
Public Sub ProcessItems()
    On Error GoTo ErrorHandler
    
    Dim settingsSheet As Worksheet
    Dim row As Long
    Dim lastRow As Long
    Dim gtin As String
    Dim drugName As String
    Dim matchRow As Long
    Dim shelf1 As String
    Dim shelf2 As String
    Dim shelf3 As String
    Dim notFoundItems As New Collection
    Dim multipleMatchItems As New Collection
    
    ' 設定シートを取得
    Set settingsSheet = ThisWorkbook.Sheets("設定")
    
    ' 最終行を取得
    lastRow = settingsSheet.Cells(settingsSheet.Rows.Count, "A").End(xlUp).row
    
    ' 処理対象がない場合
    If lastRow < 7 Then
        MsgBox "処理対象のGTINコードがありません。", vbExclamation
        Exit Sub
    End If
    
    ' 進捗状況表示
    Application.StatusBar = "医薬品情報を処理しています..."
    
    ' 画面更新を停止（パフォーマンス向上）
    Application.ScreenUpdating = False
    
    ' 設定シートA7から最終行までループ処理
    For row = 7 To lastRow
        ' GTINコードを取得
        gtin = settingsSheet.Cells(row, 1).Value
        
        ' GTINコードが空なら次の行へ
        If gtin = "" Then
            GoTo NextRow
        End If
        
        ' 医薬品名を取得
        drugName = GetDrugName(gtin)
        
        ' 医薬品名を設定シートB列に書き込む
        settingsSheet.Cells(row, 2).Value = drugName
        
        ' 医薬品名が取得できなかった場合は次の行へ
        If drugName = "" Then
            settingsSheet.Cells(row, 2).Value = "未登録"
            GoTo NextRow
        End If
        
        ' tmp_tanaシートで医薬品名に部分一致する行を検索
        matchRow = FindMedicineRowByName(drugName)
        
        ' 見つからなかった場合
        If matchRow = -1 Then
            On Error Resume Next
            notFoundItems.Add drugName
            On Error GoTo ErrorHandler
            GoTo NextRow
        End If
        
        ' 複数見つかった場合
        If matchRow = -2 Then
            On Error Resume Next
            multipleMatchItems.Add drugName
            On Error GoTo ErrorHandler
            GoTo NextRow
        End If
        
        ' 棚名を取得（設定シートB1〜B3）
        shelf1 = settingsSheet.Cells(1, 2).Value
        shelf2 = settingsSheet.Cells(2, 2).Value
        shelf3 = settingsSheet.Cells(3, 2).Value
        
        ' 棚名を更新
        OverwriteShelfNames matchRow, shelf1, shelf2, shelf3
        
NextRow:
        ' 進捗状況を更新
        If row Mod 10 = 0 Then
            Application.StatusBar = "医薬品情報を処理しています... (" & row - 6 & "/" & lastRow - 6 & ")"
            DoEvents
        End If
    Next row
    
    ' 画面更新を再開
    Application.ScreenUpdating = True
    
    ' 進捗状況表示をクリア
    Application.StatusBar = False
    
    ' 見つからなかった医薬品があれば報告
    If notFoundItems.Count > 0 Then
        Dim message As String
        Dim i As Integer
        
        message = "以下の" & notFoundItems.Count & "件の医薬品は棚番テンプレート上で見つかりませんでした:" & vbCrLf & vbCrLf
        
        For i = 1 To notFoundItems.Count
            If i <= 10 Then
                message = message & notFoundItems(i) & vbCrLf
            Else
                message = message & "... 他 " & (notFoundItems.Count - 10) & " 件"
                Exit For
            End If
        Next i
        
        MsgBox message, vbExclamation
    End If
    
    ' 複数見つかった医薬品があれば報告
    If multipleMatchItems.Count > 0 Then
        Dim message2 As String
        
        message2 = "以下の" & multipleMatchItems.Count & "件の医薬品は複数候補があり更新保留としました:" & vbCrLf & vbCrLf
        
        For i = 1 To multipleMatchItems.Count
            If i <= 10 Then
                message2 = message2 & multipleMatchItems(i) & vbCrLf
            Else
                message2 = message2 & "... 他 " & (multipleMatchItems.Count - 10) & " 件"
                Exit For
            End If
        Next i
        
        MsgBox message2, vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "医薬品情報の処理中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 医薬品コードシートからGTINコードに対応する医薬品名を取得する
Private Function GetDrugName(gtin As String) As String
    On Error GoTo ErrorHandler
    
    ' GS1CodeProcessorを使用してGTIN-14コードから医薬品情報を取得
    ' （これはすでにSheet3を使用するように修正されている）
    Dim drugInfo As DrugInfo
    drugInfo = GS1CodeProcessor.GetDrugInfoFromGS1Code(gtin)
    
    ' 結果を返す
    GetDrugName = drugInfo.DrugName
    
    Exit Function
    
ErrorHandler:
    GetDrugName = ""
End Function

' tmp_tanaシートで医薬品名を検索し、行番号を返す
Private Function FindMedicineRowByName(drugName As String) As Long
    On Error GoTo ErrorHandler
    
    Dim tmpTanaSheet As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim keywords As Variant
    Dim keyword As Variant
    Dim cellValue As String
    Dim matchCount As Integer
    Dim firstMatchRow As Long
    Dim allKeywordsMatch As Boolean
    
    ' tmp_tanaシートを取得
    Set tmpTanaSheet = ThisWorkbook.Sheets("tmp_tana")
    
    ' 最終行を取得
    lastRow = tmpTanaSheet.Cells(tmpTanaSheet.Rows.Count, "B").End(xlUp).row
    
    ' 医薬品名をキーワードに分割
    keywords = SplitDrugName(drugName)
    
    matchCount = 0
    firstMatchRow = 0
    
    ' 各行をチェック
    For row = 1 To lastRow
        cellValue = tmpTanaSheet.Cells(row, 2).Value
        
        ' セル値が空でなければ検索
        If cellValue <> "" Then
            ' すべてのキーワードが含まれるかチェック
            allKeywordsMatch = True
            
            For Each keyword In keywords
                If InStr(1, cellValue, CStr(keyword), vbTextCompare) = 0 Then
                    allKeywordsMatch = False
                    Exit For
                End If
            Next keyword
            
            ' すべてのキーワードが含まれる場合
            If allKeywordsMatch Then
                matchCount = matchCount + 1
                
                ' 最初のマッチを記録
                If firstMatchRow = 0 Then
                    firstMatchRow = row
                End If
                
                ' 2つ以上マッチした場合は複数マッチとして-2を返す
                If matchCount > 1 Then
                    FindMedicineRowByName = -2
                    Exit Function
                End If
            End If
        End If
    Next row
    
    ' マッチ結果に応じて戻り値を設定
    If matchCount = 0 Then
        FindMedicineRowByName = -1  ' 見つからない
    ElseIf matchCount = 1 Then
        FindMedicineRowByName = firstMatchRow  ' 一意に特定
    Else
        FindMedicineRowByName = -2  ' 複数マッチ
    End If
    
    Exit Function
    
ErrorHandler:
    FindMedicineRowByName = -1
End Function

' 医薬品名を検索キーワードに分割する
Private Function SplitDrugName(drugName As String) As Variant
    Dim result As Variant
    Dim tempName As String
    
    ' 全角スペース、半角スペース、括弧などで分割
    tempName = Replace(drugName, "　", " ")
    tempName = Replace(tempName, "（", " ")
    tempName = Replace(tempName, "）", " ")
    tempName = Replace(tempName, "(", " ")
    tempName = Replace(tempName, ")", " ")
    tempName = Replace(tempName, "「", " ")
    tempName = Replace(tempName, "」", " ")
    
    ' 連続スペースを1つに置換
    Do While InStr(tempName, "  ") > 0
        tempName = Replace(tempName, "  ", " ")
    Loop
    
    ' スペースで分割
    result = Split(tempName, " ")
    
    ' 空の要素を除外
    Dim i As Integer
    Dim j As Integer
    Dim count As Integer
    
    count = UBound(result) - LBound(result) + 1
    For i = LBound(result) To UBound(result)
        If Trim(result(i)) = "" Then
            For j = i To UBound(result) - 1
                result(j) = result(j + 1)
            Next j
            count = count - 1
            i = i - 1
        End If
    Next i
    
    ReDim Preserve result(0 To count - 1)
    
    SplitDrugName = result
End Function

' tmp_tanaシートの指定行に棚名を書き込む
Private Sub OverwriteShelfNames(row As Long, shelf1 As String, shelf2 As String, shelf3 As String)
    On Error GoTo ErrorHandler
    
    Dim tmpTanaSheet As Worksheet
    
    ' tmp_tanaシートを取得
    Set tmpTanaSheet = ThisWorkbook.Sheets("tmp_tana")
    
    ' 棚名1（G列）を更新
    If shelf1 <> "" Then
        tmpTanaSheet.Cells(row, 7).Value = shelf1
    End If
    
    ' 棚名2（H列）を更新
    If shelf2 <> "" Then
        tmpTanaSheet.Cells(row, 8).Value = shelf2
    End If
    
    ' 棚名3（I列）を更新
    If shelf3 <> "" Then
        tmpTanaSheet.Cells(row, 9).Value = shelf3
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "棚名の更新中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 元の棚名データを保存（Undo用）
Public Sub SaveOriginalShelfNames()
    On Error GoTo ErrorHandler
    
    Dim tmpTanaSheet As Worksheet
    Dim lastRow As Long
    
    ' tmp_tanaシートを取得
    Set tmpTanaSheet = ThisWorkbook.Sheets("tmp_tana")
    
    ' 最終行を取得
    lastRow = tmpTanaSheet.Cells(tmpTanaSheet.Rows.Count, "B").End(xlUp).row
    
    ' G〜I列のデータを配列に保存
    originalShelfNames = tmpTanaSheet.Range("G1:I" & lastRow).Value
    
    Exit Sub
    
ErrorHandler:
    MsgBox "元の棚名データの保存中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 元の棚名データに戻す（Undo）
Public Sub UndoShelfNames()
    On Error GoTo ErrorHandler
    
    Dim tmpTanaSheet As Worksheet
    Dim rowCount As Long
    
    ' 保存データがない場合
    If Not IsArray(originalShelfNames) Then
        MsgBox "元に戻すデータがありません。", vbExclamation
        Exit Sub
    End If
    
    ' tmp_tanaシートを取得
    Set tmpTanaSheet = ThisWorkbook.Sheets("tmp_tana")
    
    ' 行数を取得
    rowCount = UBound(originalShelfNames, 1)
    
    ' 保存データを書き戻す
    tmpTanaSheet.Range("G1").Resize(rowCount, 3).Value = originalShelfNames
    
    ' 完了メッセージ
    MsgBox "棚名を元の状態に戻しました。", vbInformation
    
    ' 保存データをクリア
    Erase originalShelfNames
    
    Exit Sub
    
ErrorHandler:
    MsgBox "棚名の復元中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' tmp_tanaシートをCSVファイルに保存する
Public Sub ExportTemplateCSV(Optional filePath As String = "")
    On Error GoTo ErrorHandler
    
    Dim tmpTanaSheet As Worksheet
    Dim defaultPath As String
    Dim timestamp As String
    Dim savedPath As String
    
    ' tmp_tanaシートを取得
    Set tmpTanaSheet = ThisWorkbook.Sheets("tmp_tana")
    
    ' ファイルパスが指定されていない場合
    If filePath = "" Then
        ' 設定シートのB4セルからパスを取得
        savedPath = GetTemplateOutputPath()
        
        ' 保存されたパスが空または無効な場合はデフォルトパスを使用
        If savedPath = "" Then
            ' タイムスタンプを生成（YYYYMMDD_HHMM形式）
            timestamp = Format(Now, "YYYYMMDD_HHMM")
            
            ' デフォルトパスを設定
            defaultPath = ThisWorkbook.Path & "\update_tmp_tana_" & timestamp & ".csv"
            filePath = defaultPath
        Else
            filePath = savedPath
        End If
    End If
    
    ' 上書き確認を抑制
    Application.DisplayAlerts = False
    
    ' 一時的にtmp_tanaシートをアクティブにする
    tmpTanaSheet.Activate
    
    ' CSVとして保存
    ActiveWorkbook.SaveAs filePath, xlCSV
    
    ' 元のブックを保存し直す（CSV形式にならないように）
    ActiveWorkbook.SaveAs ThisWorkbook.Path & "\temp_backup.xlsm", xlOpenXMLWorkbookMacroEnabled
    
    ' 元のブックを開き直す
    Workbooks.Open ThisWorkbook.Path & "\temp_backup.xlsm"
    
    ' 一時ファイルを削除
    Kill ThisWorkbook.Path & "\temp_backup.xlsm"
    
    ' 上書き確認を再度有効化
    Application.DisplayAlerts = True
    
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    MsgBox "CSVファイルの保存中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 設定シートのB4セルから出力先パスを取得する
Private Function GetTemplateOutputPath() As String
    On Error GoTo ErrorHandler
    
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Sheets("設定")
    
    ' B4セルからパスを取得（空の場合はデフォルトパスを使用）
    GetTemplateOutputPath = Trim(settingsSheet.Range("B4").Value)
    
    Exit Function
    
ErrorHandler:
    GetTemplateOutputPath = ""
End Function

' 設定シートのB4セルに出力先パスを保存する
Public Sub SaveTemplateOutputPath(ByVal filePath As String)
    On Error GoTo ErrorHandler
    
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Sheets("設定")
    
    ' B4セルにパスを保存
    settingsSheet.Range("B4").Value = filePath
    
    Exit Sub
    
ErrorHandler:
    MsgBox "出力先パスの保存中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 出力ファイルパスを設定するダイアログを表示する
Public Sub SetOutputFilePath()
    On Error GoTo ErrorHandler
    
    Dim currentPath As String
    Dim newPath As String
    Dim fdDialog As FileDialog
    
    ' 現在の設定を取得
    currentPath = GetTemplateOutputPath()
    
    ' ファイル選択ダイアログを表示
    Set fdDialog = Application.FileDialog(msoFileDialogSaveAs)
    
    With fdDialog
        .Title = "テンプレートファイルの出力先を選択してください"
        .InitialFileName = IIf(currentPath <> "", currentPath, ThisWorkbook.Path & "\update_tmp_tana.csv")
        ' Filtersプロパティは一部の環境でサポートされていないため削除
        
        If .Show = -1 Then
            newPath = .SelectedItems(1)
            SaveTemplateOutputPath newPath
            MsgBox "出力先を設定しました: " & newPath, vbInformation
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "出力先の設定中にエラーが発生しました: " & Err.Description, vbCritical
End Sub


