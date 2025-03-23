Option Explicit

' 定数定義
' 包装形態の定義
Private Const PACKAGE_TYPE_PTP As String = "PTP"
Private Const PACKAGE_TYPE_BULK As String = "バラ"
Private Const PACKAGE_TYPE_UNIT_DOSE As String = "分包"
Private Const PACKAGE_TYPE_SP As String = "SP"
Private Const PACKAGE_TYPE_SMALL As String = "包装小"
Private Const PACKAGE_TYPE_OTHER As String = "その他"
Private Const PACKAGE_TYPE_UNKNOWN As String = "不明"
Private Const PACKAGE_TYPE_PTP_PATIENT As String = "PTP(患者用)"
Private Const PACKAGE_TYPE_DISPENSING As String = "調剤用"

' インデックス定義（配列のインデックスを明確に）
Private Const INDEX_PTP As Long = 0
Private Const INDEX_BULK As Long = 1
Private Const INDEX_UNIT_DOSE As Long = 2
Private Const INDEX_SP As Long = 3
Private Const INDEX_SMALL As Long = 4
Private Const INDEX_OTHER As Long = 5
Private Const INDEX_UNKNOWN As Long = 6

' カテゴリの総数（配列サイズ用の定数）
Private Const CATEGORY_COUNT As Long = 7

' 処理バッチサイズ
Private Const BATCH_SIZE As Long = 50  ' 一度に処理する薬品数

' 医薬品マスターシートの設定（MainModuleからの移行）
Private Const MASTER_SHEET_NAME As String = "薬品マスター"
Private Const DRUG_CODE_COLUMN As Integer = 1  ' A列
Private Const DRUG_NAME_COLUMN As Integer = 2  ' B列
Private Const FIRST_DATA_ROW As Integer = 2    ' 2行目からデータ開始

' ラッパーモジュール - 基本機能 (Mac版)
' Mac版ではMSFormsが利用できないため、ステータスバーを使用した進捗表示に変更しています

' メイン処理を呼び出すラッパー関数
Public Sub RunDrugNameComparison()
    ' 新しい包装形態別処理を呼び出す
    RunPackageBasedProcessing
End Sub

' 新しい包装形態別処理の実行関数
Public Sub RunPackageBasedProcessing()
    On Error GoTo ErrorHandler
    
    ' アプリケーションの設定を最適化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' ワークシートの取得
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' 最終行の取得
    Dim lastRow As Long
    lastRow = settingsSheet.Cells(settingsSheet.Rows.Count, "A").End(xlUp).Row
    
    ' 医薬品コードが存在しなければエラー
    If lastRow < 7 Then
        MsgBox "医薬品コードが存在しません。", vbExclamation
        GoTo CleanupAndExit
    End If
    
    ' 処理開始
    PackageBasedProcessing
    
CleanupAndExit:
    ' アプリケーションの設定を元に戻す
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description & " (" & Err.Number & ")", vbCritical
    Debug.Print "エラー発生: " & Err.Number & " - " & Err.Description & " - " & Erl
    Resume CleanupAndExit
End Sub

' 包装形態別に医薬品を処理する関数（自動マスターシート作成機能付き）
Public Sub PackageBasedProcessing()
    On Error GoTo ErrorHandler
    
    ' 処理の開始時間を記録
    Dim startTime As Double
    startTime = Timer
    
    Debug.Print "===== PackageBasedProcessing 開始 ====="
    
    ' 最初にマスターシートを確実に作成（または検証）
    Debug.Print "マスターシートの自動確認/作成を行います"
    CreateTestDrugMaster
    
    ' 確認のためシート一覧を表示
    Debug.Print "ワークブック内のシート一覧:"
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print " - " & ws.Name
    Next ws
    
    ' ステータスバーに進捗を表示
    Application.StatusBar = "医薬品コードから医薬品名を取得しています..."
    
    ' ワークシートの取得
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' 最終行の取得
    Dim lastRow As Long
    lastRow = settingsSheet.Cells(settingsSheet.Rows.Count, "A").End(xlUp).Row
    
    ' 医薬品コードの数を計算
    Dim codeCount As Long
    codeCount = lastRow - 6  ' 7行目から開始
    
    Debug.Print "処理開始: " & codeCount & "件の医薬品コードを処理します"
    
    ' 医薬品コードと医薬品名を取得
    Dim drugCodes() As String
    Dim drugNames() As String
    ReDim drugCodes(1 To codeCount)
    ReDim drugNames(1 To codeCount)
    
    Dim i As Long, j As Long
    Dim currentRow As Long
    
    For i = 1 To codeCount
        currentRow = i + 6  ' 7行目から開始
        
        ' 医薬品コードを取得して14桁に整形
        Dim drugCode As String
        drugCode = settingsSheet.Cells(currentRow, "A").Value
        
        If Len(drugCode) > 0 Then
            Dim formattedCode As String
            formattedCode = MainModule_Mac.FormatDrugCode(drugCode)
            
            ' 整形したコードをセルに戻す（オプション）
            settingsSheet.Cells(currentRow, "A").Value = formattedCode
            
            ' 配列に保存
            drugCodes(i) = formattedCode
            
            ' 医薬品コードから医薬品名を取得
            drugNames(i) = MainModule_Mac.FindDrugNameByCode(formattedCode)
            Debug.Print "コード[" & formattedCode & "] -> 薬品名[" & drugNames(i) & "]"
            
            ' 進捗状況を更新（10件ごと）
            If i Mod 10 = 0 Then
                Application.StatusBar = "医薬品名を取得中... " & i & "/" & codeCount
                DoEvents
            End If
        Else
            drugCodes(i) = ""
            drugNames(i) = ""
        End If
    Next i
    
    ' コードを検索しても全て[マスターシートなし]だった場合の追加対策
    Dim allMasterNotFound As Boolean
    allMasterNotFound = True
    
    For i = 1 To codeCount
        If drugNames(i) <> "[マスターシートなし]" And drugNames(i) <> "" Then
            allMasterNotFound = False
            Exit For
        End If
    Next i
    
    If allMasterNotFound And codeCount > 0 Then
        Debug.Print "警告: すべての検索結果が[マスターシートなし]です。FindDrugNameByCode関数の問題を確認します。"
        
        ' 手動でマスターシートからデータを取得して上書き
        Dim masterSheet As Worksheet
        On Error Resume Next
        Set masterSheet = ThisWorkbook.Worksheets(MASTER_SHEET_NAME)
        
        If Not masterSheet Is Nothing Then
            Debug.Print "マスターシートを直接参照してデータを取得します"
            Dim lastRowMaster As Long
            lastRowMaster = masterSheet.Cells(masterSheet.Rows.Count, DRUG_CODE_COLUMN).End(xlUp).Row
            
            ' 直接マスターシートを参照して薬品名を取得
            For i = 1 To codeCount
                If drugCodes(i) <> "" Then
                    Dim found As Boolean
                    found = False
                    
                    For j = FIRST_DATA_ROW To lastRowMaster
                        If masterSheet.Cells(j, DRUG_CODE_COLUMN).Value = drugCodes(i) Then
                            drugNames(i) = masterSheet.Cells(j, DRUG_NAME_COLUMN).Value
                            Debug.Print "直接参照によるコード[" & drugCodes(i) & "] -> 薬品名[" & drugNames(i) & "]"
                            found = True
                            Exit For
                        End If
                    Next j
                    
                    If Not found Then
                        Debug.Print "コード[" & drugCodes(i) & "]に対応する薬品名が見つかりません"
                    End If
                End If
            Next i
        Else
            Debug.Print "マスターシートが見つかりません。薬品名の検索はできません。"
        End If
        
        On Error GoTo ErrorHandler
    End If
    
    ' 包装形態ごとに分類
    Application.StatusBar = "医薬品を包装形態ごとに分類しています..."
    Debug.Print "医薬品を包装形態ごとに分類します"
    
    ' 包装形態ごとの配列を初期化
    Dim packageCategories(0 To CATEGORY_COUNT - 1) As Collection
    For i = 0 To CATEGORY_COUNT - 1
        Set packageCategories(i) = New Collection
    Next i
    
    ' 分類カウンター
    Dim categoryCounts(0 To CATEGORY_COUNT - 1) As Long
    
    ' 各医薬品を包装形態ごとに分類
    For i = 1 To codeCount
        If Len(drugNames(i)) > 0 And Left(drugNames(i), 1) <> "[" Then
            Debug.Print "包装形態判定 [" & i & "]: " & drugNames(i)
            
            ' 包装形態を検出
            Dim packageType As String
            packageType = DrugNameParser_Mac.DetectPackageType(drugNames(i))
            
            Debug.Print "判定結果: " & drugNames(i) & " -> " & packageType
            
            ' 包装形態に基づいて適切なカテゴリーに分類
            Select Case packageType
                Case PACKAGE_TYPE_PTP
                    packageCategories(INDEX_PTP).Add drugNames(i)
                    categoryCounts(INDEX_PTP) = categoryCounts(INDEX_PTP) + 1
                    Debug.Print "  → PTPに分類"
                    
                Case PACKAGE_TYPE_BULK
                    packageCategories(INDEX_BULK).Add drugNames(i)
                    categoryCounts(INDEX_BULK) = categoryCounts(INDEX_BULK) + 1
                    Debug.Print "  → バラに分類"
                    
                Case PACKAGE_TYPE_UNIT_DOSE
                    packageCategories(INDEX_UNIT_DOSE).Add drugNames(i)
                    categoryCounts(INDEX_UNIT_DOSE) = categoryCounts(INDEX_UNIT_DOSE) + 1
                    Debug.Print "  → 分包に分類"
                    
                Case PACKAGE_TYPE_SP
                    packageCategories(INDEX_SP).Add drugNames(i)
                    categoryCounts(INDEX_SP) = categoryCounts(INDEX_SP) + 1
                    Debug.Print "  → SPに分類"
                    
                Case PACKAGE_TYPE_SMALL
                    packageCategories(INDEX_SMALL).Add drugNames(i)
                    categoryCounts(INDEX_SMALL) = categoryCounts(INDEX_SMALL) + 1
                    Debug.Print "  → 包装小に分類"
                    
                Case PACKAGE_TYPE_OTHER, PACKAGE_TYPE_PTP_PATIENT, PACKAGE_TYPE_DISPENSING
                    packageCategories(INDEX_OTHER).Add drugNames(i)
                    categoryCounts(INDEX_OTHER) = categoryCounts(INDEX_OTHER) + 1
                    Debug.Print "  → その他に分類"
                    
                Case Else
                    ' 不明な包装形態
                    packageCategories(INDEX_UNKNOWN).Add drugNames(i)
                    categoryCounts(INDEX_UNKNOWN) = categoryCounts(INDEX_UNKNOWN) + 1
                    Debug.Print "  → 不明に分類"
            End Select
        ElseIf Left(drugNames(i), 1) = "[" Then
            ' エラー状態の項目（[コード未登録]など）は不明カテゴリに入れる
            packageCategories(INDEX_UNKNOWN).Add drugNames(i)
            categoryCounts(INDEX_UNKNOWN) = categoryCounts(INDEX_UNKNOWN) + 1
            Debug.Print drugNames(i) & " → エラーのため不明に分類"
        End If
    Next i

    ' 分類結果を表示
    Debug.Print "分類結果:"
    Debug.Print "  PTP: " & categoryCounts(INDEX_PTP) & "件"
    Debug.Print "  バラ: " & categoryCounts(INDEX_BULK) & "件"
    Debug.Print "  分包: " & categoryCounts(INDEX_UNIT_DOSE) & "件"
    Debug.Print "  SP: " & categoryCounts(INDEX_SP) & "件"
    Debug.Print "  包装小: " & categoryCounts(INDEX_SMALL) & "件"
    Debug.Print "  その他: " & categoryCounts(INDEX_OTHER) & "件"
    Debug.Print "  不明: " & categoryCounts(INDEX_UNKNOWN) & "件"
    
    ' C列に結果を転記
    Application.StatusBar = "処理結果をC列に転記しています..."
    
    ' C列をクリア（7行目以降）
    settingsSheet.Range("C7:C" & lastRow).ClearContents
    
    ' 転記開始行
    currentRow = 7
    
    ' 包装形態ごとに順番に転記
    For i = 0 To CATEGORY_COUNT - 1
        If categoryCounts(i) > 0 Then
            For j = 1 To packageCategories(i).Count
                ' C列に転記
                settingsSheet.Cells(currentRow, "C").Value = packageCategories(i)(j)
                currentRow = currentRow + 1
                
                ' 進捗状況を更新（10件ごと）
                If j Mod 10 = 0 Then
                    Application.StatusBar = "結果を転記中... " & (currentRow - 7) & "/" & codeCount
                    DoEvents
                End If
                
                ' 最大行数を超えないようにする
                If currentRow > lastRow Then
                    Exit For
                End If
            Next j
        End If
        
        ' 最大行数を超えたらループを抜ける
        If currentRow > lastRow Then
            Exit For
        End If
    Next i
    
    ' 処理時間を計算
    Dim endTime As Double
    endTime = Timer
    Dim processingTime As Double
    processingTime = endTime - startTime
    
    ' 処理結果を表示
    Dim resultMsg As String
    resultMsg = "処理が完了しました" & vbCrLf & _
               "処理時間: " & Format(processingTime, "0.00") & "秒" & vbCrLf & _
               "処理件数: " & codeCount & "件" & vbCrLf & vbCrLf & _
               "--- 包装形態別件数 ---" & vbCrLf & _
               "PTP: " & categoryCounts(INDEX_PTP) & "件" & vbCrLf & _
               "バラ: " & categoryCounts(INDEX_BULK) & "件" & vbCrLf & _
               "分包: " & categoryCounts(INDEX_UNIT_DOSE) & "件" & vbCrLf & _
               "SP: " & categoryCounts(INDEX_SP) & "件" & vbCrLf & _
               "包装小: " & categoryCounts(INDEX_SMALL) & "件" & vbCrLf & _
               "その他: " & categoryCounts(INDEX_OTHER) & "件" & vbCrLf & _
               "不明: " & categoryCounts(INDEX_UNKNOWN) & "件"
    
    MsgBox resultMsg, vbInformation, "処理完了"
    Debug.Print "処理完了: " & Format(processingTime, "0.00") & "秒"
    Debug.Print "===== PackageBasedProcessing 終了 ====="
    
CleanupAndExit:
    ' アプリケーションの設定を元に戻す
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description & " (" & Err.Number & ")", vbCritical
    Debug.Print "エラー: " & Err.Number & " - " & Err.Description & " - " & Erl
    Resume CleanupAndExit
End Sub

' 医薬品コードから医薬品名を取得する関数
Private Sub GetDrugNamesFromCodes(ByRef drugNames() As String, ByRef drugCodes() As String, ByRef rowIndices() As Long)
    On Error GoTo ErrorHandler
    
    ' ワークシートの取得
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' 最終行の取得
    Dim lastRow As Long
    lastRow = settingsSheet.Cells(settingsSheet.Rows.Count, "A").End(xlUp).Row
    
    ' 医薬品コードの数を計算
    Dim codeCount As Long
    codeCount = lastRow - 6  ' 7行目から開始
    
    If codeCount <= 0 Then
        ' 医薬品コードが存在しない場合は空の配列を返す
        ReDim drugNames(0 To 0)
        ReDim drugCodes(0 To 0)
        ReDim rowIndices(0 To 0)
        Exit Sub
    End If
    
    ' 配列を初期化
    ReDim drugNames(1 To codeCount)
    ReDim drugCodes(1 To codeCount)
    ReDim rowIndices(1 To codeCount)
    
    ' 各行から医薬品コードを取得して医薬品名を検索
    Dim i As Long
    Dim currentItem As Long
    currentItem = 0
    
    For i = 7 To lastRow
        Dim drugCode As String
        drugCode = settingsSheet.Cells(i, "A").Value
        
        If Len(drugCode) > 0 Then
            ' カウンターをインクリメント
            currentItem = currentItem + 1
            
            ' 医薬品コードを整形
            drugCode = MainModule_Mac.FormatDrugCode(drugCode)
            
            ' 配列に格納
            drugCodes(currentItem) = drugCode
            rowIndices(currentItem) = i
            
            ' 医薬品名を取得
            drugNames(currentItem) = MainModule_Mac.FindDrugNameByCode(drugCode)
            
            ' 進捗状況を更新（10件ごと）
            If currentItem Mod 10 = 0 Then
                Application.StatusBar = "医薬品名を取得中... " & currentItem & "/" & codeCount
                DoEvents
            End If
        End If
    Next i
    
    ' 実際に取得した件数で配列のサイズを調整
    If currentItem < codeCount Then
        ReDim Preserve drugNames(1 To currentItem)
        ReDim Preserve drugCodes(1 To currentItem)
        ReDim Preserve rowIndices(1 To currentItem)
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "GetDrugNamesFromCodes エラー: " & Err.Number & " - " & Err.Description
    ' エラー時は空の配列を返す
    ReDim drugNames(0 To 0)
    ReDim drugCodes(0 To 0)
    ReDim rowIndices(0 To 0)
End Sub

' B4セルに包装形態の選択肢をドロップダウンリストとして設定する関数
Public Sub SetupPackageTypeDropdown()
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' B4セルにドロップダウンリストを設定
    With settingsSheet.Range("B4").Validation
        .Delete ' 既存の入力規則を削除
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="バラ包装,分包品"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "包装形態の選択"
        .ErrorTitle = "無効な選択"
        .InputMessage = "「バラ包装」または「分包品」を選択してください"
        .ErrorMessage = "リストから有効な包装形態を選択してください"
    End With
    
    ' B4セルの書式設定
    With settingsSheet.Range("B4")
        .Value = "バラ包装" ' デフォルト値を設定
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242) ' 薄い青色の背景
    End With
    
    ' A4セルにラベルを設定
    With settingsSheet.Range("A4")
        .Value = "包装形態:"
        .Font.Bold = True
    End With
    
    ' 現在の包装単位表示用のセルを追加
    With settingsSheet.Range("D4")
        .Value = "現在の処理:"
        .Font.Bold = True
    End With
    
    ' 現在処理中の包装単位を表示するセル
    With settingsSheet.Range("E4")
        .Value = ""
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
    End With
    
    MsgBox "包装形態のドロップダウンリストを設定しました。", vbInformation
End Sub

' ============================================================
' インポート/エクスポート関数
' ============================================================

' コードをCSVからインポートする機能
Public Sub ImportCodesFromCSV()
    ' CSVファイルの選択
    Dim filePath As String
    filePath = MainModule_Mac.GetFilePathFromDialog("CSV Files (*.csv),*.csv", "インポートするCSVファイルを選択")
    
    If Len(filePath) = 0 Then
        MsgBox "ファイルが選択されていません。", vbExclamation
        Exit Sub
    End If
    
    ' ファイルを開く
    Dim fileNum As Integer
    fileNum = FreeFile
    
    On Error GoTo ErrorHandler
    
    Open filePath For Input As #fileNum
    
    ' ファイルからデータを読み込み
    Dim line As String
    Dim codes() As String
    Dim codeCount As Long
    codeCount = 0
    
    ' サイズ不明のため、一時配列を拡張しながら読み込む
    ReDim codes(0 To 99) ' 初期サイズ
    
    ' 各行を読み込み
    Do Until EOF(fileNum)
        Line Input #fileNum, line
        
        ' 空行や不正なデータをスキップ
        If Len(Trim(line)) > 0 Then
            ' カンマ区切りの場合は最初の要素を取得
            Dim parts() As String
            If InStr(line, ",") > 0 Then
                parts = Split(line, ",")
                line = parts(0) ' 最初の要素（コード）のみ使用
            End If
            
            ' 配列にコードを追加
            If codeCount >= UBound(codes) Then
                ' 配列サイズを拡張
                ReDim Preserve codes(0 To UBound(codes) + 100)
            End If
            
            codes(codeCount) = Trim(line)
            codeCount = codeCount + 1
        End If
    Loop
    
    Close #fileNum
    
    ' 配列サイズを実際のデータ数に調整
    If codeCount > 0 Then
        ReDim Preserve codes(0 To codeCount - 1)
        
        ' 設定シートにコードを書き込み
        Dim settingsSheet As Worksheet
        Set settingsSheet = ThisWorkbook.Worksheets(1)
        
        ' A7からコードを書き込み
        Dim i As Long
        For i = 0 To codeCount - 1
            settingsSheet.Cells(i + 7, "A").Value = codes(i)
        Next i
        
        ' 整形してから医薬品名も取得
        Application.StatusBar = "医薬品コードを整形し、医薬品名を取得しています..."
        
        ' すべてのコードを一度に整形
        Dim formattedCodes() As String
        ReDim formattedCodes(0 To codeCount - 1)
        
        For i = 0 To codeCount - 1
            Dim drugCode As String
            drugCode = codes(i)
            formattedCodes(i) = MainModule_Mac.FormatDrugCode(drugCode)
            settingsSheet.Cells(i + 7, "A").Value = formattedCodes(i)
        Next i
        
        ' 医薬品名を取得して表示
        settingsSheet.Range("B7:B" & (codeCount + 6)).ClearContents
        settingsSheet.Range("C7:C" & (codeCount + 6)).ClearContents
        
        MainModule_Mac.FillDrugNamesByCode
        
        MsgBox codeCount & "件の医薬品コードをインポートしました。", vbInformation
    Else
        MsgBox "インポートするデータが見つかりませんでした。", vbExclamation
    End If
    
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Application.StatusBar = False
    MsgBox "インポート中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' メモリを解放するためのガベージコレクション関数
Private Sub CollectGarbage()
    Dim tmp As String
    tmp = Space(50000000)  ' 大きな文字列を作成
    tmp = ""  ' 解放
End Sub

' テスト用の薬品マスターシートを作成する関数（より確実なバージョン）
Public Sub CreateTestDrugMaster()
    On Error GoTo ErrorHandler
    
    Debug.Print "==== CreateTestDrugMaster 開始 ===="
    
    ' 既存の薬品マスターシートを確認
    Dim ws As Worksheet
    Dim masterExists As Boolean
    masterExists = False
    
    Debug.Print "既存シートの確認中..."
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print " - シート名: " & ws.Name
        If ws.Name = MASTER_SHEET_NAME Then
            masterExists = True
            Debug.Print " -> 薬品マスターシートが見つかりました。既存シートを使用します。"
            Exit For
        End If
    Next ws
    
    ' 薬品マスターシートがない場合は新規作成
    If Not masterExists Then
        Debug.Print "薬品マスターシートが見つかりません。新規作成します。"
        
        On Error Resume Next
        Dim newSheet As Worksheet
        Set newSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        
        If Err.Number <> 0 Then
            Debug.Print "シート作成エラー: " & Err.Description
            On Error GoTo ErrorHandler
            ' エラーリセット
            Err.Clear
            ' 代替手段: 既存シートを再利用
            Set newSheet = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
            Debug.Print "既存の最終シート " & newSheet.Name & " を再利用します"
        End If
        
        On Error GoTo ErrorHandler
        
        ' シート名を設定
        newSheet.Name = MASTER_SHEET_NAME
        Debug.Print "シート名を " & MASTER_SHEET_NAME & " に設定しました"
    End If
    
    ' 薬品マスターシートの取得
    Dim masterSheet As Worksheet
    On Error Resume Next
    Set masterSheet = ThisWorkbook.Worksheets(MASTER_SHEET_NAME)
    If Err.Number <> 0 Then
        Debug.Print "マスターシート取得エラー: " & Err.Description
        MsgBox "マスターシートの取得に失敗しました。別の方法で再試行します。", vbExclamation
        Err.Clear
        
        ' 再度全シートを検索
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name = MASTER_SHEET_NAME Then
                Set masterSheet = ws
                Debug.Print "再検索で薬品マスターシートを見つけました"
                Exit For
            End If
        Next ws
        
        If masterSheet Is Nothing Then
            Debug.Print "致命的エラー: マスターシートを取得できません"
            MsgBox "マスターシートの取得に失敗しました。処理を中止します。", vbCritical
            Exit Sub
        End If
    End If
    On Error GoTo ErrorHandler
    
    Debug.Print "薬品マスターシート " & masterSheet.Name & " を取得しました"
    
    ' 既存データのクリア
    masterSheet.Cells.Clear
    Debug.Print "薬品マスターシートの内容をクリアしました"
    
    ' ヘッダー設定
    masterSheet.Cells(1, DRUG_CODE_COLUMN).Value = "医薬品コード"
    masterSheet.Cells(1, DRUG_NAME_COLUMN).Value = "医薬品名"
    
    ' テストデータの設定
    Debug.Print "テストデータを設定します"
    
    ' コードを14桁形式に確実に整形
    Dim code1 As String, code2 As String
    code1 = FormatDrugCode("04987279551029")
    code2 = FormatDrugCode("04987279543017")
    
    ' 1行目: エムプリシティ
    masterSheet.Cells(FIRST_DATA_ROW, DRUG_CODE_COLUMN).Value = code1
    masterSheet.Cells(FIRST_DATA_ROW, DRUG_NAME_COLUMN).Value = "エムプリシティ点滴静注用４００ｍｇ 注射剤 1瓶"
    Debug.Print " - 行" & FIRST_DATA_ROW & " に設定: " & code1 & " -> エムプリシティ点滴静注用４００ｍｇ 注射剤 1瓶"
    
    ' 2行目: エリキュース
    masterSheet.Cells(FIRST_DATA_ROW + 1, DRUG_CODE_COLUMN).Value = code2
    masterSheet.Cells(FIRST_DATA_ROW + 1, DRUG_NAME_COLUMN).Value = "エリキュース錠２．５ｍｇ ＰＴＰ 10錠"
    Debug.Print " - 行" & (FIRST_DATA_ROW + 1) & " に設定: " & code2 & " -> エリキュース錠２．５ｍｇ ＰＴＰ 10錠"
    
    ' 列幅の自動調整
    masterSheet.Columns("A:B").AutoFit
    
    ' 書式を文字列に設定して先頭の0が消えないようにする
    masterSheet.Range(masterSheet.Cells(FIRST_DATA_ROW, DRUG_CODE_COLUMN), _
                      masterSheet.Cells(FIRST_DATA_ROW + 1, DRUG_CODE_COLUMN)).NumberFormat = "@"
    
    ' マスターシートが正しく設定されたか確認
    Dim testCode As String
    testCode = masterSheet.Cells(FIRST_DATA_ROW, DRUG_CODE_COLUMN).Value
    Dim testName As String
    testName = masterSheet.Cells(FIRST_DATA_ROW, DRUG_NAME_COLUMN).Value
    
    Debug.Print "設定確認 - 薬品コード: " & testCode & ", 薬品名: " & testName
    
    Debug.Print "テスト用薬品マスター作成完了 - データ行数: 2"
    MsgBox "テスト用の薬品マスターを作成しました。2件のデータを登録しました。", vbInformation
    
    Debug.Print "==== CreateTestDrugMaster 終了 ===="
    Exit Sub
    
ErrorHandler:
    Debug.Print "エラー発生: " & Err.Number & " - " & Err.Description
    MsgBox "テストデータの作成中にエラーが発生しました: " & Err.Description, vbCritical
End Sub 