Attribute VB_Name = "TestGTIN14Processing"
Option Explicit

' GTIN-14コード処理機能のテスト用モジュール

' テスト用のGTIN-14コードを処理するテスト関数
Public Sub TestGTIN14CodeProcessing()
    On Error GoTo ErrorHandler
    
    ' テスト開始メッセージ
    Debug.Print "GTIN-14コード処理機能のテストを開始します..."
    
    ' テスト用のGTIN-14コード（実際のコードに置き換えてください）
    Dim testGTIN14Code As String
    testGTIN14Code = "14912345678901" ' 14桁のテスト用コード
    
    ' GS1CodeProcessorモジュールの初期化確認
    Debug.Print "GS1CodeProcessorモジュールの初期化をテスト中..."
    
    ' 医薬品情報を配列として取得
    Dim drugInfoArray As Variant
    drugInfoArray = GS1CodeProcessor.GetDrugInfoAsArray(testGTIN14Code)
    
    ' 結果の検証
    Debug.Print "取得した医薬品情報:"
    Debug.Print "成分名: " & drugInfoArray(1)
    Debug.Print "剤形: " & drugInfoArray(2)
    Debug.Print "用量規格: " & drugInfoArray(3)
    Debug.Print "メーカー: " & drugInfoArray(4)
    Debug.Print "包装規格: " & drugInfoArray(5)
    Debug.Print "包装形態: " & drugInfoArray(6)
    Debug.Print "追加情報: " & drugInfoArray(7)
    Debug.Print "医薬品名: " & drugInfoArray(8)
    
    ' パッケージ・インジケーターの検証
    Dim pi As String
    pi = Left(testGTIN14Code, 1)
    Debug.Print "パッケージ・インジケーター: " & pi
    
    Select Case pi
        Case "0"
            Debug.Print "調剤包装単位"
        Case "1"
            Debug.Print "販売包装単位"
        Case "2"
            Debug.Print "元梱包装単位"
        Case Else
            Debug.Print "不明なパッケージ・インジケーター"
    End Select
    
    ' 包装形態の直接抽出テスト
    Debug.Print "包装形態の直接抽出をテスト中..."
    Dim drugName As String
    drugName = drugInfoArray(8)
    
    ' PackageTypeExtractorモジュールを使用して包装形態を抽出
    Dim packageType As String
    packageType = PackageTypeExtractor.ExtractPackageTypeFromDrugName(drugName)
    Debug.Print "抽出された包装形態: " & packageType
    
    ' テスト完了メッセージ
    Debug.Print "GTIN-14コード処理機能のテストが完了しました。"
    Exit Sub
    
ErrorHandler:
    Debug.Print "テスト中にエラーが発生しました: " & Err.Description
End Sub

' 包装形態の直接抽出機能をテストする関数
Public Sub TestPackageTypeExtraction()
    On Error GoTo ErrorHandler
    
    ' テスト開始メッセージ
    Debug.Print "包装形態の直接抽出機能のテストを開始します..."
    
    ' PackageTypeExtractorモジュールの初期化
    PackageTypeExtractor.InitializePackageMappings
    
    ' テスト用の医薬品名サンプル
    Dim testDrugNames(1 To 5) As String
    testDrugNames(1) = "アムロジピンOD錠5mg「トーワ」 PTP 100錠"
    testDrugNames(2) = "ロキソプロフェンNa錠60mg「サワイ」 バラ 500錠"
    testDrugNames(3) = "アセトアミノフェン「JG」原末 分包 500mg"
    testDrugNames(4) = "ムコスタ点眼液UD2%「サワイ」 0.35mL×30本"
    testDrugNames(5) = "ベタニス錠50mg PTP10錠シート"
    
    ' 各テストケースで包装形態を抽出
    Dim i As Long
    For i = 1 To UBound(testDrugNames)
        Dim packageType As String
        packageType = PackageTypeExtractor.ExtractPackageTypeFromDrugName(testDrugNames(i))
        Debug.Print "医薬品名: " & testDrugNames(i)
        Debug.Print "抽出された包装形態: " & packageType
        Debug.Print "---"
    Next i
    
    ' テスト完了メッセージ
    Debug.Print "包装形態の直接抽出機能のテストが完了しました。"
    Exit Sub
    
ErrorHandler:
    Debug.Print "テスト中にエラーが発生しました: " & Err.Description
End Sub

' MainModuleの包装形態抽出機能をテストする関数
Public Sub TestMainModulePackageExtraction()
    On Error GoTo ErrorHandler
    
    ' テスト開始メッセージ
    Debug.Print "MainModuleの包装形態抽出機能のテストを開始します..."
    
    ' テスト用の医薬品名と比較対象リスト
    Dim searchDrug As String
    searchDrug = "アムロジピンOD錠5mg「トーワ」 PTP 100錠"
    
    Dim targetDrugs(1 To 3) As String
    targetDrugs(1) = "アムロジピンOD錠5mg「トーワ」 バラ 500錠"
    targetDrugs(2) = "アムロジピンOD錠5mg「トーワ」 PTP 100錠"
    targetDrugs(3) = "アムロジピンOD錠2.5mg「トーワ」 PTP 100錠"
    
    ' 包装形態を抽出
    Dim packageType As String
    packageType = PackageTypeExtractor.ExtractPackageTypeFromDrugName(searchDrug)
    Debug.Print "検索医薬品名: " & searchDrug
    Debug.Print "抽出された包装形態: " & packageType
    
    ' 最適な一致を検索（MainModuleのFindBestMatchingDrug関数を使用）
    ' 注意: この関数はPrivateなので、テスト用に一時的にPublicに変更するか、
    ' 同様の機能を持つテスト用関数を作成する必要があります
    
    ' テスト完了メッセージ
    Debug.Print "MainModuleの包装形態抽出機能のテストが完了しました。"
    Exit Sub
    
ErrorHandler:
    Debug.Print "テスト中にエラーが発生しました: " & Err.Description
End Sub

' 全テストを実行する関数
Public Sub RunAllTests()
    Debug.Print "===== GTIN-14処理機能の全テストを開始します ====="
    Debug.Print ""
    
    ' 各テスト関数を順番に実行
    TestPackageTypeExtraction
    Debug.Print ""
    
    TestMainModulePackageExtraction
    Debug.Print ""
    
    TestGTIN14CodeProcessing
    Debug.Print ""
    
    Debug.Print "===== 全テストが完了しました ====="
End Sub
