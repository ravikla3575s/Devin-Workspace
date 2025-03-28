Attribute VB_Name = "DrugNameConverter"
Option Explicit

' ラッパーモジュール - 基本機能

' メイン処理を呼び出すラッパー関数（7行目以降の医薬品名比較）
Public Sub RunDrugNameComparison()
    ' 処理モード選択ダイアログを表示して処理を実行
    MainModule.Main
End Sub

' 医薬品名から直接包装形態を抽出する機能に変更
Public Sub SetupDirectPackageExtraction()
    ' PackageTypeExtractorモジュールを初期化
    PackageTypeExtractor.InitializePackageMappings
    
    MsgBox "医薬品名から直接包装形態を抽出する機能を有効化しました。", vbInformation
End Sub

' シート1にインストラクションを追加する関数
Public Sub AddInstructions()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    
    ' 既存の指示を削除
    ws.Range("A2:C3").ClearContents
    
    ' 指示を追加
    ws.Range("A2").Value = "【使い方】"
    ws.Range("A3").Value = "1. B7以降に検索したい医薬品名を入力"
    
    ' セルの書式設定
    ws.Range("A2").Font.Bold = True
    ws.Range("A2").Font.Size = 12
    
    ' 実行方法の指示
    ws.Range("A5").Value = "2. 下記の実行方法で処理を開始"
    ws.Range("A6").Value = "No."
    ws.Range("B6").Value = "検索医薬品名"
    ws.Range("C6").Value = "一致医薬品名"
    
    With ws.Range("A6:C6")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247) ' 薄い青の背景
    End With
    
    ' 実行方法の説明
    Dim note As String
    note = "処理実行方法: メニューから「ツール」→「マクロ」を選択し、" & vbCrLf & _
           "「RunDrugNameComparison」を選択で「実行」ボタンをクリックします。"
    
    ws.Range("A" & (ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 2)).Value = note
    
    MsgBox "使用方法の指示を追加しました。メニューから「ツール」→「マクロ」を選択し、" & vbCrLf & _
           "「RunDrugNameComparison」を選択で処理を実行してください。", vbInformation
End Sub

' ワークブックの初期化関数
Public Sub InitWorkbook()
    On Error GoTo ErrorHandler
    
    ' ワークシートの参照を取得
    Dim settingsSheet As Worksheet
    Dim targetSheet As Worksheet
    
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    Set targetSheet = ThisWorkbook.Worksheets(2)
    
    ' シート1の設定
    With settingsSheet
        ' タイトル設定
        .Range("A1:C1").Merge
        .Range("A1").Value = "医薬品名比較ツール"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.Color = RGB(180, 198, 231)
        
        ' 使い方
        .Range("A2").Value = "【使い方】"
        .Range("A2").Font.Bold = True
        .Range("A3").Value = "1. B7以降に検索したい医薬品名を入力"
        
        ' 包装形態ドロップダウンリストは廃止
        ' 医薬品名から直接包装形態を抽出するように変更
        
        ' 初期化時にPackageTypeExtractorを初期化
        PackageTypeExtractor.InitializePackageMappings
        
        ' メニュー設定ボタンを追加
        With settingsSheet.Buttons.Add(10, 30, 120, 30)
            .OnAction = "ShowMainMenu"
            .Caption = "メニュー表示"
            .Name = "MenuButton"
        End With
        
        ' 手順
        .Range("A5").Value = "2. 下記の実行方法で処理を開始"
        .Range("A5").Font.Bold = True
        
        ' ヘッダー
        .Range("A6").Value = "No."
        .Range("B6").Value = "検索医薬品名"
        .Range("C6").Value = "一致医薬品名"
        .Range("A6:C6").Font.Bold = True
        .Range("A6:C6").Interior.Color = RGB(221, 235, 247)
        
        ' 列
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 40
        
        ' 行番号
        Dim i As Long
        For i = 7 To 30
            .Cells(i, "A").Value = i - 6
        Next i
        
        ' 実行方法の説明
        .Range("A32").Value = "処理実行方法: メニューから「ツール」→「マクロ」を選択し、「RunDrugNameComparison」を実行"
        .Range("A32").Font.Italic = True
        
        ' GS1コード処理に関する説明を追加
        .Range("A34").Value = "【GS1コード処理】"
        .Range("A34").Font.Bold = True
        .Range("A35").Value = "メニューから「ツール」→「マクロ」→「RunGS1CodeProcessing」を"
        .Range("A36").Value = "選択すると、GTIN-14の14桁コードから医薬品情報を設定シートに転記できます。"
    End With
    
    ' シート2の設定
    With targetSheet
        ' タイトル
        .Range("A1:B1").Merge
        .Range("A1").Value = "比較対象医薬品リスト"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.Color = RGB(180, 198, 231)
        
        ' ヘッダー
        .Range("A2").Value = "No."
        .Range("B2").Value = "医薬品名"
        .Range("A2:B2").Font.Bold = True
        .Range("A2:B2").Interior.Color = RGB(221, 235, 247)
        
        ' 列
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 50
        
        ' 行番号
        For i = 3 To 30
            .Cells(i, "A").Value = i - 2
        Next i
    End With
    
    MsgBox "ワークブックが初期化されました。" & vbNewLine & _
           "1. 設定シートのB7以降に検索したい医薬品名を入力" & vbNewLine & _
           "2. シート2に比較対象の医薬品名を入力" & vbNewLine & _
           "3. メニューの「ツール」→「マクロ」から「RunDrugNameComparison」を実行", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

' GS1コード処理機能を実行するラッパー関数
Public Sub RunGS1CodeProcessing()
    ' MainModuleの関数を呼び出し
    MainModule.ProcessGS1DrugCode
End Sub

' メニューにGS1コード処理機能を追加する
Public Sub AddGS1ProcessingInstructions()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    
    ' GS1処理に関する説明を追加
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 2
    
    ws.Cells(lastRow, "A").Value = "【GS1コード処理機能】"
    ws.Cells(lastRow, "A").Font.Bold = True
    ws.Cells(lastRow, "A").Font.Size = 12
    
    ws.Cells(lastRow + 1, "A").Value = "メニューから「ツール」→「マクロ」→「RunGS1CodeProcessing」を"
    ws.Cells(lastRow + 2, "A").Value = "選択すると、GTIN-14の14桁コードから医薬品情報を処理できます。"
    
    MsgBox "GS1コード処理機能の説明をシートに追加しました。", vbInformation
End Sub


