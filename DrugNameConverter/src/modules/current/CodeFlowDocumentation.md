# 薬局在庫管理システム - コードフロー機能別ドキュメント

## 1. 医薬品名処理システム (Drug Name Processing System)
```vba
' DrugNameParser.bas - パース処理のメイン機能
ParseDrugString(drugStr As String) As DrugNameParts
CompareDrugStringsWithRate(sourceStr, targetStr) As Double
ExtractBaseNameSimple, ExtractFormTypeSimple, ExtractStrengthSimple, ExtractPackageTypeSimple
```
- 医薬品名をコンポーネント（基本名、剤形、規格、メーカー、包装形態）に分解
- 異なる医薬品名の一致度を計算（80%以上で一致と判定）
- StringUtilsの関数を使用して引用符間のテキスト抽出や強度比較を実行

## 2. 医薬品比較システム (Drug Comparison System)
```vba
' MainModule.bas - 比較処理の中核機能
MainProcess() - メイン比較機能
ProcessFromRow7() - 設定シートの7行目以降を処理
CalculateMatchScore() - 成分、規格、単位の一致度を計算
FindBestMatchingDrug() - キーワードベースで最適一致を検索
```
- 設定シート（Sheet1）の医薬品名とターゲットシート（Sheet2）を比較
- DrugNameParserの機能を使用して医薬品の各要素を比較
- 一致率が80%以上の場合、結果を転記

## 3. 棚管理システム (Shelf Management System)
```vba
' ShelfManager.bas - 棚管理の主要機能
Main() - 棚番一括更新処理のエントリーポイント
ProcessItems() - GTINコードを処理し棚番を更新
GetDrugName(gtin) - GTINコードから医薬品名を取得
ExportTemplateCSV() - 更新後のtmp_tanaシートをCSVに出力
```
- GTIN-14コードが記載されたCSVファイルを処理
- 現在、外部ファイル「医薬品コード.xlsx」を主な参照源として使用
- Sheet3を二次的な参照源として使用（外部ファイルで見つからない場合）
- 医薬品名を基にtmp_tanaシートの棚番情報を更新

## 4. GTIN-14コード処理システム (GTIN-14 Code Processing System)
```vba
' GS1CodeProcessor.bas - GTIN処理の核心機能
GetDrugInfoFromGS1Code(gtin14Code As String) As DrugInfo
ProcessGS1CodeAndUpdateSettings(gtin14Code As String)
```
- 14桁のGTINコードから医薬品情報を取得
- 外部ワークブック「医薬品コード.xlsx」を開いて情報を検索
- 医薬品名、成分、剤形、規格などの詳細情報を抽出

## 5. CSV取込システム (CSV Import System)
```vba
' ImportCSVToSheet2.bas - CSV取込の主な機能
ImportCSVToSheet2() - CSVファイルを選択し、データをシート2に転記
GetCSVFilePath() - CSVファイル選択ダイアログを表示
```
- 棚番テンプレートCSVファイルをシート2（ターゲット）に転記
- ファイル名が「tmp_tana.CSV」でない場合、確認ダイアログを表示
- A〜I列のデータを転記対象とする

## 6. ユーザーインターフェースシステム (User Interface System)
```vba
' DrugNameConverter.bas - UI関連の主要機能
RunDrugNameComparison() - 医薬品名比較のエントリーポイント
InitWorkbook() - ワークブックの書式設定と初期化
' DynamicShelfNameForm.frm - 動的フォームの主要機能
SetFileCount(count, fileNames) - ファイル数に応じてフォームを動的に生成
```
- ユーザーインターフェースの初期化と設定
- 動的棚名入力フォームを提供（ファイル数に応じてサイズ変更）
- マウスホイールによるスクロール機能を実装

## 7. 医薬品コード参照の現在の実装
```vba
' GS1CodeProcessor.bas:GetDrugInfoFromGS1Code - 現在の実装
Private Function GetDrugInfoFromGS1Code(ByVal gtin14Code As String) As DrugInfo
    ' 医薬品コード.xlsxを開く
    Dim drugCodeWorkbook As Workbook
    Set drugCodeWorkbook = Workbooks.Open(Application.ThisWorkbook.Path & Application.PathSeparator & "医薬品コード.xlsx")
    ' 外部ファイルからの検索処理
End Function

' ShelfManager.bas:GetDrugName - 現在の実装（フォールバック処理あり）
Private Function GetDrugName(ByVal gtin As String) As String
    ' GS1CodeProcessorを使用して医薬品名を取得（外部ファイル）
    drugInfo = GS1CodeProcessor.GetDrugInfoFromGS1Code(gtin)
    ' 見つからない場合はSheet3で検索（フォールバック）
    If drugName = "" Then
        ' シート3の検索処理
        For i = 2 To lastRow3
            If ws3.Cells(i, "F").Value = gtin Then
                drugName = ws3.Cells(i, "G").Value
                Exit For
            End If
        Next i
    End If
End Function
```

## 8. システム全体のコードフロー概要

### メインフロー
- ユーザーがマクロを実行 (`InitializeApplication` または `ShowMainMenu`) 
- 機能選択メニューが表示される
- 選択に応じて各機能のエントリーポイントが呼び出される

### 医薬品名比較フロー
- `RunDrugNameComparison` → `MainProcess` または `ProcessFromRow7`
- 医薬品名の解析 (`ParseDrugString`)
- 一致率の計算 (`CompareDrugStringsWithRate`)
- 結果の転記

### 棚番一括更新フロー
- `ShelfManager.Main`
- CSVファイル選択とGTINコード読み込み
- 動的棚名入力フォーム表示
- 医薬品名の検索 (`GetDrugName` → `GetDrugInfoFromGS1Code`)
- tmp_tanaシートの更新
- 結果のエクスポート

### CSVインポートフロー
- `ImportCSVToSheet2.ImportCSVToSheet2`
- CSVファイル選択 (`GetCSVFilePath`)
- ファイル名確認（tmp_tana.CSVでない場合）
- ターゲットシート（シート2）へのデータ転記
