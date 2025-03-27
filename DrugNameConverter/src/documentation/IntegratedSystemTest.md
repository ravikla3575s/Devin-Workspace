# IntegratedSystemTest.bas 詳細設計書

## 概要
IntegratedSystemTest.basは薬局在庫管理システムの統合テストを行うためのモジュールです。システム全体の機能が正しく連携して動作することを検証します。

## 主要機能

### TestFullSystemIntegration
```vba
Public Sub TestFullSystemIntegration()
```
**説明**: システム全体の統合テストを実行します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `InitWorkbook` (DrugNameConverter.bas), `ImportCSVToSheet2` (ImportCSVToSheet2.bas), `ShowDynamicShelfNameForm` (MainModule.bas), `ProcessFileBatch` (ProcessFileBatch.bas)

### TestDrugNameConversionFlow
```vba
Public Sub TestDrugNameConversionFlow()
```
**説明**: 薬品名変換フローのテストを実行します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `ParseDrugString` (DrugNameParser.bas), `CompareDrugStringsWithRate` (DrugNameParser.bas)

### TestShelfManagementFlow
```vba
Public Sub TestShelfManagementFlow()
```
**説明**: 棚管理フローのテストを実行します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `UpdateShelfNumbersWithShelfInfo` (ShelfManager.bas), `ExportToCSV` (ShelfManager.bas)

### TestGTINProcessingFlow
```vba
Public Sub TestGTINProcessingFlow()
```
**説明**: GTIN処理フローのテストを実行します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `IsValidGTIN14` (GS1CodeProcessor.bas), `GetDrugNameFromCode` (GS1CodeProcessor.bas)

## 補助機能

### SetupTestEnvironment
```vba
Private Sub SetupTestEnvironment()
```
**説明**: テスト環境をセットアップします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: 各テスト関数
- 呼び出し先: `InitWorkbook` (DrugNameConverter.bas)

### CleanupTestEnvironment
```vba
Private Sub CleanupTestEnvironment()
```
**説明**: テスト環境をクリーンアップします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: 各テスト関数
- 呼び出し先: なし

### CreateTestCSVFiles
```vba
Private Function CreateTestCSVFiles(ByVal count As Integer) As Variant
```
**説明**: テスト用のCSVファイルを作成します。
**引数**: 
- `count` (Integer): 作成するCSVファイルの数
**戻り値**: Variant (作成したCSVファイルのパスの配列)
**呼び出し関係**:
- 呼び出し元: 各テスト関数
- 呼び出し先: なし

### AssertEqual
```vba
Private Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, ByVal testName As String)
```
**説明**: 期待値と実際の値が等しいかどうかを検証します。
**引数**: 
- `expected` (Variant): 期待値
- `actual` (Variant): 実際の値
- `testName` (String): テスト名
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: 各テスト関数
- 呼び出し先: なし

## アルゴリズム詳細

### システム統合テストアルゴリズム
1. テスト環境をセットアップ
2. テスト用のCSVファイルを作成
3. 以下の機能を順番に実行:
   - ワークブックの初期化
   - CSVファイルのインポート
   - 動的棚名入力フォームの表示
   - ファイルバッチ処理
4. 各機能の結果が期待通りかどうかを検証
5. テスト環境をクリーンアップ

### 薬品名変換フローテストアルゴリズム
1. テスト環境をセットアップ
2. テスト用の薬品名を準備
3. 薬品名の解析と比較を実行
4. 結果が期待通りかどうかを検証
5. テスト環境をクリーンアップ

## エラーハンドリング
各テスト関数にはエラーハンドリングが実装されており、テスト中にエラーが発生した場合でも処理が継続されるよう設計されています。エラーが発生した場合は、エラー情報をログに記録し、次のテストケースの処理に進みます。

## 依存関係
- DrugNameConverter.bas: ワークブックの初期化機能を使用
- ImportCSVToSheet2.bas: CSVファイルのインポート機能を使用
- MainModule.bas: 動的棚名入力フォームの表示機能を使用
- ProcessFileBatch.bas: ファイルバッチ処理機能を使用
- DrugNameParser.bas: 薬品名の解析と比較機能を使用
- ShelfManager.bas: 棚管理機能を使用
- GS1CodeProcessor.bas: GTIN処理機能を使用
