# TestCSVImportAndDynamicForm.bas 詳細設計書

## 概要
TestCSVImportAndDynamicForm.basはCSVファイルのインポート機能と動的棚名入力フォームの連携をテストするためのモジュールです。ImportCSVToSheet2.basとDynamicShelfNameForm.frmの機能が正しく連携することを検証します。

## 主要機能

### TestCSVImportAndDynamicForm
```vba
Public Sub TestCSVImportAndDynamicForm()
```
**説明**: CSVファイルのインポートと動的棚名入力フォームの連携をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `ImportCSVToSheet2` (ImportCSVToSheet2.bas), `ShowDynamicShelfNameForm` (MainModule.bas)

### TestMultipleCSVImport
```vba
Public Sub TestMultipleCSVImport()
```
**説明**: 複数のCSVファイルをインポートする機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `ImportMultipleCSV` (ImportCSVToSheet2.bas)

### TestDynamicFormWithCSVData
```vba
Public Sub TestDynamicFormWithCSVData()
```
**説明**: CSVデータを使用して動的棚名入力フォームを表示する機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `ShowDynamicShelfNameForm` (MainModule.bas)

### TestShelfNameRetrieval
```vba
Public Sub TestShelfNameRetrieval()
```
**説明**: 動的棚名入力フォームから棚名を取得する機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `GetShelfNamesFromForm` (MainModule.bas)

## 補助機能

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

### CleanupTestFiles
```vba
Private Sub CleanupTestFiles(ByVal filePaths As Variant)
```
**説明**: テスト用のCSVファイルを削除します。
**引数**: 
- `filePaths` (Variant): 削除するCSVファイルのパスの配列
**戻り値**: なし
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

### AssertTrue
```vba
Private Sub AssertTrue(ByVal condition As Boolean, ByVal testName As String)
```
**説明**: 条件が真かどうかを検証します。
**引数**: 
- `condition` (Boolean): 検証する条件
- `testName` (String): テスト名
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: 各テスト関数
- 呼び出し先: なし

## アルゴリズム詳細

### CSVインポートと動的フォーム連携テストアルゴリズム
1. テスト用のCSVファイルを作成
2. ImportCSVToSheet2関数を呼び出してCSVファイルをインポート
3. ShowDynamicShelfNameForm関数を呼び出して動的棚名入力フォームを表示
4. フォームに表示されるファイル名が期待通りかどうかを検証
5. フォームで入力した棚名が正しく取得できるかどうかを検証
6. テスト用のCSVファイルを削除

### 複数CSVインポートテストアルゴリズム
1. 複数のテスト用CSVファイルを作成
2. ImportMultipleCSV関数を呼び出して複数のCSVファイルをインポート
3. インポートされたデータが期待通りかどうかを検証
4. テスト用のCSVファイルを削除

## エラーハンドリング
各テスト関数にはエラーハンドリングが実装されており、テスト中にエラーが発生した場合でも処理が継続されるよう設計されています。エラーが発生した場合は、エラー情報をログに記録し、次のテストケースの処理に進みます。

## 依存関係
- ImportCSVToSheet2.bas: CSVファイルのインポート機能をテストするために使用
- MainModule.bas: 動的棚名入力フォームの表示と棚名取得機能をテストするために使用
- DynamicShelfNameForm.frm: 動的棚名入力フォームの機能をテストするために使用
