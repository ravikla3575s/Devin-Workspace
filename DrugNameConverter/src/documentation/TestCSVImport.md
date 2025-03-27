# TestCSVImport.bas 詳細設計書

## 概要
TestCSVImport.basはCSVファイルのインポート機能をテストするためのモジュールです。ImportCSVToSheet2.basの機能が正しく動作することを検証します。

## 主要機能

### TestCSVImport
```vba
Public Sub TestCSVImport()
```
**説明**: CSVファイルのインポート機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `ImportCSVToSheet2` (ImportCSVToSheet2.bas)

### TestCSVFileNameVerification
```vba
Public Sub TestCSVFileNameVerification()
```
**説明**: CSVファイル名の検証機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `VerifyCSVFileName` (ImportCSVToSheet2.bas)

### TestCSVReading
```vba
Public Sub TestCSVReading()
```
**説明**: CSVファイルの読み込み機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `ReadCSVFile` (ImportCSVToSheet2.bas)

### TestWriteToSheet2
```vba
Public Sub TestWriteToSheet2()
```
**説明**: Sheet2への書き込み機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `WriteToSheet2` (ImportCSVToSheet2.bas)

## 補助機能

### CreateTestCSVFile
```vba
Private Function CreateTestCSVFile(ByVal filePath As String) As Boolean
```
**説明**: テスト用のCSVファイルを作成します。
**引数**: 
- `filePath` (String): 作成するCSVファイルのパス
**戻り値**: Boolean (ファイル作成が成功した場合はTrue)
**呼び出し関係**:
- 呼び出し元: 各テスト関数
- 呼び出し先: なし

### CleanupTestFile
```vba
Private Sub CleanupTestFile(ByVal filePath As String)
```
**説明**: テスト用のCSVファイルを削除します。
**引数**: 
- `filePath` (String): 削除するCSVファイルのパス
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

### CSVインポートテストアルゴリズム
1. テスト用のCSVファイルを作成
2. ImportCSVToSheet2関数を呼び出してCSVファイルをインポート
3. Sheet2の内容が期待通りかどうかを検証
4. テスト用のCSVファイルを削除

### ファイル名検証テストアルゴリズム
1. 有効なCSVファイル名のテストケースを準備
2. 無効なCSVファイル名のテストケースを準備
3. 各テストケースに対して以下の検証を実行:
   - 有効なファイル名がVerifyCSVFileName関数でTrueを返すことを確認
   - 無効なファイル名がVerifyCSVFileName関数でFalseを返すことを確認
4. テスト結果を表示

### CSV読み込みテストアルゴリズム
1. テスト用のCSVファイルを作成
2. ReadCSVFile関数を呼び出してCSVファイルを読み込み
3. 読み込まれたデータが期待通りかどうかを検証
4. テスト用のCSVファイルを削除

## エラーハンドリング
各テスト関数にはエラーハンドリングが実装されており、テスト中にエラーが発生した場合でも処理が継続されるよう設計されています。エラーが発生した場合は、エラー情報をログに記録し、次のテストケースの処理に進みます。

## 依存関係
- ImportCSVToSheet2.bas: CSVファイルのインポート機能をテストするために使用
