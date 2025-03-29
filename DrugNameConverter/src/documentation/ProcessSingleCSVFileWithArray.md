# ProcessSingleCSVFileWithArray.bas 詳細設計書

## 概要
ProcessSingleCSVFileWithArray.basは単一のCSVファイルを配列を使用して効率的に処理するモジュールです。GTIN-14コードの処理と棚情報の更新を最適化された方法で実行します。

## 主要機能

### ProcessSingleCSVFileWithArray
```vba
Public Sub ProcessSingleCSVFileWithArray(ByVal filePath As String, ByVal shelfNames() As String)
```
**説明**: 単一のCSVファイルを配列を使用して処理します。
**引数**: 
- `filePath` (String): 処理するCSVファイルのパス
- `shelfNames()` (String): 棚名の配列（最大3つ）
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ProcessFileBatch` (ProcessFileBatch.bas)
- 呼び出し先: `ReadCSVToArray`, `ProcessGTINCodesInArray`

### ReadCSVToArray
```vba
Private Function ReadCSVToArray(ByVal filePath As String) As Variant
```
**説明**: CSVファイルを読み込み、2次元配列として返します。
**引数**: 
- `filePath` (String): 読み込むCSVファイルのパス
**戻り値**: Variant (CSVデータを含む2次元配列)
**呼び出し関係**:
- 呼び出し元: `ProcessSingleCSVFileWithArray`
- 呼び出し先: なし

### ProcessGTINCodesInArray
```vba
Private Sub ProcessGTINCodesInArray(ByVal csvData As Variant, ByVal shelfNames() As String)
```
**説明**: 配列内のGTIN-14コードを処理します。
**引数**: 
- `csvData` (Variant): 処理するCSVデータの2次元配列
- `shelfNames()` (String): 棚名の配列（最大3つ）
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ProcessSingleCSVFileWithArray`
- 呼び出し先: `IsValidGTIN14` (GS1CodeProcessor.bas), `GetDrugNameFromCode` (GS1CodeProcessor.bas)

### UpdateSettingsSheet
```vba
Private Sub UpdateSettingsSheet(ByVal gtinCodes() As String, ByVal drugNames() As String, ByVal rowCount As Long)
```
**説明**: 設定シートをGTIN-14コードと薬品名で更新します。
**引数**: 
- `gtinCodes()` (String): GTIN-14コードの配列
- `drugNames()` (String): 薬品名の配列
- `rowCount` (Long): 処理する行数
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ProcessGTINCodesInArray`
- 呼び出し先: なし

### UpdateShelfNamesInSheet
```vba
Private Sub UpdateShelfNamesInSheet(ByVal shelfNames() As String)
```
**説明**: 設定シートのB1〜B3セルに棚名を設定します。
**引数**: 
- `shelfNames()` (String): 棚名の配列（最大3つ）
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ProcessSingleCSVFileWithArray`
- 呼び出し先: なし

## 補助機能

### ValidateShelfNames
```vba
Private Function ValidateShelfNames(ByVal shelfNames() As String) As Boolean
```
**説明**: 棚名が有効かどうかを検証します。
**引数**: 
- `shelfNames()` (String): 検証する棚名の配列
**戻り値**: Boolean (有効な場合はTrue)
**呼び出し関係**:
- 呼び出し元: `ProcessSingleCSVFileWithArray`
- 呼び出し先: なし

### LogProcessingError
```vba
Private Sub LogProcessingError(ByVal errorMsg As String, ByVal filePath As String)
```
**説明**: 処理エラーをログに記録します。
**引数**: 
- `errorMsg` (String): エラーメッセージ
- `filePath` (String): エラーが発生したファイルパス
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ProcessSingleCSVFileWithArray`
- 呼び出し先: なし

## アルゴリズム詳細

### CSVファイル処理アルゴリズム
1. CSVファイルを一度に読み込み、2次元配列に格納
2. 棚名の有効性を検証
3. 設定シートのB1〜B3セルに棚名を設定
4. 配列内のGTIN-14コードを処理:
   - 各コードの有効性を検証
   - 有効なコードから薬品名を取得
   - 取得した情報を一時配列に格納
5. 一時配列の内容を一括で設定シートに転記
6. 処理結果を表示

### パフォーマンス最適化アルゴリズム
1. 配列を使用してファイルを一度に読み込み
2. 中間処理結果を一時配列に格納
3. 一時配列の内容を一括でワークシートに転記
4. 画面更新を最小限に抑制

## エラーハンドリング
各関数にはエラーハンドリングが実装されており、ファイル処理中にエラーが発生した場合でも処理が継続されるよう設計されています。エラーが発生した場合は、エラー情報をログに記録し、処理を続行します。

## 依存関係
- GS1CodeProcessor.bas: GTIN-14コードの処理と薬品名の取得に使用
