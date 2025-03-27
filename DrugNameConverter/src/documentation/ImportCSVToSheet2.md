# ImportCSVToSheet2.bas 詳細設計書

## 概要
ImportCSVToSheet2.basはCSVファイルをSheet2（ターゲットシート）にインポートするための機能を提供するモジュールです。棚番テンプレートCSVファイルの読み込みと処理を担当します。

## 主要機能

### ImportCSVToSheet2
```vba
Public Sub ImportCSVToSheet2()
```
**説明**: CSVファイルを選択してSheet2にインポートします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `ReadCSVFile`

### ReadCSVFile
```vba
Public Function ReadCSVFile(ByVal filePath As String) As Variant
```
**説明**: CSVファイルを読み込み、2次元配列として返します。
**引数**: 
- `filePath` (String): 読み込むCSVファイルのパス
**戻り値**: Variant (CSVデータを含む2次元配列)
**呼び出し関係**:
- 呼び出し元: `ImportCSVToSheet2`
- 呼び出し先: なし

### WriteToSheet2
```vba
Public Sub WriteToSheet2(ByVal data As Variant)
```
**説明**: 2次元配列のデータをSheet2に書き込みます。
**引数**: 
- `data` (Variant): 書き込むデータの2次元配列
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ImportCSVToSheet2`
- 呼び出し先: なし

### ClearSheet2
```vba
Public Sub ClearSheet2()
```
**説明**: Sheet2の内容をクリアします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ImportCSVToSheet2`
- 呼び出し先: なし

### VerifyCSVFileName
```vba
Public Function VerifyCSVFileName(ByVal fileName As String) As Boolean
```
**説明**: CSVファイル名が有効かどうかを検証します。
**引数**: 
- `fileName` (String): 検証するファイル名
**戻り値**: Boolean (有効な場合はTrue)
**呼び出し関係**:
- 呼び出し元: `ImportCSVToSheet2`
- 呼び出し先: なし

## 補助機能

### GetFileNameFromPath
```vba
Public Function GetFileNameFromPath(ByVal filePath As String) As String
```
**説明**: ファイルパスからファイル名を抽出します。
**引数**: 
- `filePath` (String): ファイルパス
**戻り値**: String (ファイル名)
**呼び出し関係**:
- 呼び出し元: `ImportCSVToSheet2`
- 呼び出し先: なし

### FormatSheet2
```vba
Public Sub FormatSheet2()
```
**説明**: Sheet2のフォーマットを設定します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `WriteToSheet2`
- 呼び出し先: なし

## アルゴリズム詳細

### CSVインポートアルゴリズム
1. ユーザーにCSVファイルを選択させる
2. ファイル名が有効かどうかを検証
3. Sheet2の内容をクリア
4. CSVファイルを読み込み、2次元配列に格納
5. 2次元配列のデータをSheet2に書き込み
6. Sheet2のフォーマットを設定

### CSVファイル読み込みアルゴリズム
1. CSVファイルをテキストファイルとして開く
2. 行ごとに読み込み、カンマで分割
3. 分割したデータを2次元配列に格納
4. ファイルを閉じて2次元配列を返す

## エラーハンドリング
各関数にはエラーハンドリングが実装されており、ファイル処理中にエラーが発生した場合でも処理が継続されるよう設計されています。

## 依存関係
- なし（他のモジュールに依存せず、独立して機能する）
