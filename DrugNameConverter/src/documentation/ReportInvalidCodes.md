# ReportInvalidCodes.bas 詳細設計書

## 概要
ReportInvalidCodes.basは無効なGTIN-14コードを検出し、報告するためのモジュールです。データ品質の確保と問題のあるコードの特定を支援します。

## 主要機能

### ReportInvalidCodes
```vba
Public Sub ReportInvalidCodes(ByVal folderPath As String)
```
**説明**: 指定されたフォルダ内のCSVファイルから無効なGTIN-14コードを検出し、報告します。
**引数**: 
- `folderPath` (String): 検査するCSVファイルが格納されているフォルダパス
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `ScanCSVForInvalidCodes`

### ScanCSVForInvalidCodes
```vba
Private Function ScanCSVForInvalidCodes(ByVal filePath As String) As Collection
```
**説明**: 単一のCSVファイルをスキャンし、無効なGTIN-14コードを収集します。
**引数**: 
- `filePath` (String): スキャンするCSVファイルのパス
**戻り値**: Collection (無効なコードとその行番号のコレクション)
**呼び出し関係**:
- 呼び出し元: `ReportInvalidCodes`
- 呼び出し先: `IsValidGTIN14` (GS1CodeProcessor.bas)

### GenerateInvalidCodesReport
```vba
Private Sub GenerateInvalidCodesReport(ByVal invalidCodesMap As Object)
```
**説明**: 無効なコードの報告書を生成します。
**引数**: 
- `invalidCodesMap` (Object): ファイル名と無効なコードのマッピング
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ReportInvalidCodes`
- 呼び出し先: なし

### ExportInvalidCodesToCSV
```vba
Public Sub ExportInvalidCodesToCSV(ByVal invalidCodesMap As Object, ByVal outputPath As String)
```
**説明**: 無効なコードをCSVファイルにエクスポートします。
**引数**: 
- `invalidCodesMap` (Object): ファイル名と無効なコードのマッピング
- `outputPath` (String): 出力CSVファイルのパス
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ReportInvalidCodes`
- 呼び出し先: なし

## 補助機能

### GetFileNameFromPath
```vba
Private Function GetFileNameFromPath(ByVal filePath As String) As String
```
**説明**: ファイルパスからファイル名を抽出します。
**引数**: 
- `filePath` (String): ファイルパス
**戻り値**: String (ファイル名)
**呼び出し関係**:
- 呼び出し元: `ScanCSVForInvalidCodes`
- 呼び出し先: なし

### DisplayReportSummary
```vba
Private Sub DisplayReportSummary(ByVal totalFiles As Long, ByVal filesWithInvalidCodes As Long, ByVal totalInvalidCodes As Long)
```
**説明**: 報告書の概要を表示します。
**引数**: 
- `totalFiles` (Long): 処理したファイルの総数
- `filesWithInvalidCodes` (Long): 無効なコードを含むファイルの数
- `totalInvalidCodes` (Long): 検出された無効なコードの総数
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ReportInvalidCodes`
- 呼び出し先: なし

## アルゴリズム詳細

### 無効コード検出アルゴリズム
1. 指定されたフォルダ内のすべてのCSVファイルを取得
2. 各ファイルに対して以下の処理を実行:
   - ファイルを開いて行ごとに読み込み
   - 各行のGTIN-14コードを抽出
   - コードの有効性を検証
   - 無効なコードとその行番号を記録
3. 無効なコードを含むファイルとコードの情報を収集
4. 報告書を生成
5. 必要に応じてCSVファイルにエクスポート

### 報告書生成アルゴリズム
1. 無効なコードを含むファイルの数と無効なコードの総数を計算
2. 新しいワークシートを作成
3. 報告書のヘッダーを設定
4. 各ファイルの無効なコードとその行番号を記録
5. 報告書の概要を表示

## エラーハンドリング
各関数にはエラーハンドリングが実装されており、ファイル処理中にエラーが発生した場合でも処理が継続されるよう設計されています。エラーが発生した場合は、エラー情報をログに記録し、次のファイルの処理に進みます。

## 依存関係
- GS1CodeProcessor.bas: GTIN-14コードの検証に使用
