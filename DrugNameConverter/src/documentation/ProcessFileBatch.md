# ProcessFileBatch.bas 詳細設計書

## 概要
ProcessFileBatch.basは複数のCSVファイルを効率的に処理するためのバッチ処理機能を提供するモジュールです。大量のファイルを処理する際のパフォーマンス最適化を実現します。

## ユーザー定義型

### FileShelfData
```vba
Type FileShelfData
    FileName As String      ' CSVファイル名
    ShelfNames(1 To 3) As String  ' 棚名（最大3つ）
End Type
```
**説明**: CSVファイルと対応する棚名情報を格納するための構造体

## 主要機能

### ProcessFileBatch
```vba
Public Sub ProcessFileBatch(ByVal folderPath As String, ByVal fileShelfData() As FileShelfData)
```
**説明**: 複数のCSVファイルをバッチ処理します。
**引数**: 
- `folderPath` (String): CSVファイルが格納されているフォルダパス
- `fileShelfData()` (FileShelfData): ファイルと棚名の情報を含む配列
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ProcessCSVFiles` (ShelfManager.bas)
- 呼び出し先: `ProcessSingleCSVFileWithArray`, `CollectGarbage`

### OptimizePerformance
```vba
Public Sub OptimizePerformance(ByVal enable As Boolean)
```
**説明**: パフォーマンス最適化のための設定を有効/無効にします。
**引数**: 
- `enable` (Boolean): 最適化を有効にする場合はTrue
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ProcessFileBatch`
- 呼び出し先: なし

### ShowProgressBar
```vba
Public Sub ShowProgressBar(ByVal currentFile As Integer, ByVal totalFiles As Integer)
```
**説明**: 処理の進捗状況を表示します。
**引数**: 
- `currentFile` (Integer): 現在処理中のファイル番号
- `totalFiles` (Integer): 処理する総ファイル数
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ProcessFileBatch`
- 呼び出し先: なし

### LogProcessingResult
```vba
Public Sub LogProcessingResult(ByVal fileName As String, ByVal success As Boolean, Optional ByVal errorMsg As String = "")
```
**説明**: ファイル処理結果をログに記録します。
**引数**: 
- `fileName` (String): 処理したファイル名
- `success` (Boolean): 処理が成功した場合はTrue
- `errorMsg` (String): エラーメッセージ（オプション）
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ProcessFileBatch`
- 呼び出し先: なし

## 補助機能

### UpdateStatusBar
```vba
Private Sub UpdateStatusBar(ByVal message As String)
```
**説明**: ステータスバーにメッセージを表示します。
**引数**: 
- `message` (String): 表示するメッセージ
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ProcessFileBatch`, `ShowProgressBar`
- 呼び出し先: なし

### CalculateEstimatedTimeRemaining
```vba
Private Function CalculateEstimatedTimeRemaining(ByVal currentFile As Integer, ByVal totalFiles As Integer, ByVal startTime As Double) As String
```
**説明**: 残り処理時間を計算します。
**引数**: 
- `currentFile` (Integer): 現在処理中のファイル番号
- `totalFiles` (Integer): 処理する総ファイル数
- `startTime` (Double): 処理開始時間
**戻り値**: String (残り時間の文字列表現)
**呼び出し関係**:
- 呼び出し元: `ShowProgressBar`
- 呼び出し先: なし

## アルゴリズム詳細

### バッチ処理アルゴリズム
1. パフォーマンス最適化設定を有効化（画面更新と自動計算を無効化）
2. 処理開始時間を記録
3. 各ファイルに対して以下の処理を実行:
   - 進捗状況を表示
   - ファイルと対応する棚名情報を取得
   - ProcessSingleCSVFileWithArray関数を呼び出してファイルを処理
   - 処理結果をログに記録
   - 定期的にメモリを解放
4. 全ファイルの処理が完了したら、パフォーマンス最適化設定を無効化
5. 処理結果のサマリーを表示

### パフォーマンス最適化アルゴリズム
1. 画面更新を無効化（Application.ScreenUpdating = False）
2. 自動計算を無効化（Application.Calculation = xlCalculationManual）
3. イベントを無効化（Application.EnableEvents = False）
4. 処理完了後、すべての設定を元に戻す

## エラーハンドリング
各関数にはエラーハンドリングが実装されており、ファイル処理中にエラーが発生した場合でも処理が継続されるよう設計されています。エラーが発生した場合は、エラー情報をログに記録し、次のファイルの処理に進みます。

## 依存関係
- ProcessSingleCSVFileWithArray.bas: 単一CSVファイル処理に使用
- CollectGarbage.bas: メモリ解放に使用
