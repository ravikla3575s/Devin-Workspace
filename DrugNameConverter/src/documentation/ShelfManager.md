# ShelfManager.bas 詳細設計書

## 概要
ShelfManager.basは棚情報の管理と更新を行うモジュールです。CSVファイルの処理、棚情報の更新、およびデータのエクスポート機能を提供します。

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

### UpdateShelfNumbersWithShelfInfo
```vba
Public Sub UpdateShelfNumbersWithShelfInfo()
```
**説明**: 設定シートの情報を基に棚情報を更新します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `MainProcess` (MainModule.bas)
- 呼び出し先: なし

### ExportToCSV
```vba
Public Sub ExportToCSV(ByVal ws As Worksheet, ByVal filePath As String)
```
**説明**: 指定されたワークシートの内容をCSVファイルにエクスポートします。
**引数**: 
- `ws` (Worksheet): エクスポートするワークシート
- `filePath` (String): エクスポート先のファイルパス
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ExportTmpTanaToCSV`
- 呼び出し先: なし

### ExportTmpTanaToCSV
```vba
Public Sub ExportTmpTanaToCSV()
```
**説明**: tmp_tanaシートの内容をCSVファイルにエクスポートします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `ExportToCSV`

### ImportCSVFiles
```vba
Public Sub ImportCSVFiles()
```
**説明**: 複数のCSVファイルを選択してインポートします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `ShowDynamicShelfNameForm`

### ShowDynamicShelfNameForm
```vba
Public Sub ShowDynamicShelfNameForm(ByVal filePaths As Variant)
```
**説明**: 動的棚名入力フォームを表示します。
**引数**: 
- `filePaths` (Variant): 処理するCSVファイルのパスの配列
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ImportCSVFiles`
- 呼び出し先: なし

### CollectAllFileShelfData
```vba
Public Function CollectAllFileShelfData(ByVal fileCount As Integer, ByRef shelfNameForm As Object) As FileShelfData()
```
**説明**: フォームから全ファイルの棚名データを収集します。
**引数**: 
- `fileCount` (Integer): 処理するファイル数
- `shelfNameForm` (Object): 棚名入力フォーム
**戻り値**: FileShelfData() (ファイルと棚名の情報を含む配列)
**呼び出し関係**:
- 呼び出し元: `ProcessCSVFiles`
- 呼び出し先: なし

### ProcessCSVFiles
```vba
Public Sub ProcessCSVFiles(ByVal folderPath As String, ByVal fileCount As Integer, ByRef shelfNameForm As Object)
```
**説明**: 複数のCSVファイルを処理します。
**引数**: 
- `folderPath` (String): CSVファイルが格納されているフォルダパス
- `fileCount` (Integer): 処理するファイル数
- `shelfNameForm` (Object): 棚名入力フォーム
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: DynamicShelfNameForm.frm
- 呼び出し先: `CollectAllFileShelfData`, `ProcessFileBatch`

### UndoShelfNames
```vba
Public Sub UndoShelfNames()
```
**説明**: 棚名の更新を元に戻します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: なし

## 補助機能

### GetFileNameWithoutExtension
```vba
Public Function GetFileNameWithoutExtension(ByVal filePath As String) As String
```
**説明**: ファイルパスから拡張子を除いたファイル名を取得します。
**引数**: 
- `filePath` (String): ファイルパス
**戻り値**: String (拡張子を除いたファイル名)
**呼び出し関係**:
- 呼び出し元: `ProcessSingleCSVFileWithArray`
- 呼び出し先: なし

### GetFolderPath
```vba
Public Function GetFolderPath(ByVal filePath As String) As String
```
**説明**: ファイルパスからフォルダパスを取得します。
**引数**: 
- `filePath` (String): ファイルパス
**戻り値**: String (フォルダパス)
**呼び出し関係**:
- 呼び出し元: `ImportCSVFiles`
- 呼び出し先: なし

## アルゴリズム詳細

### 棚情報更新アルゴリズム
1. 設定シートから薬品名と対応する棚情報を取得
2. tmp_tanaシートの既存データをクリア
3. 取得した情報をtmp_tanaシートに転記
4. 重複する棚情報を統合

### CSVファイル処理アルゴリズム
1. ユーザーに複数のCSVファイルを選択させる
2. 動的棚名入力フォームを表示
3. 各ファイルに対応する棚名を入力
4. 入力された情報を基にCSVファイルを処理
5. 処理結果を設定シートに表示

### 最適化アルゴリズム
1. 大量のファイルを処理する場合、バッチ処理を実施
2. 画面更新と自動計算を一時的に無効化してパフォーマンスを向上
3. 配列を使用して一括読み込みと書き込みを実施
4. 定期的にメモリを解放して安定性を確保

## エラーハンドリング
各関数にはエラーハンドリングが実装されており、ファイル処理中にエラーが発生した場合でも処理が継続されるよう設計されています。

## 依存関係
- GS1CodeProcessor.bas: GTIN-14コードの処理に使用
- ProcessFileBatch.bas: ファイルバッチ処理に使用
- ProcessSingleCSVFileWithArray.bas: 単一CSVファイル処理に使用
- CollectGarbage.bas: メモリ解放に使用
