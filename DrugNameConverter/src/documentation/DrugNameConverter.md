# DrugNameConverter.bas 詳細設計書

## 概要
DrugNameConverter.basはUIとワークブック初期化を担当するモジュールです。アプリケーションの起動、ユーザーインターフェースの設定、および各種処理の呼び出し機能を提供します。

## 主要機能

### RunDrugNameComparison
```vba
Public Sub RunDrugNameComparison()
```
**説明**: 薬品名比較処理を実行します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション（ボタンクリックなど）
- 呼び出し先: `MainProcess` (MainModule.bas)

### SetupPackageTypeDropdown
```vba
Public Sub SetupPackageTypeDropdown()
```
**説明**: B4セルにパッケージタイプのドロップダウンリストを設定します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `InitWorkbook`
- 呼び出し先: なし

### AddInstructions
```vba
Public Sub AddInstructions()
```
**説明**: 設定シートに使用方法の説明を追加します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `InitWorkbook`
- 呼び出し先: なし

### InitWorkbook
```vba
Public Sub InitWorkbook()
```
**説明**: ワークブックを初期化し、必要なシートやフォーマットを設定します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: Workbook_Open イベント
- 呼び出し先: `SetupPackageTypeDropdown`, `AddInstructions`

### Workbook_Open
```vba
Private Sub Workbook_Open()
```
**説明**: ワークブックが開かれたときに実行される処理です。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: Excel自動実行
- 呼び出し先: `InitWorkbook`

### ShowStartupDialog
```vba
Public Sub ShowStartupDialog()
```
**説明**: 起動時にモード選択ダイアログを表示します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `InitWorkbook`
- 呼び出し先: なし

## 補助機能

### ClearWorksheet
```vba
Public Sub ClearWorksheet(ByVal ws As Worksheet, ByVal startRow As Long)
```
**説明**: 指定されたワークシートの内容をクリアします。
**引数**: 
- `ws` (Worksheet): クリアするワークシート
- `startRow` (Long): クリアを開始する行番号
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `InitWorkbook`
- 呼び出し先: なし

### SetupHeaderRow
```vba
Public Sub SetupHeaderRow(ByVal ws As Worksheet)
```
**説明**: 指定されたワークシートにヘッダー行を設定します。
**引数**: 
- `ws` (Worksheet): ヘッダーを設定するワークシート
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `InitWorkbook`
- 呼び出し先: なし

### FormatWorksheet
```vba
Public Sub FormatWorksheet(ByVal ws As Worksheet)
```
**説明**: 指定されたワークシートのフォーマットを設定します。
**引数**: 
- `ws` (Worksheet): フォーマットを設定するワークシート
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `InitWorkbook`
- 呼び出し先: なし

## アルゴリズム詳細

### ワークブック初期化アルゴリズム
1. ワークブックが開かれると、Workbook_Openイベントが発生
2. InitWorkbook関数が呼び出され、以下の処理を実行:
   - 必要なシートが存在しない場合は作成
   - 各シートのフォーマットを設定
   - ヘッダー行を設定
   - パッケージタイプのドロップダウンリストを設定
   - 使用方法の説明を追加
3. 起動モード選択ダイアログを表示（オプション）

### 薬品名比較処理アルゴリズム
1. ユーザーがボタンをクリックするなどのアクションを実行
2. RunDrugNameComparison関数が呼び出される
3. MainModule.basのMainProcess関数を呼び出して薬品名比較処理を実行

## エラーハンドリング
各関数にはエラーハンドリングが実装されており、初期化中にエラーが発生した場合でも処理が継続されるよう設計されています。

## 依存関係
- MainModule.bas: 薬品名比較処理の実行に使用
- ShelfManager.bas: 棚情報の管理に使用
