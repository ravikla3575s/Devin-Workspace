# DynamicShelfNameForm.frm 詳細設計書

## 概要
DynamicShelfNameForm.frmは動的棚名入力フォームを提供するユーザーフォームです。複数のCSVファイルに対応する棚名を一括で入力するためのインターフェースを提供します。

## フォーム設定

### 定数
```vba
Private Const ROW_HEIGHT As Integer = 20    ' 各行の高さ
Private Const LABEL_WIDTH As Integer = 300  ' ラベルの幅
Private Const TEXTBOX_WIDTH As Integer = 60 ' テキストボックスの幅
Private Const MARGIN As Integer = 10        ' マージン
Private Const MAX_HEIGHT As Integer = 250   ' フォームの最大高さ
Private Const MAX_FILES As Integer = 100    ' 処理可能な最大ファイル数
```
**説明**: フォームのレイアウトと制約を定義する定数

### 変数
```vba
Private fileLabels() As Label              ' ファイル名ラベル
Private shelfTextBoxes1() As TextBox       ' 棚名1入力用テキストボックス
Private shelfTextBoxes2() As TextBox       ' 棚名2入力用テキストボックス
Private shelfTextBoxes3() As TextBox       ' 棚名3入力用テキストボックス
Private WithEvents OKButton As CommandButton     ' OKボタン
Private WithEvents CancelButton As CommandButton ' キャンセルボタン
Private WithEvents ScrollFrame As Frame          ' スクロールフレーム
Private fileCount As Integer               ' 処理するファイル数
Private filePaths As Variant               ' ファイルパスの配列
```
**説明**: フォームのコントロールと状態を管理する変数

## 主要機能

### UserForm_Initialize
```vba
Private Sub UserForm_Initialize()
```
**説明**: フォームの初期化処理を行います。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: フォーム表示時に自動実行
- 呼び出し先: `CreateButtons`

### SetFileCount
```vba
Public Sub SetFileCount(ByVal count As Integer)
```
**説明**: 処理するファイル数を設定します。
**引数**: 
- `count` (Integer): 処理するファイル数
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ShowDynamicShelfNameForm` (ShelfManager.bas)
- 呼び出し先: なし

### SetFilePaths
```vba
Public Sub SetFilePaths(ByVal paths As Variant)
```
**説明**: 処理するファイルパスの配列を設定します。
**引数**: 
- `paths` (Variant): ファイルパスの配列
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ShowDynamicShelfNameForm` (ShelfManager.bas)
- 呼び出し先: なし

### CreateControls
```vba
Private Sub CreateControls()
```
**説明**: フォーム上のコントロール（ラベル、テキストボックスなど）を動的に作成します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `UserForm_Initialize`
- 呼び出し先: なし

### CreateButtons
```vba
Private Sub CreateButtons()
```
**説明**: OKボタンとキャンセルボタンを作成します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `UserForm_Initialize`
- 呼び出し先: なし

### OKButton_Click
```vba
Private Sub OKButton_Click()
```
**説明**: OKボタンがクリックされたときの処理を行います。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション（ボタンクリック）
- 呼び出し先: `ProcessCSVFiles` (ShelfManager.bas)

### CancelButton_Click
```vba
Private Sub CancelButton_Click()
```
**説明**: キャンセルボタンがクリックされたときの処理を行います。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション（ボタンクリック）
- 呼び出し先: なし

### GetShelfName
```vba
Public Function GetShelfName(ByVal index As Integer, ByVal shelfIndex As Integer) As String
```
**説明**: 指定されたインデックスの棚名を取得します。
**引数**: 
- `index` (Integer): ファイルのインデックス
- `shelfIndex` (Integer): 棚名のインデックス（1〜3）
**戻り値**: String (入力された棚名)
**呼び出し関係**:
- 呼び出し元: `CollectAllFileShelfData` (ShelfManager.bas)
- 呼び出し先: なし

## 補助機能

### AdjustScrollFrame
```vba
Private Sub AdjustScrollFrame()
```
**説明**: スクロールフレームのサイズと位置を調整します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `CreateControls`
- 呼び出し先: なし

### GetFileNameFromPath
```vba
Private Function GetFileNameFromPath(ByVal filePath As String) As String
```
**説明**: ファイルパスからファイル名を抽出します。
**引数**: 
- `filePath` (String): ファイルパス
**戻り値**: String (ファイル名)
**呼び出し関係**:
- 呼び出し元: `CreateControls`
- 呼び出し先: なし

## アルゴリズム詳細

### 動的コントロール生成アルゴリズム
1. フォームが初期化されると、処理するファイル数に基づいてコントロールを動的に生成
2. 各ファイルに対して、ファイル名ラベルと3つの棚名入力用テキストボックスを作成
3. コントロールの位置を計算し、スクロールフレーム内に配置
4. フォームの高さが最大高さを超える場合、スクロールバーを表示

### 棚名収集アルゴリズム
1. OKボタンがクリックされると、各テキストボックスから入力された棚名を収集
2. 収集した棚名情報をFileShelfData構造体の配列に格納
3. 収集した情報を基にCSVファイルを処理

## エラーハンドリング
各関数にはエラーハンドリングが実装されており、フォーム操作中にエラーが発生した場合でも処理が継続されるよう設計されています。

## 依存関係
- ShelfManager.bas: CSVファイル処理と棚情報管理に使用
- MouseScroll.bas: スクロール機能の実装に使用
