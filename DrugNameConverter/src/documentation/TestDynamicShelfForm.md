# TestDynamicShelfForm.bas 詳細設計書

## 概要
TestDynamicShelfForm.basは動的棚名入力フォームの機能をテストするためのモジュールです。DynamicShelfNameForm.frmの機能が正しく動作することを検証します。

## 主要機能

### TestDynamicShelfForm
```vba
Public Sub TestDynamicShelfForm()
```
**説明**: 動的棚名入力フォームの基本機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `ShowDynamicShelfNameForm` (MainModule.bas)

### TestDynamicShelfFormWithMultipleFiles
```vba
Public Sub TestDynamicShelfFormWithMultipleFiles()
```
**説明**: 複数のファイルを使用して動的棚名入力フォームをテストします。
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

### TestFormScrolling
```vba
Public Sub TestFormScrolling()
```
**説明**: 動的棚名入力フォームのスクロール機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `ShowDynamicShelfNameForm` (MainModule.bas)

## 補助機能

### CreateTestFileNames
```vba
Private Function CreateTestFileNames(ByVal count As Integer) As Variant
```
**説明**: テスト用のファイル名を作成します。
**引数**: 
- `count` (Integer): 作成するファイル名の数
**戻り値**: Variant (作成したファイル名の配列)
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

### 動的フォームテストアルゴリズム
1. テスト用のファイル名を作成
2. ShowDynamicShelfNameForm関数を呼び出して動的棚名入力フォームを表示
3. フォームに表示されるファイル名が期待通りかどうかを検証
4. フォームで入力した棚名が正しく取得できるかどうかを検証

### 複数ファイルテストアルゴリズム
1. 複数のテスト用ファイル名を作成
2. ShowDynamicShelfNameForm関数を呼び出して動的棚名入力フォームを表示
3. フォームに表示されるファイル名が期待通りかどうかを検証
4. フォームで入力した棚名が正しく取得できるかどうかを検証
5. フォームのスクロール機能が正しく動作するかどうかを検証

## エラーハンドリング
各テスト関数にはエラーハンドリングが実装されており、テスト中にエラーが発生した場合でも処理が継続されるよう設計されています。エラーが発生した場合は、エラー情報をログに記録し、次のテストケースの処理に進みます。

## 依存関係
- MainModule.bas: 動的棚名入力フォームの表示と棚名取得機能をテストするために使用
- DynamicShelfNameForm.frm: 動的棚名入力フォームの機能をテストするために使用
