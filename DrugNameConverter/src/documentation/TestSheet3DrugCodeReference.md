# TestSheet3DrugCodeReference.bas 詳細設計書

## 概要
TestSheet3DrugCodeReference.basはSheet3の医薬品コード参照機能をテストするためのモジュールです。GS1CodeProcessor.basの医薬品コード参照機能が正しく動作することを検証します。

## 主要機能

### TestSheet3DrugCodeReference
```vba
Public Sub TestSheet3DrugCodeReference()
```
**説明**: Sheet3の医薬品コード参照機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `GetDrugNameFromCode` (GS1CodeProcessor.bas)

### TestDrugCodeLookup
```vba
Public Sub TestDrugCodeLookup()
```
**説明**: 医薬品コードから薬品名を検索する機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `GetDrugNameFromCode` (GS1CodeProcessor.bas)

### TestMultipleDrugCodes
```vba
Public Sub TestMultipleDrugCodes()
```
**説明**: 複数の医薬品コードを処理する機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `GetDrugNameFromCode` (GS1CodeProcessor.bas)

### TestInvalidDrugCodes
```vba
Public Sub TestInvalidDrugCodes()
```
**説明**: 無効な医薬品コードを処理する機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `GetDrugNameFromCode` (GS1CodeProcessor.bas)

## 補助機能

### SetupTestData
```vba
Private Sub SetupTestData()
```
**説明**: テスト用のデータを設定します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: 各テスト関数
- 呼び出し先: なし

### CleanupTestData
```vba
Private Sub CleanupTestData()
```
**説明**: テスト用のデータをクリーンアップします。
**引数**: なし
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

### 医薬品コード参照テストアルゴリズム
1. テスト用のデータを設定
2. GetDrugNameFromCode関数を呼び出して医薬品コードから薬品名を取得
3. 取得された薬品名が期待通りかどうかを検証
4. テスト用のデータをクリーンアップ

### 複数医薬品コードテストアルゴリズム
1. 複数のテスト用医薬品コードを準備
2. 各医薬品コードに対してGetDrugNameFromCode関数を呼び出し
3. 取得された薬品名が期待通りかどうかを検証
4. テスト用のデータをクリーンアップ

## エラーハンドリング
各テスト関数にはエラーハンドリングが実装されており、テスト中にエラーが発生した場合でも処理が継続されるよう設計されています。エラーが発生した場合は、エラー情報をログに記録し、次のテストケースの処理に進みます。

## 依存関係
- GS1CodeProcessor.bas: 医薬品コード参照機能をテストするために使用
