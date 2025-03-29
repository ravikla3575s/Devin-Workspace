# TestGTIN14Processing.bas 詳細設計書

## 概要
TestGTIN14Processing.basはGTIN-14コード処理機能のテストを行うためのモジュールです。GS1CodeProcessor.basの機能が正しく動作することを検証します。

## 主要機能

### TestGTIN14Validation
```vba
Public Sub TestGTIN14Validation()
```
**説明**: GTIN-14コードの検証機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `IsValidGTIN14` (GS1CodeProcessor.bas)

### TestGTIN14Parsing
```vba
Public Sub TestGTIN14Parsing()
```
**説明**: GTIN-14コードの解析機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `ParseGTIN14` (GS1CodeProcessor.bas)

### TestDrugNameRetrieval
```vba
Public Sub TestDrugNameRetrieval()
```
**説明**: GTIN-14コードから薬品名を取得する機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `GetDrugNameFromCode` (GS1CodeProcessor.bas)

### TestBatchProcessing
```vba
Public Sub TestBatchProcessing()
```
**説明**: 複数のGTIN-14コードを一括処理する機能をテストします。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `ProcessGTINBatch` (GS1CodeProcessor.bas)

## 補助機能

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

### AssertFalse
```vba
Private Sub AssertFalse(ByVal condition As Boolean, ByVal testName As String)
```
**説明**: 条件が偽かどうかを検証します。
**引数**: 
- `condition` (Boolean): 検証する条件
- `testName` (String): テスト名
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: 各テスト関数
- 呼び出し先: なし

### GenerateTestGTIN14
```vba
Private Function GenerateTestGTIN14() As String
```
**説明**: テスト用のGTIN-14コードを生成します。
**引数**: なし
**戻り値**: String (生成されたGTIN-14コード)
**呼び出し関係**:
- 呼び出し元: 各テスト関数
- 呼び出し先: なし

## アルゴリズム詳細

### GTIN-14検証テストアルゴリズム
1. 有効なGTIN-14コードのテストケースを準備
2. 無効なGTIN-14コードのテストケースを準備
3. 各テストケースに対して以下の検証を実行:
   - 有効なコードがIsValidGTIN14関数でTrueを返すことを確認
   - 無効なコードがIsValidGTIN14関数でFalseを返すことを確認
4. テスト結果を表示

### GTIN-14解析テストアルゴリズム
1. 解析可能なGTIN-14コードのテストケースを準備
2. 各テストケースに対して以下の検証を実行:
   - ParseGTIN14関数が正しい構成要素を返すことを確認
   - 返された構成要素の値が期待値と一致することを確認
3. テスト結果を表示

### 薬品名取得テストアルゴリズム
1. 薬品名が関連付けられたGTIN-14コードのテストケースを準備
2. 各テストケースに対して以下の検証を実行:
   - GetDrugNameFromCode関数が正しい薬品名を返すことを確認
   - 存在しないコードに対して適切なエラー処理が行われることを確認
3. テスト結果を表示

## エラーハンドリング
各テスト関数にはエラーハンドリングが実装されており、テスト中にエラーが発生した場合でも処理が継続されるよう設計されています。エラーが発生した場合は、エラー情報をログに記録し、次のテストケースの処理に進みます。

## 依存関係
- GS1CodeProcessor.bas: GTIN-14コードの処理機能をテストするために使用
