# PackageTypeExtractor.bas 詳細設計書

## 概要
PackageTypeExtractor.basは薬品名からパッケージタイプを抽出するための専用モジュールです。薬品名に含まれるパッケージ情報を正確に識別し、標準化された形式で返します。

## 主要機能

### ExtractPackageType
```vba
Public Function ExtractPackageType(ByVal drugName As String) As String
```
**説明**: 薬品名からパッケージタイプを抽出します。
**引数**: 
- `drugName` (String): パッケージタイプを抽出する薬品名
**戻り値**: String (抽出されたパッケージタイプ)
**呼び出し関係**:
- 呼び出し元: `ParseDrugString` (DrugNameParser.bas)
- 呼び出し先: `ExtractPackageTypeSimple`, `NormalizePackageType`

### ExtractPackageTypeSimple
```vba
Public Function ExtractPackageTypeSimple(ByVal drugStr As String) As String
```
**説明**: 薬品名から基本的なパッケージタイプを抽出します。
**引数**: 
- `drugStr` (String): パッケージタイプを抽出する薬品名
**戻り値**: String (抽出された基本パッケージタイプ)
**呼び出し関係**:
- 呼び出し元: `ExtractPackageType`
- 呼び出し先: `ContainsText` (StringUtils.bas)

### NormalizePackageType
```vba
Public Function NormalizePackageType(ByVal packageType As String) As String
```
**説明**: 抽出されたパッケージタイプを標準形式に正規化します。
**引数**: 
- `packageType` (String): 正規化するパッケージタイプ
**戻り値**: String (正規化されたパッケージタイプ)
**呼び出し関係**:
- 呼び出し元: `ExtractPackageType`
- 呼び出し先: なし

### IsValidPackageType
```vba
Public Function IsValidPackageType(ByVal packageType As String) As Boolean
```
**説明**: 指定されたパッケージタイプが有効かどうかを検証します。
**引数**: 
- `packageType` (String): 検証するパッケージタイプ
**戻り値**: Boolean (有効な場合はTrue)
**呼び出し関係**:
- 呼び出し元: `CheckPackage` (MainModule.bas)
- 呼び出し先: なし

## 補助機能

### GetStandardPackageTypes
```vba
Public Function GetStandardPackageTypes() As Variant
```
**説明**: 標準パッケージタイプのリストを返します。
**引数**: なし
**戻り値**: Variant (標準パッケージタイプの配列)
**呼び出し関係**:
- 呼び出し元: `SetupPackageTypeDropdown` (StringUtils.bas)
- 呼び出し先: なし

### GetPackageTypeFromSuffix
```vba
Private Function GetPackageTypeFromSuffix(ByVal drugStr As String) As String
```
**説明**: 薬品名の接尾辞からパッケージタイプを取得します。
**引数**: 
- `drugStr` (String): パッケージタイプを取得する薬品名
**戻り値**: String (接尾辞から取得したパッケージタイプ)
**呼び出し関係**:
- 呼び出し元: `ExtractPackageTypeSimple`
- 呼び出し先: なし

## アルゴリズム詳細

### パッケージタイプ抽出アルゴリズム
1. 薬品名を解析し、パッケージ情報を含む部分を特定
2. 以下の順序でパッケージタイプを検索:
   - 括弧内のパッケージ情報（例: 「（PTP）」）
   - 特定のキーワード（「PTP」、「バラ」、「SP」など）
   - 薬品名の接尾辞
3. 抽出されたパッケージタイプを標準形式に正規化
4. 正規化されたパッケージタイプを返す

### パッケージタイプ正規化アルゴリズム
1. パッケージタイプの大文字・小文字を統一
2. 同義語を標準表現に変換（例: 「シート」→「PTP」）
3. 不要な文字や空白を削除
4. 標準パッケージタイプリストと照合
5. 標準形式のパッケージタイプを返す

## エラーハンドリング
各関数にはエラーハンドリングが実装されており、パッケージタイプの抽出中にエラーが発生した場合でも処理が継続されるよう設計されています。パッケージタイプが特定できない場合は、デフォルト値または空文字列を返します。

## 依存関係
- StringUtils.bas: 文字列操作ユーティリティに使用
