# GS1CodeProcessor.bas 詳細設計書

## 概要
GS1CodeProcessor.basはGS1-128およびGTIN-14コードの処理と医薬品情報の取得を行うモジュールです。コードの検証、解析、および医薬品データベースとの連携機能を提供します。

## ユーザー定義型

### DrugInfo
```vba
Type DrugInfo
    GS1Code As String       ' GS1/GTIN-14コード
    DrugName As String      ' 医薬品名
    BaseName As String      ' 基本名称
    FormType As String      ' 剤形
    Strength As String      ' 強度
    PackageType As String   ' パッケージタイプ
    PackageSpec As String   ' パッケージ仕様
    PackageAddInfo As String ' パッケージ追加情報
    PackageIndicator As String ' パッケージインジケーター
End Type
```
**説明**: GS1/GTIN-14コードから取得した医薬品情報を格納するための構造体

## 主要機能

### ValidateGTIN14
```vba
Public Function ValidateGTIN14(ByVal inputCode As String) As String
```
**説明**: GTIN-14コードを検証し、有効な形式に変換します。
**引数**: 
- `inputCode` (String): 検証するGTIN-14コード
**戻り値**: String (検証後のGTIN-14コード、無効な場合は空文字列)
**呼び出し関係**:
- 呼び出し元: `GetDrugInfoFromGS1Code`, `ProcessItems`
- 呼び出し先: なし

### IsValidGTIN14
```vba
Public Function IsValidGTIN14(ByVal gtinCode As String) As Boolean
```
**説明**: 文字列がGTIN-14コードとして有効かどうかをチェックします。
**引数**: 
- `gtinCode` (String): チェックするコード
**戻り値**: Boolean (有効な場合はTrue、無効な場合はFalse)
**呼び出し関係**:
- 呼び出し元: `ProcessSingleCSVFileWithArray`
- 呼び出し先: なし

### GetPackageIndicator
```vba
Public Function GetPackageIndicator(ByVal gs1Code As String) As String
```
**説明**: GTIN-14コードからパッケージインジケーターを取得します。
**引数**: 
- `gs1Code` (String): GTIN-14コード
**戻り値**: String (パッケージインジケーター)
**呼び出し関係**:
- 呼び出し元: `GetDrugInfoFromGS1Code`
- 呼び出し先: なし

### GetDrugInfoFromGS1Code
```vba
Public Function GetDrugInfoFromGS1Code(ByVal gs1Code As String) As DrugInfo
```
**説明**: GS1/GTIN-14コードから医薬品情報を取得します。
**引数**: 
- `gs1Code` (String): GS1/GTIN-14コード
**戻り値**: DrugInfo (取得した医薬品情報)
**呼び出し関係**:
- 呼び出し元: `ProcessGS1Code`
- 呼び出し先: `ValidateGTIN14`, `GetPackageIndicator`, `ExtractBaseNameFromDrugName`, `ExtractFormTypeFromDrugName`, `ExtractStrengthFromDrugName`, `ExtractPackageTypeFromDrugName`, `ExtractPackageSpecFromDrugName`, `ExtractPackageAddInfoFromDrugName`

### ProcessGS1Code
```vba
Public Sub ProcessGS1Code(ByVal gs1Code As String)
```
**説明**: GS1/GTIN-14コードを処理し、結果をワークシートに表示します。
**引数**: 
- `gs1Code` (String): 処理するGS1/GTIN-14コード
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: ユーザーアクション
- 呼び出し先: `GetDrugInfoFromGS1Code`

### ProcessItems
```vba
Public Sub ProcessItems()
```
**説明**: 設定シートのA7以降に入力されたGTIN-14コードを処理し、対応する医薬品名をB列に表示します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ProcessSingleCSVFileWithArray`
- 呼び出し先: `ValidateGTIN14`, `GetDrugInfoFromGS1Code`

## 補助機能

### ExtractBaseNameFromDrugName
```vba
Public Function ExtractBaseNameFromDrugName(ByVal drugName As String) As String
```
**説明**: 医薬品名から基本名称を抽出します。
**引数**: 
- `drugName` (String): 医薬品名
**戻り値**: String (抽出された基本名称)
**呼び出し関係**:
- 呼び出し元: `GetDrugInfoFromGS1Code`
- 呼び出し先: なし

### ExtractFormTypeFromDrugName
```vba
Public Function ExtractFormTypeFromDrugName(ByVal drugName As String) As String
```
**説明**: 医薬品名から剤形を抽出します。
**引数**: 
- `drugName` (String): 医薬品名
**戻り値**: String (抽出された剤形)
**呼び出し関係**:
- 呼び出し元: `GetDrugInfoFromGS1Code`
- 呼び出し先: なし

### ExtractStrengthFromDrugName
```vba
Public Function ExtractStrengthFromDrugName(ByVal drugName As String) As String
```
**説明**: 医薬品名から強度を抽出します。
**引数**: 
- `drugName` (String): 医薬品名
**戻り値**: String (抽出された強度)
**呼び出し関係**:
- 呼び出し元: `GetDrugInfoFromGS1Code`
- 呼び出し先: なし

### ExtractPackageTypeFromDrugName
```vba
Public Function ExtractPackageTypeFromDrugName(ByVal drugName As String) As String
```
**説明**: 医薬品名からパッケージタイプを抽出します。
**引数**: 
- `drugName` (String): 医薬品名
**戻り値**: String (抽出されたパッケージタイプ)
**呼び出し関係**:
- 呼び出し元: `GetDrugInfoFromGS1Code`
- 呼び出し先: なし

### ExtractPackageSpecFromDrugName
```vba
Public Function ExtractPackageSpecFromDrugName(ByVal drugName As String) As String
```
**説明**: 医薬品名からパッケージ仕様を抽出します。
**引数**: 
- `drugName` (String): 医薬品名
**戻り値**: String (抽出されたパッケージ仕様)
**呼び出し関係**:
- 呼び出し元: `GetDrugInfoFromGS1Code`
- 呼び出し先: なし

### ExtractPackageAddInfoFromDrugName
```vba
Public Function ExtractPackageAddInfoFromDrugName(ByVal drugName As String) As String
```
**説明**: 医薬品名からパッケージ追加情報を抽出します。
**引数**: 
- `drugName` (String): 医薬品名
**戻り値**: String (抽出されたパッケージ追加情報)
**呼び出し関係**:
- 呼び出し元: `GetDrugInfoFromGS1Code`
- 呼び出し先: なし

## アルゴリズム詳細

### GTIN-14コード検証アルゴリズム
1. 入力コードから数字のみを抽出
2. 数字が含まれているかチェック
3. 数字が含まれている場合は有効と判定
4. 数字が含まれていない場合は無効と判定

### 医薬品情報取得アルゴリズム
1. GTIN-14コードを検証
2. Sheet3（医薬品コードシート）からコードに一致する医薬品を検索
3. 一致する医薬品が見つかった場合、その情報を取得
4. 医薬品名を各要素（基本名称、剤形、強度など）に分解
5. DrugInfo構造体に格納して返す

## エラーハンドリング
各関数にはエラーハンドリングが実装されており、コード処理中にエラーが発生した場合でも処理が継続されるよう設計されています。無効なコードの場合は適切なエラーメッセージを表示します。

## 依存関係
- DrugNameParser.bas: 医薬品名の解析に使用
- PackageTypeExtractor.bas: パッケージ情報の抽出に使用
