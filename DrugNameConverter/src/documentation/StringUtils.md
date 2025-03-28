# StringUtils.bas 詳細設計書

## 概要
StringUtils.basは文字列操作のユーティリティ関数を提供するモジュールです。薬品名の解析や比較に必要な文字列処理機能を集約しています。

## 主要機能

### ExtractBetweenQuotes
```vba
Public Function ExtractBetweenQuotes(ByVal str As String) As String
```
**説明**: 日本語の引用符（「」）で囲まれたテキストを抽出します。
**引数**: 
- `str` (String): 処理する文字列
**戻り値**: String (引用符内のテキスト、引用符がない場合は元の文字列)
**呼び出し関係**:
- 呼び出し元: `ParseDrugString` (DrugNameParser.bas)
- 呼び出し先: なし

### RemoveParentheses
```vba
Public Function RemoveParentheses(ByVal str As String) As String
```
**説明**: 括弧（()）とその中身を削除します。
**引数**: 
- `str` (String): 処理する文字列
**戻り値**: String (括弧を削除した文字列)
**呼び出し関係**:
- 呼び出し元: `ExtractBaseNameSimple` (DrugNameParser.bas)
- 呼び出し先: なし

### RemoveSquareBrackets
```vba
Public Function RemoveSquareBrackets(ByVal str As String) As String
```
**説明**: 角括弧（[]）とその中身を削除します。
**引数**: 
- `str` (String): 処理する文字列
**戻り値**: String (角括弧を削除した文字列)
**呼び出し関係**:
- 呼び出し元: `ExtractBaseNameSimple` (DrugNameParser.bas)
- 呼び出し先: なし

### RemoveSpecificText
```vba
Public Function RemoveSpecificText(ByVal str As String, ByVal textToRemove As String) As String
```
**説明**: 指定されたテキストを文字列から削除します。
**引数**: 
- `str` (String): 処理する文字列
- `textToRemove` (String): 削除するテキスト
**戻り値**: String (指定テキストを削除した文字列)
**呼び出し関係**:
- 呼び出し元: `ExtractBaseNameSimple` (DrugNameParser.bas)
- 呼び出し先: なし

### ExtractNumbers
```vba
Public Function ExtractNumbers(ByVal str As String) As String
```
**説明**: 文字列から数字のみを抽出します。
**引数**: 
- `str` (String): 処理する文字列
**戻り値**: String (抽出された数字)
**呼び出し関係**:
- 呼び出し元: `ExtractStrengthSimple` (DrugNameParser.bas)
- 呼び出し先: なし

### ExtractUnits
```vba
Public Function ExtractUnits(ByVal str As String) As String
```
**説明**: 文字列から単位（mg、mLなど）を抽出します。
**引数**: 
- `str` (String): 処理する文字列
**戻り値**: String (抽出された単位)
**呼び出し関係**:
- 呼び出し元: `ExtractStrengthSimple` (DrugNameParser.bas)
- 呼び出し先: なし

### CompareStrength
```vba
Public Function CompareStrength(ByVal strength1 As String, ByVal strength2 As String) As Boolean
```
**説明**: 2つの強度値（用量）を比較します。
**引数**: 
- `strength1` (String): 比較する強度1
- `strength2` (String): 比較する強度2
**戻り値**: Boolean (強度が一致する場合はTrue)
**呼び出し関係**:
- 呼び出し元: `CompareDrugStringsWithRate` (DrugNameParser.bas)
- 呼び出し先: なし

### SetupPackageTypeDropdown
```vba
Public Sub SetupPackageTypeDropdown()
```
**説明**: B4セルにパッケージタイプのドロップダウンリストを設定します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `InitWorkbook` (DrugNameConverter.bas)
- 呼び出し先: なし

## 補助機能

### TrimAll
```vba
Public Function TrimAll(ByVal str As String) As String
```
**説明**: 文字列の前後と内部の余分な空白を削除します。
**引数**: 
- `str` (String): 処理する文字列
**戻り値**: String (空白を削除した文字列)
**呼び出し関係**:
- 呼び出し元: 複数の文字列処理関数
- 呼び出し先: なし

### ContainsText
```vba
Public Function ContainsText(ByVal str As String, ByVal searchText As String) As Boolean
```
**説明**: 文字列が特定のテキストを含むかどうかをチェックします。
**引数**: 
- `str` (String): 検索対象の文字列
- `searchText` (String): 検索するテキスト
**戻り値**: Boolean (テキストが含まれる場合はTrue)
**呼び出し関係**:
- 呼び出し元: `CheckPackage` (MainModule.bas)
- 呼び出し先: なし

### GetLevenshteinDistance
```vba
Public Function GetLevenshteinDistance(ByVal s1 As String, ByVal s2 As String) As Long
```
**説明**: 2つの文字列間のレーベンシュタイン距離を計算します。
**引数**: 
- `s1` (String): 比較する文字列1
- `s2` (String): 比較する文字列2
**戻り値**: Long (レーベンシュタイン距離)
**呼び出し関係**:
- 呼び出し元: `CompareStrings` (DrugNameParser.bas)
- 呼び出し先: なし

## アルゴリズム詳細

### 文字列抽出アルゴリズム
1. 正規表現パターンに基づいて文字列から特定の部分を抽出
2. 括弧や引用符などの特殊文字を処理
3. 数字や単位などの特定の要素を抽出

### 文字列比較アルゴリズム
1. レーベンシュタイン距離を使用して文字列の類似度を計算
2. 特定のパターンや要素に基づいて文字列を比較
3. 強度値の数値部分と単位部分を分離して比較

## エラーハンドリング
各関数にはエラーハンドリングが実装されており、文字列処理中にエラーが発生した場合でも処理が継続されるよう設計されています。

## 依存関係
- なし（他のモジュールに依存せず、他のモジュールから利用される基盤モジュール）
