# DrugNameParser.bas 詳細設計書

## 概要
DrugNameParser.basは薬品名を構成要素（基本名、剤形、強度、メーカー、パッケージ）に分解し、解析するための機能を提供します。薬品名の比較や類似度計算の基盤となります。

## ユーザー定義型

### DrugNameParts
```vba
Type DrugNameParts
    BaseName As String    ' 薬品の基本名称
    formType As String    ' 剤形（錠剤、カプセル等）
    strength As String    ' 強度（用量）
    maker As String       ' メーカー名
    Package As String     ' パッケージ情報
End Type
```
**説明**: 薬品名を構成する各要素を格納するための構造体

## 主要機能

### ParseDrugString
```vba
Public Function ParseDrugString(ByVal drugStr As String) As DrugNameParts
```
**説明**: 薬品名文字列を解析し、構成要素に分解します。
**引数**: 
- `drugStr` (String): 解析する薬品名文字列
**戻り値**: DrugNameParts (分解された薬品名の構成要素)
**呼び出し関係**:
- 呼び出し元: `CalculateMatchScore` (MainModule.bas)
- 呼び出し先: `ExtractBaseNameSimple`, `ExtractStrengthSimple`, `ExtractFormTypeSimple`, `ExtractPackageTypeSimple`

### ExtractBaseNameSimple
```vba
Public Function ExtractBaseNameSimple(ByVal drugStr As String) As String
```
**説明**: 薬品名から基本名称部分を抽出します。
**引数**: 
- `drugStr` (String): 薬品名文字列
**戻り値**: String (抽出された基本名称)
**呼び出し関係**:
- 呼び出し元: `ParseDrugString`
- 呼び出し先: なし

### ExtractStrengthSimple
```vba
Public Function ExtractStrengthSimple(ByVal drugStr As String) As String
```
**説明**: 薬品名から強度（用量）情報を抽出します。
**引数**: 
- `drugStr` (String): 薬品名文字列
**戻り値**: String (抽出された強度情報)
**呼び出し関係**:
- 呼び出し元: `ParseDrugString`
- 呼び出し先: なし

### ExtractFormTypeSimple
```vba
Public Function ExtractFormTypeSimple(ByVal drugStr As String) As String
```
**説明**: 薬品名から剤形情報を抽出します。
**引数**: 
- `drugStr` (String): 薬品名文字列
**戻り値**: String (抽出された剤形情報)
**呼び出し関係**:
- 呼び出し元: `ParseDrugString`
- 呼び出し先: なし

### ExtractPackageTypeSimple
```vba
Public Function ExtractPackageTypeSimple(ByVal drugStr As String) As String
```
**説明**: 薬品名からパッケージ情報を抽出します。
**引数**: 
- `drugStr` (String): 薬品名文字列
**戻り値**: String (抽出されたパッケージ情報)
**呼び出し関係**:
- 呼び出し元: `ParseDrugString`
- 呼び出し先: なし

### CompareDrugStringsWithRate
```vba
Public Function CompareDrugStringsWithRate(ByVal sourceStr As String, ByVal targetStr As String) As Double
```
**説明**: 2つの薬品名の類似度を計算します。
**引数**: 
- `sourceStr` (String): 比較元の薬品名
- `targetStr` (String): 比較先の薬品名
**戻り値**: Double (0.0〜100.0の類似度スコア)
**呼び出し関係**:
- 呼び出し元: `CalculateMatchScore` (MainModule.bas)
- 呼び出し先: `ParseDrugString`

## 補助機能

### ExtractMakerSimple
```vba
Public Function ExtractMakerSimple(ByVal drugStr As String) As String
```
**説明**: 薬品名からメーカー情報を抽出します。
**引数**: 
- `drugStr` (String): 薬品名文字列
**戻り値**: String (抽出されたメーカー情報)
**呼び出し関係**:
- 呼び出し元: `ParseDrugString`
- 呼び出し先: なし

### CompareStrings
```vba
Private Function CompareStrings(ByVal str1 As String, ByVal str2 As String) As Double
```
**説明**: 2つの文字列の類似度を計算します。
**引数**: 
- `str1` (String): 比較元の文字列
- `str2` (String): 比較先の文字列
**戻り値**: Double (0.0〜1.0の類似度)
**呼び出し関係**:
- 呼び出し元: `CompareDrugStringsWithRate`
- 呼び出し先: なし

## アルゴリズム詳細

### 薬品名解析アルゴリズム
1. 薬品名文字列を受け取る
2. 基本名称を抽出（特定のパターンや記号を基に分離）
3. 強度情報を抽出（数字+単位のパターンを検出）
4. 剤形情報を抽出（「錠」「カプセル」などの特定キーワードを検出）
5. パッケージ情報を抽出（「PTP」「バラ」などの特定キーワードを検出）
6. メーカー情報を抽出（特定のパターンを基に分離）
7. 抽出した各要素をDrugNameParts構造体に格納して返す

### 類似度計算アルゴリズム
1. 2つの薬品名を構成要素に分解
2. 各要素（基本名称、剤形、強度など）ごとに類似度を計算
3. 要素ごとの重み付けを行い、総合的な類似度スコアを算出
4. 0.0〜100.0の範囲で類似度を返す

## エラーハンドリング
各関数にはエラーハンドリングが実装されており、解析中にエラーが発生した場合でも処理が継続されるよう設計されています。

## 依存関係
- StringUtils.bas: 文字列操作ユーティリティを使用
