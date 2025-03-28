# MainModule.bas 詳細設計書

## 概要
MainModuleは薬品名比較と棚情報処理の中核機能を提供します。薬品名の比較、マッチング、および棚情報の更新を行います。

## 主要機能

### MainProcess
```vba
Public Sub MainProcess()
```
**説明**: メイン処理の入口点。設定シートの薬品名を処理し、マッチングを実行します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `RunDrugNameComparison` (DrugNameConverter.bas)
- 呼び出し先: `ProcessFromRow7`, `UpdateShelfNumbersWithShelfInfo`

### ProcessFromRow7
```vba
Public Sub ProcessFromRow7()
```
**説明**: 設定シートの7行目から薬品名を処理します。
**引数**: なし
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `MainProcess`
- 呼び出し先: `CompareAndTransferDrugNamesByPackage`

### CompareAndTransferDrugNamesByPackage
```vba
Public Sub CompareAndTransferDrugNamesByPackage(ByVal packageType As String)
```
**説明**: 指定されたパッケージタイプに基づいて薬品名を比較し転記します。
**引数**: 
- `packageType` (String): パッケージタイプ（"PTP", "バラ"など）
**戻り値**: なし
**呼び出し関係**:
- 呼び出し元: `ProcessFromRow7`
- 呼び出し先: `CalculateMatchScore`, `FindBestMatchingDrug`

### CalculateMatchScore
```vba
Public Function CalculateMatchScore(ByVal source As String, ByVal target As String) As Double
```
**説明**: 2つの薬品名の類似度スコアを計算します。
**引数**:
- `source` (String): 元の薬品名
- `target` (String): 比較対象の薬品名
**戻り値**: Double (0.0〜100.0の類似度スコア)
**呼び出し関係**:
- 呼び出し元: `CompareAndTransferDrugNamesByPackage`
- 呼び出し先: `ParseDrugString`, `CompareDrugStringsWithRate`

### FindBestMatchingDrug
```vba
Public Function FindBestMatchingDrug(ByVal drugName As String, ByVal packageType As String) As String
```
**説明**: 指定された薬品名に最も一致する薬品をターゲットシートから検索します。
**引数**:
- `drugName` (String): 検索する薬品名
- `packageType` (String): パッケージタイプ
**戻り値**: String (最も一致する薬品名)
**呼び出し関係**:
- 呼び出し元: `CompareAndTransferDrugNamesByPackage`
- 呼び出し先: `CalculateMatchScore`

### CheckPackage
```vba
Public Function CheckPackage(ByVal drugName As String, ByVal packageType As String) As Boolean
```
**説明**: 薬品名が指定されたパッケージタイプに一致するかチェックします。
**引数**:
- `drugName` (String): チェックする薬品名
- `packageType` (String): パッケージタイプ
**戻り値**: Boolean (一致する場合はTrue)
**呼び出し関係**:
- 呼び出し元: `CompareAndTransferDrugNamesByPackage`
- 呼び出し先: なし

## 定数
- `MATCH_THRESHOLD`: 薬品名マッチングの閾値（80%）
- `MAX_ROWS`: 処理する最大行数（1000）

## データフロー
1. `MainProcess` が呼び出されると、設定シートの薬品名が処理されます
2. `ProcessFromRow7` が7行目からの薬品名を処理します
3. `CompareAndTransferDrugNamesByPackage` がパッケージタイプに基づいて比較を行います
4. `CalculateMatchScore` が薬品名の類似度を計算します
5. 類似度が閾値を超えた場合、マッチングとして記録されます
6. 処理完了後、`UpdateShelfNumbersWithShelfInfo` が棚情報を更新します

## エラーハンドリング
各関数にはエラーハンドリングが実装されており、エラー発生時にはメッセージボックスでユーザーに通知します。

## 依存関係
- DrugNameParser.bas: 薬品名の解析に使用
- ShelfManager.bas: 棚情報の更新に使用
- StringUtils.bas: 文字列操作ユーティリティ
