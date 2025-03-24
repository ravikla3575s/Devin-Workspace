# 薬局在庫管理システム 詳細設計書

## 1. システム概要

薬局在庫管理システムは、医薬品名の比較・マッチング、棚情報の管理、およびGTIN-14コード処理を行うExcel VBAベースのアプリケーションです。本システムは以下の主要機能を提供します：

1. 医薬品名の解析と比較
2. 医薬品名からの包装形態の自動抽出
3. GTIN-14コードからの医薬品情報の取得
4. 棚情報の管理とCSVエクスポート

## 2. モジュール構成

システムは以下のモジュールで構成されています：

### 2.1 コアモジュール

| モジュール名 | 主な役割 | 依存モジュール |
|------------|---------|--------------|
| MainModule.bas | 医薬品名比較のメイン処理、GTIN-14処理のインターフェース | DrugNameParser.bas, StringUtils.bas, PackageTypeExtractor.bas |
| DrugNameParser.bas | 医薬品名を構造的コンポーネントに解析 | StringUtils.bas |
| DrugNameConverter.bas | UIとワークブック初期化 | MainModule.bas, PackageTypeExtractor.bas |
| ShelfManager.bas | 棚情報の管理とCSVエクスポート | - |
| StringUtils.bas | 文字列操作のユーティリティ関数 | - |

### 2.2 拡張モジュール

| モジュール名 | 主な役割 | 依存モジュール |
|------------|---------|--------------|
| PackageTypeExtractor.bas | 医薬品名から包装形態を抽出 | DrugNameParser.bas |
| GS1CodeProcessor.bas | GTIN-14コードの処理と医薬品情報の取得 | DrugNameParser.bas, PackageTypeExtractor.bas |
| TestGTIN14Processing.bas | GTIN-14処理機能のテスト | GS1CodeProcessor.bas, PackageTypeExtractor.bas |

## 3. データ構造

### 3.1 ユーザー定義型（UDT）

#### DrugNameParts（DrugNameParser.bas）
```vba
Type DrugNameParts
    BaseName As String    ' 医薬品成分名
    formType As String    ' 製剤形態（錠、カプセル、散など）
    strength As String    ' 用量規格と単位
    maker As String       ' 屋号（「〇〇」形式）
    Package As String     ' 包装情報
End Type
```

#### DrugInfo（GS1CodeProcessor.bas）
```vba
Type DrugInfo
    GS1Code As String           ' GTIN-14コード
    PackageIndicator As String  ' パッケージ・インジケーター
    DrugName As String          ' 医薬品名
    BaseName As String          ' 医薬品成分名
    FormType As String          ' 製剤形態
    Strength As String          ' 用量規格と単位
    Maker As String             ' 屋号
    PackageSpec As String       ' 包装規格
    PackageForm As String       ' 包装形態
    PackageAddInfo As String    ' 包装追加情報
End Type
```

### 3.2 定数

```vba
' MainModule.bas
Private Const MATCH_THRESHOLD As Double = 0.8 ' マッチしきい値（80%）

' PackageTypeExtractor.bas
Private packageMappings As Variant ' パッケージタイプのマッピング配列
```

## 4. 主要機能の詳細設計

### 4.1 医薬品名の解析（DrugNameParser.bas）

#### 4.1.1 ParseDrugString関数
医薬品名を構造的コンポーネントに解析します。

```vba
Public Function ParseDrugString(drugStr As String) As DrugNameParts
    ' 1. 基本名の抽出
    ' 2. 剤形タイプの抽出
    ' 3. 強度の抽出
    ' 4. メーカー名の抽出
    ' 5. パッケージ情報の抽出
End Function
```

#### 4.1.2 CompareDrugStringsWithRate関数
2つの医薬品名を比較し、類似度を計算します。

```vba
Public Function CompareDrugStringsWithRate(sourceStr As String, targetStr As String) As Double
    ' 1. 両方の医薬品名を解析
    ' 2. 各コンポーネントの類似度を計算
    ' 3. 重み付けされた総合スコアを返す
End Function
```

### 4.2 包装形態の抽出（PackageTypeExtractor.bas）

#### 4.2.1 InitializePackageMappings関数
包装形態のマッピングを初期化します。

```vba
Public Sub InitializePackageMappings()
    ' 1. マッピングコレクションを作成
    ' 2. 各マッピングペアを追加（PTP → /PTP/など）
    ' 3. 配列に変換して保持
End Sub
```

#### 4.2.2 ExtractPackageTypeFromDrugName関数
医薬品名から包装形態を抽出します。

```vba
Public Function ExtractPackageTypeFromDrugName(ByVal drugName As String) As String
    ' 1. マッピングが初期化されていない場合は初期化
    ' 2. DrugNameParserの既存の抽出方法を試す
    ' 3. スラッシュで囲まれた形式を確認
    ' 4. 各マッピング元の文字列を直接検索
    ' 5. 見つからない場合はデフォルト値を返す
End Function
```

### 4.3 GTIN-14コード処理（GS1CodeProcessor.bas）

#### 4.3.1 ValidateGTIN14関数
GTIN-14コードを検証します。

```vba
Private Function ValidateGTIN14(ByVal gtinCode As String) As String
    ' 1. 数字のみを抽出
    ' 2. 14桁であることを確認
    ' 3. 有効なコードを返す
End Function
```

#### 4.3.2 GetDrugInfoFromGS1Code関数
GTIN-14コードから医薬品情報を取得します。

```vba
Public Function GetDrugInfoFromGS1Code(ByVal gs1Code As String) As DrugInfo
    ' 1. GTIN-14コードを検証
    ' 2. 医薬品コードシートを開く
    ' 3. GS1コードに一致する医薬品を検索
    ' 4. G列（調剤包装単位名称）から医薬品名を取得
    ' 5. 医薬品名を各要素に分解
    ' 6. PackageTypeExtractorを使用して包装形態を抽出
    ' 7. 結果を返す
End Function
```

### 4.4 医薬品名比較処理（MainModule.bas）

#### 4.4.1 ProcessFromRow7関数
設定シートの7行目以降の医薬品名を処理します。

```vba
Public Sub ProcessFromRow7()
    ' 1. 設定シートとターゲットシートの参照を取得
    ' 2. 各行の医薬品名を取得
    ' 3. 医薬品名から包装形態を抽出
    ' 4. 最適な一致を検索
    ' 5. 結果を設定シートに出力
End Function
```

#### 4.4.2 FindBestMatchingDrug関数
最適な一致の医薬品を検索します。

```vba
Private Function FindBestMatchingDrug(ByVal searchDrug As String, ByVal packageType As String) As Variant
    ' 1. ターゲットシートから医薬品名を取得
    ' 2. 各医薬品名との類似度を計算
    ' 3. 最高スコアの医薬品を返す
End Function
```

## 5. データフロー

### 5.1 医薬品名比較処理のデータフロー

1. ユーザーが設定シートのB7以降に医薬品名を入力
2. `RunDrugNameComparison`関数が呼び出される
3. `ProcessFromRow7`関数が各行の医薬品名を処理
4. `ExtractPackageTypeFromDrugName`関数が医薬品名から包装形態を抽出
5. `FindBestMatchingDrug`関数が最適な一致を検索
6. `CompareDrugStringsWithRate`関数が類似度を計算
7. 結果が設定シートのC列に出力される

### 5.2 GTIN-14コード処理のデータフロー

1. ユーザーがGTIN-14コードを入力
2. `ProcessGS1DrugCode`関数が呼び出される
3. `ValidateGTIN14`関数がコードを検証
4. `GetDrugInfoFromGS1Code`関数が医薬品情報を取得
5. 医薬品コードシートのG列から医薬品名を取得
6. `ParseDrugString`関数が医薬品名を解析
7. `ExtractPackageTypeFromDrugName`関数が包装形態を抽出
8. tmp_tanaシートから一致する医薬品を検索
9. 結果が設定シートのC列に出力される

## 6. ワークシート構造

### 6.1 設定シート（Sheet1）

| セル | 内容 | 備考 |
|-----|------|------|
| A1:C1 | タイトル「医薬品名比較ツール」 | マージされたセル |
| A2 | 「【使い方】」 | 太字 |
| A3 | 使用手順1 | B4セルの参照を削除 |
| A5 | 使用手順2 | |
| A6:C6 | ヘッダー行 | 太字、背景色あり |
| A7:A30 | 行番号 | |
| B7:B30 | 検索医薬品名 | ユーザー入力 |
| C7:C30 | 一致医薬品名 | 処理結果 |
| A32 | 実行方法の説明 | イタリック体 |
| A34 | 「【GS1コード処理】」 | 太字 |
| A35:A36 | GS1コード処理の説明 | |

### 6.2 ターゲットシート（Sheet2）

| セル | 内容 | 備考 |
|-----|------|------|
| A1:B1 | タイトル「比較対象医薬品リスト」 | マージされたセル |
| A2:B2 | ヘッダー行 | 太字、背景色あり |
| A3:A30 | 行番号 | |
| B3:B30 | 医薬品名 | ユーザー入力 |

### 6.3 tmp_tanaシート

棚情報を管理するためのシートです。ShelfManagerモジュールで使用されます。

### 6.4 医薬品コードシート

医薬品コードと名称の対応表を含むシートです。GS1CodeProcessorモジュールで使用されます。

## 7. エラー処理

各モジュールには適切なエラー処理が実装されています。主なエラー処理は以下の通りです：

1. GTIN-14コードの検証エラー
2. 医薬品コードシートが見つからない場合のエラー
3. 医薬品が見つからない場合のエラー
4. 包装形態の抽出エラー

## 8. テスト計画

TestGTIN14Processing.basモジュールには以下のテスト関数が実装されています：

1. `TestPackageTypeExtraction` - 包装形態の直接抽出機能をテスト
2. `TestMainModulePackageExtraction` - MainModuleの包装形態抽出機能をテスト
3. `TestGTIN14CodeProcessing` - GTIN-14コード処理機能をテスト
4. `RunAllTests` - すべてのテストを実行

## 9. 今後の拡張性

システムは以下の点で拡張可能です：

1. 新しい包装形態の追加
2. 医薬品名解析アルゴリズムの改善
3. 他のコード体系（JANコードなど）への対応
4. ユーザーインターフェースの改善

## 10. 変更履歴

### バージョン 1.0.0
- 初期リリース

### バージョン 2.0.0
- GS1-128コード処理機能を追加

### バージョン 3.0.0
- 包装形態の直接抽出機能を追加
- B4セルのドロップダウンを廃止
- GTIN-14対応機能を追加
- 医薬品コードシートのG列（調剤包装単位名称）を使用するように変更
