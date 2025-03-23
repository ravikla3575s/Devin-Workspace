# モジュールドキュメント

## 依存関係

VBAモジュールには以下の依存関係があります：

- **MainModule.bas** は以下に依存します：
  - DrugNameParser.bas（医薬品名解析関数用）
  - StringUtils.bas（文字列操作関数用）

- **DrugNameConverter.bas** は以下に依存します：
  - MainModule.bas（MainModuleの関数を呼び出す）

- **ShelfManager.bas** は比較的独立しています

- **DrugNameParser.bas** は以下に依存します：
  - StringUtils.bas（ExtractBetweenQuotesなどの関数用）

- **StringUtils.bas** は外部依存関係がありません

## 機能

### MainModule.bas
医薬品名のマッチングと比較のためのメイン処理関数を含みます：
- MainProcess: 類似性に基づいて医薬品マッチを処理するメイン関数
- SearchAndTransferDrugData: 医薬品データを検索して転記する
- ProcessDrugNamesWithMatchRate: マッチ率計算で医薬品名を処理する
- CompareAndTransferDrugNamesByPackage: 包装タイプ別に医薬品名を比較する
- ProcessFromRow7: 7行目以降の医薬品名を処理する

### DrugNameConverter.bas
UIとワークブック初期化のためのヘルパー関数を含みます：
- RunDrugNameComparison: 医薬品名比較を実行するエントリーポイント
- SetupPackageTypeDropdown: UIにパッケージタイプのドロップダウンを設定する
- AddInstructions: Sheet1に指示を追加する
- InitWorkbook: ワークブックの書式設定と設定を初期化する

### ShelfManager.bas
棚情報を管理するための関数を含みます：
- UpdateShelfNumbersWithShelfInfo: 棚情報で棚番号を更新する
- ExportToCSV: データをCSV形式でエクスポートする

### DrugNameParser.bas
医薬品名を構造的コンポーネントに解析するための関数を含みます：
- ParseDrugString: 医薬品文字列をコンポーネントに解析するメイン関数
- ExtractBaseNameSimple: 医薬品文字列から基本名を抽出する
- ExtractFormTypeSimple: 医薬品文字列から剤形タイプを抽出する
- ExtractStrengthSimple: 医薬品文字列から強度を抽出する
- ExtractPackageTypeSimple: 医薬品文字列からパッケージタイプを抽出する
- CompareDrugStringsWithRate: 医薬品文字列を比較してマッチ率を返す

### StringUtils.bas
文字列操作のためのユーティリティ関数を含みます：
- ExtractBetweenQuotes: 日本語の引用符の間のテキストを抽出する
- ExtractStrength: 文字列から強度情報を抽出する
- ExtractNumberAndUnit: 文字列から数値と単位を抽出する
- CompareStrength: 強度文字列を比較する
- SetupPackageTypeDropdown: パッケージタイプのドロップダウンを設定する
