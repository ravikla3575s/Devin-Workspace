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
- SetupDirectPackageExtraction: 医薬品名から直接包装形態を抽出する機能を初期化する
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

## GTIN-14コード処理機能

### 新機能概要
GTIN-14（GS1-128の14桁コード）から医薬品情報を処理する機能を追加しました。この機能により：

1. GTIN-14コードから医薬品コードシートの医薬品名を読み取り
2. 医薬品名を以下の要素に分解：
   - 医薬品成分名
   - 製剤形態（錠、カプセル、散など）
   - 用量規格と単位
   - 屋号（「〇〇」形式）
   - 包装規格
   - 包装形態
   - 包装追加情報
3. tmp_tanaシートで一致する医薬品名を検索
4. 設定シートのC列に転記

### GTIN-14コードの構造
GTIN-14コードは以下の構造を持ちます：
1. **パッケージ・インジケーター（PI）** - 1桁目
   - 0：調剤包装単位
   - 1：販売包装単位
   - 2：元梱包装単位
2. **GS1事業者コード** - 2～8桁目（日本の場合、"49"で始まる）
3. **商品アイテムコード** - 9～13桁目
4. **チェックデジット** - 14桁目（誤読防止のための数値）

### モジュール間の依存関係
- `GS1CodeProcessor.bas` - GTIN-14コード処理の中核機能
- `MainModule.bas` - GTIN-14コード処理のメインインターフェース
- `DrugNameConverter.bas` - UI関連の機能とラッパー
- `DrugNameParser.bas` - 医薬品名の構文解析に使用

## 棚番一括更新システム

### 概要
棚番一括更新システムは、CSVファイルから読み込んだGTIN-14コードを使って医薬品情報を取得し、tmp_tanaシートの棚番情報を一括で更新する機能です。この機能には以下の特徴があります：

1. 最大3つのCSVファイルを一度に処理可能
2. 各CSVファイルに対応する棚名を指定可能
3. GTIN-14コードから医薬品名を自動検索
4. tmp_tanaシート内の医薬品を部分一致で検索
5. 棚名の一括更新および元に戻す機能
6. 更新結果のCSVエクスポート

### モジュール構成
- `ShelfManager.bas` - 棚番一括更新システムの中核機能
- `ShelfNameForm.frm` - 棚名入力用のユーザーフォーム
- `GS1CodeProcessor.bas` - GTIN-14コード処理機能（連携）
- `MainModule.bas` - メインメニューおよび連携機能

### 主要機能
- `Main()` - 棚番一括更新処理のエントリーポイント
- `ImportCSVFiles(folderPath)` - CSVファイルからGTIN-14コードを取り込む
- `ProcessItems()` - 取り込んだGTINコードを処理し棚番を更新
- `GetDrugName(gtin)` - GTINコードから医薬品名を取得
- `FindMedicineRowByName(drugName)` - 医薬品名からtmp_tanaの行を検索
- `UndoShelfNames()` - 棚名を元に戻す
- `ExportTemplateCSV()` - 更新後のtmp_tanaシートをCSVに出力（設定シートB4のパスを使用）
- `SetOutputFilePath()` - テンプレートファイルの出力先パスを設定
