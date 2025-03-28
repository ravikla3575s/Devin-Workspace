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

## 医薬品コード参照方法の変更

### 変更内容
- 医薬品コードの参照を外部ファイル「医薬品コード.xlsx」からワークブックのSheet3（医薬品コード）に変更
- GS1CodeProcessor.GetDrugInfoFromGS1Code関数を修正してSheet3を使用するように変更
- ShelfManager.GetDrugName関数を簡略化

### 修正理由
- 外部ファイル依存を削減し、単一ワークブック内で完結するようにする
- ファイル操作に関連するエラーを減少させる
- パフォーマンスの向上（外部ファイルのオープン/クローズが不要）

### 影響範囲
- GTIN-14コード処理機能
- 棚番一括更新機能
- その他医薬品コード参照を行う処理

### メリット
1. **運用の簡素化**
   - 外部ファイル管理が不要になる
   - シングルワークブックで完結するため配布・運用が容易になる

2. **エラー発生リスクの低減**
   - 外部ファイルが見つからない、開けないなどのエラーが発生しなくなる
   - ファイルロックによる競合が発生しない

3. **パフォーマンス向上**
   - 外部ファイルのオープン/クローズ処理が不要になり処理が高速化する
   - メモリ使用量の削減

4. **既存機能との互換性**
   - ShelfManagerは既にSheet3をフォールバック参照として使用しているため、修正の影響を最小限に抑えられる

## 棚番一括更新システム

### 概要
棚番一括更新システムは、CSVファイルから読み込んだGTIN-14コードを使って医薬品情報を取得し、tmp_tanaシートの棚番情報を一括で更新する機能です。この機能には以下の特徴があります：

1. フォルダ内のCSVファイル数を自動カウントし、動的にフォームを生成
2. 最大100個のCSVファイルを一度に処理可能
3. 各CSVファイルに対応する棚名を指定可能
4. ファイル数が多い場合はスクロール可能なフォームで対応
5. GTIN-14コードから医薬品名を自動検索
6. tmp_tanaシート内の医薬品を部分一致で検索
7. 棚名の一括更新および元に戻す機能
8. 更新結果のCSVエクスポート

### モジュール構成
- `ShelfManager.bas` - 棚番一括更新システムの中核機能
- `DynamicShelfNameForm.frm` - 動的棚名入力用のユーザーフォーム（スクロール機能付き）
- `MouseOverControl.cls` - マウス操作検知用のクラス（スクロール機能用）
- `MouseScroll.bas` - ユーザーフォームのマウスホイールスクロール機能を提供
- `GS1CodeProcessor.bas` - GTIN-14コード処理機能（連携）
- `MainModule.bas` - メインメニューおよび連携機能
- `ImportCSVToSheet2.bas` - 棚番テンプレートCSVファイルをシート2に転記する機能

### 主要機能
- `Main()` - 棚番一括更新処理のエントリーポイント
- `CountCSVFiles(folderPath)` - フォルダ内のCSVファイル数をカウント
- `GetCSVFileNames(folderPath, maxCount)` - フォルダ内のCSVファイル名を取得
- `ImportCSVFiles(folderPath)` - CSVファイルからGTIN-14コードを取り込む
- `ProcessItems()` - 取り込んだGTINコードを処理し棚番を更新
- `GetDrugName(gtin)` - GTINコードから医薬品名を取得
- `FindMedicineRowByName(drugName)` - 医薬品名からtmp_tanaの行を検索
- `UndoShelfNames()` - 棚名を元に戻す
- `ExportTemplateCSV()` - 更新後のtmp_tanaシートをCSVに出力（設定シートB4のパスを使用）
- `SetOutputFilePath()` - テンプレートファイルの出力先パスを設定

### ImportCSVToSheet2モジュール
棚番テンプレートCSVファイルをシート2（ターゲット）に転記する機能を提供します：
- `ImportCSVToSheet2()` - CSVファイルを選択し、A〜I列のデータをシート2に転記
- `GetCSVFilePath()` - CSVファイル選択ダイアログを表示
- `GetFileName()` - ファイルパスからファイル名を取得

特徴：
- ファイル名が「tmp_tana.CSV」でない場合、ユーザーに確認ダイアログを表示
- A〜I列（1行目から最終行まで）のデータを転記
- 既存データをクリアしてから新しいデータを転記
- 進捗状況をステータスバーに表示
