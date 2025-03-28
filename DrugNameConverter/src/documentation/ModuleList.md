# 薬局在庫管理システム - モジュール一覧

このドキュメントは、薬局在庫管理システムの全モジュールを一覧化したものです。各モジュールの詳細な機能説明は、対応するマークダウンファイルを参照してください。

## VBAモジュール (.bas)

1. [MainModule.bas](./MainModule.md) - 薬品名比較と棚情報処理の中核機能
2. [DrugNameParser.bas](./DrugNameParser.md) - 薬品名の解析と構成要素への分解
3. [DrugNameConverter.bas](./DrugNameConverter.md) - UIとワークブック初期化
4. [ShelfManager.bas](./ShelfManager.md) - 棚情報の管理と更新
5. [StringUtils.bas](./StringUtils.md) - 文字列操作ユーティリティ
6. [GS1CodeProcessor.bas](./GS1CodeProcessor.md) - GS1/GTIN-14コードの処理
7. [ImportCSVToSheet2.bas](./ImportCSVToSheet2.md) - CSVファイルのインポート
8. [IntegratedSystemTest.bas](./IntegratedSystemTest.md) - 統合システムテスト
9. [MouseScroll.bas](./MouseScroll.md) - マウススクロール機能
10. [PackageTypeExtractor.bas](./PackageTypeExtractor.md) - パッケージタイプの抽出
11. [ProcessFileBatch.bas](./ProcessFileBatch.md) - ファイルバッチ処理
12. [ProcessSingleCSVFileWithArray.bas](./ProcessSingleCSVFileWithArray.md) - 配列を使用した単一CSVファイル処理
13. [CollectGarbage.bas](./CollectGarbage.md) - メモリ解放促進
14. [ReportInvalidCodes.bas](./ReportInvalidCodes.md) - 無効なコードの報告
15. [TestCSVImport.bas](./TestCSVImport.md) - CSVインポートのテスト
16. [TestCSVImportAndDynamicForm.bas](./TestCSVImportAndDynamicForm.md) - CSVインポートと動的フォームのテスト
17. [TestDynamicShelfForm.bas](./TestDynamicShelfForm.md) - 動的棚名フォームのテスト
18. [TestGTIN14Processing.bas](./TestGTIN14Processing.md) - GTIN-14処理のテスト
19. [TestSheet3DrugCodeReference.bas](./TestSheet3DrugCodeReference.md) - Sheet3薬品コード参照のテスト
20. [TestShelfManagement.bas](./TestShelfManagement.md) - 棚管理のテスト

## VBAフォーム (.frm)

1. [DynamicShelfNameForm.frm](./DynamicShelfNameForm.md) - 動的棚名入力フォーム
2. [ShelfNameForm.frm](./ShelfNameForm.md) - 棚名入力フォーム

## VBAクラス (.cls)

1. [MouseOverControl.cls](./MouseOverControl.md) - マウスオーバーコントロール
