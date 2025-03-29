Attribute VB_Name = "DatabaseMenuModule"
Option Explicit

' メニューからデータベース作成・更新機能を呼び出すサブルーチン
Public Sub CreateOrUpdateDatabase()
    ' データベースシートを作成・更新する
    DatabaseSheetModule.UpdateDatabaseSheet
End Sub

' メニューからデータベース機能メニューを呼び出すサブルーチン
Public Sub ShowDatabaseMenu()
    ' データベース機能メニューを表示
    DatabaseOperationsModule.ShowDatabaseMenu
End Sub

' メニューからデータベース検索機能を呼び出すサブルーチン
Public Sub ShowDatabaseSearchForm()
    ' データベース検索機能を呼び出す
    DatabaseOperationsModule.SearchDatabase
End Sub

' メニューからデータベースCSV出力機能を呼び出すサブルーチン
Public Sub ExportDatabaseToCsv()
    ' データベースCSV出力機能を呼び出す
    DatabaseOperationsModule.ExportDatabaseToCsv
End Sub

' メニューからデータベース集計レポート作成機能を呼び出すサブルーチン
Public Sub CreateDatabaseSummaryReport()
    ' データベース集計レポート作成機能を呼び出す
    DatabaseOperationsModule.CreateDatabaseSummaryReport
End Sub
