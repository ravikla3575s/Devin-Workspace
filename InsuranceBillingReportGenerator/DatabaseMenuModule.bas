Attribute VB_Name = "DatabaseMenuModule"
Option Explicit

' メニューから売掛管理表作成・更新機能を呼び出すサブルーチン
Public Sub CreateOrUpdateDatabase()
    ' 売掛管理表シートを作成・更新する
    DatabaseSheetModule.UpdateDatabaseSheet
End Sub

' メニューから売掛管理表機能メニューを呼び出すサブルーチン
Public Sub ShowDatabaseMenu()
    ' 売掛管理表機能メニューを表示
    DatabaseOperationsModule.ShowDatabaseMenu
End Sub

' メニューから売掛管理表検索機能を呼び出すサブルーチン
Public Sub ShowDatabaseSearchForm()
    ' 売掛管理表検索機能を呼び出す
    DatabaseOperationsModule.SearchDatabase
End Sub

' メニューから売掛管理表CSV出力機能を呼び出すサブルーチン
Public Sub ExportDatabaseToCsv()
    ' 売掛管理表CSV出力機能を呼び出す
    DatabaseOperationsModule.ExportDatabaseToCsv
End Sub

' メニューから売掛管理表集計レポート作成機能を呼び出すサブルーチン
Public Sub CreateDatabaseSummaryReport()
    ' 売掛管理表集計レポート作成機能を呼び出す
    DatabaseOperationsModule.CreateDatabaseSummaryReport
End Sub
