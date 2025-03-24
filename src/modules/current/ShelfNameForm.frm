VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ShelfNameForm 
   Caption         =   "棚名入力"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ShelfNameForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ShelfNameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' キャンセルフラグ
Private mIsCancelled As Boolean

' フォーム初期化
Private Sub UserForm_Initialize()
    ' キャンセルフラグを初期化
    mIsCancelled = False
    
    ' テキストボックスの最大文字数を設定
    TextBox1.MaxLength = 5
    TextBox2.MaxLength = 5
    TextBox3.MaxLength = 5
    
    ' 設定シートから既存の棚名を取得して表示
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Sheets("設定")
    
    TextBox1.Text = settingsSheet.Cells(1, 2).Value
    TextBox2.Text = settingsSheet.Cells(2, 2).Value
    TextBox3.Text = settingsSheet.Cells(3, 2).Value
End Sub

' OKボタンクリック時の処理
Private Sub OKButton_Click()
    ' 設定シートに棚名を書き込む
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Sheets("設定")
    
    settingsSheet.Cells(1, 2).Value = TextBox1.Text
    settingsSheet.Cells(2, 2).Value = TextBox2.Text
    settingsSheet.Cells(3, 2).Value = TextBox3.Text
    
    ' フォームを閉じる
    Me.Hide
End Sub

' キャンセルボタンクリック時の処理
Private Sub CancelButton_Click()
    ' キャンセルフラグを設定
    mIsCancelled = True
    
    ' フォームを閉じる
    Me.Hide
End Sub

' キャンセルされたかどうかを返すプロパティ
Public Property Get IsCancelled() As Boolean
    IsCancelled = mIsCancelled
End Property
