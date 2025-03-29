VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SummaryUpdateForm 
   Caption         =   "まとめシート更新"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "SummaryUpdateForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SummaryUpdateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    ' まとめシート更新処理を実行
    SummaryIntegrationModule.UpdateAllReportSummaries
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    ' キャンセルボタン
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' フォームの初期化
    Me.Caption = "まとめシート更新"
    
    ' ラベルの設定
    Me.Label1.Caption = "すべてのレポートファイルのまとめシートを更新します。" & vbCrLf & _
                        "更新を実行するには「更新」ボタンをクリックしてください。"
    
    ' ボタンの設定
    Me.CommandButton1.Caption = "更新"
    Me.CommandButton2.Caption = "キャンセル"
End Sub
