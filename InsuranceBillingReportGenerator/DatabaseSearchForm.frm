VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatabaseSearchForm 
   Caption         =   "データベース検索"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   OleObjectBlob   =   "DatabaseSearchForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "DatabaseSearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' フォーム変数
Private mCancelled As Boolean

' フォーム初期化
Private Sub UserForm_Initialize()
    ' フォームの初期化
    Me.Caption = "データベース検索・フィルタリング"
    
    ' コンボボックスに請求先の選択肢を追加
    With Me.cmbBillingDestination
        .AddItem "(すべて)"
        .AddItem "社保"
        .AddItem "国保"
        .ListIndex = 0
    End With
    
    ' コンボボックスに区分の選択肢を追加
    With Me.cmbCategory
        .AddItem "(すべて)"
        .AddItem "未請求"
        .AddItem "返戻"
        .AddItem "減点"
        .AddItem "再請求"
        .AddItem "遅請求"
        .ListIndex = 0
    End With
    
    ' 日付範囲の初期化
    Me.txtDateFrom.Value = ""
    Me.txtDateTo.Value = ""
    
    ' 金額範囲の初期化
    Me.txtAmountFrom.Value = ""
    Me.txtAmountTo.Value = ""
    
    ' 検索テキストの初期化
    Me.txtSearchText.Value = ""
    
    ' キャンセルフラグを初期化
    mCancelled = False
End Sub

' 検索ボタンのクリックイベント
Private Sub btnSearch_Click()
    mCancelled = False
    Me.Hide
End Sub

' キャンセルボタンのクリックイベント
Private Sub btnCancel_Click()
    mCancelled = True
    Me.Hide
End Sub

' クリアボタンのクリックイベント
Private Sub btnClear_Click()
    ' フォームをリセット
    Me.cmbBillingDestination.ListIndex = 0
    Me.cmbCategory.ListIndex = 0
    Me.txtDateFrom.Value = ""
    Me.txtDateTo.Value = ""
    Me.txtAmountFrom.Value = ""
    Me.txtAmountTo.Value = ""
    Me.txtSearchText.Value = ""
End Sub

' キャンセルされたかどうかのプロパティ
Public Property Get Cancelled() As Boolean
    Cancelled = mCancelled
End Property

' 請求先の選択値を取得
Public Property Get SelectedBillingDestination() As String
    If Me.cmbBillingDestination.ListIndex = 0 Then
        SelectedBillingDestination = ""
    Else
        SelectedBillingDestination = Me.cmbBillingDestination.Value
    End If
End Property

' 区分の選択値を取得
Public Property Get SelectedCategory() As String
    If Me.cmbCategory.ListIndex = 0 Then
        SelectedCategory = ""
    Else
        SelectedCategory = Me.cmbCategory.Value
    End If
End Property

' 日付範囲（開始）を取得
Public Property Get DateFrom() As String
    DateFrom = Me.txtDateFrom.Value
End Property

' 日付範囲（終了）を取得
Public Property Get DateTo() As String
    DateTo = Me.txtDateTo.Value
End Property

' 金額範囲（最小）を取得
Public Property Get AmountFrom() As String
    AmountFrom = Me.txtAmountFrom.Value
End Property

' 金額範囲（最大）を取得
Public Property Get AmountTo() As String
    AmountTo = Me.txtAmountTo.Value
End Property

' 検索テキストを取得
Public Property Get SearchText() As String
    SearchText = Me.txtSearchText.Value
End Property
