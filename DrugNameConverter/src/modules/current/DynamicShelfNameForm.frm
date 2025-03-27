VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DynamicShelfNameForm 
   Caption         =   "棚名入力"
   ClientHeight    =   100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   140
   OleObjectBlob   =   "DynamicShelfNameForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "DynamicShelfNameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 定数
Private Const MAX_HEIGHT As Long = 100 ' 最大フォーム高さ（これを超えるとスクロール可能に）
Private Const ROWS_PER_FILE As Long = 2 ' ファイルごとの行数
Private Const ROW_HEIGHT As Long = 15   ' 1行の高さ
Private Const CTRL_MARGIN As Long = 5   ' コントロール間のマージン
Private Const MAX_FILES As Long = 100   ' 最大ファイル数（600棚番号対応）

' キャンセルフラグ
Private mIsCancelled As Boolean

' テキストボックスの配列
Private textBoxes() As MSForms.TextBox

' ラベルの配列（ファイル名とラベル）
Private fileLabels() As MSForms.Label
Private shelfLabels() As MSForms.Label

' CSVファイル数
Private mFileCount As Integer
   
' スクロール用のフレーム
Private scrollFrame As MSForms.Frame

' フォーム初期化
Private Sub UserForm_Initialize()
    ' キャンセルフラグを初期化
    mIsCancelled = False
       
    ' MouseScrollを無効化（フリーズ問題対応）
    ' MouseScroll.EnableMouseScroll Me, True, True, True
End Sub

' フォーム終了時
Private Sub UserForm_Terminate()
    ' 明示的にスクロールを無効化（フリーズ問題対応のため無効化）
    ' MouseScroll.DisableMouseScroll Me
End Sub

' CSVファイル数を設定し、動的にコントロールを生成する
Public Sub SetFileCount(ByVal fileCount As Integer, Optional ByVal fileNames As Variant = Nothing)
    On Error GoTo ErrorHandler
       
    Dim i As Integer
    Dim topPosition As Single
    Dim settingsSheet As Worksheet
    Dim frameHeight As Long
    Dim formHeight As Long
       
    ' ファイル数を保存（最大MAX_FILESまで）
    If fileCount > MAX_FILES Then
        mFileCount = MAX_FILES
        MsgBox MAX_FILES & "個を超えるCSVファイルが検出されました。最初の" & MAX_FILES & "個のみ処理します。", vbInformation
    Else
        mFileCount = fileCount
    End If
       
    ' 配列を初期化
    ReDim textBoxes(1 To mFileCount)
    ReDim fileLabels(1 To mFileCount)
    ReDim shelfLabels(1 To mFileCount)
       
    ' 設定シートを取得
    Set settingsSheet = ThisWorkbook.Sheets("設定")
       
    ' スクロールフレームを作成
    Set scrollFrame = Me.Controls.Add("Forms.Frame.1", "ScrollFrame", True)
    With scrollFrame
        .Caption = ""
        .Left = 5
        .Top = 5
        .Width = Me.Width - 20
           
        ' フレームの高さを計算（ファイル数に基づく）
        frameHeight = (mFileCount * ROWS_PER_FILE * ROW_HEIGHT) + 60
        .Height = frameHeight
           
        ' ボタン用の余白を追加
        formHeight = frameHeight + 60
           
        ' フォームの高さを制限し、必要に応じてスクロール可能に
        If formHeight > MAX_HEIGHT Then
            Me.Height = MAX_HEIGHT
            .ScrollBars = fmScrollBarsVertical
        Else
            Me.Height = formHeight
            .ScrollBars = fmScrollBarsNone
        End If
    End With
       
    ' テキストボックスとラベルを生成
    For i = 1 To mFileCount
        ' ファイル名ラベルを作成
        Set fileLabels(i) = scrollFrame.Controls.Add("Forms.Label.1", "FileLabel" & i, True)
        With fileLabels(i)
            .Caption = IIf(Not IsEmpty(fileNames) And i <= UBound(fileNames), "ファイル: " & fileNames(i), "ファイル " & i)
            .Left = CTRL_MARGIN
            .Top = CTRL_MARGIN + ((i - 1) * ROWS_PER_FILE * ROW_HEIGHT)
            .Width = scrollFrame.Width - 15
            .Height = ROW_HEIGHT
        End With
           
        ' 棚名ラベルを作成
        Set shelfLabels(i) = scrollFrame.Controls.Add("Forms.Label.1", "ShelfLabel" & i, True)
        With shelfLabels(i)
            .Caption = "棚名 " & i & ":"
            .Left = CTRL_MARGIN
            .Top = CTRL_MARGIN + ((i - 1) * ROWS_PER_FILE * ROW_HEIGHT) + ROW_HEIGHT
            .Width = 30
            .Height = ROW_HEIGHT
        End With
           
        ' テキストボックスを作成
        Set textBoxes(i) = scrollFrame.Controls.Add("Forms.TextBox.1", "TextBox" & i, True)
        With textBoxes(i)
            .Left = CTRL_MARGIN + 10
            .Top = CTRL_MARGIN + ((i - 1) * ROWS_PER_FILE * ROW_HEIGHT) + ROW_HEIGHT
            .Width = 25
            .Height = ROW_HEIGHT - 5
            .MaxLength = 5
               
            ' 設定シートから既存の棚名を取得して表示（B1〜BN）
            If i <= MAX_FILES Then  ' 設定シートの制限を考慮
                .Text = settingsSheet.Cells(i, 2).Value
            End If
        End With
    Next i
       
    ' OKボタンの位置を調整
    OKButton.Top = Me.Height - 30
    OKButton.Left = 10
    OKButton.Width = 40
       
    ' キャンセルボタンの位置を調整
    CancelButton.Top = Me.Height - 30
    CancelButton.Left = 60
    CancelButton.Width = 40
       
    Exit Sub
       
ErrorHandler:
    MsgBox "フォームの初期化中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' OKボタンクリック時の処理
Private Sub OKButton_Click()
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim settingsSheet As Worksheet
    
    ' 設定シートを取得
    Set settingsSheet = ThisWorkbook.Sheets("設定")
    
    ' テキストボックスの値を設定シートに書き込む
    For i = 1 To mFileCount
        If i <= MAX_FILES Then  ' 設定シートの制限を考慮
            settingsSheet.Cells(i, 2).Value = textBoxes(i).Text
        End If
    Next i
    
    ' フォームを閉じる
    Me.Hide
    
    Exit Sub
    
ErrorHandler:
    MsgBox "設定の保存中にエラーが発生しました: " & Err.Description, vbCritical
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

' ファイル数を返すプロパティ
Public Property Get FileCount() As Integer
    FileCount = mFileCount
End Property

' 棚名を取得するプロパティ
Public Property Get ShelfName(ByVal index As Integer) As String
    If index >= 1 And index <= mFileCount Then
        ShelfName = textBoxes(index).Text
    Else
        ShelfName = ""
    End If
End Property
