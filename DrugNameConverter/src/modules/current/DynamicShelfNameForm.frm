VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DynamicShelfNameForm 
   Caption         =   "棚名入力"
   ClientHeight    =   600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   520
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
Private Const MAX_HEIGHT As Long = 250 ' 最大フォーム高さ（約5ファイル表示）
Private Const ROWS_PER_FILE As Long = 2 ' ファイルごとの行数（ファイル名 + 棚名入力欄の行）
Private Const ROW_HEIGHT As Long = 20   ' 1行の高さ
Private Const CTRL_MARGIN As Long = 5  ' コントロール間のマージン
Private Const MAX_FILES As Long = 100   ' 最大ファイル数（600棚番号対応）
Private Const SHELF_INPUTS_PER_FILE As Long = 3 ' 各ファイルに対する棚名入力欄の数

' キャンセルフラグ
Private mIsCancelled As Boolean

' テキストボックスの配列
Private textBoxes() As MSForms.TextBox
Private textBoxes2() As MSForms.TextBox
Private textBoxes3() As MSForms.TextBox

' ラベルの配列（ファイル名とラベル）
Private fileLabels() As MSForms.Label
Private shelfLabels() As MSForms.Label
Private shelfLabels2() As MSForms.Label
Private shelfLabels3() As MSForms.Label

' CSVファイル数
Private mFileCount As Integer
   
' スクロール用のフレーム
Private scrollFrame As MSForms.Frame

' ボタン変数の宣言
Private WithEvents OKButton As MSForms.CommandButton
Private WithEvents CancelButton As MSForms.CommandButton

' フォーム初期化
Private Sub UserForm_Initialize()
    ' キャンセルフラグを初期化
    mIsCancelled = False
    
    ' デフォルトのボタンを作成
    CreateDefaultButtons
       
    ' MouseScrollを無効化（フリーズ問題対応）
    ' MouseScroll.EnableMouseScroll Me, True, True, True
End Sub

' デフォルトのボタンを作成する
Private Sub CreateDefaultButtons()
    ' OKボタンを動的に作成
    Set OKButton = Me.Controls.Add("Forms.CommandButton.1", "OKButton", True)
    With OKButton
        .Caption = "OK"
        .Top = Me.Height - 30
        .Left = 10
        .Width = 60
        .Height = 25
    End With
    
    ' キャンセルボタンを動的に作成
    Set CancelButton = Me.Controls.Add("Forms.CommandButton.1", "CancelButton", True)
    With CancelButton
        .Caption = "キャンセル"
        .Top = Me.Height - 30
        .Left = 80
        .Width = 80
        .Height = 25
    End With
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
    ReDim textBoxes2(1 To mFileCount)
    ReDim textBoxes3(1 To mFileCount)
    ReDim fileLabels(1 To mFileCount)
    ReDim shelfLabels(1 To mFileCount)
    ReDim shelfLabels2(1 To mFileCount)
    ReDim shelfLabels3(1 To mFileCount)
       
    ' 設定シートを取得
    Set settingsSheet = ThisWorkbook.Sheets("設定")
       
    ' スクロールフレームを作成
    Set scrollFrame = Me.Controls.Add("Forms.Frame.1", "ScrollFrame", True)
    With scrollFrame
        .Caption = ""
        .Left = 5
        .Top = 5
        .Width = Me.Width - 20
           
        ' フレームの高さを計算（約5ファイル表示）
        Dim visibleFiles As Integer
        visibleFiles = 5 ' 最大表示ファイル数
        
        ' 実際のファイル数と表示ファイル数を比較して小さい方を使用
        visibleFiles = IIf(mFileCount < visibleFiles, mFileCount, visibleFiles)
        
        ' ファイル表示領域の高さを計算
        frameHeight = (visibleFiles * ROWS_PER_FILE * ROW_HEIGHT) + 30
        .Height = frameHeight
           
        ' フォームサイズを固定（ボタン用の余白を追加）
        formHeight = MAX_HEIGHT
        Me.Height = formHeight
           
        ' スクロール設定
        If mFileCount > visibleFiles Then
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = (mFileCount * ROWS_PER_FILE * ROW_HEIGHT) + 20  ' スクロール領域の高さを設定
            .ScrollWidth = Me.Width - 30 ' スクロール領域の幅を設定
        Else
            .ScrollBars = fmScrollBarsNone
        End If
        
        ' フォームの幅を設定（すべての棚名入力欄が表示されるように）
        Me.Width = 520
    End With
       
    ' テキストボックスとラベルを生成
    For i = 1 To mFileCount
        ' ファイル名ラベルを作成
        Set fileLabels(i) = scrollFrame.Controls.Add("Forms.Label.1", "FileLabel" & i, True)
        With fileLabels(i)
            .Caption = IIf(Not IsEmpty(fileNames) And i <= UBound(fileNames), "ファイル: " & fileNames(i), "ファイル " & i)
            .Left = CTRL_MARGIN
            .Top = CTRL_MARGIN + ((i - 1) * ROWS_PER_FILE * ROW_HEIGHT)
            .Width = scrollFrame.Width - 20
            .Height = ROW_HEIGHT
        End With
           
        ' 棚名1ラベルを作成
        Set shelfLabels(i) = scrollFrame.Controls.Add("Forms.Label.1", "ShelfLabel" & i & "_1", True)
        With shelfLabels(i)
            .Caption = "棚名1:"
            .Left = CTRL_MARGIN
            .Top = CTRL_MARGIN + ((i - 1) * ROWS_PER_FILE * ROW_HEIGHT) + ROW_HEIGHT
            .Width = 40
            .Height = ROW_HEIGHT
        End With
           
        ' 棚名1テキストボックスを作成
        Set textBoxes(i) = scrollFrame.Controls.Add("Forms.TextBox.1", "TextBox" & i & "_1", True)
        With textBoxes(i)
            .Left = CTRL_MARGIN + 50
            .Top = CTRL_MARGIN + ((i - 1) * ROWS_PER_FILE * ROW_HEIGHT) + ROW_HEIGHT
            .Width = 60
            .Height = ROW_HEIGHT
            .MaxLength = 5
               
            ' 初期値は空に設定
            .Text = ""
        End With
        
        ' 棚名2ラベルを作成
        Set shelfLabels2(i) = scrollFrame.Controls.Add("Forms.Label.1", "ShelfLabel" & i & "_2", True)
        With shelfLabels2(i)
            .Caption = "棚名2:"
            .Left = CTRL_MARGIN + 120
            .Top = CTRL_MARGIN + ((i - 1) * ROWS_PER_FILE * ROW_HEIGHT) + ROW_HEIGHT
            .Width = 40
            .Height = ROW_HEIGHT
        End With
        
        ' 棚名2テキストボックスを作成
        Set textBoxes2(i) = scrollFrame.Controls.Add("Forms.TextBox.1", "TextBox" & i & "_2", True)
        With textBoxes2(i)
            .Left = CTRL_MARGIN + 170
            .Top = CTRL_MARGIN + ((i - 1) * ROWS_PER_FILE * ROW_HEIGHT) + ROW_HEIGHT
            .Width = 60
            .Height = ROW_HEIGHT
            .MaxLength = 5
            
            ' 初期値は空に設定
            .Text = ""
        End With
        
        ' 棚名3ラベルを作成
        Set shelfLabels3(i) = scrollFrame.Controls.Add("Forms.Label.1", "ShelfLabel" & i & "_3", True)
        With shelfLabels3(i)
            .Caption = "棚名3:"
            .Left = CTRL_MARGIN + 240
            .Top = CTRL_MARGIN + ((i - 1) * ROWS_PER_FILE * ROW_HEIGHT) + ROW_HEIGHT
            .Width = 40
            .Height = ROW_HEIGHT
        End With
        
        ' 棚名3テキストボックスを作成
        Set textBoxes3(i) = scrollFrame.Controls.Add("Forms.TextBox.1", "TextBox" & i & "_3", True)
        With textBoxes3(i)
            .Left = CTRL_MARGIN + 290
            .Top = CTRL_MARGIN + ((i - 1) * ROWS_PER_FILE * ROW_HEIGHT) + ROW_HEIGHT
            .Width = 60
            .Height = ROW_HEIGHT
            .MaxLength = 5
            
            ' 初期値は空に設定
            .Text = ""
        End With
    Next i
       
    ' OKボタンの位置を固定（常に表示されるように）
    OKButton.Top = formHeight - 40
    OKButton.Left = 10
    OKButton.Width = 60
    OKButton.Height = 25
    OKButton.ZOrder (0) ' 前面に配置
       
    ' キャンセルボタンの位置を固定（常に表示されるように）
    CancelButton.Top = formHeight - 40
    CancelButton.Left = 80
    CancelButton.Width = 80
    CancelButton.Height = 25
    CancelButton.ZOrder (0) ' 前面に配置
    
    ' ボタンが確実に表示されるようにスクロール領域の下端に余白を追加
    With scrollFrame
        If .ScrollBars = fmScrollBarsVertical Then
            .ScrollHeight = .ScrollHeight + 50
        End If
    End With
       
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
            ' 3つの棚名を設定シートの異なる列に保存
            settingsSheet.Cells(i, 2).Value = textBoxes(i).Text   ' 棚名1
            settingsSheet.Cells(i, 3).Value = textBoxes2(i).Text  ' 棚名2
            settingsSheet.Cells(i, 4).Value = textBoxes3(i).Text  ' 棚名3
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

' 棚名を取得するプロパティ（棚名1）
Public Property Get ShelfName(ByVal index As Integer) As String
    If index >= 1 And index <= mFileCount Then
        ShelfName = textBoxes(index).Text
    Else
        ShelfName = ""
    End If
End Property

' 棚名2を取得するプロパティ
Public Property Get ShelfName2(ByVal index As Integer) As String
    If index >= 1 And index <= mFileCount Then
        ShelfName2 = textBoxes2(index).Text
    Else
        ShelfName2 = ""
    End If
End Property

' 棚名3を取得するプロパティ
Public Property Get ShelfName3(ByVal index As Integer) As String
    If index >= 1 And index <= mFileCount Then
        ShelfName3 = textBoxes3(index).Text
    Else
        ShelfName3 = ""
    End If
End Property
