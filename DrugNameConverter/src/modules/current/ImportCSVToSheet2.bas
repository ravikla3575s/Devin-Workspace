Attribute VB_Name = "ImportCSVToSheet2"
Option Explicit

' 棚番テンプレートCSVファイルをシート2に転記する
Public Sub ImportCSVToSheet2()
    On Error GoTo ErrorHandler
    
    Dim filePath As String
    Dim fileNum As Integer
    Dim line As String
    Dim row As Long
    Dim cols As Variant
    Dim i As Integer
    Dim targetSheet As Worksheet
    Dim fileName As String
    Dim userResponse As VbMsgBoxResult
    
    ' ファイル選択ダイアログを表示
    filePath = GetCSVFilePath()
    If filePath = "" Then
        MsgBox "ファイルが選択されていないため、処理を中止します。", vbExclamation
        Exit Sub
    End If
    
    ' ファイル名を取得
    fileName = GetFileName(filePath)
    
    ' ファイル名が「tmp_tana.CSV」でない場合、確認ダイアログを表示
    If LCase(fileName) <> "tmp_tana.csv" Then
        userResponse = MsgBox("選択されたファイル名は「tmp_tana.CSV」ではありません。" & vbCrLf & _
                             "ファイル名: " & fileName & vbCrLf & vbCrLf & _
                             "このファイルを使用して更新しますか？", _
                             vbQuestion + vbYesNo, "ファイル名確認")
        
        ' ユーザーがキャンセルした場合は処理中止
        If userResponse = vbNo Then
            MsgBox "処理を中止しました。", vbInformation
            Exit Sub
        End If
    End If
    
    ' ターゲットシート（シート2）を取得
    Set targetSheet = ThisWorkbook.Sheets("ターゲット")
    
    ' 既存データをクリア
    targetSheet.UsedRange.ClearContents
    
    ' 進捗状況表示
    Application.StatusBar = "CSVファイルを読み込んでいます..."
    
    ' 画面更新を停止（パフォーマンス向上）
    Application.ScreenUpdating = False
    
    ' CSVファイルを開く
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    ' 開始行を設定
    row = 1
    
    ' ファイルの各行を読み込む
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        
        ' 空行をスキップ
        If Trim(line) <> "" Then
            ' カンマで分割
            cols = Split(line, ",")
            
            ' A〜I列にデータを転記（列数が足りない場合は空白）
            For i = 0 To 8  ' A〜I列（0〜8）
                If i <= UBound(cols) Then
                    targetSheet.Cells(row, i + 1).Value = cols(i)
                End If
            Next i
            
            ' 次の行へ
            row = row + 1
        End If
        
        ' 進捗状況を更新
        If row Mod 100 = 0 Then
            Application.StatusBar = "CSVファイルを読み込んでいます... (" & row & "行)"
            DoEvents
        End If
    Loop
    
    ' ファイルを閉じる
    Close #fileNum
    
    ' 画面更新を再開
    Application.ScreenUpdating = True
    
    ' 進捗状況表示をクリア
    Application.StatusBar = False
    
    ' 完了メッセージ
    MsgBox row - 1 & "行のデータをシート2に転記しました。" & vbCrLf & _
           "ファイル: " & filePath, vbInformation
    
    Exit Sub
    
ErrorHandler:
    ' エラー発生時の処理
    On Error Resume Next
    Close #fileNum
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "CSVファイルの読み込み中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' CSVファイルのパスを取得する
Private Function GetCSVFilePath() As String
    Dim fileDialog As FileDialog
    
    Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .Title = "棚番テンプレートCSVファイルを選択してください"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            GetCSVFilePath = .SelectedItems(1)
        Else
            GetCSVFilePath = ""
        End If
    End With
End Function

' ファイルパスからファイル名を取得する
Private Function GetFileName(filePath As String) As String
    Dim parts As Variant
    Dim lastPart As String
    
    ' パスの区切り文字（\）で分割
    parts = Split(filePath, "\")
    
    ' 最後の部分がファイル名
    lastPart = parts(UBound(parts))
    
    GetFileName = lastPart
End Function
