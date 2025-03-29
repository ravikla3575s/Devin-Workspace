Attribute VB_Name = "MenuIntegrationModule"
Option Explicit

' メニューを作成するサブルーチン
Public Sub CreateCustomMenu()
    On Error Resume Next
    
    ' 既存のメニューを削除
    Application.CommandBars("Worksheet Menu Bar").Controls("保険請求管理").Delete
    
    ' 新しいメニューを作成
    Dim menu_bar As CommandBar
    Dim menu_item As CommandBarControl
    Dim sub_menu As CommandBarPopup
    
    Set menu_bar = Application.CommandBars("Worksheet Menu Bar")
    Set menu_item = menu_bar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    menu_item.Caption = "保険請求管理"
    
    ' CSVファイル処理メニュー
    With menu_item.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = "CSVファイルから報告書作成"
        .OnAction = "MainModule.CreateReportsFromCSV"
        .FaceId = 23
    End With
    
    ' パス設定メニュー
    With menu_item.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = "テンプレート・保存先設定"
        .OnAction = "SetTemplateAndSavePath.SetPaths"
        .FaceId = 17
    End With
    
    ' セパレーター
    menu_item.Controls.Add Type:=msoControlSeparator, Temporary:=True
    
    ' まとめシート更新メニュー
    With menu_item.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = "まとめシート更新"
        .OnAction = "SummaryMenuModule.ShowSummaryUpdateForm"
        .FaceId = 37
    End With
    
    ' 半期決算書作成メニュー
    With menu_item.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = "半期売掛金繰越額計算"
        .OnAction = "HalfYearMenuModule.ShowHalfYearCalculationForm"
        .FaceId = 183
    End With
    
    ' 請求誤差追求報告書メニュー
    With menu_item.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = "請求誤差追求報告書作成"
        .OnAction = "BillingDiscrepancyMenuModule.ShowBillingDiscrepancyForm"
        .FaceId = 184
    End With
    
    ' セパレーター
    menu_item.Controls.Add Type:=msoControlSeparator, Temporary:=True
    
    ' ヘルプメニュー
    With menu_item.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = "ヘルプ"
        .OnAction = "MenuIntegrationModule.ShowHelp"
        .FaceId = 487
    End With
End Sub

' ヘルプを表示するサブルーチン
Public Sub ShowHelp()
    MsgBox "保険請求管理システム ヘルプ" & vbCrLf & vbCrLf & _
           "【CSVファイルから報告書作成】" & vbCrLf & _
           "CSVファイルを選択して報告書を作成します。" & vbCrLf & vbCrLf & _
           "【テンプレート・保存先設定】" & vbCrLf & _
           "テンプレートファイルと保存先フォルダを設定します。" & vbCrLf & vbCrLf & _
           "【まとめシート更新】" & vbCrLf & _
           "既存の報告書のまとめシートを更新します。" & vbCrLf & vbCrLf & _
           "【半期売掛金繰越額計算】" & vbCrLf & _
           "半期ごとの売掛金繰越額を計算します。" & vbCrLf & vbCrLf & _
           "【請求誤差追求報告書作成】" & vbCrLf & _
           "減点・査定データから請求誤差追求報告書を作成します。", _
           vbInformation, "ヘルプ"
End Sub

' ワークブックを開いたときに自動的にメニューを作成
Public Sub Auto_Open()
    CreateCustomMenu
End Sub

' ワークブックを閉じるときにメニューを削除
Public Sub Auto_Close()
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls("保険請求管理").Delete
End Sub
