Attribute VB_Name = "BillingDiscrepancyMenuModule"
Option Explicit

' メニューから請求誤差追求報告書作成機能を呼び出すサブルーチン
Public Sub ShowBillingDiscrepancyForm()
    ' 請求誤差追求報告書作成機能を呼び出す
    BillingDiscrepancyModule.CreateBillingDiscrepancyReport
End Sub
