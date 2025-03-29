Attribute VB_Name = "HalfYearMenuModule"
Option Explicit

' メニューから半期決算書作成機能を呼び出すサブルーチン
Public Sub ShowHalfYearCalculationForm()
    ' 半期決算書作成機能を呼び出す
    HalfYearCalculationModule.CalculateAccountsReceivable
End Sub
