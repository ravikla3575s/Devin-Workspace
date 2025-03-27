Attribute VB_Name = "ReportInvalidCodes"
Option Explicit

' 無効なGTINコードを報告するユーティリティ関数
Public Sub ReportInvalidCodes(ByRef invalidCodes As Collection)
    On Error GoTo ErrorHandler
    
    ' 無効なGTINコードがあれば報告
    If invalidCodes.Count > 0 Then
        Dim message As String
        Dim i As Integer
        
        message = "以下の" & invalidCodes.Count & "件のコードは14桁の数字ではないため、処理対象外としました:" & vbCrLf & vbCrLf
        
        For i = 1 To invalidCodes.Count
            If i <= 10 Then
                message = message & invalidCodes(i) & vbCrLf
            Else
                message = message & "... 他 " & (invalidCodes.Count - 10) & " 件"
                Exit For
            End If
        Next i
        
        MsgBox message, vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "無効なコードの報告中にエラーが発生しました: " & Err.Description, vbCritical
End Sub
