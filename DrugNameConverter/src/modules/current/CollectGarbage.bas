Attribute VB_Name = "CollectGarbage"
Option Explicit

' メモリ解放を促進（VBAにはガベージコレクションがないため、明示的に促進）
Public Sub CollectGarbage()
    On Error Resume Next
    
    ' 大きな文字列を確保して解放することでメモリ解放を促進
    Dim tmp As String
    tmp = Space$(50000000 / 2)  ' 約25MBの文字列を確保
    tmp = ""  ' 解放
    
    ' 明示的にGCを呼び出し
    Application.MemoryFree
    
    ' 一時的にDoEventsを呼び出してUIの応答性を維持
    DoEvents
End Sub
