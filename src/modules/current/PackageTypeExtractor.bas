Attribute VB_Name = "PackageTypeExtractor"
Option Explicit

' パッケージタイプのマッピングを保持する配列
Private packageMappings As Variant

' モジュール初期化時にマッピングをロード
Public Sub InitializePackageMappings()
    On Error GoTo ErrorHandler
    
    ' CSVから変換マッピングをロード
    Dim mappings As Collection
    Set mappings = New Collection
    
    ' マッピングデータの定義（CSV対応）
    ' 変換前,変換後の形式
    AddMapping mappings, "PTP", "/PTP/"
    AddMapping mappings, "ＰＴＰ", "/PTP/"
    AddMapping mappings, "バラ", "/バラ/"
    AddMapping mappings, "調剤用", "/調剤用/"
    AddMapping mappings, "分包", "/分包/"
    AddMapping mappings, "包装小", "/包装小/"
    AddMapping mappings, "SP", "/SP/"
    AddMapping mappings, "ＳＰ", "/SP/"
    AddMapping mappings, "PTP(患者用)", "/PTP(患者用)/"
    AddMapping mappings, "ＰＴＰ（患者用）", "/PTP(患者用)/"
    AddMapping mappings, "その他", "/その他(なし)/"
    AddMapping mappings, "未定義", "/未定義/"
    
    ' マッピングを配列に変換して保持
    Dim i As Long
    ReDim packageMappings(1 To mappings.Count, 1 To 2)
    
    For i = 1 To mappings.Count
        Dim pair As Variant
        pair = mappings(i)
        packageMappings(i, 1) = pair(0) ' 変換前
        packageMappings(i, 2) = pair(1) ' 変換後
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "パッケージマッピングの初期化中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' マッピングにペアを追加するヘルパー関数
Private Sub AddMapping(ByRef coll As Collection, ByVal source As String, ByVal target As String)
    Dim pair(0 To 1) As String
    pair(0) = source
    pair(1) = target
    coll.Add pair
End Sub

' 医薬品名から包装形態を抽出する関数
Public Function ExtractPackageTypeFromDrugName(ByVal drugName As String) As String
    On Error GoTo ErrorHandler
    
    ' マッピングが初期化されていない場合は初期化
    If IsEmpty(packageMappings) Then
        InitializePackageMappings
    End If
    
    ' まず既存の抽出方法を試す
    Dim packageType As String
    packageType = DrugNameParser.ExtractPackageTypeSimple(drugName)
    
    ' パッケージタイプが見つかった場合はマッピングを適用
    If Len(packageType) > 0 Then
        ExtractPackageTypeFromDrugName = ConvertPackageType(packageType)
        Exit Function
    End If
    
    ' スラッシュで囲まれた形式を確認
    Dim startPos As Long, endPos As Long
    startPos = InStr(1, drugName, "/")
    If startPos > 0 Then
        endPos = InStr(startPos + 1, drugName, "/")
        If endPos > startPos Then
            packageType = Mid(drugName, startPos + 1, endPos - startPos - 1)
            ExtractPackageTypeFromDrugName = ConvertPackageType(packageType)
            Exit Function
        End If
    End If
    
    ' 各マッピング元の文字列を直接検索
    Dim i As Long
    For i = 1 To UBound(packageMappings, 1)
        If InStr(1, drugName, packageMappings(i, 1), vbTextCompare) > 0 Then
            ExtractPackageTypeFromDrugName = packageMappings(i, 2)
            Exit Function
        End If
    Next i
    
    ' デフォルト値（見つからない場合）
    ExtractPackageTypeFromDrugName = "/未定義/"
    Exit Function
    
ErrorHandler:
    MsgBox "包装形態の抽出中にエラーが発生しました: " & Err.Description, vbCritical
    ExtractPackageTypeFromDrugName = "/未定義/"
End Function

' パッケージタイプを変換する関数
Public Function ConvertPackageType(ByVal packageType As String) As String
    ' マッピングが初期化されていない場合は初期化
    If IsEmpty(packageMappings) Then
        InitializePackageMappings
    End If
    
    Dim i As Long
    For i = 1 To UBound(packageMappings, 1)
        If StrComp(packageType, packageMappings(i, 1), vbTextCompare) = 0 Then
            ConvertPackageType = packageMappings(i, 2)
            Exit Function
        End If
    Next i
    
    ' 見つからない場合は元の値をスラッシュで囲む
    ConvertPackageType = "/" & packageType & "/"
End Function
