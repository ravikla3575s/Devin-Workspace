Attribute VB_Name = "GS1CodeProcessor"
Option Explicit

' GS1-128の14桁コードから医薬品情報を処理する構造体
Public Type DrugInfo
    GS1Code As String           ' GS1-128コード
    DrugName As String          ' 医薬品名
    BaseName As String          ' 医薬品成分名
    FormType As String          ' 製剤形態（錠、カプセル、散など）
    Strength As String          ' 用量規格と単位
    Maker As String             ' 屋号（「〇〇」形式）
    PackageSpec As String       ' 包装規格
    PackageForm As String       ' 包装形態
    PackageAddInfo As String    ' 包装追加情報
End Type

' GS1-128コードから医薬品情報を取得する関数
Public Function GetDrugInfoFromGS1Code(ByVal gs1Code As String) As DrugInfo
    On Error GoTo ErrorHandler
    
    Dim result As DrugInfo
    Dim wbDrugCode As Workbook
    Dim wsDrugCode As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim drugName As String
    
    ' 医薬品コードシートを開く
    Set wbDrugCode = Workbooks.Open(Application.ThisWorkbook.Path & Application.PathSeparator & "医薬品コード.xlsx")
    Set wsDrugCode = wbDrugCode.Sheets(1)
    
    ' GS1コードを結果に格納
    result.GS1Code = gs1Code
    
    ' 最終行を取得
    lastRow = wsDrugCode.Cells(wsDrugCode.Rows.Count, 1).End(xlUp).Row
    
    ' GS1コードに一致する医薬品を検索
    For i = 2 To lastRow ' ヘッダー行をスキップ
        If wsDrugCode.Cells(i, 1).Value = gs1Code Then
            drugName = wsDrugCode.Cells(i, 2).Value
            result.DrugName = drugName
            
            ' 医薬品名を各要素に分解
            Dim drugParts As DrugNameParts
            drugParts = ParseDrugString(drugName)
            
            ' 構造体にデータを格納
            result.BaseName = drugParts.BaseName
            result.FormType = drugParts.formType
            result.Strength = drugParts.strength
            result.Maker = drugParts.maker
            result.PackageForm = drugParts.Package
            
            ' 追加情報の処理（必要に応じて実装）
            result.PackageSpec = ExtractPackageSpecFromDrugName(drugName)
            result.PackageAddInfo = ExtractPackageAddInfoFromDrugName(drugName)
            
            Exit For
        End If
    Next i
    
    ' 医薬品コードシートを閉じる
    wbDrugCode.Close SaveChanges:=False
    
    GetDrugInfoFromGS1Code = result
    Exit Function
    
ErrorHandler:
    If Not wbDrugCode Is Nothing Then
        wbDrugCode.Close SaveChanges:=False
    End If
    MsgBox "GS1コード処理中にエラーが発生しました: " & Err.Description, vbCritical
    ' 空の結果を返す
    GetDrugInfoFromGS1Code = result
End Function

' 包装規格を医薬品名から抽出する関数
Private Function ExtractPackageSpecFromDrugName(ByVal drugName As String) As String
    ' 数字+単位のパターンを探す（例: 100錠、10カプセルなど）
    Dim regex As Object
    Dim matches As Object
    Dim result As String
    
    ' CreateObjectを使わないバージョン
    Dim i As Long
    Dim inNumber As Boolean
    Dim numStart As Long
    
    inNumber = False
    result = ""
    
    For i = 1 To Len(drugName)
        Dim c As String
        c = Mid(drugName, i, 1)
        
        If IsNumeric(c) Then
            If Not inNumber Then
                inNumber = True
                numStart = i
            End If
        ElseIf c = " " And inNumber Then
            ' スペースは含める
        Else
            If inNumber Then
                ' 数字の後に単位があるか確認
                Dim units As Variant
                units = Array("錠", "カプセル", "包", "枚", "本", "袋", "瓶", "管")
                
                Dim j As Long
                For j = 0 To UBound(units)
                    If InStr(i, drugName, units(j)) > 0 Then
                        result = Mid(drugName, numStart, i - numStart + Len(units(j)))
                        Exit Function
                    End If
                Next j
                
                inNumber = False
            End If
        End If
    Next i
    
    ExtractPackageSpecFromDrugName = result
End Function

' 包装追加情報を医薬品名から抽出する関数
Private Function ExtractPackageAddInfoFromDrugName(ByVal drugName As String) As String
    ' 括弧内の情報を抽出
    Dim startPos As Long, endPos As Long
    Dim result As String
    
    startPos = InStr(1, drugName, "(")
    If startPos = 0 Then startPos = InStr(1, drugName, "（")
    
    If startPos > 0 Then
        endPos = InStr(startPos, drugName, ")")
        If endPos = 0 Then endPos = InStr(startPos, drugName, "）")
        
        If endPos > startPos Then
            result = Mid(drugName, startPos, endPos - startPos + 1)
        End If
    End If
    
    ExtractPackageAddInfoFromDrugName = result
End Function

' GS1コードから医薬品名を検索して設定シートに転記する
Public Sub ProcessGS1CodeAndUpdateSettings(ByVal gs1Code As String)
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' GS1コードから医薬品情報を取得
    Dim drugInfo As DrugInfo
    drugInfo = GetDrugInfoFromGS1Code(gs1Code)
    
    ' 医薬品情報が取得できなかった場合
    If Len(drugInfo.DrugName) = 0 Then
        MsgBox "指定されたGS1コード: " & gs1Code & " に対応する医薬品が見つかりませんでした。", vbExclamation
        GoTo CleanExit
    End If
    
    ' tmp_tanaシートから一致する医薬品を検索
    Dim wsTmpTana As Worksheet
    Dim lastRowTana As Long
    Dim i As Long
    Dim matchFound As Boolean
    
    Set wsTmpTana = ThisWorkbook.Worksheets("tmp_tana")
    lastRowTana = wsTmpTana.Cells(wsTmpTana.Rows.Count, 2).End(xlUp).Row
    matchFound = False
    
    For i = 2 To lastRowTana
        ' 医薬品名の一致をチェック
        If InStr(1, wsTmpTana.Cells(i, 2).Value, drugInfo.DrugName, vbTextCompare) > 0 Then
            ' 設定シートのC列に転記
            Dim settingsSheet As Worksheet
            Dim emptyRow As Long
            
            Set settingsSheet = ThisWorkbook.Worksheets(1)
            
            ' 空いている行を見つける（C7以降）
            For emptyRow = 7 To 50
                If Len(Trim(settingsSheet.Cells(emptyRow, "C").Value)) = 0 Then
                    Exit For
                End If
            Next emptyRow
            
            ' 見つかった行に転記
            settingsSheet.Cells(emptyRow, "C").Value = wsTmpTana.Cells(i, 2).Value
            
            matchFound = True
            Exit For
        End If
    Next i
    
    ' マッチする医薬品が見つからなかった場合
    If Not matchFound Then
        MsgBox "指定された医薬品「" & drugInfo.DrugName & "」に対応するtmp_tanaの商品が見つかりませんでした。", vbExclamation
    End If
    
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "処理中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' GS1コードから医薬品情報を二次元配列として取得
Public Function GetDrugInfoAsArray(ByVal gs1Code As String) As Variant
    Dim drugInfo As DrugInfo
    Dim result(1 To 8) As Variant
    
    ' 医薬品情報を取得
    drugInfo = GetDrugInfoFromGS1Code(gs1Code)
    
    ' 二次元配列に格納
    result(1) = drugInfo.BaseName        ' 医薬品成分名
    result(2) = drugInfo.FormType        ' 製剤形態
    result(3) = drugInfo.Strength        ' 用量規格と単位
    result(4) = drugInfo.Maker           ' 屋号
    result(5) = drugInfo.PackageSpec     ' 包装規格
    result(6) = drugInfo.PackageForm     ' 包装形態
    result(7) = drugInfo.PackageAddInfo  ' 包装追加情報
    result(8) = drugInfo.DrugName        ' 完全な医薬品名
    
    GetDrugInfoAsArray = result
End Function
