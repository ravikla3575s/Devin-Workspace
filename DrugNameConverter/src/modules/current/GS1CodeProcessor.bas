Attribute VB_Name = "GS1CodeProcessor"
Option Explicit

' GTIN-14コードから医薬品情報を処理する構造体
Public Type DrugInfo
    GS1Code As String           ' GTIN-14コード
    PackageIndicator As String  ' パッケージ・インジケーター（0:調剤包装単位, 1:販売包装単位, 2:元梱包装単位）
    DrugName As String          ' 医薬品名
    BaseName As String          ' 医薬品成分名
    FormType As String          ' 製剤形態（錠、カプセル、散など）
    Strength As String          ' 用量規格と単位
    Maker As String             ' 屋号（「〇〇」形式）
    PackageSpec As String       ' 包装規格
    PackageForm As String       ' 包装形態
    PackageAddInfo As String    ' 包装追加情報
End Type

' GTIN-14コードを検証する関数
Private Function ValidateGTIN14(ByVal gtinCode As String) As String
    ' 数字のみを抽出
    Dim i As Long
    Dim result As String
    
    result = ""
    For i = 1 To Len(gtinCode)
        If IsNumeric(Mid(gtinCode, i, 1)) Then
            result = result & Mid(gtinCode, i, 1)
        End If
    Next i
    
    ' 14桁であることを確認
    If Len(result) <> 14 Then
        result = ""
    End If
    
    ValidateGTIN14 = result
End Function

' GTIN-14コードからパッケージ・インジケーターを取得する関数
Private Function GetPackageIndicator(ByVal gtin14 As String) As String
    If Len(gtin14) >= 1 Then
        GetPackageIndicator = Left(gtin14, 1)
    Else
        GetPackageIndicator = ""
    End If
End Function

' GTIN-14コードから医薬品情報を取得する関数
Public Function GetDrugInfoFromGS1Code(ByVal gs1Code As String) As DrugInfo
    On Error GoTo ErrorHandler
    
    Dim result As DrugInfo
    Dim ws3 As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim drugName As String
    Dim validatedCode As String
    Dim found As Boolean
    
    ' GTIN-14コードを検証
    validatedCode = ValidateGTIN14(gs1Code)
    
    ' 無効なコードの場合
    If Len(validatedCode) <> 14 Then
        MsgBox "入力されたコードは有効なGTIN-14形式ではありません。14桁の数字を入力してください。", vbExclamation
        Exit Function
    End If
    
    ' 医薬品コードシート（Sheet3）を取得
    Set ws3 = ThisWorkbook.Worksheets(3) ' 医薬品コードシート
    
    ' GS1コードを結果に格納
    result.GS1Code = validatedCode
    
    ' パッケージ・インジケーターを取得
    result.PackageIndicator = GetPackageIndicator(validatedCode)
    
    ' 最終行を取得
    lastRow = ws3.Cells(ws3.Rows.Count, "F").End(xlUp).Row
    
    ' 検索フラグを初期化
    found = False
    
    ' Sheet3でGTINコードに一致する医薬品を検索
    For i = 2 To lastRow ' ヘッダー行をスキップ
        ' データベース内のコードも検証して比較
        Dim dbCode As String
        dbCode = ValidateGTIN14(CStr(ws3.Cells(i, "F").Value))
        
        If CStr(dbCode) = CStr(validatedCode) Then
            ' G列から医薬品名を取得
            drugName = ws3.Cells(i, "G").Value
            result.DrugName = drugName
            
            ' 医薬品名を各要素に分解
            Dim drugParts As DrugNameParts
            drugParts = ParseDrugString(drugName)
            
            ' 構造体にデータを格納
            result.BaseName = drugParts.BaseName
            result.FormType = drugParts.formType
            result.Strength = drugParts.strength
            result.Maker = drugParts.maker
            
            ' 包装形態を抽出
            ' PackageTypeExtractorモジュールを使用して包装形態を抽出
            result.PackageForm = PackageTypeExtractor.ExtractPackageTypeFromDrugName(drugName)
            
            ' 追加情報の処理
            result.PackageSpec = ExtractPackageSpecFromDrugName(drugName)
            result.PackageAddInfo = ExtractPackageAddInfoFromDrugName(drugName)
            
            found = True
            Exit For
        End If
    Next i
    
    ' 見つからなかった場合のデフォルト値設定
    If Not found Then
        result.GS1Code = CStr(validatedCode)
        result.DrugName = "/未登録/"
    End If
    
    GetDrugInfoFromGS1Code = result
    Exit Function
    
ErrorHandler:
    ' エラーハンドリング
    result.GS1Code = validatedCode
    result.DrugName = "/エラー: " & Err.Description & "/"
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
    If Len(CStr(drugInfo.DrugName)) = 0 Then
        MsgBox "指定されたGS1コード: " & CStr(gs1Code) & " に対応する医薬品が見つかりませんでした。", vbExclamation
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
