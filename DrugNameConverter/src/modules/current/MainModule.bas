Attribute VB_Name = "MainModule"
Option Explicit

' メインの処理関数：医薬品名の一致度に基づいて転記
Public Sub MainProcess()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets(1)
    Set ws2 = ThisWorkbook.Worksheets(2)
    Set ws3 = ThisWorkbook.Worksheets(3)
    
    Dim lastRow1 As Long, lastRow2 As Long
    lastRow1 = ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row
    
    Const MATCH_THRESHOLD As Double = 80 ' 一致率のしきい値（80%）
    
    Dim i As Long, j As Long
    For i = 2 To lastRow1
        Dim sourceStr As String
        sourceStr = ws1.Cells(i, "B").Value
        
        If Len(sourceStr) > 0 Then
            Dim maxMatchRate As Double
            Dim bestMatchIndex As Long
            maxMatchRate = 0
            bestMatchIndex = 0
            
            For j = 2 To lastRow2
                Dim targetStr As String
                targetStr = ws2.Cells(j, "B").Value
                
                Dim currentMatchRate As Double
                currentMatchRate = CompareDrugStringsWithRate(sourceStr, targetStr)
                
                If currentMatchRate > maxMatchRate Then
                    maxMatchRate = currentMatchRate
                    bestMatchIndex = j
                End If
            Next j
            
            ' 結果の出力
            If maxMatchRate >= MATCH_THRESHOLD Then
                ws1.Cells(i, "C").Value = ws2.Cells(bestMatchIndex, "B").Value
                ws1.Cells(i, "D").Value = maxMatchRate & "%"
                
                ' 一致した各要素の詳細を出力（デバッグ用）
                Dim sourceParts As DrugNameParts
                Dim targetParts As DrugNameParts
                sourceParts = ParseDrugString(sourceStr)
                targetParts = ParseDrugString(ws2.Cells(bestMatchIndex, "B").Value)
                
                ws1.Cells(i, "E").Value = "基本名:" & sourceParts.BaseName & _
                                         " 剤形:" & sourceParts.formType & _
                                         " 規格:" & sourceParts.strength & _
                                         " メーカー:" & sourceParts.maker
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "処理が完了しました。"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description
End Sub

' 医薬品名の検索と転記関数
Public Sub SearchAndTransferDrugData()
    On Error GoTo ErrorHandler
    
    '画面更新を一時停止してパフォーマンス向上
    Application.ScreenUpdating = False
    
    'ワークシートの設定
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets(1)
    Set ws2 = ThisWorkbook.Worksheets(2)
    Set ws3 = ThisWorkbook.Worksheets(3)
    
    '最終行の取得
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row
    lastRow3 = ws3.Cells(ws3.Rows.Count, "F").End(xlUp).Row
    
    Dim i As Long
    Dim inputValue As Variant
    
    '各行で、列の値を処理
    For i = 2 To lastRow1  'ヘッダーをスキップ
        inputValue = ws1.Cells(i, "A").Value
        
        '入力値を処理する関数を呼び出し
        ProcessInputValue inputValue, ws1, ws2, ws3, i, lastRow2, lastRow3
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "処理が完了しました。"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description
End Sub

' 入力値を処理する関数
Private Sub ProcessInputValue(ByVal inputValue As Variant, _
                            ByRef ws1 As Worksheet, _
                            ByRef ws2 As Worksheet, _
                            ByRef ws3 As Worksheet, _
                            ByVal currentRow As Long, _
                            ByVal lastRow2 As Long, _
                            ByVal lastRow3 As Long)
                            
    Dim drugNameFromSheet3 As String
    Dim drugNameFromSheet2 As String
    Dim packageType As String
    Dim j As Long, k As Long
    
    'Sheet3から薬品名を検索
    For k = 2 To lastRow3
        drugNameFromSheet3 = ws3.Cells(k, "F").Value
        If InStr(1, inputValue, drugNameFromSheet3) > 0 Then
            'Sheet2から対応する薬品名を検索
            For j = 2 To lastRow2
                drugNameFromSheet2 = ws2.Cells(j, "B").Value
                If drugNameFromSheet2 = drugNameFromSheet3 Then
                    '包タイプを取得
                    packageType = GetPackageType(inputValue)
                    
                    'データを転記
                    ws1.Cells(currentRow, "B").Value = ws2.Cells(j, "A").Value
                    ws1.Cells(currentRow, "C").Value = packageType
                    Exit For
                End If
            Next j
            Exit For
        End If
    Next k
End Sub

' 一致率計算による医薬品処理関数
Public Sub ProcessDrugNamesWithMatchRate()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets(1)
    Set ws2 = ThisWorkbook.Worksheets(2)
    
    Dim i As Long, j As Long
    Dim lastRow1 As Long, lastRow2 As Long
    Const MATCH_THRESHOLD As Double = 80 ' 一致率のしきい値（80%）
    
    lastRow1 = ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row
    
    For i = 2 To lastRow1
        Dim sourceStr As String
        Dim maxMatchRate As Double
        Dim bestMatchIndex As Long
        
        sourceStr = ws1.Cells(i, "B").Value
        maxMatchRate = 0
        bestMatchIndex = 0
        
        For j = 2 To lastRow2
            Dim targetStr As String
            Dim currentMatchRate As Double
            
            targetStr = ws2.Cells(j, "B").Value
            currentMatchRate = CompareDrugStringsWithRate(sourceStr, targetStr)
            
            If currentMatchRate > maxMatchRate Then
                maxMatchRate = currentMatchRate
                bestMatchIndex = j
            End If
        Next j
        
        ' しきい値以上の一致率があった場合のみ転記
        If maxMatchRate >= MATCH_THRESHOLD Then
            ws1.Cells(i, "C").Value = ws2.Cells(bestMatchIndex, "B").Value
            ws1.Cells(i, "D").Value = maxMatchRate & "%"
        End If
    Next i
    
    MsgBox "処理が完了しました。"
End Sub

' 設定シートの包形態を考慮して医薬品比較と転記を行う
Public Sub CompareAndTransferDrugNamesByPackage()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' ワークシートの設定
    Dim wsSettings As Worksheet, wsTarget As Worksheet
    Set wsSettings = ThisWorkbook.Worksheets(1) ' 設定シート
    Set wsTarget = ThisWorkbook.Worksheets(2)   ' 比較対象のシート
    
    ' 包装形態を医薬品名から直接抽出するように変更
    Dim packageType As String
    
    ' 最終行を取得
    Dim lastRowSettings As Long, lastRowTarget As Long
    lastRowSettings = wsSettings.Cells(wsSettings.Rows.Count, "B").End(xlUp).Row
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "B").End(xlUp).Row
    
    ' 検索対象と比較対象の医薬品名を配列に格納
    Dim searchDrugs() As String
    Dim targetDrugs() As String
    Dim i As Long, j As Long
    
    ' 検索医薬品用の配列を初期化
    ReDim searchDrugs(1 To lastRowSettings - 1) ' ヘッダー行を除外
    For i = 2 To lastRowSettings
        searchDrugs(i - 1) = wsSettings.Cells(i, "B").Value
    Next i
    
    ' 比較対象用の配列を初期化
    ReDim targetDrugs(1 To lastRowTarget - 1) ' ヘッダー行を除外
    For i = 2 To lastRowTarget
        targetDrugs(i - 1) = wsTarget.Cells(i, "B").Value
    Next i
    
    ' 各検索医薬品に対して比較処理
    For i = 2 To lastRowSettings
        Dim searchDrug As String
        searchDrug = wsSettings.Cells(i, "B").Value
        
        If Len(searchDrug) > 0 Then
            ' 医薬品名から直接包装形態を抽出
            packageType = PackageTypeExtractor.ExtractPackageTypeFromDrugName(searchDrug)
            
            Dim bestMatch As String
            bestMatch = FindBestMatchWithPackage(searchDrug, targetDrugs, packageType)
            
            If Len(bestMatch) > 0 Then
                ' 一致した医薬品名をC列に転記
                wsSettings.Cells(i, "C").Value = bestMatch
            Else
                ' 一致しなかった場合は空欄にする
                wsSettings.Cells(i, "C").Value = ""
            End If
        End If
    Next i
    
CleanExit:
    Application.ScreenUpdating = True
    MsgBox "処理が完了しました。", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

' 医薬品名の成分、規格、単位の一致度を計算
Public Function CalculateMatchScore(ByRef searchParts As DrugNameParts, ByRef targetParts As DrugNameParts) As Double
    Dim score As Double
    Dim totalWeight As Double
    
    score = 0
    totalWeight = 0
    
    ' 基本名の比較（重み: 50%）
    If StrComp(searchParts.BaseName, targetParts.BaseName, vbTextCompare) = 0 Then
        score = score + 50
    End If
    totalWeight = totalWeight + 50
    
    ' 剤形の比較（重み: 20%）
    If StrComp(searchParts.formType, targetParts.formType, vbTextCompare) = 0 Then
        score = score + 20
    End If
    totalWeight = totalWeight + 20
    
    ' 規格の比較（重み: 30%）
    If CompareStrength(searchParts.strength, targetParts.strength) Then
        score = score + 30
    End If
    totalWeight = totalWeight + 30
    
    ' スコアの正規化（百分率）
    If totalWeight > 0 Then
        CalculateMatchScore = (score / totalWeight) * 100
    Else
        CalculateMatchScore = 0
    End If
End Function

' 包形態を考慮して最適な医薬品名の一致を見つける関数
Private Function FindBestMatchWithPackage(ByVal searchDrug As String, ByRef targetDrugs() As String, ByVal requiredPackage As String) As String
    Dim i As Long
    Dim bestMatchScore As Double
    Dim bestMatchIndex As Long
    Dim searchParts As DrugNameParts
    
    ' 検索対象の医薬品名を分解
    searchParts = ParseDrugString(searchDrug)
    bestMatchScore = 0
    bestMatchIndex = -1
    
    ' 各比較対象に対してスコアを計算
    For i = LBound(targetDrugs) To UBound(targetDrugs)
        Dim targetParts As DrugNameParts
        Dim currentScore As Double
        Dim hasRequiredPackage As Boolean
        
        ' 比較対象の医薬品名を分解
        targetParts = ParseDrugString(targetDrugs(i))
        
        ' 包形態の確認
        hasRequiredPackage = (InStr(1, targetParts.Package, requiredPackage, vbTextCompare) > 0)
        
        If hasRequiredPackage Then
            ' 基本名、規格、単位の一致度を計算
            currentScore = CalculateMatchScore(searchParts, targetParts)
            
            If currentScore > bestMatchScore Then
                bestMatchScore = currentScore
                bestMatchIndex = i
            End If
        End If
    Next i
    
    ' 一定以上のスコアがある場合のみ結果を返す
    If bestMatchScore >= 70 And bestMatchIndex >= 0 Then ' 70%以上の一致率
        FindBestMatchWithPackage = targetDrugs(bestMatchIndex)
    Else
        FindBestMatchWithPackage = ""
    End If
End Function

' 7行目以降の医薬品名を比較と転記を行う関数
Public Sub ProcessFromRow7()
    On Error GoTo ErrorHandler
    
    ' 初期設定
    Application.ScreenUpdating = False
    
    ' ワークシート参照の取得
    Dim settingsSheet As Worksheet, targetSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1) ' 設定シート
    Set targetSheet = ThisWorkbook.Worksheets(2)   ' 比較対象のシート
    
    ' 医薬品名から直接包装形態を抽出するように変更
    ' PackageTypeExtractorモジュールを初期化
    PackageTypeExtractor.InitializePackageMappings
    
    ' 最終行の取得
    Dim lastRowSettings As Long, lastRowTarget As Long
    lastRowSettings = settingsSheet.Cells(settingsSheet.Rows.Count, "B").End(xlUp).Row
    lastRowTarget = targetSheet.Cells(targetSheet.Rows.Count, "B").End(xlUp).Row
    
    ' 比較対象医薬品を配列に格納
    Dim targetDrugs() As String
    ReDim targetDrugs(1 To lastRowTarget - 1)
    
    Dim i As Long
    For i = 2 To lastRowTarget
        targetDrugs(i - 1) = targetSheet.Cells(i, "B").Value
    Next i
    
    ' 医薬品名の比較と転記（7行目から開始）
    Dim searchDrug As String, bestMatch As String
    Dim processedCount As Long, skippedCount As Long
    processedCount = 0
    skippedCount = 0
    
    For i = 7 To lastRowSettings ' 処理を7行目以降から開始
        searchDrug = settingsSheet.Cells(i, "B").Value
        
        If Len(searchDrug) > 0 Then
            ' 医薬品名から直接包装形態を抽出
            Dim packageType As String
            packageType = PackageTypeExtractor.ExtractPackageTypeFromDrugName(searchDrug)
            
            ' 最適な一致を検索
            bestMatch = FindBestMatchingDrug(searchDrug, targetDrugs, packageType)
            
            ' 一致する結果があれば転記、なければスキップ
            If Len(bestMatch) > 0 Then
                settingsSheet.Cells(i, "C").Value = bestMatch
                processedCount = processedCount + 1
            Else
                ' 一致しない場合は何もしない（空文字で上書きしない）
                skippedCount = skippedCount + 1
            End If
        End If
    Next i
    
CleanExit:
    Application.ScreenUpdating = True
    MsgBox "処理が完了しました。" & vbCrLf & _
           processedCount & "件の医薬品名が一致しました。" & vbCrLf & _
           skippedCount & "件の医薬品名は一致するものが見つかりませんでした。", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' 最適一致する医薬品名を見つける関数
Private Function FindBestMatchingDrug(ByVal searchDrug As String, ByRef targetDrugs() As String, ByVal packageType As String) As String
    Dim i As Long
    Dim bestMatchIndex As Long, bestMatchScore As Long, currentScore As Long
    
    bestMatchIndex = -1
    bestMatchScore = 0
    
    ' 検索対象をキーワードに分解
    Dim keywords As Variant
    keywords = ExtractKeywords(searchDrug)
    
    ' 包形態の特殊処理
    Dim skipPackageCheck As Boolean
    skipPackageCheck = (packageType = "/未定義/" Or packageType = "/その他(なし)/")
    
    ' 各比較対象に対して処理
    For i = LBound(targetDrugs) To UBound(targetDrugs)
        If Len(targetDrugs(i)) > 0 Then
            ' 包形態チェック
            Dim matchesPackage As Boolean
            
            If skipPackageCheck Then
                ' 未定義またはその他の場合は包形態チェックをスキップ
                matchesPackage = True
            Else
                ' 包形態が一致するか確認
                Dim targetPackageType As String
                targetPackageType = PackageTypeExtractor.ExtractPackageTypeFromDrugName(targetDrugs(i))
                matchesPackage = (targetPackageType = packageType)
            End If
            
            If matchesPackage Then
                ' キーワード一致率を計算
                currentScore = CalcMatchScore(keywords, targetDrugs(i))
                
                ' より高いスコアを記録
                If currentScore > bestMatchScore Then
                    bestMatchScore = currentScore
                    bestMatchIndex = i
                End If
            End If
        End If
    Next i
    
    ' 結果を返す（閾値以上のスコアの場合のみ）
    If bestMatchScore >= 50 And bestMatchIndex >= 0 Then
        FindBestMatchingDrug = targetDrugs(bestMatchIndex)
    Else
        FindBestMatchingDrug = ""
    End If
End Function

' 医薬品名からキーワードを抽出する関数
Private Function ExtractKeywords(ByVal drugName As String) As Variant
    ' 全角スペースを半角に変換
    drugName = Replace(drugName, "　", " ")
    
    ' スペースで分割して配列に格納
    Dim words As Variant, result() As String
    Dim i As Long, validCount As Long
    
    words = Split(drugName, " ")
    ReDim result(UBound(words))
    validCount = 0
    
    ' 空でない要素のみ取得
    For i = 0 To UBound(words)
        If Trim(words(i)) <> "" Then
            result(validCount) = LCase(Trim(words(i)))
            validCount = validCount + 1
        End If
    Next i
    
    ' 結果が空の場合の処理
    If validCount = 0 Then
        ReDim result(0)
        result(0) = LCase(Trim(drugName))
        validCount = 1
    End If
    
    ReDim Preserve result(validCount - 1)
    ExtractKeywords = result
End Function

' キーワードの一致率を計算する関数
Private Function CalcMatchScore(ByRef keywords As Variant, ByVal targetDrug As String) As Long
    Dim i As Long, matchCount As Long
    Dim lowerTargetDrug As String
    
    lowerTargetDrug = LCase(targetDrug)
    matchCount = 0
    
    ' 各キーワードが含まれているかチェック
    For i = 0 To UBound(keywords)
        If InStr(1, lowerTargetDrug, keywords(i), vbTextCompare) > 0 Then
            matchCount = matchCount + 1
        End If
    Next i
    
    ' 一致率を計算（百分率）
    If UBound(keywords) >= 0 Then
        CalcMatchScore = (matchCount * 100) / (UBound(keywords) + 1)
    Else
        CalcMatchScore = 0
    End If
End Function

' 包形態が一致するかチェックする関数（CreateObjectを使わないバージョン）
Private Function CheckPackage(ByVal drugName As String, ByVal packageType As String) As Boolean
    ' PackageTypeExtractorモジュールを使用して包装形態を抽出
    Dim extractedPackageType As String
    extractedPackageType = PackageTypeExtractor.ExtractPackageTypeFromDrugName(drugName)
    
    ' 抽出した包装形態と指定された包装形態を比較
    CheckPackage = (extractedPackageType = packageType)
End Function

' GTIN-14コードから医薬品情報を処理するメイン関数
Public Sub ProcessGS1DrugCode()
    On Error GoTo ErrorHandler
    
    ' GTIN-14コードの入力を求める
    Dim gtin14Code As String
    gtin14Code = InputBox("GTIN-14の14桁コードを入力してください:", "医薬品コード処理")
    
    If Len(gtin14Code) = 0 Then
        Exit Sub
    End If
    
    ' 14桁であることを確認
    If Len(gtin14Code) <> 14 Or Not IsNumeric(gtin14Code) Then
        MsgBox "GTIN-14コードは14桁の数字である必要があります。", vbExclamation
        Exit Sub
    End If
    
    ' GTIN-14コードを処理
    GS1CodeProcessor.ProcessGS1CodeAndUpdateSettings gtin14Code
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

' GTIN-14コードから医薬品情報を配列で取得して表示するデモ関数
Public Sub DemoDisplayDrugInfoFromGS1()
    On Error GoTo ErrorHandler
    
    ' GTIN-14コードの入力を求める
    Dim gtin14Code As String
    gtin14Code = InputBox("GTIN-14の14桁コードを入力してください:", "医薬品情報表示")
    
    If Len(gtin14Code) = 0 Then
        Exit Sub
    End If
    
    ' 14桁であることを確認
    If Len(gtin14Code) <> 14 Or Not IsNumeric(gtin14Code) Then
        MsgBox "GTIN-14コードは14桁の数字である必要があります。", vbExclamation
        Exit Sub
    End If
    
    ' 医薬品情報を配列として取得
    Dim drugInfoArray As Variant
    drugInfoArray = GS1CodeProcessor.GetDrugInfoAsArray(gtin14Code)
    
    ' 結果を表示
    Dim resultMsg As String
    resultMsg = "医薬品情報:" & vbCrLf & _
               "成分名: " & drugInfoArray(1) & vbCrLf & _
               "剤形: " & drugInfoArray(2) & vbCrLf & _
               "用量規格: " & drugInfoArray(3) & vbCrLf & _
               "メーカー: " & drugInfoArray(4) & vbCrLf & _
               "包装規格: " & drugInfoArray(5) & vbCrLf & _
               "包装形態: " & drugInfoArray(6) & vbCrLf & _
               "追加情報: " & drugInfoArray(7) & vbCrLf & _
               "医薬品名: " & drugInfoArray(8) & vbCrLf & _
               "パッケージ・インジケーター: " & Left(gtin14Code, 1) & " (" & GetPackageIndicatorDescription(Left(gtin14Code, 1)) & ")"
    
    MsgBox resultMsg, vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

' パッケージ・インジケーターの説明を取得する関数
Private Function GetPackageIndicatorDescription(ByVal indicator As String) As String
    Select Case indicator
        Case "0"
            GetPackageIndicatorDescription = "調剤包装単位"
        Case "1"
            GetPackageIndicatorDescription = "販売包装単位"
        Case "2"
            GetPackageIndicatorDescription = "元梱包装単位"
        Case Else
            GetPackageIndicatorDescription = "不明"
    End Select
End Function

' メインメニューを表示する関数
Public Sub ShowMainMenu()
    Dim choice As VbMsgBoxResult
    
    choice = MsgBox("薬局在庫管理システム - 機能選択" & vbCrLf & vbCrLf & _
                   "「はい」：医薬品名比較機能" & vbCrLf & _
                   "「いいえ」：棚番一括更新機能" & vbCrLf & _
                   "「キャンセル」：GTIN-14コード処理" & vbCrLf & vbCrLf & _
                   "※CSVインポート機能はマクロ「ImportCSVToSheet2.ImportCSVToSheet2」を実行", _
                   vbYesNoCancel + vbQuestion, "メイン機能選択")
    
    Select Case choice
        Case vbYes
            ' 医薬品名比較機能
            RunDrugNameComparison
            
        Case vbNo
            ' 棚番一括更新機能
            ShelfManager.Main
            
        Case vbCancel
            ' GTIN-14コード処理
            ProcessGS1DrugCode
            
    End Select
End Sub

' アプリケーション起動時の初期化関数
Public Sub InitializeApplication()
    ' ユーザーフォームを初期化（必要な場合）
    If Not IsFormLoaded("ShelfNameForm") Then
        Load ShelfNameForm
    End If
    
    ' その他の初期化処理
    
    ' メインメニューを表示
    ShowMainMenu
End Sub

' フォームが既にロードされているか確認
Private Function IsFormLoaded(formName As String) As Boolean
    Dim i As Integer
    For i = 0 To VBA.UserForms.Count - 1
        If VBA.UserForms(i).Name = formName Then
            IsFormLoaded = True
            Exit Function
        End If
    Next i
    IsFormLoaded = False
End Function





