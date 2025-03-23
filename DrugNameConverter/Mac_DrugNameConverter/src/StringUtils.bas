Option Explicit

' ==============================================
' 文字列処理ユーティリティ関数
' ==============================================

' 「」で囲まれたテキストを抽出する関数（正規表現を使わないバージョン）
Public Function ExtractBetweenQuotes(ByVal text As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(1, text, "「")
    If startPos > 0 Then
        endPos = InStr(startPos + 1, text, "」")
        If endPos > startPos Then
            ExtractBetweenQuotes = Mid(text, startPos + 1, endPos - startPos - 1)
        Else
            ExtractBetweenQuotes = ""
        End If
    Else
        ExtractBetweenQuotes = ""
    End If
End Function

' レーベンシュタイン距離を計算する関数（2つの文字列間の編集距離）
Public Function LevenshteinDistance(ByVal s1 As String, ByVal s2 As String) As Long
    On Error GoTo ErrorHandler
    
    Dim i As Long, j As Long
    Dim cost As Long
    
    ' 文字列の長さを取得
    Dim len1 As Long, len2 As Long
    len1 = Len(s1)
    len2 = Len(s2)
    
    ' 距離行列を初期化
    Dim d() As Long
    ReDim d(0 To len1, 0 To len2)
    
    ' ベースケースを初期化
    For i = 0 To len1
        d(i, 0) = i
    Next i
    
    For j = 0 To len2
        d(0, j) = j
    Next j
    
    ' 距離行列を埋める
    For i = 1 To len1
        For j = 1 To len2
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            
            d(i, j) = Application.WorksheetFunction.Min( _
                d(i - 1, j) + 1, _      ' 削除
                d(i, j - 1) + 1, _      ' 挿入
                d(i - 1, j - 1) + cost) ' 置換
        Next j
    Next i
    
    ' 結果を返す
    LevenshteinDistance = d(len1, len2)
    Exit Function
    
ErrorHandler:
    Debug.Print "LevenshteinDistance エラー: " & Err.Number & " - " & Err.Description
    LevenshteinDistance = 0
End Function

' 2つの文字列の類似度を計算する関数
Public Function GetSimilarity(ByVal string1 As String, ByVal string2 As String) As Double
    On Error GoTo ErrorHandler
    
    ' 両方の文字列が空の場合は完全一致とみなす
    If Len(string1) = 0 And Len(string2) = 0 Then
        GetSimilarity = 1
        Exit Function
    End If
    
    ' どちらかの文字列が空の場合は類似度0
    If Len(string1) = 0 Or Len(string2) = 0 Then
        GetSimilarity = 0
        Exit Function
    End If
    
    ' 大文字小文字を区別しない比較のため小文字に統一
    string1 = LCase(string1)
    string2 = LCase(string2)
    
    ' 完全一致の場合
    If string1 = string2 Then
        GetSimilarity = 1
        Exit Function
    End If
    
    ' レーベンシュタイン距離を計算
    Dim distance As Long
    distance = LevenshteinDistance(string1, string2)
    
    ' 長い方の文字列の長さを基準に正規化
    Dim maxLength As Long
    maxLength = Application.WorksheetFunction.Max(Len(string1), Len(string2))
    
    ' 類似度を計算 (0に近いほど違い、1に近いほど似ている)
    GetSimilarity = 1 - (distance / CDbl(maxLength))
    Exit Function
    
ErrorHandler:
    Debug.Print "GetSimilarity エラー: " & Err.Number & " - " & Err.Description
    GetSimilarity = 0
End Function

' 規格（強度）を抽出する関数（正規表現を使わない版）
Public Function ExtractStrength(ByVal text As String) As String
    Dim i As Long
    Dim numStart As Long
    Dim result As String
    Dim inNumber As Boolean
    Dim units As Variant
    
    units = Array("mg", "g", "ml", "μg")
    inNumber = False
    numStart = 0
    
    For i = 1 To Len(text)
        Dim c As String
        c = Mid(text, i, 1)
        
        If IsNumeric(c) Or c = "." Then
            If Not inNumber Then
                inNumber = True
                numStart = i
            End If
        ElseIf c = " " And inNumber Then
            ' スペースは許容
        Else
            If inNumber Then
                ' 数字の後に単位があるか確認
                Dim j As Long
                Dim found As Boolean
                found = False
                
                For j = 0 To UBound(units)
                    If LCase(Mid(text, i, Len(units(j)))) = LCase(units(j)) Then
                        result = Mid(text, numStart, i - numStart + Len(units(j)))
                        found = True
                        Exit For
                    End If
                Next j
                
                If found Then
                    ExtractStrength = result
                    Exit Function
                End If
                
                inNumber = False
            End If
        End If
    Next i
    
    ExtractStrength = ""
End Function

' 数値と単位を分離する関数（正規表現を使わないバージョン）
Public Sub ExtractNumberAndUnit(ByVal str As String, ByRef num As Double, ByRef unit As String)
    Dim i As Long
    Dim numStr As String
    Dim unitStr As String
    Dim numStart As Long
    Dim inNumber As Boolean
    
    inNumber = False
    numStart = 0
    numStr = ""
    unitStr = ""
    
    For i = 1 To Len(str)
        Dim c As String
        c = Mid(str, i, 1)
        
        If IsNumeric(c) Or c = "." Then
            If Not inNumber Then
                inNumber = True
                numStart = i
            End If
        ElseIf c = " " And inNumber Then
            ' スペースは数字と見なす
        Else
            If inNumber Then
                numStr = Mid(str, numStart, i - numStart)
                unitStr = Mid(str, i)
                Exit For
            End If
        End If
    Next i
    
    ' 単位から不要な文字を削除
    unitStr = Trim(unitStr)
    
    ' 単位の標準化
    If LCase(Left(unitStr, 2)) = "mg" Then
        unitStr = "mg"
    ElseIf LCase(Left(unitStr, 1)) = "g" Then
        unitStr = "g"
    ElseIf LCase(Left(unitStr, 2)) = "ml" Then
        unitStr = "ml"
    ElseIf LCase(Left(unitStr, 2)) = "μg" Then
        unitStr = "μg"
    End If
    
    ' 結果を設定
    If Len(numStr) > 0 Then
        On Error Resume Next
        num = CDbl(numStr)
        If Err.Number <> 0 Then
            num = 0
        End If
        On Error GoTo 0
        unit = LCase(unitStr)
    Else
        num = 0
        unit = ""
    End If
End Sub

' 規格（強度）を比較する関数
Public Function CompareStrength(ByVal str1 As String, ByVal str2 As String) As Boolean
    ' 数値と単位を分離して比較
    Dim num1 As Double, num2 As Double
    Dim unit1 As String, unit2 As String
    
    ' 数値と単位を抽出
    ExtractNumberAndUnit str1, num1, unit1
    ExtractNumberAndUnit str2, num2, unit2
    
    ' 数値と単位が両方一致する場合のみTrue
    CompareStrength = (num1 = num2) And (StrComp(unit1, unit2, vbTextCompare) = 0)
End Function

' 薬品名から特定のキーワードを抽出する関数
Public Function ExtractKeywords(ByVal drugName As String) As Variant
    ' キーワードを格納する配列
    Dim keywords() As String
    ReDim keywords(0 To 9)  ' 最大10個のキーワードを想定
    Dim keywordCount As Long
    keywordCount = 0
    
    ' 空文字チェック
    If Len(drugName) = 0 Then
        ReDim keywords(0 To 0)
        keywords(0) = ""
        ExtractKeywords = keywords
        Exit Function
    End If
    
    ' エラーメッセージのチェック（[コード未登録]など）
    If Left(drugName, 1) = "[" And Right(drugName, 1) = "]" Then
        ReDim keywords(0 To 0)
        keywords(0) = drugName
        ExtractKeywords = keywords
        Exit Function
    End If
    
    ' 数字と単位のパターンを抽出（mg、g、mLなど）
    Dim strengthPattern As String
    strengthPattern = ExtractStrength(drugName)
    
    If Len(strengthPattern) > 0 Then
        keywords(keywordCount) = strengthPattern
        keywordCount = keywordCount + 1
    End If
    
    ' 一般的な薬品名の特徴的な単語を抽出
    Dim commonWords As Variant
    commonWords = Array("錠", "カプセル", "顆粒", "散", "シロップ", "注射", "軟膏", "点眼", "坐剤", "貼付")
    
    Dim i As Long
    For i = LBound(commonWords) To UBound(commonWords)
        If InStr(1, drugName, commonWords(i), vbTextCompare) > 0 Then
            keywords(keywordCount) = commonWords(i)
            keywordCount = keywordCount + 1
            
            ' 最大キーワード数に達したらループを抜ける
            If keywordCount >= 10 Then
                Exit For
            End If
        End If
    Next i
    
    ' 結果を適切なサイズの配列に調整
    If keywordCount > 0 Then
        ReDim Preserve keywords(0 To keywordCount - 1)
    Else
        ReDim keywords(0 To 0)
        keywords(0) = ""
    End If
    
    ExtractKeywords = keywords
End Function

' 文字列の中から最も重要と思われる単語を抽出する関数
Public Function ExtractImportantWords(ByVal text As String, Optional ByVal maxWords As Long = 3) As Variant
    Dim result() As String
    ReDim result(0 To maxWords - 1)
    Dim resultCount As Long
    resultCount = 0
    
    ' 空文字チェック
    If Len(text) = 0 Then
        ReDim result(0 To 0)
        result(0) = ""
        ExtractImportantWords = result
        Exit Function
    End If
    
    ' 記号や余分なスペースを除去
    Dim cleanText As String
    cleanText = text
    
    ' 括弧内の内容を除去
    cleanText = RemoveBracketsContent(cleanText)
    
    ' 単語に分割（日本語なので文字ごとに区切る）
    Dim words() As String
    Dim wordCount As Long
    ReDim words(0 To Len(cleanText) - 1)
    wordCount = 0
    
    Dim i As Long
    Dim currentWord As String
    currentWord = ""
    
    For i = 1 To Len(cleanText)
        Dim c As String
        c = Mid(cleanText, i, 1)
        
        ' 記号や空白の場合は区切り文字として扱う
        If c = " " Or c = "　" Or c = "," Or c = "." Or c = "、" Or c = "。" Then
            If Len(currentWord) > 0 Then
                words(wordCount) = currentWord
                wordCount = wordCount + 1
                currentWord = ""
            End If
        Else
            currentWord = currentWord & c
        End If
    Next i
    
    ' 最後の単語を追加
    If Len(currentWord) > 0 Then
        words(wordCount) = currentWord
        wordCount = wordCount + 1
    End If
    
    ' 配列のサイズを調整
    If wordCount > 0 Then
        ReDim Preserve words(0 To wordCount - 1)
    Else
        ReDim words(0 To 0)
        words(0) = ""
    End If
    
    ' 結果を最大単語数に制限
    Dim actualMaxWords As Long
    actualMaxWords = Application.WorksheetFunction.Min(maxWords, wordCount)
    
    If actualMaxWords > 0 Then
        ReDim result(0 To actualMaxWords - 1)
        For i = 0 To actualMaxWords - 1
            result(i) = words(i)
        Next i
    Else
        ReDim result(0 To 0)
        result(0) = ""
    End If
    
    ExtractImportantWords = result
End Function

' 括弧内の内容を除去する関数
Private Function RemoveBracketsContent(ByVal text As String) As String
    Dim result As String
    Dim pos1 As Long, pos2 As Long
    result = text
    
    ' すべての括弧とその内容を繰り返し削除
    Do
        ' 丸括弧を検索
        pos1 = InStr(1, result, "(")
        If pos1 > 0 Then
            pos2 = InStr(pos1, result, ")")
            If pos2 > 0 Then
                result = Left(result, pos1 - 1) & Mid(result, pos2 + 1)
                ' 処理を継続
                GoTo ContinueLoop
            End If
        End If
        
        ' 角括弧を検索
        pos1 = InStr(1, result, "[")
        If pos1 > 0 Then
            pos2 = InStr(pos1, result, "]")
            If pos2 > 0 Then
                result = Left(result, pos1 - 1) & Mid(result, pos2 + 1)
                ' 処理を継続
                GoTo ContinueLoop
            End If
        End If
        
        ' 括弧が見つからなければループを抜ける
        Exit Do
        
ContinueLoop:
    Loop
    
    RemoveBracketsContent = Trim(result)
End Function

' B4セルに包装形態の選択肢をドロップダウンリストとして設定する関数
Public Sub SetupPackageTypeDropdown()
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' B4セルにドロップダウンリストを設定
    With settingsSheet.Range("B4").Validation
        .Delete ' 既存の入力規則を削除
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="(未定義),その他(なし),包装小,調剤用,PTP,分包,バラ,SP,PTP(患者用)"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "包装形態の選択"
        .ErrorTitle = "無効な選択"
        .InputMessage = "リストから包装形態を選択してください"
        .ErrorMessage = "リストから有効な包装形態を選択してください"
    End With
    
    ' B4セルの書式設定
    With settingsSheet.Range("B4")
        .Value = "PTP" ' デフォルト値を設定
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242) ' 薄い青色の背景
    End With
    
    ' A4セルにラベルを設定
    With settingsSheet.Range("A4")
        .Value = "包装形態:"
        .Font.Bold = True
    End With
    
    ' B3セルにタイトルを設定
    With settingsSheet.Range("A1:C1")
        .Merge
        .Value = "医薬品名比較ツール"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(180, 198, 231) ' 青色の背景
    End With
    
    ' 列ヘッダーを設定
    settingsSheet.Range("A6").Value = "No."
    settingsSheet.Range("B6").Value = "検索医薬品名"
    settingsSheet.Range("C6").Value = "一致医薬品名"
    
    With settingsSheet.Range("A6:C6")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247) ' 薄い青色の背景
    End With
    
    ' 列幅を調整
    settingsSheet.Columns("A").ColumnWidth = 5
    settingsSheet.Columns("B").ColumnWidth = 30
    settingsSheet.Columns("C").ColumnWidth = 40
    
    ' 行番号を設定（7行目から30行目まで）
    Dim i As Long
    For i = 7 To 30
        settingsSheet.Cells(i, "A").Value = i - 6
    Next i
    
    MsgBox "包装形態のドロップダウンリストを設定しました。", vbInformation
End Sub


