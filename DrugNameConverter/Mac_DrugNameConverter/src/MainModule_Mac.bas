Option Explicit

' 医薬品マスターシートの設定
Public Const MASTER_SHEET_NAME As String = "薬品マスター"
Public Const DRUG_CODE_COLUMN As Integer = 1  ' A列
Public Const DRUG_NAME_COLUMN As Integer = 2  ' B列
Public Const FIRST_DATA_ROW As Integer = 2    ' 2行目からデータ開始

' ============================================================
' 医薬品情報検索・操作関数
' ============================================================

' 医薬品コードを14桁に整形する関数
Public Function FormatDrugCode(ByVal drugCode As String) As String
    On Error GoTo ErrorHandler
    
    ' 空文字や数値以外の文字を除去
    Dim cleanCode As String
    Dim i As Long
    
    cleanCode = ""
    For i = 1 To Len(drugCode)
        If IsNumeric(Mid(drugCode, i, 1)) Then
            cleanCode = cleanCode & Mid(drugCode, i, 1)
        End If
    Next i
    
    ' 14桁に調整
    If Len(cleanCode) > 14 Then
        ' 14桁を超える場合は左から14桁を使用
        FormatDrugCode = Left(cleanCode, 14)
    ElseIf Len(cleanCode) < 14 Then
        ' 14桁未満の場合は右寄せでゼロ埋め
        FormatDrugCode = String(14 - Len(cleanCode), "0") & cleanCode
    Else
        ' ちょうど14桁の場合はそのまま
        FormatDrugCode = cleanCode
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "FormatDrugCode エラー: " & Err.Number & " - " & Err.Description
    FormatDrugCode = drugCode ' エラー時は元の値を返す
End Function

' 医薬品コードから医薬品名を検索する関数（書式比較に対応）
Public Function FindDrugNameByCode(ByVal drugCode As String) As String
    On Error GoTo ErrorHandler
    
    Debug.Print "--- FindDrugNameByCode 開始 ---"
    Debug.Print "検索コード: " & drugCode
    
    ' 引数チェック
    If Len(drugCode) = 0 Then
        Debug.Print "空の医薬品コードが指定されました"
        FindDrugNameByCode = ""
        Exit Function
    End If
    
    ' 薬品マスターシートの存在確認
    Dim masterSheet As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' シートが存在するか確認
    Dim sheetExists As Boolean
    sheetExists = False
    
    Debug.Print "マスターシート検索中..."
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Debug.Print " - シート確認: " & ws.Name
        If ws.Name = MASTER_SHEET_NAME Then
            sheetExists = True
            Set masterSheet = ws
            Debug.Print " -> マスターシート '" & ws.Name & "' が見つかりました"
            Exit For
        End If
    Next ws
    
    If Not sheetExists Then
        ' デバッグ情報を追加
        Debug.Print "薬品マスターシート '" & MASTER_SHEET_NAME & "' が見つかりません"
        Debug.Print "利用可能なシート:"
        For Each ws In wb.Worksheets
            Debug.Print " - " & ws.Name
        Next ws
        
        FindDrugNameByCode = "[マスターシートなし]"
        Exit Function
    End If
    
    ' マスターシートが存在する場合
    Debug.Print "薬品マスターシート '" & MASTER_SHEET_NAME & "' を使用します"
    
    ' 医薬品コードを整形
    Dim formattedCode As String
    formattedCode = FormatDrugCode(drugCode)
    Debug.Print "整形後のコード: " & formattedCode
    
    ' 検索範囲の設定
    Dim lastRow As Long
    lastRow = masterSheet.Cells(masterSheet.Rows.Count, DRUG_CODE_COLUMN).End(xlUp).Row
    
    Debug.Print "マスターシートの最終行: " & lastRow
    
    ' データが存在するか確認
    If lastRow < FIRST_DATA_ROW Then
        Debug.Print "薬品マスターシートにデータが存在しません"
        FindDrugNameByCode = "[データなし]"
        Exit Function
    End If
    
    Debug.Print "薬品マスターシートの検索範囲: " & FIRST_DATA_ROW & "行目〜" & lastRow & "行目"
    
    ' マスターシートのデータをすべて表示（デバッグ用）
    Debug.Print "マスターシートの内容:"
    Dim i As Long
    For i = FIRST_DATA_ROW To lastRow
        Debug.Print " - 行" & i & ": コード[" & masterSheet.Cells(i, DRUG_CODE_COLUMN).Value & _
                  "] 薬品名[" & masterSheet.Cells(i, DRUG_NAME_COLUMN).Value & "]"
    Next i
    
    ' 複数の書式で比較する検索
    Debug.Print "複数書式での検索を実行..."
    
    ' 1. 完全一致検索
    For i = FIRST_DATA_ROW To lastRow
        If CStr(masterSheet.Cells(i, DRUG_CODE_COLUMN).Value) = formattedCode Then
            FindDrugNameByCode = masterSheet.Cells(i, DRUG_NAME_COLUMN).Value
            Debug.Print "完全一致で見つかりました: " & FindDrugNameByCode
            Exit Function
        End If
    Next i
    
    ' 2. 先頭の0を除いた数値として比較
    Dim numericCode As String
    numericCode = CStr(CLng(Val(formattedCode)))
    Debug.Print "数値変換後のコード: " & numericCode
    
    For i = FIRST_DATA_ROW To lastRow
        Dim masterCodeNumeric As String
        masterCodeNumeric = CStr(CLng(Val(CStr(masterSheet.Cells(i, DRUG_CODE_COLUMN).Value))))
        
        Debug.Print "比較: " & numericCode & " vs " & masterCodeNumeric
        
        If numericCode = masterCodeNumeric Then
            FindDrugNameByCode = masterSheet.Cells(i, DRUG_NAME_COLUMN).Value
            Debug.Print "数値比較で見つかりました: " & FindDrugNameByCode
            Exit Function
        End If
    Next i
    
    ' 3. 末尾13桁での比較
    Dim code13Digits As String
    If Len(formattedCode) >= 13 Then
        code13Digits = Right(formattedCode, 13)
        Debug.Print "末尾13桁: " & code13Digits
        
        For i = FIRST_DATA_ROW To lastRow
            Dim masterCode As String
            masterCode = CStr(masterSheet.Cells(i, DRUG_CODE_COLUMN).Value)
            
            If Len(masterCode) >= 13 Then
                Dim masterCode13Digits As String
                masterCode13Digits = Right(masterCode, 13)
                
                If code13Digits = masterCode13Digits Then
                    FindDrugNameByCode = masterSheet.Cells(i, DRUG_NAME_COLUMN).Value
                    Debug.Print "末尾13桁比較で見つかりました: " & FindDrugNameByCode
                    Exit Function
                End If
            End If
        Next i
    End If
    
    ' 4. 文字列の末尾部分での比較（先頭の0を除いた部分）
    Debug.Print "文字列末尾での部分比較を実行..."
    For i = FIRST_DATA_ROW To lastRow
        Dim mc As String
        mc = CStr(masterSheet.Cells(i, DRUG_CODE_COLUMN).Value)
        
        ' 末尾部分が一致するか確認
        If Right(formattedCode, Len(mc)) = mc Or Right(mc, Len(numericCode)) = numericCode Then
            FindDrugNameByCode = masterSheet.Cells(i, DRUG_NAME_COLUMN).Value
            Debug.Print "末尾部分比較で見つかりました: " & FindDrugNameByCode
            Exit Function
        End If
    Next i
    
    ' 医薬品コードが見つからない場合
    Debug.Print "コード '" & drugCode & "' は薬品マスターに未登録です"
    FindDrugNameByCode = "[コード未登録]"
    Debug.Print "--- FindDrugNameByCode 終了 ---"
    
    Exit Function
    
ErrorHandler:
    Debug.Print "FindDrugNameByCode エラー: " & Err.Number & " - " & Err.Description
    FindDrugNameByCode = "[エラー発生]"
End Function

' 医薬品コードを元に医薬品名を設定する関数
Public Sub FillDrugNamesByCode()
    On Error GoTo ErrorHandler
    
    ' ワークシート参照の取得
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' 最終行の取得
    Dim lastRow As Long
    lastRow = settingsSheet.Cells(settingsSheet.Rows.Count, "A").End(xlUp).Row
    
    ' 各行の医薬品コードから医薬品名を取得してC列に表示
    Dim i As Long
    For i = 7 To lastRow
        Application.StatusBar = "医薬品名取得中: " & (i - 6) & "/" & (lastRow - 6) & "..."
        DoEvents
        
        Dim drugCode As String
        drugCode = settingsSheet.Cells(i, "A").Value
        
        If Len(drugCode) > 0 Then
            ' 医薬品コードを14桁に整形
            drugCode = FormatDrugCode(drugCode)
            settingsSheet.Cells(i, "A").Value = drugCode
            
            ' 医薬品名を取得してC列に設定
            Dim drugName As String
            drugName = FindDrugNameByCode(drugCode)
            
            settingsSheet.Cells(i, "C").Value = drugName
        End If
    Next i
    
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "医薬品名の設定中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 文字列の類似度を計算する関数
Public Function CalculateSimilarity(ByVal str1 As String, ByVal str2 As String) As Double
    On Error GoTo ErrorHandler
    
    ' 空文字列チェック
    If Len(str1) = 0 Or Len(str2) = 0 Then
        CalculateSimilarity = 0
        Exit Function
    End If
    
    ' 大文字小文字を区別しない比較のため小文字に統一
    str1 = LCase(str1)
    str2 = LCase(str2)
    
    ' 同一文字列の場合は類似度100%
    If str1 = str2 Then
        CalculateSimilarity = 1
        Exit Function
    End If
    
    ' バイグラム類似度（2文字の連続部分文字列の一致率）を計算
    Dim bigrams1 As Collection, bigrams2 As Collection
    Set bigrams1 = New Collection
    Set bigrams2 = New Collection
    
    ' バイグラムの生成
    Dim i As Long
    Dim bigram As String
    
    ' str1のバイグラム
    For i = 1 To Len(str1) - 1
        bigram = Mid(str1, i, 2)
        On Error Resume Next ' 重複エラーをスキップ
        bigrams1.Add bigram, bigram
        On Error GoTo ErrorHandler
    Next i
    
    ' str2のバイグラム
    For i = 1 To Len(str2) - 1
        bigram = Mid(str2, i, 2)
        On Error Resume Next ' 重複エラーをスキップ
        bigrams2.Add bigram, bigram
        On Error GoTo ErrorHandler
    Next i
    
    ' 一致するバイグラムをカウント
    Dim matchCount As Long
    matchCount = 0
    
    On Error Resume Next
    For i = 1 To bigrams1.Count
        Dim key As String
        key = bigrams1(i)
        
        ' バイグラムが存在するか確認
        Dim element As Variant
        On Error Resume Next
        element = bigrams2(key)
        If Err.Number = 0 Then
            matchCount = matchCount + 1
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' 類似度を計算（ダイス係数: 2*一致数 / (集合1のサイズ + 集合2のサイズ)）
    If bigrams1.Count + bigrams2.Count > 0 Then
        CalculateSimilarity = (2 * matchCount) / CDbl(bigrams1.Count + bigrams2.Count)
    Else
        CalculateSimilarity = 0
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "CalculateSimilarity エラー: " & Err.Number & " - " & Err.Description
    CalculateSimilarity = 0
End Function

' メモリ解放関数
Public Sub CollectGarbage()
    Dim tmp As String
    tmp = Space(50000000)  ' 大きな文字列を作成
    tmp = ""  ' 解放
End Sub

' 医薬品名から最も一致する医薬品を検索する関数（段階的に処理）
Public Function FindBestMatchingDrug(ByVal searchTerm As String, ByRef targetDrugs() As String, Optional ByVal packageType As String = "") As String
    On Error GoTo ErrorHandler
    
    ' 引数チェック
    If Len(searchTerm) = 0 Then
        FindBestMatchingDrug = ""
        Exit Function
    End If
    
    ' エラーメッセージの処理
    If Left(searchTerm, 1) = "[" And Right(searchTerm, 1) = "]" Then
        ' エラーメッセージ（[マスターシートなし]など）はそのまま返す
        FindBestMatchingDrug = ""
        Exit Function
    End If
    
    ' 検索対象の配列が空の場合
    If UBound(targetDrugs) < LBound(targetDrugs) Then
        FindBestMatchingDrug = ""
        Exit Function
    End If
    
    ' デバッグ出力
    Debug.Print "FindBestMatchingDrug: 検索開始"
    Debug.Print "検索語: " & searchTerm & ", 包装形態: " & packageType
    
    ' 検索語の薬品名を解析
    Dim searchParts As DrugNameParts
    searchParts = DrugNameParser_Mac.ParseDrugString(searchTerm)
    
    ' 段階1: 医薬品名での完全一致検索（ベースネームの完全一致）
    Debug.Print "段階1: 医薬品名での完全一致検索"
    Dim exactBasenameMatch As String
    exactBasenameMatch = FindExactBasenameMatch(searchParts, targetDrugs)
    
    If Len(exactBasenameMatch) > 0 Then
        Debug.Print "  医薬品名完全一致: " & exactBasenameMatch
        FindBestMatchingDrug = exactBasenameMatch
        Exit Function
    End If
    
    ' 段階2: 医薬品名+規格での検索（ベースネーム+規格の一致）
    Debug.Print "段階2: 医薬品名+規格での検索"
    Dim basenameStrengthMatch As String
    basenameStrengthMatch = FindBasenameStrengthMatch(searchParts, targetDrugs)
    
    If Len(basenameStrengthMatch) > 0 Then
        Debug.Print "  医薬品名+規格一致: " & basenameStrengthMatch
        FindBestMatchingDrug = basenameStrengthMatch
        Exit Function
    End If
    
    ' 段階3: 包装形態を考慮した類似度検索（従来のロジック）
    Debug.Print "段階3: 包装形態を考慮した類似度検索"
    Dim fuzzyMatch As String
    fuzzyMatch = FindFuzzyMatch(searchTerm, targetDrugs, packageType)
    
    If Len(fuzzyMatch) > 0 Then
        Debug.Print "  類似度検索一致: " & fuzzyMatch
        FindBestMatchingDrug = fuzzyMatch
        Exit Function
    End If
    
    ' いずれの方法でも一致が見つからなかった場合
    Debug.Print "一致する薬品が見つかりませんでした"
    FindBestMatchingDrug = ""
    Exit Function
    
ErrorHandler:
    Debug.Print "FindBestMatchingDrug エラー: " & Err.Number & " - " & Err.Description
    FindBestMatchingDrug = ""
End Function

' 医薬品名（ベースネーム）で完全一致する薬品を探す関数
Private Function FindExactBasenameMatch(ByRef searchParts As DrugNameParts, ByRef targetDrugs() As String) As String
    Dim i As Long
    
    ' 検索対象の基本名称が空の場合は処理しない
    If Len(searchParts.BaseName) = 0 Then
        FindExactBasenameMatch = ""
        Exit Function
    End If
    
    Dim targetParts As DrugNameParts
    
    For i = LBound(targetDrugs) To UBound(targetDrugs)
        If Len(targetDrugs(i)) > 0 Then
            targetParts = DrugNameParser_Mac.ParseDrugString(targetDrugs(i))
            
            ' 基本名称の完全一致チェック（大文字小文字を区別しない）
            If StrComp(searchParts.BaseName, targetParts.BaseName, vbTextCompare) = 0 Then
                FindExactBasenameMatch = targetDrugs(i)
                Exit Function
            End If
        End If
    Next i
    
    FindExactBasenameMatch = ""
End Function

' 医薬品名+規格で一致する薬品を探す関数
Private Function FindBasenameStrengthMatch(ByRef searchParts As DrugNameParts, ByRef targetDrugs() As String) As String
    Dim i As Long
    
    ' 検索対象の基本名称または規格が空の場合は処理しない
    If Len(searchParts.BaseName) = 0 Or Len(searchParts.Strength) = 0 Then
        FindBasenameStrengthMatch = ""
        Exit Function
    End If
    
    Dim targetParts As DrugNameParts
    Dim bestMatchIndex As Long
    Dim bestMatchScore As Double
    bestMatchScore = 0
    bestMatchIndex = -1
    
    For i = LBound(targetDrugs) To UBound(targetDrugs)
        If Len(targetDrugs(i)) > 0 Then
            targetParts = DrugNameParser_Mac.ParseDrugString(targetDrugs(i))
            
            ' 基本名称の類似度（0.7以上）と規格の完全一致
            Dim basenameSimilarity As Double
            basenameSimilarity = DrugNameParser_Mac.GetSimilarity(searchParts.BaseName, targetParts.BaseName)
            
            If basenameSimilarity >= 0.7 And StrComp(searchParts.Strength, targetParts.Strength, vbTextCompare) = 0 Then
                If basenameSimilarity > bestMatchScore Then
                    bestMatchScore = basenameSimilarity
                    bestMatchIndex = i
                End If
            End If
        End If
    Next i
    
    If bestMatchIndex >= 0 Then
        FindBasenameStrengthMatch = targetDrugs(bestMatchIndex)
    Else
        FindBasenameStrengthMatch = ""
    End If
End Function

' 包装形態を考慮した類似度検索（従来のロジック）
Private Function FindFuzzyMatch(ByVal searchTerm As String, ByRef targetDrugs() As String, ByVal packageType As String) As String
    On Error GoTo ErrorHandler
    
    ' エラーメッセージの処理
    If Left(searchTerm, 1) = "[" And Right(searchTerm, 1) = "]" Then
        ' エラーメッセージ（[マスターシートなし]など）はそのまま返す
        FindFuzzyMatch = ""
        Exit Function
    End If
    
    ' キーワードを抽出
    Dim keywords As Variant
    keywords = DrugNameParser_Mac.ExtractKeywords(searchTerm)
    
    ' キーワードのデバッグ出力
    Dim keywordStr As String, j As Long
    keywordStr = "キーワード: "
    If Not IsEmpty(keywords) Then
        For j = LBound(keywords) To UBound(keywords)
            keywordStr = keywordStr & keywords(j) & ", "
        Next j
    End If
    Debug.Print keywordStr
    
    ' 最高スコアの初期化
    Dim bestScore As Double
    bestScore = 0
    
    Dim bestMatch As String
    bestMatch = ""
    
    ' 検索語の包装形態を解析
    Dim searchParts As DrugNameParts
    searchParts = DrugNameParser_Mac.ParseDrugString(searchTerm)
    
    ' 各ターゲット薬品との一致度を計算
    Dim i As Long
    Dim matchCount As Long, processedCount As Long
    Dim highMatchCount As Long, mediumMatchCount As Long
    matchCount = 0
    processedCount = 0
    highMatchCount = 0
    mediumMatchCount = 0
    
    ' より安全なインデックス処理のためにOn Errorを使用
    On Error Resume Next
    
    ' 上位マッチを格納する配列
    Dim allMatches() As Variant
    ReDim allMatches(0 To 20, 0 To 1) ' 最大20件の上位マッチを保存 [ターゲット, スコア]
    Dim matchesCount As Long
    matchesCount = 0
    
    For i = LBound(targetDrugs) To UBound(targetDrugs)
        Dim target As String
        target = targetDrugs(i)
        
        If Len(target) > 0 Then
            processedCount = processedCount + 1
            
            ' 包装形態のチェック
            Dim targetParts As DrugNameParts
            targetParts = DrugNameParser_Mac.ParseDrugString(target)
            
            Dim packageMatch As Boolean
            Dim packageBonus As Double
            packageMatch = True
            packageBonus = 0
            
            ' 包装形態が指定されている場合、一致するか確認
            If Len(packageType) > 0 Then
                ' 包装形態の照合をより柔軟に
                packageMatch = IsPackageTypeMatching(target, packageType, targetParts)
                
                If packageMatch Then
                    packageBonus = 0.1 ' 明示的な一致にはボーナス
                End If
                
                ' 包装形態が一致しない場合でも、基本名と規格が似ていれば許容
                If Not packageMatch Then
                    ' 基本名が類似していて規格が一致する場合
                    If DrugNameParser_Mac.GetSimilarity(searchParts.BaseName, targetParts.BaseName) >= 0.8 And _
                       (Len(searchParts.Strength) = 0 Or Len(targetParts.Strength) = 0 Or _
                        searchParts.Strength = targetParts.Strength) Then
                        packageMatch = True
                        Debug.Print "  包装形態不一致だが基本名類似と規格一致による許容: " & target
                    End If
                End If
            End If
            
            ' 包装形態が一致するか一致しなくても基本名が似ている場合にスコア計算
            If packageMatch Then
                ' スコアの計算
                Dim score As Double
                score = 0
                
                ' 完全一致の場合は最高スコア
                If searchTerm = target Then
                    score = 1
                    Debug.Print "  完全一致: " & target & " (スコア=1.0)"
                Else
                    ' 薬品名解析による詳細な一致度を計算
                    Dim compareScore As Double
                    compareScore = DrugNameParser_Mac.CompareDrugStringsWithRate(searchTerm, target)
                    
                    ' キーワードごとの一致度を加算
                    Dim keywordScore As Double
                    keywordScore = 0
                    
                    ' キーワード一致のデバッグ情報
                    Dim matchedKeywords As String
                    matchedKeywords = ""
                    
                    If Not IsEmpty(keywords) Then
                        Dim keywordCount As Long
                        keywordCount = UBound(keywords) - LBound(keywords) + 1
                        
                        If keywordCount > 0 Then
                            For j = LBound(keywords) To UBound(keywords)
                                Dim keyword As String
                                keyword = keywords(j)
                                
                                If Len(keyword) > 0 And InStr(1, target, keyword, vbTextCompare) > 0 Then
                                    ' キーワードの重みに応じてスコアを加算（オーバーフロー対策）
                                    keywordScore = keywordScore + (1# / CDbl(keywordCount))
                                    matchedKeywords = matchedKeywords & keyword & ", "
                                End If
                            Next j
                        End If
                    End If
                    
                    ' キーワードスコアを加算（比率調整）
                    score = (compareScore * 0.7) + (keywordScore * 0.3)
                    
                    ' 包装形態ボーナスを追加
                    score = score + packageBonus
                    
                    ' 最大スコアは1.0
                    If score > 1 Then score = 1
                    
                    ' デバッグ出力
                    Debug.Print "  対象: " & target
                    Debug.Print "    比較スコア: " & compareScore & ", キーワードスコア: " & keywordScore
                    Debug.Print "    包装ボーナス: " & packageBonus
                    Debug.Print "    一致キーワード: " & matchedKeywords
                    Debug.Print "    最終スコア: " & score
                End If
                
                ' マッチレベルのカウント
                If score >= 0.5 Then
                    highMatchCount = highMatchCount + 1
                ElseIf score >= 0.3 Then
                    mediumMatchCount = mediumMatchCount + 1
                End If
                
                ' スコアが一定以上の場合は上位マッチリストに追加
                If score >= 0.2 Then
                    If matchesCount < UBound(allMatches, 1) Then
                        allMatches(matchesCount, 0) = target
                        allMatches(matchesCount, 1) = score
                        matchesCount = matchesCount + 1
                    Else
                        ' 上位マッチリストが一杯の場合、最も低いスコアと比較
                        Dim lowestIndex As Long
                        Dim lowestScore As Double
                        lowestIndex = 0
                        lowestScore = CDbl(allMatches(0, 1))
                        
                        ' 安全なループ処理
                        Dim k As Long
                        For k = 1 To matchesCount - 1
                            If k <= UBound(allMatches, 1) Then
                                If CDbl(allMatches(k, 1)) < lowestScore Then
                                    lowestIndex = k
                                    lowestScore = CDbl(allMatches(k, 1))
                                End If
                            End If
                        Next k
                        
                        ' より高いスコアの場合は置き換え
                        If score > lowestScore Then
                            allMatches(lowestIndex, 0) = target
                            allMatches(lowestIndex, 1) = score
                        End If
                    End If
                End If
                
                ' 最高スコアの更新
                If score > bestScore Then
                    bestScore = score
                    bestMatch = target
                    Debug.Print "  ★新たな最高スコア: " & score & " -> " & target
                End If
                
                ' スコアがしきい値を超えたカウント
                If score >= 0.2 Then
                    matchCount = matchCount + 1
                End If
            Else
                Debug.Print "  包装形態不一致: " & target & " (スキップ)"
            End If
        End If
    Next i
    
    ' エラー状態をリセット
    On Error GoTo ErrorHandler
    
    ' 上位マッチのデバッグ出力
    Debug.Print "上位マッチ一覧:"
    For i = 0 To matchesCount - 1
        Debug.Print "  " & allMatches(i, 0) & " (スコア: " & allMatches(i, 1) & ")"
    Next i
    
    ' 最終結果のデバッグ出力
    Debug.Print "検索結果: スコア " & bestScore & " の " & bestMatch
    Debug.Print "処理した薬品数: " & processedCount
    Debug.Print "高一致(>=0.5): " & highMatchCount & "件, 中一致(>=0.3): " & mediumMatchCount & "件"
    
    ' しきい値の動的調整
    Dim threshold As Double
    threshold = 0.3  ' 基本しきい値
    
    ' 結果の分布に応じて閾値を調整
    If highMatchCount >= 3 Then
        ' 高一致が多い場合は閾値を上げる
        threshold = 0.5
        Debug.Print "多数の高一致: しきい値を0.5に上げて再判定"
    ElseIf highMatchCount = 0 And mediumMatchCount = 0 And matchCount > 0 Then
        ' 高一致も中一致もないが低一致がある場合は閾値を下げる
        threshold = 0.2
        Debug.Print "一致が少ない: しきい値を0.2に下げて再判定"
    End If
    
    ' 十分な一致度がある場合のみ結果を返す
    If bestScore >= threshold Then
        FindFuzzyMatch = bestMatch
        Debug.Print "最終選択: " & bestMatch & " (スコア " & bestScore & ")"
    Else
        FindFuzzyMatch = ""
        Debug.Print "一致する薬品が見つかりませんでした (最高スコア " & bestScore & ")"
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "FindFuzzyMatch エラー: " & Err.Number & " - " & Err.Description
    FindFuzzyMatch = ""
End Function

' 包装形態の一致を確認する関数
Private Function IsPackageTypeMatching(ByVal targetDrug As String, ByVal packageType As String, ByRef targetParts As DrugNameParts) As Boolean
    ' 対象に指定の包装形態が含まれているか確認
    If InStr(1, packageType, "バラ", vbTextCompare) > 0 Then
        ' バラ包装系なら"バラ"または"調剤用"が含まれているか確認
        If InStr(1, targetDrug, "バラ", vbTextCompare) > 0 Or _
           InStr(1, targetDrug, "調剤用", vbTextCompare) > 0 Then
            IsPackageTypeMatching = True
            Debug.Print "  包装形態一致(バラ): " & targetDrug
            Exit Function
        End If
    ElseIf InStr(1, packageType, "PTP", vbTextCompare) > 0 Then
        ' PTP包装系なら"PTP"が含まれているか確認
        If InStr(1, targetDrug, "PTP", vbTextCompare) > 0 Then
            IsPackageTypeMatching = True
            Debug.Print "  包装形態一致(PTP): " & targetDrug
            Exit Function
        End If
    ElseIf InStr(1, packageType, "分包", vbTextCompare) > 0 Then
        ' 分包系なら"分包"が含まれているか確認
        If InStr(1, targetDrug, "分包", vbTextCompare) > 0 Then
            IsPackageTypeMatching = True
            Debug.Print "  包装形態一致(分包): " & targetDrug
            Exit Function
        End If
    ElseIf InStr(1, packageType, "SP", vbTextCompare) > 0 Then
        ' SP包装系なら"SP"が含まれているか確認
        If InStr(1, targetDrug, "SP", vbTextCompare) > 0 Then
            IsPackageTypeMatching = True
            Debug.Print "  包装形態一致(SP): " & targetDrug
            Exit Function
        End If
    ElseIf InStr(1, packageType, "包装小", vbTextCompare) > 0 Then
        ' 包装小系なら"包装小"が含まれているか確認
        If InStr(1, targetDrug, "包装小", vbTextCompare) > 0 Then
            IsPackageTypeMatching = True
            Debug.Print "  包装形態一致(包装小): " & targetDrug
            Exit Function
        End If
    End If
    
    ' ターゲットに包装形態が明記されていない場合
    If Len(targetParts.Package) = 0 Then
        ' 包装形態が特定できない場合はマッチしていると判断
        IsPackageTypeMatching = True
        Debug.Print "  包装形態なしによる許容: " & targetDrug
        Exit Function
    End If
    
    ' 一致しない場合
    IsPackageTypeMatching = False
End Function

' 一致マーカーを処理中のセルに追加する関数
Public Sub AddMatchMarker(ByVal targetCell As Range, ByVal matchType As String)
    With targetCell
        Select Case matchType
            Case "完全一致"
                .Interior.Color = RGB(198, 239, 206) ' 薄い緑
            Case "部分一致"
                .Interior.Color = RGB(255, 235, 156) ' 薄い黄色
            Case "不一致"
                .Interior.Color = RGB(255, 199, 206) ' 薄い赤
        End Select
    End With
End Sub

' ファイル選択ダイアログを表示し、選択されたファイルパスを返す関数
Public Function GetFilePathFromDialog(Optional ByVal fileFilter As String = "Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", Optional ByVal dialogTitle As String = "ファイルを選択") As String
    Dim filePath As String
    filePath = Application.GetOpenFilename(fileFilter, , dialogTitle, , False)
    
    ' キャンセルされた場合は空文字を返す
    If filePath = "False" Then
        GetFilePathFromDialog = ""
    Else
        GetFilePathFromDialog = filePath
    End If
End Function

' データのバックアップを作成する関数
Public Sub BackupWorksheet(ByVal sourceSheet As Worksheet, ByVal backupName As String)
    ' バックアップシートがすでに存在する場合は削除
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(backupName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' シートをコピーして名前を変更
    sourceSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ActiveSheet.Name = backupName
    
    ' バックアップ日時を記録
    ActiveSheet.Cells(1, 1).Value = "バックアップ: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

' ============================================================
' 設定と初期化関数
' ============================================================

' 指定されたシートの内容をクリアする関数
Public Sub ClearSheet(ByVal ws As Worksheet, Optional ByVal startRow As Long = 1, Optional ByVal preserveFormatting As Boolean = True)
    If preserveFormatting Then
        ' フォーマットを保持しながらクリア
        ws.Cells.ClearContents
    Else
        ' すべてクリア
        ws.Cells.Clear
    End If
End Sub

' 使用方法の説明を追加する関数
Public Sub AddInstructions()
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' A35から下の内容をクリア（既存の指示があれば削除）
    settingsSheet.Range("A35:E50").ClearContents
    
    ' 使用方法の説明を追加
    Dim instructions As Variant
    instructions = Array("■ 使用方法", _
                         "1. 医薬品コードシートに医薬品コードを入力します（A7セルから下）", _
                         "2. 設定シートで包装形態を選択します（B4セル）", _
                         "3. メニューから「ツール」→「マクロ」→「ProcessDrugCodesAndCompare」を実行します", _
                         "", _
                         "■ 動作内容", _
                         "* 処理中はステータスバーに進捗状況が表示されます", _
                         "* 最初の医薬品名から自動的に包装形態が判定され、最適な結果が得られるように処理されます", _
                         "* 一致した医薬品名はB列に表示されます", _
                         "", _
                         "■ パッケージタイプについて", _
                         "* 「バラ包装」：バラや調剤用の薬品を優先的に検索します", _
                         "* 「分包品」：PTP、分包、SP包装の薬品を優先的に検索します", _
                         "", _
                         "■ 処理のパフォーマンス", _
                         "* 最初の医薬品から包装規格を自動判定し、処理を最適化します", _
                         "* 処理された結果は、完了時にメッセージボックスで表示されます")
    
    ' 説明文を表示
    Dim i As Long
    For i = LBound(instructions) To UBound(instructions)
        settingsSheet.Cells(35 + i, 1).Value = instructions(i)
    Next i
    
    ' 書式設定
    With settingsSheet.Range("A35")
        .Font.Bold = True
        .Font.Size = 11
    End With
End Sub 