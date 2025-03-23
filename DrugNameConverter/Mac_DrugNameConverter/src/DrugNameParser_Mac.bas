Option Explicit

' 薬品名の構成要素を格納するための型定義
Public Type DrugNameParts
    BaseName As String     ' 基本名称
    FormType As String     ' 剤形
    Strength As String     ' 規格・含量
    Maker As String        ' メーカー
    Package As String      ' 包装形態
    PackageSize As String  ' 包装単位
End Type

' 包装形態の定義（他のモジュールと共通）
Private Const PACKAGE_TYPE_PTP As String = "PTP"
Private Const PACKAGE_TYPE_BULK As String = "バラ"
Private Const PACKAGE_TYPE_UNIT_DOSE As String = "分包"
Private Const PACKAGE_TYPE_SP As String = "SP"
Private Const PACKAGE_TYPE_SMALL As String = "包装小"
Private Const PACKAGE_TYPE_OTHER As String = "その他"
Private Const PACKAGE_TYPE_UNKNOWN As String = "不明"
Private Const PACKAGE_TYPE_PTP_PATIENT As String = "PTP(患者用)"
Private Const PACKAGE_TYPE_DISPENSING As String = "調剤用"

' 文字列を解析して構成要素に分解する関数
Public Function ParseDrugString(ByVal drugString As String) As DrugNameParts
    Dim result As DrugNameParts
    
    ' デバッグ出力
    Debug.Print "解析対象薬品名: " & drugString
    
    ' 空文字列のチェック
    If Len(drugString) = 0 Then
        Debug.Print "  空の薬品名"
        ParseDrugString = result
        Exit Function
    End If
    
    ' 全角文字を半角に変換（特に数字やカッコ）
    drugString = ConvertToHalfWidth(drugString)
    
    ' 括弧内容の抽出
    Dim bracketsContent As String
    bracketsContent = ExtractBracketsContent(drugString)
    Debug.Print "  括弧内容: " & bracketsContent
    
    ' 基本名称（括弧を除いた部分）
    result.BaseName = RemoveBracketsContent(drugString)
    Debug.Print "  基本名称: " & result.BaseName
    
    ' 剤形の抽出
    result.FormType = ExtractFormType(result.BaseName)
    Debug.Print "  剤形: " & result.FormType
    
    ' 規格・含量の抽出
    result.Strength = ExtractStrength(result.BaseName)
    Debug.Print "  規格: " & result.Strength
    
    ' メーカーの抽出（括弧内から）
    result.Maker = ExtractMaker(bracketsContent)
    Debug.Print "  メーカー: " & result.Maker
    
    ' 包装形態の抽出（括弧内から）
    result.Package = ExtractPackageTypeSimple(bracketsContent)
    
    ' 包装形態が見つからない場合は薬品名全体から再検索
    If Len(result.Package) = 0 Then
        result.Package = ExtractPackageTypeSimple(drugString)
    End If
    Debug.Print "  包装形態: " & result.Package
    
    ' 包装単位の抽出（括弧内から）
    result.PackageSize = ExtractPackageSize(bracketsContent)
    Debug.Print "  包装単位: " & result.PackageSize
    
    ParseDrugString = result
End Function

' 全角文字を半角に変換する関数
Private Function ConvertToHalfWidth(ByVal text As String) As String
    Dim i As Long
    Dim result As String
    Dim c As String
    
    result = ""
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        ' 全角数字を半角に変換
        Select Case c
            Case "０": result = result & "0"
            Case "１": result = result & "1"
            Case "２": result = result & "2"
            Case "３": result = result & "3"
            Case "４": result = result & "4"
            Case "５": result = result & "5"
            Case "６": result = result & "6"
            Case "７": result = result & "7"
            Case "８": result = result & "8"
            Case "９": result = result & "9"
            Case "（": result = result & "("
            Case "）": result = result & ")"
            Case "［": result = result & "["
            Case "］": result = result & "]"
            Case Else: result = result & c
        End Select
    Next i
    
    ConvertToHalfWidth = result
End Function

' 括弧内の内容を抽出する関数
Private Function ExtractBracketsContent(ByVal drugString As String) As String
    Dim result As String
    Dim startPos As Long, endPos As Long
    
    ' 最初の開き括弧を検索
    startPos = InStr(1, drugString, "(")
    If startPos = 0 Then
        ' 丸括弧がない場合は角括弧をチェック
        startPos = InStr(1, drugString, "[")
        If startPos = 0 Then
            ' どちらの括弧もない場合
            ExtractBracketsContent = ""
            Exit Function
        Else
            ' 対応する閉じ角括弧を検索
            endPos = InStr(startPos + 1, drugString, "]")
        End If
    Else
        ' 対応する閉じ丸括弧を検索
        endPos = InStr(startPos + 1, drugString, ")")
    End If
    
    ' 閉じ括弧が見つからない場合
    If endPos = 0 Then
        ExtractBracketsContent = ""
        Exit Function
    End If
    
    ' 括弧内の内容を抽出（括弧自体は含まない）
    result = Mid(drugString, startPos + 1, endPos - startPos - 1)
    
    ExtractBracketsContent = result
End Function

' 括弧内の内容を除去する関数（Mac環境向け実装）
Private Function RemoveBracketsContent(ByVal drugString As String) As String
    Dim result As String
    Dim pos1 As Long, pos2 As Long
    result = drugString
    
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

' 剤形を抽出する関数（改善版）
Private Function ExtractFormType(ByVal drugName As String) As String
    ' 完全一致型
    Dim exactFormTypes As Variant
    exactFormTypes = Array("錠", "カプセル", "細粒", "顆粒", "散", "シロップ", "液", "注射液", "注", "軟膏", "クリーム", "ゲル", "ローション", "点眼液", "目薬", "点鼻液", "吸入液", "貼付剤", "パッチ", "坐剤", "ドライシロップ", "内用液")
    
    ' 部分一致型（より具体的な剤形が先）
    Dim partialFormTypes As Variant
    partialFormTypes = Array("フィルムコーティング錠", "フィルム錠", "徐放錠", "口腔内崩壊錠", "OD錠", "腸溶錠", "チュアブル錠", "分散錠", "舌下錠", "硬カプセル", "軟カプセル", "細粒剤", "顆粒剤", "散剤", "エマルション", "クリーム剤", "外用液", "軟膏剤", "ゲル剤")
    
    ' まず部分一致型で検索（より具体的な剤形を優先）
    Dim i As Long
    For i = LBound(partialFormTypes) To UBound(partialFormTypes)
        If InStr(drugName, partialFormTypes(i)) > 0 Then
            ExtractFormType = partialFormTypes(i)
            Exit Function
        End If
    Next i
    
    ' 次に完全一致型で検索
    For i = LBound(exactFormTypes) To UBound(exactFormTypes)
        If InStr(drugName, exactFormTypes(i)) > 0 Then
            ExtractFormType = exactFormTypes(i)
            Exit Function
        End If
    Next i
    
    ' 特別なケース：OD錠（口腔内崩壊錠）
    If InStr(drugName, "OD") > 0 And InStr(drugName, "錠") > 0 Then
        ExtractFormType = "OD錠"
        Exit Function
    End If
    
    ExtractFormType = ""
End Function

' 規格・含量を抽出する関数（Mac環境向け実装）
Private Function ExtractStrength(ByVal drugName As String) As String
    Dim result As String
    result = ""
    
    ' 数字+単位のパターンを検索（正規表現の代わりに文字列操作で実装）
    Dim i As Long, j As Long
    Dim units As Variant
    units = Array("mg", "g", "mL", "μg", "単位", "IU", "%")
    
    ' 数字の開始位置を探す
    For i = 1 To Len(drugName)
        If IsNumeric(Mid(drugName, i, 1)) Then
            ' 数字が見つかったら、その後の単位を確認
            Dim numEnd As Long
            numEnd = i
            
            ' 数字部分の終了位置を特定
            Do While numEnd <= Len(drugName) And (IsNumeric(Mid(drugName, numEnd, 1)) Or Mid(drugName, numEnd, 1) = ".")
                numEnd = numEnd + 1
            Loop
            
            ' 単位を探す
            For j = LBound(units) To UBound(units)
                If InStr(numEnd, drugName, units(j)) = numEnd Then
                    ' 数値+単位を取得
                    result = Mid(drugName, i, numEnd - i + Len(units(j)))
                    ExtractStrength = Trim(result)
                    Exit Function
                End If
            Next j
        End If
    Next i
    
    ExtractStrength = ""
End Function

' メーカー名を抽出する関数
Private Function ExtractMaker(ByVal bracketsContent As String) As String
    Dim makers As Variant
    makers = Array("武田", "第一三共", "アステラス", "エーザイ", "田辺三菱", "大塚", "アストラゼネカ", "ノバルティス", "ファイザー", "MSD", "バイエル", "大正", "中外", "参天", "久光", "杏林", "沢井", "東和", "日医工", "あすか", "ニプロ", "サンド", "陽進堂", "科研", "キョーリン", "ツムラ", "日本ケミファ", "トーアエイヨー", "共和", "明治", "救急", "持田", "ゼリア", "小野", "協和", "Meiji Seika", "テバ", "富士", "マイラン", "ヤンセン", "ギリアド", "シオノギ", "塩野義", "アッヴィ", "ブリストル", "テルモ", "帝人", "キッセイ", "ロシュ", "グラクソ", "サノフィ", "大日本住友", "興和", "鳥居")
    
    Dim i As Long
    For i = LBound(makers) To UBound(makers)
        If InStr(1, bracketsContent, makers(i), vbTextCompare) > 0 Then
            ExtractMaker = makers(i)
            Exit Function
        End If
    Next i
    
    ExtractMaker = ""
End Function

' 包装単位を抽出する関数
Private Function ExtractPackageSize(ByVal bracketsContent As String) As String
    ' 数字+単位のパターン（例：100錠、500g）
    Dim i As Long
    Dim inNumber As Boolean
    Dim numStart As Long
    Dim numEnd As Long
    
    inNumber = False
    
    For i = 1 To Len(bracketsContent)
        If IsNumeric(Mid(bracketsContent, i, 1)) Then
            If Not inNumber Then
                inNumber = True
                numStart = i
            End If
        ElseIf inNumber Then
            numEnd = i - 1
            
            ' 単位が続くか確認
            Dim units As Variant
            units = Array("錠", "カプセル", "g", "mg", "mL", "袋", "本", "個", "枚")
            
            Dim j As Long
            For j = LBound(units) To UBound(units)
                If InStr(i, bracketsContent, units(j)) = i Then
                    ' 数字+単位を抽出
                    ExtractPackageSize = Mid(bracketsContent, numStart, numEnd - numStart + 1 + Len(units(j)))
                    Exit Function
                End If
            Next j
            
            inNumber = False
        End If
    Next i
    
    ExtractPackageSize = ""
End Function

' 医薬品名から包装形態を判定する関数
Public Function DetectPackageType(ByVal drugName As String) As String
    On Error GoTo ErrorHandler
    
    ' デバッグ出力をさらに詳細に
    Debug.Print "===== DetectPackageType呼び出し ====="
    Debug.Print "入力薬品名: [" & drugName & "]"
    
    ' 空の医薬品名の場合
    If Len(drugName) = 0 Then
        Debug.Print "  結果: 空の医薬品名 → " & PACKAGE_TYPE_UNKNOWN
        DetectPackageType = PACKAGE_TYPE_UNKNOWN
        Exit Function
    End If
    
    ' エラーメッセージの場合は不明を返す
    If Left(drugName, 1) = "[" And Right(drugName, 1) = "]" Then
        Debug.Print "  結果: エラーメッセージ → " & PACKAGE_TYPE_UNKNOWN
        DetectPackageType = PACKAGE_TYPE_UNKNOWN
        Exit Function
    End If
    
    ' ここで薬品名のトリムを行う（前後の空白を除去）
    drugName = Trim(drugName)
    Debug.Print "トリム後の薬品名: [" & drugName & "]"
    
    ' 画像に示された9種類の包装形態パターン
    Dim packageTypePatterns As Variant
    packageTypePatterns = Array("(未定義)", "その他(なし)", "包装小", "調剤用", "PTP", "分包", "バラ", "SP", "PTP(患者用)")
    
    ' 1. 最初にスペースで区切られた単語から包装形態を探す
    ' データベースの形式に合わせた処理（例: エリキュース錠２．５ｍｇ ＰＴＰ 10錠）
    Dim words As Variant
    words = Split(drugName, " ")
    
    Debug.Print "  単語分解："
    Dim i As Long, word As String
    For i = LBound(words) To UBound(words)
        word = Trim(words(i))
        Debug.Print "    単語[" & i & "]: [" & word & "]"
        
        ' 全角半角を考慮して包装形態を検出
        If word = "PTP" Or word = "ＰＴＰ" Then
            Debug.Print "  結果: スペース区切りから検出 PTP → " & PACKAGE_TYPE_PTP
            DetectPackageType = PACKAGE_TYPE_PTP
            Exit Function
        ElseIf word = "PTP(患者用)" Or word = "ＰＴＰ(患者用)" Then
            Debug.Print "  結果: スペース区切りから検出 PTP(患者用) → " & PACKAGE_TYPE_PTP_PATIENT
            DetectPackageType = PACKAGE_TYPE_PTP_PATIENT
            Exit Function
        ElseIf word = "バラ" Then
            Debug.Print "  結果: スペース区切りから検出 バラ → " & PACKAGE_TYPE_BULK
            DetectPackageType = PACKAGE_TYPE_BULK
            Exit Function
        ElseIf word = "分包" Then
            Debug.Print "  結果: スペース区切りから検出 分包 → " & PACKAGE_TYPE_UNIT_DOSE
            DetectPackageType = PACKAGE_TYPE_UNIT_DOSE
            Exit Function
        ElseIf word = "SP" Or word = "ＳＰ" Then
            Debug.Print "  結果: スペース区切りから検出 SP → " & PACKAGE_TYPE_SP
            DetectPackageType = PACKAGE_TYPE_SP
            Exit Function
        ElseIf word = "包装小" Then
            Debug.Print "  結果: スペース区切りから検出 包装小 → " & PACKAGE_TYPE_SMALL
            DetectPackageType = PACKAGE_TYPE_SMALL
            Exit Function
        ElseIf word = "調剤用" Then
            Debug.Print "  結果: スペース区切りから検出 調剤用 → " & PACKAGE_TYPE_DISPENSING
            DetectPackageType = PACKAGE_TYPE_DISPENSING
            Exit Function
        ElseIf word = "(未定義)" Or word = "その他(なし)" Then
            Debug.Print "  結果: スペース区切りから検出 その他 → " & PACKAGE_TYPE_OTHER
            DetectPackageType = PACKAGE_TYPE_OTHER
            Exit Function
        ElseIf word = "注射剤" Then
            Debug.Print "  結果: スペース区切りから注射剤を検出 → " & PACKAGE_TYPE_OTHER
            DetectPackageType = PACKAGE_TYPE_OTHER
            Exit Function
        End If
    Next i
    
    ' 2. 「/包装形態/」形式を処理
    Dim startPos As Long, endPos As Long
    startPos = InStr(1, drugName, "/")
    
    If startPos > 0 Then
        endPos = InStr(startPos + 1, drugName, "/")
        
        If endPos > startPos Then
            ' /で囲まれた部分を抽出
            Dim packageType As String
            packageType = Mid(drugName, startPos + 1, endPos - startPos - 1)
            Debug.Print "  「/包装形態/」形式から抽出: " & packageType
            
            ' 抽出した包装形態をトリム（前後の空白を除去）
            packageType = Trim(packageType)
            Debug.Print "  トリム後の包装形態: " & packageType
            
            ' 包装形態を判定 - 完全一致を優先
            Select Case packageType
                Case "PTP", "ＰＴＰ"
                    Debug.Print "  結果: 完全一致 PTP → " & PACKAGE_TYPE_PTP
                    DetectPackageType = PACKAGE_TYPE_PTP
                    Exit Function
                Case "PTP(患者用)"
                    Debug.Print "  結果: 完全一致 PTP(患者用) → " & PACKAGE_TYPE_PTP_PATIENT
                    DetectPackageType = PACKAGE_TYPE_PTP_PATIENT
                    Exit Function
                Case "調剤用"
                    Debug.Print "  結果: 完全一致 調剤用 → " & PACKAGE_TYPE_DISPENSING
                    DetectPackageType = PACKAGE_TYPE_DISPENSING
                    Exit Function
                Case "バラ"
                    Debug.Print "  結果: 完全一致 バラ → " & PACKAGE_TYPE_BULK
                    DetectPackageType = PACKAGE_TYPE_BULK
                    Exit Function
                Case "分包"
                    Debug.Print "  結果: 完全一致 分包 → " & PACKAGE_TYPE_UNIT_DOSE
                    DetectPackageType = PACKAGE_TYPE_UNIT_DOSE
                    Exit Function
                Case "SP", "ＳＰ"
                    Debug.Print "  結果: 完全一致 SP → " & PACKAGE_TYPE_SP
                    DetectPackageType = PACKAGE_TYPE_SP
                    Exit Function
                Case "包装小"
                    Debug.Print "  結果: 完全一致 包装小 → " & PACKAGE_TYPE_SMALL
                    DetectPackageType = PACKAGE_TYPE_SMALL
                    Exit Function
                Case "(未定義)", "その他(なし)"
                    Debug.Print "  結果: 完全一致 その他 → " & PACKAGE_TYPE_OTHER
                    DetectPackageType = PACKAGE_TYPE_OTHER
                    Exit Function
            End Select
            
            ' 完全一致しない場合は部分一致を検索
            For i = LBound(packageTypePatterns) To UBound(packageTypePatterns)
                If InStr(1, packageType, packageTypePatterns(i), vbTextCompare) > 0 Then
                    Debug.Print "  結果: 部分一致 " & packageTypePatterns(i)
                    Select Case packageTypePatterns(i)
                        Case "PTP"
                            Debug.Print "  → " & PACKAGE_TYPE_PTP
                            DetectPackageType = PACKAGE_TYPE_PTP
                            Exit Function
                        Case "PTP(患者用)"
                            Debug.Print "  → " & PACKAGE_TYPE_PTP_PATIENT
                            DetectPackageType = PACKAGE_TYPE_PTP_PATIENT
                            Exit Function
                        Case "調剤用"
                            Debug.Print "  → " & PACKAGE_TYPE_DISPENSING
                            DetectPackageType = PACKAGE_TYPE_DISPENSING
                            Exit Function
                        Case "バラ"
                            Debug.Print "  → " & PACKAGE_TYPE_BULK
                            DetectPackageType = PACKAGE_TYPE_BULK
                            Exit Function
                        Case "分包"
                            Debug.Print "  → " & PACKAGE_TYPE_UNIT_DOSE
                            DetectPackageType = PACKAGE_TYPE_UNIT_DOSE
                            Exit Function
                        Case "SP"
                            Debug.Print "  → " & PACKAGE_TYPE_SP
                            DetectPackageType = PACKAGE_TYPE_SP
                            Exit Function
                        Case "包装小"
                            Debug.Print "  → " & PACKAGE_TYPE_SMALL
                            DetectPackageType = PACKAGE_TYPE_SMALL
                            Exit Function
                        Case "(未定義)", "その他(なし)"
                            Debug.Print "  → " & PACKAGE_TYPE_OTHER
                            DetectPackageType = PACKAGE_TYPE_OTHER
                            Exit Function
                    End Select
                End If
            Next i
            Debug.Print "  結果: /で囲まれた部分に一致なし"
        End If
    End If
    
    ' 3. 括弧内の情報を処理
    Dim bracketStartPos As Long, bracketEndPos As Long
    bracketStartPos = InStr(1, drugName, "(")
    
    If bracketStartPos > 0 Then
        bracketEndPos = InStr(bracketStartPos, drugName, ")")
        
        If bracketEndPos > bracketStartPos Then
            Dim bracketContent As String
            bracketContent = Mid(drugName, bracketStartPos + 1, bracketEndPos - bracketStartPos - 1)
            Debug.Print "  括弧内の内容: " & bracketContent
            
            ' 括弧内で包装形態パターンが見つかるか確認
            For i = LBound(packageTypePatterns) To UBound(packageTypePatterns)
                If InStr(1, bracketContent, packageTypePatterns(i), vbTextCompare) > 0 Then
                    Select Case packageTypePatterns(i)
                        Case "PTP"
                            Debug.Print "  結果: 括弧内に PTP → " & PACKAGE_TYPE_PTP
                            DetectPackageType = PACKAGE_TYPE_PTP
                            Exit Function
                        Case "PTP(患者用)"
                            Debug.Print "  結果: 括弧内に PTP(患者用) → " & PACKAGE_TYPE_PTP_PATIENT
                            DetectPackageType = PACKAGE_TYPE_PTP_PATIENT
                            Exit Function
                        Case "調剤用"
                            Debug.Print "  結果: 括弧内に 調剤用 → " & PACKAGE_TYPE_DISPENSING
                            DetectPackageType = PACKAGE_TYPE_DISPENSING
                            Exit Function
                        Case "バラ"
                            Debug.Print "  結果: 括弧内に バラ → " & PACKAGE_TYPE_BULK
                            DetectPackageType = PACKAGE_TYPE_BULK
                            Exit Function
                        Case "分包"
                            Debug.Print "  結果: 括弧内に 分包 → " & PACKAGE_TYPE_UNIT_DOSE
                            DetectPackageType = PACKAGE_TYPE_UNIT_DOSE
                            Exit Function
                        Case "SP"
                            Debug.Print "  結果: 括弧内に SP → " & PACKAGE_TYPE_SP
                            DetectPackageType = PACKAGE_TYPE_SP
                            Exit Function
                        Case "包装小"
                            Debug.Print "  結果: 括弧内に 包装小 → " & PACKAGE_TYPE_SMALL
                            DetectPackageType = PACKAGE_TYPE_SMALL
                            Exit Function
                        Case "(未定義)", "その他(なし)"
                            Debug.Print "  結果: 括弧内に その他 → " & PACKAGE_TYPE_OTHER
                            DetectPackageType = PACKAGE_TYPE_OTHER
                            Exit Function
                    End Select
                End If
            Next i
        End If
    End If
    
    ' 4. 医薬品名全体から直接検索（完全一致のみ）
    For i = LBound(packageTypePatterns) To UBound(packageTypePatterns)
        ' 医薬品名に完全な単語として包装形態が含まれているか確認
        If InStr(1, " " & drugName & " ", " " & packageTypePatterns(i) & " ", vbTextCompare) > 0 Then
            Select Case packageTypePatterns(i)
                Case "PTP"
                    Debug.Print "  結果: 薬品名全体に PTP → " & PACKAGE_TYPE_PTP
                    DetectPackageType = PACKAGE_TYPE_PTP
                    Exit Function
                Case "PTP(患者用)"
                    Debug.Print "  結果: 薬品名全体に PTP(患者用) → " & PACKAGE_TYPE_PTP_PATIENT
                    DetectPackageType = PACKAGE_TYPE_PTP_PATIENT
                    Exit Function
                Case "調剤用"
                    Debug.Print "  結果: 薬品名全体に 調剤用 → " & PACKAGE_TYPE_DISPENSING
                    DetectPackageType = PACKAGE_TYPE_DISPENSING
                    Exit Function
                Case "バラ"
                    Debug.Print "  結果: 薬品名全体に バラ → " & PACKAGE_TYPE_BULK
                    DetectPackageType = PACKAGE_TYPE_BULK
                    Exit Function
                Case "分包"
                    Debug.Print "  結果: 薬品名全体に 分包 → " & PACKAGE_TYPE_UNIT_DOSE
                    DetectPackageType = PACKAGE_TYPE_UNIT_DOSE
                    Exit Function
                Case "SP"
                    Debug.Print "  結果: 薬品名全体に SP → " & PACKAGE_TYPE_SP
                    DetectPackageType = PACKAGE_TYPE_SP
                    Exit Function
                Case "包装小"
                    Debug.Print "  結果: 薬品名全体に 包装小 → " & PACKAGE_TYPE_SMALL
                    DetectPackageType = PACKAGE_TYPE_SMALL
                    Exit Function
                Case "(未定義)", "その他(なし)"
                    Debug.Print "  結果: 薬品名全体に その他 → " & PACKAGE_TYPE_OTHER
                    DetectPackageType = PACKAGE_TYPE_OTHER
                    Exit Function
            End Select
        End If
    Next i
    
    ' 5. 全角文字を考慮した追加チェック
    If InStr(1, drugName, "ＰＴＰ", vbTextCompare) > 0 Then
        Debug.Print "  結果: 全角文字 ＰＴＰ → " & PACKAGE_TYPE_PTP
        DetectPackageType = PACKAGE_TYPE_PTP
        Exit Function
    ElseIf InStr(1, drugName, "ＳＰ", vbTextCompare) > 0 Then
        Debug.Print "  結果: 全角文字 ＳＰ → " & PACKAGE_TYPE_SP
        DetectPackageType = PACKAGE_TYPE_SP
        Exit Function
    End If
    
    ' 判定できない場合は不明を返す
    Debug.Print "  結果: 包装形態判定不能 → " & PACKAGE_TYPE_UNKNOWN
    DetectPackageType = PACKAGE_TYPE_UNKNOWN
    Exit Function
    
ErrorHandler:
    Debug.Print "DetectPackageType エラー: " & Err.Number & " - " & Err.Description
    DetectPackageType = PACKAGE_TYPE_UNKNOWN
End Function

' シンプルな包装形態抽出関数
Public Function ExtractPackageTypeSimple(ByVal text As String) As String
    ' デバッグ出力をさらに詳細に
    Debug.Print "--- ExtractPackageTypeSimple呼び出し ---"
    Debug.Print "  入力テキスト: " & text
    
    ' ここで入力テキストのトリムを行う（前後の空白を除去）
    text = Trim(text)
    Debug.Print "  トリム後のテキスト: " & text
    
    ' 画像に示された9種類の包装形態パターンのみを配列に格納
    Dim packageTypePatterns As Variant
    packageTypePatterns = Array("(未定義)", "その他(なし)", "包装小", "調剤用", "PTP", "分包", "バラ", "SP", "PTP(患者用)")
    
    ' 包装形態パターンの完全一致チェック
    Dim i As Long
    For i = LBound(packageTypePatterns) To UBound(packageTypePatterns)
        ' 完全一致または単語として含まれているか確認
        If packageTypePatterns(i) = text Or _
           InStr(1, " " & text & " ", " " & packageTypePatterns(i) & " ", vbTextCompare) > 0 Then
            Debug.Print "  検出: " & packageTypePatterns(i) & " (完全一致/単語一致)"
            ' パターンと包装形態定数のマッピング
            Select Case packageTypePatterns(i)
                Case "PTP"
                    Debug.Print "  結果: " & PACKAGE_TYPE_PTP
                    ExtractPackageTypeSimple = PACKAGE_TYPE_PTP
                    Exit Function
                Case "PTP(患者用)"
                    Debug.Print "  結果: " & PACKAGE_TYPE_PTP_PATIENT
                    ExtractPackageTypeSimple = PACKAGE_TYPE_PTP_PATIENT
                    Exit Function
                Case "調剤用"
                    Debug.Print "  結果: " & PACKAGE_TYPE_DISPENSING
                    ExtractPackageTypeSimple = PACKAGE_TYPE_DISPENSING
                    Exit Function
                Case "バラ"
                    Debug.Print "  結果: " & PACKAGE_TYPE_BULK
                    ExtractPackageTypeSimple = PACKAGE_TYPE_BULK
                    Exit Function
                Case "分包"
                    Debug.Print "  結果: " & PACKAGE_TYPE_UNIT_DOSE
                    ExtractPackageTypeSimple = PACKAGE_TYPE_UNIT_DOSE
                    Exit Function
                Case "SP"
                    Debug.Print "  結果: " & PACKAGE_TYPE_SP
                    ExtractPackageTypeSimple = PACKAGE_TYPE_SP
                    Exit Function
                Case "包装小"
                    Debug.Print "  結果: " & PACKAGE_TYPE_SMALL
                    ExtractPackageTypeSimple = PACKAGE_TYPE_SMALL
                    Exit Function
                Case "(未定義)", "その他(なし)"
                    Debug.Print "  結果: " & PACKAGE_TYPE_OTHER
                    ExtractPackageTypeSimple = PACKAGE_TYPE_OTHER
                    Exit Function
            End Select
        End If
    Next i
    
    ' 部分一致チェック（単語境界を考慮しない）
    For i = LBound(packageTypePatterns) To UBound(packageTypePatterns)
        If InStr(1, text, packageTypePatterns(i), vbTextCompare) > 0 Then
            Debug.Print "  検出: " & packageTypePatterns(i) & " (部分一致)"
            ' パターンと包装形態定数のマッピング
            Select Case packageTypePatterns(i)
                Case "PTP"
                    Debug.Print "  結果: " & PACKAGE_TYPE_PTP
                    ExtractPackageTypeSimple = PACKAGE_TYPE_PTP
                    Exit Function
                Case "PTP(患者用)"
                    Debug.Print "  結果: " & PACKAGE_TYPE_PTP_PATIENT
                    ExtractPackageTypeSimple = PACKAGE_TYPE_PTP_PATIENT
                    Exit Function
                Case "調剤用"
                    Debug.Print "  結果: " & PACKAGE_TYPE_DISPENSING
                    ExtractPackageTypeSimple = PACKAGE_TYPE_DISPENSING
                    Exit Function
                Case "バラ"
                    Debug.Print "  結果: " & PACKAGE_TYPE_BULK
                    ExtractPackageTypeSimple = PACKAGE_TYPE_BULK
                    Exit Function
                Case "分包"
                    Debug.Print "  結果: " & PACKAGE_TYPE_UNIT_DOSE
                    ExtractPackageTypeSimple = PACKAGE_TYPE_UNIT_DOSE
                    Exit Function
                Case "SP"
                    Debug.Print "  結果: " & PACKAGE_TYPE_SP
                    ExtractPackageTypeSimple = PACKAGE_TYPE_SP
                    Exit Function
                Case "包装小"
                    Debug.Print "  結果: " & PACKAGE_TYPE_SMALL
                    ExtractPackageTypeSimple = PACKAGE_TYPE_SMALL
                    Exit Function
                Case "(未定義)", "その他(なし)"
                    Debug.Print "  結果: " & PACKAGE_TYPE_OTHER
                    ExtractPackageTypeSimple = PACKAGE_TYPE_OTHER
                    Exit Function
            End Select
        End If
    Next i
    
    ' 全角記号も考慮する
    If InStr(1, text, "ＰＴＰ", vbTextCompare) > 0 Then
        Debug.Print "  結果: " & PACKAGE_TYPE_PTP & " (全角処理)"
        ExtractPackageTypeSimple = PACKAGE_TYPE_PTP
        Exit Function
    ElseIf InStr(1, text, "ＳＰ", vbTextCompare) > 0 Then
        Debug.Print "  結果: " & PACKAGE_TYPE_SP & " (全角処理)"
        ExtractPackageTypeSimple = PACKAGE_TYPE_SP
        Exit Function
    End If
    
    ' 包装形態が見つからない場合は空文字を返す
    Debug.Print "  結果: 包装形態なし"
    ExtractPackageTypeSimple = ""
End Function