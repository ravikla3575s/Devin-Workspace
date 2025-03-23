Attribute VB_Name = "MainModule"
Option Explicit

' ���C���̏����֐��F��i���̈�v���Ɋ�Â��ē]�L
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
    
    Const MATCH_THRESHOLD As Double = 80 ' ��v���̂������l�i80%�j
    
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
            
            ' ���ʂ̏o��
            If maxMatchRate >= MATCH_THRESHOLD Then
                ws1.Cells(i, "C").Value = ws2.Cells(bestMatchIndex, "B").Value
                ws1.Cells(i, "D").Value = maxMatchRate & "%"
                
                ' ��v�����e�v�f�̏ڍׂ��o�́i�f�o�b�O�p�j
                Dim sourceParts As DrugNameParts
                Dim targetParts As DrugNameParts
                sourceParts = ParseDrugString(sourceStr)
                targetParts = ParseDrugString(ws2.Cells(bestMatchIndex, "B").Value)
                
                ws1.Cells(i, "E").Value = "��{��:" & sourceParts.BaseName & _
                                         " �܌^:" & sourceParts.formType & _
                                         " �K�i:" & sourceParts.strength & _
                                         " ���[�J�[:" & sourceParts.maker
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "�������������܂����B"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "�G���[���������܂���: " & Err.Description
End Sub

' ��i���̌����Ɠ]�L�֐�
Public Sub SearchAndTransferDrugData()
    On Error GoTo ErrorHandler
    
    '��ʍX�V���ꎞ��~���ăp�t�H�[�}���X����
    Application.ScreenUpdating = False
    
    '���[�N�V�[�g�̐ݒ�
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets(1)
    Set ws2 = ThisWorkbook.Worksheets(2)
    Set ws3 = ThisWorkbook.Worksheets(3)
    
    '�ŏI�s�̎擾
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row
    lastRow3 = ws3.Cells(ws3.Rows.Count, "F").End(xlUp).Row
    
    Dim i As Long
    Dim inputValue As Variant
    
    '�e�s��A��̒l������
    For i = 2 To lastRow1  '�w�b�_�[���X�L�b�v
        inputValue = ws1.Cells(i, "A").Value
        
        '���͒l����������֐����Ăяo��
        ProcessInputValue inputValue, ws1, ws2, ws3, i, lastRow2, lastRow3
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "�������������܂����B"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "�G���[���������܂���: " & Err.Description
End Sub

' ���͒l����������֐�
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
    
    'Sheet3�����ܖ�������
    For k = 2 To lastRow3
        drugNameFromSheet3 = ws3.Cells(k, "F").Value
        If InStr(1, inputValue, drugNameFromSheet3) > 0 Then
            'Sheet2����Ή������ܖ�������
            For j = 2 To lastRow2
                drugNameFromSheet2 = ws2.Cells(j, "B").Value
                If drugNameFromSheet2 = drugNameFromSheet3 Then
                    '��^�C�v���擾
                    packageType = GetPackageType(inputValue)
                    
                    '�f�[�^��]�L
                    ws1.Cells(currentRow, "B").Value = ws2.Cells(j, "A").Value
                    ws1.Cells(currentRow, "C").Value = packageType
                    Exit For
                End If
            Next j
            Exit For
        End If
    Next k
End Sub

' ��v���v�Z�ɂ���i�������֐�
Public Sub ProcessDrugNamesWithMatchRate()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets(1)
    Set ws2 = ThisWorkbook.Worksheets(2)
    
    Dim i As Long, j As Long
    Dim lastRow1 As Long, lastRow2 As Long
    Const MATCH_THRESHOLD As Double = 80 ' ��v���̂������l�i80%�j
    
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
        
        ' �������l�ȏ�̈�v�����������ꍇ�̂ݓ]�L
        If maxMatchRate >= MATCH_THRESHOLD Then
            ws1.Cells(i, "C").Value = ws2.Cells(bestMatchIndex, "B").Value
            ws1.Cells(i, "D").Value = maxMatchRate & "%"
        End If
    Next i
    
    MsgBox "�������������܂����B"
End Sub

' �ݒ�V�[�g�̕�`�Ԃ��l���������i����r�Ɠ]�L���s��
Public Sub CompareAndTransferDrugNamesByPackage()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' ���[�N�V�[�g�̐ݒ�
    Dim wsSettings As Worksheet, wsTarget As Worksheet
    Set wsSettings = ThisWorkbook.Worksheets(1) ' �ݒ�V�[�g
    Set wsTarget = ThisWorkbook.Worksheets(2)   ' ��r�Ώۂ̃V�[�g
    
    ' B4�Z�������`�Ԃ��擾
    Dim packageType As String
    packageType = wsSettings.Range("B4").Value
    
    ' �ŏI�s���擾
    Dim lastRowSettings As Long, lastRowTarget As Long
    lastRowSettings = wsSettings.Cells(wsSettings.Rows.Count, "B").End(xlUp).Row
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "B").End(xlUp).Row
    
    ' �����ΏۂƔ�r�Ώۂ̈��i����z��Ɋi�[
    Dim searchDrugs() As String
    Dim targetDrugs() As String
    Dim i As Long, j As Long
    
    ' �������i�p�̔z���������
    ReDim searchDrugs(1 To lastRowSettings - 1) ' �w�b�_�[�s������
    For i = 2 To lastRowSettings
        searchDrugs(i - 1) = wsSettings.Cells(i, "B").Value
    Next i
    
    ' ��r�Ώۗp�̔z���������
    ReDim targetDrugs(1 To lastRowTarget - 1) ' �w�b�_�[�s������
    For i = 2 To lastRowTarget
        targetDrugs(i - 1) = wsTarget.Cells(i, "B").Value
    Next i
    
    ' �e�������i�ɑ΂��Ĕ�r����
    For i = 2 To lastRowSettings
        Dim searchDrug As String
        searchDrug = wsSettings.Cells(i, "B").Value
        
        If Len(searchDrug) > 0 Then
            Dim bestMatch As String
            bestMatch = FindBestMatchWithPackage(searchDrug, targetDrugs, packageType)
            
            If Len(bestMatch) > 0 Then
                ' ��v�������i����C��ɓ]�L
                wsSettings.Cells(i, "C").Value = bestMatch
            Else
                ' ��v���Ȃ������ꍇ�͋󗓂ɂ���
                wsSettings.Cells(i, "C").Value = ""
            End If
        End If
    Next i
    
CleanExit:
    Application.ScreenUpdating = True
    MsgBox "�������������܂����B", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical
End Sub

' ���i���̐����A�K�i�A�P�ʂ̈�v�x���v�Z
Public Function CalculateMatchScore(ByRef searchParts As DrugNameParts, ByRef targetParts As DrugNameParts) As Double
    Dim score As Double
    Dim totalWeight As Double
    
    score = 0
    totalWeight = 0
    
    ' �������̔�r�i�d��: 50%�j
    If StrComp(searchParts.BaseName, targetParts.BaseName, vbTextCompare) = 0 Then
        score = score + 50
    End If
    totalWeight = totalWeight + 50
    
    ' �܌^�̔�r�i�d��: 20%�j
    If StrComp(searchParts.formType, targetParts.formType, vbTextCompare) = 0 Then
        score = score + 20
    End If
    totalWeight = totalWeight + 20
    
    ' �K�i�̔�r�i�d��: 30%�j
    If CompareStrength(searchParts.strength, targetParts.strength) Then
        score = score + 30
    End If
    totalWeight = totalWeight + 30
    
    ' �X�R�A�̐��K���i�S�����j
    If totalWeight > 0 Then
        CalculateMatchScore = (score / totalWeight) * 100
    Else
        CalculateMatchScore = 0
    End If
End Function

' ��`�Ԃ��l�������œK�Ȉ��i���̈�v����������
Private Function FindBestMatchWithPackage(ByVal searchDrug As String, ByRef targetDrugs() As String, ByVal requiredPackage As String) As String
    Dim i As Long
    Dim bestMatchScore As Double
    Dim bestMatchIndex As Long
    Dim searchParts As DrugNameParts
    
    ' �����Ώۂ̈��i���𕪉�
    searchParts = ParseDrugString(searchDrug)
    bestMatchScore = 0
    bestMatchIndex = -1
    
    ' �e��r�Ώۂɑ΂��ăX�R�A���v�Z
    For i = LBound(targetDrugs) To UBound(targetDrugs)
        Dim targetParts As DrugNameParts
        Dim currentScore As Double
        Dim hasRequiredPackage As Boolean
        
        ' ��r�Ώۂ̈��i���𕪉�
        targetParts = ParseDrugString(targetDrugs(i))
        
        ' ��`�Ԃ̊m�F
        hasRequiredPackage = (InStr(1, targetParts.Package, requiredPackage, vbTextCompare) > 0)
        
        If hasRequiredPackage Then
            ' �������A�K�i�A�P�ʂ̈�v�x���v�Z
            currentScore = CalculateMatchScore(searchParts, targetParts)
            
            If currentScore > bestMatchScore Then
                bestMatchScore = currentScore
                bestMatchIndex = i
            End If
        End If
    Next i
    
    ' ���ȏ�̃X�R�A������ꍇ�̂݌��ʂ�Ԃ�
    If bestMatchScore >= 70 And bestMatchIndex >= 0 Then ' 70%�ȏ�̈�v��
        FindBestMatchWithPackage = targetDrugs(bestMatchIndex)
    Else
        FindBestMatchWithPackage = ""
    End If
End Function

' 7�s�ڈȍ~�̈��i����r�Ɠ]�L���s���֐�
Public Sub ProcessFromRow7()
    On Error GoTo ErrorHandler
    
    ' �����ݒ�
    Application.ScreenUpdating = False
    
    ' ���[�N�V�[�g�Q�Ƃ̎擾
    Dim settingsSheet As Worksheet, targetSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1) ' �ݒ�V�[�g
    Set targetSheet = ThisWorkbook.Worksheets(2)   ' ��r�Ώۂ̃V�[�g
    
    ' ��`�Ԃ̎擾�Ɗm�F
    Dim packageType As String
    packageType = settingsSheet.Range("B4").Value
    
    ' �L���ȕ�`�Ԃ��`�F�b�N
    Dim validPackageTypes As Variant
    validPackageTypes = Array("(����`)", "���̑�(�Ȃ�)", "���", "���ܗp", "PTP", "����", "�o��", "SP", "PTP(���җp)")
    
    Dim isValidPackage As Boolean
    Dim i As Long
    isValidPackage = False
    
    For i = LBound(validPackageTypes) To UBound(validPackageTypes)
        If packageType = validPackageTypes(i) Then
            isValidPackage = True
            Exit For
        End If
    Next i
    
    If Not isValidPackage Then
        MsgBox "B4�Z���ɗL���ȕ�`�Ԃ�ݒ肵�Ă��������B" & vbCrLf & _
               "�L���Ȓl: (����`), ���̑�(�Ȃ�), ���, ���ܗp, PTP, ����, �o��, SP, PTP(���җp)", vbExclamation
        GoTo CleanExit
    End If
    
    ' �ŏI�s�̎擾
    Dim lastRowSettings As Long, lastRowTarget As Long
    lastRowSettings = settingsSheet.Cells(settingsSheet.Rows.Count, "B").End(xlUp).Row
    lastRowTarget = targetSheet.Cells(targetSheet.Rows.Count, "B").End(xlUp).Row
    
    ' ��r�Ώۖ�i����z��Ɋi�[
    Dim targetDrugs() As String
    ReDim targetDrugs(1 To lastRowTarget - 1)
    
    For i = 2 To lastRowTarget
        targetDrugs(i - 1) = targetSheet.Cells(i, "B").Value
    Next i
    
    ' ���i���̔�r�Ɠ]�L�i7�s�ڂ���J�n�j
    Dim searchDrug As String, bestMatch As String
    Dim processedCount As Long, skippedCount As Long
    processedCount = 0
    skippedCount = 0
    
    For i = 7 To lastRowSettings ' ������7�s�ڈȍ~������
        searchDrug = settingsSheet.Cells(i, "B").Value
        
        If Len(searchDrug) > 0 Then
            ' �œK�Ȉ�v������
            bestMatch = FindBestMatchingDrug(searchDrug, targetDrugs, packageType)
            
            ' ��v���錋�ʂ�����Γ]�L�A�Ȃ���΃X�L�b�v
            If Len(bestMatch) > 0 Then
                settingsSheet.Cells(i, "C").Value = bestMatch
                processedCount = processedCount + 1
            Else
                ' ��v���Ȃ��ꍇ�͉������Ȃ��i�󕶎��ŏ㏑�����Ȃ��j
                skippedCount = skippedCount + 1
            End If
        End If
    Next i
    
CleanExit:
    Application.ScreenUpdating = True
    MsgBox "�������������܂����B" & vbCrLf & _
           processedCount & "���̈��i������v���܂����B" & vbCrLf & _
           skippedCount & "���̈��i���͈�v������̂�������܂���ł����B", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' �ł���v������i������������֐�
Private Function FindBestMatchingDrug(ByVal searchDrug As String, ByRef targetDrugs() As String, ByVal packageType As String) As String
    Dim i As Long
    Dim bestMatchIndex As Long, bestMatchScore As Long, currentScore As Long
    
    bestMatchIndex = -1
    bestMatchScore = 0
    
    ' �����Ώۂ��L�[���[�h�ɕ���
    Dim keywords As Variant
    keywords = ExtractKeywords(searchDrug)
    
    ' ��`�Ԃ̓��ʏ���
    Dim skipPackageCheck As Boolean
    skipPackageCheck = (packageType = "(����`)" Or packageType = "���̑�(�Ȃ�)")
    
    ' �e��r�Ώۂɑ΂��ď���
    For i = LBound(targetDrugs) To UBound(targetDrugs)
        If Len(targetDrugs(i)) > 0 Then
            ' ��`�ԃ`�F�b�N
            Dim matchesPackage As Boolean
            
            If skipPackageCheck Then
                ' ����`�܂��͂��̑��̏ꍇ�͕�`�ԃ`�F�b�N���X�L�b�v
                matchesPackage = True
            Else
                ' ��`�Ԃ���v���邩�m�F
                matchesPackage = CheckPackage(targetDrugs(i), packageType)
            End If
            
            If matchesPackage Then
                ' �L�[���[�h��v�����v�Z
                currentScore = CalcMatchScore(keywords, targetDrugs(i))
                
                ' ��荂���X�R�A���L�^
                If currentScore > bestMatchScore Then
                    bestMatchScore = currentScore
                    bestMatchIndex = i
                End If
            End If
        End If
    Next i
    
    ' ���ʂ�Ԃ��i臒l�ȏ�̃X�R�A�̏ꍇ�̂݁j
    If bestMatchScore >= 50 And bestMatchIndex >= 0 Then
        FindBestMatchingDrug = targetDrugs(bestMatchIndex)
    Else
        FindBestMatchingDrug = ""
    End If
End Function

' ���i������L�[���[�h�𒊏o����֐�
Private Function ExtractKeywords(ByVal drugName As String) As Variant
    ' �S�p�X�y�[�X�𔼊p�ɕϊ�
    drugName = Replace(drugName, "�@", " ")
    
    ' �X�y�[�X�ŕ������Ĕz��Ɋi�[
    Dim words As Variant, result() As String
    Dim i As Long, validCount As Long
    
    words = Split(drugName, " ")
    ReDim result(UBound(words))
    validCount = 0
    
    ' ��łȂ��v�f�̂ݎ擾
    For i = 0 To UBound(words)
        If Trim(words(i)) <> "" Then
            result(validCount) = LCase(Trim(words(i)))
            validCount = validCount + 1
        End If
    Next i
    
    ' ���ʂ���̏ꍇ�̏���
    If validCount = 0 Then
        ReDim result(0)
        result(0) = LCase(Trim(drugName))
        validCount = 1
    End If
    
    ReDim Preserve result(validCount - 1)
    ExtractKeywords = result
End Function

' �L�[���[�h�̈�v�����v�Z����֐�
Private Function CalcMatchScore(ByRef keywords As Variant, ByVal targetDrug As String) As Long
    Dim i As Long, matchCount As Long
    Dim lowerTargetDrug As String
    
    lowerTargetDrug = LCase(targetDrug)
    matchCount = 0
    
    ' �e�L�[���[�h���܂܂�Ă��邩�`�F�b�N
    For i = 0 To UBound(keywords)
        If InStr(1, lowerTargetDrug, keywords(i), vbTextCompare) > 0 Then
            matchCount = matchCount + 1
        End If
    Next i
    
    ' ��v�����v�Z�i�S�����j
    If UBound(keywords) >= 0 Then
        CalcMatchScore = (matchCount * 100) / (UBound(keywords) + 1)
    Else
        CalcMatchScore = 0
    End If
End Function

' ��`�Ԃ���v���邩�`�F�b�N����֐��iCreateObject���g��Ȃ��o�[�W�����j
Private Function CheckPackage(ByVal drugName As String, ByVal packageType As String) As Boolean
    ' ��`�Ԃ̃o���G�[�V�������`
    Dim PTPVariations As Variant
    Dim BulkVariations As Variant
    Dim SPVariations As Variant
    Dim DividedVariations As Variant
    Dim SmallPackageVariations As Variant
    Dim DispensingVariations As Variant
    Dim PatientPTPVariations As Variant
    
    ' �e��`�Ԃٕ̈\�L��z��Œ�`
    PTPVariations = Array("PTP", "�o�s�o", "P.T.P.", "P.T.P")
    BulkVariations = Array("�o��", "���", "BARA", "�o����")
    SPVariations = Array("SP", "�r�o", "S.P")
    DividedVariations = Array("����", "�Ԃ�ۂ�", "����i")
    SmallPackageVariations = Array("���", "���")
    DispensingVariations = Array("���ܗp", "����", "���ܗp�")
    PatientPTPVariations = Array("PTP(���җp)", "���җpPTP", "���җp")
    
    ' ��`�Ԃɉ������ϐ���I��
    Dim variations As Variant
    
    Select Case packageType
        Case "PTP"
            variations = PTPVariations
        Case "�o��"
            variations = BulkVariations
        Case "SP"
            variations = SPVariations
        Case "����"
            variations = DividedVariations
        Case "���"
            variations = SmallPackageVariations
        Case "���ܗp"
            variations = DispensingVariations
        Case "PTP(���җp)"
            variations = PatientPTPVariations
        Case Else
            ' ��`����Ă��Ȃ��ꍇ�͕����񊮑S��v�Ŋm�F
            CheckPackage = (InStr(1, drugName, packageType, vbTextCompare) > 0)
            Exit Function
    End Select
    
    ' �e�o���G�[�V�����Ŋm�F
    Dim j As Long
    For j = LBound(variations) To UBound(variations)
        If InStr(1, drugName, variations(j), vbTextCompare) > 0 Then
            CheckPackage = True
            Exit Function
        End If
    Next j
    
    CheckPackage = False
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





