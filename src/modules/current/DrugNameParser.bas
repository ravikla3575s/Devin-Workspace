Attribute VB_Name = "DrugNameParser"
Option Explicit

' ��i���̍\����
Public Type DrugNameParts
    BaseName As String
    formType As String
    strength As String
    maker As String
    Package As String
End Type

' ��i������͂��č\���̂ɕ�������֐�
Public Function ParseDrugString(ByVal drugStr As String) As DrugNameParts
    Dim result As DrugNameParts
    Dim tempStr As String
    
    ' �S�p�����𔼊p�ɕϊ�
    tempStr = StrConv(drugStr, vbNarrow)
    
    ' ���[�J�[���𒊏o (�u�v��)
    Dim makerMatch As String
    makerMatch = ExtractBetweenQuotes(tempStr)
    result.maker = makerMatch
    
    ' �K�i�𒊏o (����+�P��)
    Dim strengthMatch As String
    strengthMatch = ExtractStrengthSimple(tempStr)
    result.strength = strengthMatch
    
    ' �܌^�𒊏o
    Dim formMatch As String
    formMatch = ExtractFormTypeSimple(tempStr)
    result.formType = formMatch
    
    ' ��`�Ԃ𒊏o
    result.Package = ExtractPackageTypeSimple(tempStr)
    
    ' ��{���𒊏o�i���[�J�[���ƋK�i�̑O�܂Łj
    result.BaseName = ExtractBaseNameSimple(tempStr, result.maker, result.strength, result.formType)
    
    ParseDrugString = result
End Function

' ��i���̊�{�����𒊏o����֐��i���K�\�����g��Ȃ��o�[�W�����j
Public Function ExtractBaseNameSimple(ByVal text As String, _
                                    ByVal maker As String, _
                                    ByVal strength As String, _
                                    ByVal formType As String) As String
    Dim result As String
    result = text
    
    ' ���[�J�[��������
    If maker <> "" Then
        result = Replace(result, "�u" & maker & "�v", "")
    End If
    
    ' �K�i������
    If strength <> "" Then
        result = Replace(result, strength, "")
    End If
    
    ' �܌^������
    If formType <> "" Then
        result = Replace(result, formType, "")
    End If
    
    ' ���ʕ\���������i��F10���j- ���K�\�����g��Ȃ��o�[�W����
    Dim i As Long
    Dim parts() As String
    parts = Split(result, " ")
    
    For i = 0 To UBound(parts)
        If IsNumericWithSuffix(parts(i)) Then
            parts(i) = ""
        End If
    Next i
    
    result = Join(parts, " ")
    
    ' ���ꕶ���Ɨ]���ȋ󔒂�����
    result = Replace(result, "�@", " ")  ' �S�p�X�y�[�X�𔼊p��
    result = Trim(result)
    
    ' �A������X�y�[�X��P��̃X�y�[�X�ɒu��
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    ExtractBaseNameSimple = result
End Function

' ����+�P�ʂ��ǂ������`�F�b�N����i��F10���j
Private Function IsNumericWithSuffix(ByVal text As String) As Boolean
    Dim i As Long
    Dim hasDigit As Boolean
    Dim hasSuffix As Boolean
    
    hasDigit = False
    
    For i = 1 To Len(text)
        If IsNumeric(Mid(text, i, 1)) Then
            hasDigit = True
        End If
    Next i
    
    ' �P�ʂ̃��X�g
    Dim units As Variant
    units = Array("��", "�J�v�Z��", "��", "��", "��", "�{", "��", "��", "�g", "��")
    
    hasSuffix = False
    For i = 0 To UBound(units)
        If InStr(text, units(i)) > 0 Then
            hasSuffix = True
            Exit For
        End If
    Next i
    
    IsNumericWithSuffix = hasDigit And hasSuffix
End Function

' �K�i�i���x�j�𒊏o����֐��i���K�\�����g��Ȃ��o�[�W�����j
Public Function ExtractStrengthSimple(ByVal text As String) As String
    Dim i As Long, j As Long
    Dim numStart As Long
    Dim result As String
    Dim inNumber As Boolean
    Dim units As Variant
    
    units = Array("mg", "g", "ml", "��g")
    inNumber = False
    numStart = 0
    result = ""
    
    For i = 1 To Len(text)
        Dim c As String
        c = Mid(text, i, 1)
        
        If IsNumeric(c) Or c = "." Then
            If Not inNumber Then
                inNumber = True
                numStart = i
            End If
        ElseIf c = " " And inNumber Then
            ' �X�y�[�X�͋��e
        Else
            If inNumber Then
                ' �����̌�ɒP�ʂ����邩�m�F
                For j = 0 To UBound(units)
                    If LCase(Mid(text, i, Len(units(j)))) = LCase(units(j)) Then
                        result = Mid(text, numStart, i - numStart + Len(units(j)))
                        Exit For
                    End If
                Next j
                
                If result <> "" Then
                    Exit For
                End If
                
                inNumber = False
            End If
        End If
    Next i
    
    ExtractStrengthSimple = result
End Function

' �܌^�𒊏o����֐��i���K�\�����g��Ȃ��o�[�W�����j
Public Function ExtractFormTypeSimple(ByVal text As String) As String
    Dim forms As Variant
    Dim i As Long
    
    forms = Array("��", "�J�v�Z��", "�ח�", "����", "�U", "�V���b�v", "�h���C�V���b�v", _
                  "���ˉt", "���˗p", "��p", "�N���[��", "�Q��", "�e�[�v", "�p�b�v", "�_��t")
    
    For i = 0 To UBound(forms)
        If InStr(text, forms(i)) > 0 Then
            ExtractFormTypeSimple = forms(i)
            Exit Function
        End If
    Next i
    
    ExtractFormTypeSimple = ""
End Function

' ��`�Ԃ𒊏o����֐��i���K�\�����g��Ȃ��o�[�W�����j
Public Function ExtractPackageTypeSimple(ByVal text As String) As String
    Dim packages As Variant
    Dim i As Long
    
    packages = Array("(����`)", "���̑�(�Ȃ�)", "���", "���ܗp", "PTP", "����", "�o��", "SP", "PTP(���җp)")
    
    For i = 0 To UBound(packages)
        If InStr(1, text, packages(i), vbTextCompare) > 0 Then
            ' ����������`�Ԃ����̂܂ܕԂ��iNormalizePackageType�͎g��Ȃ��j
            ExtractPackageTypeSimple = packages(i)
            Exit Function
        End If
    Next i
    
    ' �X���b�V���ň͂܂ꂽ����������
    Dim startPos As Long, endPos As Long
    startPos = InStr(1, text, "/")
    If startPos > 0 Then
        endPos = InStr(startPos + 1, text, "/")
        If endPos > startPos Then
            ' �X���b�V���Ԃ̕���������̂܂ܕԂ��iNormalizePackageType�͎g��Ȃ��j
            ExtractPackageTypeSimple = Mid(text, startPos + 1, endPos - startPos - 1)
            Exit Function
        End If
    End If
    
    ExtractPackageTypeSimple = ""
End Function

' �p�b�P�[�W�^�C�v�擾�i�X���b�V���Ԃ̕�����j
Public Function GetPackageType(ByVal text As String) As String
    Dim startPos As Long, endPos As Long
    
    startPos = InStr(1, text, "/")
    If startPos > 0 Then
        endPos = InStr(startPos + 1, text, "/")
        If endPos > startPos Then
            GetPackageType = Mid(text, startPos + 1, endPos - startPos - 1)
        Else
            GetPackageType = ""
        End If
    Else
        GetPackageType = ""
    End If
End Function

' ��i���̔�r�֐�
Public Function CompareDrugStringsWithRate(ByVal sourceStr As String, ByVal targetStr As String) As Double
    Dim sourceParts As DrugNameParts
    Dim targetParts As DrugNameParts
    Dim matchCount As Integer
    Dim totalItems As Integer
    
    sourceParts = ParseDrugString(sourceStr)
    targetParts = ParseDrugString(targetStr)
    
    totalItems = 5 ' ��{���A�܌^�A�K�i�A���[�J�[�A���5����
    matchCount = 0
    
    ' ��{���̔�r�i���S��v�j
    If StrComp(sourceParts.BaseName, targetParts.BaseName, vbTextCompare) = 0 Then
        matchCount = matchCount + 1
    End If
    
    ' �܌^�̔�r�i���S��v�j
    If StrComp(sourceParts.formType, targetParts.formType, vbTextCompare) = 0 Then
        matchCount = matchCount + 1
    End If
    
    ' �K�i�̔�r�i���l�ƒP�ʂ𐳋K�����Ĕ�r�j
    If CompareStrength(sourceParts.strength, targetParts.strength) Then
        matchCount = matchCount + 1
    End If
    
    ' ���[�J�[���̔�r�i���S��v�j
    If StrComp(sourceParts.maker, targetParts.maker, vbTextCompare) = 0 Then
        matchCount = matchCount + 1
    End If
    
    ' ��`�Ԃ̔�r�i������x�̗h������e�j
    ' ComparePackageType�֐��̑���ɒP���ȕ������r���g�p
    If StrComp(sourceParts.Package, targetParts.Package, vbTextCompare) = 0 Then
        matchCount = matchCount + 1
    End If
    
    ' ��v�����v�Z�i�S�����j
    CompareDrugStringsWithRate = (matchCount / totalItems) * 100
End Function





