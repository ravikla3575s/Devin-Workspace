Attribute VB_Name = "StringUtils"
Option Explicit

' �u�v�ň͂܂ꂽ�e�L�X�g�𒊏o����֐��i���K�\�����g��Ȃ��o�[�W�����j
Public Function ExtractBetweenQuotes(ByVal text As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(1, text, "�u")
    If startPos > 0 Then
        endPos = InStr(startPos + 1, text, "�v")
        If endPos > startPos Then
            ExtractBetweenQuotes = Mid(text, startPos + 1, endPos - startPos - 1)
        Else
            ExtractBetweenQuotes = ""
        End If
    Else
        ExtractBetweenQuotes = ""
    End If
End Function

' �K�i�i���x�j�𒊏o����֐��i���K�\�����g��Ȃ��Łj
Public Function ExtractStrength(ByVal text As String) As String
    Dim i As Long
    Dim numStart As Long
    Dim result As String
    Dim inNumber As Boolean
    Dim units As Variant
    
    units = Array("mg", "g", "ml", "��g")
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
            ' �X�y�[�X�͋��e
        Else
            If inNumber Then
                ' �����̌�ɒP�ʂ����邩�m�F
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

' ���l�ƒP�ʂ𕪗�����֐��i���K�\�����g��Ȃ��o�[�W�����j
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
            ' �X�y�[�X�͐����ƌ��Ȃ�
        Else
            If inNumber Then
                numStr = Mid(str, numStart, i - numStart)
                unitStr = Mid(str, i)
                Exit For
            End If
        End If
    Next i
    
    ' �P�ʂ���s�v�ȕ������폜
    unitStr = Trim(unitStr)
    
    ' �P�ʂ̕W����
    If LCase(Left(unitStr, 2)) = "mg" Then
        unitStr = "mg"
    ElseIf LCase(Left(unitStr, 1)) = "g" Then
        unitStr = "g"
    ElseIf LCase(Left(unitStr, 2)) = "ml" Then
        unitStr = "ml"
    ElseIf LCase(Left(unitStr, 2)) = "��g" Then
        unitStr = "��g"
    End If
    
    ' ���ʂ�ݒ�
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

' �K�i�i���x�j���r����֐�
Public Function CompareStrength(ByVal str1 As String, ByVal str2 As String) As Boolean
    ' ���l�ƒP�ʂ𕪗����Ĕ�r
    Dim num1 As Double, num2 As Double
    Dim unit1 As String, unit2 As String
    
    ' ���l�ƒP�ʂ𒊏o
    ExtractNumberAndUnit str1, num1, unit1
    ExtractNumberAndUnit str2, num2, unit2
    
    ' ���l�ƒP�ʂ�������v����ꍇ�̂�True
    CompareStrength = (num1 = num2) And (StrComp(unit1, unit2, vbTextCompare) = 0)
End Function

' B4�Z���ɕ�`�Ԃ̑I�������h���b�v�_�E�����X�g�Ƃ��Đݒ肷��֐�
Public Sub SetupPackageTypeDropdown()
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' B4�Z���Ƀh���b�v�_�E�����X�g��ݒ�
    With settingsSheet.Range("B4").Validation
        .Delete ' �����̓��͋K�����폜
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="(����`),���̑�(�Ȃ�),���,���ܗp,PTP,����,�o��,SP,PTP(���җp)"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "��`�Ԃ̑I��"
        .ErrorTitle = "�����ȑI��"
        .InputMessage = "���X�g�����`�Ԃ�I�����Ă�������"
        .ErrorMessage = "���X�g����L���ȕ�`�Ԃ�I�����Ă�������"
    End With
    
    ' B4�Z���̏����ݒ�
    With settingsSheet.Range("B4")
        .Value = "PTP" ' �f�t�H���g�l��ݒ�
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242) ' �����F�̔w�i
    End With
    
    ' A4�Z���Ƀ��x����ݒ�
    With settingsSheet.Range("A4")
        .Value = "��`��:"
        .Font.Bold = True
    End With
    
    ' B3�Z���Ƀ^�C�g����ݒ�
    With settingsSheet.Range("A1:C1")
        .Merge
        .Value = "���i����r�c�[��"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(180, 198, 231) ' �F�̔w�i
    End With
    
    ' ��w�b�_�[��ݒ�
    settingsSheet.Range("A6").Value = "No."
    settingsSheet.Range("B6").Value = "�������i��"
    settingsSheet.Range("C6").Value = "��v���i��"
    
    With settingsSheet.Range("A6:C6")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247) ' �����F�̔w�i
    End With
    
    ' �񕝂𒲐�
    settingsSheet.Columns("A").ColumnWidth = 5
    settingsSheet.Columns("B").ColumnWidth = 30
    settingsSheet.Columns("C").ColumnWidth = 40
    
    ' �s�ԍ���ݒ�i7�s�ڂ���30�s�ڂ܂Łj
    Dim i As Long
    For i = 7 To 30
        settingsSheet.Cells(i, "A").Value = i - 6
    Next i
    
    MsgBox "��`�Ԃ̃h���b�v�_�E�����X�g��ݒ肵�܂����B", vbInformation
End Sub




