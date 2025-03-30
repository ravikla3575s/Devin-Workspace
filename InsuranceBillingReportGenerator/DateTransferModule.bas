Attribute VB_Name = "DateTransferModule"
Option Explicit

' ������̒萔��`
Private Const BILLING_SHAHO As String = "�Е�"
Private Const BILLING_KOKUHO As String = "����"

' ���Z�v�g�󋵂̒萔��`
Private Const STATUS_UNCLAIMED As Long = 1    ' ������
Private Const STATUS_RECLAIM As Long = 2      ' �Đ���
Private Const STATUS_RETURN As Long = 3       ' �Ԗ�
Private Const STATUS_ADJUSTMENT As Long = 4    ' ��������

' �e�󋵂̊J�n�s
Private Type StartRows
    Unclaimed As Long    ' �������J�n�s
    Reclaim As Long      ' �Đ����J�n�s
    Return As Long       ' �ԖߊJ�n�s
    Adjustment As Long   ' ��������J�n�s
End Type

' �����悲�Ƃ̃��[�N�V�[�g��
Private Const WS_SHAHO As String = "�Еۖ������ꗗ"
Private Const WS_KOKUHO As String = "���ۖ������ꗗ"

' ���C�������֐�
Private Function ProcessBillingData(ByVal dispensing_year As Integer, ByVal dispensing_month As Integer, _
                                  ByVal status As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' �ЕہE���ۂ��ꂼ��̔z���������
    Dim shahoData() As Variant
    Dim kuhoData() As Variant
    ReDim shahoData(1 To 8, 1 To 1)
    ReDim kuhoData(1 To 8, 1 To 1)
    
    ' �J�E���^�[������
    Dim shahoCount As Long: shahoCount = 1
    Dim kuhoCount As Long: kuhoCount = 1
    
    ' �J�n�s�̎擾
    Dim shahoStartRows As StartRows
    Dim kuhoStartRows As StartRows
    Call InitializeStartRows(shahoStartRows, kuhoStartRows)
    
    ' �t�H�[������
    Dim billing_form As New UnclaimedBillingForm
    Dim continue_input As Boolean
    continue_input = True
    
    Do While continue_input
        billing_form.SetDispensingDate dispensing_year, dispensing_month
        billing_form.Show
        
        If Not billing_form.DialogResult Then
            If shahoCount = 1 And kuhoCount = 1 Then
                ' �f�[�^�����͂ŃL�����Z��
                ProcessBillingData = True
                Exit Function
            Else
                ' �����f�[�^������ꍇ�͊m�F
                If MsgBox("���͍ς݂̃f�[�^��j�����Ă�낵���ł����H", vbYesNo + vbQuestion) = vbYes Then
                    Exit Do
                End If
            End If
        Else
            ' ������ɉ����ēK�؂Ȕz��Ɋi�[
            If billing_form.BillingDestination = BILLING_SHAHO Then
                ' �Е۔z��̊g���`�F�b�N
                If shahoCount > UBound(shahoData, 2) Then
                    ReDim Preserve shahoData(1 To 8, 1 To shahoCount)
                End If
                Call StoreDataInArray(shahoData, shahoCount, billing_form, dispensing_year, dispensing_month)
                shahoCount = shahoCount + 1
            Else
                ' ���۔z��̊g���`�F�b�N
                If kuhoCount > UBound(kuhoData, 2) Then
                    ReDim Preserve kuhoData(1 To 8, 1 To kuhoCount)
                End If
                Call StoreDataInArray(kuhoData, kuhoCount, billing_form, dispensing_year, dispensing_month)
                kuhoCount = kuhoCount + 1
            End If
            
            continue_input = billing_form.ContinueInput
        End If
    Loop
    
    ' �f�[�^�̓]�L����
    If shahoCount > 1 Then
        Call WriteDataToWorksheet(shahoData, shahoCount - 1, WS_SHAHO, GetStartRow(shahoStartRows, status))
    End If
    
    If kuhoCount > 1 Then
        Call WriteDataToWorksheet(kuhoData, kuhoCount - 1, WS_KOKUHO, GetStartRow(kuhoStartRows, status))
    End If
    
    ProcessBillingData = True
    Exit Function
    
ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical
    ProcessBillingData = False
End Function

' �J�n�s�̏�����
Private Sub InitializeStartRows(ByRef shahoRows As StartRows, ByRef kuhoRows As StartRows)
    ' �Еۂ̊J�n�s
    With shahoRows
        .Unclaimed = 2      ' �������J�n�s
        .Reclaim = 8        ' �Đ����J�n�s
        .Return = 14        ' �ԖߊJ�n�s
        .Adjustment = 20    ' ��������J�n�s
    End With
    
    ' ���ۂ̊J�n�s
    With kuhoRows
        .Unclaimed = 2
        .Reclaim = 8
        .Return = 14
        .Adjustment = 20
    End With
End Sub

' ��Ԃɉ������J�n�s�̎擾
Private Function GetStartRow(ByRef rows As StartRows, ByVal status As Long) As Long
    Select Case status
        Case STATUS_UNCLAIMED
            GetStartRow = rows.Unclaimed
        Case STATUS_RECLAIM
            GetStartRow = rows.Reclaim
        Case STATUS_RETURN
            GetStartRow = rows.Return
        Case STATUS_ADJUSTMENT
            GetStartRow = rows.Adjustment
    End Select
End Function

' �z��ւ̃f�[�^�i�[
Private Sub StoreDataInArray(ByRef dataArray() As Variant, ByVal CurrentIndex As Long, _
                           ByVal form As UnclaimedBillingForm, ByVal year As Integer, ByVal month As Integer)
    With form
        dataArray(1, CurrentIndex) = .PatientName
        dataArray(2, CurrentIndex) = "R" & year & "." & Format(month, "00")
        dataArray(3, CurrentIndex) = .MedicalInstitution
        dataArray(4, CurrentIndex) = .UnclaimedReason
        dataArray(5, CurrentIndex) = .BillingDestination
        dataArray(6, CurrentIndex) = .InsuranceRatio
        dataArray(7, CurrentIndex) = .BillingPoints
        dataArray(8, CurrentIndex) = .Remarks
    End With
End Sub

' ���[�N�V�[�g�ւ̃f�[�^�]�L
Private Sub WriteDataToWorksheet(ByRef dataArray() As Variant, ByVal dataCount As Long, _
                               ByVal wsName As String, ByVal startRow As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(wsName)
    
    ' ���݂̍s�����m�F
    Dim currentRows As Long
    currentRows = ws.Range("A" & startRow).End(xlDown).row - startRow + 1
    
    ' 5�s�ȏ�̃f�[�^������ꍇ�A�s��ǉ�
    If currentRows >= 5 Then
        ws.rows(startRow + 5).Resize(dataCount).Insert Shift:=xlDown
    End If
    
    ' �f�[�^�̓]�L
    With ws
        .Range(.Cells(startRow, 1), .Cells(startRow + dataCount - 1, 8)).value = _
            WorksheetFunction.Transpose(WorksheetFunction.Transpose(dataArray))
        
        ' �����ݒ�
        .Range(.Cells(startRow, 1), .Cells(startRow + dataCount - 1, 8)).Borders.LineStyle = xlContinuous
    End With
End Sub

Sub ImportCsvData(csv_file_path As String, ws As Worksheet, file_type As String, Optional check_status As Boolean = False)
    Dim file_system_local As Object, text_stream As Object
    Dim column_map As Object
    Dim line_text As String
    Dim data_array As Variant
    Dim row_index As Long, col_index As Long
    Dim key As Variant

    On Error GoTo ImportError
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set file_system_local = CreateObject("Scripting.FileSystemObject")
    Set text_stream = file_system_local.OpenTextFile(csv_file_path, 1, False, -2)
    Set column_map = GetColumnMapping(file_type)

    ' �w�b�_�s���쐬
    ws.Cells.Clear
    col_index = 1
    For Each key In column_map.Keys
        ws.Cells(1, col_index).value = column_map(key)
        col_index = col_index + 1
    Next key

    ' CSV�t�@�C����ǂݍ��݁A�f�[�^������]�L
    row_index = 2  ' �f�[�^��2�s�ڂ���J�n
    
    ' CSV��1�s�ڂ�2�s�ځi�w�b�_�[�j��ǂݔ�΂�
    If Not text_stream.AtEndOfStream Then
        text_stream.SkipLine  ' 1�s�ڂ��X�L�b�v
        If Not text_stream.AtEndOfStream Then
            text_stream.SkipLine  ' 2�s�ڂ��X�L�b�v
        End If
    End If
    
    ' �c��̃f�[�^��]�L
    Do While Not text_stream.AtEndOfStream
        line_text = text_stream.ReadLine
        data_array = Split(line_text, ",")
        
        ' �����m��󋵂̃`�F�b�N�icheck_status��True�̏ꍇ�j
        Dim should_transfer As Boolean
        should_transfer = True
        
        If check_status Then
            ' �����m��󋵂�30��ځi�C���f�b�N�X29�j�ɂ���
            If UBound(data_array) >= 29 Then
                ' �����m��󋵂�1�ȊO�̏ꍇ�ɓ]�L
                should_transfer = (Trim(data_array(29)) <> "1")
                
                ' �f�o�b�O�o�͂�ǉ�
                Debug.Print "Row " & row_index & " status: " & Trim(data_array(29)) & _
                          ", Transfer: " & should_transfer
            End If
        End If
        
        If should_transfer Then
            col_index = 1
            For Each key In column_map.Keys
                If key - 1 <= UBound(data_array) Then
                    ws.Cells(row_index, col_index).value = Trim(data_array(key - 1))
                End If
                col_index = col_index + 1
            Next key
            row_index = row_index + 1
        End If
    Loop
    text_stream.Close

    ws.Cells.EntireColumn.AutoFit

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
ImportError:
    MsgBox "CSV�f�[�^�Ǎ����ɃG���[���������܂���: " & Err.Description, vbCritical, "�G���["
    If Not text_stream Is Nothing Then text_stream.Close
    Set text_stream = Nothing
    Set file_system_local = Nothing
    Set column_map = Nothing
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical
End Sub

Function GetColumnMapping(file_type As String) As Object
    Dim column_map As Object
    Set column_map = CreateObject("Scripting.Dictionary")
    Dim k As Integer

    Select Case file_type
        Case "�U���z���׏�"
            column_map.Add 2, "�f�Ái���܁j�N��"
            column_map.Add 5, "��t�ԍ�"
            column_map.Add 14, "����"
            column_map.Add 16, "���N����"
            column_map.Add 22, "��Õی�_�����_��"
            column_map.Add 23, "��Õی�_����_��"
            column_map.Add 24, "��Õی�_�ꕔ���S��"
            column_map.Add 25, "��Õی�_���z"
            ' ��1�`��5����i�e10��Ԋu: �����_���E����_���E���ҕ��S���E���z�j
            For k = 1 To 5
                column_map.Add 33 + (k - 1) * 10, "��" & k & "����_�����_��"
                column_map.Add 34 + (k - 1) * 10, "��" & k & "����_����_��"
                column_map.Add 35 + (k - 1) * 10, "��" & k & "����_���ҕ��S��"
                column_map.Add 36 + (k - 1) * 10, "��" & k & "����_���z"
            Next k
            column_map.Add 82, "�Z��z���v"
        Case "�����m���"
            ' �����m��CSV�ifixf�f�[�^�j�̗�Ή�
            column_map.Add 4, "�f�Ái���܁j�N��"
            column_map.Add 5, "����"
            column_map.Add 7, "���N����"
            column_map.Add 9, "��Ë@�֖���"
            column_map.Add 13, "�����v�_��"
            For k = 1 To 4
                column_map.Add 16 + (k - 1) * 3, "��" & k & "����_�����_��"
            Next k
            column_map.Add 30, "�����m���"
            column_map.Add 31, "�G���[�敪"
        Case "�����_�A����"
            column_map.Add 2, "���ܔN��"
            column_map.Add 4, "��t�ԍ�"
            column_map.Add 11, "�敪"
            column_map.Add 14, "�V�l���Ƌ敪"
            column_map.Add 15, "����"
            column_map.Add 21, "�����_��(���z)"
            column_map.Add 22, "���R"
        Case "�Ԗߓ���"
            column_map.Add 2, "���ܔN��(YYMM)"
            column_map.Add 3, "��t�ԍ�"
            column_map.Add 4, "�ی��Ҕԍ�"
            column_map.Add 7, "����"
            column_map.Add 9, "�����_��"
            column_map.Add 10, "��܈ꕔ���S��"
            column_map.Add 12, "�ꕔ���S���z"
            column_map.Add 13, "����S���z"
            column_map.Add 14, "���R�R�[�h"
        Case Else
            ' ���̑��̃f�[�^��ʂ�����Βǉ�
    End Select

    Set GetColumnMapping = column_map
End Function

Sub TransferBillingDetails(report_wb As Workbook, csv_file_name As String, dispensing_year As String, _
                         dispensing_month As String, Optional check_status As Boolean = False)
    On Error GoTo ErrorHandler
    
    Dim ws_main As Worksheet, ws_details As Worksheet
    Dim csv_yymm As String
    Dim payer_type As String
    Dim start_row_dict As Object
    Dim rebill_dict As Object, late_dict As Object, unpaid_dict As Object, assessment_dict As Object
    
    ' �ەt�����̌����擾
    Dim details_sheet_name As String
    details_sheet_name = UtilityModule.ConvertToCircledNumber(CInt(dispensing_month))
    
    Debug.Print "Looking for details sheet: " & details_sheet_name
    
    ' �ڍ׃V�[�g�̑��݊m�F
    On Error Resume Next
    Set ws_details = report_wb.Sheets(details_sheet_name)
    On Error GoTo ErrorHandler
    
    If ws_details Is Nothing Then
        MsgBox "�ڍ׃V�[�g '" & details_sheet_name & "' ��������܂���B", vbExclamation, "�G���["
        Exit Sub
    End If
    
    ' ���C���V�[�g�͑��݊m�F�����Ɏ擾
    Set ws_main = report_wb.Sheets(1)
    
    ' ���ܔN���Ɛ�����敪�̎擾
    csv_yymm = GetDispenseYearMonth(ws_main)
    payer_type = GetPayerType(csv_file_name)
    
    If payer_type = "�J��" Then
        Debug.Print "�J�Ѓf�[�^�̂��߁A�������X�L�b�v���܂��B"
        Exit Sub
    End If
    
    ' �ڍ׃V�[�g��̊e�J�e�S���J�n�s���擾
    Set start_row_dict = UtilityModule.GetCategoryStartRows(ws_details, payer_type)
    
    If start_row_dict.count = 0 Then
        Debug.Print "WARNING: �J�e�S���̊J�n�s��������܂���: " & payer_type
        Exit Sub
    End If
    
    ' �f�[�^�̕��ނƎ����̍쐬
    Set rebill_dict = CreateObject("Scripting.Dictionary")
    Set late_dict = CreateObject("Scripting.Dictionary")
    Set unpaid_dict = CreateObject("Scripting.Dictionary")
    Set assessment_dict = CreateObject("Scripting.Dictionary")
    
    ' ���C���V�[�g�̃f�[�^�𕪗�
    If check_status Then
        Call ClassifyMainSheetDataWithStatus(ws_main, csv_yymm, csv_file_name, _
                                           rebill_dict, late_dict, unpaid_dict, assessment_dict)
    Else
        Call ClassifyMainSheetData(ws_main, csv_yymm, csv_file_name, _
                                 rebill_dict, late_dict, unpaid_dict, assessment_dict)
    End If
    
    ' �s�̒ǉ�����
    Call InsertAdditionalRows(ws_details, start_row_dict, rebill_dict.count, late_dict.count, assessment_dict.count)
    
    ' �f�[�^�̓]�L
    Call WriteDataToDetails(ws_details, start_row_dict, rebill_dict, late_dict, unpaid_dict, assessment_dict, payer_type)
    
    ' FIXF�t�@�C���̏ꍇ�A���������Z�v�g�̊m�F�i�ڍ׃V�[�g��n���j
    If InStr(LCase(csv_file_name), "fixf") > 0 Then
        Call CheckAndRegisterUnclaimedBilling(CInt(dispensing_year), CInt(dispensing_month), ws_details)
    End If
    
    Exit Sub

ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in TransferBillingDetails"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "Details sheet name: " & details_sheet_name
    Debug.Print "File name: " & csv_file_name
    Debug.Print "Payer type: " & payer_type
    Debug.Print "=================================="
    
    MsgBox "�f�[�^�]�L���ɃG���[���������܂����B" & vbCrLf & _
           "�G���[�ԍ�: " & Err.Number & vbCrLf & _
           "�G���[���e: " & Err.Description & vbCrLf & _
           "�ڍ׃V�[�g: " & details_sheet_name, _
           vbCritical, "�G���["
End Sub

Private Function GetDispenseYearMonth(ws As Worksheet) As String
    GetDispenseYearMonth = ""
    If ws.Cells(2, 2).value <> "" Then
        GetDispenseYearMonth = Right(CStr(ws.Cells(2, 2).value), 4)
        If InStr(GetDispenseYearMonth, "�N") > 0 Or InStr(GetDispenseYearMonth, "��") > 0 Then
            GetDispenseYearMonth = Replace(Replace(GetDispenseYearMonth, "�N", ""), "��", "")
        End If
    End If
End Function

Private Function GetPayerType(csv_file_name As String) As String
    Dim base_name As String, payer_code As String
    
    base_name = csv_file_name
    If InStr(base_name, ".") > 0 Then base_name = Left(base_name, InStrRev(base_name, ".") - 1)
    
    If Len(base_name) >= 7 Then
        payer_code = Mid(base_name, 7, 1)
    Else
        payer_code = ""
    End If
    
    Select Case payer_code
        Case "1": GetPayerType = "�Е�"
        Case "2": GetPayerType = "����"
        Case Else: GetPayerType = "�J��"
    End Select
End Function

Private Sub ClassifyMainSheetData(ws As Worksheet, csv_yymm As String, csv_file_name As String, _
    ByRef rebill_dict As Object, ByRef late_dict As Object, ByRef unpaid_dict As Object, ByRef assessment_dict As Object)
    
    Dim last_row As Long, row As Long
    Dim dispensing_code As String, dispensing_ym As String
    Dim row_data As Variant
    
    last_row = ws.Cells(ws.rows.count, "D").End(xlUp).row
    
    For row = 2 To last_row
        dispensing_code = ws.Cells(row, 2).value
        dispensing_ym = DateConversionModule.ConvertToWesternDate(dispensing_code)
        
        If csv_yymm <> "" And Right(dispensing_code, 4) <> csv_yymm Then
            row_data = Array(ws.Cells(row, 4).value, dispensing_ym, ws.Cells(row, 5).value, ws.Cells(row, 10).value)
            
            If InStr(LCase(csv_file_name), "fixf") > 0 Then
                late_dict(ws.Cells(row, 1).value) = row_data
            ElseIf InStr(LCase(csv_file_name), "fmei") > 0 Then
                rebill_dict(ws.Cells(row, 1).value) = row_data
            ElseIf InStr(LCase(csv_file_name), "zogn") > 0 Then
                unpaid_dict(ws.Cells(row, 1).value) = row_data
            ElseIf InStr(LCase(csv_file_name), "henr") > 0 Then
                assessment_dict(ws.Cells(row, 1).value) = row_data
            End If
        End If
    Next row
End Sub

Private Sub ClassifyMainSheetDataWithStatus(ws As Worksheet, csv_yymm As String, csv_file_name As String, _
    ByRef rebill_dict As Object, ByRef late_dict As Object, ByRef unpaid_dict As Object, ByRef assessment_dict As Object)
    
    Dim last_row As Long, row As Long
    Dim dispensing_code As String, dispensing_ym As String
    Dim row_data As Variant
    
    last_row = ws.Cells(ws.rows.count, "D").End(xlUp).row
    
    For row = 2 To last_row
        ' �����m��󋵂��`�F�b�N�iAD�� = 30��ځj
        If ws.Cells(row, 30).value = "2" Then
            dispensing_code = ws.Cells(row, 2).value
            dispensing_ym = UtilityModule.ConvertToWesternDate(dispensing_code)
            
            If csv_yymm <> "" And Right(dispensing_code, 4) <> csv_yymm Then
                row_data = Array(ws.Cells(row, 4).value, dispensing_ym, ws.Cells(row, 5).value, ws.Cells(row, 10).value)
                
                If InStr(LCase(csv_file_name), "fixf") > 0 Then
                    late_dict(ws.Cells(row, 1).value) = row_data
                ElseIf InStr(LCase(csv_file_name), "fmei") > 0 Then
                    rebill_dict(ws.Cells(row, 1).value) = row_data
                ElseIf InStr(LCase(csv_file_name), "zogn") > 0 Then
                    unpaid_dict(ws.Cells(row, 1).value) = row_data
                ElseIf InStr(LCase(csv_file_name), "henr") > 0 Then
                    assessment_dict(ws.Cells(row, 1).value) = row_data
                End If
            End If
        End If
    Next row
End Sub

Private Type UnclaimedRecord
    PatientName As String
    DispensingDate As String
    MedicalInstitution As String
    UnclaimedReason As String
    BillingDestination As String
    InsuranceRatio As Integer
    BillingPoints As Long
    Remarks As String
End Type

Private Function CheckAndRegisterUnclaimedBilling(ByVal dispensing_year As Integer, ByVal dispensing_month As Integer, _
                                            Optional ByVal ws_details As Worksheet = Nothing) As Boolean
    On Error GoTo ErrorHandler
    
    Dim response As VbMsgBoxResult
    response = MsgBox("���������Z�v�g�̓��͂��J�n���܂����H", vbYesNo + vbQuestion)
    
    If response = vbYes Then
        ' ���������Z�v�g�f�[�^���i�[����񎟌��z��
        Dim unclaimedData() As Variant
        ReDim unclaimedData(1 To 8, 1 To 1)
        Dim currentColumn As Long
        currentColumn = 1
        
        Dim unclaimed_form As New UnclaimedBillingForm
        Dim continue_input As Boolean
        continue_input = True
        
        Do While continue_input
            ' ���ܔN����ݒ�
            unclaimed_form.SetDispensingDate dispensing_year, dispensing_month
            
            ' �ҏW���[�h�̏ꍇ�A�f�[�^�����[�h
            If unclaimed_form.CurrentIndex < currentColumn Then
                unclaimed_form.LoadData unclaimedData, unclaimed_form.CurrentIndex
            End If
            
            unclaimed_form.Show
            
            If Not unclaimed_form.DialogResult Then
                ' �L�����Z���{�^���������ꂽ�ꍇ
                If currentColumn = 1 Then
                    ' �f�[�^�����͂ŃL�����Z��
                    CheckAndRegisterUnclaimedBilling = True
                    Exit Function
                Else
                    ' �����f�[�^������ꍇ�͊m�F
                    If MsgBox("���͍ς݂̃f�[�^��j�����Ă�낵���ł����H", vbYesNo + vbQuestion) = vbYes Then
                        Exit Do
                    End If
                End If
            Else
                ' �z��̃T�C�Y���g���i�K�v�ȏꍇ�j
                If currentColumn > UBound(unclaimedData, 2) Then
                    ReDim Preserve unclaimedData(1 To 8, 1 To currentColumn)
                End If
                
                ' �f�[�^��z��Ɋi�[
                With unclaimed_form
                    unclaimedData(1, currentColumn) = .PatientName
                    unclaimedData(2, currentColumn) = "R" & dispensing_year & "." & Format(dispensing_month, "00")
                    unclaimedData(3, currentColumn) = .MedicalInstitution
                    unclaimedData(4, currentColumn) = .UnclaimedReason
                    unclaimedData(5, currentColumn) = .BillingDestination
                    unclaimedData(6, currentColumn) = .InsuranceRatio
                    unclaimedData(7, currentColumn) = .BillingPoints
                    unclaimedData(8, currentColumn) = .Remarks
                End With
                
                If .ContinueInput Then
                    ' ���փ{�^���������ꂽ�ꍇ
                    currentColumn = currentColumn + 1
                    continue_input = True
                Else
                    ' �����{�^���������ꂽ�ꍇ
                    continue_input = False
                End If
            End If
        Loop
        
        ' �f�[�^��1���ȏ���͂���Ă���ꍇ�̂݁AExcel�ɓ]�L
        If currentColumn > 0 Then
            If ws_details Is Nothing Then
                Set ws_details = ThisWorkbook.Worksheets("�������ꗗ")
            End If
            
            ' �ŏI�s�̎擾
            Dim lastRow As Long
            lastRow = ws_details.Cells(ws_details.rows.count, "A").End(xlUp).row
            
            ' �f�[�^�̓]�L
            With ws_details
                .Range(.Cells(lastRow + 1, 1), .Cells(lastRow + currentColumn, 8)).value = _
                    WorksheetFunction.Transpose(WorksheetFunction.Transpose(unclaimedData))
                
                ' �����ݒ�
                .Range(.Cells(lastRow + 1, 1), .Cells(lastRow + currentColumn, 8)).Borders.LineStyle = xlContinuous
            End With
        End If
    End If
    
    CheckAndRegisterUnclaimedBilling = True
    Exit Function

ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical
    CheckAndRegisterUnclaimedBilling = False
End Function

Private Sub InsertAdditionalRows(ws As Worksheet, start_row_dict As Object, rebill_count As Long, late_count As Long, assessment_count As Long)
    Dim ws_details As Worksheet
    Set ws_details = ws
    
    Dim row_index As Long
    Dim start_row As Long
    Dim end_row As Long
    Dim i_key As Long
    
    ' �e�J�e�S���̊J�n�s���擾
    For Each key In start_row_dict.Keys
        start_row = start_row_dict(key)
        end_row = start_row + 1
        
        ' �s�̒ǉ�
        ws_details.rows(end_row).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        ws_details.Cells(end_row, 1).value = key
        
        ' �f�[�^�̓]�L
        If rebill_count > 0 Then
            ws_details.Cells(end_row, 2).value = "�Đ���"
            rebill_count = rebill_count - 1
        ElseIf late_count > 0 Then
            ws_details.Cells(end_row, 2).value = "�x����"
            late_count = late_count - 1
        ElseIf assessment_count > 0 Then
            ws_details.Cells(end_row, 2).value = "�Z��"
            assessment_count = assessment_count - 1
        End If
    Next key
End Sub

Private Sub WriteDataToDetails(ws As Worksheet, start_row_dict As Object, rebill_dict As Object, late_dict As Object, unpaid_dict As Object, assessment_dict As Object, payer_type As String)
    Dim ws_details As Worksheet
    Set ws_details = ws
    
    Dim row_index As Long
    Dim start_row As Long
    Dim end_row As Long
    Dim i_key As Long
    
    ' �e�J�e�S���̊J�n�s���擾
    For Each key In start_row_dict.Keys
        start_row = start_row_dict(key)
        end_row = start_row + 1
        
        ' �f�[�^�̓]�L
        If rebill_dict.exists(key) Then
            ws_details.Cells(end_row, 2).value = rebill_dict(key)(0)
            ws_details.Cells(end_row, 3).value = rebill_dict(key)(1)
            ws_details.Cells(end_row, 4).value = rebill_dict(key)(2)
            ws_details.Cells(end_row, 5).value = rebill_dict(key)(3)
        ElseIf late_dict.exists(key) Then
            ws_details.Cells(end_row, 2).value = late_dict(key)(0)
            ws_details.Cells(end_row, 3).value = late_dict(key)(1)
            ws_details.Cells(end_row, 4).value = late_dict(key)(2)
            ws_details.Cells(end_row, 5).value = late_dict(key)(3)
        ElseIf unpaid_dict.exists(key) Then
            ws_details.Cells(end_row, 2).value = unpaid_dict(key)(0)
            ws_details.Cells(end_row, 3).value = unpaid_dict(key)(1)
            ws_details.Cells(end_row, 4).value = unpaid_dict(key)(2)
            ws_details.Cells(end_row, 5).value = unpaid_dict(key)(3)
        ElseIf assessment_dict.exists(key) Then
            ws_details.Cells(end_row, 2).value = assessment_dict(key)(0)
            ws_details.Cells(end_row, 3).value = assessment_dict(key)(1)
            ws_details.Cells(end_row, 4).value = assessment_dict(key)(2)
            ws_details.Cells(end_row, 5).value = assessment_dict(key)(3)
        End If
    Next key
End Sub

