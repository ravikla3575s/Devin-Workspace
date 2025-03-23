Attribute VB_Name = "DrugNameConverter"
Option Explicit

' ���b�p�[���W���[�� - ��{�@�\

' ���C���������Ăяo�����b�p�[�֐��i7�s�ڈȍ~�̈��i����r�j
Public Sub RunDrugNameComparison()
    ' MainModule�̊֐����Ăяo��
    MainModule.ProcessFromRow7
End Sub

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
    
    MsgBox "��`�Ԃ̃h���b�v�_�E�����X�g��ݒ肵�܂����B", vbInformation
End Sub

' �V�[�g1�ɃC���X�g���N�V������ǉ�����֐�
Public Sub AddInstructions()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    
    ' �����̎w�����폜
    ws.Range("A2:C3").ClearContents
    
    ' �w����ǉ�
    ws.Range("A2").Value = "�y�g�����z"
    ws.Range("A3").Value = "1. B4�̕�`�Ԃ�I�����ĉ�����"
    ws.Range("A4").Value = "��`��:"
    ws.Range("B4").Font.Bold = True
    
    ' �Z���̏����ݒ�
    ws.Range("A2").Font.Bold = True
    ws.Range("A2").Font.Size = 12
    
    ' ���s���@�̎w��
    ws.Range("A5").Value = "2. B7�ȍ~�Ɍ���������i�������"
    ws.Range("A6").Value = "No."
    ws.Range("B6").Value = "�������i��"
    ws.Range("C6").Value = "��v���i��"
    
    With ws.Range("A6:C6")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247) ' �����F�̔w�i
    End With
    
    ' ���s���@�̈ē�
    Dim note As String
    note = "�����s���@: ���j���[����u�c�[���v���u�}�N���v��I�����A" & vbCrLf & _
           "�uRunDrugNameComparison�v��I��Łu���s�v�{�^�����N���b�N���܂��B"
    
    ws.Range("A" & (ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 2)).Value = note
    
    MsgBox "�g�p���@�̎w����ǉ����܂����B���j���[����u�c�[���v���u�}�N���v��I�����A" & vbCrLf & _
           "�uRunDrugNameComparison�v��I��ŏ��������s���Ă��������B", vbInformation
End Sub

' ���[�N�u�b�N�̏������֐�
Public Sub InitWorkbook()
    On Error GoTo ErrorHandler
    
    ' ���[�N�V�[�g�̎Q�Ƃ��擾
    Dim settingsSheet As Worksheet
    Dim targetSheet As Worksheet
    
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    Set targetSheet = ThisWorkbook.Worksheets(2)
    
    ' �V�[�g1�̐ݒ�
    With settingsSheet
        ' �^�C�g���ݒ�
        .Range("A1:C1").Merge
        .Range("A1").Value = "���i����r�c�[��"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.Color = RGB(180, 198, 231)
        
        ' �g����
        .Range("A2").Value = "�y�g�����z"
        .Range("A2").Font.Bold = True
        .Range("A3").Value = "1. B4�̕�`�Ԃ�I��"
        .Range("A4").Value = "��`��:"
        .Range("A4").Font.Bold = True
        
        ' �h���b�v�_�E�����X�g�ݒ�
        With .Range("B4").Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, _
                 Formula1:="(����`),���̑�(�Ȃ�),���,���ܗp,PTP,����,�o��,SP,PTP(���җp)"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        
        ' B4�Z���ݒ�
        .Range("B4").Value = "PTP"
        .Range("B4").Font.Bold = True
        .Range("B4").Interior.Color = RGB(217, 225, 242)
        
        ' �菇
        .Range("A5").Value = "2. B7�ȍ~�Ɍ���������i�������"
        .Range("A5").Font.Bold = True
        
        ' �w�b�_�[
        .Range("A6").Value = "No."
        .Range("B6").Value = "�������i��"
        .Range("C6").Value = "��v���i��"
        .Range("A6:C6").Font.Bold = True
        .Range("A6:C6").Interior.Color = RGB(221, 235, 247)
        
        ' ��
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 40
        
        ' �s�ԍ�
        Dim i As Long
        For i = 7 To 30
            .Cells(i, "A").Value = i - 6
        Next i
        
        ' ���s���@�̈ē�
        .Range("A32").Value = "�����s���@: ���j���[����u�c�[���v���u�}�N���v��I�����A�uRunDrugNameComparison�v�����s"
        .Range("A32").Font.Italic = True
        
        ' GS1�R�[�h�����ɊւẴ����������
        .Range("A34").Value = "�yGS1�R�[�h�����z"
        .Range("A34").Font.Bold = True
        .Range("A35").Value = "���j���[����u�c�[���v���u�}�N���v���uRunGS1CodeProcessing�v��"
        .Range("A36").Value = "GS1-128�̂P�S���R�[�h���痤�i�����������ݒ�V�[�g�ɓ]�L�ł��܂��B"
    End With
    
    ' �V�[�g2�̐ݒ�
    With targetSheet
        ' �^�C�g��
        .Range("A1:B1").Merge
        .Range("A1").Value = "��r�Ώۈ��i���X�g"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.Color = RGB(180, 198, 231)
        
        ' �w�b�_�[
        .Range("A2").Value = "No."
        .Range("B2").Value = "���i��"
        .Range("A2:B2").Font.Bold = True
        .Range("A2:B2").Interior.Color = RGB(221, 235, 247)
        
        ' ��
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 50
        
        ' �s�ԍ�
        For i = 3 To 30
            .Cells(i, "A").Value = i - 2
        Next i
    End With
    
    MsgBox "���[�N�u�b�N�����������܂����B" & vbNewLine & _
           "1. �ݒ�V�[�g��B4�Z���ŕ�`�Ԃ�I��" & vbNewLine & _
           "2. �V�[�g2�ɔ�r�Ώۂ̈��i�������" & vbNewLine & _
           "3. �ݒ�V�[�g��B7�ȍ~�Ɍ���������i�������" & vbNewLine & _
           "4. ���j���[�́u�c�[���v���u�}�N���v����uRunDrugNameComparison�v�����s", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical
End Sub
' GS1コード処理機能を実行するラッパー関数
Public Sub RunGS1CodeProcessing()
    ' MainModuleの関数を呼び出し
    MainModule.ProcessGS1DrugCode
End Sub

' メニューにGS1コード処理機能を追加する
Public Sub AddGS1ProcessingInstructions()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    
    ' GS1処理に関する説明を追加
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 2
    
    ws.Cells(lastRow, "A").Value = "【GS1コード処理機能】"
    ws.Cells(lastRow, "A").Font.Bold = True
    ws.Cells(lastRow, "A").Font.Size = 12
    
    ws.Cells(lastRow + 1, "A").Value = "メニューから「ツール」→「マクロ」→「RunGS1CodeProcessing」を"
    ws.Cells(lastRow + 2, "A").Value = "選択すると、GS1-128の14桁コードから医薬品情報を処理できます。"
    
    MsgBox "GS1コード処理機能の説明をシートに追加しました。", vbInformation
End Sub


