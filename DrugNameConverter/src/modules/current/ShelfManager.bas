Attribute VB_Name = "ShelfManager"
Sub UpdateShelfNumbersWithShelfInfo()
    Dim wsTana As Worksheet
    Dim wbMed As Workbook
    Dim wsMed As Worksheet
    Dim lastRowTana As Long
    Dim lastRowMed As Long
    Dim i As Long, j As Long
    Dim medName As String
    Dim medCodeName As String
    Dim shelf1 As String, shelf2 As String, shelf3 As String
    Dim outputFilePath As String
    
    ' �Ώۂ̃V�[�g��ݒ�
    Set wsTana = ThisWorkbook.Worksheets("tmp_tana")
    
    ' �I�������擾�iA1:B3�ɒI��1�`3�̏�񂪂���Ɖ���j
    shelf1 = ThisWorkbook.Sheets(1).Cells(1, 2).Value ' �I��1
    shelf2 = ThisWorkbook.Sheets(1).Cells(2, 2).Value ' �I��2
    shelf3 = ThisWorkbook.Sheets(1).Cells(3, 2).Value ' �I��3
    
    ' tmp_tana�V�[�g�̍ŏI�s���擾
    lastRowTana = wsTana.Cells(wsTana.Rows.Count, 1).End(xlUp).Row
    
    ' ���i�R�[�h�t�@�C�����J��
    Set wbMed = Workbooks.Open("/Users/yoshipc/Desktop/���i�R�[�h.xlsx")
    Set wsMed = wbMed.Sheets("�V�[�g1 - ���i�R�[�h")
    lastRowMed = wsMed.Cells(wsMed.Rows.Count, 1).End(xlUp).Row
    
    ' A4�ȍ~�̃Z���ɋL�ڂ��ꂽ���i�����X�g��tmp_tana�̈��i�𕔕���v����
    Dim readRow As Long
    readRow = 4 ' A4����ǂݎ��J�n�Ɖ���
    
    Do While ThisWorkbook.Sheets(1).Cells(readRow, 3).Value <> ""
        medName = ThisWorkbook.Sheets(1).Cells(readRow, 3).Value
        
        ' tmp_tana�̊e��i���ƕ�����v����
        For i = 2 To lastRowTana
            If InStr(1, wsTana.Cells(i, 2).Value, medName, vbTextCompare) > 0 Then
                ' ������v�����s�ɒI�Ԃ�ݒ�i�󗓂̏ꍇ�͕ύX���Ȃ��j
                If shelf1 <> "" Then wsTana.Cells(i, 7).Value = "[" & shelf1 & "]"
                If shelf2 <> "" Then wsTana.Cells(i, 8).Value = "[" & shelf2 & "]"
                If shelf3 <> "" Then wsTana.Cells(i, 9).Value = "[" & shelf3 & "]"
                Exit For
            End If
        Next i
        
        readRow = readRow + 1
    Loop
    
    ' CSV�t�@�C���Ƃ��ďo�͂���
    outputFilePath = Application.ThisWorkbook.Path & Application.PathSeparator & "updated_tmp_tana.csv"
    Call ExportToCSV(wsTana, outputFilePath)
    
    ' ���i�R�[�h�t�@�C�������
    wbMed.Close SaveChanges:=False
    
    MsgBox "�I�Ԃ̍X�V���������ACSV�t�@�C���Ƃ��ĕۑ����܂����B"
End Sub

' �V�[�g��CSV�t�@�C���Ƃ��ďo�͂���T�u�v���V�[�W��
Sub ExportToCSV(ws As Worksheet, filePath As String)
    Dim csvData As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    
    ' �V�[�g�̍ŏI�s�ƍŏI����擾
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' �f�[�^��CSV�`���ɕϊ�
    For i = 1 To lastRow
        For j = 1 To lastCol
            csvData = csvData & ws.Cells(i, j).Value
            If j < lastCol Then csvData = csvData & ","
        Next j
        csvData = csvData & vbNewLine
    Next i
    
    ' �t�@�C���ɏ�������
    Open filePath For Output As #1
    Print #1, csvData
    Close #1
End Sub


