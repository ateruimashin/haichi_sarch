Attribute VB_Name = "Module2"
Public passwordResult As Boolean
'���C������'
Sub getDataMain()
    Call password
    Call queriesReflesh
    Call margeData
End Sub
'�p�X���[�h�F�؂�����'
Private Sub password()
    passwordResult = False
    UserForm2.Show
    If passwordResult = False Then
        End
    End If
End Sub
'�e�[�u�����݃`�F�b�N'
Private Sub queriesReflesh()
    'data�V�[�g�Ƀe�[�u�������݂��邩�m�F'
    With Worksheets("data")
        If .ListObjects.Count = 0 Then
            '�e�[�u�����Ȃ��ꍇ�̓G���[���b�Z�[�W���o��'
            MsgBox "�f�[�^�����݂��܂���B�}�j���A���ɉ����čĐڑ����Ă��������B"
            End
        Else
            '�e�[�u�������݂���ꍇ�͍X�V����'
            ActiveWorkbook.Connections("�N�G�� - toExcel (6)").Refresh
        End If
    End With
End Sub
'����ƃf�[�^������������'
Private Sub margeData()
    Dim supreadsheetData        '�X�v���b�h�V�[�g���瓾���f�[�^��sarchable�V�[�g�̃f�[�^'
    Dim meiboData               '����̃f�[�^'
    Dim lastRow, lastColumn     '�ŏI�s�A�ŏI��'
    Dim i As Long               'For����index�p'
    Dim resultRg As Range       '�������ʂ�Range�I�u�W�F�N�g�p'

    
    '�X�v���b�h�V�[�g���瓾���f�[�^��sarchable�ɃR�s�['
    Set supreadsheetData = Worksheets("data").UsedRange
    'sarchable�V�[�g�S�̂̃Z�����N���A'
    Worksheets("sarchable").UsedRange.ClearContents
    '�f�[�^���R�s�['
    supreadsheetData.Copy Destination:=Worksheets("sarchable").Range("A1")
    With Worksheets("sarchable")
        'B,C���}������(�u�t�ԍ��Ɠd�b�ԍ�������)'
        .Columns("B:C").Insert
        'B1��C1�ɗ�̖��O������'
        .Range("B1").Value = "���O"
        .Range("C1").Value = "�d�b�ԍ�"
    End With
    
    '����̃f�[�^��Range�I�u�W�F�N�g�Ƃ��Ď擾'
    With Worksheets("meibo").UsedRange
        lastRow = .Cells(Rows.Count, 1).End(xlUp).Row    '�ŏI�s�̎擾'
        lastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column '�ŏI��̎擾'
        meiboData = .Range(.Cells(2, 1), .Cells(lastRow, lastColumn)).Value  '�ŏI�s�܂Ńf�[�^���擾����'
    End With
    
    'sarchable�V�[�g�̍u�t�ԍ����u�t���납�猟�����A���O�Ɠd�b�ԍ�������'
    With Worksheets("sarchable")
        '�d�b�ԍ��̗�̌`���𕶎���ɂ���'
        .UsedRange.Columns("B:C").NumberFormatLocal = "@"
        '�u�t�ԍ����u�t���납�猟������'
        For i = LBound(meiboData) To UBound(meiboData)
            Set resultRg = .UsedRange.Columns(1).Find(meiboData(i, 1), LookIn:=xlValues)
            '�������sarchable�f�[�^��B���C��Ƀf�[�^����������'
            If Not resultRg Is Nothing Then
                '���O����������'
                .Cells(resultRg.Row, 2).Value = meiboData(i, 2)
                '�d�b�ԍ��𕶎���Ƃ��ď�������'
                .Cells(resultRg.Row, 3).Value = meiboData(i, 3)
            End If
        Next i
    End With
    
End Sub
