VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�u�t�����V�X�e��"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10980
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboBox1_Change()
    Call Sheet1.makeCombobox2
End Sub

Private Sub ComboBox2_Change()
    Call Sheet1.makeCombobox3
End Sub

'���������s����'
Private Sub CommandButton1_Click()
    Dim lastRow As Long, lastColumn As Long '�ŏI�s�A�ŏI��̈ʒu'
    Dim allData, resultData()   '�S�Ẵf�[�^���i�[����A���ʂ̃f�[�^���i�[����'
    Dim i As Long, j As Long, cnt As Long   'for���Ƃ��Ŏg�������̕ϐ�'
    Dim sex As String   '�I�v�V�����{�^���̒l���i�[����ϐ�'
    Dim subjectNum As Long  '�Ȗڔԍ����i�[����ϐ�'

    '��������f�[�^�̑S�̂�allData�Ɋi�[����'
   With Worksheets("sarchable") 'sarchable�V�[�g���Q�Ƃ���'
        lastRow = .Cells(Rows.Count, 1).End(xlUp).Row    '�ŏI�s�̎擾'
        lastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column '�ŏI��̎擾'
        allData = .Range(.Cells(1, 1), .Cells(lastRow, lastColumn)).Value  '�ŏI�s�܂Ńf�[�^���擾����'
    End With
    
    '�Ȗڔԍ����擾'
    subjectNum = Sheet1.subjectNum(ComboBox1.Value, ComboBox2.Value, ComboBox3.Value)
    
    '�����̂݌���,�Ȗڂ̂݌����A����&�Ȗڌ����̂����ꂩ�ɕ��򂷂�'
    If TextBox1.Value = "" And Not subjectNum = -1 Then
        '�Ȗڂ̌���'
        Call sarchSubjectOnly(lastRow, subjectNum, allData)
    ElseIf Not TextBox1.Value = "" And subjectNum = -1 Then
        '�����̂݌���'
        Call sarchTutorOnly(lastColumn, allData)
    ElseIf Not TextBox1 = "" And Not subjectNum = -1 Then
        '����&�Ȗڌ���'
        Call sarchOnly(subjectNum, allData)
    End If
End Sub

'�Ȗڂ݂̂̌���'
Sub sarchSubjectOnly(lastRow As Long, subjectNum As Long, allData As Variant)

    '�I�v�V�����{�^���̏�Ԃ��擾(���ʂ��擾)'
    If OptionButton1 = True Then        '�w��Ȃ�'
        sex = OptionButton1.Caption
    ElseIf OptionButton2 = True Then    '�j��'
        sex = OptionButton2.Caption
    ElseIf OptionButton3 = True Then    '����'
        sex = OptionButton3.Caption
    End If
    
     '�������ʂ��i�[���邽�߂ɓ��I�m�ۂ���'
    ReDim resultData(1 To lastRow, 1 To 3)
    
    '�����ň�v�����f�[�^��resultData�Ɋi�[����'
    '�R���{�{�b�N�X�����ׂĖ��߂ĂȂ����̓���(�ُ�I��)'
    If subjectNum = -1 Then
        MsgBox "�w�N�A�ȖځA�ڍׂȉȖڂ͂��ׂđI�����Ă�������"
    Else
        '����:�w��Ȃ��̎�'
        If sex = "�w��Ȃ�" Then
            For i = LBound(allData) To UBound(allData)
                If i = 1 Or allData(i, subjectNum) = "�͂�" Then
                    cnt = cnt + 1
                    resultData(cnt, 1) = allData(i, 1) '�u�t�ԍ�'
                    resultData(cnt, 2) = allData(i, 2) '�u�t��'
                    resultData(cnt, 3) = allData(i, 3) '�d�b�ԍ�'
                End If
            Next i
        Else
            '���ʎw�肠��̂Ƃ�'
            For i = LBound(allData) To UBound(allData)
                If i = 1 Or (allData(i, 4) Like sex And allData(i, subjectNum) = "�͂�") Then
                    cnt = cnt + 1
                    resultData(cnt, 1) = allData(i, 1) '�u�t�ԍ�'
                    resultData(cnt, 2) = allData(i, 2) '�u�t��'
                    resultData(cnt, 3) = allData(i, 3) '�d�b�ԍ�'
                End If
            Next i
        End If
    End If
    '���X�g�{�b�N�ɕ\��'
    With ListBox1
        .ColumnCount = 3
        .ColumnWidths = "50;70;50"
        .List = resultData
    End With
End Sub

'�����̂݌���'
Private Sub sarchTutorOnly(lastColumn As Long, allData As Variant)
    Dim tutorName As String
    Dim i As Long, j As Long, k As Long
    Dim flag As Boolean
    ReDim resultData(1 To lastColumn)
    tutorName = TextBox1.Value
    '�t���O������'
    flag = False
    
    '�u�t������'
    For i = LBound(allData) To UBound(allData)
        If InStr(allData(i, 2), tutorName) Then
            '�Y���u�t�Ɋւ����񂾂��ꎟ���z��Ɋi�['
            For j = 5 To lastColumn
                    resultData(j - 4) = allData(i, j)
            Next j
            '�u�t������������t���O�𗧂Ăă��[�v�𔲂���'
            flag = True
            Exit For
        End If
    Next i
    
    '�t���O�ŕ��򂳂���'
    If flag = True Then
        '�􉽂Ƒ㐔�̏�����z�������폜����'
        For k = 18 To UBound(resultData)
            resultData(k - 2) = resultData(k)
        Next k
        '���ʂ�\��������'
        Call showResult(resultData)
    Else
        MsgBox "�u�t��������܂���ł���", vbOKOnly, "�u�t��������Ȃ�"
    End If
End Sub

'�����̂݌����̌��ʍ쐬'
Private Sub showResult(resultData As Variant)
    Dim lastRow As Long
    Dim subjectCnt As Long, resultCnt As Long, i As Long, j As Long
    Dim subjectsList
    ReDim showList(1 To 22, 1 To 6)
    '�J�E���g�̏�����'
    subjectCnt = 18
    resultCnt = 1
    '�Ȗڃf�[�^�̎擾(�d�l�œ񎟌��z��Ŏ擾����)'
    With Worksheets("subjects")
        lastRow = .Cells(Rows.Count, 1).End(xlUp).Row    '�ŏI�s�̎擾'
        subjectsList = .Range(.Cells(1, 3), .Cells(lastRow, 3)).Value  '�ŏI�s�܂Ńf�[�^���擾����'
    End With
    '���炩����showList�ɓ���Ă����f�[�^'
    showList(1, 1) = "���w��"
    showList(3, 1) = "���w��"
    showList(5, 1) = "���w��"
    showList(7, 1) = "���Z��"
    showList(19, 1) = "�p�ꌟ��"
    '�Ȗڃ��X�g�ƕ\���`���̏��Ԃ��قȂ�̂Ŏ蓮�œ���Ă����f�[�^'
    showList(1, 2) = subjectsList(2, 1) '�p��'
    showList(1, 3) = subjectsList(3, 1) '���w'
    showList(1, 4) = subjectsList(5, 1) '����'
    showList(1, 5) = subjectsList(7, 1) '����'
    showList(1, 6) = subjectsList(9, 1) '�Љ�'
    showList(3, 2) = ""                 '�󌱉p��̗��͋�'
    showList(3, 3) = subjectsList(4, 1) '�󌱐��w'
    showList(3, 4) = subjectsList(6, 1) '�󌱍���'
    showList(3, 5) = subjectsList(8, 1) '�󌱗���'
    showList(3, 6) = subjectsList(10, 1) '�󌱎Љ�'
    showList(5, 2) = subjectsList(11, 1) '�p��'
    showList(5, 3) = subjectsList(12, 1) '���w'
    showList(5, 4) = subjectsList(15, 1) '�p��'
    showList(5, 5) = subjectsList(16, 1) '���w'
    showList(5, 6) = subjectsList(17, 1) '�p��'
    'showList�ɉȖڃf�[�^�����Ă���'
    For i = 1 To 22
        If i = 1 Or i = 3 Or i = 5 Then
            GoTo Continue
        End If
        For j = 2 To 6
            '�ϗ������o�ς̂��Ƃ͉��s������'
            If i = 17 And j = 3 Then
                GoTo Continue
            End If
            '�z��̊�s�ڂ͉Ȗڂ����'
            If i Mod 2 = 1 Then
                showList(i, j) = subjectsList(subjectCnt, 1)
                subjectCnt = subjectCnt + 1
            Else
            '�z��̋����s�ڂ̓f�[�^�����'
                '�ϗ������o�ς̂��Ƃ͉��s������'
                If i = 18 And j = 3 Then
                    GoTo Continue
                End If
                If i = 4 And j = 2 Then
                    GoTo jContinue
                End If
                showList(i, j) = resultData(resultCnt)
                resultCnt = resultCnt + 1
            End If
jContinue:
        Next j
Continue:
    Next i
    With ListBox1
        .ColumnCount = 6
        .ColumnWidths = "50;60;60;60;60;60"
        .List = showList
    End With
End Sub

'����&�Ȗڌ���'
Private Sub sarchOnly(subjectNum As Long, allData As Variant)
    Dim tutorName As String
    Dim i As Long
    Dim flag As Boolean
    '�t���O�̏�����'
    flag = False
    '�u�t���擾'
    tutorName = TextBox1.Value
    '�����ƉȖڂŌ���'
    For i = LBound(allData) To UBound(allData)
        If InStr(allData(i, 2), tutorName) And allData(i, subjectNum) = "�͂�" Then
        '�\�Ȃ�t���O�𗧂Ă�For���𔲂���'
            tutorName = allData(i, 2)
            flag = True
            Exit For
        End If
    Next i
    '�t���O�ɂ���ă��b�Z�[�W�{�b�N�X���o��'
    If flag = True Then
        MsgBox tutorName & "��" & ComboBox3.Value & "�������\�ł��B", vbOKOnly, "��������H"
    Else
        MsgBox tutorName & "��" & ComboBox3.Value & "�������ł��܂���B", vbOKOnly, "��������H"
    End If
End Sub

'�������ʂ�������'
Private Sub CommandButton2_Click()
    Call UserForm_Initialize
End Sub

'���[�U�t�H�[���̏�����'
Private Sub UserForm_Initialize()
    
    '�I�v�V�����{�^���̏�Ԃ�������'
    OptionButton1 = True
    '�R���{�{�b�N�X��������'
    Call Sheet1.box_Initalize
    '�e�L�X�g�{�b�N�X�̏�����'
    TextBox1.Text = ""
    '���X�g�{�b�N�X�̏�����'
    ListBox1.Clear
    
End Sub

