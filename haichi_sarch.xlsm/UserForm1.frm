VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6225
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
    Dim lastRow As Long, lastColumn As Long '�ŏI�s�̈ʒu'
    Dim allData, resultData()   '�S�Ẵf�[�^���i�[����A���ʂ̃f�[�^���i�[����'
    Dim i As Long, j As Long, cnt As Long   'for���Ƃ��Ŏg�������̕ϐ�'
    Dim sex As String   '�I�v�V�����{�^���̒l���i�[����ϐ�'
    Dim subjectNum As Long  '�Ȗڔԍ����i�[����ϐ�'

    '��������f�[�^�̑S�̂�allData�Ɋi�[����'
   With Worksheets("meibo") 'meibo�V�[�g���Q�Ƃ���'
        lastRow = .Cells(Rows.Count, 1).End(xlUp).Row    '�ŏI�s�̎擾'
        lastColumn = .Cells(Columns.Count, 1).End(xlToRight).Column '�ŏI��̎擾'
        allData = .Range(.Cells(1, 1), .Cells(lastRow, lastColumn)).Value  '�ŏI�s�܂Ńf�[�^���擾����'
    End With
    
    '�I�v�V�����{�^���̏�Ԃ��擾(���ʂ��擾)'
    If OptionButton1 = True Then        '�w��Ȃ�'
        sex = OptionButton1.Caption
    ElseIf OptionButton2 = True Then    '�j��'
        sex = OptionButton2.Caption
    ElseIf OptionButton3 = True Then    '����'
        sex = OptionButton3.Caption
    End If

    '�Ȗڔԍ����擾'
    subjectNum = Sheet1.subjectNum(ComboBox1.Value, ComboBox2.Value, ComboBox3.Value)
    
    ReDim resultData(1 To lastRow, 1 To 3)  '�������ʂ��i�[���邽�߂ɓ��I�m�ۂ���'
    
    '�����ň�v�����f�[�^��resultData�Ɋi�[����'
    '�R���{�{�b�N�X�����ׂĖ��߂ĂȂ����̓���(�ُ�I��)'
    If subjectNum = -1 Then
        MsgBox "�w�N�A�ȖځA�ڍׂȉȖڂ͂��ׂđI�����Ă�������"
    Else
        '����:�w��Ȃ��̎�'
        If sex = "�w��Ȃ�" Then
            For i = LBound(allData) To UBound(allData)
                If i = 1 Or allData(i, subjectNum) = 1 Then
                    cnt = cnt + 1
                    resultData(cnt, 1) = allData(i, 1) '�u�t�ԍ�'
                    resultData(cnt, 2) = allData(i, 2) '�u�t��'
                    resultData(cnt, 3) = allData(i, 4) '�d�b�ԍ�'
                End If
            Next i
        Else
            '���ʎw�肠��̂Ƃ�'
            For i = LBound(allData) To UBound(allData)
                If i = 1 Or (allData(i, 3) Like sex And allData(i, subjectNum) = 1) Then
                    cnt = cnt + 1
                    resultData(cnt, 1) = allData(i, 1) '�u�t�ԍ�'
                    resultData(cnt, 2) = allData(i, 2) '�u�t��'
                    resultData(cnt, 3) = allData(i, 4) '�d�b�ԍ�'
                End If
            Next i
        End If
    End If
    '���X�g�{�b�N�ɕ\��'
    With ListBox1
        .ColumnCount = 3
        .ColumnWidths = "50;50;50"
        .List = resultData
    End With
End Sub


'���[�U�t�H�[���̏�����'
Private Sub UserForm_Initialize()
    
    '�I�v�V�����{�^���̏�Ԃ�������'
    OptionButton1 = True
    '�R���{�{�b�N�X��������'
    Call Sheet1.box_Initalize
    '[������]�������猟���͎������Ă��Ȃ��̂�'
    TextBox1.Value = "���������͖������ł�"
    
End Sub

