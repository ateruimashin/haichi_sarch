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
'���������s����'
Private Sub CommandButton1_Click()
    Dim lastRow As Long, lastColumn As Long '�ŏI�s�̈ʒu'
    Dim allData, resultData()   '�S�Ẵf�[�^���i�[����A���ʂ̃f�[�^���i�[����'
    Dim i As Long, j As Long, cnt As Long   'for���Ƃ��Ŏg�������̕ϐ�'
    Dim sex As String   '�I�v�V�����{�^���̒l���i�[����ϐ�'
    Dim myBook As Workbook  '���[�N�u�b�N'
    Set myBook = Workbooks("haich_sarch_data.xlsx")
    
    '��������f�[�^�̑S�̂�allData�Ɋi�[����'
   With myBook.Worksheets("mainData")
        lastRow = .Cells(Rows.Count, 1).End(xlUp).Row    '�ŏI�s�̎擾'
        lastColumn = .Cells(Columns.Count, 1).End(xlToRight).Column '�ŏI��̎擾'
        allData = .Range(.Cells(1, 1), .Cells(lastRow, lastColumn)).Value  '�ŏI�s�܂Ńf�[�^���擾����'
    End With
    
    '�I�v�V�����{�^���̏�Ԃ��擾(���ʂ��擾)'
    If OptionButton1 = True Then
        sex = OptionButton1.Caption
    ElseIf OptionButton2 = True Then
        sex = OptionButton2.Caption
    End If
    
    ReDim resultData(1 To lastRow, 1 To 3)  '�������ʂ��i�[���邽�߂ɓ��I�m�ۂ���'
    '�����ň�v�����f�[�^��resultData�Ɋi�[����'
    For i = LBound(allData) To UBound(allData)
        If allData(i, 3) Like sex Then
            cnt = cnt + 1
            resultData(cnt, 1) = allData(i, 1) '�u�t�ԍ�'
            resultData(cnt, 2) = allData(i, 2) '�u�t��'
            resultData(cnt, 3) = allData(i, 4) '�d�b�ԍ�'
        End If
    Next i
    
    With ListBox1
        .ColumnCount = 3
        .ColumnWidths = "50;50;50"
        .List = resultData
    End With
End Sub


