VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "�p�X���[�h�F��"
   ClientHeight    =   1800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim myMsg As Integer
    If TextBox1.Value = "password" Then
        myMsg = MsgBox("�F�؂��܂���", vbOKOnly)
        passwordResult = True
        Unload Me
    Else
        myMsg = MsgBox("�F�؂ł��܂���ł���", _
        vbOKOnly + vbInformation, "�p�X���[�h�F��")
        
        With TextBox1
            .Value = ""
            .SetFocus
        End With
    End If
End Sub

Private Sub CommandButton2_Click()
    passwordResult = False
    Unload Me
End Sub
