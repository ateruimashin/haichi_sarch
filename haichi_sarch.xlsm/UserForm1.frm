VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6225
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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

'検索を実行する'
Private Sub CommandButton1_Click()
    Dim lastRow As Long, lastColumn As Long '最終行の位置'
    Dim allData, resultData()   '全てのデータを格納する、結果のデータを格納する'
    Dim i As Long, j As Long, cnt As Long   'for文とかで使ういつもの変数'
    Dim sex As String   'オプションボタンの値を格納する変数'
    Dim subjectNum As Long  '科目番号を格納する変数'

    '検索するデータの全体をallDataに格納する'
   With Worksheets("meibo") 'meiboシートを参照する'
        lastRow = .Cells(Rows.Count, 1).End(xlUp).Row    '最終行の取得'
        lastColumn = .Cells(Columns.Count, 1).End(xlToRight).Column '最終列の取得'
        allData = .Range(.Cells(1, 1), .Cells(lastRow, lastColumn)).Value  '最終行までデータを取得する'
    End With
    
    'オプションボタンの状態を取得(性別を取得)'
    If OptionButton1 = True Then        '指定なし'
        sex = OptionButton1.Caption
    ElseIf OptionButton2 = True Then    '男性'
        sex = OptionButton2.Caption
    ElseIf OptionButton3 = True Then    '女性'
        sex = OptionButton3.Caption
    End If

    '科目番号を取得'
    subjectNum = Sheet1.subjectNum(ComboBox1.Value, ComboBox2.Value, ComboBox3.Value)
    
    ReDim resultData(1 To lastRow, 1 To 3)  '検索結果を格納するために動的確保する'
    
    '検索で一致したデータをresultDataに格納する'
    'コンボボックスをすべて埋めてない時の動作(異常終了)'
    If subjectNum = -1 Then
        MsgBox "学年、科目、詳細な科目はすべて選択してください"
    Else
        '性別:指定なしの時'
        If sex = "指定なし" Then
            For i = LBound(allData) To UBound(allData)
                If i = 1 Or allData(i, subjectNum) = 1 Then
                    cnt = cnt + 1
                    resultData(cnt, 1) = allData(i, 1) '講師番号'
                    resultData(cnt, 2) = allData(i, 2) '講師名'
                    resultData(cnt, 3) = allData(i, 4) '電話番号'
                End If
            Next i
        Else
            '性別指定ありのとき'
            For i = LBound(allData) To UBound(allData)
                If i = 1 Or (allData(i, 3) Like sex And allData(i, subjectNum) = 1) Then
                    cnt = cnt + 1
                    resultData(cnt, 1) = allData(i, 1) '講師番号'
                    resultData(cnt, 2) = allData(i, 2) '講師名'
                    resultData(cnt, 3) = allData(i, 4) '電話番号'
                End If
            Next i
        End If
    End If
    'リストボックに表示'
    With ListBox1
        .ColumnCount = 3
        .ColumnWidths = "50;50;50"
        .List = resultData
    End With
End Sub


'ユーザフォームの初期化'
Private Sub UserForm_Initialize()
    
    'オプションボタンの状態を初期化'
    OptionButton1 = True
    'コンボボックスを初期化'
    Call Sheet1.box_Initalize
    '[未実装]氏名から検索は実装していないので'
    TextBox1.Value = "氏名検索は未実装です"
    
End Sub

