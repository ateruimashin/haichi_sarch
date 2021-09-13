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


'検索を実行する'
Private Sub CommandButton1_Click()
    Dim lastRow As Long, lastColumn As Long '最終行の位置'
    Dim allData, resultData()   '全てのデータを格納する、結果のデータを格納する'
    Dim i As Long, j As Long, cnt As Long   'for文とかで使ういつもの変数'
    Dim sex As String   'オプションボタンの値を格納する変数'

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
    
    ReDim resultData(1 To lastRow, 1 To 3)  '検索結果を格納するために動的確保する'
    '検索で一致したデータをresultDataに格納する'
    '性別:指定なしの時'
    If sex = "指定なし" Then
        For i = LBound(allData) To UBound(allData)
            cnt = cnt + 1
            resultData(cnt, 1) = allData(i, 1) '講師番号'
            resultData(cnt, 2) = allData(i, 2) '講師名'
            resultData(cnt, 3) = allData(i, 4) '電話番号'
        Next i
    Else
        '性別指定ありのとき'
        For i = LBound(allData) To UBound(allData)
            If i = 1 Or allData(i, 3) Like sex Then '一番上に講師名などを表示するためにi=1(1行目)のみ別で分岐'
                cnt = cnt + 1
                resultData(cnt, 1) = allData(i, 1) '講師番号'
                resultData(cnt, 2) = allData(i, 2) '講師名'
                resultData(cnt, 3) = allData(i, 4) '電話番号'
            End If
        Next i
    End If
    
    With ListBox1
        .ColumnCount = 3
        .ColumnWidths = "50;50;50"
        .List = resultData
    End With
End Sub


'ユーザフォームの初期化'
Private Sub UserForm_Initialize()
    
    'オプションボタンの状態を取得(性別を取得)'
    OptionButton1 = True
    
    
End Sub

