VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "講師検索システム"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10980
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
    Dim lastRow As Long, lastColumn As Long '最終行、最終列の位置'
    Dim allData, resultData()   '全てのデータを格納する、結果のデータを格納する'
    Dim i As Long, j As Long, cnt As Long   'for文とかで使ういつもの変数'
    Dim sex As String   'オプションボタンの値を格納する変数'
    Dim subjectNum As Long  '科目番号を格納する変数'

    '検索するデータの全体をallDataに格納する'
   With Worksheets("sarchable") 'sarchableシートを参照する'
        lastRow = .Cells(Rows.Count, 1).End(xlUp).Row    '最終行の取得'
        lastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column '最終列の取得'
        allData = .Range(.Cells(1, 1), .Cells(lastRow, lastColumn)).Value  '最終行までデータを取得する'
    End With
    
    '科目番号を取得'
    subjectNum = Sheet1.subjectNum(ComboBox1.Value, ComboBox2.Value, ComboBox3.Value)
    
    '氏名のみ検索,科目のみ検索、氏名&科目検索のいずれかに分岐する'
    If TextBox1.Value = "" And Not subjectNum = -1 Then
        '科目の検索'
        Call sarchSubjectOnly(lastRow, subjectNum, allData)
    ElseIf Not TextBox1.Value = "" And subjectNum = -1 Then
        '氏名のみ検索'
        Call sarchTutorOnly(lastColumn, allData)
    ElseIf Not TextBox1 = "" And Not subjectNum = -1 Then
        '氏名&科目検索'
        Call sarchOnly(subjectNum, allData)
    End If
End Sub

'科目のみの検索'
Sub sarchSubjectOnly(lastRow As Long, subjectNum As Long, allData As Variant)

    'オプションボタンの状態を取得(性別を取得)'
    If OptionButton1 = True Then        '指定なし'
        sex = OptionButton1.Caption
    ElseIf OptionButton2 = True Then    '男性'
        sex = OptionButton2.Caption
    ElseIf OptionButton3 = True Then    '女性'
        sex = OptionButton3.Caption
    End If
    
     '検索結果を格納するために動的確保する'
    ReDim resultData(1 To lastRow, 1 To 3)
    
    '検索で一致したデータをresultDataに格納する'
    'コンボボックスをすべて埋めてない時の動作(異常終了)'
    If subjectNum = -1 Then
        MsgBox "学年、科目、詳細な科目はすべて選択してください"
    Else
        '性別:指定なしの時'
        If sex = "指定なし" Then
            For i = LBound(allData) To UBound(allData)
                If i = 1 Or allData(i, subjectNum) = "はい" Then
                    cnt = cnt + 1
                    resultData(cnt, 1) = allData(i, 1) '講師番号'
                    resultData(cnt, 2) = allData(i, 2) '講師名'
                    resultData(cnt, 3) = allData(i, 3) '電話番号'
                End If
            Next i
        Else
            '性別指定ありのとき'
            For i = LBound(allData) To UBound(allData)
                If i = 1 Or (allData(i, 4) Like sex And allData(i, subjectNum) = "はい") Then
                    cnt = cnt + 1
                    resultData(cnt, 1) = allData(i, 1) '講師番号'
                    resultData(cnt, 2) = allData(i, 2) '講師名'
                    resultData(cnt, 3) = allData(i, 3) '電話番号'
                End If
            Next i
        End If
    End If
    'リストボックに表示'
    With ListBox1
        .ColumnCount = 3
        .ColumnWidths = "50;70;50"
        .List = resultData
    End With
End Sub

'氏名のみ検索'
Private Sub sarchTutorOnly(lastColumn As Long, allData As Variant)
    Dim tutorName As String
    Dim i As Long, j As Long, k As Long
    Dim flag As Boolean
    ReDim resultData(1 To lastColumn)
    tutorName = TextBox1.Value
    'フラグ初期化'
    flag = False
    
    '講師名検索'
    For i = LBound(allData) To UBound(allData)
        If InStr(allData(i, 2), tutorName) Then
            '該当講師に関する情報だけ一次元配列に格納'
            For j = 5 To lastColumn
                    resultData(j - 4) = allData(i, j)
            Next j
            '講師が見つかったらフラグを立ててループを抜ける'
            flag = True
            Exit For
        End If
    Next i
    
    'フラグで分岐させる'
    If flag = True Then
        '幾何と代数の情報をを配列内から削除する'
        For k = 18 To UBound(resultData)
            resultData(k - 2) = resultData(k)
        Next k
        '結果を表示させる'
        Call showResult(resultData)
    Else
        MsgBox "講師が見つかりませんでした", vbOKOnly, "講師が見つからない"
    End If
End Sub

'氏名のみ検索の結果作成'
Private Sub showResult(resultData As Variant)
    Dim lastRow As Long
    Dim subjectCnt As Long, resultCnt As Long, i As Long, j As Long
    Dim subjectsList
    ReDim showList(1 To 22, 1 To 6)
    'カウントの初期化'
    subjectCnt = 18
    resultCnt = 1
    '科目データの取得(仕様で二次元配列で取得する)'
    With Worksheets("subjects")
        lastRow = .Cells(Rows.Count, 1).End(xlUp).Row    '最終行の取得'
        subjectsList = .Range(.Cells(1, 3), .Cells(lastRow, 3)).Value  '最終行までデータを取得する'
    End With
    'あらかじめshowListに入れておくデータ'
    showList(1, 1) = "小学生"
    showList(3, 1) = "中学受験"
    showList(5, 1) = "中学生"
    showList(7, 1) = "高校生"
    showList(19, 1) = "英語検定"
    '科目リストと表示形式の順番が異なるので手動で入れておくデータ'
    showList(1, 2) = subjectsList(2, 1) '英語'
    showList(1, 3) = subjectsList(3, 1) '数学'
    showList(1, 4) = subjectsList(5, 1) '国語'
    showList(1, 5) = subjectsList(7, 1) '理科'
    showList(1, 6) = subjectsList(9, 1) '社会'
    showList(3, 2) = ""                 '受験英語の欄は空欄'
    showList(3, 3) = subjectsList(4, 1) '受験数学'
    showList(3, 4) = subjectsList(6, 1) '受験国語'
    showList(3, 5) = subjectsList(8, 1) '受験理科'
    showList(3, 6) = subjectsList(10, 1) '受験社会'
    showList(5, 2) = subjectsList(11, 1) '英語'
    showList(5, 3) = subjectsList(12, 1) '数学'
    showList(5, 4) = subjectsList(15, 1) '英語'
    showList(5, 5) = subjectsList(16, 1) '数学'
    showList(5, 6) = subjectsList(17, 1) '英語'
    'showListに科目データを入れていく'
    For i = 1 To 22
        If i = 1 Or i = 3 Or i = 5 Then
            GoTo Continue
        End If
        For j = 2 To 6
            '倫理政治経済のあとは改行したい'
            If i = 17 And j = 3 Then
                GoTo Continue
            End If
            '配列の奇数行目は科目を入力'
            If i Mod 2 = 1 Then
                showList(i, j) = subjectsList(subjectCnt, 1)
                subjectCnt = subjectCnt + 1
            Else
            '配列の偶数行目はデータを入力'
                '倫理政治経済のあとは改行したい'
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

'氏名&科目検索'
Private Sub sarchOnly(subjectNum As Long, allData As Variant)
    Dim tutorName As String
    Dim i As Long
    Dim flag As Boolean
    'フラグの初期化'
    flag = False
    '講師名取得'
    tutorName = TextBox1.Value
    '氏名と科目で検索'
    For i = LBound(allData) To UBound(allData)
        If InStr(allData(i, 2), tutorName) And allData(i, subjectNum) = "はい" Then
        '可能ならフラグを立ててFor文を抜ける'
            tutorName = allData(i, 2)
            flag = True
            Exit For
        End If
    Next i
    'フラグによってメッセージボックスを出す'
    If flag = True Then
        MsgBox tutorName & "は" & ComboBox3.Value & "を教務可能です。", vbOKOnly, "教えられる？"
    Else
        MsgBox tutorName & "は" & ComboBox3.Value & "を教務できません。", vbOKOnly, "教えられる？"
    End If
End Sub

'検索結果を初期化'
Private Sub CommandButton2_Click()
    Call UserForm_Initialize
End Sub

'ユーザフォームの初期化'
Private Sub UserForm_Initialize()
    
    'オプションボタンの状態を初期化'
    OptionButton1 = True
    'コンボボックスを初期化'
    Call Sheet1.box_Initalize
    'テキストボックスの初期化'
    TextBox1.Text = ""
    'リストボックスの初期化'
    ListBox1.Clear
    
End Sub

