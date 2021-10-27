Attribute VB_Name = "Module2"
Public passwordResult As Boolean
'メイン部分'
Sub getDataMain()
    Call password
    Call queriesReflesh
    Call margeData
End Sub
'パスワード認証をする'
Private Sub password()
    passwordResult = False
    UserForm2.Show
    If passwordResult = False Then
        End
    End If
End Sub
'テーブル存在チェック'
Private Sub queriesReflesh()
    'dataシートにテーブルが存在するか確認'
    With Worksheets("data")
        If .ListObjects.Count = 0 Then
            'テーブルがない場合はエラーメッセージを出す'
            MsgBox "データが存在しません。マニュアルに沿って再接続してください。"
            End
        Else
            'テーブルが存在する場合は更新する'
            ActiveWorkbook.Connections("クエリ - toExcel (6)").Refresh
        End If
    End With
End Sub
'名簿とデータを結合させる'
Private Sub margeData()
    Dim supreadsheetData        'スプレッドシートから得たデータとsarchableシートのデータ'
    Dim meiboData               '名簿のデータ'
    Dim lastRow, lastColumn     '最終行、最終列'
    Dim i As Long               'For文のindex用'
    Dim resultRg As Range       '検索結果のRangeオブジェクト用'

    
    'スプレッドシートから得たデータをsarchableにコピー'
    Set supreadsheetData = Worksheets("data").UsedRange
    'sarchableシート全体のセルをクリア'
    Worksheets("sarchable").UsedRange.ClearContents
    'データをコピー'
    supreadsheetData.Copy Destination:=Worksheets("sarchable").Range("A1")
    With Worksheets("sarchable")
        'B,C列を挿入する(講師番号と電話番号が入る)'
        .Columns("B:C").Insert
        'B1とC1に列の名前を入れる'
        .Range("B1").Value = "名前"
        .Range("C1").Value = "電話番号"
    End With
    
    '名簿のデータをRangeオブジェクトとして取得'
    With Worksheets("meibo").UsedRange
        lastRow = .Cells(Rows.Count, 1).End(xlUp).Row    '最終行の取得'
        lastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column '最終列の取得'
        meiboData = .Range(.Cells(2, 1), .Cells(lastRow, lastColumn)).Value  '最終行までデータを取得する'
    End With
    
    'sarchableシートの講師番号を講師名簿から検索し、名前と電話番号を入れる'
    With Worksheets("sarchable")
        '電話番号の列の形式を文字列にする'
        .UsedRange.Columns("B:C").NumberFormatLocal = "@"
        '講師番号を講師名簿から検索する'
        For i = LBound(meiboData) To UBound(meiboData)
            Set resultRg = .UsedRange.Columns(1).Find(meiboData(i, 1), LookIn:=xlValues)
            '見つかればsarchableデータのB列とC列にデータを書き込む'
            If Not resultRg Is Nothing Then
                '名前を書き込む'
                .Cells(resultRg.Row, 2).Value = meiboData(i, 2)
                '電話番号を文字列として書き込む'
                .Cells(resultRg.Row, 3).Value = meiboData(i, 3)
            End If
        Next i
    End With
    
End Sub
