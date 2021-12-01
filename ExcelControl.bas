Attribute VB_Name = "ExcelControl"
Option Explicit

'************************
'エクセルシート成形処理
'************************

Sub シートクリア(対象シート As String, 開始行 As Long, 開始列 As Long, 対象列数 As Long, 行カウントする列 As Long)
'テーブル取得時、出力エリアの値をクリア

Dim maxRow As Long

If Worksheets(対象シート).Cells(開始行, 行カウントする列).Value = "" Then
    Exit Sub
End If

If Worksheets(対象シート).Cells(開始行 + 1, 行カウントする列).Value = "" Then
    
    maxRow = 開始行

Else
    
    maxRow = Worksheets(対象シート).Cells(開始行, 行カウントする列).End(xlDown).Row

End If

    Worksheets(対象シート).Range(Worksheets(対象シート).Cells(開始行, 開始列), Worksheets(対象シート).Cells(maxRow, 開始列 + 対象列数 - 1)).Select
    Selection.ClearContents
    
    With Selection.Font
        .Name = "游ゴシック"
        .FontStyle = "標準"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    
    Worksheets(対象シート).Cells(開始行, 開始列).Select
End Sub


Sub シートソート(対象シート As String, 開始行 As Long, 開始列 As Long, 対象列数 As Long, 行カウントする列 As Long, ソートキー列 As Long, 並び順 As String)
'テーブル取得時、指定があった場合ソートする（エクセルの機能で）、ただし現状ソートキーに指定できるのは１項目のみ

Const asc As String = "Ascending"
Const dec As String = "Descending"


    Dim maxRow As Long
    
    maxRow = Worksheets(対象シート).Cells(開始行, 行カウントする列).End(xlDown).Row
   
    'Columns("A:G").Select
    Worksheets(対象シート).Range(Worksheets(対象シート).Cells(開始行, 開始列), Worksheets(対象シート).Cells(maxRow, 開始列 + 対象列数 - 1)).Select
    
    ActiveWorkbook.Worksheets(対象シート).Sort.SortFields.Clear
    
    If 並び順 = "a" Then
        ActiveWorkbook.Worksheets(対象シート).Sort.SortFields.Add2 key:=Range(Worksheets(対象シート).Cells(開始行, ソートキー列), Worksheets(対象シート).Cells(maxRow, ソートキー列)), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    Else
    '降順
        ActiveWorkbook.Worksheets(対象シート).Sort.SortFields.Add2 key:=Range(Worksheets(対象シート).Cells(開始行, ソートキー列), Worksheets(対象シート).Cells(maxRow, ソートキー列)), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    End If
    
    With ActiveWorkbook.Worksheets(対象シート).Sort
        .SetRange Range(Worksheets(対象シート).Cells(開始行, 開始列), Worksheets(対象シート).Cells(maxRow, 開始列 + 対象列数 - 1))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

