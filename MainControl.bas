Attribute VB_Name = "MainControl"
Option Explicit

'************************
'テーブル取得・登録処理
'************************

Public GetMultiRecordJson As Object
Public GetSingleRecordJson As Object
Public CreateRecordJson As Object
Public objHTTP As Object

Public Enum テーブル取得入力エリア
    開始行 = 6
    開始列 = 1
End Enum

Public Enum テーブル取得出力エリア
    開始行 = 23
    開始列 = 1
End Enum

Public Enum テーブル登録入力エリア
    開始行 = 10
    開始列 = 1
End Enum

Sub アクティブシートクリア()

Dim maxCol As Long
Dim sheetName As String

sheetName = ActiveSheet.Name

maxCol = Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, テーブル取得出力エリア.開始列).End(xlToRight).Column

Call シートクリア(sheetName, テーブル取得出力エリア.開始行 + 1, テーブル取得出力エリア.開始列, maxCol, テーブル取得出力エリア.開始列)

End Sub


Sub テーブル取得メイン()

Dim sheetName As String
Dim maxCol As Long
Dim maxRow As Long
Dim j As Long
Dim i As Long
Dim filterFirst As Boolean
Dim strUrl As String

Dim Parse As Object
Dim DesStr As Variant
Dim Pj() As Variant

Dim sortKey As Long
Dim sortOrder As String


Application.ScreenUpdating = False

'事前にシートをクリアしておく
Call アクティブシートクリア

filterFirst = True

'アクティブシート名を取得
sheetName = ActiveSheet.Name

'リクエストのためのjson用Dictionaryを用意する
Set GetMultiRecordJson = New Dictionary
Set GetMultiRecordJson = SetApiKey(GetEnvironmentApiKey, GetMultiRecordJson)

'入力エリアに設定された条件値の数をEnd(xlToRight).Columnで取得する
maxCol = Worksheets(sheetName).Cells(テーブル取得入力エリア.開始行, テーブル取得入力エリア.開始列).End(xlToRight).Column

Application.StatusBar = "Pleasanterテーブル取得開始..."

'設定された条件値の分ループして、リクエストするjson用Dictionaryを生成する
For j = 1 To maxCol

    If Worksheets(sheetName).Cells(テーブル取得入力エリア.開始行 + 1, j).Value <> "" Then
        
        If filterFirst = True Then
            'フィルターがあれば初回だけ作る
            Set GetMultiRecordJson = SetView(GetMultiRecordJson)
            'フィルターがあれば初回だけ作る
            Set GetMultiRecordJson = SetColumnFilterHash(GetMultiRecordJson)
            filterFirst = False
        End If
        
        '関数にキーと値を渡してjson用Dictionaryに設定する
        Set GetMultiRecordJson = SetColumnFilterHashItem(Worksheets(sheetName).Cells(テーブル取得入力エリア.開始行, j).Value, _
                                                    Worksheets(sheetName).Cells(テーブル取得入力エリア.開始行 + 1, j).Value, _
                                                    GetMultiRecordJson)
        
    End If

DoEvents

Next j

'リクエスト用のURLを設定する
strUrl = Worksheets(sheetName).Cells(3, 2).Value & "/api/items/" & Worksheets(sheetName).Cells(2, 2).Value & "/get"

'HTTPリクエスト発行
Set objHTTP = HTTPリクエスト発行(strUrl, GetMultiRecordJson)

'レスポンスjsonをパース
Set Parse = JsonConverter.ParseJson(objHTTP.responseText)

Application.StatusBar = "Pleasanterテーブル取得終了..."

'取得件数をチェック、0件なら処理終了
If Parse("Response")("Data").count <= 0 Then
    
    Application.StatusBar = False
    
    Exit Sub

End If

'出力エリアに設定された出力カラムの数をEnd(xlToRight).Columnで取得する
maxCol = Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, テーブル取得出力エリア.開始列).End(xlToRight).Column

Application.StatusBar = "Excelに展開中..."

ReDim Pj(Parse("Response")("Data").count - 1, maxCol)

i = 0
sortKey = 0

'取得データ件数分ループ
For Each DesStr In Parse("Response")("Data")
     
    Application.StatusBar = "Excelに展開中...(" & (i + 1) & "/" & Parse("Response")("Data").count & "件終了)"
    
    '出力エリアに設定した出力カラム数分ループ
    For j = 0 To maxCol - 1
        
        Select Case True
        
        '項目種類ごとに編集
        Case InStr(Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, j + 1).Value, "Class")
            Pj(i, j) = DesStr("ClassHash")(Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, j + 1).Value)
        
        Case InStr(Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, j + 1).Value, "Num")
            Pj(i, j) = DesStr("NumHash")(Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, j + 1).Value)
        
        Case InStr(Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, j + 1).Value, "Date")
            Pj(i, j) = DesStr("DateHash")(Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, j + 1).Value)
        
        Case InStr(Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, j + 1).Value, "Check")
            Pj(i, j) = DesStr("CheckHash")(Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, j + 1).Value)
        
        Case InStr(Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, j + 1).Value, "Description")
            Pj(i, j) = DesStr("DescriptionHash")(Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, j + 1).Value)
            
        Case Else
            
            'IDの場合はリンクをつける
            If Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, j + 1).Value = "IssueId" Or _
               Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, j + 1).Value = "ResultId" Then
               
               Pj(i, j) = "=HYPERLINK(""" & Worksheets(sheetName).Cells(3, 2).Value & "/items/" & _
                    DesStr(Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, j + 1).Value) & _
                    "/"",""" & DesStr(Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, j + 1).Value) & """)"
            
            Else
                
               Pj(i, j) = DesStr(Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行, j + 1).Value)
            
            End If
            
        End Select
        
        'ソート用のカラムを取得しとく
        If Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行 - 1, j + 1).Value <> "" Then
            sortKey = j + 1
            sortOrder = Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行 - 1, j + 1).Value
            
        End If
        
        DoEvents
    
    Next j
        
     i = i + 1
DoEvents
Next

Worksheets(sheetName).Range(Worksheets(sheetName).Cells(テーブル取得出力エリア.開始行 + 1, テーブル取得出力エリア.開始列), _
                    Worksheets(sheetName).Cells(Parse("Response")("Data").count + テーブル取得出力エリア.開始行, maxCol)) = Pj

If sortKey > 0 Then
    'エクセル機能でソート
    Call シートソート(sheetName, テーブル取得出力エリア.開始行 + 1, テーブル取得出力エリア.開始列, _
                        maxCol, テーブル取得出力エリア.開始列, sortKey, sortOrder)
End If

Application.ScreenUpdating = True

Application.StatusBar = False


End Sub

Sub テーブル登録メイン()

Dim sheetName As String
Dim maxCol As Long
Dim maxRow As Long
Dim j As Long
Dim i As Long
Dim filterFirst As Boolean
Dim classHashFirst As Boolean
Dim numHashFirst As Boolean
Dim dateHashFirst As Boolean

Dim strUrl As String

Dim Parse As Object
Dim DesStr As Variant
Dim Pj() As Variant

filterFirst = True
classHashFirst = True
numHashFirst = True
dateHashFirst = True

'アクティブシート名を取得
sheetName = ActiveSheet.Name

'リクエストのためのjson用Dictionaryを用意する
Set CreateRecordJson = New Dictionary
Set CreateRecordJson = SetApiKey(GetEnvironmentApiKey, CreateRecordJson)

'入力エリアに設定された条件値の数をEnd(xlToRight).Columnで取得する
maxCol = Worksheets(sheetName).Cells(テーブル登録入力エリア.開始行, テーブル登録入力エリア.開始列).End(xlToRight).Column

'設定された条件値の分ループして、リクエストするjson用Dictionaryを生成する
For j = 1 To maxCol

        Select Case True
        
        '項目種類ごとに編集する（Check項目には対応できていません！！）
        Case InStr(Worksheets(sheetName).Cells(テーブル登録入力エリア.開始行, j).Value, "Class")
            
            If classHashFirst = True Then
                '初回だけ作る
                Set CreateRecordJson = SetClassHash(CreateRecordJson)
                classHashFirst = False
            End If
            
            Set CreateRecordJson = SetClassHashItem(Worksheets(sheetName).Cells(テーブル登録入力エリア.開始行, j), _
                                                    Worksheets(sheetName).Cells(テーブル登録入力エリア.開始行 + 1, j), _
                                                    CreateRecordJson)
            
        Case InStr(Worksheets(sheetName).Cells(テーブル登録入力エリア.開始行, j + 1).Value, "Num")
        
            If numHashFirst = True Then
                '初回だけ作る
                Set CreateRecordJson = SetNumHash(CreateRecordJson)
                numHashFirst = False
            End If
            
            Set CreateRecordJson = SetNumHashItem(Worksheets(sheetName).Cells(テーブル登録入力エリア.開始行, j), _
                                                    Worksheets(sheetName).Cells(テーブル登録入力エリア.開始行 + 1, j), _
                                                    CreateRecordJson)
        
        Case InStr(Worksheets(sheetName).Cells(テーブル登録入力エリア.開始行, j + 1).Value, "Date")
        
            If dateHashFirst = True Then
                '初回だけ作る
                Set CreateRecordJson = SetDateHash(CreateRecordJson)
                dateHashFirst = False
            End If
            
            Set CreateRecordJson = SetDateHashItem(Worksheets(sheetName).Cells(テーブル登録入力エリア.開始行, j), _
                                                    Worksheets(sheetName).Cells(テーブル登録入力エリア.開始行 + 1, j), _
                                                    CreateRecordJson)

        Case Else
            Set CreateRecordJson = SetItem(Worksheets(sheetName).Cells(テーブル登録入力エリア.開始行, j), _
                                                    Worksheets(sheetName).Cells(テーブル登録入力エリア.開始行 + 1, j), _
                                                    CreateRecordJson)
        
        End Select
        
Next j

'URLを生成
If Worksheets(sheetName).Cells(3, 2).Value & Worksheets(sheetName).Cells(3, 2).Value = "" Then
'登録IDが無ければ新規

    strUrl = Worksheets(sheetName).Cells(4, 2).Value & "/api/items/" & Worksheets(sheetName).Cells(2, 2).Value & "/create"

Else
'存在したら更新

    strUrl = Worksheets(sheetName).Cells(4, 2).Value & "/api/items/" & Worksheets(sheetName).Cells(3, 2).Value & "/update"

End If

'HTTPリクエスト発行
Set objHTTP = HTTPリクエスト発行(strUrl, CreateRecordJson)
    
'レスポンスjsonをパース
Set Parse = JsonConverter.ParseJson(objHTTP.responseText)

'エクセルに結果を転記
Worksheets(sheetName).Cells(6, 2).Value = objHTTP.statusText & "(Status:" & objHTTP.Status & ")"
Worksheets(sheetName).Cells(7, 2).Value = "=HYPERLINK(""" & Worksheets(sheetName).Cells(4, 2).Value & _
    "/items/" & Parse("Id") & "/"",""" & Parse("Id") & """)"

End Sub


