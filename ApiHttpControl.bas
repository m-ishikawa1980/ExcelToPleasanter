Attribute VB_Name = "ApiHttpControl"
Option Explicit

'********************************************************
'API発行のためのjson定義作成処理、HTTPリクエスト発行処理
'********************************************************

Function SetApiKey(ApiKey As String, jsonObject As Object) As Object

    'ApiKeyを設定
    jsonObject.Add "ApiVersion", "1.1"
    jsonObject.Add "ApiKey", ApiKey
    
    Set SetApiKey = jsonObject

End Function

Function SetView(jsonObject As Object) As Object

    'Viewを設定
    jsonObject.Add "View", New Dictionary
    Set SetView = jsonObject

End Function


Function SetColumnFilterHash(jsonObject As Object) As Object

    'SetColumnFilterHashを設定
    jsonObject("View").Add "ColumnFilterHash", New Dictionary
    Set SetColumnFilterHash = jsonObject

End Function

Function SetColumnFilterHashItem(key As String, Val As Variant, jsonObject As Object) As Object

    'SetColumnFilterHashに設定する条件とするキー（Key）と条件値（Val）を設定
    jsonObject("View")("ColumnFilterHash").Add key, "[" & Val & "]"
    Set SetColumnFilterHashItem = jsonObject

End Function


Function SetItem(key As String, Val As String, jsonObject As Object) As Object
    
    'レコード登録時のキー（Key）と登録値（Val）を設定
    jsonObject.Add key, Val
    Set SetItem = jsonObject

End Function



Function SetClassHash(jsonObject As Object) As Object
    
    'レコード登録時のClassHashを設定
    jsonObject.Add "ClassHash", New Dictionary
    Set SetClassHash = jsonObject

End Function

Function SetClassHashItem(key As String, Val As String, jsonObject As Object) As Object
    
    'レコード登録時のClassHashに設定するキー（Key）と登録値（Val）を設定
    jsonObject("ClassHash").Add key, Val
    Set SetClassHashItem = jsonObject

End Function


Function SetNumHash(jsonObject As Object) As Object

    'レコード登録時のNumHashを設定
    jsonObject.Add "NumHash", New Dictionary
    Set SetNumHash = jsonObject

End Function

Function SetNumHashItem(key As String, Val As String, jsonObject As Object) As Object
    
    'レコード登録時のNumHashに設定するキー（Key）と登録値（Val）を設定
    jsonObject("NumHash").Add key, Val
    Set SetNumHashItem = jsonObject

End Function

Function SetDateHash(jsonObject As Object) As Object

    'レコード登録時のDateHashを設定
    jsonObject.Add "DateHash", New Dictionary
    Set SetDateHash = jsonObject

End Function

Function SetDateHashItem(key As String, Val As String, jsonObject As Object) As Object
    
    'レコード登録時のDateHashに設定するキー（Key）と登録値（Val）を設定
    jsonObject("DateHash").Add key, Val
    Set SetDateHashItem = jsonObject

End Function

Function SetSorterHash(jsonObject As Object) As Object

    'レコード登録時のDateHashに設定するキー（Key）と登録値（Val）を設定
    jsonObject.Add "ColumnSorterHash", New Dictionary
    Set SetSorterHash = jsonObject

End Function

Function SetSorterHashItem(key As String, Val As String, jsonObject As Object) As Object

    jsonObject("ColumnSorterHash").Add key, Val
    Set SetSorterHashItem = jsonObject

End Function

Function HTTPリクエスト発行(url As String, jsonObject As Object) As Object

    ' HTTPリクエスト発行
    Dim objHTTP As Object
    Set objHTTP = CreateObject("msxml2.xmlhttp")
    objHTTP.Open "POST", url, False
    
    objHTTP.setRequestHeader "Content-Type", "application/json;charset=utf-8"
    objHTTP.send JsonConverter.ConvertToJson(jsonObject)
    
    Set HTTPリクエスト発行 = objHTTP

End Function

