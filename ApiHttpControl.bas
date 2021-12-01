Attribute VB_Name = "ApiHttpControl"
Option Explicit

'********************************************************
'API���s�̂��߂�json��`�쐬�����AHTTP���N�G�X�g���s����
'********************************************************

Function SetApiKey(ApiKey As String, jsonObject As Object) As Object

    'ApiKey��ݒ�
    jsonObject.Add "ApiVersion", "1.1"
    jsonObject.Add "ApiKey", ApiKey
    
    Set SetApiKey = jsonObject

End Function

Function SetView(jsonObject As Object) As Object

    'View��ݒ�
    jsonObject.Add "View", New Dictionary
    Set SetView = jsonObject

End Function


Function SetColumnFilterHash(jsonObject As Object) As Object

    'SetColumnFilterHash��ݒ�
    jsonObject("View").Add "ColumnFilterHash", New Dictionary
    Set SetColumnFilterHash = jsonObject

End Function

Function SetColumnFilterHashItem(key As String, Val As Variant, jsonObject As Object) As Object

    'SetColumnFilterHash�ɐݒ肷������Ƃ���L�[�iKey�j�Ə����l�iVal�j��ݒ�
    jsonObject("View")("ColumnFilterHash").Add key, "[" & Val & "]"
    Set SetColumnFilterHashItem = jsonObject

End Function


Function SetItem(key As String, Val As String, jsonObject As Object) As Object
    
    '���R�[�h�o�^���̃L�[�iKey�j�Ɠo�^�l�iVal�j��ݒ�
    jsonObject.Add key, Val
    Set SetItem = jsonObject

End Function



Function SetClassHash(jsonObject As Object) As Object
    
    '���R�[�h�o�^����ClassHash��ݒ�
    jsonObject.Add "ClassHash", New Dictionary
    Set SetClassHash = jsonObject

End Function

Function SetClassHashItem(key As String, Val As String, jsonObject As Object) As Object
    
    '���R�[�h�o�^����ClassHash�ɐݒ肷��L�[�iKey�j�Ɠo�^�l�iVal�j��ݒ�
    jsonObject("ClassHash").Add key, Val
    Set SetClassHashItem = jsonObject

End Function


Function SetNumHash(jsonObject As Object) As Object

    '���R�[�h�o�^����NumHash��ݒ�
    jsonObject.Add "NumHash", New Dictionary
    Set SetNumHash = jsonObject

End Function

Function SetNumHashItem(key As String, Val As String, jsonObject As Object) As Object
    
    '���R�[�h�o�^����NumHash�ɐݒ肷��L�[�iKey�j�Ɠo�^�l�iVal�j��ݒ�
    jsonObject("NumHash").Add key, Val
    Set SetNumHashItem = jsonObject

End Function

Function SetDateHash(jsonObject As Object) As Object

    '���R�[�h�o�^����DateHash��ݒ�
    jsonObject.Add "DateHash", New Dictionary
    Set SetDateHash = jsonObject

End Function

Function SetDateHashItem(key As String, Val As String, jsonObject As Object) As Object
    
    '���R�[�h�o�^����DateHash�ɐݒ肷��L�[�iKey�j�Ɠo�^�l�iVal�j��ݒ�
    jsonObject("DateHash").Add key, Val
    Set SetDateHashItem = jsonObject

End Function

Function SetSorterHash(jsonObject As Object) As Object

    '���R�[�h�o�^����DateHash�ɐݒ肷��L�[�iKey�j�Ɠo�^�l�iVal�j��ݒ�
    jsonObject.Add "ColumnSorterHash", New Dictionary
    Set SetSorterHash = jsonObject

End Function

Function SetSorterHashItem(key As String, Val As String, jsonObject As Object) As Object

    jsonObject("ColumnSorterHash").Add key, Val
    Set SetSorterHashItem = jsonObject

End Function

Function HTTP���N�G�X�g���s(url As String, jsonObject As Object) As Object

    ' HTTP���N�G�X�g���s
    Dim objHTTP As Object
    Set objHTTP = CreateObject("msxml2.xmlhttp")
    objHTTP.Open "POST", url, False
    
    objHTTP.setRequestHeader "Content-Type", "application/json;charset=utf-8"
    objHTTP.send JsonConverter.ConvertToJson(jsonObject)
    
    Set HTTP���N�G�X�g���s = objHTTP

End Function

