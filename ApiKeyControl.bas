Attribute VB_Name = "ApiKeyControl"
Option Explicit

'**********************************************************
'ApiKey取得処理：ユーザ環境変数"ApiKey"からApiKeyの値を取得
'**********************************************************

Function GetEnvironmentApiKey() As String
    
    Dim wsh As New IWshRuntimeLibrary.wshShell
    Dim env As IWshRuntimeLibrary.WshEnvironment
    Dim s
    Dim i
    Dim varName As String
    Dim ret As String
        
    varName = "ApiKey"
        
        
    '環境変数を取得
    Set env = wsh.Environment("User")
        
    If env(varName) <> "" Then
    
        ret = env(varName)
        
    Else
    
        ret = ""
        
    End If
    
    Set env = Nothing
    Set wsh = Nothing
    
    GetEnvironmentApiKey = ret
    
End Function
