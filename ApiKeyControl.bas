Attribute VB_Name = "ApiKeyControl"
Option Explicit

'**********************************************************
'ApiKey�擾�����F���[�U���ϐ�"ApiKey"����ApiKey�̒l���擾
'**********************************************************

Function GetEnvironmentApiKey() As String
    
    Dim wsh As New IWshRuntimeLibrary.wshShell
    Dim env As IWshRuntimeLibrary.WshEnvironment
    Dim s
    Dim i
    Dim varName As String
    Dim ret As String
        
    varName = "ApiKey"
        
        
    '���ϐ����擾
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
