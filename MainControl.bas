Attribute VB_Name = "MainControl"
Option Explicit

'************************
'�e�[�u���擾�E�o�^����
'************************

Public GetMultiRecordJson As Object
Public GetSingleRecordJson As Object
Public CreateRecordJson As Object
Public objHTTP As Object

Public Enum �e�[�u���擾���̓G���A
    �J�n�s = 6
    �J�n�� = 1
End Enum

Public Enum �e�[�u���擾�o�̓G���A
    �J�n�s = 23
    �J�n�� = 1
End Enum

Public Enum �e�[�u���o�^���̓G���A
    �J�n�s = 10
    �J�n�� = 1
End Enum

Sub �A�N�e�B�u�V�[�g�N���A()

Dim maxCol As Long
Dim sheetName As String

sheetName = ActiveSheet.Name

maxCol = Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, �e�[�u���擾�o�̓G���A.�J�n��).End(xlToRight).Column

Call �V�[�g�N���A(sheetName, �e�[�u���擾�o�̓G���A.�J�n�s + 1, �e�[�u���擾�o�̓G���A.�J�n��, maxCol, �e�[�u���擾�o�̓G���A.�J�n��)

End Sub


Sub �e�[�u���擾���C��()

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

'���O�ɃV�[�g���N���A���Ă���
Call �A�N�e�B�u�V�[�g�N���A

filterFirst = True

'�A�N�e�B�u�V�[�g�����擾
sheetName = ActiveSheet.Name

'���N�G�X�g�̂��߂�json�pDictionary��p�ӂ���
Set GetMultiRecordJson = New Dictionary
Set GetMultiRecordJson = SetApiKey(GetEnvironmentApiKey, GetMultiRecordJson)

'���̓G���A�ɐݒ肳�ꂽ�����l�̐���End(xlToRight).Column�Ŏ擾����
maxCol = Worksheets(sheetName).Cells(�e�[�u���擾���̓G���A.�J�n�s, �e�[�u���擾���̓G���A.�J�n��).End(xlToRight).Column

Application.StatusBar = "Pleasanter�e�[�u���擾�J�n..."

'�ݒ肳�ꂽ�����l�̕����[�v���āA���N�G�X�g����json�pDictionary�𐶐�����
For j = 1 To maxCol

    If Worksheets(sheetName).Cells(�e�[�u���擾���̓G���A.�J�n�s + 1, j).Value <> "" Then
        
        If filterFirst = True Then
            '�t�B���^�[������Ώ��񂾂����
            Set GetMultiRecordJson = SetView(GetMultiRecordJson)
            '�t�B���^�[������Ώ��񂾂����
            Set GetMultiRecordJson = SetColumnFilterHash(GetMultiRecordJson)
            filterFirst = False
        End If
        
        '�֐��ɃL�[�ƒl��n����json�pDictionary�ɐݒ肷��
        Set GetMultiRecordJson = SetColumnFilterHashItem(Worksheets(sheetName).Cells(�e�[�u���擾���̓G���A.�J�n�s, j).Value, _
                                                    Worksheets(sheetName).Cells(�e�[�u���擾���̓G���A.�J�n�s + 1, j).Value, _
                                                    GetMultiRecordJson)
        
    End If

DoEvents

Next j

'���N�G�X�g�p��URL��ݒ肷��
strUrl = Worksheets(sheetName).Cells(3, 2).Value & "/api/items/" & Worksheets(sheetName).Cells(2, 2).Value & "/get"

'HTTP���N�G�X�g���s
Set objHTTP = HTTP���N�G�X�g���s(strUrl, GetMultiRecordJson)

'���X�|���Xjson���p�[�X
Set Parse = JsonConverter.ParseJson(objHTTP.responseText)

Application.StatusBar = "Pleasanter�e�[�u���擾�I��..."

'�擾�������`�F�b�N�A0���Ȃ珈���I��
If Parse("Response")("Data").count <= 0 Then
    
    Application.StatusBar = False
    
    Exit Sub

End If

'�o�̓G���A�ɐݒ肳�ꂽ�o�̓J�����̐���End(xlToRight).Column�Ŏ擾����
maxCol = Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, �e�[�u���擾�o�̓G���A.�J�n��).End(xlToRight).Column

Application.StatusBar = "Excel�ɓW�J��..."

ReDim Pj(Parse("Response")("Data").count - 1, maxCol)

i = 0
sortKey = 0

'�擾�f�[�^���������[�v
For Each DesStr In Parse("Response")("Data")
     
    Application.StatusBar = "Excel�ɓW�J��...(" & (i + 1) & "/" & Parse("Response")("Data").count & "���I��)"
    
    '�o�̓G���A�ɐݒ肵���o�̓J�����������[�v
    For j = 0 To maxCol - 1
        
        Select Case True
        
        '���ڎ�ނ��ƂɕҏW
        Case InStr(Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, j + 1).Value, "Class")
            Pj(i, j) = DesStr("ClassHash")(Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, j + 1).Value)
        
        Case InStr(Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, j + 1).Value, "Num")
            Pj(i, j) = DesStr("NumHash")(Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, j + 1).Value)
        
        Case InStr(Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, j + 1).Value, "Date")
            Pj(i, j) = DesStr("DateHash")(Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, j + 1).Value)
        
        Case InStr(Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, j + 1).Value, "Check")
            Pj(i, j) = DesStr("CheckHash")(Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, j + 1).Value)
        
        Case InStr(Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, j + 1).Value, "Description")
            Pj(i, j) = DesStr("DescriptionHash")(Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, j + 1).Value)
            
        Case Else
            
            'ID�̏ꍇ�̓����N������
            If Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, j + 1).Value = "IssueId" Or _
               Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, j + 1).Value = "ResultId" Then
               
               Pj(i, j) = "=HYPERLINK(""" & Worksheets(sheetName).Cells(3, 2).Value & "/items/" & _
                    DesStr(Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, j + 1).Value) & _
                    "/"",""" & DesStr(Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, j + 1).Value) & """)"
            
            Else
                
               Pj(i, j) = DesStr(Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s, j + 1).Value)
            
            End If
            
        End Select
        
        '�\�[�g�p�̃J�������擾���Ƃ�
        If Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s - 1, j + 1).Value <> "" Then
            sortKey = j + 1
            sortOrder = Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s - 1, j + 1).Value
            
        End If
        
        DoEvents
    
    Next j
        
     i = i + 1
DoEvents
Next

Worksheets(sheetName).Range(Worksheets(sheetName).Cells(�e�[�u���擾�o�̓G���A.�J�n�s + 1, �e�[�u���擾�o�̓G���A.�J�n��), _
                    Worksheets(sheetName).Cells(Parse("Response")("Data").count + �e�[�u���擾�o�̓G���A.�J�n�s, maxCol)) = Pj

If sortKey > 0 Then
    '�G�N�Z���@�\�Ń\�[�g
    Call �V�[�g�\�[�g(sheetName, �e�[�u���擾�o�̓G���A.�J�n�s + 1, �e�[�u���擾�o�̓G���A.�J�n��, _
                        maxCol, �e�[�u���擾�o�̓G���A.�J�n��, sortKey, sortOrder)
End If

Application.ScreenUpdating = True

Application.StatusBar = False


End Sub

Sub �e�[�u���o�^���C��()

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

'�A�N�e�B�u�V�[�g�����擾
sheetName = ActiveSheet.Name

'���N�G�X�g�̂��߂�json�pDictionary��p�ӂ���
Set CreateRecordJson = New Dictionary
Set CreateRecordJson = SetApiKey(GetEnvironmentApiKey, CreateRecordJson)

'���̓G���A�ɐݒ肳�ꂽ�����l�̐���End(xlToRight).Column�Ŏ擾����
maxCol = Worksheets(sheetName).Cells(�e�[�u���o�^���̓G���A.�J�n�s, �e�[�u���o�^���̓G���A.�J�n��).End(xlToRight).Column

'�ݒ肳�ꂽ�����l�̕����[�v���āA���N�G�X�g����json�pDictionary�𐶐�����
For j = 1 To maxCol

        Select Case True
        
        '���ڎ�ނ��ƂɕҏW����iCheck���ڂɂ͑Ή��ł��Ă��܂���I�I�j
        Case InStr(Worksheets(sheetName).Cells(�e�[�u���o�^���̓G���A.�J�n�s, j).Value, "Class")
            
            If classHashFirst = True Then
                '���񂾂����
                Set CreateRecordJson = SetClassHash(CreateRecordJson)
                classHashFirst = False
            End If
            
            Set CreateRecordJson = SetClassHashItem(Worksheets(sheetName).Cells(�e�[�u���o�^���̓G���A.�J�n�s, j), _
                                                    Worksheets(sheetName).Cells(�e�[�u���o�^���̓G���A.�J�n�s + 1, j), _
                                                    CreateRecordJson)
            
        Case InStr(Worksheets(sheetName).Cells(�e�[�u���o�^���̓G���A.�J�n�s, j + 1).Value, "Num")
        
            If numHashFirst = True Then
                '���񂾂����
                Set CreateRecordJson = SetNumHash(CreateRecordJson)
                numHashFirst = False
            End If
            
            Set CreateRecordJson = SetNumHashItem(Worksheets(sheetName).Cells(�e�[�u���o�^���̓G���A.�J�n�s, j), _
                                                    Worksheets(sheetName).Cells(�e�[�u���o�^���̓G���A.�J�n�s + 1, j), _
                                                    CreateRecordJson)
        
        Case InStr(Worksheets(sheetName).Cells(�e�[�u���o�^���̓G���A.�J�n�s, j + 1).Value, "Date")
        
            If dateHashFirst = True Then
                '���񂾂����
                Set CreateRecordJson = SetDateHash(CreateRecordJson)
                dateHashFirst = False
            End If
            
            Set CreateRecordJson = SetDateHashItem(Worksheets(sheetName).Cells(�e�[�u���o�^���̓G���A.�J�n�s, j), _
                                                    Worksheets(sheetName).Cells(�e�[�u���o�^���̓G���A.�J�n�s + 1, j), _
                                                    CreateRecordJson)

        Case Else
            Set CreateRecordJson = SetItem(Worksheets(sheetName).Cells(�e�[�u���o�^���̓G���A.�J�n�s, j), _
                                                    Worksheets(sheetName).Cells(�e�[�u���o�^���̓G���A.�J�n�s + 1, j), _
                                                    CreateRecordJson)
        
        End Select
        
Next j

'URL�𐶐�
If Worksheets(sheetName).Cells(3, 2).Value & Worksheets(sheetName).Cells(3, 2).Value = "" Then
'�o�^ID��������ΐV�K

    strUrl = Worksheets(sheetName).Cells(4, 2).Value & "/api/items/" & Worksheets(sheetName).Cells(2, 2).Value & "/create"

Else
'���݂�����X�V

    strUrl = Worksheets(sheetName).Cells(4, 2).Value & "/api/items/" & Worksheets(sheetName).Cells(3, 2).Value & "/update"

End If

'HTTP���N�G�X�g���s
Set objHTTP = HTTP���N�G�X�g���s(strUrl, CreateRecordJson)
    
'���X�|���Xjson���p�[�X
Set Parse = JsonConverter.ParseJson(objHTTP.responseText)

'�G�N�Z���Ɍ��ʂ�]�L
Worksheets(sheetName).Cells(6, 2).Value = objHTTP.statusText & "(Status:" & objHTTP.Status & ")"
Worksheets(sheetName).Cells(7, 2).Value = "=HYPERLINK(""" & Worksheets(sheetName).Cells(4, 2).Value & _
    "/items/" & Parse("Id") & "/"",""" & Parse("Id") & """)"

End Sub


