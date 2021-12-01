Attribute VB_Name = "ExcelControl"
Option Explicit

'************************
'�G�N�Z���V�[�g���`����
'************************

Sub �V�[�g�N���A(�ΏۃV�[�g As String, �J�n�s As Long, �J�n�� As Long, �Ώۗ� As Long, �s�J�E���g����� As Long)
'�e�[�u���擾���A�o�̓G���A�̒l���N���A

Dim maxRow As Long

If Worksheets(�ΏۃV�[�g).Cells(�J�n�s, �s�J�E���g�����).Value = "" Then
    Exit Sub
End If

If Worksheets(�ΏۃV�[�g).Cells(�J�n�s + 1, �s�J�E���g�����).Value = "" Then
    
    maxRow = �J�n�s

Else
    
    maxRow = Worksheets(�ΏۃV�[�g).Cells(�J�n�s, �s�J�E���g�����).End(xlDown).Row

End If

    Worksheets(�ΏۃV�[�g).Range(Worksheets(�ΏۃV�[�g).Cells(�J�n�s, �J�n��), Worksheets(�ΏۃV�[�g).Cells(maxRow, �J�n�� + �Ώۗ� - 1)).Select
    Selection.ClearContents
    
    With Selection.Font
        .Name = "���S�V�b�N"
        .FontStyle = "�W��"
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
    
    
    Worksheets(�ΏۃV�[�g).Cells(�J�n�s, �J�n��).Select
End Sub


Sub �V�[�g�\�[�g(�ΏۃV�[�g As String, �J�n�s As Long, �J�n�� As Long, �Ώۗ� As Long, �s�J�E���g����� As Long, �\�[�g�L�[�� As Long, ���я� As String)
'�e�[�u���擾���A�w�肪�������ꍇ�\�[�g����i�G�N�Z���̋@�\�Łj�A����������\�[�g�L�[�Ɏw��ł���̂͂P���ڂ̂�

Const asc As String = "Ascending"
Const dec As String = "Descending"


    Dim maxRow As Long
    
    maxRow = Worksheets(�ΏۃV�[�g).Cells(�J�n�s, �s�J�E���g�����).End(xlDown).Row
   
    'Columns("A:G").Select
    Worksheets(�ΏۃV�[�g).Range(Worksheets(�ΏۃV�[�g).Cells(�J�n�s, �J�n��), Worksheets(�ΏۃV�[�g).Cells(maxRow, �J�n�� + �Ώۗ� - 1)).Select
    
    ActiveWorkbook.Worksheets(�ΏۃV�[�g).Sort.SortFields.Clear
    
    If ���я� = "a" Then
        ActiveWorkbook.Worksheets(�ΏۃV�[�g).Sort.SortFields.Add2 key:=Range(Worksheets(�ΏۃV�[�g).Cells(�J�n�s, �\�[�g�L�[��), Worksheets(�ΏۃV�[�g).Cells(maxRow, �\�[�g�L�[��)), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    Else
    '�~��
        ActiveWorkbook.Worksheets(�ΏۃV�[�g).Sort.SortFields.Add2 key:=Range(Worksheets(�ΏۃV�[�g).Cells(�J�n�s, �\�[�g�L�[��), Worksheets(�ΏۃV�[�g).Cells(maxRow, �\�[�g�L�[��)), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    End If
    
    With ActiveWorkbook.Worksheets(�ΏۃV�[�g).Sort
        .SetRange Range(Worksheets(�ΏۃV�[�g).Cells(�J�n�s, �J�n��), Worksheets(�ΏۃV�[�g).Cells(maxRow, �J�n�� + �Ώۗ� - 1))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

