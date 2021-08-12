Attribute VB_Name = "mdl_Algo"
Option Explicit

Function LsDist(baseText As String, tryText As String) As Double
'# ������̔�r
'Levenshtein�����ŗގ��x�𑪒肵���v�x��Ԃ�
'Levenshtein����
'   ���̕�������������ύX����΁A�v�l������ɂȂ邩�񐔂ő���

'Arg
'Param1(String):    baseText    ��r���̕�����
'Param2(String):    tryText       ��r�Ώۂ̕�����

'Return(Double):    ������̈�v�x  min:0/Max:1

Dim matrix As Variant
Dim i As Integer, j As Integer, cost As Integer
Dim missCnt As Integer

    LsDist = 0
    
    If (baseText = tryText) Then
        LsDist = Format(1, "0.00")
        Exit Function
    End If
    If (Len(baseText) = 0) Then
        LsDist = Format(0, "0.00")
        Exit Function
    End If
    
    ReDim matrix(Len(baseText), Len(tryText))

    For i = 0 To Len(baseText)
        matrix(i, 0) = i
    Next i
    
    For j = 0 To Len(tryText)
        matrix(0, j) = j
    Next j
    
    For i = 1 To Len(baseText)
        For j = 1 To Len(tryText)
            cost = IIf(Mid$(baseText, i, 1) = Mid$(tryText, j, 1), 0, 1)
            matrix(i, j) = WorksheetFunction.Min(matrix(i - 1, j) + 1, matrix(i, j - 1) + 1, matrix(i - 1, j - 1) + cost)
            
                 'matrix(i - 1, j) + 1              '�v�f�̍폜
                 'matrix(i, j - 1) + 1              '�v�f�̑}��
                 'matrix(i - 1, j - 1) + cost    '�v�f�̒u��
        Next j
    Next i
    
    missCnt = matrix(Len(baseText), Len(tryText))
    
    '��v�x��Ԃ�
    LsDist = (missCnt / Len(baseText))
    LsDist = 1 - LsDist / Len(baseText)
    LsDist = Format(LsDist, "0.00")
    If LsDist < 0 Then LsDist = Format(0, "0.00")
End Function

Function GetFileFromFolder(ByVal folderPath As String)
'�w�肵���p�X�̃t�H���_���T�u�t�H���_���ɂ���t�@�C�������ׂĒT��

Dim fso As Object
Dim objfile As Object, objCFolder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
Dim filePath As String
    
'�t�@�C�����擾����
If IsExist(folderPath & "\*") = True Then
    For Each objfile In fso.getFolder(folderPath).Files
        Logging ("fileName   " & objfile.Name)
        filePath = objFilePath
    
        Call ����
    Next
End If

'�T�u�t�H���_���擾���A�T����i�߂�
If IsExist(folderPath) = True Then
    For Each objFolder In fso.getFolder(folderPath).subfolders
        Logging ("folderName     " & objCFolder.Name)
        Call GetFileFromFolder(folderPath & "\" & objCFolder.Name)
    Next
End If

errExit:
    Logging ("Err on GetFileFromFolder")
    Exit Function
End Function
