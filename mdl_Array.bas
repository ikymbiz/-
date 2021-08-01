Attribute VB_Name = "mdl_Array"
'# mdl_Array
'## Summary
'�z������ʓI�Ɏg�����߂̃��W���[��

Option Explicit

Function IsInitialized(arr As Variant) As Boolean
'�z�񂪏���������Ă��邩���m�F����
'
'Parameters
'------
'- arr:Variant
'�ΏۂƂȂ�z��
'
'Returns
'------
'- Boolean
'    - True      :�z�񂪏������ς�
'    - False     :�z��ɒl���i�[����Ă���
    
Dim Length As Long
On Error GoTo Not_Err
    
    Length = UBound(arr)
    IsInitialized = False
    Exit Function

Not_Err:
    IsInitialized = True
End Function

'Private Sub TEST_IsInitialized()
'Dim arrA(1) As Variant
'Dim arrB() As Variant
'
'    arrA(0) = 2
'    Debug.Print IsInitialized(arrA)
'    Debug.Print IsInitialized(arrB)
'End Sub

Function DeleteSameValue(arr As Variant) As Variant
'�z����̏d�����폜����
'
'Parameters
'------
'- arr:Variant
'�ΏۂƂȂ�z��

Dim dic As Object                      '�d�����������l���i�[����
Dim i, j                                       '���[�v�J�E���^
Dim iLen                                     '�z��v�f��
Dim arrEdit() As Variant             '�ҏW��̔z��

'__init__
    Set dic = CreateObject("Scripting.Dictionary")
    ReDim arrEdit(0)
    iLen = UBound(arr)

'__main__
    For i = 0 To iLen
        '�z��ɖ��o�^�̏ꍇ
        If (dic.Exists(arr(i)) = False) Then
            Call dic.Add(arr(i), arr(i))      '�ǉ�
            
            '�d�����Ȃ��l�̂ݕҏW��̔z��Ɋi�[����
            arrEdit(UBound(arrEdit)) = arr(i)
            ReDim Preserve arrEdit(UBound(arrEdit) + 1)
        Else
        End If
    Next
    
    '�z��Ɋi�[�ς݂̏ꍇ
    If IsEmpty(arrEdit(0)) = False Then
       '�]���ȗ̈���폜
       ReDim Preserve arrEdit(UBound(arrEdit) - 1)
    End If
    
    DeleteSameValue = arrEdit
End Function

'Private Sub TEST_DeleteSameValue()
'Dim arr(5) As Variant, arrB As Variant
'Dim i As Integer
'
'    arr(0) = "AAA"
'    arr(1) = "AAC"
'    arr(2) = "AAC"
'
'    For i = LBound(arr) To UBound(arr)
'        Debug.Print (arr(i))
'    Next
'
'    Stop
'
'    arrB = DeleteSameValue(arr())
'
'    For i = LBound(arrB) To UBound(arrB)
'        Debug.Print (arrB(i))
'    Next
'End Sub

Function MergeArray(arrA As Variant, arrB As Variant) As Variant
'2�̔z��𓝍�����B

Dim newArray() As Variant
Dim i As Integer
Dim itemCounter As Integer

'__init__
    '�z��̊m�F
    If IsInitialized(arrA) = True Then GoTo NextProc
    If IsInitialized(arrB) = True Then GoTo EndProc

    itemCounter = 0

'__main__
    '1 �ڂ̔z��̓��e��V�����z��Ɋi�[����
    For i = LBound(arrA) To UBound(arrA)
        ReDim Preserve newArray(itemCounter)
        newArray(itemCounter) = arrA(i)
        itemCounter = itemCounter + 1
    Next i
    
NextProc:
    '2�ڂ̔z��̓��e��V�����z��Ɋi�[����
    For i = LBound(arrB) To UBound(arrB)
        ReDim Preserve newArray(itemCounter)
        newArray(itemCounter) = arrB(i)
        itemCounter = itemCounter + 1
    Next i

EndProc:
    MergeArray = newArray
End Function

'Private Sub TEST_MergeArray()
'Dim arr As Variant, arrA(2) As Variant, arrB(2) As Variant
'Dim i As Integer
'
'    arrA(0) = "AAA"
'    arrA(1) = "AAC"
'    arrA(2) = "AAC"
'
'    arrB(0) = "AB"
'    arrB(1) = "AC"
'    arrB(2) = "AD"
'
'    For i = LBound(arrA) To UBound(arrA)
'        Debug.Print (arrA(i))
'    Next
'
'    Stop
'
'    For i = LBound(arrB) To UBound(arrB)
'        Debug.Print (arrB(i))
'    Next
'
'    Stop
'    Debug.Print "----"
'
'    arr = MergeArray(arrA, arrB)
'
'    For i = LBound(arr) To UBound(arr)
'        Debug.Print (arr(i))
'    Next
'End Sub
