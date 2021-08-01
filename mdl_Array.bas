Attribute VB_Name = "mdl_Array"
'# mdl_Array
'## Summary
'配列を効果的に使うためのモジュール

Option Explicit

Function IsInitialized(arr As Variant) As Boolean
'配列が初期化されているかを確認する｡
'
'Parameters
'------
'- arr:Variant
'対象となる配列
'
'Returns
'------
'- Boolean
'    - True      :配列が初期化済み
'    - False     :配列に値が格納されている
    
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
'配列内の重複を削除する｡
'
'Parameters
'------
'- arr:Variant
'対象となる配列

Dim dic As Object                      '重複を除いた値を格納する
Dim i, j                                       'ループカウンタ
Dim iLen                                     '配列要素数
Dim arrEdit() As Variant             '編集後の配列

'__init__
    Set dic = CreateObject("Scripting.Dictionary")
    ReDim arrEdit(0)
    iLen = UBound(arr)

'__main__
    For i = 0 To iLen
        '配列に未登録の場合
        If (dic.Exists(arr(i)) = False) Then
            Call dic.Add(arr(i), arr(i))      '追加
            
            '重複がない値のみ編集後の配列に格納する
            arrEdit(UBound(arrEdit)) = arr(i)
            ReDim Preserve arrEdit(UBound(arrEdit) + 1)
        Else
        End If
    Next
    
    '配列に格納済みの場合
    If IsEmpty(arrEdit(0)) = False Then
       '余分な領域を削除
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
'2つの配列を統合する。

Dim newArray() As Variant
Dim i As Integer
Dim itemCounter As Integer

'__init__
    '配列の確認
    If IsInitialized(arrA) = True Then GoTo NextProc
    If IsInitialized(arrB) = True Then GoTo EndProc

    itemCounter = 0

'__main__
    '1 つ目の配列の内容を新しい配列に格納する
    For i = LBound(arrA) To UBound(arrA)
        ReDim Preserve newArray(itemCounter)
        newArray(itemCounter) = arrA(i)
        itemCounter = itemCounter + 1
    Next i
    
NextProc:
    '2つ目の配列の内容を新しい配列に格納する
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
