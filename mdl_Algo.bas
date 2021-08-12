Attribute VB_Name = "mdl_Algo"
Option Explicit

Function LsDist(baseText As String, tryText As String) As Double
'# 文字列の比較
'Levenshtein距離で類似度を測定し､一致度を返す
'Levenshtein距離
'   元の文字列を何文字変更すれば、思考文字列になるか回数で測る

'Arg
'Param1(String):    baseText    比較元の文字列
'Param2(String):    tryText       比較対象の文字列

'Return(Double):    文字列の一致度  min:0/Max:1

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
            
                 'matrix(i - 1, j) + 1              '要素の削除
                 'matrix(i, j - 1) + 1              '要素の挿入
                 'matrix(i - 1, j - 1) + cost    '要素の置換
        Next j
    Next i
    
    missCnt = matrix(Len(baseText), Len(tryText))
    
    '一致度を返す
    LsDist = (missCnt / Len(baseText))
    LsDist = 1 - LsDist / Len(baseText)
    LsDist = Format(LsDist, "0.00")
    If LsDist < 0 Then LsDist = Format(0, "0.00")
End Function

Function GetFileFromFolder(ByVal folderPath As String)
'指定したパスのフォルダ→サブフォルダ内にあるファイルをすべて探す

Dim fso As Object
Dim objfile As Object, objCFolder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
Dim filePath As String
    
'ファイルを取得する
If IsExist(folderPath & "\*") = True Then
    For Each objfile In fso.getFolder(folderPath).Files
        Logging ("fileName   " & objfile.Name)
        filePath = objFilePath
    
        Call 処理
    Next
End If

'サブフォルダを取得し、探索を進める
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
