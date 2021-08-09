Attribute VB_Name = "mdl_Log"
Option Explicit

Function Logging(varData As Variant)
'# 目的
'引数をイミディエイトウィンドウに表示するとともにテキストデータとして保存する

'## 処理内容
'- 実行ファイルの直下に"_Log"ファイルを作成
'- logファイルを作成
'- logファイルに記述内容を保存
'- 記述内容をイミディエイトウィンドウに表示する

Dim strPath As String
Dim lngFileNum As Long
Dim strLogFile As String

On Error Resume Next
    strPath = ThisWorkbook.path
    strPath = strPath & "\_Log"
On Error GoTo errExit
    If IsExist(strPath) = False Then MkDir strPath
    
    strLogFile = strPath & "\" & Format(Date, "YYYY.MM.DD") & ".log"
    lngFileNum = FreeFile()

    Open strLogFile For Append As #lngFileNum
    Print #lngFileNum, varData
    Close #lngFileNum
    Debug.Print varData
Exit Function
errExit:
    strLogFile = ThisWorkbook.path & "\Err.log"
    lngFileNum = FreeFile()
    
    Open strLogFile For Append As #lngFileNum
    Print #lngFileNum, varData
    Close #lngFileNum
    Debug.Print "ERR    " & ThisWorkbook.FullName & "Logging    " & varData
End Function

Function ErrLog(fncName As String) As String
'エラーログを出力する

Dim str As String
    str = str & ">  Err" & vbCrLf
    str = str & "    " & Now & vbCrLf
    str = str & ">>ErrNo      " & Err.Number
    str = str & ">>               " & Err.HelpContext & vbCrLf
    str = str & ">>               " & Err.HelpFile & vbCrLf
    str = str & ">>               " & Err.Description & vbCrLf
    str = str & ">>fncName " & fncName & vbCrLf
    str = str & ">>  Line       " & Erl
    

End Function


