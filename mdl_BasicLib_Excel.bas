Attribute VB_Name = "mdl_BasicLib_Excel"
Option Explicit

Function rowLast(sheetName As String, column As Long) As Long
'最終行を求める

'Arg
'sheetName     検索するシート名
'column            検索する列番号

On Error GoTo errExit
    rowLast = Sheets(sheetName).Columns(column).Find(What:="*", _
                                                                                            LookIn:=xlFormulas, _
                                                                                            SearchOrder:=xlByRows, _
                                                                                            SearchDirection:=xlPrevious).row
                                                                                
    Exit Function
errExit:
    rowLast = 0
End Function

Function colLast(sheetName As String, row As Long) As Integer
'最終列を求める

'Arg
'sheetName     検索するシート名
'column            検索する列番号

On Error GoTo errExit
    colLast = Sheets(sheetName).Rows(row).Find(What:="*", _
                                                                                LookIn:=xlFormulas, _
                                                                                SearchOrder:=xlByColumns, _
                                                                                SearchDirection:=xlPrevious).column
                                                                                
    Exit Function
errExit:
    colLast = 0
End Function

Function SetWidth(Width As Integer, StartCol As Integer, LastCol As Integer)
'セルの幅を設定する

    Range(Columns(StartCol), Columns(LastCol)).ColumnWidth = Width
End Function

Function SetHight(Height As Integer, StartRow As Integer, LastRow As Integer)
'セルの高さを設定する

    Range(Rows(StartRow), Rows(LastRow)).RowHeight = Height
End Function

Function SetFileReadOnly()
'ファイルを読取り専用にする

On Error Resume Next
    ActiveWorkbook.Saved = True
    ActiveWorkbook.ChangeFileAccess (xlReadOnly)
End Function

Function SetFileReadWrite()
'ファイルを読取り専用を解除する

On Error Resume Next
    ActiveWorkbook.Saved = True
    ActiveWorkbook.ChangeFileAccess (xlReadWrite)
End Function

Function IsReadOnly()
'ファイルが読み取り専用か確認する

    IsReadOnly = ActiveWorkbook.ReadOnly
End Function

Function KillOwn()
'プログラムファイル自身を削除する
'読み取り専用で開き、読み取り元ファイルを削除する

    Call SetFileReadOnly
    Kill ThisWorkbook.FullName
End Function

Function CellColor(rngR As Range, _
                                intColorR As Long, intColorG As Long, intColorB As Long, _
                                Optional dblTintAndShade As Double)
'RGBスケールでセルの色を変える

'RGBパラメータ
'   https://ironodata.info/

    With rngR.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = RGB(intColorR, intColorG, intColorB)
                    .TintAndShade = dblTintAndShade
                    .PatternTintAndShade = 0
    End With
End Function
                                
Function ClearColor(rngR As Range)
'セルの色設定をクリアする
    With rngR.Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
    End With
End Function

Function GetFilePath() As String
'ダイアログからファイルを選択し、ファイルパスを取得する

Dim FilePath As String

    FilePath = Application.GetOpenFilename
    
    If FilePath = "False" Then
        GetFilePath = False
        Exit Function
    Else
    End If
    
    GetFilePath = FilePath
End Function

Function GetDirPath() As String
'ダイアログからフォルダを選択し、パスを取得する

Dim FilePath As String

    FilePath = Application.FileDialog(msoFileDialogFolderPicker).Show
    
    If FilePath = "False" Then
        GetDirPath = False
        Exit Function
    Else
    End If
    
    GetDirPath = FilePath
End Function

Function GetFileName(FilePath As String, Optional ExtensionFlg As Boolean = True) As String
'引数で指定されたファイル名を取得する

'Arg     ExtensionFlg
'    True：Returnに拡張子あり
'     False:Returnに拡張子なし

    If ExtensionFlg = True Then
        GetFileName = Mid(FilePath, InStrRev(FilePath, "\") + 1)
    Else
        GetFileName = Replace(FilePath, Left(FilePath, InStrRev(FilePath, "\")), "")
        GetFileName = Replace(GetFileName, GetExtension(FilePath), "")
        GetFileName = Left(GetFileName, Len(GetFileName) - 1)
    End If
End Function

Function GetExtension(FilePath As String) As String
'引数で指定されたファイルの拡張子を返す

Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetExtension = FSO.GetExtensionName(FilePath)
End Function

Function GetPCName() As String
'PCの名前を取得する

Dim WshNetworkObject As Object

    Set WshNetworkObject = CreateObject("Wscript.Network")
    GetPCName = WshNetworkObject.ComputerName
End Function

'Function GetUserID() As String
''ユーザIDを取得する
'
'Dim objSysInfo As Object
'Dim objUser As Object
'
'    Set objSysInfo = CreateObject("ADSysteminfo")
'    Set objUser = GetObject("LDAP://" & objSysInfo.UserName)
'    GetUserID = objUser.Name
'End Function

Function IsExist(FilePath As String) As Boolean
'ファイル、ディレクトリの存在確認をする

    If InStrRev(FilePath, ",") > 0 Then
        If Dir(FilePath) <> "" Then
            IsExist = True
        Else
            IsExist = False
        End If
    Else
        If Dir(FilePath, vbDirectory) <> "" Then
            IsExist = True
        Else
            IsExist = False
        End If
    End If
End Function

Function IsAppActivate(Title As String, Optional WaitTime As Single = 3) As Boolean
'指定のウィンドウがアクティブか確認する

Dim StartTime As Single
Dim ElapesedTime As Single

On Error Resume Next
    
    '開始時間セット
    StartTime = Timer
    
    '一定時間の間、一定間隔ごとに処理を試みる
    Do While ElapesedTime < WaitTime     '経過時間 <= 間隔(秒)
        
        '対象画面が起動しているか確認する
        AppActivate (Title)
        
        If Err = 0 Then
            IsAppActivate = True
            Exit Function
        Else
        End If
        
        WaitTimeFor (0.1)                           '処理間隔
        ElapesedTime = Timer - StartTime     '経過時間算出

    Loop
    
    '画面が見つからないときはFlaseで返す
    On Error GoTo 0
        IsAppActivate = False
End Function

Function OpenDir(DirPath As String, Optional WaitTime As Single = 0.7)
'フォルダパスを指定してディレクトリを開く

Dim StartTime As Single

    If IsExist(fokderpath) = False Then GoTo errExist
    
    Shell "C:\Windows\Explore.exe" & FolderPath, vbNormalFocus
    WaitTimeFor (WaitTime)
    StartTime = Timer
    
    'フォルダが表示されるまで待つ
    '５秒待って表示されなかったらエラーを出す
    Do Until IsAppActivate(GetFileName(FolderPath)) = True
        DoEvents
        
        If Timer - StartTime > 5 Then
            OpenDir = False
            Exit Do
        Else
        End If
        
    Loop
    
    OpenDir = True
    Exit Function

errExit:
    OpenDir = False
End Function

Function strTime(Time As Date) As String
'時刻を文字列で"hhnn"で返す

    strTime = Format(Time, "hhnn")
End Function

Function WaitTimeFor(WaitSecounds As Single)
'指定の秒数処理を待機させる

Dim StartTime As String
    StartTime = Timer
    
    Do While Timer < StartTime + WaitSecounds
        DoEvents
    Loop
End Function


'=== Remarks ===
'Sort
'Sub sort()
'Dim sheetName As String
'Dim ranges As Range
'    Call Sheets(sheetName).UsedRange.sort(key1:=Range(ranges), Order1:=xlAscending, Header:=xlYes)
'End Sub
