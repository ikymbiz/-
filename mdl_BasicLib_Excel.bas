Attribute VB_Name = "mdl_BasicLib_Excel"
Option Explicit

Function rowLast(sheetName As String, column As Long) As Long
'�ŏI�s�����߂�

'Arg
'sheetName     ��������V�[�g��
'column            ���������ԍ�

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
'�ŏI������߂�

'Arg
'sheetName     ��������V�[�g��
'column            ���������ԍ�

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
'�Z���̕���ݒ肷��

    Range(Columns(StartCol), Columns(LastCol)).ColumnWidth = Width
End Function

Function SetHight(Height As Integer, StartRow As Integer, LastRow As Integer)
'�Z���̍�����ݒ肷��

    Range(Rows(StartRow), Rows(LastRow)).RowHeight = Height
End Function

Function SetFileReadOnly()
'�t�@�C����ǎ���p�ɂ���

On Error Resume Next
    ActiveWorkbook.Saved = True
    ActiveWorkbook.ChangeFileAccess (xlReadOnly)
End Function

Function SetFileReadWrite()
'�t�@�C����ǎ���p����������

On Error Resume Next
    ActiveWorkbook.Saved = True
    ActiveWorkbook.ChangeFileAccess (xlReadWrite)
End Function

Function IsReadOnly()
'�t�@�C�����ǂݎ���p���m�F����

    IsReadOnly = ActiveWorkbook.ReadOnly
End Function

Function KillOwn()
'�v���O�����t�@�C�����g���폜����
'�ǂݎ���p�ŊJ���A�ǂݎ�茳�t�@�C�����폜����

    Call SetFileReadOnly
    Kill ThisWorkbook.FullName
End Function

Function CellColor(rngR As Range, _
                                intColorR As Long, intColorG As Long, intColorB As Long, _
                                Optional dblTintAndShade As Double)
'RGB�X�P�[���ŃZ���̐F��ς���

'RGB�p�����[�^
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
'�Z���̐F�ݒ���N���A����
    With rngR.Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
    End With
End Function

Function GetFilePath() As String
'�_�C�A���O����t�@�C����I�����A�t�@�C���p�X���擾����

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
'�_�C�A���O����t�H���_��I�����A�p�X���擾����

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
'�����Ŏw�肳�ꂽ�t�@�C�������擾����

'Arg     ExtensionFlg
'    True�FReturn�Ɋg���q����
'     False:Return�Ɋg���q�Ȃ�

    If ExtensionFlg = True Then
        GetFileName = Mid(FilePath, InStrRev(FilePath, "\") + 1)
    Else
        GetFileName = Replace(FilePath, Left(FilePath, InStrRev(FilePath, "\")), "")
        GetFileName = Replace(GetFileName, GetExtension(FilePath), "")
        GetFileName = Left(GetFileName, Len(GetFileName) - 1)
    End If
End Function

Function GetExtension(FilePath As String) As String
'�����Ŏw�肳�ꂽ�t�@�C���̊g���q��Ԃ�

Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetExtension = FSO.GetExtensionName(FilePath)
End Function

Function GetPCName() As String
'PC�̖��O���擾����

Dim WshNetworkObject As Object

    Set WshNetworkObject = CreateObject("Wscript.Network")
    GetPCName = WshNetworkObject.ComputerName
End Function

'Function GetUserID() As String
''���[�UID���擾����
'
'Dim objSysInfo As Object
'Dim objUser As Object
'
'    Set objSysInfo = CreateObject("ADSysteminfo")
'    Set objUser = GetObject("LDAP://" & objSysInfo.UserName)
'    GetUserID = objUser.Name
'End Function

Function IsExist(FilePath As String) As Boolean
'�t�@�C���A�f�B���N�g���̑��݊m�F������

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
'�w��̃E�B���h�E���A�N�e�B�u���m�F����

Dim StartTime As Single
Dim ElapesedTime As Single

On Error Resume Next
    
    '�J�n���ԃZ�b�g
    StartTime = Timer
    
    '��莞�Ԃ̊ԁA���Ԋu���Ƃɏ��������݂�
    Do While ElapesedTime < WaitTime     '�o�ߎ��� <= �Ԋu(�b)
        
        '�Ώۉ�ʂ��N�����Ă��邩�m�F����
        AppActivate (Title)
        
        If Err = 0 Then
            IsAppActivate = True
            Exit Function
        Else
        End If
        
        WaitTimeFor (0.1)                           '�����Ԋu
        ElapesedTime = Timer - StartTime     '�o�ߎ��ԎZ�o

    Loop
    
    '��ʂ�������Ȃ��Ƃ���Flase�ŕԂ�
    On Error GoTo 0
        IsAppActivate = False
End Function

Function OpenDir(DirPath As String, Optional WaitTime As Single = 0.7)
'�t�H���_�p�X���w�肵�ăf�B���N�g�����J��

Dim StartTime As Single

    If IsExist(fokderpath) = False Then GoTo errExist
    
    Shell "C:\Windows\Explore.exe" & FolderPath, vbNormalFocus
    WaitTimeFor (WaitTime)
    StartTime = Timer
    
    '�t�H���_���\�������܂ő҂�
    '�T�b�҂��ĕ\������Ȃ�������G���[���o��
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
'�����𕶎����"hhnn"�ŕԂ�

    strTime = Format(Time, "hhnn")
End Function

Function WaitTimeFor(WaitSecounds As Single)
'�w��̕b��������ҋ@������

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
