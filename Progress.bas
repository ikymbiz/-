Attribute VB_Name = "Progress"
Option Explicit

'----------------------------------------
'ÅEêiíªï\é¶
'----------------------------------------
Public Sub Application_StatusBar_Progress(ByVal Message As String, _
ByVal Delimiter As String, _
ByVal StartValue As Long, ByVal Value As Long, ByVal EndValue As Long, _
Optional ReverseFlag As Boolean = False)

    Application.StatusBar = ProgressText( _
        Message, Delimiter, _
        StartValue, Value, EndValue, _
        ReverseFlag)

End Sub

Public Function ProgressText( _
ByVal Message As String, ByVal Delimiter As String, _
ByVal StartValue As Long, ByVal Value As Long, ByVal EndValue As Long, _
Optional ReverseFlag As Boolean = False, _
Optional PercentVisible As Boolean = True)
    Dim Result As String
    If ReverseFlag = False Then
        Result = _
            Message + Delimiter + _
            CStr(Value - StartValue + 1) + "/" + _
            CStr(EndValue - StartValue + 1)
        If PercentVisible Then
            Result = Result + Delimiter + _
            CStr(Format((Value - StartValue + 1) / (EndValue - StartValue + 1) * 100, "0.00")) + "%"
        End If

    Else
        Result = _
            Message + Delimiter + _
            CStr(Value - StartValue + 1) + "/" + _
            CStr(EndValue - StartValue + 1)
        If PercentVisible Then
            Result = Result + Delimiter + _
            CStr(Format(100 - ((Value - StartValue + 1) / (EndValue - StartValue + 1) * 100), "0.00")) + "%"
    End If
    End If
    ProgressText = Result
End Function

