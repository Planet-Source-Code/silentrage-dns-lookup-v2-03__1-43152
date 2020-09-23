Attribute VB_Name = "modINI"
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Dim Buffer As String * 4096
Public Function GetKeys(Section As String)
    Length = GetPrivateProfileSection(Section, Buffer, 4096, App.Path & "\strings.ini")
    If Length Then
        Entries = Left(Buffer, Length)
        Pos = InStr(Entries, "=")
        While Pos > 0
            R = Mid(Entries, Pos, InStr(Entries, vbNullChar) - Pos + 1)
            Entries = Replace(Entries, R, "|")
            Pos = InStr(Entries, "=")
        Wend
        GetKeys = Split(Left(Entries, Len(Entries) - 1), "|")
    End If
End Function
Public Function GetInt(Section As String, Value As String) As Long
    GetInt = GetPrivateProfileInt(Section, Value, 0, App.Path & "\strings.ini")
End Function
Public Function GetStr(Section As String, Value As String) As String
    Length = GetPrivateProfileString(Section, Value, "", Buffer, 255, App.Path & "\strings.ini")
    If Length Then
        GetStr = Left(Buffer, Length)
    Else
        GetStr = Value
    End If
End Function
