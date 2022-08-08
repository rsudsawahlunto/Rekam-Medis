Attribute VB_Name = "modIniFile"
Option Explicit

Public strLokasiFile As String

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpDefault As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function funcGetFromINI(ByVal strSection As String, ByVal strKey As String, _
    ByVal strDefault As String, ByVal strIniFile As String) As String
    Dim strBuffer As String, lngRet As Long

    strBuffer = String$(255, 0)
    lngRet = GetPrivateProfileString(strSection, strKey, "", strBuffer, Len(strBuffer), strIniFile)
    If lngRet = 0 Then
        If strDefault <> "" Then funcAddToINI strSection, strKey, strDefault, strIniFile
        funcGetFromINI = strDefault
    Else
        funcGetFromINI = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    End If
End Function

Public Function funcAddToINI(ByVal strSection As String, ByVal strKey As String, _
    ByVal strValue As String, ByVal strIniFile As String) As Boolean
    Dim lngRet As Long

    lngRet = WritePrivateProfileString(strSection, strKey, strValue, strIniFile)
    funcAddToINI = (lngRet)
End Function
