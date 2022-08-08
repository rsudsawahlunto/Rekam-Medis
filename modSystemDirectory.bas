Attribute VB_Name = "modSystemDirectory"
Option Explicit

Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" _
(ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" _
(ByVal pidl As Long, ByVal pszPath As String) As Long

Public Enum enmSpecialFolderName
    CSIDL_DESKTOP = &H0 '// The Desktop - virtual folder
    CSIDL_PROGRAMS = 2 '// Program Files
    CSIDL_CONTROLS = 3 '// Control Panel - virtual folder
    CSIDL_PRINTERS = 4 '// Printers - virtual folder
    CSIDL_DOCUMENTS = 5 '// My Documents
    CSIDL_FAVORITES = 6 '// Favourites
    CSIDL_STARTUP = 7 '// Startup Folder
    CSIDL_RECENT = 8 '// Recent Documents
    CSIDL_SENDTO = 9 '// Send To Folder
    CSIDL_BITBUCKET = 10 '// Recycle Bin - virtual folder
    CSIDL_STARTMENU = 11 '// Start Menu
    CSIDL_DESKTOPFOLDER = 16 '// Desktop folder
    CSIDL_DRIVES = 17 '// My Computer - virtual folder
    CSIDL_NETWORK = 18 '// Network Neighbourhood - virtual folder
    CSIDL_NETHOOD = 19 '// NetHood Folder
    CSIDL_FONTS = 20 '// Fonts folder
    CSIDL_SHELLNEW = 21 '// ShellNew folder
End Enum

Public Type SHITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Const MAX_PATH As Integer = 260

Public Function funcGetSpecialFolder(ByVal hwndForm As Long, ByVal SpecialFolderName As enmSpecialFolderName) As String
    Dim sPath As String
    Dim IDL As ITEMIDLIST

    funcGetSpecialFolder = ""
    If SHGetSpecialFolderLocation(hwndForm, SpecialFolderName, IDL) = 0 Then
        sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(IDL.mkid.cb, sPath) Then
            funcGetSpecialFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1) & ""
        End If
    End If
End Function

