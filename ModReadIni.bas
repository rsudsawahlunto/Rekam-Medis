Attribute VB_Name = "ModReadIni"
Public Declare Function GetSystemDirectory Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Integer) As Long
Declare Function GetWindowsDirectory Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Integer) As Long
Declare Function OSGetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnString$, ByVal NumBytes As Long, ByVal FileName$) As Long
Declare Function OSWritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$) As Long
Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
'Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpCaption As Any) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function ShowWindow Lib "user32" (ByVal handle As Long, ByVal cmd As Long) As Long
Declare Function Sfocus Lib "user32" Alias "SetFocus" (ByVal handle As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
'Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long

'Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Public Declare Function FlushConsoleInputBuffer Lib "kernel32" (ByVal hConsoleInput As Long) As Long

Function ReadINI(ByVal szApp$, ByVal szItem$, ByVal szDefault$, ByVal szFile$) As String
    Dim tmp As String
    Dim x As Integer

    tmp = String$(2048, 32)
    '  OSGetPrivateProfileString("windows", "Run", "", tmp, Len(tmp), "WIN.INI")
    '     WIN.INI
    '     [windows]   <- szApp$
    '     Run = 1 < -szItem$(Item = value)
    'szDefault$ -> Nilai awal jika item tidak bernilai
    x = OSGetPrivateProfileString(szApp$, szItem$, szDefault, tmp, Len(tmp), szFile$)
    'ReadINI = Mid$(tmp, 1, x)
    tmp = Mid$(tmp, 1, x)
    ReadINI = IIf(tmp = "", szDefault, tmp)
End Function

Public Function GetINI(ByVal Section, ByVal Key, ByVal def, FileIni$) As String
    Dim retVal As String, AppName As String
    Dim worked As Long, l As Long
    retVal = String$(2048, 0)
    l = Len(retVal)
    worked = OSGetPrivateProfileString(Section, Key, def, retVal, l, App.Path & "\" & FileIni)
    GetINI = IIf(worked = 0, def, Left(retVal, InStr(retVal, Chr(0)) - 1))
End Function

Public Sub SetINI(Section$, Key$, value$, FileIni$)
    Dim x As Long
    'If OSWritePrivateProfileString(section, key, Value, App.Path & "\" & gINIFile) Then
    x = OSWritePrivateProfileString(Section, Key, value, FileIni)
    If x <> 1 Then
        MsgBox "File " & FileIni$ & " couldn't be edited !", vbCritical, "Error 'INI' Fileini"
    End If
End Sub

Public Function GetWindowsDir() As String
    Dim Temp$, x As Long
    Temp$ = String$(145, 0)              ' Size Buffer
    x = GetWindowsDirectory(Temp$, 145)  ' Make API Call
    Temp$ = Left$(Temp$, x)              ' Trim Buffer

    If Right$(Temp$, 1) <> "\" Then      ' Add \ if necessary
        GetWindowsDir$ = Temp$ & "\"
    Else
        GetWindowsDir$ = Temp$
    End If
End Function

Public Function GetWindowsSysDir() As String
    Dim Temp$, x As Long
    Temp$ = String$(145, 0)                 ' Size Buffer
    x = GetSystemDirectory(Temp$, 145)      ' Make API Call
    Temp$ = Left$(Temp$, x)                 ' Trim Buffer

    If Right$(Temp$, 1) <> "\" Then         ' Add \ if necessary
        GetWindowsSysDir$ = Temp$ & "\"
    Else
        GetWindowsSysDir$ = Temp$
    End If
End Function

Function GetDefaultPrinter() As String
    Dim cFile$, tmp As String, x  As Long
    cFile$ = GetWindowsDir$() & "win.ini"
    tmp = String$(255, 32)
    x = OSGetPrivateProfileString("windows", "device", "", tmp, Len(tmp), cFile$)
    GetDefaultPrinter = Mid$(tmp, 1, x)
End Function
