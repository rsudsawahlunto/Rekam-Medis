Attribute VB_Name = "modAPI"
Option Explicit

'modul gantiwarna
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'modul gantiwarna

'modul form transparant
Public Const LWA_COLORKEY = 1
Public Const LWA_ALPHA = 2
Public Const LWA_BOTH = 3
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = -20

Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal color As Long, ByVal x As Byte, ByVal alpha As Long) As Boolean
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'modul form transparant

'modul Flash
Public Declare Function ShellExecute& Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Const SW_NORMAL = 1
Dim sd As String * 256
Dim WinDir, FileName As String
'modul Flash

'modul gantiwarna
Public Sub gantiwarna(ctl As Control, ByVal baru As Long, ByVal lama As Long, ByVal x As Single, ByVal y As Single)
    With ctl
        If (.Width < x) Or (x < 0) Or (.Height < y) Or (y < 0) Then
            ReleaseCapture
            .BackColor = lama
        Else
            SetCapture .hWnd
            .BackColor = baru
        End If
    End With
    On Error GoTo 0
End Sub

'modul gantiwarna

'modul form transparant
Sub SetTranslucent(ThehWnd As Long, color As Long, nTrans As Integer, flag As Byte)
    On Error GoTo errRtn
    Dim attrib As Long
    attrib = GetWindowLong(ThehWnd, GWL_EXSTYLE)
    SetWindowLong ThehWnd, GWL_EXSTYLE, attrib Or WS_EX_LAYERED
    SetLayeredWindowAttributes ThehWnd, color, nTrans, flag
    Exit Sub
errRtn:
    MsgBox Err.Description & " Source :" & Err.Description
End Sub

'modul form transparant

'modul Flash
Public Sub OpenWebsite(strWebsite As String)
    If ShellExecute(&O0, "Open", strWebsite, vbNullString, vbNullString, SW_NORMAL) < 33 Then
    End If
End Sub

Public Function PlayFlashMovie(frm As Form)
    Dim x As Long
    'get the windows system directory
    'X=lenght of the windows directory's string
    x& = GetSystemDirectory(sd, Len(sd))
    WinDir = Left(sd, x)

    FileName = WinDir & "\Medifirst3D.swf"
    If Not FileExists(FileName) Then
        frm.Flash1.Visible = False
        Exit Function
    End If
    With frm.Flash1
        .Movie = FileName
        .Play
    End With
    Exit Function
End Function

Function FileExists(Path As String) As Boolean
    Dim Temp As String
    If Path = "" Then Exit Function
    'try to Dir the current path. if the file exists, Dir returns
    'its name, else it returns a null string
    Temp = Dir(Path)
    If Temp <> "" Then FileExists = True Else FileExists = False
End Function

'modul Flash
