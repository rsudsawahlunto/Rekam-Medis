Attribute VB_Name = "WheelHook"
Public MyProperty As Object
Public Const MK_CONTROL = &H8
Public Const MK_LBUTTON = &H1
Public Const MK_RBUTTON = &H2
Public Const MK_MBUTTON = &H10
Public Const MK_SHIFT = &H4
Public Const GWL_WNDPROC = -4
Public Const WM_MOUSEWHEEL = &H20A
Public LocalHwnd As Long
Public LocalPrevWndProc As Long
Public MyControl As Object
Public Declare Function CallWindowProc Lib "user32.dll" Alias _
"CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
ByVal msg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias _
"SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong _
As Long) As Long

Public Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal _
    wParam As Long, ByVal lParam As Long) As Long
    Dim MouseKeys As Long
    Dim Rotation As Long
    Dim Xpos As Long
    Dim Ypos As Long
    If Lmsg = WM_MOUSEWHEEL Then
        MouseKeys = wParam And 65535
        Rotation = wParam / 65536
        Xpos = lParam And 65535
        Ypos = lParam / 65536
        If Rotation = -120 Then
            MyProperty.Scroll 0, 3
        Else
            MyProperty.Scroll 0, -3
        End If
    End If
    WindowProc = CallWindowProc(LocalPrevWndProc, Lwnd, Lmsg, wParam, lParam)
End Function

Public Sub WheelHook(PassedControl As Object)
    On Error Resume Next
    Set MyControl = PassedControl
    LocalHwnd = PassedControl.hWnd
    LocalPrevWndProc = SetWindowLong(LocalHwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Sub WheelUnHook()
    Dim WorkFlag As Long
    On Error Resume Next
    WorkFlag = SetWindowLong(LocalHwnd, GWL_WNDPROC, LocalPrevWndProc)
    Set MyControl = Nothing
End Sub

