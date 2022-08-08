VERSION 5.00
Begin VB.Form frmSplashScreen 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   840
      Top             =   1080
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   5955
      Left            =   0
      Picture         =   "frmSplashScreen.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   4230
   End
   Begin VB.Image Image2 
      Height          =   5955
      Left            =   0
      Picture         =   "frmSplashScreen.frx":72B7
      Top             =   0
      Visible         =   0   'False
      Width           =   4230
   End
End
Attribute VB_Name = "frmSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MoveScreen As Boolean, color As Long, flag As Byte
Dim MousX, MousY, CurrX, CurrY As Integer

Private Sub Command3_Click()
    End
End Sub

Private Sub Form_Activate()
    On Error GoTo errRtn
    color = RGB(0, 0, 255): flag = 0
    flag = flag Or LWA_COLORKEY: frmSplashScreen.Show
    SetTranslucent frmSplashScreen.hWnd, color, 0, flag
    Exit Sub
errRtn:
    MsgBox Err.Description & " Source : " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmLogin.Show
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveScreen = True: MousX = x: MousY = y
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MoveScreen Then
        CurrX = Me.Left - MousX + x
        CurrY = Me.Top - MousY + y
        Me.Move CurrX, CurrY
    End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveScreen = False
End Sub

Private Sub Timer1_Timer()
    Dim i As Double
    Image1.Visible = True
    For i = 1 To 1000000
        DoEvents
        If i = 400000 Then Image2.Visible = True
    Next i
    Unload Me
End Sub
