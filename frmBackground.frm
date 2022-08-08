VERSION 5.00
Begin VB.Form frmBackground 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Background Medifirst2000"
   ClientHeight    =   9810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19950
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmBackground.frx":0000
   ScaleHeight     =   9810
   ScaleWidth      =   19950
   ShowInTaskbar   =   0   'False
   Begin VB.Image imgLogoRS 
      Height          =   1815
      Left            =   12960
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Label lblNamaRS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   6240
      Width           =   105
   End
End
Attribute VB_Name = "frmBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Exit Sub
    PopupMenu MDIUtama.mnuTransaksi
End Sub
