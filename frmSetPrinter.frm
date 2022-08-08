VERSION 5.00
Begin VB.Form frmSetPrinter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Setting Server Printer Barcode"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetPrinter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5175
   Begin VB.CommandButton cmdLanjut 
      Caption         =   "&Lanjut"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoadReg 
      Caption         =   "&Load dari Registry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   5175
      Begin VB.ComboBox cbPrinterBarcode 
         Height          =   330
         Left            =   240
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   600
         Width           =   4695
      End
      Begin VB.CommandButton cmdSaveReg 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   1800
         TabIndex        =   0
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   3480
         TabIndex        =   1
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Server Printer Barcode"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         TabIndex        =   5
         Top             =   315
         Width           =   2355
      End
   End
   Begin VB.Image Image1 
      Height          =   1905
      Left            =   0
      Picture         =   "frmSetPrinter.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5205
   End
End
Attribute VB_Name = "frmSetPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBatal_Click()
    Unload Me
End Sub

Private Sub cmdLoadReg_Click()
'    mstrServerPrinterBarcode = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "ServerPrinterBarcode")
    mstrServerPrinterBarcode = ReadINI("Default Printer", "PrinterBarcode", "", "C:\SettingPrinter.ini")
    cbPrinterBarcode.Text = mstrServerPrinterBarcode
End Sub

Private Sub cmdSaveReg_Click()
    On Error GoTo errorLanjut
    mstrServerPrinterBarcode = cbPrinterBarcode.Text
'    Call CreateKey("HKEY_CURRENT_USER\Software\Medifirst2000")
'    Call CreateKey("HKEY_CURRENT_USER\Software\Medifirst2000\Standard")
'    Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "ServerPrinterBarcode", mstrServerPrinterBarcode)
    
    WriteIniValue "C:\SettingPrinter.ini", "Default Printer", "PrinterBarcode", mstrServerPrinterBarcode
    Unload Me
    Exit Sub
errorLanjut:
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    For Each prn In Printers
        cbPrinterBarcode.AddItem prn.DeviceName
    Next
    Call cmdLoadReg_Click
End Sub

Private Sub txtServerName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSaveReg.SetFocus
End Sub
