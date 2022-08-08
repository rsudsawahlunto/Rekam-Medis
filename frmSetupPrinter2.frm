VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSetupPrinter2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Setting Printer"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetupPrinter2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   12585
   Begin VB.Frame Frame6 
      Caption         =   "Setting Printer Label"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6360
      TabIndex        =   5
      Top             =   7800
      Width           =   6015
      Begin VB.ComboBox cboPrinter2 
         Height          =   330
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   680
         Width           =   4335
      End
      Begin VB.ComboBox cboPrinter1 
         Height          =   330
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   300
         Width           =   4335
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Printer 2"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Printer 1"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Timer tmrPrinter 
      Left            =   2760
      Top             =   600
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   8040
      Width           =   6015
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Simpan"
         Height          =   450
         Left            =   2940
         TabIndex        =   4
         Top             =   200
         Width           =   1455
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   450
         Left            =   4440
         TabIndex        =   3
         Top             =   200
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Setting Printer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7950
      Left            =   0
      TabIndex        =   1
      Top             =   1065
      Width           =   12495
      Begin VB.Frame Frame13 
         Caption         =   "Legal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   6360
         TabIndex        =   35
         Top             =   3840
         Width           =   6015
         Begin MSFlexGridLib.MSFlexGrid fgLegal 
            Height          =   1455
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   2566
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "A4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   6360
         TabIndex        =   32
         Top             =   2040
         Width           =   6015
         Begin MSFlexGridLib.MSFlexGrid fgA4 
            Height          =   1455
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   2566
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
         End
      End
      Begin VB.Frame Frame11 
         Height          =   975
         Left            =   120
         TabIndex        =   27
         Top             =   5640
         Width           =   6015
         Begin VB.ComboBox cboDuplexing 
            Height          =   330
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   600
            Width           =   4575
         End
         Begin VB.ComboBox cboJnsDriver 
            Height          =   330
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "Duplexing"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   795
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "Jenis Driver"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Seperempat A4 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   25
         Top             =   3840
         Width           =   6015
         Begin MSFlexGridLib.MSFlexGrid fgSeperempatA42 
            Height          =   1455
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   2566
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Seperempat A4 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   6015
         Begin MSFlexGridLib.MSFlexGrid fgSeperempatA41 
            Height          =   1455
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   2566
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   6360
         TabIndex        =   17
         Top             =   5760
         Width           =   3525
         Begin VB.ComboBox cboUkuranKertas 
            Height          =   330
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   60
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.Frame Frame3 
            Caption         =   "Orientasi Kertas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   3255
            Begin VB.OptionButton OptOrien 
               Caption         =   "Portrait"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   240
               TabIndex        =   20
               Top             =   240
               Width           =   945
            End
            Begin VB.OptionButton OptOrien 
               Caption         =   "Landscape"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   1440
               TabIndex        =   19
               Top             =   240
               Width           =   1170
            End
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "Ukuran Kertas"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   22
            Top             =   345
            Visible         =   0   'False
            Width           =   1140
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Interval Print"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9960
         TabIndex        =   14
         Top             =   5640
         Width           =   2415
         Begin VB.TextBox txtIntervalPrint 
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Detik"
            Height          =   375
            Left            =   1080
            TabIndex        =   16
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Seperempat A4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   6360
         TabIndex        =   12
         Top             =   240
         Width           =   6015
         Begin MSFlexGridLib.MSFlexGrid fgSeperempatA4 
            Height          =   1455
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   2566
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Printer Kartu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6015
         Begin MSFlexGridLib.MSFlexGrid fgPrinterKartu 
            Height          =   1455
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   2566
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
         End
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7920
      TabIndex        =   0
      Top             =   5640
      Width           =   855
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Image Image3 
      Height          =   945
      Left            =   10800
      Picture         =   "frmSetupPrinter2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmSetupPrinter2.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7680
      Picture         =   "frmSetupPrinter2.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmSetupPrinter2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Const strChecked = "þ"
Const strUnChecked = "q"
Public icol, iCols As Long
Public irow, irows As Long
Public a, b, c As Integer
Public j, k, l As Integer
Public sPrinterLegal As String
Public indexGrid As Integer
Private Sub cboDuplexing_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then OptOrien(0).SetFocus
End Sub

Private Sub cboJnsDriver_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboDuplexing.SetFocus
End Sub

Private Sub cmdBatal_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

sPrinterLabel1 = cboPrinter1.Text
sPrinterLabel2 = cboPrinter2.Text
sDriver = cboJnsDriver.Text

Call CreateKey("HKEY_CURRENT_USER\Software\Medifirst2000")
Call CreateKey("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000")

'Untuk Printer Kartu
With fgPrinterKartu
sPrinter = ""
    For i = 1 To .Row
        If .TextMatrix(i, 1) = strChecked Then
            If sPrinter = "" Then
                sPrinter = .TextMatrix(i, 0) & ";"
            Else
                sPrinter = sPrinter & "" & .TextMatrix(i, 0) & ";"
            End If
            Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "Printer1", sPrinter)
        End If
    Next i
End With

'Untuk Printer SeperempatA41
With fgSeperempatA41
sPrinter2 = ""
    For i = 1 To .Row
        If .TextMatrix(i, 1) = strChecked Then
            If sPrinter2 = "" Then
                sPrinter2 = .TextMatrix(i, 0) & ";"
            Else
                sPrinter2 = sPrinter2 & "" & .TextMatrix(i, 0) & ";"
            End If
            Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "Printer2", sPrinter2)
        End If
    Next i
End With

'Untuk Printer SeperempatA42
With fgSeperempatA42
sPrinter3 = ""
    For i = 1 To .Row
        If .TextMatrix(i, 1) = strChecked Then
            If sPrinter3 = "" Then
                sPrinter3 = .TextMatrix(i, 0) & ";"
            Else
                sPrinter3 = sPrinter3 & "" & .TextMatrix(i, 0) & ";"
            End If
            Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "Printer3", sPrinter3)
        End If
    Next i
End With

'Untuk Printer SeperempatA4
With fgSeperempatA4
sPrinter4 = ""
    For i = 1 To .Row
        If .TextMatrix(i, 1) = strChecked Then
            If sPrinter4 = "" Then
                sPrinter4 = .TextMatrix(i, 0) & ";"
            Else
                sPrinter4 = sPrinter4 & "" & .TextMatrix(i, 0) & ";"
            End If
            Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "Printer4", sPrinter4)
        End If
    Next i
End With

'Untuk Printer A4
With fgA4
sPrinter5 = ""
    For i = 1 To .Row
        If .TextMatrix(i, 1) = strChecked Then
            If sPrinter5 = "" Then
                sPrinter5 = .TextMatrix(i, 0) & ";"
            Else
                sPrinter5 = sPrinter5 & "" & .TextMatrix(i, 0) & ";"
            End If
            Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "Printer5", sPrinter5)
        End If
    Next i
End With

'Untuk Printer Legal
With fgLegal
sPrinterLegal = ""
    For i = 1 To .Row
        If .TextMatrix(i, 1) = strChecked Then
            If sPrinterLegal = "" Then
                sPrinterLegal = .TextMatrix(i, 0) & ";"
            Else
                sPrinterLegal = sPrinterLegal & "" & .TextMatrix(i, 0) & ";"
            End If
            Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "PrinterLegal", sPrinterLegal)
        End If
    Next i
End With

' Untuk Printer Label
Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "PrinterLabel1", sPrinterLabel1)
Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "PrinterLabel2 ", sPrinterLabel2)

Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "JenisDriver", sDriver)
Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "Duplexing", cboDuplexing.Text)

If txtIntervalPrint.Text = "" Then
    intTimerPrinter = 1
Else
    intTimerPrinter = txtIntervalPrint.Text
End If

MsgBox "Setting Printer Selesai", vbInformation
cmdOK.Enabled = False
cmdBatal.SetFocus

End Sub

Private Sub fgA4_Click()
    With fgA4
        If .TextMatrix(.Row, 1) = strUnChecked Then
            .TextMatrix(.Row, 1) = strChecked
        Else
            .TextMatrix(.Row, 1) = strUnChecked
        End If
    End With
End Sub
Private Sub fgLegal_Click()
    With fgLegal
        If .TextMatrix(.Row, 1) = strUnChecked Then
            .TextMatrix(.Row, 1) = strChecked
        Else
            .TextMatrix(.Row, 1) = strUnChecked
        End If
    End With
End Sub

Private Sub fgPrinterKartu_Click()
    With fgPrinterKartu
        If .TextMatrix(.Row, 1) = strUnChecked Then
            .TextMatrix(.Row, 1) = strChecked
        Else
            .TextMatrix(.Row, 1) = strUnChecked
        End If
    End With
End Sub

Private Sub fgSeperempatA4_Click()
    With fgSeperempatA4
        If .TextMatrix(.Row, 1) = strUnChecked Then
            .TextMatrix(.Row, 1) = strChecked
        Else
            .TextMatrix(.Row, 1) = strUnChecked
        End If
    End With

End Sub

Private Sub fgSeperempatA41_Click()
    With fgSeperempatA41
        If .TextMatrix(.Row, 1) = strUnChecked Then
            .TextMatrix(.Row, 1) = strChecked
        Else
            .TextMatrix(.Row, 1) = strUnChecked
        End If
    End With

End Sub

Private Sub fgSeperempatA42_Click()
    With fgSeperempatA42
        If .TextMatrix(.Row, 1) = strUnChecked Then
            .TextMatrix(.Row, 1) = strChecked
        Else
            .TextMatrix(.Row, 1) = strUnChecked
        End If
    End With

End Sub

Private Sub Form_Load()
On Error GoTo pesan

Call PlayFlashMovie(Me)
Call centerForm(Me, MDIUtama)
indexGrid = 1
For Each prn In Printers
' Untuk Load Printer
    Call SetPrinterKartu
    Call SetSeperempatA41
    Call SetSeperempatA42
    Call SetSeperempatA4
    Call SetA4
    Call SetLegal
    
'Untuk Load Checklis
    Call CheckPrinterKartu
    Call CheckSeperempatA41
    Call CheckSeperempatA42
    Call CheckSeperempatA4
    Call CheckA4
    Call CheckLegal
    
    cboPrinter1.AddItem prn.DeviceName
    cboPrinter2.AddItem prn.DeviceName
    cboJnsDriver.AddItem prn.DriverName
    indexGrid = indexGrid + 1
Next

Addcbo cboUkuranKertas, "Default", crDefaultPaperSize
    Addcbo cboUkuranKertas, "Letter", crPaperLetter
    Addcbo cboUkuranKertas, "Small Letter", crPaperLetterSmall
    Addcbo cboUkuranKertas, "Legal", crPaperLegal
    Addcbo cboUkuranKertas, "10x14", crPaper10x14
    Addcbo cboUkuranKertas, "11x17", crPaper11x17
    Addcbo cboUkuranKertas, "A3", crPaperA3
    Addcbo cboUkuranKertas, "A4", crPaperA4
    Addcbo cboUkuranKertas, "A4 Small", crPaperA4Small
    Addcbo cboUkuranKertas, "A5", crPaperA5
    Addcbo cboUkuranKertas, "B4", crPaperB4
    Addcbo cboUkuranKertas, "B5", crPaperB5
    Addcbo cboUkuranKertas, "C Sheet", crPaperCsheet
    Addcbo cboUkuranKertas, "D Sheet", crPaperDsheet
    Addcbo cboUkuranKertas, "Envelope 9", crPaperEnvelope9
    Addcbo cboUkuranKertas, "Envelope 10", crPaperEnvelope10
    Addcbo cboUkuranKertas, "Envelope 11", crPaperEnvelope11
    Addcbo cboUkuranKertas, "Envelope 12", crPaperEnvelope12
    Addcbo cboUkuranKertas, "Envelope 14", crPaperEnvelope14
    Addcbo cboUkuranKertas, "Envelope B4", crPaperEnvelopeB4
    Addcbo cboUkuranKertas, "Envelope B5", crPaperEnvelopeB5
    Addcbo cboUkuranKertas, "Envelope B6", crPaperEnvelopeB6
    Addcbo cboUkuranKertas, "Envelope C3", crPaperEnvelopeC3
    Addcbo cboUkuranKertas, "Envelope C4", crPaperEnvelopeC4
    Addcbo cboUkuranKertas, "Envelope C5", crPaperEnvelopeC5
    Addcbo cboUkuranKertas, "Envelope C6", crPaperEnvelopeC6
    Addcbo cboUkuranKertas, "Envelope C65", crPaperEnvelopeC65
    Addcbo cboUkuranKertas, "Envelope DL", crPaperEnvelopeDL
    Addcbo cboUkuranKertas, "Envelope Italy", crPaperEnvelopeItaly
    Addcbo cboUkuranKertas, "Envelope Monarch", crPaperEnvelopeMonarch
    Addcbo cboUkuranKertas, "Envelope Personal", crPaperEnvelopePersonal
    Addcbo cboUkuranKertas, "E Sheet", crPaperEsheet
    Addcbo cboUkuranKertas, "Executive", crPaperExecutive
    Addcbo cboUkuranKertas, "Fanfold Legal German", crPaperFanfoldLegalGerman
    Addcbo cboUkuranKertas, "Fanfold Standard German", crPaperFanfoldStdGerman
    Addcbo cboUkuranKertas, "Fanfold US", crPaperFanfoldUS
    Addcbo cboUkuranKertas, "Folio", crPaperFolio
    Addcbo cboUkuranKertas, "Ledger", crPaperLedger
    Addcbo cboUkuranKertas, "Note", crPaperNote
    Addcbo cboUkuranKertas, "Quarto", crPaperQuarto
    Addcbo cboUkuranKertas, "Statement", crPaperStatement
    Addcbo cboUkuranKertas, "Tabloid", crPaperTabloid
   

    'dupexing
    Addcbo cboDuplexing, "Default", crPRDPDefault
    Addcbo cboDuplexing, "Simplex", crPRDPSimplex
    Addcbo cboDuplexing, "Horizontal", crPRDPHorizontal
    Addcbo cboDuplexing, "Vertical", crPRDPVertical
Exit Sub
pesan:
Call msubPesanError
End Sub

Private Sub Addcbo(cbo As ComboBox, Name As String, index As Integer)
    cbo.AddItem Name                        ' Add the name of the item to the combo box
    cbo.ItemData(cbo.NewIndex) = index      ' Set the .itemdata(.listindex) for later retrieval
End Sub

Private Sub OptOrien_KeyPress(index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then cmdOK.SetFocus
End Sub

Private Sub SetPrinterKartu()

    With fgPrinterKartu
        .TextMatrix(0, 0) = "Nama Printer"
        .TextMatrix(0, 1) = ""
        
        .ColWidth(0) = 4800
        .ColWidth(1) = 400
         .AddItem prn.DeviceName
        
    End With

End Sub

Private Sub SetSeperempatA41()

    With fgSeperempatA41
        .TextMatrix(0, 0) = "Nama Printer"
        .TextMatrix(0, 1) = ""
        
        .ColWidth(0) = 4800
        .ColWidth(1) = 400
        
    
        .AddItem prn.DeviceName
        
    End With

End Sub
Private Sub SetSeperempatA42()

    With fgSeperempatA42
        .TextMatrix(0, 0) = "Nama Printer"
        .TextMatrix(0, 1) = ""
        
        .ColWidth(0) = 4800
        .ColWidth(1) = 400
        
    
        .AddItem prn.DeviceName
        
    End With

End Sub
Private Sub SetSeperempatA4()

    With fgSeperempatA4
        .TextMatrix(0, 0) = "Nama Printer"
        .TextMatrix(0, 1) = ""
        
        .ColWidth(0) = 4800
        .ColWidth(1) = 400
        
    
        .AddItem prn.DeviceName
        
    End With

End Sub
Private Sub SetA4()

    With fgA4
        .TextMatrix(0, 0) = "Nama Printer"
        .TextMatrix(0, 1) = ""
        
        .ColWidth(0) = 4800
        .ColWidth(1) = 400
        
    
        .AddItem prn.DeviceName
        
    End With

End Sub
Private Sub SetLegal()

    With fgLegal
        .TextMatrix(0, 0) = "Nama Printer"
        .TextMatrix(0, 1) = ""
        
        .ColWidth(0) = 4800
        .ColWidth(1) = 400
        
    
        .AddItem prn.DeviceName
        
    End With

End Sub

Private Sub CheckPrinterKartu()

With fgPrinterKartu

irows = .Rows - 1
iCols = .Cols

    For irow = 1 To irows
        j = irow
        For icol = 1 To iCols - 1
            .Row = j
            .Col = icol
            .CellFontName = "Wingdings"
            .CellFontSize = 14
            .CellAlignment = flexAlignCenterCenter
    
            .TextMatrix(j, icol) = strChecked
            Dim position As Integer
            Dim perinterAktif  As String
            perinterAktif = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "Printer1")
            position = InStr(perinterAktif, .TextMatrix(j, 0))
            If position > 0 Then
              .TextMatrix(j, icol) = strChecked
            Else
              .TextMatrix(j, icol) = strUnChecked
            End If
        Next icol
    Next irow
End With

End Sub
Private Sub CheckSeperempatA41()

With fgSeperempatA41

irows = .Rows - 1
iCols = .Cols

    For irow = 1 To irows
        j = irow
        For icol = 1 To iCols - 1
            .Row = j
            .Col = icol
            .CellFontName = "Wingdings"
            .CellFontSize = 14
            .CellAlignment = flexAlignCenterCenter
    
            .TextMatrix(j, icol) = strUnChecked
             Dim position As Integer
            Dim perinterAktif  As String
            perinterAktif = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "Printer2")
            position = InStr(perinterAktif, .TextMatrix(j, 0))
            If position > 0 Then
              .TextMatrix(j, icol) = strChecked
            Else
              .TextMatrix(j, icol) = strUnChecked
            End If
        Next icol
    Next irow
End With

End Sub
Private Sub CheckSeperempatA42()

With fgSeperempatA42

irows = .Rows - 1
iCols = .Cols

    For irow = 1 To irows
        j = irow
        For icol = 1 To iCols - 1
            .Row = j
            .Col = icol
            .CellFontName = "Wingdings"
            .CellFontSize = 14
            .CellAlignment = flexAlignCenterCenter
    
            .TextMatrix(j, icol) = strUnChecked
             Dim position As Integer
            Dim perinterAktif  As String
            perinterAktif = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "Printer3")
            position = InStr(perinterAktif, .TextMatrix(j, 0))
            If position > 0 Then
              .TextMatrix(j, icol) = strChecked
            Else
              .TextMatrix(j, icol) = strUnChecked
            End If
        Next icol
    Next irow
End With

End Sub
Private Sub CheckSeperempatA4()

With fgSeperempatA4

irows = .Rows - 1
iCols = .Cols

    For irow = 1 To irows
        j = irow
        For icol = 1 To iCols - 1
            .Row = j
            .Col = icol
            .CellFontName = "Wingdings"
            .CellFontSize = 14
            .CellAlignment = flexAlignCenterCenter
    
            .TextMatrix(j, icol) = strUnChecked
            
            Dim position As Integer
            Dim perinterAktif  As String
            perinterAktif = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "Printer4")
            position = InStr(perinterAktif, .TextMatrix(j, 0))
            If position > 0 Then
              .TextMatrix(j, icol) = strChecked
            Else
              .TextMatrix(j, icol) = strUnChecked
            End If
            
        Next icol
    Next irow
End With

End Sub
Private Sub CheckA4()

With fgA4

irows = .Rows - 1
iCols = .Cols

    For irow = 1 To irows
        j = irow
        For icol = 1 To iCols - 1
            .Row = j
            .Col = icol
            .CellFontName = "Wingdings"
            .CellFontSize = 14
            .CellAlignment = flexAlignCenterCenter
    
            .TextMatrix(j, icol) = strUnChecked
            Dim position As Integer
            Dim perinterAktif  As String
            perinterAktif = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "Printer5")
            position = InStr(perinterAktif, .TextMatrix(j, 0))
            If position > 0 Then
              .TextMatrix(j, icol) = strChecked
            Else
              .TextMatrix(j, icol) = strUnChecked
            End If
        Next icol
    Next irow
End With

End Sub

Private Sub CheckLegal()

With fgLegal

irows = .Rows - 1
iCols = .Cols

    For irow = 1 To irows
        j = irow
        For icol = 1 To iCols - 1
            .Row = j
            .Col = icol
            .CellFontName = "Wingdings"
            .CellFontSize = 14
            .CellAlignment = flexAlignCenterCenter
    
            .TextMatrix(j, icol) = strUnChecked
        Next icol
    Next irow
End With

End Sub

