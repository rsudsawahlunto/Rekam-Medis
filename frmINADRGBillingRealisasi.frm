VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmINADRGBillingRealisasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Realisasi INADRG dan Billing System"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   Icon            =   "frmINADRGBillingRealisasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   8055
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   8055
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   8055
      Begin VB.Frame Frame4 
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   5055
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   2640
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   126943235
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "ddMMMMyyyy"
            Format          =   126943235
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   2320
            TabIndex        =   10
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   1680
            TabIndex        =   7
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   360
         Left            =   5280
         TabIndex        =   2
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Instalasi"
         Height          =   255
         Left            =   5280
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   11
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
      Height          =   975
      Left            =   0
      Picture         =   "frmINADRGBillingRealisasi.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6120
      Picture         =   "frmINADRGBillingRealisasi.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmINADRGBillingRealisasi.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmINADRGBillingRealisasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCetak_Click()
    If Periksa("datacombo", dcInstalasi, "Nama Instalasi kosong") = False Then Exit Sub
    strNNamaInstalasi = dcInstalasi.Text

    frmctkINADRGBillingRealisasi.Show
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcInstalasi.MatchedWithList = True Then cmdCetak.SetFocus
        strSQL = "select kdInstalasi, NamaInstalasi from Instalasi where kdInstalasi in ('02', '03') and StatusEnabled='1' and (Namainstalasi LIKE '%" & dcInstalasi.Text & "%') order by NamaInstalasi"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcInstalasi.Text = ""
            Exit Sub
        End If
        dcInstalasi.BoundText = rs(0).value
        dcInstalasi.Text = rs(1).value
    End If
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub dtpAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subLoadDcSource
    dtpAwal.value = Format(Now, "dd MMMM yyyy 00:00:00")
    dtpAkhir.value = Format(Now, "dd MMMM yyyy 23:59:59")
End Sub

Private Sub subLoadDcSource()
    Call msubDcSource(dcInstalasi, rs, "select kdInstalasi, NamaInstalasi from Instalasi where kdInstalasi in ('02', '03') and StatusEnabled='1' order by NamaInstalasi")
    dcInstalasi.BoundText = "02"
End Sub
