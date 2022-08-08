VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRekapitulasiHarianPasienRI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pelayanan Rawat Inap"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   Icon            =   "frmRekapHarianPasienRI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   9270
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   9255
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
         Height          =   615
         Left            =   6720
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Spreadsheet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   9255
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
         TabIndex        =   5
         Top             =   120
         Width           =   2415
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   120
            TabIndex        =   6
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
            CustomFormat    =   "MMMM yyyy"
            Format          =   61669379
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   5400
            TabIndex        =   7
            Top             =   240
            Visible         =   0   'False
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
            CustomFormat    =   "MMMM yyyy"
            Format          =   61669379
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   5040
            TabIndex        =   8
            Top             =   315
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   360
         Left            =   2760
         TabIndex        =   9
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo dcSubInstalasi 
         Height          =   360
         Left            =   5880
         TabIndex        =   11
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin VB.Label Label1 
         Caption         =   "Sub Instalasi"
         Height          =   255
         Left            =   5880
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Ruang Rawat"
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7320
      Picture         =   "frmRekapHarianPasienRI.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRekapHarianPasienRI.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmRekapitulasiHarianPasienRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''splakuk 2009-06-10
Option Explicit

'Private Sub chk1_Click()
'    If chk1.Value = vbChecked Then
'        Label8.Visible = True
'        dtpAkhir.Visible = True
'    Else
'        Label8.Visible = False
'        dtpAkhir.Visible = False
'    End If
'End Sub

Private Sub cmdCetak_Click()
    If Periksa("datacombo", dcRuangan, "Nama Ruangan kosong") = False Then Exit Sub
    If Periksa("datacombo", dcSubInstalasi, "Sub Instalasi / Kasus Penyakit kosong") = False Then Exit Sub
    strNNamaSubInstalasi = dcSubInstalasi.Text
    strNNamaRuangan = dcRuangan.Text
    frmCetakUreqRekapHarianPasienRI.Show
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcRuangan_Change()
    Call msubDcSource(dcSubInstalasi, rs, "select kdsubinstalasi, NamaSubInstalasi from V_SubInstalasiRuangan where kdruangan = '" & dcRuangan.BoundText & "' order by NamaSubInstalasi")
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subLoadDcSource
    dtpAwal.Value = Format(Now, "dd MMMM yyyy 00:00:00")
    dtpAkhir.Value = Format(Now, "dd MMMM yyyy 23:59:59")
End Sub

Private Sub subLoadDcSource()
    Call msubDcSource(dcRuangan, rs, "select kdRuangan, NamaRuangan from Ruangan where kdInstalasi = '03' order by NamaRuangan")
    'Call msubDcSource(dcRuangan, rs, "select kdsubinstalasi, NamaSubInstalasi from V_SubInstalasiRuangan where kdruangan = '" & dcRuangan.BoundText & "' order by NamaRuangan")
End Sub

