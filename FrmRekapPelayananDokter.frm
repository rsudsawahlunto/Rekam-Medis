VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10a.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRekapPelayananDokter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Rekapitulasi Pelayanan Dokter"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9315
   Icon            =   "FrmRekapPelayananDokter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   9315
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   9255
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   4440
         TabIndex        =   7
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   2640
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   360
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
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
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   360
         Left            =   4680
         TabIndex        =   14
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
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
      Begin VB.Label Label3 
         Caption         =   "Ruangan"
         Height          =   255
         Left            =   4680
         TabIndex        =   13
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Instalasi"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   9255
      Begin VB.Frame Frame4 
         Caption         =   "Tanggal Pelayanan"
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
         Left            =   4200
         TabIndex        =   1
         Top             =   120
         Width           =   4935
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   120
            TabIndex        =   2
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
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   66715651
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   2640
            TabIndex        =   3
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
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   66715651
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   2280
            TabIndex        =   4
            Top             =   315
            Width           =   375
         End
      End
      Begin MSDataListLib.DataCombo dcDokter 
         Height          =   360
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Nama Dokter"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3735
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "0"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7320
      Picture         =   "FrmRekapPelayananDokter.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "FrmRekapPelayananDokter.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "FrmRekapPelayananDokter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCetak_Click()
    If Periksa("datacombo", dcDokter, "Nama Dokter kosong") = False Then Exit Sub
    If Periksa("datacombo", dcInstalasi, "Nama Instalasi kosong") = False Then Exit Sub
    If Periksa("datacombo", dcRuangan, "Nama Ruangan kosong") = False Then Exit Sub
    strNamaDokter = FrmRekapPelayananDokter.dcDokter.Text
    FrmCetakLapRekapPelayananDokter.Show
    
    End Sub

Private Sub cmdTutup_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subLoadDcSource
    dtpAwal.Value = Format(Now, "dd MMMM yyyy 00:00:00")
    dtpAkhir.Value = Format(Now, "dd MMMM yyyy 23:59:59")

End Sub
Private Sub subLoadDcSource()
On Error GoTo errLoad

    Call msubDcSource(dcDokter, rs, "SELECT IdPegawai, NamaLengkap FROM DataPegawai where KdJenisPegawai='001' ORDER BY NamaLengkap ")
    If rs.EOF = False Then dcDokter.BoundText = rs(0).Value

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcInstalasi_GotFocus()
On Error GoTo errLoad
'Dim tempKode As String
'
'    tempKode = dcInstalasi.BoundText
'
'        strSQL = "SELECT DISTINCT KdInstalasi,NamaInstalasi FROM V_KelasPelayanan order by NamaInstalasi"
'    Call msubDcSource(dcInstalasi, rs, strSQL)
'    dcInstalasi.BoundText = tempKode
'
    Call msubDcSource(dcInstalasi, rs, "SELECT KdInstalasi, NamaInstalasi FROM Instalasi ORDER BY NamaInstalasi")
    dcInstalasi.Text = rs(1).Value
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcRuangan_GotFocus()
On Error GoTo errLoad
'Dim tempKode As String
'
'    tempKode = dcRuangan.BoundText
'        strSQL = "SELECT distinct KdRuangan, NamaRuangan FROM V_KelasPelayanan order by NamaRuangan"
'    Call msubDcSource(dcRuangan, rs, strSQL)
'    dcRuangan.BoundText = tempKode

    Call msubDcSource(dcRuangan, rs, "select kdRuangan, NamaRuangan from Ruangan where kdInstalasi in (select kdInstalasi from Instalasi where kdInstalasi like '%" & dcInstalasi.BoundText & "%') order by NamaRuangan")
    dcRuangan.Text = rs(1).Value

Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Private Sub LoadDataCombo(NCek As String, Optional BoolCek As Boolean)
'On Error GoTo errLoad
'
'    Call msubDcSource(dcInstalasi, rs, "SELECT KdInstalasi, NamaInstalasi FROM Instalasi ORDER BY NamaInstalasi")
'    dcInstalasi.Text = rs(1).Value
'
'    Call msubDcSource(dcRuangan, rs, "select kdRuangan, NamaRuangan from Ruangan where kdInstalasi in (select kdInstalasi from Instalasi where kdInstalasi like '%" & dcInstalasi.BoundText & "%') order by NamaRuangan")
'    dcRuangan.Text = rs(1).Value
'
'
'
'
'End Sub
