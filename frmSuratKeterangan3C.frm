VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSuratKeterangan3C 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Surat Keterangan"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSuratKeterangan3C.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   8910
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   6360
      TabIndex        =   25
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox txtKesimpulan2 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2595
      Left            =   0
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   5040
      Width           =   8895
   End
   Begin VB.TextBox txtUmur 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   3960
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtSex 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   7680
      TabIndex        =   11
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "JAWABAN PEMERIKSAAN RONTGEN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   0
      TabIndex        =   12
      Top             =   1080
      Width           =   8895
      Begin VB.TextBox txtNoRontgen 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   6720
         TabIndex        =   7
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtKiriman 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   6720
         TabIndex        =   8
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtNoCM 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   6720
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtAlamat 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         TabIndex        =   4
         Top             =   1650
         Width           =   1815
      End
      Begin VB.TextBox txtTindakan 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   9
         Top             =   2160
         Width           =   7215
      End
      Begin VB.TextBox txtPekerjaan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtNama 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   6720
         TabIndex        =   5
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   126484483
         UpDown          =   -1  'True
         CurrentDate     =   40087
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "No Rontgen"
         Height          =   210
         Left            =   5280
         TabIndex        =   23
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Kiriman dari"
         Height          =   210
         Left            =   5280
         TabIndex        =   22
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "No Pendaftaran"
         Height          =   210
         Left            =   5280
         TabIndex        =   21
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal"
         Height          =   210
         Left            =   5280
         TabIndex        =   20
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Umur"
         Height          =   210
         Left            =   3480
         TabIndex        =   19
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nama Tindakan"
         Height          =   210
         Left            =   240
         TabIndex        =   18
         Top             =   2175
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Alamat"
         Height          =   210
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pekerjaan"
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1020
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   14
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Catatan Expertise"
      Height          =   210
      Left            =   120
      TabIndex        =   24
      Top             =   4800
      Width           =   1440
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6960
      Picture         =   "frmSuratKeterangan3C.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmSuratKeterangan3C.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmSuratKeterangan3C.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmSuratKeterangan3C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOut_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    frmCetakSuratKeterangan3.Show
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim kata As String
    On Error GoTo hell

    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)

    mstrNoPen = frmTransaksiPasien.txtNoPendaftaran

    strSQL = "SELECT * FROM JawabanPemeriksaanRontgen WHERE NoPendaftaran = '" & mstrNoPen & "'"

    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    If rs.RecordCount = 0 Then Exit Sub

    txtNoRontgen = rs.Fields(0).value
    txtNoCM.Text = rs.Fields(1).value
    txtNama.Text = rs.Fields(2).value
    txtSex.Text = rs.Fields(3).value
    txtUmur.Text = rs.Fields(4).value
    txtPekerjaan.Text = rs.Fields(5).value
    txtAlamat.Text = rs.Fields(6).value
    dtpAwal.value = rs.Fields(7).value
    txtKiriman.Text = rs.Fields(8).value
    txtTindakan.Text = rs.Fields(9).value
    txtKesimpulan2.Text = rs.Fields(11).value
    txtKesimpulan2.Enabled = False

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub txtKesimpulan2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdPrint.SetFocus
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtSex_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtUmur_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub
