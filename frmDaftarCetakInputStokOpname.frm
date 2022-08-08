VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmDaftarCetakInputStokOpname 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Cetak Lembar Input Stok Opname"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarCetakInputStokOpname.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   8940
   Begin VB.CheckBox chkTampilHarga 
      Caption         =   "Tampilkan Harga"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   7050
      TabIndex        =   4
      Top             =   2280
      Width           =   1650
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "C&etak"
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   2280
      Width           =   1650
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kriteria Cetak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   8895
      Begin MSDataListLib.DataCombo dcJenisAwal 
         Height          =   405
         Left            =   1560
         TabIndex        =   0
         Top             =   480
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   714
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcJenisAkhir 
         Height          =   405
         Left            =   5400
         TabIndex        =   1
         Top             =   480
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   714
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
         Height          =   210
         Index           =   1
         Left            =   4980
         TabIndex        =   6
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Barang"
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   1005
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   7
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
      Picture         =   "frmDaftarCetakInputStokOpname.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7080
      Picture         =   "frmDaftarCetakInputStokOpname.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarCetakInputStokOpname.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmDaftarCetakInputStokOpname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub subLoadDcSource()
    strSQL = "SELECT KdDetailJenisBarang, DetailJenisBarang FROM V_DetailJenisBrgPerKelompokBrg WHERE KdKelompokBarang = '" & mstrKdKelompokBarang & "' AND StatusEnabled<>0 Order By DetailJenisBarang"
    Call msubDcSource(dcJenisAwal, rs, strSQL)
    Call msubDcSource(dcJenisAkhir, rs, strSQL)
End Sub

Private Sub cmdCetak_Click()
    mstrFilter = ""
    If Periksa("datacombo", dcJenisAwal, "Jenis barang awal kosong") = False Then Exit Sub
    If Periksa("datacombo", dcJenisAkhir, "Jenis barang akhir kosong") = False Then Exit Sub

    If mstrKdKelompokBarang = "02" Then         'medis
        strSQL = "SELECT * " & _
        " FROM V_DataStokBarangMedisNonRekap " & _
        " WHERE KdRuangan = '" & mstrKdRuangan & "' AND (JenisBarang BETWEEN '" & dcJenisAwal.Text & "' and '" & dcJenisAkhir.Text & "') ORDER BY JenisBarang, NamaBarang"
    ElseIf mstrKdKelompokBarang = "01" Then     'non medis
        strSQL = "SELECT * " & _
        " FROM V_DataStokBarangNonMedisNonRekap " & _
        " WHERE KdRuangan = '" & mstrKdRuangan & "' AND (JenisBarang BETWEEN '" & dcJenisAwal.Text & "' and '" & dcJenisAkhir.Text & "') ORDER BY JenisBarang, NamaBarang"
    End If
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        MsgBox "Tidak ada data, cek lagi kriteria cetak", vbExclamation, "Informasi"
        Exit Sub
    End If

    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    On Error GoTo hell
    frmCetakInputStokOpname.chkHarga = chkTampilHarga.value
    frmCetakInputStokOpname.Show
    Exit Sub
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcNamaBarangAkhir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dcJenisAkhir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dcJenisAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJenisAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subLoadDcSource
End Sub

