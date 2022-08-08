VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDaftarPemeriksaanPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Daftar Pemeriksaan Pasien"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPemeriksaanPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6435
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   6315
      Begin VB.Frame Frame4 
         Caption         =   "Dokter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3120
         TabIndex        =   10
         Top             =   1080
         Width           =   2415
         Begin MSDataListLib.DataCombo dcNamaDokter 
            Height          =   330
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   330
         Left            =   480
         TabIndex        =   2
         Top             =   1320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Frame Frame2 
         Caption         =   "Instalasi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   9
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Frame Frame3 
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
         Height          =   855
         Left            =   200
         TabIndex        =   8
         Top             =   150
         Width           =   5895
         Begin MSComCtl2.DTPicker DTPickerAwal 
            Height          =   375
            Left            =   720
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   64618499
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker DTPickerAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   64618499
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   11
            Top             =   315
            Width           =   255
         End
      End
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   3240
      Width           =   1665
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
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
      Left            =   4560
      Picture         =   "frmDaftarPemeriksaanPasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPemeriksaanPasien.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPemeriksaanPasien.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmDaftarPemeriksaanPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCetak_Click()
On Error GoTo errLoad

'If dcInstalasi.Text = "" Then
    'MsgBox "Pilih Instalasi", vbCritical, "Warning"
    'dcInstalasi.SetFocus
    'Exit Sub
'End If

If dcNamaDokter.Text = "" Then
    MsgBox "Pilih nama Dokter", vbCritical, "Warning"
    dcNamaDokter.SetFocus
    Exit Sub
End If

    
'
'strSQL = "SELECT DISTINCT TglPelayanan,NoCM, NamaPasien, Umur, NamaDiagnosa, JenisPelayanan, DokterPemeriksa, NamaInstalasi FROM V_RekapitulasiPelayananDokterPerPasienDetail " & _
'             " where KdInstalasi like '%" & dcInstalasi.BoundText & "%' " & _
'             " and IdDokter like '%" & dcNamaDokter.BoundText & "%' AND TglPelayanan BETWEEN '" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:59") & "'" & _
'             " Group by TglPelayanan,NoCM, NamaPasien, Umur, NamaDiagnosa, JenisPelayanan, DokterPemeriksa, NamaInstalasi " & _
'             " order by NoCM "

strSQL = "SELECT DISTINCT TglPelayanan,NoCM, NamaPasien, Umur, NamaDiagnosa, JenisPelayanan, Tindakan, DokterPemeriksa, NamaInstalasi FROM V_RekapitulasiPelayananDokterPerPasienDetail " & _
             " where KdInstalasi like '%" & dcInstalasi.BoundText & "%' " & _
             " and IdDokter like '%" & dcNamaDokter.BoundText & "%' AND TglPelayanan BETWEEN '" & Format(DTPickerAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd HH:mm:59") & "'" & _
             " and KdJnsPelayanan <> '001' and KdJnsPelayanan <> '301' and KdPelayananRS <> '301007' " & _
             " Group by TglPelayanan,NoCM, NamaPasien, Umur, NamaDiagnosa, JenisPelayanan, Tindakan, DokterPemeriksa, NamaInstalasi " & _
             " order by NoCM "

'Notes = KdJnsPelayanan = 001 => Administrasi
'        KdJnsPelayanan = 301 => Sewa Kamar Rawat Inap
'        KdPelayananRS  = 301007 => JasaPerawat
             
 Call msubRecFO(dbRst, strSQL)
  
 frmCetakDaftarPemeriksaanPasien.Show
 
Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub


Private Sub DTPickerAkhir_Change()
    DTPickerAkhir.MaxDate = Now
End Sub

Private Sub DTPickerAwal_Change()
    DTPickerAwal.MaxDate = Now
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
   Call subLoadDcSource
    With Me
        .DTPickerAwal.Value = Format(Now, "dd MMM yyyy 00:00:00")
        .DTPickerAkhir.Value = Now
    End With
    strSQL = "SELECT KdInstalasi, NamaInstalasi FROM Instalasi" 'WHERE KdInstalasi NOT IN ('05','07','13','14','15','17','18','19','20','21','23')"
    Call msubDcSource(dcInstalasi, dbRst, strSQL)
Exit Sub
errLoad:
    Call msubPesanError
End Sub
Private Sub subLoadDcSource()
On Error GoTo errLoad
   
    Call msubDcSource(dcNamaDokter, rs, "SELECT dbo.DataPegawai.IdPegawai, dbo.DataPegawai.NamaLengkap, dbo.DataPegawai.KdJenisPegawai, dbo.DataCurrentPegawai.KdStatus From  dbo.DataPegawai LEFT OUTER JOIN  dbo.DataCurrentPegawai ON dbo.DataPegawai.IdPegawai = dbo.DataCurrentPegawai.IdPegawai WHERE     (dbo.DataPegawai.KdJenisPegawai = '001') AND (dbo.DataCurrentPegawai.KdStatus = '01') AND NamaLengkap LIKE '%" & dcNamaDokter.Text & "%' order by NamaLengkap")
    
Exit Sub
errLoad:
    msubPesanError
End Sub


