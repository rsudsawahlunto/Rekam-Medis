VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDataPengunjung 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Pengunjung Rumah Sakit"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDataPengunjung.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10815
   Begin VB.Frame Frame5 
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
      Left            =   0
      TabIndex        =   24
      Top             =   4680
      Width           =   10725
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   360
         Left            =   3480
         TabIndex        =   25
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama Ruangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1440
         TabIndex        =   26
         Top             =   240
         Width           =   1380
      End
   End
   Begin VB.Frame fraButton 
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
      Left            =   0
      TabIndex        =   16
      Top             =   6480
      Width           =   10725
      Begin VB.TextBox txtDiagnosa 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3600
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Master Laporan"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   7080
         TabIndex        =   18
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   8880
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kd Diagnosa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2280
         TabIndex        =   35
         Top             =   270
         Width           =   1140
      End
   End
   Begin VB.Frame fraPeriode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1060
      Left            =   0
      TabIndex        =   7
      Top             =   5400
      Width           =   10725
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
         Height          =   825
         Left            =   720
         TabIndex        =   8
         Top             =   120
         Width           =   8895
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Group By"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   170
            Width           =   2895
            Begin VB.OptionButton optGroupBy 
               Caption         =   "Bulan"
               Height          =   210
               Index           =   1
               Left            =   960
               TabIndex        =   12
               Top             =   220
               Width           =   855
            End
            Begin VB.OptionButton optGroupBy 
               Caption         =   "Hari"
               Height          =   210
               Index           =   0
               Left            =   240
               TabIndex        =   11
               Top             =   220
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton optGroupBy 
               Caption         =   "Tahun"
               Height          =   210
               Index           =   2
               Left            =   1920
               TabIndex        =   10
               Top             =   220
               Width           =   855
            End
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   3720
            TabIndex        =   13
            Top             =   270
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
            OLEDropMode     =   1
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   141819907
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   6240
            TabIndex        =   14
            Top             =   270
            Width           =   2055
            _ExtentX        =   3625
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
            Format          =   141819907
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5880
            TabIndex        =   15
            Top             =   330
            Width           =   255
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kriteria Laporan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   10725
      Begin VB.OptionButton optIndeksPeny 
         Caption         =   "Laporan Indeks Penyakit"
         Height          =   255
         Left            =   6240
         TabIndex        =   34
         Top             =   2520
         Width           =   2295
      End
      Begin VB.OptionButton optKDRS 
         Caption         =   "KDRS"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   2520
         Width           =   2295
      End
      Begin VB.OptionButton optTindakanOperasi 
         Caption         =   "TINDAKAN OPERASI"
         Height          =   255
         Left            =   6240
         TabIndex        =   32
         Top             =   2160
         Width           =   2295
      End
      Begin VB.OptionButton optDPJP 
         Caption         =   "DPJP"
         Height          =   255
         Left            =   6240
         TabIndex        =   31
         Top             =   1800
         Width           =   1935
      End
      Begin VB.OptionButton optKunjRJ 
         Caption         =   "Kunjungan Pasien Rawat Jalan berdasarkan Poli"
         Height          =   255
         Left            =   6240
         TabIndex        =   30
         Top             =   1440
         Width           =   4215
      End
      Begin VB.OptionButton optJenisBayarKelas 
         Caption         =   "Data Pengunjung Berdasarkan Jenis Pembayaran && Kelas"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   5055
      End
      Begin VB.OptionButton optRekapFisioterapi 
         Caption         =   "Rekapitulasi Jumlah Pemeriksaan Fisioterapi"
         Height          =   255
         Left            =   6240
         TabIndex        =   28
         Top             =   1080
         Width           =   4215
      End
      Begin VB.OptionButton optJumlahPasienUSG 
         Caption         =   "Jumlah Pasien USG"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   1935
      End
      Begin VB.OptionButton optJumlahPasienEKG 
         Caption         =   "Jumlah Pasien EKG"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   1935
      End
      Begin VB.OptionButton optDataPasienRJByKodeWilayah 
         Caption         =   "Data Pasien Berdasarkan Kode Wilayah"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   4695
      End
      Begin VB.OptionButton optKecelakaan 
         Caption         =   "Daftar Pasien Kecelakaan Lalu Lintas"
         Height          =   255
         Left            =   6240
         TabIndex        =   20
         Top             =   720
         Width           =   3495
      End
      Begin VB.OptionButton optRekapUTD 
         Caption         =   "Rekap Jumlah Tindakan UTD"
         Height          =   255
         Left            =   6240
         TabIndex        =   19
         Top             =   360
         Width           =   3495
      End
      Begin VB.OptionButton optPasienRIWilayahJK 
         Caption         =   "Data Pasien Rawat Inap Berdasarkan Wilayah Menurut Jenis Kelamin"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   6015
      End
      Begin VB.OptionButton optJenisBayar 
         Caption         =   "Data Pengunjung Berdasarkan Jenis Pembayaran"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   4335
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
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   3960
      Width           =   10725
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   360
         Left            =   3480
         TabIndex        =   2
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Instalasi Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1755
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
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
   Begin VB.Image Image4 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataPengunjung.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmDataPengunjung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCetak_Click()
    Call Kriterialaporan
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        MsgBox "Tidak Ada Data", vbExclamation, "Validasi"
    Else
        If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
        frmCetakKunjunganRS.Show
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    frmMasterLaporan.Show
End Sub

Private Sub dcInstalasi_Change()
    If (optDataPasienRJByKodeWilayah.value = True Or optJenisBayar.value = True) And dcInstalasi.BoundText = "01" Then
        dcRuangan.Enabled = True
    Else
        dcRuangan.Enabled = False
    End If
    
    strSQL = "SELECT KdRuangan, NamaRuangan " & _
           "FROM Ruangan " & _
           "WHERE KdInstalasi = '" & dcInstalasi.BoundText & "' ORDER BY NamaRuangan"
    Call msubDcSource(dcRuangan, rs, strSQL)
    
    If optJumlahPasienEKG.value = True Or optJumlahPasienUSG.value = True Then dcRuangan.Enabled = True
    If optJenisBayar.value = True Then
        strCetak = "Jenis Pembayaran RJGD"
    End If
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    With Me
        .dtpAwal.value = Now
        .dtpAkhir.value = Now
        
        dtpAwal.CustomFormat = "dd MMMM yyyyy"
        dtpAkhir.CustomFormat = "dd MMMM yyyyy"
        
    End With
    
    strCetak = "Jenis Pembayaran RJGD"
    Call subDcSource
    dcRuangan.Enabled = False
End Sub

Private Sub subDcSource()
    strSQL = "SELECT KdInstalasi, NamaInstalasi " & _
    " From instalasi" & _
    " WHERE (KdInstalasi IN ('01', '02', '03', '06')) and StatusEnabled='1'"
    Call msubDcSource(dcInstalasi, rs, strSQL)
End Sub

Private Sub optDPJP_Click()
    strCetak = "DPJP"
    dcInstalasi.Enabled = False
    dcRuangan.Enabled = False
    dcInstalasi.Text = ""
    dcRuangan.Text = ""
    optGroupBy(0).Enabled = False
    optGroupBy(1).Enabled = True
    optGroupBy(2).Enabled = True
    optGroupBy(1).value = True
    txtdiagnosa.Enabled = False
End Sub

Private Sub optGroupBy_Click(index As Integer)
    Select Case index
        Case 0
            dtpAwal.CustomFormat = "dd MMMM yyyyy"
            dtpAkhir.CustomFormat = "dd MMMM yyyyy"
            optGroupBy(1).value = False
            optGroupBy(2).value = False

        Case 1
            dtpAkhir.CustomFormat = "MMMM yyyyy"
            dtpAwal.CustomFormat = "MMMM yyyyy"
            optGroupBy(0).value = False
            optGroupBy(2).value = False

        Case 2
            dtpAkhir.CustomFormat = "yyyyy"
            dtpAwal.CustomFormat = "yyyyy"
            optGroupBy(0).value = False
            optGroupBy(1).value = False
        
        Case 3
            dtpAwal.CustomFormat = "dd MMMM yyyyy"
            dtpAkhir.CustomFormat = "dd MMMM yyyyy"
            optGroupBy(1).value = False
            optGroupBy(2).value = False
    End Select
End Sub

Private Sub Kriterialaporan()
    Select Case strCetak
        Case "TindakanOperasi"
            If optGroupBy(1).value = True Then
                strSQL = "select NamaLengkap,NamaPenjamin,COUNT(NamaLengkap) as Jml from V_TindOK " & _
                         "where (MONTH(TglPulang) between '" & dtpAwal.Month & "' and '" & dtpAkhir.Month & "') " & _
                         "and (YEAR(TglPulang) between '" & dtpAwal.Year & "' and '" & dtpAkhir.Year & "') " & _
                         "group by NamaLengkap,NamaPenjamin"
            Else
                strSQL = "select NamaLengkap,NamaPenjamin,COUNT(NamaLengkap) as Jml from V_TindOK " & _
                         "where (YEAR(TglPulang) between '" & dtpAwal.Year & "' and '" & dtpAkhir.Year & "') " & _
                         "group by NamaLengkap,NamaPenjamin"
            End If
        Case "DPJP"
            If optGroupBy(1).value = True Then
                strSQL = "select NamaPenjamin,NamaLengkap,COUNT(NamaLengkap) as jml from V_DPJP " & _
                         "where (MONTH(TglPulang) between '" & dtpAwal.Month & "' and '" & dtpAkhir.Month & "') " & _
                         "and (YEAR(TglPulang) between '" & dtpAwal.Year & "' and '" & dtpAkhir.Year & "') " & _
                         "group by NamaPenjamin,NamaLengkap"
            Else
                strSQL = "select NamaPenjamin,NamaLengkap,COUNT(NamaLengkap) as jml from V_DPJP " & _
                         "where (YEAR(TglPulang) between '" & dtpAwal.Year & "' and '" & dtpAkhir.Year & "') " & _
                         "group by NamaPenjamin,NamaLengkap"
            End If
        Case "KunjRJ"
''            strSQL = "SELECT RegistrasiRJ.ThnTglMasuk,dbo.RegistrasiRJ.BlnTglMasuk, dbo.UbahBulan(RegistrasiRJ.BlnTglMasuk) AS Bln, Ruangan.NamaRuangan, RegistrasiRJ.StatusPasien, COUNT(RegistrasiRJ.NoPendaftaran) AS Jml " & _
''                     "FROM RegistrasiRJ INNER JOIN Ruangan ON RegistrasiRJ.KdRuangan = Ruangan.KdRuangan " & _
''                     "WHERE (NOT (RegistrasiRJ.NoPendaftaran IN (SELECT NoPendaftaran FROM PasienBatalDirawat))) AND Ruangan.KdInstalasi in ('02','06') AND ThnTglMasuk BETWEEN '" & dtpAwal.Year & "' AND '" & dtpAkhir.Year & "' " & _
''                     "GROUP BY RegistrasiRJ.ThnTglMasuk, RegistrasiRJ.BlnTglMasuk, Ruangan.NamaRuangan, RegistrasiRJ.StatusPasien"
            strSQL = "SELECT YEAR(TglPulang) AS ThnTglMasuk, MONTH(TglPulang) AS BlnTglMasuk, dbo.UbahBulan(MONTH(TglPulang)) AS Bln, " & _
                     "NamaRuangan, StatusPasien, 1 AS Jml FROM V_DataKunjunganPasienMasukBWilayahDStatusRJ2 " & _
                     "WHERE YEAR(TglPulang) BETWEEN '" & dtpAwal.Year & "' AND '" & dtpAkhir.Year & "' AND KdInstalasi in ('02','06')"
        Case "Jenis Pembayaran RJGD"
'            If dcInstalasi.BoundText = "01" Or dcInstalasi.BoundText = "02" Then
'                If optGroupBy(0).value = True Then
'                    strSQL = "SELECT 'Tahun' + STR(YEAR(TglPendaftaran)) as Tahun, dbo.UbahBulan(MONTH(TglPendaftaran)) As Bulan, MONTH(TglPendaftaran) As Bulan2, NamaRuangan, Kriteria, JmlPasien, KelompokPasien " & _
'                             "FROM   V_DataKunjunganPasienMasukBWilayahDStatusIGDNew " & _
'                             "WHERE (TglPendaftaran BETWEEN '" & Format(dtpAwal.value, "yyyy-mm-dd 00:00:00") & "' AND '" & Format(dtpAkhir.value, "yyyy-mm-dd 23:59:59") & "')  and KdInstalasi ='" & dcInstalasi.BoundText & "'"
'                ElseIf optGroupBy(1).value = True Then
'                    strSQL = "SELECT 'Tahun' + STR(YEAR(TglPendaftaran)) as Tahun, dbo.UbahBulan(MONTH(TglPendaftaran)) As Bulan, MONTH(TglPendaftaran) As Bulan2,NamaRuangan, Kriteria, JmlPasien, KelompokPasien " & _
'                             "FROM   V_DataKunjunganPasienMasukBWilayahDStatusIGDNew " & _
'                             "WHERE (month(TglPendaftaran) BETWEEN '" & dtpAwal.Month & "' AND '" & dtpAkhir.Month & "' AND year (tglpendaftaran) between '" & dtpAwal.Year & "' AND '" & dtpAkhir.Year & "') and KdInstalasi ='" & dcInstalasi.BoundText & "'"
'                Else
'                    strSQL = "SELECT 'Tahun' + STR(YEAR(TglPendaftaran)) as Tahun, dbo.UbahBulan(MONTH(TglPendaftaran)) As Bulan, MONTH(TglPendaftaran) As Bulan2, NamaRuangan, Kriteria, JmlPasien, KelompokPasien " & _
'                             "FROM   V_DataKunjunganPasienMasukBWilayahDStatusIGDNew " & _
'                             "WHERE (year(tglpendaftaran) between '" & dtpAwal.Year & "' AND '" & dtpAkhir.Year & "') and KdInstalasi ='" & dcInstalasi.BoundText & "'"
'                End If
'            Else
'                If optGroupBy(0).value = True Then
'                    strSQL = "SELECT 'Tahun' + STR(YEAR(TglPendaftaran)) as Tahun, dbo.UbahBulan(MONTH(TglPendaftaran)) As Bulan, MONTH(TglPendaftaran) As Bulan2, NamaRuangan, Kriteria, JmlPasien, KelompokPasien " & _
'                             "FROM   V_DataKunjunganPasienMasukBWilayahDStatusIGDNew " & _
'                             "WHERE (TglPendaftaran BETWEEN '" & Format(dtpAwal.value, "yyyy-mm-dd 00:00:00") & "' AND '" & Format(dtpAkhir.value, "yyyy-mm-dd 23:59:59") & "')  and KdInstalasi ='" & dcInstalasi.BoundText & "'"
'                ElseIf optGroupBy(1).value = True Then
'                    strSQL = "SELECT 'Tahun' + STR(YEAR(TglPendaftaran)) as Tahun, dbo.UbahBulan(MONTH(TglPendaftaran)) As Bulan, MONTH(TglPendaftaran) As Bulan2,NamaRuangan, Kriteria, JmlPasien, KelompokPasien " & _
'                             "FROM   V_DataKunjunganPasienMasukBWilayahDStatusIGDNew " & _
'                             "WHERE (month(TglPendaftaran) BETWEEN '" & dtpAwal.Month & "' AND '" & dtpAkhir.Month & "' AND year (tglpendaftaran) between '" & dtpAwal.Year & "' AND '" & dtpAkhir.Year & "') and KdInstalasi ='" & dcInstalasi.BoundText & "'"
'                Else
'                    strSQL = "SELECT 'Tahun' + STR(YEAR(TglPendaftaran)) as Tahun, dbo.UbahBulan(MONTH(TglPendaftaran)) As Bulan, MONTH(TglPendaftaran) As Bulan2, NamaRuangan, Kriteria, JmlPasien, KelompokPasien " & _
'                             "FROM   V_DataKunjunganPasienMasukBWilayahDStatusIGDNew " & _
'                             "WHERE (year(tglpendaftaran) between '" & dtpAwal.Year & "' AND '" & dtpAkhir.Year & "') and KdInstalasi ='" & dcInstalasi.BoundText & "'"
'                End If
'            End If
            If dcInstalasi.BoundText = "01" Then
                If optGroupBy(0).value = True Then
                    strSQL = "SELECT 'Tahun' + STR(YEAR(TglPulang)) as Tahun, dbo.UbahBulan(MONTH(TglPulang)) As Bulan, MONTH(TglPulang) As Bulan2, NamaRuangan, Kriteria, JmlPasien, KelompokPasien FROM   V_DataKunjunganPasienMasukBWilayahDStatusIGDNew3 WHERE (TglPulang between '" & Format(dtpAwal.value, "yyyy-mm-dd 00:00:00") & "' AND '" & Format(dtpAkhir.value, "yyyy-mm-dd 23:59:59") & "') and KdInstalasi ='" & dcInstalasi.BoundText & "' AND NamaRuangan LIKE '%" & dcRuangan.Text & "%'"
                ElseIf optGroupBy(1).value = True Then
                    strSQL = "SELECT 'Tahun' + STR(YEAR(TglPulang)) as Tahun, dbo.UbahBulan(MONTH(TglPulang)) As Bulan, MONTH(TglPulang) As Bulan2, NamaRuangan, Kriteria, JmlPasien, KelompokPasien FROM   V_DataKunjunganPasienMasukBWilayahDStatusIGDNew3 " & _
                             "WHERE (MONTH(TglPulang) between '" & dtpAwal.Month & "' AND '" & dtpAkhir.Month & "') " & _
                             "AND (YEAR(TglPulang) between '" & dtpAwal.Year & "' AND '" & dtpAkhir.Year & "') " & _
                             "and KdInstalasi ='" & dcInstalasi.BoundText & "' AND NamaRuangan LIKE '%" & dcRuangan.Text & "%'"
                Else
                    strSQL = "SELECT 'Tahun' + STR(YEAR(TglPulang)) as Tahun, dbo.UbahBulan(MONTH(TglPulang)) As Bulan, MONTH(TglPulang) As Bulan2, NamaRuangan, Kriteria, JmlPasien, KelompokPasien FROM   V_DataKunjunganPasienMasukBWilayahDStatusIGDNew3 " & _
                             "WHERE (YEAR(TglPulang) between '" & dtpAwal.Year & "' AND '" & dtpAkhir.Year & "') and KdInstalasi ='" & dcInstalasi.BoundText & "' AND NamaRuangan LIKE '%" & dcRuangan.Text & "%'"
                End If
            Else
                If optGroupBy(0).value = True Then
                    strSQL = "SELECT 'Tahun' + STR(YEAR(TglPulang)) as Tahun, dbo.UbahBulan(MONTH(TglPulang)) As Bulan, MONTH(TglPulang) As Bulan2, NamaRuangan, Kriteria, JmlPasien, KelompokPasien FROM   V_DataKunjunganPasienMasukBWilayahDStatusIGDNew3 WHERE (TglPulang between '" & Format(dtpAwal.value, "yyyy-mm-dd 00:00:00") & "' AND '" & Format(dtpAkhir.value, "yyyy-mm-dd 23:59:59") & "') and KdInstalasi ='" & dcInstalasi.BoundText & "'"
                ElseIf optGroupBy(1).value = True Then
                    strSQL = "SELECT 'Tahun' + STR(YEAR(TglPulang)) as Tahun, dbo.UbahBulan(MONTH(TglPulang)) As Bulan, MONTH(TglPulang) As Bulan2, NamaRuangan, Kriteria, JmlPasien, KelompokPasien FROM   V_DataKunjunganPasienMasukBWilayahDStatusIGDNew3 " & _
                             "WHERE (MONTH(TglPulang) between '" & dtpAwal.Month & "' AND '" & dtpAkhir.Month & "') " & _
                             "AND (YEAR(TglPulang) between '" & dtpAwal.Year & "' AND '" & dtpAkhir.Year & "') " & _
                             "and KdInstalasi ='" & dcInstalasi.BoundText & "'"
                Else
                    strSQL = "SELECT 'Tahun' + STR(YEAR(TglPulang)) as Tahun, dbo.UbahBulan(MONTH(TglPulang)) As Bulan, MONTH(TglPulang) As Bulan2, NamaRuangan, Kriteria, JmlPasien, KelompokPasien FROM   V_DataKunjunganPasienMasukBWilayahDStatusIGDNew3 " & _
                             "WHERE (YEAR(TglPulang) between '" & dtpAwal.Year & "' AND '" & dtpAkhir.Year & "') and KdInstalasi ='" & dcInstalasi.BoundText & "'"
                End If
            End If
        Case "Jenis Pembayaran RJGD2"
            If optGroupBy(0).value = True Then
                strSQL = "SELECT KelompokPasien,Kriteria,DeskKelas, 1 AS Jml FROM V_DataKunjunganPasienMasukBWilayahDStatusRI " & _
                         "WHERE (TglPulang between '" & Format(dtpAwal.value, "yyyy-mm-dd 00:00:00") & "' AND '" & Format(dtpAkhir.value, "yyyy-mm-dd 23:59:59") & "')"
            ElseIf optGroupBy(1).value = True Then
                strSQL = "SELECT KelompokPasien,Kriteria,DeskKelas, 1 AS Jml FROM V_DataKunjunganPasienMasukBWilayahDStatusRI " & _
                         "WHERE (MONTH(TglPulang) between '" & dtpAwal.Month & "' AND '" & dtpAkhir.Month & "') " & _
                         "AND (YEAR(TglPulang) between '" & dtpAwal.Year & "' AND '" & dtpAkhir.Year & "') "
            Else
                strSQL = "SELECT KelompokPasien,Kriteria,DeskKelas, 1 AS Jml FROM V_DataKunjunganPasienMasukBWilayahDStatusRI " & _
                         "WHERE (YEAR(TglPulang) between '" & dtpAwal.Year & "' AND '" & dtpAkhir.Year & "')"
            End If
            
        Case "Rekapitulasi Jumlah Pemeriksaan UTD"
                    strSQL = "SELECT 'TAHUN' + STR(YEAR(BiayaPelayanan.TglPelayanan)) AS Tahun, dbo.UbahBulan(MONTH(BiayaPelayanan.TglPelayanan)) AS Bulan, MONTH(BiayaPelayanan.TglPelayanan) AS Bulan2,SUM(BiayaPelayanan.JmlPelayanan) AS JumlahPelayanan, ListPelayananRS.NamaPelayanan " & _
                             "FROM BiayaPelayanan INNER JOIN ListPelayananRS ON BiayaPelayanan.KdPelayananRS = ListPelayananRS.KdPelayananRS " & _
                             "WHERE (BiayaPelayanan.KdRuangan = '903') AND (BiayaPelayanan.KdPelayananRS = '086010') and (year(BiayaPelayanan.TglPelayanan) between '" & dtpAwal.Year & "' AND '" & dtpAkhir.Year & "') " & _
                             "GROUP BY YEAR(BiayaPelayanan.TglPelayanan), MONTH(BiayaPelayanan.TglPelayanan), ListPelayananRS.NamaPelayanan, 'TAHUN' + STR(YEAR(BiayaPelayanan.TglPelayanan)),dbo.UbahBulan(MONTH(BiayaPelayanan.TglPelayanan))"
        Case "WilayahJekelRI"
            If optGroupBy(0).value = True Then
                strSQL = "select TglKeluar, NamaRuangan,Case When Kriteria='L' then 'Luar Kota' Else 'Dalam Kota' end as Kriteria,DeskKelas, JenisKelamin, 1 as Jml from V_DataKunjunganPasienMasukBWilayahDStatusIGDNew4 " & _
                         " where TglKeluar between '" & Format(dtpAwal.value, "yyyy-mm-dd 00:00:00") & "' and '" & Format(dtpAkhir.value, "yyyy-mm-dd 23:59:59") & "'"
            ElseIf optGroupBy(1).value = True Then
                strSQL = "select TglKeluar, NamaRuangan,Case When Kriteria='L' then 'Luar Kota' Else 'Dalam Kota' end as Kriteria,DeskKelas, JenisKelamin, 1 as Jml from V_DataKunjunganPasienMasukBWilayahDStatusIGDNew4 " & _
                         " where (year(TglKeluar) between '" & dtpAwal.Year & "' and '" & dtpAkhir.Year & "') " & _
                         "AND (MONTH(TglKeluar) between '" & dtpAwal.Month & "' and '" & dtpAkhir.Month & "')"
            Else
                strSQL = "select TglKeluar, NamaRuangan,Case When Kriteria='L' then 'Luar Kota' Else 'Dalam Kota' end as Kriteria,DeskKelas, JenisKelamin, 1 as Jml from V_DataKunjunganPasienMasukBWilayahDStatusIGDNew4 " & _
                         " where year(TglKeluar) between '" & dtpAwal.Year & "' and '" & dtpAkhir.Year & "'"
            End If

        Case "Daftar Pasien Kecelakaan"
'                    strSQL = "SELECT Pasien.NamaLengkap, dbo.S_HitungUmur(Pasien.TglLahir, PeriksaDiagnosa.TglMasuk) AS Umur, PasienDaftar.TglPendaftaran, CASE WHEN PeriksaDiagnosa.KdJenisDiagnosa = '05' THEN PeriksaDiagnosa.KdDiagnosa ELSE '-' END AS Diagnosa, Instalasi.NamaInstalasi " & _
'                             "FROM PeriksaDiagnosa RIGHT OUTER JOIN PasienDaftar INNER JOIN RegistrasiIGD ON PasienDaftar.NoPendaftaran = RegistrasiIGD.NoPendaftaran RIGHT OUTER JOIN Ruangan INNER JOIN Instalasi ON Ruangan.KdInstalasi = Instalasi.KdInstalasi ON PasienDaftar.KdRuanganAkhir = Ruangan.KdRuangan LEFT OUTER JOIN Pasien ON PasienDaftar.NoCM = Pasien.NoCM ON PeriksaDiagnosa.NoPendaftaran = PasienDaftar.NoPendaftaran " & _
'                             "WHERE (RegistrasiIGD.KdRujukanAsal IN ('07', '16')) and TglPendaftaran between '" & Format(dtpAwal.value, "yyyy-mm-dd 00:00:00") & "' and '" & Format(dtpAkhir.value, "yyyy-mm-dd 23:59:59") & "'"
                    strSQL = "SELECT     KasusKecelakaan.NoCM, Pasien.NamaLengkap, dbo.S_HitungUmur(Pasien.TglLahir, KasusKecelakaan.TglKecelakaan) AS Umur, PasienDaftar.TglPendaftaran, " & _
                             "CASE WHEN PeriksaDiagnosa.KdJenisDiagnosa = '05' THEN PeriksaDiagnosa.KdDiagnosa ELSE '-' END AS Diagnosa, Instalasi.NamaInstalasi " & _
                             "FROM         KasusKecelakaan INNER JOIN " & _
                             "Pasien ON KasusKecelakaan.NoCM = Pasien.NoCM INNER JOIN " & _
                             "PasienDaftar ON KasusKecelakaan.NoPendaftaran = PasienDaftar.NoPendaftaran INNER JOIN " & _
                             "Ruangan ON PasienDaftar.KdRuanganAkhir = Ruangan.KdRuangan INNER JOIN " & _
                             "Instalasi ON Ruangan.KdInstalasi = Instalasi.KdInstalasi LEFT OUTER JOIN " & _
                             "PeriksaDiagnosa ON KasusKecelakaan.NoPendaftaran = PeriksaDiagnosa.NoPendaftaran " & _
                             "WHERE PasienDaftar.TglPendaftaran BETWEEN '" & Format(dtpAwal.value, "yyyy-mm-dd 00:00:00") & "' AND '" & Format(dtpAkhir.value, "yyyy-mm-dd 23:59:59") & "'"
        Case "Data Pasien RJ By Kode Wilayah"
'                    strSQL = "SELECT dbo.UbahBulan(MONTH(RegistrasiRJ.TglMasuk)) AS Bulan, MONTH(RegistrasiRJ.TglMasuk) AS Bulan1, YEAR(RegistrasiRJ.TglMasuk) AS Tahun, COUNT(CASE WHEN dbo.Ambil_KodeWilayah(Pasien.Alamat, 'W') = 'D' THEN 1 ELSE 0 END) AS 'Dalam', COUNT(CASE WHEN dbo.Ambil_KodeWilayah(Pasien.Alamat, 'W') = 'L' THEN 1 ELSE 0 END) AS 'Luar' " & _
'                             "FROM RegistrasiRJ INNER JOIN Pasien ON RegistrasiRJ.NoCM = Pasien.NoCM " & _
'                             "WHERE (year(TglMasuk) between '" & dtpAwal.Year & "' AND '" & dtpAkhir.Year & "') " & _
'                             "GROUP BY YEAR(RegistrasiRJ.TglMasuk), MONTH(RegistrasiRJ.TglMasuk) " & _
'                             "ORDER BY Tahun, Bulan1"
            
            If dcInstalasi.BoundText = "01" Or dcInstalasi.BoundText = "02" Or dcInstalasi.BoundText = "06" Then
                If dcInstalasi.BoundText = "01" Then
    '                strSQL = "select YEAR(TglPendaftaran) as Tahun,dbo.UbahBulan(MONTH(TglPendaftaran)) as Bulan,MONTH(TglPendaftaran) as Bulan1, SUM(case when Kriteria='D' then 1 else 0 end) as Dalam, SUM(case when Kriteria='L' then 1 else 0 end) as Luar from V_DataKunjunganPasienMasukBWilayahDStatusIGDNew where KdInstalasi='" & dcInstalasi.BoundText & "' and (YEAR(TglPendaftaran) between '" & dtpAwal.Year & "' and '" & dtpAkhir.Year & "') group by year(TglPendaftaran),month(TglPendaftaran)"
                    strSQL = "select YEAR(TglPulang) as Tahun,dbo.UbahBulan(MONTH(TglPulang)) as Bulan,MONTH(TglPulang) as Bulan1, SUM(case when Kriteria='D' then 1 else 0 end) as Dalam, SUM(case when Kriteria='L' then 1 else 0 end) as Luar from V_DataKunjunganPasienMasukBWilayahDStatusIGDNew3 where KdInstalasi='" & dcInstalasi.BoundText & "' and (YEAR(TglPulang) between '" & dtpAwal.Year & "' and '" & dtpAkhir.Year & "') AND NamaRuangan LIKE '%" & dcRuangan.Text & "%' group by year(TglPulang),month(TglPulang) ORDER By Bulan1"
                Else
                    strSQL = "select YEAR(TglPulang) as Tahun,dbo.UbahBulan(MONTH(TglPulang)) as Bulan,MONTH(TglPulang) as Bulan1, SUM(case when Kriteria='D' then 1 else 0 end) as Dalam, SUM(case when Kriteria='L' then 1 else 0 end) as Luar from V_DataKunjunganPasienMasukBWilayahDStatusIGDNew3 where KdInstalasi='" & dcInstalasi.BoundText & "' and (YEAR(TglPulang) between '" & dtpAwal.Year & "' and '" & dtpAkhir.Year & "') group by year(TglPulang),month(TglPulang) ORDER By Bulan1"
                End If
            Else
'                strSQL = "select YEAR(TglKeluar) as Tahun,dbo.UbahBulan(MONTH(TglKeluar)) as Bulan,MONTH(TglKeluar) as Bulan1, SUM(case when Kriteria='D' then 1 else 0 end) as Dalam, SUM(case when Kriteria='L' then 1 else 0 end) as Luar, DeskKelas from V_DataKunjunganPasienMasukBWilayahDStatusIGDNew2 where KdInstalasi='" & dcInstalasi.BoundText & "' and (YEAR(TglKeluar) between '" & dtpAwal.Year & "' and '" & dtpAkhir.Year & "') group by year(TglKeluar),month(TglKeluar), DeskKelas order by month(TglKeluar)"
                strSQL = "select YEAR(TglKeluar) as Tahun,dbo.UbahBulan(MONTH(TglKeluar)) as Bulan, MONTH(TglKeluar) as Bulan1, case when Kriteria='D' then 'DALAM KOTA' else 'LUAR KOTA' end as Kriteria, DeskKelas, 1 as Jml from V_DataKunjunganPasienMasukBWilayahDStatusIGDNew4 where KdInstalasi='" & dcInstalasi.BoundText & "' and (YEAR(TglKeluar) between '" & dtpAwal.Year & "' and '" & dtpAkhir.Year & "') group by TglKeluar, DeskKelas, Kriteria ORDER BY MONTH(TglKeluar)"
            End If
        Case "Jumlah Pasien EKG"
                strSQL = "SELECT case when ListPelayananRS.NamaPelayanan like '%EKG%' then 1 else 0 end as Jumlah, ListPelayananRS.NamaPelayanan, Ruangan.NamaRuangan,BiayaPelayanan.KdRuangan, dbo.UbahBulan(month(BiayaPelayanan.TglPelayanan)) as Bulan, month(BiayaPelayanan.TglPelayanan) as Bulan1, YEAR(BiayaPelayanan.TglPelayanan) as Tahun, Penjamin.NamaPenjamin " & _
                         "FROM Penjamin INNER JOIN PasienDaftar ON Penjamin.IdPenjamin = PasienDaftar.IdPenjamin RIGHT OUTER JOIN ListPelayananRS INNER JOIN BiayaPelayanan ON ListPelayananRS.KdPelayananRS = BiayaPelayanan.KdPelayananRS INNER JOIN Ruangan ON BiayaPelayanan.KdRuangan = Ruangan.KdRuangan ON PasienDaftar.NoPendaftaran = BiayaPelayanan.NoPendaftaran " & _
                         "WHERE BiayaPelayanan.KdRuangan LIKE '%" & dcRuangan.BoundText & "%' and ListPelayananRS.NamaPelayanan LIKE '%EKG%' and (YEAR(TglPelayanan) between '" & dtpAwal.Year & "' and '" & dtpAkhir.Year & "')"
        Case "Jumlah Pasien USG"
                strSQL = "SELECT case when ListPelayananRS.NamaPelayanan like '%USG%' then 1 else 0 end as Jumlah, ListPelayananRS.NamaPelayanan, Ruangan.NamaRuangan,BiayaPelayanan.KdRuangan, dbo.UbahBulan(month(BiayaPelayanan.TglPelayanan)) as Bulan, month(BiayaPelayanan.TglPelayanan) as Bulan1, YEAR(BiayaPelayanan.TglPelayanan) as Tahun, Penjamin.NamaPenjamin " & _
                         "FROM Penjamin INNER JOIN PasienDaftar ON Penjamin.IdPenjamin = PasienDaftar.IdPenjamin RIGHT OUTER JOIN ListPelayananRS INNER JOIN BiayaPelayanan ON ListPelayananRS.KdPelayananRS = BiayaPelayanan.KdPelayananRS INNER JOIN Ruangan ON BiayaPelayanan.KdRuangan = Ruangan.KdRuangan ON PasienDaftar.NoPendaftaran = BiayaPelayanan.NoPendaftaran " & _
                         "WHERE BiayaPelayanan.KdRuangan='" & dcRuangan.BoundText & "' and ListPelayananRS.NamaPelayanan LIKE '%USG%' and (YEAR(TglPelayanan) between '" & dtpAwal.Year & "' and '" & dtpAkhir.Year & "')"
        Case "Rekap Fisioterapi"
                strSQL = "SELECT 1 AS jumlah, Case when Instalasi.KdInstalasi='02' Then 'RJ' else 'RI' end as Instalasi, dbo.UbahBulan(MONTH(BiayaPelayanan.TglPelayanan)) AS bulan, MONTH(BiayaPelayanan.TglPelayanan) AS bulan1, YEAR(BiayaPelayanan.TglPelayanan) AS tahun, ListPelayananRS.NamaPelayanan " & _
                         "FROM BiayaPelayanan LEFT OUTER JOIN Ruangan ON BiayaPelayanan.KdRuanganAsal = Ruangan.KdRuangan LEFT OUTER JOIN Instalasi ON Ruangan.KdInstalasi = Instalasi.KdInstalasi LEFT OUTER JOIN ListPelayananRS ON BiayaPelayanan.KdPelayananRS = ListPelayananRS.KdPelayananRS " & _
                         "WHERE (BiayaPelayanan.KdRuangan = '601') AND (ListPelayananRS.KdPelayananRS <> '001001') AND (ListPelayananRS.KdPelayananRS <> '888001') AND (ListPelayananRS.KdPelayananRS <> '888007') AND (Instalasi.KdInstalasi = '02' OR Instalasi.KdInstalasi = '03') and (YEAR(TglPelayanan) between '" & dtpAwal.Year & "' and '" & dtpAkhir.Year & "')"
        Case "KDRS"
                If dcInstalasi.BoundText <> "" Then
                    strFilter = "AND Instalasi.KdInstalasi='" & dcInstalasi.BoundText & "'"
                Else
                    strFilter = ""
                End If
                
                If optGroupBy(0).value = True Then
                    strFilter = strFilter & "AND PasienDaftar.TglPulang BETWEEN '" & Format(dtpAwal.value, "yyyy-mm-dd 00:00:00") & "' AND '" & Format(dtpAkhir.value, "yyyy-mm-dd 23:59:59") & "'"
                End If
                strSQL = "SELECT PeriksaDiagnosa.NoPendaftaran, PeriksaDiagnosa.KdDiagnosa, PeriksaDiagnosa.NoCM, Pasien.NamaLengkap, dbo.S_HitungUmurTahun(Pasien.TglLahir, PasienDaftar.TglPendaftaran) AS Umur, " & _
                         "CASE WHEN Pasien.JenisKelamin = 'L' THEN 'Laki-laki' ELSE 'Perempuan' END AS JenisKelamin, CASE WHEN DetailPasien.NamaIbu IS NULL " & _
                         "THEN '-' ELSE DetailPasien.NamaIbu END AS NamaIbu, Pasien.Alamat, CONVERT(CHAR,PasienDaftar.TglPendaftaran,103) AS TglPendaftaran, CONVERT(CHAR,PasienDaftar.TglPulang,103) AS TglPulang, KondisiPulang.KondisiPulang, Instalasi.KdInstalasi, " & _
                         "instalasi.NamaInstalasi , Ruangan.KdRuangan, Ruangan.namaRuangan " & _
                         "FROM PeriksaDiagnosa INNER JOIN Pasien ON PeriksaDiagnosa.NoCM = Pasien.NoCM INNER JOIN PasienDaftar ON PeriksaDiagnosa.NoPendaftaran = PasienDaftar.NoPendaftaran INNER JOIN " & _
                         "DetailPasien ON Pasien.NoCM = DetailPasien.NoCM INNER JOIN Ruangan ON PasienDaftar.KdRuanganAkhir = Ruangan.KdRuangan INNER JOIN Instalasi ON Ruangan.KdInstalasi = Instalasi.KdInstalasi LEFT OUTER JOIN " & _
                         "KondisiPulang INNER JOIN PasienPulang ON KondisiPulang.KdKondisiPulang = PasienPulang.KdKondisiPulang ON PeriksaDiagnosa.NoPendaftaran = PasienPulang.NoPendaftaran " & _
                         "WHERE (PeriksaDiagnosa.KdDiagnosa IN ('P71.3', 'A80', 'B05', 'A90', 'A91', 'B54', 'B50'))" & strFilter & _
                         " ORDER BY PasienDaftar.TglPulang"
        
        Case "IndeksPeny"
            strSQL = "SELECT PeriksaDiagnosa.KdDiagnosa, Diagnosa.NamaDiagnosa, PeriksaDiagnosa.NoCM, Pasien.NamaLengkap, PasienDaftar.TglPendaftaran, PasienDaftar.TglPulang, DATEDIFF(DAY,PasienDaftar.TglPendaftaran, PasienDaftar.TglPulang) + 1 AS LOS, DATEDIFF(YEAR, Pasien.TglLahir, PasienDaftar.TglPendaftaran) AS Umur, " & _
                     "Ruangan.NamaRuangan, dbo.FB_TakeNilaiBiaya(PeriksaDiagnosa.NoPendaftaran, 'A', 'TB') AS TotalBiaya, dbo.FB_TakeBiayaTotalKelas(PeriksaDiagnosa.NoPendaftaran, '05') AS VIP, " & _
                     "dbo.FB_TakeBiayaTotalKelas(PeriksaDiagnosa.NoPendaftaran, '03') AS I, dbo.FB_TakeBiayaTotalKelas(PeriksaDiagnosa.NoPendaftaran, '02') AS II, " & _
                     "dbo.FB_TakeBiayaTotalKelas(PeriksaDiagnosa.NoPendaftaran, '01') AS III " & _
                     "FROM PeriksaDiagnosa INNER JOIN " & _
                     "Diagnosa ON PeriksaDiagnosa.KdDiagnosa = Diagnosa.KdDiagnosa INNER JOIN " & _
                     "Pasien ON PeriksaDiagnosa.NoCM = Pasien.NoCM INNER JOIN " & _
                     "PasienDaftar ON PeriksaDiagnosa.NoPendaftaran = PasienDaftar.NoPendaftaran INNER JOIN " & _
                     "Ruangan ON PasienDaftar.KdRuanganAkhir = Ruangan.KdRuangan " & _
                     "WHERE (PasienDaftar.TglPulang BETWEEN '" & Format(dtpAwal.value, "yyyy-mm-dd 00:00:00") & "' AND '" & Format(dtpAkhir.value, "yyyy-mm-dd 23:59:59") & "') AND (PeriksaDiagnosa.KdDiagnosa LIKE '%" & txtdiagnosa.Text & "%')"
    End Select
End Sub

Private Sub optIndeksPeny_Click()
    strCetak = "IndeksPeny"
    dcInstalasi.Enabled = False
    optGroupBy(0).Enabled = True
    optGroupBy(1).Enabled = False
    optGroupBy(2).Enabled = False
    optGroupBy(0).value = True
    dcRuangan.Enabled = False
    txtdiagnosa.Enabled = True
End Sub

Private Sub optJenisBayar_Click()
'    If dcInstalasi.BoundText = "02" Or dcInstalasi.BoundText = "01" Then
        strCetak = "Jenis Pembayaran RJGD"
'    ElseIf dcInstalasi.BoundText = "03" Then
'        strCetak = "Jenis Pembayaran RI"
'    End If
    
    dcInstalasi.Enabled = True
    optGroupBy(0).Enabled = True
    optGroupBy(1).Enabled = True
    dcRuangan.Enabled = False
    txtdiagnosa.Enabled = False
End Sub


Private Sub optDataPasienRJByKodeWilayah_Click()
    strCetak = "Data Pasien RJ By Kode Wilayah"
    dcInstalasi.Enabled = True
    optGroupBy(0).Enabled = False
    optGroupBy(1).Enabled = False
'    optGroupBy(0).Enabled = True
'    optGroupBy(1).Enabled = True
    optGroupBy(2).Enabled = True
    optGroupBy(2).value = True
    dcRuangan.Enabled = False
    txtdiagnosa.Enabled = False
End Sub

Private Sub optJenisBayarKelas_Click()

    strCetak = "Jenis Pembayaran RJGD2"
    
    dcInstalasi.Enabled = False
    dcRuangan.Enabled = False
    
    optGroupBy(0).Enabled = True
    optGroupBy(1).Enabled = True
    optGroupBy(2).Enabled = True
    txtdiagnosa.Enabled = False
End Sub

Private Sub optJumlahPasienEKG_Click()
    strCetak = "Jumlah Pasien EKG"
    dcInstalasi.Enabled = True
    optGroupBy(0).Enabled = False
    optGroupBy(1).Enabled = False
    optGroupBy(2).Enabled = True
    optGroupBy(2).value = True
    dcRuangan.Enabled = True
    txtdiagnosa.Enabled = False
End Sub

Private Sub optJumlahPasienUSG_Click()
    strCetak = "Jumlah Pasien USG"
    dcInstalasi.Enabled = True
    optGroupBy(0).Enabled = False
    optGroupBy(1).Enabled = False
    optGroupBy(2).Enabled = True
    optGroupBy(2).value = True
    dcRuangan.Enabled = True
    txtdiagnosa.Enabled = False
End Sub

Private Sub optKDRS_Click()
    strCetak = "KDRS"
    dcInstalasi.Enabled = True
    dcRuangan.Enabled = True
    dcInstalasi.Text = ""
    dcRuangan.Text = ""
    optGroupBy(0).Enabled = True
    optGroupBy(1).Enabled = False
    optGroupBy(2).Enabled = False
    optGroupBy(0).value = True
    txtdiagnosa.Enabled = False
End Sub

Private Sub optKecelakaan_Click()
    strCetak = "Daftar Pasien Kecelakaan"
    dcInstalasi.Enabled = False
    optGroupBy(0).Enabled = True
    optGroupBy(1).Enabled = False
    optGroupBy(2).Enabled = False
    dcRuangan.Enabled = False
    txtdiagnosa.Enabled = False
End Sub

Private Sub optKunjRJ_Click()
    strCetak = "KunjRJ"
    dcInstalasi.Enabled = False
    dcInstalasi.Text = ""
    optGroupBy(0).Enabled = False
    optGroupBy(1).Enabled = False
    optGroupBy(2).Enabled = True
    optGroupBy(2).value = True
    dcRuangan.Enabled = False
    txtdiagnosa.Enabled = False
End Sub

Private Sub optPasienRIWilayahJK_Click()
    strCetak = "WilayahJekelRI"
    dcInstalasi.Enabled = False
    dcInstalasi.Text = ""
'    optGroupBy(0).Enabled = False
'    optGroupBy(1).Enabled = False
    optGroupBy(0).Enabled = True
    optGroupBy(1).Enabled = True
    optGroupBy(2).Enabled = True
'    optGroupBy(2).value = True
    dcRuangan.Enabled = False
    txtdiagnosa.Enabled = False
End Sub

Private Sub optRekapFisioterapi_Click()
    strCetak = "Rekap Fisioterapi"
    dcInstalasi.Enabled = False
    dcInstalasi.Text = ""
    optGroupBy(0).Enabled = False
    optGroupBy(1).Enabled = False
    optGroupBy(2).Enabled = True
    optGroupBy(2).value = True
    dcRuangan.Enabled = False
    txtdiagnosa.Enabled = False
End Sub

Private Sub optRekapUTD_Click()
    strCetak = "Rekapitulasi Jumlah Pemeriksaan UTD"
    dcInstalasi.Enabled = False
    optGroupBy(0).Enabled = False
    optGroupBy(1).Enabled = False
    optGroupBy(2).value = True
    dcRuangan.Enabled = False
    txtdiagnosa.Enabled = False
End Sub

Private Sub optTindakanOperasi_Click()
    strCetak = "TindakanOperasi"
    dcInstalasi.Enabled = False
    dcRuangan.Enabled = False
    dcInstalasi.Text = ""
    dcRuangan.Text = ""
    optGroupBy(0).Enabled = False
    optGroupBy(1).Enabled = True
    optGroupBy(2).Enabled = True
    optGroupBy(1).value = True
    txtdiagnosa.Enabled = False
End Sub
