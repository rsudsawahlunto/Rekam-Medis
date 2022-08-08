VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLapRKP_KPSK2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kunjungan Pasien "
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLapRKP_KPSK2.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   9405
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
      TabIndex        =   10
      Top             =   1680
      Width           =   9405
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
         Height          =   830
         Left            =   240
         TabIndex        =   11
         Top             =   120
         Width           =   8895
         Begin VB.Frame Frame1 
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
            TabIndex        =   13
            Top             =   170
            Width           =   3015
            Begin VB.OptionButton optGroupBy 
               Caption         =   "Total"
               Height          =   210
               Index           =   3
               Left            =   3000
               TabIndex        =   16
               Top             =   220
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.OptionButton optGroupBy 
               Caption         =   "Tahun"
               Height          =   210
               Index           =   2
               Left            =   1920
               TabIndex        =   3
               Top             =   220
               Width           =   975
            End
            Begin VB.OptionButton optGroupBy 
               Caption         =   "Hari"
               Height          =   210
               Index           =   0
               Left            =   240
               TabIndex        =   1
               Top             =   220
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton optGroupBy 
               Caption         =   "Bulan"
               Height          =   210
               Index           =   1
               Left            =   960
               TabIndex        =   2
               Top             =   220
               Width           =   735
            End
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   3720
            TabIndex        =   4
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
            OLEDropMode     =   1
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   127270915
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   6240
            TabIndex        =   5
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
            Format          =   127270915
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
            TabIndex        =   12
            Top             =   330
            Width           =   255
         End
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
      TabIndex        =   9
      Top             =   2760
      Width           =   9405
      Begin VB.CommandButton cmdgrafik 
         Caption         =   "&Grafik"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   5640
         TabIndex        =   7
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   7440
         TabIndex        =   8
         Top             =   240
         Width           =   1695
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
      TabIndex        =   14
      Top             =   960
      Width           =   9405
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   360
         Left            =   3480
         TabIndex        =   0
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
         TabIndex        =   15
         Top             =   240
         Width           =   1755
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   17
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
      Left            =   7560
      Picture         =   "frmLapRKP_KPSK2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmLapRKP_KPSK2.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLapRKP_KPSK2.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmLapRKP_KPSK2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrInstalasi2 As String

Sub Kriterialaporan()
    On Error GoTo hell

    Dim mdtBulan As Integer
    Dim MdtTahun As Integer

    If (optGroupBy(0).value = True) Or optGroupBy(3).value = True Then
        mdTglAwal = dtpAwal.value 'TglAwal
        mdTglAkhir = dtpAkhir.value 'TglAkhir
        mstrKdInstalasi = dcInstalasi.BoundText
        mstrInstalasi2 = dcInstalasi.Text
        Select Case strCetak
            Case "LapKunjunganJenisStatus"
                strCetak2 = IIf(optGroupBy(3).value = True, "LapKunjunganJenisStatusTotal", "LapKunjunganJenisStatusHari")
                strSQL = "Select * from V_DatakunjunganPasienMasukBjenisBstausPasien " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                "and KdInstalasi ='" & mstrKdInstalasi & "'"

            Case "LapKunjunganSt_PnyktPsn"
                strCetak2 = IIf(optGroupBy(3).value = True, "LapKunjunganSt_PnyktPsnTotal", "LapKunjunganSt_PnyktPsnHari")

                strSQL = "Select * from V_DataKunjunganPasienMasukBstatusBkasusPenyakit " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                "and KdInstalasi ='" & mstrKdInstalasi & "'"

            Case "LapKunjunganRujukanBStatus"
                strCetak2 = IIf(optGroupBy(3).value = True, "LapKunjunganRujukanBStatusTotal", "LapKunjunganRujukanBStatusHari")

                strSQL = "Select * from V_DataKunjunganPasienMasukBsetatusBRujukan " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                "and KdInstalasi ='" & mstrKdInstalasi & "'"

            Case "LapKunjunganKonPulang_Status"
                strCetak2 = IIf(optGroupBy(3).value = True, "LapKunjunganKonPulang_StatusTotal", "LapKunjunganKonPulang_StatusHari")

                strSQL = "Select * from V_DataKunjunganPasienKeluarBKondisiPulang_Bstatus " & _
                "WHERE (KdInstalasi ='" & mstrKdInstalasi & "') and ( TglKeluar BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglKeluar asc"

            Case "LapKunjunganJenisOperasi_Status"
                strCetak2 = IIf(optGroupBy(3).value = True, "LapKunjunganJenisOperasi_StatusTotal", "LapKunjunganJenisOperasi_StatusHari")

                strSQL = "Select * from V_DataKunjunganPasienMasukIBSBJenisOperasiBstatus " & _
                "WHERE (KdInstalasi ='" & mstrKdInstalasi & "') and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPendaftaran asc"

            Case "LapKunjunganKelasStatus"
                strCetak2 = IIf(optGroupBy(3).value = True, "LapKunjunganKelasStatusTotal", "LapKunjunganKelasStatusHari")

                strSQL = "Select TglPendaftaran,RuanganPelayanan,Judul,Detail,Jk,JmlPasien from V_DataKunjunganPasienMasukBsetatusBKelas " & _
                "WHERE (KdInstalasi ='" & mstrKdInstalasi & "' ) and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPendaftaran asc"

            Case "LapKunjunganBDiagnosa"
                strCetak2 = IIf(optGroupBy(3).value = True, "LapKunjunganBDiagnosaTotal", "LapKunjunganBDiagnosaHari")

                strSQL = "Select * from V_DataDiagnosaPasienPH " & _
                "WHERE (TglPeriksa BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and KdInstalasi ='" & mstrKdInstalasi & "'"

            Case "LapKunjunganPasienBDiagnosaWilayah"
                strCetak2 = IIf(optGroupBy(3).value = True, "LapKunjunganPasienBDiagnosaWilayahTotal", "LapKunjunganPasienBDiagnosaWilayahHari")

                strSQL = "Select * from V_DataDiagnosaPasienPH " & _
                "WHERE (TglPeriksa BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and KdInstalasi ='" & mstrKdInstalasi & "' "

            Case "LapKunjunganBwilayah"
                strCetak2 = IIf(optGroupBy(3).value = True, "LapKunjunganBwilayahTotal", "LapKunjunganBwilayahHari")
                strSQL = "Select TglPendaftaran,RuanganPelayanan,Judul,Detail,Jk,JmlPasien from V_DataKunjunganPasienMasukBWilayah " & _
                "WHERE  KdInstalasi ='" & mstrKdInstalasi & "' and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPendaftaran asc"

        End Select

    ElseIf optGroupBy(1).value = True Then
        mdTglAwal = dtpAwal.value 'TglAwal
        mdTglAkhir = dtpAkhir.value
        mdtBulan = CStr(Format(dtpAkhir.value, "mm"))
        MdtTahun = CStr(Format(dtpAkhir.value, "yyyy"))
        mdTglAkhir = CDate(Format(dtpAkhir.value, "yyyy-mm") & "-" & funcHitungHari(mdtBulan, MdtTahun) & " 23:59:59")
        mstrKdInstalasi = dcInstalasi.BoundText  'KdInstalasi
        mstrInstalasi2 = dcInstalasi.Text

        Select Case strCetak
            Case "LapKunjunganJenisStatus"
                strCetak2 = "LapKunjunganJenisStatusBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(TglPendaftaran)  AS TglPendaftaran, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM   V_DatakunjunganPasienMasukBjenisBstausPasien " _
                & "WHERE (month(TglPendaftaran) BETWEEN '" _
                & Month(mdTglAwal) & "' AND '" & Month(mdTglAkhir) & "' AND year (tglpendaftaran) between '" _
                & Year(mdTglAwal) & "' AND '" & Year(mdTglAkhir) & "') " & _
                "and KdInstalasi ='" & mstrKdInstalasi & "'"

            Case "LapKunjunganBwilayah"
                strCetak2 = "LapKunjunganBwilayahBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(TglPendaftaran) AS TglPendaftaran, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM  V_DataKunjunganPasienMasukBWilayah " _
                & "WHERE (TglPendaftaran BETWEEN '" _
                & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
                & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "')" & _
                "and  KdInstalasi ='" & mstrKdInstalasi & "'"

            Case "LapKunjunganSt_PnyktPsn"
                strCetak2 = "LapKunjunganSt_PnyktPsnBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(TglPendaftaran) AS TglPendaftaran, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM  V_DataKunjunganPasienMasukBstatusBkasusPenyakit " _
                & "WHERE (month(TglPendaftaran) BETWEEN '" _
                & Month(mdTglAwal) & "' AND '" & Month(mdTglAkhir) & "' AND year (tglpendaftaran) between '" _
                & Year(mdTglAwal) & "' AND '" & Year(mdTglAkhir) & "') " & _
                "and KdInstalasi ='" & mstrKdInstalasi & "'"

            Case "LapKunjunganRujukanBStatus"
                strCetak2 = "LapKunjunganRujukanBStatusBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(TglPendaftaran)  AS TglPendaftaran, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM  V_DataKunjunganPasienMasukBsetatusBRujukan " _
                & "WHERE (month(TglPendaftaran) BETWEEN '" _
                & Month(mdTglAwal) & "' AND '" & Month(mdTglAkhir) & "' AND year (tglpendaftaran) between '" _
                & Year(mdTglAwal) & "' AND '" & Year(mdTglAkhir) & "') " & _
                "and KdInstalasi ='" & mstrKdInstalasi & "' "

            Case "LapKunjunganKonPulang_Status"
                strCetak2 = "LapKunjunganKonPulang_StatusBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(TglKeluar)  AS TglKeluar, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM  V_DataKunjunganPasienKeluarBKondisiPulang_Bstatus " _
                & "WHERE (month(TglKeluar) BETWEEN '" _
                & Month(mdTglAwal) & "' AND '" & Month(mdTglAkhir) & "' AND year (tglkeluar) between '" _
                & Year(mdTglAkhir) & "' AND '" & Year(mdTglAkhir) & "')" & _
                "and KdInstalasi ='" & mstrKdInstalasi & "' "

            Case "LapKunjunganJenisOperasi_Status"
                strCetak2 = "LapKunjunganJenisOperasi_StatusBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(TglPendaftaran)  AS TglPendaftaran, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM  V_DataKunjunganPasienMasukIBSBJenisOperasiBstatus " _
                & "WHERE (month(TglPendaftaran) BETWEEN '" _
                & Month(mdTglAwal) & "' AND '" & Month(mdTglAkhir) & "' AND year (tglpendaftaran) between '" _
                & Year(mdTglAwal) & "' AND '" & Year(mdTglAkhir) & "') " & _
                "and KdInstalasi ='" & mstrKdInstalasi & "'"

            Case "LapKunjunganKelasStatus"
                strCetak2 = "LapKunjunganKelasStatusBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(TglPendaftaran)  AS TglPendaftaran, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM  V_DataKunjunganPasienMasukBsetatusBKelas " _
                & "WHERE (month(TglPendaftaran) BETWEEN '" _
                & Month(mdTglAwal) & "' AND '" & Month(mdTglAkhir) & "' AND year (tglpendaftaran) between '" _
                & Year(mdTglAwal) & "' AND '" & Year(mdTglAkhir) & "') " & _
                "And KdInstalasi ='" & mstrKdInstalasi & "'"

            Case "LapKunjunganBDiagnosa"
                strCetak2 = "LapKunjunganBDiagnosaBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(tglperiksa)  AS tglperiksa, RuanganPelayanan, KdDiagnosa,Diagnosa, StatusKasus, JK, JmlPasien  FROM  V_DataDiagnosaPasienPH " _
                & "WHERE (month(TglPeriksa) BETWEEN '" _
                & Month(mdTglAwal) & "' AND '" & Month(mdTglAkhir) & "' AND year (tglperiksa) between '" _
                & Year(mdTglAwal) & "' AND '" & Year(mdTglAkhir) & "') " & _
                "and KdInstalasi ='" & mstrKdInstalasi & "'"

            Case "LapKunjunganPasienBDiagnosaWilayah"
                strCetak2 = "LapKunjunganPasienBDiagnosaWilayahBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(tglperiksa)  AS tglperiksa,Kecamatan, StatusKasus, JK, JmlPasien  FROM  V_DataDiagnosaPasienPH " _
                & "WHERE (month(TglPeriksa) BETWEEN '" _
                & Month(mdTglAwal) & "' AND '" & Month(mdTglAkhir) & "' AND year (tglperiksa) between '" _
                & Year(mdTglAwal) & "' AND '" & Year(mdTglAkhir) & "') " & _
                "and KdInstalasi ='" & mstrKdInstalasi & "'"
        End Select

    ElseIf optGroupBy(2).value = True Then
        mdTglAwal = CDate("01-01-" & Format(dtpAwal.value, "yyyy HH:MM:SS")) 'TglAwal
        mdTglAkhir = CDate("31-12-" & Format(dtpAkhir.value, "yyyy HH:MM:SS")) 'TglAkhir
        mstrKdInstalasi = dcInstalasi.BoundText 'KdInstalasi
        mstrInstalasi2 = dcInstalasi.Text

        Select Case strCetak
            Case "LapKunjunganJenisStatus"
                strCetak2 = "LapKunjunganJenisStatusTahun"
                strSQL = "Select * from V_DatakunjunganPasienMasukBjenisBstausPasien " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                "and KdInstalasi ='" & mstrKdInstalasi & "'"

            Case "LapKunjunganBwilayah"
                strCetak2 = "LapKunjunganBwilayahTahun"
                strSQL = "Select * from V_DataKunjunganPasienMasukBWilayah " & _
                "WHERE (kdRuanganPelayanan='" & mstrKdRuangan & "') and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and KdInstalasi ='" & mstrKdInstalasi & "'order by TglPendaftaran asc"

            Case "LapKunjunganSt_PnyktPsn"
                strCetak2 = "LapKunjunganSt_PnyktPsnTahun"
                strSQL = "Select * from V_DataKunjunganPasienMasukBstatusBkasusPenyakit " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                "and KdInstalasi ='" & mstrKdInstalasi & "' "

            Case "LapKunjunganRujukanBStatus"
                strCetak2 = "LapKunjunganRujukanBStatusTahun"
                strSQL = "Select * from V_DataKunjunganPasienMasukBsetatusBRujukan " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                "and KdInstalasi ='" & mstrKdInstalasi & "'"

            Case "LapKunjunganKonPulang_Status"
                strCetak2 = "LapKunjunganKonPulang_StatusTahun"
                strSQL = "Select * from V_DataKunjunganPasienKeluarBKondisiPulang_Bstatus " & _
                "WHERE (KdInstalasi ='" & mstrKdInstalasi & "' ) and ( TglKeluar BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                "  order by tglKeluar asc"

            Case "LapKunjunganJenisOperasi_Status"
                strCetak2 = "LapKunjunganJenisOperasi_StatusTahun"
                strSQL = "Select * from V_DataKunjunganPasienMasukIBSBJenisOperasiBstatus " & _
                "WHERE ( KdInstalasi ='" & mstrKdInstalasi & "') and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " order by TglPendaftaran asc"

            Case "LapKunjunganKelasStatus"
                strCetak2 = "LapKunjunganKelasStatusTahun"
                strSQL = "Select * from V_DataKunjunganPasienMasukBsetatusBKelas " & _
                "WHERE ( KdInstalasi ='" & mstrKdInstalasi & "') and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " order by TglPendaftaran asc"

            Case "LapKunjunganBDiagnosa"
                strCetak2 = "LapKunjunganBDiagnosaTahun"
                strSQL = "Select * from V_DataDiagnosaPasienPH " & _
                "WHERE (tglperiksa BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and KdInstalasi ='" & mstrKdInstalasi & "'"

            Case "LapKunjunganPasienBDiagnosaWilayah"
                strCetak2 = "LapKunjunganPasienBDiagnosaWilayahTahun"
                strSQL = "Select * from V_DataDiagnosaPasienPH_New " & _
                "WHERE (Tglperiksa BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and KdInstalasi ='" & mstrKdInstalasi & "'"
        End Select
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    If Periksa("datacombo", dcInstalasi, "Instalasi kosong") = False Then Exit Sub
    Call Kriterialaporan
    If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then MsgBox "Tidak Ada Data", vbExclamation, "Validasi": Exit Sub
    FrmCetakLapKunjunganPasien.Show
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdgrafik_Click()
    On Error GoTo hell
    If Periksa("datacombo", dcInstalasi, "Data instalasi kosong") = False Then Exit Sub
    Call Kriterialaporan
    If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then MsgBox "Tidak Ada Data", vbExclamation, "Validasi": Exit Sub
    FrmCetakLaporandalamBentukGrafik.Show
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcInstalasi.MatchedWithList = True Then optGroupBy(0).SetFocus
        strSQL = " SELECT KdInstalasi, NamaInstalasi " & _
        " From instalasi" & _
        " WHERE (KdInstalasi IN ('01', '02', '03', '04', '06', '08', '09', '10', '11', '16')) and StatusEnabled='1' and (Namainstalasi LIKE '%" & dcInstalasi.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcInstalasi.Text = ""
            Exit Sub
        End If
        dcInstalasi.BoundText = rs(0).value
        dcInstalasi.Text = rs(1).value
    End If
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    With Me
        .dtpAwal.value = Now
        .dtpAkhir.value = Now
    End With
    Call subDcSource
    Call cekOpt
End Sub

Private Sub cekOpt()
    If optGroupBy(0).value = True Then
        Call optGroupBy_Click(0)
    ElseIf optGroupBy(1).value = True Then
        Call optGroupBy_Click(1)
    ElseIf optGroupBy(2).value = True Then
        Call optGroupBy_Click(2)
    End If
End Sub

Private Sub optGroupBy_Click(Index As Integer)
    Select Case Index
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

Private Sub optGroupBy_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub subDcSource()
    On Error GoTo hell
    strSQL = "SELECT KdInstalasi, NamaInstalasi " & _
    " From instalasi" & _
    " WHERE (KdInstalasi IN ('01', '02', '03', '04', '06', '08', '09', '10', '11', '16')) and StatusEnabled='1'"
    Call msubDcSource(dcInstalasi, rs, strSQL)

    Exit Sub
hell:
    Call msubPesanError
End Sub

