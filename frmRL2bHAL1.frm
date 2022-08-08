VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRL2bHAL1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL2b "
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   Icon            =   "frmRL2bHAL1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   6240
   Begin VB.OptionButton Option21 
      Caption         =   "Hal. 17"
      Height          =   495
      Left            =   4440
      TabIndex        =   34
      Top             =   4800
      Width           =   1215
   End
   Begin VB.OptionButton Option18 
      Caption         =   "Hal. 14"
      Height          =   495
      Left            =   3120
      TabIndex        =   31
      Top             =   5760
      Width           =   1215
   End
   Begin VB.OptionButton Option20 
      Caption         =   "Hal. 16"
      Height          =   495
      Left            =   4440
      TabIndex        =   33
      Top             =   4320
      Width           =   1215
   End
   Begin VB.OptionButton Option19 
      Caption         =   "Hal. 15"
      Height          =   495
      Left            =   3120
      TabIndex        =   32
      Top             =   6240
      Width           =   1215
   End
   Begin VB.OptionButton Option17 
      Caption         =   "Hal. 13"
      Height          =   495
      Left            =   3120
      TabIndex        =   30
      Top             =   5280
      Width           =   1215
   End
   Begin VB.OptionButton Option16 
      Caption         =   "Hal. 12"
      Height          =   495
      Left            =   3120
      TabIndex        =   29
      Top             =   4800
      Width           =   1215
   End
   Begin VB.OptionButton Option15 
      Caption         =   "Hal. 11"
      Height          =   495
      Left            =   3120
      TabIndex        =   28
      Top             =   4320
      Width           =   1215
   End
   Begin VB.OptionButton Option14 
      Caption         =   "Hal. 10"
      Height          =   495
      Left            =   1560
      TabIndex        =   27
      Top             =   6240
      Width           =   1215
   End
   Begin VB.OptionButton Option13 
      Caption         =   "Hal. 9"
      Height          =   495
      Left            =   1560
      TabIndex        =   26
      Top             =   5760
      Width           =   1215
   End
   Begin VB.OptionButton Option12 
      Caption         =   "Hal. 8"
      Height          =   495
      Left            =   1560
      TabIndex        =   25
      Top             =   5280
      Width           =   1215
   End
   Begin VB.OptionButton Option11 
      Caption         =   "Hal. 7"
      Height          =   495
      Left            =   1560
      TabIndex        =   24
      Top             =   4800
      Width           =   1215
   End
   Begin VB.OptionButton Option10 
      Caption         =   "Hal. 6"
      Height          =   495
      Left            =   1560
      TabIndex        =   23
      Top             =   4320
      Width           =   1215
   End
   Begin VB.OptionButton Option9 
      Caption         =   "Hal. 5"
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   6240
      Width           =   1215
   End
   Begin VB.OptionButton Option8 
      Caption         =   "Hal. 4"
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   5760
      Width           =   1215
   End
   Begin VB.OptionButton Option7 
      Caption         =   "Hal. 3"
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   5280
      Width           =   1215
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Hal. 2"
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   4800
      Width           =   1215
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Hal. 1"
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   4320
      Width           =   1215
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
      TabIndex        =   11
      Top             =   7200
      Width           =   6285
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   3480
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   240
         Width           =   1905
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   2295
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
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
         CurrentDate     =   40544
      End
   End
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
      Height          =   705
      Left            =   3480
      TabIndex        =   7
      Top             =   3120
      Width           =   2295
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
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
         CurrentDate     =   40544
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Triwulan"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   5655
      Begin VB.OptionButton Option1 
         Caption         =   "Triwulan1"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Triwulan4"
         Height          =   495
         Left            =   4200
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Triwulan3"
         Height          =   495
         Left            =   2880
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Triwulan2"
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Triwulan"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtptahun 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   360
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
         CustomFormat    =   "yyyy"
         Format          =   127270915
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   16
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
   Begin VB.Frame Frame2 
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
      Height          =   6015
      Left            =   0
      TabIndex        =   14
      Top             =   1080
      Width           =   6255
      Begin VB.Frame Frame5 
         Caption         =   "Halaman"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   120
         TabIndex        =   17
         Top             =   2880
         Width           =   5655
      End
      Begin VB.Label Label1 
         Caption         =   "s/d"
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   2280
         Width           =   375
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL2bHAL1.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRL2bHAL1.frx":2328
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   3360
      Picture         =   "frmRL2bHAL1.frx":4CE9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2955
   End
End
Attribute VB_Name = "frmRL2bHAL1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project/reference/microsoft excel 12.0 object library
'Selalu gunakan format file excel 2003  .xls sebagai standar agar pengguna excel 2003 atau diatasnya dpt menggunakan report laporannya
'Catatan: Format excel 2000 tidak dpt mengoperasikan beberapa fungsi yg ada pada excell 2003 atau diatasnya

Option Explicit

Dim awal As String
Dim akhir As String

'Special Buat Excel
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel.Range
Dim oResizeRange As Excel.Range
Dim j As String
'Special Buat Excel

Private Sub Check1_Click()
    If Check1.value = 0 Then
        dtpAwal.Enabled = True
        dtpAkhir.Enabled = True
        dtptahun.Enabled = False
        Option1.Enabled = False
        Option2.Enabled = False
        Option3.Enabled = False
        Option4.Enabled = False
        dtpAwal.value = Now
        dtpAkhir.value = Now
        dtpAkhir.CustomFormat = "dd MMMM yyyy"
        dtpAwal.CustomFormat = "dd MMMM yyyy"
    Else
        dtpAwal.Enabled = False
        dtpAkhir.Enabled = False
        dtptahun.Enabled = True
        Option1.Enabled = True
        Option2.Enabled = True
        Option3.Enabled = True
        Option4.Enabled = True
        dtpAkhir.CustomFormat = "MMMM dd"
        dtpAwal.CustomFormat = "MMMM dd"
        dtptahun.value = Now
    End If
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell

    If Option5.value = True Then
        'Hal1
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.1.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where a.qnodtd between '482' and'516'" _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell
        End If

    ElseIf Option6.value = True Then
        'Hal2
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.2.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where a.qnodtd between '517' and'555'" _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell
        End If

    ElseIf Option7.value = True Then
        'Hal3
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.3.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where a.qnodtd between '556' and '589'" _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell
        End If

    ElseIf Option8.value = True Then
        'Hal4
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.4.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where a.qnodtd between '590' and'623'" _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell
        End If

    ElseIf Option9.value = True Then
        'Hal5
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.5.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where a.qnodtd between '624' and'655'" _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell
        End If

    ElseIf Option10.value = True Then
        'Hal6
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.6.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where a.qnodtd between '656' and'684'" _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell
        End If

    ElseIf Option11.value = True Then
        'Hal7
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.7.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where a.qnodtd between '685' and'719'" _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell
        End If

    ElseIf Option12.value = True Then
        'Hal8
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.8.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where a.qnodtd between '720' and'759'" _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell
        End If

    ElseIf Option13.value = True Then
        'Hal9
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.9.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where a.qnodtd between '760' and'792'" _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell
        End If

    ElseIf Option14.value = True Then
        'Hal10
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.10.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where a.qnodtd between '793' and'828'" _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell
        End If

    ElseIf Option15.value = True Then
        'Hal11
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.11.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where a.qnodtd between '829' and'865'" _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell
        End If

    ElseIf Option16.value = True Then
        'Hal12
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.12.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where a.qnodtd between '866' and'895'" _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell
        End If

    ElseIf Option17.value = True Then
        'Hal13
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.13.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where a.qnodtd between '896' and'931'" _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell
        End If

    ElseIf Option18.value = True Then
        'Hal14
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.14.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where a.qnodtd between '932' and'958'" _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell
        End If

    ElseIf Option19.value = True Then

        'Hal15
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.15.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where a.qnodtd between '959' and'977'" _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell

            Set rsx = Nothing
            strSQL = "select sum(Kel_Umur1)as Kel_Umur1, sum(Kel_Umur2)as Kel_Umur2, sum(Kel_Umur3)as Kel_Umur3, sum(Kel_Umur4)as Kel_Umur4, " & _
            " sum(Kel_Umur5)as Kel_Umur5, sum(Kel_Umur6)as Kel_Umur6, sum(Kel_Umur7)as Kel_Umur7, sum(Kel_Umur8)as Kel_Umur8, " & _
            " sum(Kel_L)as Kel_L, sum(Kel_P)as Kel_P, sum(Total)as Total, sum(Kel_kunj)as Kel_kunj from rl2b where bulan between '" & Format(dtpAwal, "MM ") & "' AND '" & Format(dtpAkhir, "MM") & "'" & _
            " and tahun = '" & Format(dtpAwal, "yyyy ") & "' and qnodtd between '482'and '977'"
            Call msubRecFO(rsx, strSQL)
            j = 30
            With oSheet
                .Cells(j, 9) = Trim(IIf(IsNull(rsx!Kel_Umur1.value), 0, (rsx!Kel_Umur1.value)))
                .Cells(j, 10) = Trim(IIf(IsNull(rsx!Kel_Umur2.value), 0, (rsx!Kel_Umur2.value)))
                .Cells(j, 11) = Trim(IIf(IsNull(rsx!Kel_Umur3.value), 0, (rsx!Kel_Umur3.value)))
                .Cells(j, 12) = Trim(IIf(IsNull(rsx!Kel_Umur4.value), 0, (rsx!Kel_Umur4.value)))
                .Cells(j, 13) = Trim(IIf(IsNull(rsx!Kel_Umur5.value), 0, (rsx!Kel_Umur5.value)))
                .Cells(j, 14) = Trim(IIf(IsNull(rsx!Kel_Umur6.value), 0, (rsx!Kel_Umur6.value)))
                .Cells(j, 15) = Trim(IIf(IsNull(rsx!Kel_Umur7.value), 0, (rsx!Kel_Umur7.value)))
                .Cells(j, 16) = Trim(IIf(IsNull(rsx!Kel_Umur8.value), 0, (rsx!Kel_Umur8.value)))
                .Cells(j, 17) = Trim(IIf(IsNull(rsx!Kel_L.value), 0, (rsx!Kel_L.value)))
                .Cells(j, 18) = Trim(IIf(IsNull(rsx!Kel_P.value), 0, (rsx!Kel_P.value)))
                .Cells(j, 19) = Trim(IIf(IsNull(rsx!total.value), 0, (rsx!total.value)))
                .Cells(j, 20) = Trim(IIf(IsNull(rsx!kel_Kunj.value), 0, (rsx!kel_Kunj.value)))
            End With
        End If

    ElseIf Option20.value = True Then

        'Hal16
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.16.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where (a.qnodtd between '978' and'1005') and (a.QNoDTD not In ('996','997','998','999')) " _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell
        End If

    ElseIf Option21.value = True Then

        'Hal17
        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2B Data Keadaan Morbiditas Rawat Jalan Hal.17.xls")
        Set oSheet = oWB.ActiveSheet

        If Check1.value = vbChecked And Option1.value = True Then
            oSheet.Cells(4, 9).value = "I"
        ElseIf Check1.value = vbChecked And Option2.value = True Then
            oSheet.Cells(4, 9).value = "II"
        ElseIf Check1.value = vbChecked And Option3.value = True Then
            oSheet.Cells(4, 9).value = "III"
        ElseIf Check1.value = vbChecked And Option4.value = True Then
            oSheet.Cells(4, 9).value = "IV"
        ElseIf Check1.value = vbUnchecked Then
            oSheet.Cells(4, 9).value = ""
        End If

        oSheet.Cells(4, 12).value = Format(frmRL2bHAL1.dtptahun.value, "yyyy")

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("h6", "h7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("s6", "s7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, a.NamaDTD, a.NoDTerperinci, isnull(sum(b.Kel_Umur1), 0) as Kel_Umur1, isnull(sum(b.Kel_Umur2), 0) as Kel_Umur2, " _
        & "isnull(sum(b.Kel_Umur3), 0) as Kel_Umur3, isnull(sum(b.Kel_Umur4), 0) as Kel_Umur4, isnull(sum(b.Kel_Umur5), 0) as Kel_Umur5, isnull(sum(b.Kel_Umur6), 0) as Kel_Umur6, " _
        & "isnull(sum(b.Kel_Umur7), 0) as Kel_Umur7, isnull(sum(b.Kel_Umur8), 0) as Kel_Umur8, isnull(sum(b.Kel_L), 0) as Kel_L, isnull(sum(b.Kel_P), 0) as Kel_P, isnull(sum(b.Kel_Kunj), 0) as Kel_Kunj, " _
        & "isnull(sum(b.Kel_L), 0) + isnull(sum(b.Kel_P), 0) as Total FROM DIAGNOSADTD a left outer join " _
        & "(select a.tglPeriksa, b.NoDTD, sum(a.JmlPasienKel1) as Kel_Umur1, sum(a.JmlPasienKel2) as Kel_Umur2, " _
        & "sum(a.JmlPasienKel3) as Kel_Umur3, sum(a.JmlPasienKel4) as Kel_Umur4, sum(a.JmlPasienKel5) as Kel_Umur5, sum(a.JmlPasienKel6) as Kel_Umur6, sum(a.JmlPasienKel7) as Kel_Umur7, " _
        & "sum(a.JmlPasienKel8) as Kel_Umur8, sum(JmlPasienOutPria) as Kel_L, sum(a.JmlPasienOutWanita) as Kel_P, sum(a.JmlPasienOutHidup) as Kel_H, sum(a.JmlPasienOutMati) as Kel_M, sum(a.JmlKunjungan) as Kel_Kunj, " _
        & "a.KdRuangan, d.KdInstalasi, a.NoPendaftaran from PeriksaDiagnosa a inner join Diagnosa b on a.kdDiagnosa = b.kdDiagnosa " _
        & "inner join registrasiRJ c on a.NoPendaftaran = c.NoPendaftaran inner join Ruangan d on a.kdRUangan = d.kdRUangan left outer join PasienBatalDirawat e on a.NoPendaftaran = e.NoPendaftaran " _
        & "WHERE   a.TglPeriksa BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and e.NoPendaftaran is Null " _
        & "group by a.tglPeriksa, b.NoDTD, a.KdRuangan, d.KdInstalasi, a.NoPendaftaran) as b on a.NoDTD = b.NoDTD " _
        & "where a.qnodtd in ('1006','1007','996','997','998','999')" _
        & "group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD order by a.NoDTD, a.NamaDTD"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 11
            Call setcell

            Set rsx = Nothing
            strSQL = "select sum(Kel_Umur1)as Kel_Umur1, sum(Kel_Umur2)as Kel_Umur2, sum(Kel_Umur3)as Kel_Umur3, sum(Kel_Umur4)as Kel_Umur4, " & _
            " sum(Kel_Umur5)as Kel_Umur5, sum(Kel_Umur6)as Kel_Umur6, sum(Kel_Umur7)as Kel_Umur7, sum(Kel_Umur8)as Kel_Umur8, " & _
            " sum(Kel_L)as Kel_L, sum(Kel_P)as Kel_P, sum(Total)as Total, sum(Kel_kunj)as Kel_kunj from rl2b where bulan between '" & Format(dtpAwal, "MM ") & "' AND '" & Format(dtpAkhir, "MM") & "'" & _
            " and tahun = '" & Format(dtpAwal, "yyyy ") & "'and  qnodtd between '978' and '1007'"
            Call msubRecFO(rsx, strSQL)
            j = 17
            With oSheet
                .Cells(j, 9) = Trim(IIf(IsNull(rsx!Kel_Umur1.value), 0, (rsx!Kel_Umur1.value)))
                .Cells(j, 10) = Trim(IIf(IsNull(rsx!Kel_Umur2.value), 0, (rsx!Kel_Umur2.value)))
                .Cells(j, 11) = Trim(IIf(IsNull(rsx!Kel_Umur3.value), 0, (rsx!Kel_Umur3.value)))
                .Cells(j, 12) = Trim(IIf(IsNull(rsx!Kel_Umur4.value), 0, (rsx!Kel_Umur4.value)))
                .Cells(j, 13) = Trim(IIf(IsNull(rsx!Kel_Umur5.value), 0, (rsx!Kel_Umur5.value)))
                .Cells(j, 14) = Trim(IIf(IsNull(rsx!Kel_Umur6.value), 0, (rsx!Kel_Umur6.value)))
                .Cells(j, 15) = Trim(IIf(IsNull(rsx!Kel_Umur7.value), 0, (rsx!Kel_Umur7.value)))
                .Cells(j, 16) = Trim(IIf(IsNull(rsx!Kel_Umur8.value), 0, (rsx!Kel_Umur8.value)))
                .Cells(j, 17) = Trim(IIf(IsNull(rsx!Kel_L.value), 0, (rsx!Kel_L.value)))
                .Cells(j, 18) = Trim(IIf(IsNull(rsx!Kel_P.value), 0, (rsx!Kel_P.value)))
                .Cells(j, 19) = Trim(IIf(IsNull(rsx!total.value), 0, (rsx!total.value)))
                .Cells(j, 20) = Trim(IIf(IsNull(rsx!kel_Kunj.value), 0, (rsx!kel_Kunj.value)))
            End With
        End If

    End If
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtptahun_Change()
    dtptahun.MaxDate = Now
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    With Me
        .dtpAwal.value = Now
        .dtpAkhir.value = Now
        .dtptahun.value = Now
    End With
    Check1.value = 1
    Option1.value = 1
End Sub

Private Sub Option1_Click()
    awal = CStr(dtptahun.Year) + "/01/01"
    akhir = CStr(dtptahun.Year) + "/03/31"

    dtpAwal = awal
    dtpAkhir = akhir
End Sub

Private Sub Option10_Click()
    Call clear
End Sub

Private Sub Option11_Click()
    Call clear
End Sub

Private Sub Option12_Click()
    Call clear
End Sub

Private Sub Option13_Click()
    Call clear
End Sub

Private Sub Option14_Click()
    Call clear
End Sub

Private Sub Option15_Click()
    Call clear
End Sub

Private Sub Option16_Click()
    Call clear
End Sub

Private Sub Option17_Click()
    Call clear
End Sub

Private Sub Option18_Click()
    Call clear
End Sub

Private Sub Option19_Click()
    Call clear
End Sub

Private Sub Option2_Click()
    awal = CStr(dtptahun.Year) + "/04/01"
    akhir = CStr(dtptahun.Year) + "/06/30"

    dtpAwal = awal
    dtpAkhir = akhir
End Sub

Private Sub Option20_Click()
    Call clear
End Sub

Private Sub Option21_Click()
    Call clear
End Sub

Private Sub Option3_Click()
    awal = CStr(dtptahun.Year) + "/07/01"
    akhir = CStr(dtptahun.Year) + "/09/30"

    dtpAwal = awal
    dtpAkhir = akhir
End Sub

Private Sub Option4_Click()
    awal = CStr(dtptahun.Year) + "/10/01"
    akhir = CStr(dtptahun.Year) + "/12/31"

    dtpAwal = awal
    dtpAkhir = akhir
End Sub

Private Sub clear()
    Check1.Enabled = False
    dtptahun.Enabled = False
    dtpAwal.Enabled = False
    dtpAkhir.Enabled = False
    Option1.Enabled = False
    Option2.Enabled = False
    Option3.Enabled = False
    Option4.Enabled = False
End Sub

Private Sub Option5_Click()
    Call clear
End Sub

Private Sub Option6_Click()
    Call clear
End Sub

Private Sub Option7_Click()
    Call clear
End Sub

Private Sub Option8_Click()
    Call clear
End Sub

Private Sub Option9_Click()
    Call clear
End Sub

Private Sub setcell()
    While Not rs.EOF
        With oSheet
            .Cells(j, 9) = Trim(IIf(IsNull(rs!Kel_Umur1.value), 0, (rs!Kel_Umur1.value)))
            .Cells(j, 10) = Trim(IIf(IsNull(rs!Kel_Umur2.value), 0, (rs!Kel_Umur2.value)))
            .Cells(j, 11) = Trim(IIf(IsNull(rs!Kel_Umur3.value), 0, (rs!Kel_Umur3.value)))
            .Cells(j, 12) = Trim(IIf(IsNull(rs!Kel_Umur4.value), 0, (rs!Kel_Umur4.value)))
            .Cells(j, 13) = Trim(IIf(IsNull(rs!Kel_Umur5.value), 0, (rs!Kel_Umur5.value)))
            .Cells(j, 14) = Trim(IIf(IsNull(rs!Kel_Umur6.value), 0, (rs!Kel_Umur6.value)))
            .Cells(j, 15) = Trim(IIf(IsNull(rs!Kel_Umur7.value), 0, (rs!Kel_Umur7.value)))
            .Cells(j, 16) = Trim(IIf(IsNull(rs!Kel_Umur8.value), 0, (rs!Kel_Umur8.value)))
            .Cells(j, 17) = Trim(IIf(IsNull(rs!Kel_L.value), 0, (rs!Kel_L.value)))
            .Cells(j, 20) = Trim(IIf(IsNull(rs!kel_Kunj.value), 0, (rs!kel_Kunj.value)))
        End With
        j = j + 1
        rs.MoveNext
    Wend
End Sub

