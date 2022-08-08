VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDaftarPasienRL22 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien RL 2.2"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienRL22.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   10050
   Begin VB.Frame frameJudul 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   8
      Top             =   1080
      Width           =   9975
      Begin VB.Frame Frame1 
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
         Left            =   3840
         TabIndex        =   9
         Top             =   120
         Width           =   5775
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   345
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   127270915
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   345
            Left            =   3480
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   127270915
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   10
            Top             =   307
            Width           =   255
         End
      End
      Begin VB.Label LblJumData 
         AutoSize        =   -1  'True
         Caption         =   "10 / 100 Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1155
      End
   End
   Begin MSDataGridLib.DataGrid dgPasienRL22 
      Height          =   5175
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cari Data Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   7320
      Width           =   9975
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   450
         Left            =   8040
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Nama Pasien / No. CM"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2640
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   8340
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8811
            Text            =   "Cetak Data Individual Pasien Rawat Inap (F11)"
            TextSave        =   "Cetak Data Individual Pasien Rawat Inap (F11)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8811
            Text            =   "Refresh Data (F5)"
            TextSave        =   "Refresh Data (F5)"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   13
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
      Left            =   12840
      Picture         =   "frmDaftarPasienRL22.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPasienRL22.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPasienRL22.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDaftarPasienRL22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project/reference/microsoft excel 12.0 object library
'Selalu gunakan format file excel 2003  .xls sebagai standar agar pengguna excel 2003 atau diatasnya dpt menggunakan report laporannya
'Catatan: Format excel 2000 tidak dpt mengoperasikan beberapa fungsi yg ada pada excell 2003 atau diatasnya

Option Explicit

'Special Buat Excel
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel.Range
Dim oResizeRange As Excel.Range
Dim j As String
'Special Buat Excel

Private Sub cmdCari_Click()
    On Error GoTo hell
    lblJumData.Caption = "0/0"
    Set rs = Nothing
    strSQL = "select NoCM, NamaLengkap, TglPendaftaran, TglPulang from Rl2_2 where ([NoCM] like '%" & txtParameter.Text & "%' OR [NamaLengkap] like '%" & txtParameter.Text & "%') AND (TglPENDAFTARAN between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "')"
    rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    Set dgPasienRL22.DataSource = rs
    Call SetGridPasienRL22
    lblJumData.Caption = "1 / " & dgPasienRL22.ApproxCount & " Data"
    If dgPasienRL22.ApproxCount = 0 Then dtpAwal.SetFocus Else dgPasienRL22.SetFocus
    Exit Sub
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgPasienRL22_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    lblJumData.Caption = dgPasienRL22.Bookmark & " / " & dgPasienRL22.ApproxCount & " Data"
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo hell
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    Select Case KeyCode
        Case vbKeyF5
            Call cmdCari_Click
        Case vbKeyF11
            If dgPasienRL22.ApproxCount = 0 Then Exit Sub

            'Buka Excel
            Set oXL = CreateObject("Excel.Application")
            oXL.Visible = True
            'Buat Buka Template
            Set oWB = oXL.Workbooks.Open(App.Path & "\Data Individual Morbiditas Pasien RI_RL2.2.xls")
            Set oSheet = oWB.ActiveSheet

            Set rsb = Nothing
            strSQL = "select * from profilrs"
            Call msubRecFO(rsb, strSQL)

            Set oResizeRange = oSheet.Range("f6")
            oResizeRange.value = Trim(rsb!NamaRS)

            Set oResizeRange = oSheet.Range("r5")
            oResizeRange.value = Trim(rsb!KdRs)

            mstrNoCM = dgPasienRL22.Columns("nocm")
            strSQL = "Select * from Rl2_2 where nocm= '" & mstrNoCM & "'"
            Call msubRecFO(rs, strSQL)

            With oSheet
                .Cells(6, 18) = Trim(IIf(IsNull(rs!NoCM.value), "", (rs!NoCM.value)))
                .Cells(8, 8) = Trim(IIf(IsNull(rs!TglPendaftaran.value), "", (rs!TglPendaftaran.value)))
                .Cells(9, 8) = Trim(IIf(IsNull(rs!Tglbersalin.value), "", (rs!Tglbersalin.value)))
                .Cells(10, 8) = Trim(IIf(IsNull(rs!TglPulang.value), "", (rs!TglPulang.value)))
                .Cells(11, 6) = Trim(IIf(IsNull(rs!DeskKelas.value), "", (rs!DeskKelas.value)))

                If rs!CaraMasuk.value = "Unit Gawat Darurat" Then
                    oSheet.Cells(12, 3) = "V"
                    oSheet.Cells(12, 7) = ""
                    oSheet.Cells(12, 10) = ""
                End If
                If rs!CaraMasuk.value = "Unit Rawat Jalan" Then
                    oSheet.Cells(12, 3) = ""
                    oSheet.Cells(12, 7) = "V"
                    oSheet.Cells(12, 10) = ""
                End If
                If rs!CaraMasuk.value = "Langsung Rawat Inap" Then
                    oSheet.Cells(12, 3) = ""
                    oSheet.Cells(12, 7) = ""
                    oSheet.Cells(12, 10) = "V"
                End If

                If rs!RujukanAsal.value = "RS Pemerintah" Or rs!RujukanAsal.value = "RS Swasta" Or rs!RujukanAsal.value = "Lain - Lain" Or rs!RujukanAsal.value = "Kecelakaan" Or rs!RujukanAsal.value = "IGD" Or rs!RujukanAsal.value = "Poliklinik" Or rs!RujukanAsal.value = "Rawat Inap" Or rs!RujukanAsal.value = "Intern" Or rs!RujukanAsal.value = "Praktek Swasta" Then
                    oSheet.Cells(14, 3) = "V"
                    oSheet.Cells(14, 7) = ""
                    oSheet.Cells(14, 9) = ""
                    oSheet.Cells(14, 12) = ""
                    oSheet.Cells(15, 3) = ""
                    oSheet.Cells(15, 7) = ""
                    oSheet.Cells(15, 9) = ""
                    oSheet.Cells(15, 12) = ""
                End If
                If rs!RujukanAsal.value = "Puskesmas" Then
                    oSheet.Cells(14, 3) = ""
                    oSheet.Cells(14, 7) = ""
                    oSheet.Cells(14, 9) = ""
                    oSheet.Cells(14, 12) = ""
                    oSheet.Cells(15, 3) = "V"
                    oSheet.Cells(15, 7) = ""
                    oSheet.Cells(15, 9) = ""
                    oSheet.Cells(15, 12) = ""
                End If
                If rs!RujukanAsal.value = "Kasus Polisi" Then
                    oSheet.Cells(14, 3) = ""
                    oSheet.Cells(14, 7) = ""
                    oSheet.Cells(14, 9) = ""
                    oSheet.Cells(14, 12) = "V"
                    oSheet.Cells(15, 3) = ""
                    oSheet.Cells(15, 7) = ""
                    oSheet.Cells(15, 9) = ""
                    oSheet.Cells(15, 12) = ""
                End If
                If rs!RujukanAsal.value = "Datang Sendiri" Or IsNull(rs!RujukanAsal.value) Then
                    oSheet.Cells(14, 3) = ""
                    oSheet.Cells(14, 7) = ""
                    oSheet.Cells(14, 9) = ""
                    oSheet.Cells(14, 12) = ""
                    oSheet.Cells(15, 3) = ""
                    oSheet.Cells(15, 7) = ""
                    oSheet.Cells(15, 9) = ""
                    oSheet.Cells(15, 12) = "V"
                End If

                .Cells(16, 6) = Trim(IIf(IsNull(rs!alamat.value), "", (rs!alamat.value)))
                .Cells(17, 3) = Trim(IIf(IsNull(rs!Propinsi.value), "", (rs!Propinsi.value)))
                .Cells(17, 9) = Trim(IIf(IsNull(rs!Kota.value), "", (rs!Kota.value)))
                .Cells(17, 14) = Trim(IIf(IsNull(rs!Kecamatan.value), "", (rs!Kecamatan.value)))
                .Cells(18, 6) = Trim(IIf(IsNull(rs!tgllahir.value), "", (rs!tgllahir.value)))

                If rs![TempatMelahirkan].value = "RS" Then
                    oSheet.Cells(20, 7) = "V"
                    oSheet.Cells(20, 12) = ""
                Else
                    oSheet.Cells(20, 7) = ""
                    oSheet.Cells(20, 12) = "V"
                End If

                If rs!Pendidikan.value = "Tidak Sekolah" Or rs!Pendidikan.value = "Belum Sekolah" Or IsNull(rs!Pendidikan.value) Then
                    oSheet.Cells(21, 4) = "V"
                    oSheet.Cells(21, 7) = ""
                    oSheet.Cells(21, 10) = ""
                    oSheet.Cells(22, 4) = ""
                    oSheet.Cells(22, 7) = ""
                    oSheet.Cells(22, 10) = ""
                    oSheet.Cells(23, 4) = ""
                End If
                If rs!Pendidikan.value = "TK" Then
                    oSheet.Cells(21, 4) = ""
                    oSheet.Cells(21, 7) = "V"
                    oSheet.Cells(21, 10) = ""
                    oSheet.Cells(22, 4) = ""
                    oSheet.Cells(22, 7) = ""
                    oSheet.Cells(22, 10) = ""
                    oSheet.Cells(23, 4) = ""
                End If
                If rs!Pendidikan.value = "SD" Then
                    oSheet.Cells(21, 4) = ""
                    oSheet.Cells(21, 7) = ""
                    oSheet.Cells(21, 10) = "V"
                    oSheet.Cells(22, 4) = ""
                    oSheet.Cells(22, 7) = ""
                    oSheet.Cells(22, 10) = ""
                    oSheet.Cells(23, 4) = ""
                End If
                If rs!Pendidikan.value = "SLTP" Then
                    oSheet.Cells(21, 4) = ""
                    oSheet.Cells(21, 7) = ""
                    oSheet.Cells(21, 10) = ""
                    oSheet.Cells(22, 4) = "V"
                    oSheet.Cells(22, 7) = ""
                    oSheet.Cells(22, 10) = ""
                    oSheet.Cells(23, 4) = ""
                End If
                If rs!Pendidikan.value = "SLTA" Or rs!Pendidikan.value = "SMK" Or rs!Pendidikan.value = "STM" Or rs!Pendidikan.value = "SR" Or rs!Pendidikan.value = "SPK" Then
                    oSheet.Cells(21, 4) = ""
                    oSheet.Cells(21, 7) = ""
                    oSheet.Cells(21, 10) = ""
                    oSheet.Cells(22, 4) = ""
                    oSheet.Cells(22, 7) = "V"
                    oSheet.Cells(22, 10) = ""
                    oSheet.Cells(23, 4) = ""
                End If
                If rs!Pendidikan.value = "Diploma I" Or rs!Pendidikan.value = "Diploma II" Or rs!Pendidikan.value = "Diploma III" Or rs!Pendidikan.value = "Diploma IV" Then
                    oSheet.Cells(21, 4) = ""
                    oSheet.Cells(21, 7) = ""
                    oSheet.Cells(21, 10) = ""
                    oSheet.Cells(22, 4) = ""
                    oSheet.Cells(22, 7) = ""
                    oSheet.Cells(22, 10) = "V"
                    oSheet.Cells(23, 4) = ""
                End If
                If rs!Pendidikan.value = "Dokter" Or rs!Pendidikan.value = "MAN" Or rs!Pendidikan.value = "MI" Or rs!Pendidikan.value = "MTsN" Or rs!Pendidikan.value = "Profesi" Or rs!Pendidikan.value = "Professor" Or rs!Pendidikan.value = "S1" Or rs!Pendidikan.value = "S1 Apoteker" Or rs!Pendidikan.value = "S1 Keperawatan" Or rs!Pendidikan.value = "S2" Or rs!Pendidikan.value = "S2 Farmasi/Apoteker" Or rs!Pendidikan.value = "S2 Keperawatan" Or rs!Pendidikan.value = "S2-Manjaemen Farmasi" Or rs!Pendidikan.value = "S3" Or rs!Pendidikan.value = "S3 Farmasi/Apoteker" Or rs!Pendidikan.value = "S3 Keperawatan" Then
                    oSheet.Cells(21, 4) = ""
                    oSheet.Cells(21, 7) = ""
                    oSheet.Cells(21, 10) = ""
                    oSheet.Cells(22, 4) = ""
                    oSheet.Cells(22, 7) = ""
                    oSheet.Cells(22, 10) = ""
                    oSheet.Cells(23, 4) = "V"
                End If

                .Cells(24, 6) = Trim(IIf(IsNull(rs!Pekerjaan.value), "", (rs!Pekerjaan.value)))

                If rs!AnteNatal.value = "0" Or IsNull(rs!AnteNatal.value) Then
                    oSheet.Cells(25, 5) = "V"
                    oSheet.Cells(25, 7) = ""
                    oSheet.Cells(25, 9) = ""
                End If
                If rs!AnteNatal.value >= 1 Or rs!AnteNatal.value <= 8 Then
                    oSheet.Cells(25, 5) = ""
                    oSheet.Cells(25, 7) = "V"
                    oSheet.Cells(25, 9) = ""
                End If
                If rs!AnteNatal.value > 8 Then
                    oSheet.Cells(25, 5) = ""
                    oSheet.Cells(25, 7) = ""
                    oSheet.Cells(25, 9) = "V"
                End If

                If rs!JenisPersalinan.value = "Normal" Then
                    oSheet.Cells(26, 3) = "V"
                    oSheet.Cells(26, 7) = ""
                    oSheet.Cells(26, 10) = ""
                    oSheet.Cells(27, 3) = ""
                    oSheet.Cells(27, 7) = ""
                End If
                If rs!JenisPersalinan.value = "Ekstraksi Vakum" Then
                    oSheet.Cells(26, 3) = ""
                    oSheet.Cells(26, 7) = "V"
                    oSheet.Cells(26, 10) = ""
                    oSheet.Cells(27, 3) = ""
                    oSheet.Cells(27, 7) = ""
                End If
                If rs!JenisPersalinan.value = "Ekstraksi Cunam" Then
                    oSheet.Cells(26, 3) = ""
                    oSheet.Cells(26, 7) = ""
                    oSheet.Cells(26, 10) = "V"
                    oSheet.Cells(27, 3) = ""
                    oSheet.Cells(27, 7) = ""
                End If
                If rs!JenisPersalinan.value = "Seksio Sesaria" Then
                    oSheet.Cells(26, 3) = ""
                    oSheet.Cells(26, 7) = ""
                    oSheet.Cells(26, 10) = ""
                    oSheet.Cells(27, 3) = "V"
                    oSheet.Cells(27, 7) = ""
                End If
                If rs!JenisPersalinan.value = "Lainnya" Or IsNull(rs!JenisPersalinan.value) Then
                    oSheet.Cells(26, 3) = ""
                    oSheet.Cells(26, 7) = ""
                    oSheet.Cells(26, 10) = ""
                    oSheet.Cells(27, 3) = ""
                    oSheet.Cells(27, 7) = "V"
                End If

                .Cells(28, 6) = Trim(IIf(IsNull(rs!NamaDiagnosa.value), "", (rs!NamaDiagnosa.value)))
                .Cells(30, 6) = Trim(IIf(IsNull(rs!DiagnosaTindakan.value), "", (rs!DiagnosaTindakan.value)))
                .Cells(31, 6) = Trim(IIf(IsNull(rs![PenyebabAbortus].value), "", (rs![PenyebabAbortus].value)))
                .Cells(32, 6) = Trim(IIf(IsNull(rs![MasaGestasi].value), "", (rs![MasaGestasi].value)))

                If rs!JenisPegawai.value = "Dokter" Then
                    oSheet.Cells(33, 4) = "V"
                    oSheet.Cells(33, 7) = ""
                    oSheet.Cells(33, 10) = ""
                    oSheet.Cells(33, 4) = ""
                    oSheet.Cells(33, 7) = ""
                End If
                If rs!JenisPegawai.value = "Bidan" Then
                    oSheet.Cells(33, 4) = ""
                    oSheet.Cells(33, 7) = "V"
                    oSheet.Cells(33, 10) = ""
                    oSheet.Cells(33, 4) = ""
                    oSheet.Cells(33, 7) = ""
                End If
                If rs!JenisPegawai.value = "Konservasi Gigi" Or rs!JenisPegawai.value = "Asisten Anestesi (Paramedis)" Or rs!JenisPegawai.value = "Asisten Operasi (Paramedis)" Or rs!JenisPegawai.value = "Perawat" Or rs!JenisPegawai.value = "Tenaga Penunjang" Then
                    oSheet.Cells(33, 4) = ""
                    oSheet.Cells(33, 7) = ""
                    oSheet.Cells(33, 10) = "V"
                    oSheet.Cells(33, 4) = ""
                    oSheet.Cells(33, 7) = ""
                End If
                If rs!JenisPegawai.value = "Dukun" Then
                    oSheet.Cells(33, 4) = ""
                    oSheet.Cells(33, 7) = ""
                    oSheet.Cells(33, 10) = ""
                    oSheet.Cells(33, 4) = "V"
                    oSheet.Cells(33, 7) = ""
                End If
                If rs!JenisPegawai.value = "Programmer" Or IsNull(rs!JenisPegawai.value) Or rs!JenisPegawai.value = "Sistem Administrator" Or rs!JenisPegawai.value = "Supir Ambulance" Or rs!JenisPegawai.value = "Direksi" Or rs!JenisPegawai.value = "Staff SIM" Or rs!JenisPegawai.value = "Administrasi" Or rs!JenisPegawai.value = "Operator" Or rs!JenisPegawai.value = "Pelaksana" Or rs!JenisPegawai.value = "Pelaksana Instalasi" Or rs!JenisPegawai.value = "Apoteker" Or rs!JenisPegawai.value = "Assisten Apoteker" Or rs!JenisPegawai.value = "Staff Farmasi" Or rs!JenisPegawai.value = "Psikologi" Then
                    oSheet.Cells(33, 4) = ""
                    oSheet.Cells(33, 7) = ""
                    oSheet.Cells(33, 10) = ""
                    oSheet.Cells(33, 4) = ""
                    oSheet.Cells(33, 7) = "V"
                End If

                If IsNull(rs!kdInfeksiNosokomial.value) Then
                    oSheet.Cells(35, 5) = ""
                    oSheet.Cells(35, 7) = "V"
                Else
                    oSheet.Cells(35, 5) = "V"
                    oSheet.Cells(35, 7) = ""
                End If

                If rs!PenyebabIN.value = "Staphylococcus" Then
                    oSheet.Cells(36, 4) = "V"
                    oSheet.Cells(36, 7) = ""
                    oSheet.Cells(36, 10) = ""
                    oSheet.Cells(37, 4) = ""
                    oSheet.Cells(37, 7) = ""
                    oSheet.Cells(37, 10) = ""
                    oSheet.Cells(38, 4) = ""
                    oSheet.Cells(38, 7) = ""
                    oSheet.Cells(38, 10) = ""
                End If
                If rs!PenyebabIN.value = "Streptococus" Then
                    oSheet.Cells(36, 4) = ""
                    oSheet.Cells(36, 7) = ""
                    oSheet.Cells(36, 10) = ""
                    oSheet.Cells(37, 4) = "V"
                    oSheet.Cells(37, 7) = ""
                    oSheet.Cells(37, 10) = ""
                    oSheet.Cells(38, 4) = ""
                    oSheet.Cells(38, 7) = ""
                    oSheet.Cells(38, 10) = ""
                End If
                If rs!PenyebabIN.value = "Pneumococus" Then
                    oSheet.Cells(36, 4) = ""
                    oSheet.Cells(36, 7) = ""
                    oSheet.Cells(36, 10) = ""
                    oSheet.Cells(37, 4) = ""
                    oSheet.Cells(37, 7) = ""
                    oSheet.Cells(37, 10) = ""
                    oSheet.Cells(38, 4) = "V"
                    oSheet.Cells(38, 7) = ""
                    oSheet.Cells(38, 10) = ""
                End If
                If rs!PenyebabIN.value = "E. Koli" Then
                    oSheet.Cells(36, 4) = ""
                    oSheet.Cells(36, 7) = "V"
                    oSheet.Cells(36, 10) = ""
                    oSheet.Cells(37, 4) = ""
                    oSheet.Cells(37, 7) = ""
                    oSheet.Cells(37, 10) = ""
                    oSheet.Cells(38, 4) = ""
                    oSheet.Cells(38, 7) = ""
                    oSheet.Cells(38, 10) = ""
                End If
                If rs!PenyebabIN.value = "Klebxiella" Then
                    oSheet.Cells(36, 4) = ""
                    oSheet.Cells(36, 7) = ""
                    oSheet.Cells(36, 10) = ""
                    oSheet.Cells(37, 4) = ""
                    oSheet.Cells(37, 7) = "V"
                    oSheet.Cells(37, 10) = ""
                    oSheet.Cells(38, 4) = ""
                    oSheet.Cells(38, 7) = ""
                    oSheet.Cells(38, 10) = ""
                End If
                If rs!PenyebabIN.value = "Pseudomonas" Then
                    oSheet.Cells(36, 4) = ""
                    oSheet.Cells(36, 7) = ""
                    oSheet.Cells(36, 10) = ""
                    oSheet.Cells(37, 4) = ""
                    oSheet.Cells(37, 7) = ""
                    oSheet.Cells(37, 10) = ""
                    oSheet.Cells(38, 4) = ""
                    oSheet.Cells(38, 7) = "V"
                    oSheet.Cells(38, 10) = ""
                End If
                If rs!PenyebabIN.value = "Proteus" Then
                    oSheet.Cells(36, 4) = ""
                    oSheet.Cells(36, 7) = ""
                    oSheet.Cells(36, 10) = "V"
                    oSheet.Cells(37, 4) = ""
                    oSheet.Cells(37, 7) = ""
                    oSheet.Cells(37, 10) = ""
                    oSheet.Cells(38, 4) = ""
                    oSheet.Cells(38, 7) = ""
                    oSheet.Cells(38, 10) = ""
                End If
                If rs!PenyebabIN.value = "Lain-Lain" Then
                    oSheet.Cells(36, 4) = ""
                    oSheet.Cells(36, 7) = ""
                    oSheet.Cells(36, 10) = ""
                    oSheet.Cells(37, 4) = ""
                    oSheet.Cells(37, 7) = ""
                    oSheet.Cells(37, 10) = "V"
                    oSheet.Cells(38, 4) = ""
                    oSheet.Cells(38, 7) = ""
                    oSheet.Cells(38, 10) = ""
                End If
                If rs!PenyebabIN.value = "Tidak Tahu" Or IsNull(rs!PenyebabIN.value) Then
                    oSheet.Cells(36, 4) = ""
                    oSheet.Cells(36, 7) = ""
                    oSheet.Cells(36, 10) = ""
                    oSheet.Cells(37, 4) = ""
                    oSheet.Cells(37, 7) = ""
                    oSheet.Cells(37, 10) = ""
                    oSheet.Cells(38, 4) = ""
                    oSheet.Cells(38, 7) = ""
                    oSheet.Cells(38, 10) = "V"
                End If

                .Cells(39, 8) = Trim(IIf(IsNull(rs!KetunaanKelainan.value), "", (rs!KetunaanKelainan.value)))
                .Cells(40, 9) = Trim(IIf(IsNull(rs!JmlDarah.value), "", (rs!JmlDarah.value)))
                .Cells(41, 8) = Trim(IIf(IsNull(rs![TglMelahirkanSebelumnya].value), "", (rs![TglMelahirkanSebelumnya].value)))
                .Cells(43, 8) = Trim(IIf(IsNull(rs![TglMelahirkanSekarang].value), "", (rs![TglMelahirkanSekarang].value)))

                If rs!JmlBayi.value = "1" Then
                    oSheet.Cells(44, 4) = "V"
                    oSheet.Cells(44, 7) = ""
                End If
                If rs!JmlBayi.value > 1 Then
                    oSheet.Cells(44, 4) = ""
                    oSheet.Cells(44, 7) = "V"
                End If

                If rs!KeadaanLahirBayi.value = "Lahir Hidup" Then
                    oSheet.Cells(45, 4) = "V"
                    oSheet.Cells(45, 7) = ""
                    oSheet.Cells(45, 10) = ""
                End If
                If rs!KeadaanLahirBayi.value = "Lahir Mati" Then
                    oSheet.Cells(45, 4) = ""
                    oSheet.Cells(45, 7) = "V"
                    oSheet.Cells(45, 10) = ""
                End If
                If rs!KeadaanLahirBayi.value = "Kembar Hidup & Mati" Then
                    oSheet.Cells(45, 4) = ""
                    oSheet.Cells(45, 7) = ""
                    oSheet.Cells(45, 10) = "V"
                End If

                .Cells(46, 7) = Trim(IIf(IsNull(rs![ParitasKe].value), "", (rs![ParitasKe].value)))
                .Cells(47, 6) = Trim(IIf(IsNull(rs![JmlLahirHidup].value), "", (rs![JmlLahirHidup].value)))
                .Cells(48, 6) = Trim(IIf(IsNull(rs![JmlLahirMati].value), "", (rs![JmlLahirMati].value)))
                .Cells(50, 6) = Trim(IIf(IsNull(rs![JmlAbortus].value), "", (rs![JmlAbortus].value)))

                If rs![KondisiKeluar].value = "Sembuh" Then
                    oSheet.Cells(52, 4) = "V"
                    oSheet.Cells(52, 7) = ""
                    oSheet.Cells(52, 9) = ""
                    oSheet.Cells(52, 12) = ""
                End If
                If rs![KondisiKeluar].value = "Belum Sembuh" Then
                    oSheet.Cells(52, 4) = ""
                    oSheet.Cells(52, 7) = "V"
                    oSheet.Cells(52, 9) = ""
                    oSheet.Cells(52, 12) = ""
                End If
                If rs![KondisiKeluar].value = "Meninggal < 48 Jam" Then
                    oSheet.Cells(52, 4) = ""
                    oSheet.Cells(52, 7) = ""
                    oSheet.Cells(52, 9) = "V"
                    oSheet.Cells(52, 12) = ""
                End If
                If rs![KondisiKeluar].value = "Meninggal > 48 Jam" Then
                    oSheet.Cells(52, 4) = ""
                    oSheet.Cells(52, 7) = ""
                    oSheet.Cells(52, 9) = ""
                    oSheet.Cells(52, 12) = "V"
                End If

                If rs!StatusKeluar.value = "Pulang" Then
                    oSheet.Cells(55, 3) = "V"
                    oSheet.Cells(55, 7) = ""
                    oSheet.Cells(55, 9) = ""
                    oSheet.Cells(55, 13) = ""
                End If
                If rs!StatusKeluar.value = "Dirujuk" Then
                    oSheet.Cells(55, 3) = ""
                    oSheet.Cells(55, 7) = "V"
                    oSheet.Cells(55, 9) = ""
                    oSheet.Cells(55, 13) = ""
                End If
                If rs!StatusKeluar.value = "Pulang Paksa" Then
                    oSheet.Cells(55, 3) = ""
                    oSheet.Cells(55, 7) = ""
                    oSheet.Cells(55, 9) = "V"
                    oSheet.Cells(55, 13) = ""
                End If
                If rs!StatusKeluar.value = "Pindah Kamar" Or rs!StatusKeluar.value = "Permintaan Sendiri" Or rs!StatusKeluar.value = "Mati < 48 Jam" Or rs!StatusKeluar.value = "Mati > 48 Jam" Or rs!StatusKeluar.value = "Melarikan Diri" Or rs!StatusKeluar.value = "Dirawat" Or rs!StatusKeluar.value = "Meninggal" Or IsNull(rs!StatusKeluar.value) Then
                    oSheet.Cells(55, 3) = ""
                    oSheet.Cells(55, 7) = ""
                    oSheet.Cells(55, 9) = ""
                    oSheet.Cells(55, 13) = "V"
                End If

                If rs![CaraPembayaran].value = "Membayar" Then
                    oSheet.Cells(56, 4) = "V"
                    oSheet.Cells(56, 7) = ""
                    oSheet.Cells(56, 10) = ""
                    oSheet.Cells(57, 4) = ""
                    oSheet.Cells(57, 7) = ""
                    oSheet.Cells(57, 10) = ""
                End If
                If rs![CaraPembayaran].value = "Askes" Then
                    oSheet.Cells(56, 4) = ""
                    oSheet.Cells(56, 7) = "V"
                    oSheet.Cells(56, 10) = ""
                    oSheet.Cells(57, 4) = ""
                    oSheet.Cells(57, 7) = ""
                    oSheet.Cells(57, 10) = ""
                End If
                If rs![CaraPembayaran].value = "Kontrak" Then
                    oSheet.Cells(56, 4) = ""
                    oSheet.Cells(56, 7) = ""
                    oSheet.Cells(56, 10) = "V"
                    oSheet.Cells(57, 4) = ""
                    oSheet.Cells(57, 7) = ""
                    oSheet.Cells(57, 10) = ""
                End If
                If rs![CaraPembayaran].value = "JPKM" Then
                    oSheet.Cells(56, 4) = ""
                    oSheet.Cells(56, 7) = ""
                    oSheet.Cells(56, 10) = ""
                    oSheet.Cells(57, 4) = "V"
                    oSheet.Cells(57, 7) = ""
                    oSheet.Cells(57, 10) = ""
                End If
                If rs![CaraPembayaran].value = "Keringanan" Then
                    oSheet.Cells(56, 4) = ""
                    oSheet.Cells(56, 7) = ""
                    oSheet.Cells(56, 10) = ""
                    oSheet.Cells(57, 4) = ""
                    oSheet.Cells(57, 7) = "V"
                    oSheet.Cells(57, 10) = ""
                End If
                If rs![CaraPembayaran].value = "Keterangan Tidak Mampu" Or rs![CaraPembayaran].value = "Kartu Sehat" Then
                    oSheet.Cells(56, 4) = ""
                    oSheet.Cells(56, 7) = ""
                    oSheet.Cells(56, 10) = ""
                    oSheet.Cells(57, 4) = ""
                    oSheet.Cells(57, 7) = ""
                    oSheet.Cells(57, 10) = "V"
                End If
            End With
    End Select
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpAwal.value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.value = Now
    frameJudul.Caption = "Daftar Pasien RL 2.2 "
    Call cmdCari_Click
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub SetGridPasienRL22()
    With dgPasienRL22
        .Columns(0).Width = 1000
        .Columns(1).Width = 1500
        .Columns(2).Width = 3000
        .Columns(3).Width = 3000
    End With
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdCari_Click
        txtParameter.SetFocus
    End If
End Sub
