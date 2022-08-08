VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDaftarPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Formulir Data Individual Kepegawaian"
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
   Icon            =   "frmDaftarPegawai.frx":0000
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
            Format          =   126943235
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
            Format          =   126943235
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
   Begin MSDataGridLib.DataGrid dgPegawai 
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
      Caption         =   "Cari Data Pegawai"
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
      Begin VB.CommandButton cmdclose 
         Caption         =   "Tutu&p"
         Height          =   450
         Index           =   0
         Left            =   8040
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan NamaPegawai / NIP"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2445
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
            Text            =   "Refresh Data (F5)"
            TextSave        =   "Refresh Data (F5)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8811
            Text            =   "Cetak Data Pegawai (F11)"
            TextSave        =   "Cetak Data Pegawai (F11)"
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
      Picture         =   "frmDaftarPegawai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPegawai.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPegawai.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDaftarPegawai"
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
Dim x As String
'Special Buat Excel

Private Sub cmdCari_Click()
    On Error GoTo hell
    lblJumData.Caption = "0/0"
    Set rs = Nothing
    strSQL = "select NamaLengkap, TglMasuk, IdPegawai, NIP from RL4a_1 where ([NIP] like '%" & txtParameter.Text & "%' OR [NamaLengkap] like '%" & txtParameter.Text & "%') AND (TglMasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "')or tglmasuk is null"
    rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    Set dgPegawai.DataSource = rs
    Call SetGridPegawai
    lblJumData.Caption = "1 / " & dgPegawai.ApproxCount & " Data"
    If dgPegawai.ApproxCount = 0 Then dtpAwal.SetFocus Else dgPegawai.SetFocus
    Exit Sub
hell:
End Sub

Private Sub cmdclose_Click(Index As Integer)
    Unload Me
End Sub

Private Sub dgPegawai_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    lblJumData.Caption = dgPegawai.Bookmark & " / " & dgPegawai.ApproxCount & " Data"
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
            If dgPegawai.ApproxCount = 0 Then Exit Sub
            If dgPegawai.Columns("IdPegawai") = "" Then MsgBox "Tidak Memiliki IdPegawai", vbInformation, "Information": Exit Sub
            'Buka Excel
            Set oXL = CreateObject("Excel.Application")
            oXL.Visible = True
            'Buat Buka Template
            Set oWB = oXL.Workbooks.Open(App.Path & "\Formulir Data Individual Kepegawaian RL4a.xls")
            Set oSheet = oWB.ActiveSheet

            Set rsb = Nothing
            strSQL = "select * from profilrs"
            Call msubRecFO(rsb, strSQL)

            Set oResizeRange = oSheet.Range("s11")
            oResizeRange.value = Trim(rsb!NamaRS)

            Set oResizeRange = oSheet.Range("s13")
            oResizeRange.value = Trim(rsb!KdRs)

            mstrIdPegawai = dgPegawai.Columns("IdPegawai")
            strSQL = "Select * from RL4a_1 where IdPegawai= '" & mstrIdPegawai & "'"
            Call msubRecFO(rs, strSQL)

            With oSheet
                .Cells(19, 13) = Trim(IIf(IsNull(rs!NIP.value), "", (rs!NIP.value)))
                .Cells(20, 13) = Trim(IIf(IsNull(rs!NamaLengkap.value), "", (rs!NamaLengkap.value)))

                If rs!JenisKelamin.value = "L" Then
                    oSheet.Cells(21, 13) = "V"
                    oSheet.Cells(21, 17) = ""
                End If
                If rs!JenisKelamin.value = "P" Then
                    oSheet.Cells(21, 13) = ""
                    oSheet.Cells(21, 17) = "V"
                End If

                .Cells(22, 13) = Trim(IIf(IsNull(rs!TempatLahir.value), "", (rs!TempatLahir.value)))
                .Cells(22, 20) = Trim(IIf(IsNull(rs!tgllahir.value), "", (rs!tgllahir.value)))

                If rs!Agama.value = "Islam" Then
                    oSheet.Cells(24, 13) = "1"
                End If
                If rs!Agama.value = "Kristen Protestan" Then
                    oSheet.Cells(24, 13) = "2"
                End If
                If rs!Agama.value = "Kristen Katolik" Then
                    oSheet.Cells(24, 13) = "3"
                End If
                If rs!Agama.value = "Hindu" Then
                    oSheet.Cells(24, 13) = "4"
                End If
                If rs!Agama.value = "Budha" Then
                    oSheet.Cells(24, 13) = "5"
                End If

                If rs!StatusPerkawinan.value = "Kawin" Or rs!StatusPerkawinan.value = "Menikah" Then
                    oSheet.Cells(26, 13) = "1"
                End If
                If rs!StatusPerkawinan.value = "Lajang" Or rs!StatusPerkawinan.value = "Belum" Or IsNull(rs!StatusPerkawinan.value) Then
                    oSheet.Cells(26, 13) = "2"
                End If
                If rs!StatusPerkawinan.value = "Janda" Then
                    oSheet.Cells(26, 13) = "3"
                End If
                If rs!StatusPerkawinan.value = "Duda" Then
                    oSheet.Cells(26, 13) = "4"
                End If

                .Cells(30, 13) = Trim(IIf(IsNull(rs![NamaSuamiIstri].value), "", (rs![NamaSuamiIstri].value)))
                .Cells(31, 13) = Trim(IIf(IsNull(rs![TempatLahirSuamiIstri].value), "", (rs![TempatLahirSuamiIstri].value)))
                .Cells(31, 20) = Trim(IIf(IsNull(rs![TglLahirSuamiIstri].value), "", (rs![TglLahirSuamiIstri].value)))
                .Cells(33, 13) = Trim(IIf(IsNull(rs![TglNikah].value), "", (rs![TglNikah].value)))
                .Cells(34, 13) = Trim(IIf(IsNull(rs![Pekerjaan].value), "", (rs![Pekerjaan].value)))
                .Cells(35, 13) = Trim(IIf(IsNull(rs![NoSeriSuamiIstri].value), "", (rs![NoSeriSuamiIstri].value)))
                .Cells(42, 5) = Trim(IIf(IsNull(rs![NamaAnak].value), "", (rs![NamaAnak].value)))
                .Cells(42, 14) = Trim(IIf(IsNull(rs![TglLahirAnak].value), "", (rs![TglLahirAnak].value)))
                .Cells(42, 22) = Trim(IIf(IsNull(rs![jk].value), "", (rs![jk].value)))
                .Cells(46, 14) = Trim(IIf(IsNull(rs!TglMasuk.value), "", (rs!TglMasuk.value)))
                .Cells(47, 14) = Trim(IIf(IsNull(rs!KualifikasiJurusan.value), "", (rs!KualifikasiJurusan.value)))

                If rs!TypePegawai.value = "CPNS" Then
                    oSheet.Cells(49, 13) = "1"
                End If
                If rs!TypePegawai.value = "PNS" Then
                    oSheet.Cells(49, 13) = "2"
                End If
                If rs!TypePegawai.value = "PTT" Then
                    oSheet.Cells(49, 13) = "3"
                End If

                If rs![JenisKepegawaian].value = "DEPKES" Then
                    oSheet.Cells(50, 13) = "1"
                End If
                If rs![JenisKepegawaian].value = "PNS DAERAH" Then
                    oSheet.Cells(50, 13) = "2"
                End If
                If rs![JenisKepegawaian].value = "DEPDIKNAS" Then
                    oSheet.Cells(50, 13) = "3"
                End If
                If rs![JenisKepegawaian].value = "TNI/POLRI" Then
                    oSheet.Cells(50, 13) = "4"
                End If
                If rs![JenisKepegawaian].value = "DEP.LAIN/BUMN" Then
                    oSheet.Cells(50, 13) = "5"
                End If
                If rs![JenisKepegawaian].value = "SWASTA" Then
                    oSheet.Cells(50, 13) = "6"
                End If
                If rs![JenisKepegawaian].value = "KONTRAK" Then
                    oSheet.Cells(50, 13) = "7"
                End If

                .Cells(52, 14) = Trim(IIf(IsNull(rs!idpegawai.value), "", (rs!idpegawai.value)))
                .Cells(53, 14) = Trim(IIf(IsNull(rs![GolongaAkhir].value), "", (rs![GolongaAkhir].value)))

            End With

            'Buka Excel
            Set oXL = CreateObject("Excel.Application")
            oXL.Visible = True
            'Buat Buka Template
            Set oWB = oXL.Workbooks.Open(App.Path & "\Formulir Data Individual Kepegawaian RL4b.xls")
            Set oSheet = oWB.ActiveSheet

            mstrIdPegawai = dgPegawai.Columns("IdPegawai")
            strSQL = "Select * from RL4a_2 where IdPegawai= '" & mstrIdPegawai & "'"
            Call msubRecFO(rs, strSQL)

            With oSheet
                .Cells(5, 13) = Trim(IIf(IsNull(rs!NoSK.value), "", (rs!NoSK.value)))
                .Cells(6, 22) = Trim(IIf(IsNull(rs!TglMulai.value), "", (rs!TglMulai.value)))
                .Cells(8, 22) = Trim(IIf(IsNull(rs![UnitKerja].value), "", (rs![UnitKerja].value)))
                .Cells(10, 22) = Trim(IIf(IsNull(rs![SatuanTugas].value), "", (rs![SatuanTugas].value)))
                .Cells(12, 22) = Trim(IIf(IsNull(rs!KotaKodyaKab.value), "", (rs!KotaKodyaKab.value)))
                .Cells(13, 13) = Trim(IIf(IsNull(rs!Propinsi.value), "", (rs!Propinsi.value)))
            End With

            mstrIdPegawai = dgPegawai.Columns("IdPegawai")
            strSQL = "Select * from RL4_3 where IdPegawai= '" & mstrIdPegawai & "'"
            Call msubRecFO(rsx, strSQL)

            If rsx.RecordCount > 0 Then
                rsx.MoveFirst
                j = 19
                x = 1
                While Not rsx.EOF
                    With oSheet
                        .Cells(j, 3) = x
                        .Cells(j, 4) = Trim(IIf(IsNull(rsx!NamaPangkat.value), "", (rsx!NamaPangkat.value)))
                        .Cells(j, 10) = Trim(IIf(IsNull(rsx!KdGolongan.value), "", (rsx!KdGolongan.value)))
                        .Cells(j, 13) = Trim(IIf(IsNull(rsx!TglSK.value), "", (rsx!TglSK.value)))
                        .Cells(j, 21) = Trim(IIf(IsNull(rsx!TandaTanganSK.value), "", (rsx!TandaTanganSK.value)))
                    End With
                    x = x + 1
                    j = j + 1
                    rsx.MoveNext
                Wend
            End If

            mstrIdPegawai = dgPegawai.Columns("IdPegawai")
            strSQL = "Select * from RL4a_4 where IdPegawai= '" & mstrIdPegawai & "'"
            Call msubRecFO(rsx, strSQL)

            If rsx.RecordCount > 0 Then
                rsx.MoveFirst
                j = 32
                x = 1
                While Not rsx.EOF
                    With oSheet
                        .Cells(j, 3) = x
                        .Cells(j, 4) = Trim(IIf(IsNull(rsx!NamaJabatan.value), "", (rsx!NamaJabatan.value)))
                        .Cells(j, 10) = Trim(IIf(IsNull(rsx!TglMulaiBerlaku.value), "", (rsx!TglMulaiBerlaku.value)))
                        .Cells(j, 18) = Trim(IIf(IsNull(rsx!TglAkhirBerlaku.value), "", (rsx!TglAkhirBerlaku.value)))
                        .Cells(j, 26) = Trim(IIf(IsNull(rsx!JenisJabatan.value), "", (rsx!JenisJabatan.value)))
                        .Cells(j, 28) = Trim(IIf(IsNull(rsx!NamaEselon.value), "", (rsx!NamaEselon.value)))
                    End With
                    x = x + 1
                    j = j + 1
                    rsx.MoveNext
                Wend
            End If

            mstrIdPegawai = dgPegawai.Columns("IdPegawai")
            strSQL = "Select * from RL4a_5 where IdPegawai= '" & mstrIdPegawai & "'"
            Call msubRecFO(rsx, strSQL)

            If rsx.RecordCount > 0 Then
                rsx.MoveFirst
                j = 43
                x = 1
                While Not rsx.EOF
                    With oSheet
                        .Cells(j, 3) = x
                        .Cells(j, 4) = Trim(IIf(IsNull(rsx!Pendidikan.value), "", (rsx!Pendidikan.value)))
                        .Cells(j, 10) = Trim(IIf(IsNull(rsx![ThnLulus].value), "", (rsx![ThnLulus].value)))
                        .Cells(j, 14) = Trim(IIf(IsNull(rsx!FakultasJurusan.value), "", (rsx!FakultasJurusan.value)))
                        .Cells(j, 22) = Trim(IIf(IsNull(rsx!NamaPendidikan.value), "", (rsx!NamaPendidikan.value)))
                    End With
                    x = x + 1
                    j = j + 1
                    rsx.MoveNext
                Wend
            End If

            'Buka Excel
            Set oXL = CreateObject("Excel.Application")
            oXL.Visible = True
            'Buat Buka Template
            Set oWB = oXL.Workbooks.Open(App.Path & "\Formulir Data Individual Kepegawaian RL4c.xls")
            Set oSheet = oWB.ActiveSheet

            mstrIdPegawai = dgPegawai.Columns("IdPegawai")
            strSQL = "Select * from RL4a_6a where IdPegawai= '" & mstrIdPegawai & "'"
            Call msubRecFO(rsx, strSQL)

            If rsx.RecordCount > 0 Then
                rsx.MoveFirst
                j = 8
                x = 1
                While Not rsx.EOF
                    With oSheet
                        .Cells(j, 3) = x
                        .Cells(j, 4) = Trim(IIf(IsNull(rsx!NamaTugas.value), "", (rsx!NamaTugas.value)))
                        .Cells(j, 13) = Trim(IIf(IsNull(rsx![tahun].value), "", (rsx![tahun].value)))
                        .Cells(j, 17) = Trim(IIf(IsNull(rsx![Lamanya].value), "", (rsx![Lamanya].value)))
                        .Cells(j, 21) = Trim(IIf(IsNull(rsx![Penyelenggara].value), "", (rsx![Penyelenggara].value)))
                    End With
                    x = x + 1
                    j = j + 1
                    rsx.MoveNext
                Wend
            End If

            mstrIdPegawai = dgPegawai.Columns("IdPegawai")
            strSQL = "Select * from RL4a_6b where IdPegawai= '" & mstrIdPegawai & "'"
            Call msubRecFO(rsx, strSQL)

            If rsx.RecordCount > 0 Then
                rsx.MoveFirst
                j = 20
                x = 1
                While Not rsx.EOF
                    With oSheet

                        .Cells(j, 3) = x
                        .Cells(j, 4) = Trim(IIf(IsNull(rsx!NamaPelatihan.value), "", (rsx!NamaPelatihan.value)))
                        .Cells(j, 13) = Trim(IIf(IsNull(rsx![tahun].value), "", (rsx![tahun].value)))
                        .Cells(j, 17) = Trim(IIf(IsNull(rsx!LamaWaktu.value), "", (rsx!LamaWaktu.value)))
                        .Cells(j, 21) = Trim(IIf(IsNull(rsx!InstansiPenyelenggara.value), "", (rsx!InstansiPenyelenggara.value)))
                    End With
                    x = x + 1
                    j = j + 1
                    rsx.MoveNext
                Wend
            End If

            mstrIdPegawai = dgPegawai.Columns("IdPegawai")
            strSQL = "Select * from RL4a_7 where IdPegawai= '" & mstrIdPegawai & "'"
            Call msubRecFO(rsx, strSQL)

            If rsx.RecordCount > 0 Then
                rsx.MoveFirst
                j = 34
                x = 1
                While Not rsx.EOF
                    With oSheet
                        .Cells(j, 3) = x
                        .Cells(j, 4) = Trim(IIf(IsNull(rsx!NamaPenghargaan.value), "", (rsx!NamaPenghargaan.value)))
                        .Cells(j, 13) = Trim(IIf(IsNull(rsx!TglDiperoleh.value), "", (rsx!TglDiperoleh.value)))
                        .Cells(j, 21) = Trim(IIf(IsNull(rsx!NomorPiagam.value), "", (rsx!NomorPiagam.value)))
                        .Cells(j, 25) = Trim(IIf(IsNull(rsx!InstansiPemberi.value), "", (rsx!InstansiPemberi.value)))
                    End With
                    x = x + 1
                    j = j + 1
                    rsx.MoveNext
                Wend
            End If

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
    Call cmdCari_Click
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub SetGridPegawai()
    With dgPegawai
        .Columns(0).Width = 3000
        .Columns(1).Width = 1500
        .Columns(2).Width = 2000
        .Columns(3).Width = 2000
    End With
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdCari_Click
        txtParameter.SetFocus
    End If
End Sub
