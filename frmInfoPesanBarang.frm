VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInfoPesanBarang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Informasi Pemesanan Barang"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInfoPesanBarang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   13830
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   7920
      Width           =   13815
      Begin VB.CommandButton cmdRetur 
         Caption         =   "&Retur Pemesanan"
         Height          =   520
         Left            =   4800
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdDetail 
         Caption         =   "&Detail Pemesanan"
         Height          =   520
         Left            =   6840
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdVerifikasi 
         Caption         =   "&Verifikasi"
         Height          =   520
         Left            =   8880
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   520
         Left            =   12240
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   520
         Left            =   10560
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Informasi Pemesanan Barang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   13815
      Begin VB.Frame fraDetail 
         Caption         =   "Detail Pemesanan Barang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   600
         TabIndex        =   17
         Top             =   2280
         Visible         =   0   'False
         Width           =   12855
         Begin VB.TextBox TxtIsi 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.CommandButton cmdSimpanKonfirmasi 
            Caption         =   "Simpan Konfirmasi"
            Height          =   520
            Left            =   9120
            TabIndex        =   19
            Top             =   3720
            Width           =   1935
         End
         Begin VB.CommandButton cmdTutupDetail 
            Caption         =   "Tutup"
            Height          =   520
            Left            =   11160
            TabIndex        =   18
            Top             =   3720
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid dgDetail 
            Height          =   3375
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   5953
            _Version        =   393216
            FixedCols       =   0
         End
      End
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
         Left            =   7920
         TabIndex        =   11
         Top             =   240
         Width           =   5775
         Begin VB.CommandButton cmdTampilkan 
            Caption         =   "&Cari"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpTglAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   2
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   107610115
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpTglAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   3
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   107610115
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   12
            Top             =   315
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status Barang"
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
         Left            =   4440
         TabIndex        =   10
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton optVerifikasi 
            Caption         =   "Konfirmasi"
            Height          =   375
            Left            =   3360
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.OptionButton optBelum 
            Caption         =   "Belum Dikirim"
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optSudah 
            Caption         =   "Sudah Diterima"
            Height          =   375
            Left            =   1680
            TabIndex        =   1
            Top             =   240
            Width           =   1575
         End
      End
      Begin MSDataGridLib.DataGrid dgInfoPesanBrg 
         Height          =   5535
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   9763
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   15
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
      Left            =   12000
      Picture         =   "frmInfoPesanBarang.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmInfoPesanBarang.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmInfoPesanBarang.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmInfoPesanBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim noKonfirmasi As String
Dim strNomorRetur As String
Private Sub subLoadTextIsi()

        Dim i As Integer

        txtIsi.Left = dgDetail.Left
    
        
        For i = 0 To dgDetail.Col - 1
                txtIsi.Left = txtIsi.Left + dgDetail.ColWidth(i)
        Next i

        txtIsi.Visible = True
        txtIsi.Top = dgDetail.Top - 7
    
        For i = 0 To dgDetail.Row - 1
                txtIsi.Top = txtIsi.Top + dgDetail.RowHeight(i)
        Next i
    
        If dgDetail.TopRow > 1 Then
                txtIsi.Top = txtIsi.Top - ((dgDetail.TopRow - 1) * dgDetail.RowHeight(1))
        End If
    
        txtIsi.Width = dgDetail.ColWidth(dgDetail.Col)
        '    txtIsi.Height = fgData.RowHeight(fgData.Row)
    
        txtIsi.Visible = True
        txtIsi.SelStart = Len(txtIsi.Text)
        txtIsi.SetFocus
End Sub

Private Function sp_KonfirmasiDetailOrder(f_noOrder As String, f_JmlKonfirmasi As Integer, f_NamaKonfirmasi As String, f_KdBarang As String) As Boolean
On Error GoTo Errload
   sp_KonfirmasiDetailOrder = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("noOrder", adChar, adParamInput, 10, f_noOrder)
        .Parameters.Append .CreateParameter("KdBarang", adChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("JmlKonfirmasi", adInteger, adParamInput, , CInt(f_JmlKonfirmasi))
        .Parameters.Append .CreateParameter("NamaKonfirmasi", adChar, adParamInput, 150, f_NamaKonfirmasi)
        .Parameters.Append .CreateParameter("NoKonfirmasi", adChar, adParamInput, 10, noKonfirmasi)
        
    
        .ActiveConnection = dbConn
        .CommandText = "AU_KonfirmasiDetailOrder"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Error - Ada kesalahan dalam penyimpanan data struk terima, Hubungi administrator", vbCritical, "Error"
            sp_KonfirmasiDetailOrder = False
        Else
            'noKonfirmasi = .Parameters("OutKode").Value
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
Errload:
    sp_KonfirmasiDetailOrder = False
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Call msubPesanError
End Function

Private Function sp_KonfirmasiDetailKirim(f_NoKirim As String, f_JmlKonfirmasi As Integer, f_NamaKonfirmasi As String, f_KdBarang As String) As Boolean
On Error GoTo Errload
    sp_KonfirmasiDetailKirim = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("noKirim", adVarChar, adParamInput, 10, f_NoKirim)
        .Parameters.Append .CreateParameter("KdBarang", adChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("JmlKonfirmasi", adInteger, adParamInput, , CInt(f_JmlKonfirmasi))
        .Parameters.Append .CreateParameter("NamaKonfirmasi", adVarChar, adParamInput, 150, f_NamaKonfirmasi)
        .Parameters.Append .CreateParameter("NoKonfirmasi", adVarChar, adParamInput, 10, noKonfirmasi)
        
        
    
        .ActiveConnection = dbConn
        .CommandText = "AU_KonfirmasiDetailKirim"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Error - Ada kesalahan dalam penyimpanan data struk terima, Hubungi administrator", vbCritical, "Error"
            sp_KonfirmasiDetailKirim = False
        Else
            'noKonfirmasi = .Parameters("OutKode").Value
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
Errload:
    sp_KonfirmasiDetailKirim = False
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Call msubPesanError
End Function


Private Function sp_Konfirmasi() As Boolean
On Error GoTo Errload
    sp_Konfirmasi = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoKonfirmasi", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("TglKonfirmasi", adDate, adParamInput, , Format(Now, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("NamaKonfirmasi", adVarChar, adParamInput, 150, Null)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutKode", adChar, adParamOutput, 10, Null)
    
        .ActiveConnection = dbConn
        .CommandText = "Add_Konfirmasi"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Error - Ada kesalahan dalam penyimpanan data struk terima, Hubungi administrator", vbCritical, "Error"
            sp_Konfirmasi = False
        Else
            noKonfirmasi = .Parameters("OutKode").Value
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
Errload:
    sp_Konfirmasi = False
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Call msubPesanError
End Function

Private Sub subSetGridDetail()

        On Error GoTo Errload

        With dgDetail
                .Visible = True
                .Clear
                .Rows = 2
                .Cols = 9
        
                .RowHeight(0) = 400
                .TextMatrix(0, 0) = "Tanggal"
                .TextMatrix(0, 1) = "No Pesan"
                .TextMatrix(0, 2) = "Jenis Barang"
                .TextMatrix(0, 3) = "Nama Barang"
                .TextMatrix(0, 4) = "Jml Pesan"
                .TextMatrix(0, 5) = "Tanggal Terima"
                .TextMatrix(0, 6) = "Jml Konfirmasi"
                .TextMatrix(0, 7) = "Nama Konfirmasi"
                .TextMatrix(0, 8) = "Kd Barang"
                

                .ColWidth(0) = 2200
                .ColWidth(1) = 0
                .ColWidth(2) = 900
                .ColWidth(3) = 2300
                .ColWidth(4) = 1000
                .ColWidth(5) = 2200
                .ColWidth(6) = 1300 '0
                .ColWidth(7) = 2000
                .ColWidth(8) = 0
                .ColAlignment(0) = flexAlignCenterCenter
                .ColAlignment(1) = flexAlignCenterCenter
                .ColAlignment(2) = flexAlignCenterCenter
                .ColAlignment(3) = flexAlignCenterCenter
                .ColAlignment(4) = flexAlignCenterCenter
                .ColAlignment(5) = flexAlignCenterCenter
                .ColAlignment(6) = flexAlignCenterCenter
                .ColAlignment(7) = flexAlignLeftCenter
                .ColAlignment(8) = flexAlignCenterCenter
                
        End With
        
        Exit Sub

Errload:
        Call msubPesanError
End Sub
Private Sub cmdCetak_Click()
On Error GoTo hell
If dgInfoPesanBrg.ApproxCount = 0 Then Exit Sub
'On Error Resume Next
    mdTglAwal = dtpTglAwal.Value
    mdTglAkhir = dtpTglAkhir.Value
    If optBelum.Value = True Then
        strCetak = "Belum Diterima"
    ElseIf optSudah.Value = True Then
        strCetak = "Sudah Diterima"
    ElseIf optVerifikasi.Value = True Then
        strCetak = "Sudah Diverifikasi"
    End If
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
       Set rsx = Nothing
        If optBelum.Value = True Then
        strSQLx = "select distinct [Tgl. Pesan] as Tanggal, [No. Pesan],Tujuan, [Jenis Barang],[Nama Barang], [Jml. Pesan],[Nama Pemesan],NoKirim,Noterima " & _
                 " from V_InfoPemesananBrgRuangan " & _
                 " where ([Tgl. Pesan] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "') and NoKirim is null and NoKonfirmasi is null and KdRuangan='" & mstrKdRuangan & " '"
          strjudul = "Tanggal Pesan"
          strjudulRuangan = "Tujuan"
         ElseIf optSudah.Value = True Then

        strSQLx = "select  distinct [Tgl. Kirim] as Tanggal, NoOrder as [No. Pesan], [Ruangan Pengirim] as Tujuan,[Jenis Barang],[Nama Barang], [Jml. Kirim] as [Jml. Pesan],[Nama Pemesan],NoKirim " & _
                 "from V_InfoPengirimanBrgRuangan " & _
                 "where   ([Tgl. Kirim] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "') and NoKirim is not null and NoKonfirmasi is not null AND KdRuanganTujuan='" & mstrKdRuangan & "' AND NoRetur is null "
          strjudul = "Tanggal Kirim"
          strjudulRuangan = "Pengirim"
          ElseIf optVerifikasi.Value = True Then
        
         strSQLx = "select [Tgl. Pesan] AS Tanggal, NoKonfirmasi ,[No. Pesan], Tujuan,[Jenis Barang], [Nama Barang], jmlBarangKonfirmasi as [Jml. Pesan] ,[Nama Pemesan],NoKirim " & _
                 "from V_InfoPemesananBrgRuanganKirimx " & _
                 "where  KdRuangan='" & mstrKdRuangan & "' and([Tgl. Pesan] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "') and  NoKonfirmasi IS NOT NULL "

          strjudul = "Tanggal Terima"
          strjudulRuangan = "Pengirim"
        
    End If
    
         rsx.Open strSQLx, dbConn, adOpenDynamic, adLockOptimistic
    frmCetakInfoPesanBarang.Show
hell:
End Sub

Private Sub cmdDetail_Click()
    If dgInfoPesanBrg.ApproxCount = 0 Then Exit Sub
    fraDetail.Visible = True
    Call SubLoadDetail
End Sub
Private Sub SubLoadDetail()
On Error GoTo Errload
If cmdSimpanKonfirmasi.Enabled = False Then
    cmdSimpanKonfirmasi.Enabled = True
Else
    cmdSimpanKonfirmasi = False
End If
    If (optBelum.Value = True) Then
        cmdSimpanKonfirmasi.Visible = IIf(dgInfoPesanBrg.Columns("kdRuanganTujuan").Value = mstrKdRuangan, True, False)
    ElseIf (optSudah.Value = True) Then
        cmdSimpanKonfirmasi.Visible = IIf(dgInfoPesanBrg.Columns("kdRuangan").Value = mstrKdRuangan, True, False)
    End If
    
    Set rsx = Nothing
    If optBelum.Value = True Then
        strSQLx = "select distinct [Tgl. Pesan] as Tanggal, [No. Pesan], [Jenis Barang],KdBarang,[Nama Barang], [Jml. Pesan],JmlKonfirmasi,NamaKonfirmasi,TglTerima" & _
                 " from V_InfoPemesananBrgRuangan " & _
                 " where [No. Pesan] = '" & dgInfoPesanBrg.Columns("No. Pesan") & "' and ([Tgl. Pesan] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "') "
    ElseIf optSudah.Value = True Then

        strSQLx = "select  [Tgl. Pesan] as Tanggal,NoKirim as [No. Pesan],KdBarang, [Jenis Barang],[Nama Barang], [Jml. Pesan] as [Jml. Pesan],JmlBarangKonfirmasi,NamaKonfirmasi,TglTerima " & _
                 "from V_InfoPemesananBrgRuanganKirimx " & _
                 "where NoKirim = '" & dgInfoPesanBrg.Columns("NoKirim") & "' " '  and keterangan = 'Kirim'"

    ElseIf optVerifikasi.Value = True Then

     strSQLx = "select distinct [Tgl. Pesan] AS Tanggal, NoKonfirmasi as [No. Konfirmasi], [Jenis Barang], [Nama Barang], JmlBarangKonfirmasi as [Jml. Konfirmasi], NoTerima " & _
                 "from V_InfoPemesananBrgRuanganKirimx " & _
                 "where NoKonfirmasi = '" & dgInfoPesanBrg.Columns("NoKonfirmasi") & "' and ([Tgl. Pesan] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "')"


        
    End If
    
         rsx.Open strSQLx, dbConn, adOpenDynamic, adLockOptimistic
         subSetGridDetail
        ' Set dgDetail.DataSource = rsx
    Dim i As Integer
    For i = 1 To rsx.RecordCount
         With dgDetail
                .TextMatrix(i, 0) = rsx("Tanggal")
                .TextMatrix(i, 1) = rsx("No. Pesan")
                .TextMatrix(i, 2) = rsx("Jenis Barang")
                .TextMatrix(i, 3) = rsx("Nama Barang")
                .TextMatrix(i, 4) = rsx("Jml. Pesan")
                .TextMatrix(i, 5) = IIf(IsNull(rsx("TglTerima")), "", (rsx("TglTerima")))
                If optBelum.Value = True Then
                    .TextMatrix(i, 6) = IIf(IsNull(rsx("JmlKonfirmasi")), "0", rsx("JmlKonfirmasi"))
                Else
                    .TextMatrix(i, 6) = IIf(IsNull(rsx("jmlBarangKonfirmasi")), "0", rsx("jmlBarangKonfirmasi"))
                End If
                .TextMatrix(i, 7) = IIf(IsNull(rsx("NamaKonfirmasi")), "", (rsx("NamaKonfirmasi")))
                .TextMatrix(i, 8) = rsx("KdBarang")
         End With
        dgDetail.Rows = dgDetail.Rows + 1

        rsx.MoveNext
    Next i
        'Set dgDetail.DataSource = rsx
        
    'With dgDetail
     '   .Columns(4).Width = 1900
      '  .Columns(0).Width = 1900
       ' .Columns(1).Width = 1900
        '.Columns(2).Width = 1900
        '.Columns(3).Width = 2800
    'End With

Exit Sub
Errload:
    Call msubPesanError
End Sub
Private Function sp_Retur() As Boolean
    On Error GoTo Errload

    sp_Retur = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRetur", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("TglRetur", adDate, adParamInput, , Format(Now, "yyyy/MM/dd HH:mm:ss"))
        '.Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, dgInfoPesanBrg.Columns("KdRuanganTujuan"))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 50, "Retur Order Pesanan")
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputNoRetur", adChar, adParamOutput, 10, Null)
        .Parameters.Append .CreateParameter("NoKirim", adChar, adParamInput, 10, dgInfoPesanBrg.Columns("NoKirim"))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_ReturDetailOrder"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data Retur", vbCritical, "Validasi"
            sp_Retur = False
        Else
            
            strNomorRetur = .Parameters("OutputNoRetur").Value
            Call Add_HistoryLoginActivity("Add_Retur")
        End If
    End With
    Set dbcmd = Nothing
    Call deleteADOCommandParameters(dbcmd)
    Exit Function
Errload:
    sp_Retur = False
    Call msubPesanError
End Function

Private Sub cmdRetur_Click()
On Error Resume Next
Set rs = Nothing
strSQL = "Select NoRetur from DetailOrderRuanganRetur where NoKirim ='" & dgInfoPesanBrg.Columns("NoKirim") & "'"
Call msubRecFO(rs, strSQL)

If dgInfoPesanBrg.Columns("Tujuan") = "" Then Exit Sub
    If (dgInfoPesanBrg.Columns("KdRuangan") = mstrKdRuangan) Then
        If (dgInfoPesanBrg.Columns("NoKonfirmasi") = "") Then
                If rs.EOF = False Then
                    MsgBox "Barang sudah pernah di retur", vbCritical
                    Exit Sub
                Else
                If MsgBox("Yakin akan retur barang", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
                    If (sp_Retur = False) Then Exit Sub
                        MsgBox "Retur pemesanan barang berhasil", vbInformation
                End If
            End If
        Else
            MsgBox "Data sudah di verifikasi, sehingga tidak bisa di retur", vbInformation
        End If
    End If
'End If
Exit Sub
End Sub

Private Sub cmdSimpanKonfirmasi_Click()
On Error GoTo Errload
    If sp_Konfirmasi() = False Then Exit Sub
    Dim i As Integer
     For i = 1 To dgDetail.Rows - 2
        With dgDetail
            If (optBelum.Value = True) Then
                If sp_KonfirmasiDetailOrder(.TextMatrix(i, 1), .TextMatrix(i, 6), .TextMatrix(i, 7), .TextMatrix(i, 8)) = False Then Exit Sub
            ElseIf (optSudah.Value = True) Then
                If sp_KonfirmasiDetailKirim(.TextMatrix(i, 1), .TextMatrix(i, 6), .TextMatrix(i, 7), .TextMatrix(i, 8)) = False Then Exit Sub
            End If
        End With
    Next i
    MsgBox "Data Konfirmasi telah tersimpan", vbInformation
    cmdSimpanKonfirmasi.Enabled = False
Exit Sub
Errload:
    Call msubPesanError
End Sub

Public Sub cmdTampilkan_Click()
 If optBelum.Value = True Then
        cmdVerifikasi.Enabled = False
    Else
        cmdVerifikasi.Enabled = True
    End If
    If mstrKdKelompokBarang = "02" Then     'medis
        If optBelum.Value = True Then
           cmdDetail.Caption = "Detail Pemesanan"
           Set rs = Nothing
           strSQL = "select distinct [Tgl. Pesan], [No. Pesan],Tujuan,NoKirim,NoKonfirmasi,Keterangan, [Ruangan Pemesan],kdruanganTujuan,KdRuangan  " & _
                    "from InfoPesanBarang " & _
                    "where  (KdRuangan='" & mstrKdRuangan & "' or kdruanganTujuan='" & mstrKdRuangan & "') " & _
                    "and([Tgl. Pesan] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "') and NoKirim Is Null"
           cmdVerifikasi.Enabled = False
           
        ElseIf optSudah.Value = True Then
            cmdDetail.Caption = "Detail Pengiriman"
           Set rs = Nothing
           strSQL = "select distinct [No. Pesan],Tujuan,NoKirim,NoKonfirmasi,Keterangan, [Ruangan Pemesan],kdruanganTujuan,KdRuangan  " & _
                    "from InfoPesanBarangDetailYangSudakTerima " & _
                    "where NoKirim is not null and Keterangan = 'Kirim'   and( KdRuangan='" & mstrKdRuangan & "' or kdRuanganTujuan='" & mstrKdRuangan & "' )  and([Tgl. Pesan] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "')"
                        ' " and Keterangan = 'Kirim' " & _
                    cmdVerifikasi.Enabled = True
                  
            
        Else
           Set rs = Nothing
           cmdDetail.Caption = "Detail Konfirmasi"
           strSQL = "select distinct [No. Pesan],Tujuan,NoKirim,NoKonfirmasi,Keterangan, [Ruangan Pemesan],KdRuangan,kdruanganTujuan  " & _
                    "from V_InfoPemesananBrgRuanganOrderKirimxx " & _
                    "where NoKirim is not null and NoKonfirmasi is Not null and KdRuangan='" & mstrKdRuangan & "'  and Keterangan = 'Konfirmasi' " & _
                    "and([Tgl. Pesan] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "')"
           cmdVerifikasi.Enabled = False
           
        End If
    ElseIf mstrKdKelompokBarang = "01" Then     'non medis
        If optBelum.Value = True Then
           Set rs = Nothing
           strSQL = "select * from V_InfoPemesananBrgRuanganNM where NoKirim is null and KdRuangan='" & mstrKdRuangan & "' and([Tgl. Pesan] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "')"
        Else
           Set rs = Nothing
           strSQL = "select * from V_InfoPemesananBrgRuanganNM where NoKirim is not null and KdRuangan='" & mstrKdRuangan & "' and([Tgl. Pesan] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "')"
        End If
    End If
    rs.Open strSQL, dbConn, adOpenDynamic, adLockReadOnly
    Set dgInfoPesanBrg.DataSource = rs
    With dgInfoPesanBrg
    Dim i As Integer
         For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i
        .Columns(0).Width = 1900
        .Columns(1).Width = 1800
        .Columns(2).Width = 2500
        .Columns(3).Width = 0
        .Columns(4).Width = 0
        .Columns(5).Width = 1500
'        .Columns(4).Width = 0
'        .Columns(5).Width = 0
'        .Columns(6).Width = 0
'        .Columns(7).Width = 1800
'        .Columns(8).Width = 0
''        .Columns(19).Width = 1500
'        .Columns(18).Width = 1500
'        .Columns(19).Width = 1500
    End With
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdTutupDetail_Click()
    fraDetail.Visible = False
End Sub

Private Sub cmdVerifikasi_Click()
On Error Resume Next
Dim i As Integer
If (dgInfoPesanBrg.Columns("KdRuangan") <> mstrKdRuangan) Then
    MsgBox "Verifikasi harus di ruangan pemesan", vbCritical
    Exit Sub
End If

If (dgInfoPesanBrg.Columns("NoKonfirmasi") <> "") Then
    MsgBox "Data sudah pernah di verifikasi sebelumnya", vbInformation
    Exit Sub
End If

Set rs = Nothing
strSQL = "Select NoRetur from detailorderruanganretur where NoKirim ='" & dgInfoPesanBrg.Columns("NoKirim") & "'"
Call msubRecFO(rs, strSQL)
If rs.EOF = False Then
    MsgBox "Barang yang sudah diretur tidak bisa di verifikasi", vbCritical
    Exit Sub
End If

If dgInfoPesanBrg.ApproxCount = 0 Then Exit Sub
If dgInfoPesanBrg.Columns("Keterangan").Value = "Pesan" Then dgInfoPesanBrg.Bookmark = dgInfoPesanBrg.Bookmark + 1
'Me.Enabled = False
    With frmKonfirmasiPenerimaanBarang
    If (optSudah.Value = True) Then '
        strSQL = "select distinct [Tgl. Pesan], [No. Pesan], Tujuan, NoKirim, KdRuanganTujuan, NoKonfirmasi " & _
                    "from V_InfoPemesananBrgRuanganKirimx " & _
                    "where NoKirim='" & dgInfoPesanBrg.Columns(2) & "'  and KdRuangan='" & mstrKdRuangan & "' " & _
                    "and([Tgl. Pesan] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "')"
    Else
    strSQL = "select distinct [Tgl. Pesan], [No. Pesan], Tujuan, NoKirim, KdRuanganTujuan, NoKonfirmasi " & _
                    "from V_InfoPemesananBrgRuanganOrderKirimxx " & _
                    "where  NoKirim='" & dgInfoPesanBrg.Columns(2) & "' and NoKirim is not null and NoKonfirmasi is null and KdRuangan='" & mstrKdRuangan & "' " & _
                    "and([Tgl. Pesan] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "')"
    End If
        
        Call msubRecFO(rs, strSQL)

    If rs.EOF = False Then
        If (Not IsNull(rs("No. Pesan").Value)) Then
            .txtNoOrder.Text = rs("No. Pesan").Value
        End If
        .dtpTglOrder.Value = rs(0).Value ' dgInfoPesanBrg.Columns(0).Value
        .txtRuanganTujuanPemesanan.Text = dgInfoPesanBrg.Columns(2).Value
        .txtKdRuanganPengirim.Text = rs("KdRuanganTujuan").Value
        .txtNoKirim.Text = dgInfoPesanBrg.Columns(2)
        .txtRuanganPengirim.Text = rs("Tujuan").Value
        mstrNoKirim = dgInfoPesanBrg.Columns(3)
        '.txtKdRuanganPengirim.Text = rs("Tujuan").Value

        
        If rs("NoKonfirmasi").Value = "" Or IsNull(rs("NoKonfirmasi")) Then
            .txtNoKonfirmasi.Text = ""
        Else
            .txtNoKonfirmasi.Text = rs("NoKonfirmasi").Value
        End If
        If (optSudah.Value = True) Then '
         strSQL = "Select [Nama Barang], AsalBarang, JmlOrder, JmlKirim, KdBarang, KdAsal, NoTerima " & _
                 "from InfoPesanBarangDetailYangSudakTerima " & _
                 "Where NoKirim = '" & .txtNoKirim.Text & "' "
        Else
            strSQL = "Select [Nama Barang], AsalBarang, JmlOrder, JmlKirim, KdBarang, KdAsal, NoTerima " & _
                 "from V_StrukKirimRuanganCetakM " & _
                 "Where NoKirim = '" & .txtNoKirim.Text & "' "
        End If
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = True Then Exit Sub
        
       .fgData.Rows = rs.RecordCount + 1
       For i = 1 To rs.RecordCount
            .fgData.TextMatrix(i, 0) = rs(0).Value
            .fgData.TextMatrix(i, 1) = rs(1).Value
            If IsNull(rs(2)) Then
            .fgData.TextMatrix(i, 2) = 0
            Else
            .fgData.TextMatrix(i, 2) = rs(2).Value
            End If
            .fgData.TextMatrix(i, 3) = rs(3).Value
            .fgData.TextMatrix(i, 4) = rs(3).Value
            .fgData.TextMatrix(i, 5) = ""
            .fgData.TextMatrix(i, 6) = rs(4).Value
            .fgData.TextMatrix(i, 7) = rs(5).Value
            .fgData.TextMatrix(i, 8) = rs(6).Value

       rs.MoveNext
       Next i
        .Show
Else
    MsgBox "Data Kosong", vbInformation
End If
    End With
Exit Sub
'err:
'    Call msubPesanError
'    'Resume 0
End Sub

Private Sub dgDetail_KeyPress(KeyAscii As Integer)
'
 Select Case dgDetail.Col
            
                Case 6, 7 'Jumlah Konfirmasi Barang
                        'TxtIsiRacikan.MaxLength = 20
                        Call subLoadTextIsi
                        txtIsi.Text = dgDetail.TextMatrix(dgDetail.Row, dgDetail.Col)
                        'TxtIsiRacikan.Text = Chr(KeyAscii)
                        'TxtIsiRacikan.SelStart = Len(TxtIsiRacikan.Text)
                        '            fgRacikan.Rows = fgRacikan.Rows + 1
End Select
End Sub



'Private Sub dgInfoPesanBrg_Click()
''WheelHook.WheelUnHook
''        Set MyProperty = dgInfoPesanBrg
''        WheelHook.WheelHook dgInfoPesanBrg
'End Sub

Private Sub dtpTglAkhir_Change()
    dtpTglAkhir.MaxDate = Now
End Sub

Private Sub dtpTglAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       optBelum.SetFocus
    End If
End Sub

Private Sub dtpTglAwal_Change()
    dtpTglAwal.MaxDate = Now
End Sub

Private Sub dtpTglAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       dtpTglAkhir.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call openConnection
    optBelum.Value = True
    
    If optBelum.Value = True Then
        cmdVerifikasi.Enabled = False
    Else
        cmdVerifikasi.Enabled = True
    End If
        
    dtpTglAkhir.Value = Now
    dtpTglAwal.Value = Now
    Call cmdTampilkan_Click
    
End Sub

Private Sub optBelum_Click()
    cmdRetur.Enabled = False
    Call cmdTampilkan_Click
    'dgInfoPesanBrg.Columns(0).Caption = "Tgl Pesan"
End Sub

Private Sub optBelum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdTampilkan.SetFocus
    End If
End Sub

Private Sub optVerifikasi_Click()
    Call cmdTampilkan_Click
    cmdVerifikasi.Enabled = False
   ' dgInfoPesanBrg.Columns(0).Caption = "Tgl Konfirmasi"
    
End Sub

Private Sub optSudah_Click()
    cmdRetur.Enabled = True
    Call cmdTampilkan_Click
    'dgInfoPesanBrg.Columns(0).Caption = "Tgl Kirim"
End Sub

Private Sub txtIsi_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13) Then
        dgDetail.TextMatrix(dgDetail.Row, dgDetail.Col) = txtIsi.Text
        txtIsi.Visible = False
        txtIsi.Text = ""
        dgDetail.SetFocus
    End If
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Select Case dgDetail.Col
                Case 6 'Jumlah Konfirmasi
                        Call SetKeyPressToNumber(KeyAscii)
                        If dgDetail.TextMatrix(dgDetail.Row, 6) > dgDetail.TextMatrix(dgDetail.Row, 4) Then
                            MsgBox "Jumlah konfirmasi tidak bisa melebihi jumlah pesan", vbInformation
                            dgDetail.TextMatrix(dgDetail.Row, 6) = dgDetail.TextMatrix(dgDetail.Row, 4)
                        Exit Sub
                        End If
    End Select
End Sub
