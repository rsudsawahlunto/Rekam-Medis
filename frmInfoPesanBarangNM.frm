VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInfoPesanBarangNM 
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
   Icon            =   "frmInfoPesanBarangNM.frx":0000
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
         Left            =   5280
         TabIndex        =   19
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdDetail 
         Caption         =   "&Detail"
         Height          =   520
         Left            =   7200
         TabIndex        =   16
         Top             =   240
         Width           =   1575
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
         Caption         =   "Detail"
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
         Width           =   12735
         Begin VB.CommandButton cmdSimpanKonfirmasi 
            Caption         =   "&Simpan Konfirmasi"
            Height          =   520
            Left            =   8760
            TabIndex        =   22
            Top             =   3720
            Width           =   1815
         End
         Begin VB.TextBox txtIsi 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   3840
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   480
            Visible         =   0   'False
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid dgDetail 
            Height          =   3255
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   5741
            _Version        =   393216
         End
         Begin VB.CommandButton cmdTutupDetail 
            Caption         =   "Tutup"
            Height          =   520
            Left            =   10800
            TabIndex        =   18
            Top             =   3720
            Width           =   1575
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
            Format          =   122486787
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
            Format          =   122486787
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
      Picture         =   "frmInfoPesanBarangNM.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmInfoPesanBarangNM.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmInfoPesanBarangNM.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmInfoPesanBarangNM"
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


Private Sub cmdCetak_Click()
On Error GoTo hell
If dgInfoPesanBrg.ApproxCount = 0 Then Exit Sub
'On Error Resume Next
    mdTglAwal = dtpTglAwal.value
    mdTglAkhir = dtpTglAkhir.value
    If optBelum.value = True Then
        strCetak = "Belum Diterima"
    ElseIf optSudah.value = True Then
        strCetak = "Sudah Diterima"
    ElseIf optVerifikasi.value = True Then
        strCetak = "Sudah Diverifikasi"
    End If
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakInfoPesanBarangNM.Show
hell:
End Sub

Private Sub cmdDetail_Click()
    If dgInfoPesanBrg.ApproxCount = 0 Then Exit Sub
    fraDetail.Visible = True
    Call SubLoadDetail
End Sub
Private Sub SubLoadDetail()
On Error GoTo errLoad
If cmdSimpanKonfirmasi.Enabled = False Then
    cmdSimpanKonfirmasi.Enabled = True
Else
    cmdSimpanKonfirmasi = False
End If
    If (optBelum.value = True) Then
        cmdSimpanKonfirmasi.Visible = IIf(dgInfoPesanBrg.Columns("kdRuanganTujuan").value = mstrKdRuangan, True, False)
    ElseIf (optSudah.value = True) Then
        cmdSimpanKonfirmasi.Visible = IIf(dgInfoPesanBrg.Columns("kdRuangan").value = mstrKdRuangan, True, False)
    End If

    Set rsx = Nothing
    If optBelum.value = True Then
        strSQLX = "select DISTINCT [Tgl. Pesan] as Tanggal, [No. Pesan], [Jenis Barang], KdBarang,[Nama Barang], [Jml. Pesan], JmlBarangKonfirmasi,NoRegisterAsset,NamaAsal,NamaKonfirmasi " & _
                 "from V_InfoPemesananBrgRuanganNM " & _
                 "where [No. Pesan] = '" & dgInfoPesanBrg.Columns("No. Pesan") & "' and ([Tgl. Pesan] between '" & Format(dtpTglAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.value, "yyyy/MM/dd 23:59:59") & "') "
    ElseIf optSudah.value = True Then
        strSQLX = "select DISTINCT TglKirim AS Tanggal, NoKirim as [No. Pesan],[Jenis Barang], KdBarang,[Nama Barang], [Jml. Pesan], JmlBarangKonfirmasi,NoRegisterAsset,NamaAsal,NamaKonfirmasi " & _
                 "from V_InfoPemesananBrgRuanganNMKirimx " & _
                 "where NoKirim = '" & dgInfoPesanBrg.Columns("NoKirim") & "' and (TglKirim between '" & Format(dtpTglAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.value, "yyyy/MM/dd 23:59:59") & "')"
    ElseIf optVerifikasi.value = True Then
        strSQLX = "select [Tgl. Pesan] as TglKonfirmasi,[No. Pesan],[Nama Barang], JmlBarangKonfirmasi as Jumlah,NoRegisterAsset " & _
                 "from V_InfoPemesananBrgRuanganNMKirimx " & _
                 "where NoKonfirmasi = '" & dgInfoPesanBrg.Columns("NoKonfirmasi") & "' and ([Tgl. Pesan] between '" & Format(dtpTglAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.value, "yyyy/MM/dd 23:59:59") & "') and NoKirim is not null and NoKonfirmasi is not null"
 
        
    End If
    
         rsx.Open strSQLX, dbConn, adOpenDynamic, adLockOptimistic
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
                '.TextMatrix(i, 5) = rsx("TglTerima")
                .TextMatrix(i, 5) = IIf(IsNull(rsx("JmlBarangKonfirmasi")), "0", rsx("JmlBarangKonfirmasi"))
             '   .TextMatrix(i, 6) = IIf(IsNull(rsx("NamaKonfirmasi")), "", Trim$((rsx("NamaKonfirmasi"))))
                
                If (IsNull(rsx("NamaKonfirmasi"))) Then
                    .TextMatrix(i, 6) = ""
                Else
                    .TextMatrix(i, 6) = Trim$((rsx("NamaKonfirmasi")))
                End If
                '.TextMatrix(i, 7) = IIf(IsNull(rsx("NamaKonfirmasi")), "0", (rsx("NamaKonfirmasi")))
                .TextMatrix(i, 7) = rsx("KdBarang")
                .TextMatrix(i, 8) = IIf(IsNull(rsx("NoRegisterAsset")), "0", (rsx("NoRegisterAsset")))
                .TextMatrix(i, 9) = IIf(IsNull(rsx("NamaAsal")), "", (rsx("NamaAsal")))
         End With
        dgDetail.Rows = dgDetail.Rows + 1

        rsx.MoveNext
    Next i
    
    'Set dgDetail.DataSource = rsx
'    With dgDetail
'        .Columns(0).Width = 1900
'        .Columns(1).Width = 1500
'        .Columns(2).Width = 2800
'        .Columns(3).Width = 800
'        If optBelum.Value = False Then
'            .Columns(4).Width = 2500
'        End If
'    End With

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subSetGridDetail()

        On Error GoTo errLoad

        With dgDetail
                .Visible = True
                .Clear
                .Rows = 2
                .Cols = 10
        
                .RowHeight(0) = 400
                .TextMatrix(0, 0) = "Tanggal"
                .TextMatrix(0, 1) = "No Pesan"
                .TextMatrix(0, 2) = "Jenis Barang"
                .TextMatrix(0, 3) = "Nama Barang"
                .TextMatrix(0, 4) = "Jml Kirim"
                '.TextMatrix(0, 5) = "Tanggal Terima"
                .TextMatrix(0, 5) = "Jml Konfirmasi"
                .TextMatrix(0, 6) = "Nama Konfirmasi"
                '.TextMatrix(0, 7) = "Nama Konfirmasi"
                .TextMatrix(0, 7) = "Kd Barang"
                .TextMatrix(0, 8) = "No Register Asset"
                .TextMatrix(0, 9) = "Asal Barang"
               

                .ColWidth(0) = 2200
                .ColWidth(1) = 0
                .ColWidth(2) = 2500
                .ColWidth(3) = 3000
                .ColWidth(4) = 1000
                '.ColWidth(5) = 2200
                .ColWidth(5) = 1300 '0
                '.ColWidth(7) = 2000
                .ColWidth(6) = 2000
                .ColWidth(7) = 0
                .ColWidth(8) = 1500
                .ColWidth(9) = 150
                
                .ColAlignment(0) = flexAlignCenterCenter
                .ColAlignment(1) = flexAlignCenterCenter
                .ColAlignment(2) = flexAlignCenterCenter
                .ColAlignment(3) = flexAlignCenterCenter
                .ColAlignment(4) = flexAlignCenterCenter
                .ColAlignment(5) = flexAlignCenterCenter
                .ColAlignment(6) = flexAlignCenterCenter
'                .ColAlignment(7) = flexAlignLeftCenter
                .ColAlignment(7) = flexAlignCenterCenter
                
                
        End With
        
        Exit Sub

errLoad:
        Call msubPesanError
End Sub

Private Function sp_Retur() As Boolean
    On Error GoTo errLoad

    sp_Retur = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRetur", adChar, adParamInput, 10, strNomorRetur)
        .Parameters.Append .CreateParameter("TglRetur", adDate, adParamInput, , Format(Now, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        '.Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, dgInfoPesanBrg.Columns("KdRuanganTujuan"))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 50, "Retur Order Pesanan")
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputNoRetur", adChar, adParamOutput, 10, Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, dgInfoPesanBrg.Columns(0))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_ReturDetailOrderNM"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data Retur", vbCritical, "Validasi"
        Else
            strNomorRetur = .Parameters("OutputNoRetur").value
            Call Add_HistoryLoginActivity("Add_Retur")
        End If
    End With
    Set dbcmd = Nothing
    Call deleteADOCommandParameters(dbcmd)
    Exit Function
errLoad:
    sp_Retur = False
    Call msubPesanError
End Function

Private Function sp_Konfirmasi() As Boolean
On Error GoTo errLoad
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
        
        If .Parameters("return_value").value <> 0 Then
            MsgBox "Error - Ada kesalahan dalam penyimpanan data struk terima, Hubungi administrator", vbCritical, "Error"
            sp_Konfirmasi = False
        Else
            noKonfirmasi = .Parameters("OutKode").value
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
errLoad:
    sp_Konfirmasi = False
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Call msubPesanError
End Function


Private Function sp_KonfirmasiDetailKirimNM(f_NoKirim As String, f_JmlKonfirmasi As Integer, f_NamaKonfirmasi As String, f_KdBarang As String, f_NoRegisterAsset) As Boolean
On Error GoTo errLoad
    sp_KonfirmasiDetailKirimNM = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("noKirim", adChar, adParamInput, 10, f_NoKirim)
        .Parameters.Append .CreateParameter("KdBarang", adChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("JmlKonfirmasi", adInteger, adParamInput, , CInt(f_JmlKonfirmasi))
        .Parameters.Append .CreateParameter("NamaKonfirmasi", adVarChar, adParamInput, 150, f_NamaKonfirmasi)
        .Parameters.Append .CreateParameter("NoKonfirmasi", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("NoRegisterAsset", adVarChar, adParamInput, 15, f_NoRegisterAsset)
        
    
        .ActiveConnection = dbConn
        .CommandText = "AU_KonfirmasiDetailKirimNM"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").value <> 0 Then
            MsgBox "Error - Ada kesalahan dalam penyimpanan data struk terima, Hubungi administrator", vbCritical, "Error"
            sp_KonfirmasiDetailKirimNM = False
        Else
            'noKonfirmasi = .Parameters("OutKode").Value
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
errLoad:
    sp_KonfirmasiDetailKirimNM = False
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Call msubPesanError
End Function

Private Function sp_KonfirmasiDetailOrder(f_noOrder As String, f_JmlKonfirmasi As Integer, f_NamaKonfirmasi As String, f_KdBarang As String) As Boolean
On Error GoTo errLoad
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
        .CommandText = "AU_KonfirmasiDetailOrderNM"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").value <> 0 Then
            MsgBox "Error - Ada kesalahan dalam penyimpanan data struk terima, Hubungi administrator", vbCritical, "Error"
            sp_KonfirmasiDetailOrder = False
        Else
            'noKonfirmasi = .Parameters("OutKode").Value
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
errLoad:
    sp_KonfirmasiDetailOrder = False
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Call msubPesanError
End Function

Private Sub cmdRetur_Click()
    If optBelum.value = True Then
        If dgInfoPesanBrg.ApproxCount = 0 Then Exit Sub
        If MsgBox("Yakin akan retur barang", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
'            If (dgInfoPesanBrg.Columns("NoKonfirmasi") = "") Then
                If (sp_Retur = True) Then
                    MsgBox "Data sudah Di retur"
                End If
 '           Else
  '              MsgBox "Data sudah di terima, sehingga tidak bisa di retur"
   '         End If
        End If
    End If
End Sub

Private Sub cmdSimpanKonfirmasi_Click()
On Error GoTo errLoad
    If sp_Konfirmasi() = False Then Exit Sub
    Dim i As Integer
     For i = 1 To dgDetail.Rows - 2
        With dgDetail
            If (optBelum.value = True) Then
                If sp_KonfirmasiDetailOrder(.TextMatrix(i, 1), .TextMatrix(i, 5), .TextMatrix(i, 6), .TextMatrix(i, 7)) = False Then Exit Sub
                'If sp_KonfirmasiDetailKirimNM(.TextMatrix(i, 1), .TextMatrix(i, 5), .TextMatrix(i, 6), .TextMatrix(i, 7), .TextMatrix(i, 8)) = False Then Exit Sub
            ElseIf (optSudah.value = True) Then
                If sp_KonfirmasiDetailKirimNM(.TextMatrix(i, 1), .TextMatrix(i, 5), .TextMatrix(i, 6), .TextMatrix(i, 7), .TextMatrix(i, 8)) = False Then Exit Sub
            End If
        End With
    Next i
    MsgBox "Data Konfirmasi telah tersimpan", vbInformation
    cmdSimpanKonfirmasi.Enabled = False
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Public Sub cmdTampilkan_Click()
    If mstrKdKelompokBarang = "02" Then     'medis
        If optBelum.value = True Then
           Set rs = Nothing
           strSQL = "select distinct * " & _
                    "from V_InfoPemesananBrgRuanganOrderKirimxx " & _
                    "where NoKirim is null and KdRuangan='" & mstrKdRuangan & "' " & _
                    "and([Tgl. Pesan] between '" & Format(dtpTglAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.value, "yyyy/MM/dd 23:59:59") & "')"
           cmdVerifikasi.Enabled = False
           
        ElseIf optSudah.value = True Then
           Set rs = Nothing
           strSQL = "select distinct * " & _
                    "from V_InfoPemesananBrgRuanganOrderKirimxx " & _
                    "where NoKirim is not null and NoKonfirmasi is null and KdRuangan='" & mstrKdRuangan & "' " & _
                    "and([Tgl. Pesan] between '" & Format(dtpTglAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.value, "yyyy/MM/dd 23:59:59") & "')"
 
                    cmdVerifikasi.Enabled = True
                  
            
        Else
           Set rs = Nothing
           strSQL = "select distinct * " & _
                    "from V_InfoPemesananBrgRuanganOrderKirimxx " & _
                    "where NoKirim is not null and NoKonfirmasi is Not null and KdRuangan='" & mstrKdRuangan & "' " & _
                    "and([Tgl. Pesan] between '" & Format(dtpTglAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.value, "yyyy/MM/dd 23:59:59") & "')"
           cmdVerifikasi.Enabled = False
           
        End If
    ElseIf mstrKdKelompokBarang = "01" Then     'non medis
        If optBelum.value = True Then
           Set rs = Nothing
           strSQL = "select distinct [No. Pesan], Tujuan, NoKirim, NoKonfirmasi,[Ruangan Pemesan],KdRuanganTujuan,KdRuangan from V_InfoPemesananBrgRuanganNMOrderKirimxx where NoKirim is null and (KdRuangan='" & mstrKdRuangan & "' or KdRuanganTujuan='" & mstrKdRuangan & "') and([Tgl. Pesan] between '" & Format(dtpTglAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.value, "yyyy/MM/dd 23:59:59") & "')"
           cmdVerifikasi.Enabled = False
        ElseIf optSudah.value = True Then
           Set rs = Nothing
           strSQL = "select distinct [No. Pesan], Tujuan as Pengirim, NoKirim, NoKonfirmasi,[Ruangan Pemesan],KdRuanganTujuan,KdRuangan from V_InfoPemesananBrgRuanganNMOrderKirimxx where NoKirim is Not null  and (KdRuangan='" & mstrKdRuangan & "' or KdRuanganTujuan='" & mstrKdRuangan & "') and(TglKirim between '" & Format(dtpTglAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.value, "yyyy/MM/dd 23:59:59") & "')"
           cmdVerifikasi.Enabled = True
        Else
            Set rs = Nothing
           strSQL = "select distinct [No. Pesan], Tujuan as Pengirim, NoKirim, NoKonfirmasi,[Ruangan Pemesan],KdRuanganTujuan,KdRuangan from V_InfoPemesananBrgRuanganNMOrderKirimxx where NoKirim is not null and NoKonfirmasi is not null and KdRuangan='" & mstrKdRuangan & "' and([Tgl. Pesan] between '" & Format(dtpTglAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.value, "yyyy/MM/dd 23:59:59") & "')"
           cmdVerifikasi.Enabled = False
        
        End If
    End If
    rs.Open strSQL, dbConn, adOpenDynamic, adLockOptimistic
    Set dgInfoPesanBrg.DataSource = rs
    With dgInfoPesanBrg
    Dim i As Integer
         For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i
        .Columns(0).Width = 1500
        .Columns(1).Width = 2000
        .Columns(2).Width = 1500
        .Columns(3).Width = 1500
        .Columns(4).Width = 1500
'        .Columns(5).Width = 0
'        .Columns(6).Width = 0
'        .Columns(7).Width = 1800
'        .Columns(8).Width = 0
''        .Columns(19).Width = 1500
'        .Columns(18).Width = 1500
    End With
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdTutupDetail_Click()
    fraDetail.Visible = False
End Sub

Private Sub cmdVerifikasi_Click()
Dim i As Integer

If dgInfoPesanBrg.ApproxCount = 0 Then Exit Sub

If (dgInfoPesanBrg.Columns("KdRuangan") <> mstrKdRuangan) Then
    MsgBox "Verifikasi harus di ruangan pemesan"
    Exit Sub
End If

If (dgInfoPesanBrg.Columns("KdRuangan") = "") Then
    'MsgBox "Verifikasi harus di ruangan pemesan"
    Exit Sub
End If

'If (dgInfoPesanBrg.Columns("NoKonfirmasi") <> "") Then
'    MsgBox "Data sudah di verifikasi"
'    Exit Sub
'End If

If dgInfoPesanBrg.ApproxCount = 0 Then Exit Sub
Me.Enabled = False
mstrKdKelompokBarang = "01"
    With frmKonfirmasiPenerimaanBarangNM
    .txtNamaFormPengirim.Text = Me.Name
        strSQL = "select distinct [Tgl. Pesan], [No. Pesan], Tujuan, NoKirim, KdRuanganTujuan, NoKonfirmasi, NoTerima " & _
                    "from V_InfoPemesananBrgRuanganNMOrderKirimxx " & _
                    "where NoKirim is not null and NoKonfirmasi is null and KdRuangan='" & mstrKdRuangan & "' " & _
                    "and(TglKirim between '" & Format(dtpTglAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.value, "yyyy/MM/dd 23:59:59") & "')"
        Call msubRecFO(rs, strSQL)
'
    If rs.EOF = True Then Exit Sub
        
        If IsNull(rs(1)) Then
            .txtNoOrder = ""
            Else
        .txtNoOrder.Text = rs(1)
        End If
        If IsNull(rs(0)) Then
        .dtpTglOrder.value = Now
        Else
        .dtpTglOrder.value = rs(0)
        End If
        
        .txtRuanganTujuanPemesanan.Text = rs(4)
        
        If IsNull(rs(3)) Then
        .txtNoKirim.Text = ""
        mstrNoKirim = ""
        Else
        .txtNoKirim.Text = dgInfoPesanBrg.Columns(2).value
        mstrNoKirim = rs(3)
        End If
        .txtRuanganPengirim.Text = rs(2)
        
        .txtKdRuanganPengirim.Text = rs(4)

        
'        If dgInfoPesanBrg.Columns(18) = "" Then
'            .txtNoKonfirmasi.Text = ""
'        Else
'            .txtNoKonfirmasi.Text = dgInfoPesanBrg.Columns(18).Value
'        End If
        
        strSQL = "Select distinct NamaBarang, AsalBarang, JmlOrder, JmlKirim, KdBarang, KdAsal, '0000000000' as NoTerima, NoRegisterAsset " & _
                 "from V_StrukKirimRuanganCetakNM " & _
                 "Where NoKirim = '" & dgInfoPesanBrg.Columns(2).value & "' "
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = True Then Exit Sub
        
       .fgData.Rows = rs.RecordCount + 1
       For i = 1 To rs.RecordCount
            .fgData.TextMatrix(i, 0) = rs(0).value
            .fgData.TextMatrix(i, 1) = rs(1).value
            If IsNull(rs(2)) Then
            .fgData.TextMatrix(i, 2) = 0
            Else
            .fgData.TextMatrix(i, 2) = rs(2).value
            End If
            .fgData.TextMatrix(i, 3) = rs(3).value
            .fgData.TextMatrix(i, 4) = rs(3).value
            .fgData.TextMatrix(i, 5) = rs(7).value
            .fgData.TextMatrix(i, 6) = rs(4).value
            .fgData.TextMatrix(i, 7) = rs(5).value
            .fgData.TextMatrix(i, 8) = rs(6).value
            .fgData.TextMatrix(i, 9) = ""
         

       rs.MoveNext
       Next i
        .Show
'Else
'    MsgBox "Data Kosong", vbInformation
'End If
    End With
Exit Sub
End Sub

Private Sub dgDetail_KeyPress(KeyAscii As Integer)
'
 Select Case dgDetail.Col
            
                Case 5 'Jumlah Konfirmasi Barang
                        'TxtIsiRacikan.MaxLength = 20
                        Call subLoadTextIsi
                        txtIsi.Text = dgDetail.TextMatrix(dgDetail.Row, dgDetail.Col)
                        'TxtIsiRacikan.Text = Chr(KeyAscii)
                        'TxtIsiRacikan.SelStart = Len(TxtIsiRacikan.Text)
                        '            fgRacikan.Rows = fgRacikan.Rows + 1
                Case 6
                    Call subLoadTextIsi
                    txtIsi.Text = dgDetail.TextMatrix(dgDetail.Row, dgDetail.Col)
                    
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
'    Call openConnection
    optBelum.value = True
    
    If optBelum.value = True Then
        cmdVerifikasi.Enabled = False
    Else
        cmdVerifikasi.Enabled = True
    End If
        
    dtpTglAkhir.value = Now
    dtpTglAwal.value = Now
    Call cmdTampilkan_Click
    
End Sub

Private Sub optBelum_Click()
    cmdRetur.Visible = True
    Call cmdTampilkan_Click
End Sub

Private Sub optBelum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdTampilkan.SetFocus
    End If
End Sub

Private Sub optVerifikasi_Click()
    Call cmdTampilkan_Click
    cmdVerifikasi.Enabled = False
    
End Sub

Private Sub optSudah_Click()
    cmdRetur.Visible = False
    Call cmdTampilkan_Click
End Sub

Private Sub txtIsi_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13) Then
    
        dgDetail.TextMatrix(dgDetail.Row, dgDetail.Col) = txtIsi.Text
        txtIsi.Visible = False
        dgDetail.SetFocus
        txtIsi.Text = ""
    End If
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Select Case dgDetail.Col
                Case 5 'Jumlah Konfirmasi
                        Call SetKeyPressToNumber(KeyAscii)
    End Select
End Sub

