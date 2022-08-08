VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPOAKaryawan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pemakaian Obat & Alkes Karyawan"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPOAKaryawan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmPOAKaryawan.frx":0CCA
   ScaleHeight     =   6510
   ScaleWidth      =   11010
   Begin VB.TextBox txtAsalBarang 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3240
      TabIndex        =   28
      Text            =   "txtAsalBarang"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtSatuan 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      TabIndex        =   27
      Text            =   "txtSatuan"
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtKdAsal 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      TabIndex        =   26
      Text            =   "txtKdAsal"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtKdBarang 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   25
      Text            =   "txtKdBarang"
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtKdDokter 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   24
      Text            =   "txtKdDokter"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid dgDokter 
      Height          =   3015
      Left            =   7080
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
            LCID            =   1057
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
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dgHargaBrg 
      Height          =   3015
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
            LCID            =   1057
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
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid dgData 
      Height          =   2655
      Left            =   0
      TabIndex        =   9
      Top             =   3120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   50
      Cols            =   6
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   8577768
      ForeColorFixed  =   -2147483627
      ForeColorSel    =   -2147483628
      BackColorBkg    =   16777215
      FocusRect       =   0
      HighLight       =   2
      FillStyle       =   1
      GridLines       =   3
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   17
      Top             =   2040
      Width           =   10935
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         Height          =   375
         Left            =   8760
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton btnHapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   9735
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtjml 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4200
         MaxLength       =   4
         TabIndex        =   6
         Text            =   "1"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txthargasatuan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txttotbiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtNamaBrg 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah"
         Height          =   210
         Left            =   4200
         TabIndex        =   22
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Harga Satuan"
         Height          =   210
         Left            =   5280
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya"
         Height          =   210
         Left            =   6960
         TabIndex        =   20
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Nama Barang"
         Height          =   210
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   14
      Top             =   960
      Width           =   10935
      Begin VB.TextBox txtKeperluan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   480
         Width           =   4695
      End
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7080
         TabIndex        =   2
         Top             =   480
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
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
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   106954755
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Keperluan Pemakaian"
         Height          =   210
         Left            =   2280
         TabIndex        =   23
         Top             =   240
         Width           =   1725
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Penanggung Jawab"
         Height          =   210
         Left            =   7080
         TabIndex        =   16
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Pemakaian"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1560
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   19
      Top             =   5760
      Width           =   10935
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   9480
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   8040
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   29
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
      Left            =   9120
      Picture         =   "frmPOAKaryawan.frx":190C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPOAKaryawan.frx":2694
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPOAKaryawan.frx":3CF2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmPOAKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim subBolTampil As Boolean

Private Sub btnHapus_Click()
    With dgData
        If .Row = .Rows Then Exit Sub
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        Call msubRemoveItem(dgData, .Row)
        intRowNow = .Rows - 1
    End With
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errSimpan
    If dgData.Rows = 2 Then Exit Sub

    Set dbcmd = New ADODB.Command
    For i = 1 To dgData.Rows - 2
        With dbcmd
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
            .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, dgData.TextMatrix(i, 0))
            .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, dgData.TextMatrix(i, 2))
            .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
            .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, dgData.TextMatrix(i, 5))
            .Parameters.Append .CreateParameter("JmlBrg", adInteger, adParamInput, , dgData.TextMatrix(i, 3))
            .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , CCur(dgData.TextMatrix(i, 4)))
            .Parameters.Append .CreateParameter("TglPemakaian", adDate, adParamInput, , Format(dgData.TextMatrix(i, 7), "yyyy/MM/dd HH:mm:ss"))
            .Parameters.Append .CreateParameter("Keperluan", adVarChar, adParamInput, 100, dgData.TextMatrix(i, 9))
            .Parameters.Append .CreateParameter("PenanggungJawab", adChar, adParamInput, 10, dgData.TextMatrix(i, 6))
            .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, UserID)

            .ActiveConnection = dbConn
            .CommandText = "dbo.Add_PemakaianObatAlkesKaryawan"
            .CommandType = adCmdStoredProc
            .Execute

            If Not (.Parameters("RETURN_VALUE").value = 0) Then
                MsgBox "Ada Kesalahan dalam Penyimpanan Biaya Pelayanan Pasien", vbCritical, "Validasi"
            End If
            Set dbcmd = Nothing
        End With
    Next i
    Call Add_HistoryLoginActivity("Add_PemakaianObatAlkesKaryawan")
    MsgBox "Pemasukan Biaya Pelayanan Pasien Sukses", vbInformation, "Informasi"
    cmdSimpan.Enabled = False
    Exit Sub
errSimpan:
    Set dbcmd = Nothing
    msubPesanError
End Sub

Private Sub cmdTambah_Click()
    On Error GoTo errTambah
    Dim adoCommand As ADODB.Command
    Dim i As Integer

    If Periksa("text", txtDokter, "Penanggung jawab kosong") = False Then Exit Sub
    If sp_CekStokBarang(txtKdBarang.Text, txtKdAsal.Text, txtSatuan.Text, txtjml.Text) = False Then Exit Sub

    With dgData
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = txtKdBarang.Text And .TextMatrix(i, 2) = txtKdAsal.Text And .TextMatrix(i, 5) = txtSatuan.Text Then Exit Sub
        Next i
        intRowNow = .Rows - 1

        .TextMatrix(intRowNow, 0) = txtKdBarang.Text
        .TextMatrix(intRowNow, 1) = txtNamaBrg.Text
        .TextMatrix(intRowNow, 2) = txtKdAsal.Text
        .TextMatrix(intRowNow, 3) = CInt(txtjml.Text)

        strSQL = "SELECT dbo.FB_TakeTarifOAKaryawan(" & CCur(txthargasatuan) & ") AS HargaSatuan "
        Call msubRecFO(dbRst, strSQL)
        .TextMatrix(intRowNow, 4) = dbRst(0).value

        .TextMatrix(intRowNow, 5) = txtSatuan.Text
        .TextMatrix(intRowNow, 6) = txtKdDokter.Text
        .TextMatrix(intRowNow, 7) = Format(dtpTglPeriksa, "dd/mm/yyyy HH:mm:ss")
        .TextMatrix(intRowNow, 8) = txtAsalBarang.Text
        .TextMatrix(intRowNow, 9) = txtKeperluan.Text
        .Rows = .Rows + 1
        .SetFocus
    End With
    txtKdDokter.Text = ""
    txtKdBarang.Text = ""
    txtKdAsal.Text = ""
    txtSatuan.Text = ""
    txtAsalBarang.Text = ""
    txtNamaBrg.Text = ""
    dgHargaBrg.Visible = False
    txtjml.Text = 1
    txthargasatuan.Text = 0
    txttotbiaya.Text = 0

    Exit Sub
errTambah:
    msubPesanError
End Sub

Private Function sp_CekStokBarang(f_KdBarang As String, f_KdAsal As String, f_Satuan As String, f_Jumlah As Integer) As Boolean
    On Error GoTo errLoad
    sp_CekStokBarang = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, f_Satuan)
        .Parameters.Append .CreateParameter("JmlBrg", adInteger, adParamInput, , txtjml)
        .Parameters.Append .CreateParameter("OutputPesan", adChar, adParamOutput, 1, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Check_StokBarangRuangan"
        .CommandType = adCmdStoredProc
        .Execute

        If (.Parameters("OutputPesan").value = "T") Then
            MsgBox "Stok Barang Tidak cukup", vbCritical, "Validasi"
            sp_CekStokBarang = False
        End If
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    sp_CekStokBarang = False
    msubPesanError ("sp_CekStokBarang")
End Function

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then dgDokter.Visible = False: txtDokter.SetFocus
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        subBolTampil = True
        txtDokter.Text = dgDokter.Columns(1)
        subBolTampil = False
        txtKdDokter.Text = dgDokter.Columns(0)
        dgDokter.Visible = False
        txtNamaBrg.SetFocus
    End If
End Sub

Private Sub dgHargaBrg_DblClick()
    Call dgHargaBrg_KeyPress(13)
End Sub

Private Sub dgHargaBrg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtKdBarang.Text = dgHargaBrg.Columns("KdBarang")
        txthargasatuan.Text = dgHargaBrg.Columns("HargaBarang")
        txtSatuan.Text = dgHargaBrg.Columns("Satuan")
        txtKdAsal.Text = dgHargaBrg.Columns("KdAsal")
        txtAsalBarang.Text = dgHargaBrg.Columns("AsalBarang")
        txtNamaBrg.Text = dgHargaBrg.Columns("NamaBarang")

        dgHargaBrg.Visible = False
        txtjml.SetFocus
    End If
End Sub

Private Sub dtpTglPeriksa_Change()
    If dtpTglPeriksa.value < mdTglMasuk Then dtpTglPeriksa.value = mdTglMasuk
End Sub

Private Sub dtpTglPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKeperluan.SetFocus
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    cmdSimpan.Enabled = True
    dtpTglPeriksa.value = Now
    Call centerForm(Me, MDIUtama)
    Call setgrid
    txtKdDokter.Text = ""
    txtKdBarang.Text = ""
    txtKdAsal.Text = ""
    txtSatuan.Text = ""
    txtAsalBarang.Text = ""
End Sub

Private Sub txtDokter_Change()
    If subBolTampil = True Then Exit Sub
    Call subLoadDokter
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter()
    On Error GoTo errLoad

    strSQL = "SELECT TOP 100 IdPegawai,[Nama Lengkap],JK,[Jenis Pegawai] FROM V_DaftarPegawai WHERE [Nama Lengkap] like '%" & txtDokter.Text & "%'"
    Call msubRecFO(rs, strSQL)
    Set dgDokter.DataSource = rs
    With dgDokter
        .Top = 1800
        .Left = 240
        .Columns(0).Width = 1200
        .Columns(1).Width = 5000
        .Columns(2).Width = 400
        .Columns(3).Width = 2000
        .Visible = True
    End With

    Exit Sub
errLoad:
    Call msubPesanError("SubLoadDokter")
End Sub

Private Sub txtDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then If dgDokter.Visible = True Then dgDokter.SetFocus
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 27 Then
        txtDokter = ""
        frameDokter.Visible = False
    End If
    If KeyAscii = 13 Then
        dgDokter.SetFocus
    End If
    If KeyAscii = 39 Then KeyAscii = 0
hell:
End Sub

Private Sub txthargasatuan_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtjml_Change()
    On Error GoTo a:

    txttotbiaya.Text = txtjml.Text * txthargasatuan.Text
    txttotbiaya.Text = Format(txttotbiaya.Text, "#,###.00")
    Exit Sub
a:
    txttotbiaya = 0
End Sub

Private Sub txtjml_GotFocus()
    txtjml.SelStart = 0
    txtjml.SelLength = Len(txtjml.Text)
End Sub

Private Sub txtjml_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTambah.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtjml_LostFocus()
    txtjml.Text = Val(txtjml.Text)
End Sub

Private Sub txtKeperluan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDokter.SetFocus
End Sub

Private Sub txtNamaBrg_Change()
    If subBolTampil = True Then Exit Sub
    Call subLoadbarang
End Sub

Private Sub txtNamaBrg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If dgHargaBrg.Visible = False Then Exit Sub
        dgHargaBrg.SetFocus
    End If
End Sub

Private Sub txtNamaBrg_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 27 Then
        txtNamaBrg = ""
        frameHargaBrg.Visible = False
    End If
    If KeyAscii = 13 Then
        If dgHargaBrg.Visible = True Then dgHargaBrg.SetFocus Else txtjml.SetFocus
    End If
hell:
End Sub

'untuk meload data Barang di grid
Private Sub subLoadbarang()
    On Error GoTo errLoad

    strSQL = "SELECT DISTINCT JenisBarang, NamaBarang, AsalBarang, Satuan, HargaBarang, KdBarang, KdAsal FROM V_HargaNettoBarang WHERE NamaBarang like '%" & txtNamaBrg.Text & "%' and kdruangan='" & mstrKdRuangan & "'  AND (JenisHargaNetto = '1')"
    Call msubRecFO(rs, strSQL)
    Set dgHargaBrg.DataSource = rs
    With dgHargaBrg
        .Top = 2880
        .Left = 240
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i
        .Columns("JenisBarang").Width = 2000
        .Columns("NamaBarang").Width = 4500
        .Columns("AsalBarang").Width = 1100
        .Columns("Satuan").Width = 700
        .Columns("HargaBarang").Width = 1500
        .Columns("HargaBarang").NumberFormat = "#,###.00"
        .Columns("HargaBarang").Alignment = dbgRight
        .Visible = True
    End With

    Exit Sub
errLoad:
    Call msubPesanError("subLoadbarang")
End Sub

Private Sub setgrid()
    With dgData
        .clear
        .Rows = 2
        .Cols = 10
        .TextMatrix(0, 0) = "Kode Barang"
        .TextMatrix(0, 1) = "Nama Barang"
        .TextMatrix(0, 2) = "Kode Asal"
        .TextMatrix(0, 3) = "Jumlah"
        .TextMatrix(0, 4) = "Harga Satuan"
        .TextMatrix(0, 5) = "Satuan"
        .TextMatrix(0, 6) = "Kode Dokter"
        .TextMatrix(0, 7) = "tgl Pelayanan"
        .TextMatrix(0, 8) = "Asal Barang"
        .TextMatrix(0, 9) = "Keperluan"
        .ColWidth(0) = 1200
        .ColWidth(1) = 2500
        .ColWidth(2) = 0
        .ColWidth(3) = 700
        .ColWidth(4) = 1200
        .ColWidth(5) = 700
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 1000
        .ColWidth(9) = 2000
    End With
End Sub

Private Sub txttotbiaya_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub
