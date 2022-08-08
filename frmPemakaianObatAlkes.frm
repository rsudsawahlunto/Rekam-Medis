VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPemakaianObatAlkes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pemakaian Obat & Alkes"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPemakaianObatAlkes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   9330
   Begin VB.Frame frameHargaBrg 
      Caption         =   "Data Harga Barang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   18
      Top             =   2880
      Visible         =   0   'False
      Width           =   9375
      Begin MSDataGridLib.DataGrid dgHargaBrg 
         Height          =   2175
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3836
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
            Size            =   8.25
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
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frameDokter 
      Caption         =   "Data Dokter Pemeriksa"
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
      Left            =   9720
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   7575
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   2295
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4048
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
            Size            =   8.25
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
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lvPemeriksa 
      Height          =   1815
      Left            =   9600
      TabIndex        =   19
      Top             =   270
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nama Pemeriksa"
         Object.Width           =   2540
      EndProperty
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
      TabIndex        =   20
      Top             =   3000
      Width           =   9375
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   8160
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   6960
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
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
      TabIndex        =   21
      Top             =   2040
      Width           =   9375
      Begin VB.TextBox txtNamaBrg 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txttotbiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txthargasatuan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtjml 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3480
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "1"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton btnHapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   8160
         TabIndex        =   11
         Top             =   420
         Width           =   1095
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         Height          =   375
         Left            =   6960
         TabIndex        =   10
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Nama Barang"
         Height          =   210
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya"
         Height          =   210
         Left            =   5400
         TabIndex        =   24
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Harga Satuan"
         Height          =   210
         Left            =   4200
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah"
         Height          =   210
         Left            =   3480
         TabIndex        =   22
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame fraDataDokterPerawat 
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
      TabIndex        =   26
      Top             =   960
      Width           =   9375
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
      Begin VB.CheckBox chkDilayaniDokter 
         Caption         =   "Dokter Pemeriksa "
         Height          =   255
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox txtNamaPerawat 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5760
         TabIndex        =   4
         Text            =   "txtNamaPerawat"
         Top             =   480
         Width           =   3375
      End
      Begin VB.CheckBox chkPerawat 
         Caption         =   "Paramedis"
         Height          =   255
         Left            =   5760
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   525
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
         Format          =   115736579
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Periksa"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1260
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgPerawatPerPelayanan 
      Height          =   1455
      Left            =   3360
      TabIndex        =   15
      Top             =   4320
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2566
      _Version        =   393216
      FixedCols       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid GridAlat 
      Height          =   1935
      Left            =   0
      TabIndex        =   14
      Top             =   3840
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3413
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   28
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
      Picture         =   "frmPemakaianObatAlkes.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPemakaianObatAlkes.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPemakaianObatAlkes.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmPemakaianObatAlkes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strpelayanan As String
Dim strkdbarang As String
Dim intJml As Integer
Dim strKdAsal As String
Dim strSatuanJml As String
Dim strnoreg As String
Dim strKdKelas As String
Dim boo As Boolean
Dim strAsalBarang As String
Dim strFilterDokter As String
Dim intRowNow As Integer

Dim subKdPemeriksa() As String
Dim subJmlTotal As Integer
Dim i As Integer
Dim j As Integer
Dim subcurHargaSatuan As Currency

Private Sub btnHapus_Click()
    With GridAlat
        If .Row = .Rows Then Exit Sub
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        Call msubRemoveItem(GridAlat, .Row)
        intRowNow = .Rows - 1
    End With
End Sub

Private Sub chkDilayaniDokter_Click()
    On Error GoTo errLoad

    If chkDilayaniDokter.value = 0 Then
        txtDokter.Enabled = False
        txtDokter.Text = ""

        If frameDokter.Visible = True Then frameDokter.Visible = False
    Else
        lvPemeriksa.Visible = False

        txtDokter.Enabled = True
        strSQL = "SELECT dbo.RegistrasiRJ.IdDokter, dbo.DataPegawai.NamaLengkap " & _
        " FROM dbo.RegistrasiRJ INNER JOIN dbo.DataPegawai ON dbo.RegistrasiRJ.IdDokter = dbo.DataPegawai.IdPegawai " & _
        " WHERE (dbo.RegistrasiRJ.NoPendaftaran = '" & mstrNoPen & "')"
        Call msubRecFO(rs, strSQL)

        If Not rs.EOF Then
            txtDokter.Text = rs(1).value
            mstrKdDokter = rs(0).value
            intJmlDokter = rs.RecordCount
            frameDokter.Visible = False
        End If
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub chkDilayaniDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkDilayaniDokter.value = 0 Then
            chkPerawat.SetFocus
        Else
            txtDokter.SetFocus
        End If
    End If
End Sub

Private Sub chkPerawat_Click()
    If chkPerawat.value = vbChecked Then
        strSQL = "SELECT IdPegawai FROM V_DaftarPemeriksaPasien WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            txtNamaPerawat.Text = strNmPegawai
            If lvPemeriksa.ListItems.Count > 0 Then
                lvPemeriksa.ListItems.Item("key" & strIDPegawaiAktif).Checked = True
                Call lvPemeriksa_ItemCheck(lvPemeriksa.ListItems.Item("key" & strIDPegawaiAktif))
            End If
        Else
            txtNamaPerawat.Text = ""
        End If
    Else
        txtNamaPerawat.Text = ""
    End If
    lvPemeriksa.Visible = False
End Sub

Private Sub chkPerawat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkPerawat.value = vbChecked Then
            txtNamaPerawat.SetFocus
        Else
            txtNamaBrg.SetFocus
        End If
    End If
End Sub

Private Sub cmdSimpan_Click()
    Dim adoCommand As New ADODB.Command
    Dim i As Integer
    If GridAlat.Rows = 2 Then Exit Sub

    For i = 1 To GridAlat.Rows - 2
        With adoCommand
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
            .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, GridAlat.TextMatrix(i, 0))
            .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, GridAlat.TextMatrix(i, 2))
            .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuanganPasien)
            .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, GridAlat.TextMatrix(i, 5))
            .Parameters.Append .CreateParameter("JmlBrg", adInteger, adParamInput, , CInt(GridAlat.TextMatrix(i, 3)))
            .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, strnoreg)
            .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
            .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, strKdKelas)
            .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , CCur(GridAlat.TextMatrix(i, 4)))
            .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(GridAlat.TextMatrix(i, 7), "yyyy/MM/dd HH:mm:ss"))
            .Parameters.Append .CreateParameter("NoLabRad", adChar, adParamInput, 10, Null)
            .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, GridAlat.TextMatrix(i, 6))
            .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
            .Parameters.Append .CreateParameter("IdPegawai2", adChar, adParamInput, 10, Null)
            .Parameters.Append .CreateParameter("StatusStok", adChar, adParamInput, 1, "0")

            .ActiveConnection = dbConn
            .CommandText = "Add_PemakaianObatAlkes"
            .CommandType = adCmdStoredProc
            .Execute

            If Not (.Parameters("RETURN_VALUE").value = 0) Then
                MsgBox "Ada Kesalahan dalam Penyimpanan Biaya Pelayanan Pasien", vbCritical, "Validasi"
            End If
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
        End With

    Next i

    For i = 1 To fgPerawatPerPelayanan.Rows - 1
        With fgPerawatPerPelayanan
            If sp_PetugasPemeriksaOA(.TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 5), .TextMatrix(i, 6), .TextMatrix(i, 7)) = False Then Exit Sub
        End With
    Next i

    Call Add_HistoryLoginActivity("Add_PemakaianObatAlkes+Add_PetugasPemeriksaOA")
    cmdSimpan.Enabled = False
    frmTransaksiPasien.subpemakaianobatalkes
End Sub

'simpan data perawat
Private Function sp_PetugasPemeriksaOA(F_dtTanggalPelayanan As Date, F_strKodeBarang As String, F_strKodeAsal As String, F_strSatuanJml As String, F_StrIdPerawat As String, F_IdUser As String) As Boolean
    On Error GoTo errLoad

    sp_PetugasPemeriksaOA = False

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, strnoreg)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuanganPasien)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(F_dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdBarang", adChar, adParamInput, 9, F_strKodeBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, F_strKodeAsal)
        .Parameters.Append .CreateParameter("SatuanJml", adChar, adParamInput, 1, F_strSatuanJml)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, F_StrIdPerawat)  'kode perawat
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, F_IdUser)

        .ActiveConnection = dbConn
        .CommandText = "Add_PetugasPemeriksaOA"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data petugas pemeriksa BP", vbExclamation, "Validasi"
            sp_PetugasPemeriksaOA = False

        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
        sp_PetugasPemeriksaOA = True
    End With

    Exit Function
errLoad:
    Call msubPesanError
    sp_PetugasPemeriksaOA = False
End Function

Private Sub cmdTambah_Click()
    Dim adoCommand As ADODB.Command
    If (mstrKdDokter = "") And (chkDilayaniDokter.value = 1) Then
        MsgBox "Pilih dulu Dokter Pemeriksa Pasien", vbCritical, "Validasi"
        txtDokter.SetFocus
        Exit Sub
    End If
    If txtjml = "" Then
        MsgBox "Jumlah harus diisi", vbInformation, "Informasi"
        Exit Sub
    End If

    If chkPerawat.value = vbChecked And subJmlTotal = 0 Then
        MsgBox "Nama perawat kosong", vbCritical, "Validasi"
        lvPemeriksa.Visible = True
        txtNamaPerawat.SetFocus
        Exit Sub
    End If

    Call sp_cekstobarang(dbcmd, strkdbarang, strKdAsal, mstrKdRuanganPasien, strSatuanJml, txtjml)
    If boo = False Then Exit Sub

    Dim i As Integer
    With GridAlat
        If strkdbarang = "" Then Exit Sub
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = strkdbarang And .TextMatrix(i, 2) = strKdAsal And .TextMatrix(i, 5) = strSatuanJml Then Exit Sub
        Next i
        intRowNow = .Rows - 1

        .TextMatrix(intRowNow, 0) = strkdbarang
        .TextMatrix(intRowNow, 1) = txtNamaBrg.Text
        .TextMatrix(intRowNow, 2) = strKdAsal
        .TextMatrix(intRowNow, 3) = CInt(txtjml.Text)

        strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & strKdAsal & "', " & CCur(txthargasatuan) & ")  as HargaSatuan"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then subcurHargaSatuan = 0 Else subcurHargaSatuan = rs(0).value
        .TextMatrix(intRowNow, 4) = subcurHargaSatuan

        .TextMatrix(intRowNow, 5) = strSatuanJml
        If chkDilayaniDokter.value = 1 Then
            .TextMatrix(intRowNow, 6) = mstrKdDokter
        Else
            .TextMatrix(intRowNow, 6) = UserID
        End If
        .TextMatrix(intRowNow, 7) = Format(dtpTglPeriksa, "dd/mm/yyyy HH:mm:ss")
        .TextMatrix(intRowNow, 8) = strAsalBarang
        .Rows = .Rows + 1
        .SetFocus
    End With

    If chkPerawat.value = vbChecked Then Call subLoadPelayananPerPerawat
    txtNamaBrg.Text = ""
    txtjml.Text = 1
    txthargasatuan.Text = 0
    txttotbiaya.Text = 0
    frameHargaBrg.Visible = False
    chkPerawat.SetFocus
End Sub

Private Sub subLoadPelayananPerPerawat()
    With fgPerawatPerPelayanan
        For i = 1 To subJmlTotal
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = strnoreg
            .TextMatrix(.Rows - 1, 1) = mstrKdRuanganPasien
            .TextMatrix(.Rows - 1, 2) = dtpTglPeriksa.value
            .TextMatrix(.Rows - 1, 3) = GridAlat.TextMatrix(GridAlat.Rows - 2, 0) 'kode barang
            .TextMatrix(.Rows - 1, 4) = GridAlat.TextMatrix(GridAlat.Rows - 2, 2) 'kode asal
            .TextMatrix(.Rows - 1, 5) = GridAlat.TextMatrix(GridAlat.Rows - 2, 5) 'satuan
            .TextMatrix(.Rows - 1, 6) = Mid(subKdPemeriksa(i), 4, Len(subKdPemeriksa(i)) - 3)
            .TextMatrix(.Rows - 1, 7) = strIDPegawaiAktif
        Next i
    End With

    subJmlTotal = 0
    txtNamaPerawat.BackColor = &HFFFFFF
    ReDim Preserve subKdPemeriksa(subJmlTotal)
    chkPerawat.value = vbUnchecked
End Sub

Private Sub sp_cekstobarang(ByVal adoCommand As ADODB.Command, strkdbarang As String, strKdAsal As String, strKdRuangan As String, strSatuanJml As String, txtjml As Integer)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, strkdbarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, strKdAsal)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, strKdRuangan)
        .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, strSatuanJml)
        .Parameters.Append .CreateParameter("JmlBrg", adInteger, adParamInput, , txtjml)
        .Parameters.Append .CreateParameter("OutputPesan", adChar, adParamOutput, 1, Null)
        .ActiveConnection = dbConn
        .CommandText = "dbo.Check_StokBarangRuangan"
        .CommandType = adCmdStoredProc
        .Execute

        If (.Parameters("OutputPesan").value = "T") Then
            MsgBox "Stok Barang Tidak cukup", vbCritical, "Validasi"
            boo = False
        Else
            boo = True
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

Private Sub cmdTutup_Click()
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan data pemakaian obat & alkes?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Unload Me
    frmTransaksiPasien.Enabled = True
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        txtDokter.Text = dgDokter.Columns(1).value
        mstrKdDokter = dgDokter.Columns(0).value
        If mstrKdDokter = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        chkDilayaniDokter.value = 1
        frameDokter.Visible = False
        chkPerawat.SetFocus
    End If
    If KeyAscii = 27 Then
        frameDokter.Visible = False
    End If
End Sub

Private Sub dgHargaBrg_DblClick()
    Call dgHargaBrg_KeyPress(13)
End Sub

Private Sub dgHargaBrg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJml = 0 Then Exit Sub
        Dim strKdBrg As String
        Dim intHrgSat As Currency
        Dim strSatJml As String
        Dim strKdAsl As String
        Dim strNmAsl As String
        strKdBrg = dgHargaBrg.Columns("KdBarang").value
        intHrgSat = dgHargaBrg.Columns("HargaBarang").value
        strKdAsl = dgHargaBrg.Columns("KdAsal").value
        strNmAsl = dgHargaBrg.Columns("AsalBarang").value
        txtNamaBrg.Text = dgHargaBrg.Columns("NamaBarang").value
        strkdbarang = strKdBrg
        txthargasatuan = intHrgSat
        strSatuanJml = strSatJml
        strKdAsal = strKdAsl
        strAsalBarang = strNmAsl
        If strkdbarang = "" Then
            MsgBox "Nama Barang belum dipilih", vbCritical, "Validasi"
            txtNamaBrg.Text = ""
            dgHargaBrg.SetFocus
            Exit Sub
        End If
        Call txtjml_Change
        frameHargaBrg.Visible = False
        txtjml.SetFocus
    End If
    If KeyAscii = 27 Then
        frameHargaBrg.Visible = False
    End If
End Sub

Private Sub dtpTglPeriksa_Change()
    If dtpTglPeriksa.value < mdTglMasuk Then dtpTglPeriksa = mdTglMasuk
End Sub

Private Sub dtpTglPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then chkDilayaniDokter.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call PlayFlashMovie(Me)
    cmdSimpan.Enabled = True

    dtpTglPeriksa.value = Now
    strnoreg = mstrNoPen
    strKdKelas = mstrKdKelas
    Call centerForm(Me, MDIUtama)
    frmTransaksiPasien.Enabled = False
    Call setgrid

    subJmlTotal = 0
    Call subSetGridPerawatPerPelayanan
    Call subLoadListPemeriksa

    chkPerawat.value = vbChecked
    lvPemeriksa.Visible = False

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmTransaksiPasien.Enabled = True
End Sub

Private Sub lvPemeriksa_DblClick()
    Call lvPemeriksa_KeyPress(13)
End Sub

Private Sub lvPemeriksa_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim blnSelected As Boolean
    If Item.Checked = True Then
        subJmlTotal = subJmlTotal + 1
        ReDim Preserve subKdPemeriksa(subJmlTotal)
        subKdPemeriksa(subJmlTotal) = Item.Key
    Else
        blnSelected = False
        For i = 1 To subJmlTotal
            If subKdPemeriksa(i) = Item.Key Then blnSelected = True
            If blnSelected = True Then
                If i = subJmlTotal Then
                    subKdPemeriksa(i) = ""
                Else
                    subKdPemeriksa(i) = subKdPemeriksa(i + 1)
                End If
            End If
        Next i
        subJmlTotal = subJmlTotal - 1
    End If

    If subJmlTotal = 0 Then
        txtNamaPerawat.BackColor = &HFFFFFF
    Else
        txtNamaPerawat.BackColor = &HC0FFFF
    End If
End Sub

Private Sub lvPemeriksa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lvPemeriksa.Visible = False: txtNamaPerawat.SetFocus
End Sub

Private Sub txtDokter_Change()
    strFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
    mstrKdDokter = ""
    frameDokter.Visible = True
    Call subLoadDokter
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter()
    On Error Resume Next
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & strFilterDokter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDokter = rs.RecordCount
    Set dgDokter.DataSource = rs
    With dgDokter
        .Columns(0).Width = 1200
        .Columns(1).Width = 2500
        .Columns(2).Width = 400
        .Columns(3).Width = 2000
    End With
    frameDokter.Left = 0
    frameDokter.Top = 1800
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 27 Then
        txtDokter = ""
        frameDokter.Visible = False
    End If
    If KeyAscii = 13 Then
        If frameDokter.Visible = True Then
            dgDokter.SetFocus
        Else
            chkPerawat.SetFocus
        End If
    End If
    Call SetKeyPressToChar(KeyAscii)
hell:
End Sub

Private Sub txthargasatuan_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtjml_Change()
    On Error GoTo a:
    txttotbiaya = txtjml * txthargasatuan
    Exit Sub
a:
    txttotbiaya = ""
End Sub

Private Sub txtjml_GotFocus()
    txtjml.SelStart = 0
    txtjml.SelLength = Len(txtjml.Text)
End Sub

Private Sub txtjml_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTambah.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtjml_LostFocus()
    If txtjml.Text = "" Then txtjml.Text = 1: Exit Sub
    If txtjml.Text = 0 Then txtjml.Text = 1
End Sub

Private Sub txtNamaBrg_Change()
    strpelayanan = "WHERE NamaBarang like '%" & txtNamaBrg.Text & "%' and KdRuangan='" & mstrKdRuanganPasien & "' AND KdKelompokPasien = '" & mstrKdJenisPasien & "' AND IdPenjamin = '" & mstrKdPenjaminPasien & "'"
    strkdbarang = ""
    frameHargaBrg.Visible = True
    Call subLoadbarang
End Sub

'untuk meload data Barang di grid
Private Sub subLoadbarang()
    On Error Resume Next
    strSQL = "SELECT * FROM V_HargaNettoBarang " & strpelayanan
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJml = rs.RecordCount
    Set dgHargaBrg.DataSource = rs
    With dgHargaBrg
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i
        .Columns("JenisBarang").Width = 1200
        .Columns("NamaBarang").Width = 3165
        .Columns("AsalBarang").Width = 1000
        .Columns("JenisPasien").Width = 1100
        .Columns("Satuan").Width = 675
        .Columns("HargaBarang").Width = 1200
        .Columns("HargaBarang").NumberFormat = "#,###"
        .Columns("HargaBarang").Alignment = dbgRight
    End With
    frameHargaBrg.Left = 0
    frameHargaBrg.Top = 3000
End Sub

Private Sub setgrid()
    With GridAlat
        .clear
        .Rows = 2
        .Cols = 9
        .TextMatrix(0, 0) = "Kode Barang"
        .TextMatrix(0, 1) = "Nama Barang"
        .TextMatrix(0, 2) = "Kode Asal"
        .TextMatrix(0, 3) = "Jumlah"
        .TextMatrix(0, 4) = "Harga Satuan"
        .TextMatrix(0, 5) = "Sat"
        .TextMatrix(0, 6) = "Kode Dokter"
        .TextMatrix(0, 7) = "tgl Pelayanan"
        .TextMatrix(0, 8) = "Asal Barang"
        .ColWidth(0) = 1200
        .ColWidth(1) = 4300
        .ColWidth(2) = 0
        .ColWidth(3) = 800
        .ColWidth(4) = 1200
        .ColWidth(5) = 400
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 1000
    End With
End Sub

Private Sub txtNamaBrg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If frameHargaBrg.Visible = False Then Exit Sub
        dgHargaBrg.SetFocus
    End If
End Sub

Private Sub txtNamaBrg_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 27 Then
        txtNamaBrg = ""
        txtjml.Text = ""
        txthargasatuan.Text = ""
        txttotbiaya.Text = ""
        frameHargaBrg.Visible = False
    End If
    If KeyAscii = 13 Then
        If frameHargaBrg.Visible = True Then
            dgHargaBrg.SetFocus
        Else
            txtjml.SetFocus
        End If
    End If
hell:
End Sub

Private Sub txtNamaPerawat_Change()
    On Error GoTo errLoad

    Call subLoadListPemeriksa("where [Nama Pemeriksa] LIKE '%" & txtNamaPerawat.Text & "%'")
    lvPemeriksa.Visible = True

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtNamaPerawat_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If lvPemeriksa.Visible = True Then If lvPemeriksa.ListItems.Count > 0 Then lvPemeriksa.SetFocus
        Case vbKeyEscape
            lvPemeriksa.Visible = False
    End Select
End Sub

Private Sub txtNamaPerawat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If lvPemeriksa.Visible = True Then
            lvPemeriksa.SetFocus
        Else
            txtNamaBrg.SetFocus
        End If
    End If
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub subSetGridPerawatPerPelayanan()
    With fgPerawatPerPelayanan
        .Cols = 8
        .Rows = 1

        .TextMatrix(0, 0) = "NoPendaftaran"
        .TextMatrix(0, 1) = "Kode Ruangan"
        .TextMatrix(0, 2) = "Tgl Pelayanan"
        .TextMatrix(0, 3) = "Kode Barang"
        .TextMatrix(0, 4) = "Kode Asal"
        .TextMatrix(0, 5) = "Satuan"
        .TextMatrix(0, 6) = "IdPegawai"
        .TextMatrix(0, 7) = "IdUser"
    End With
End Sub

Private Sub subLoadListPemeriksa(Optional strKriteria As String)
    Dim strKey As String

    strSQL = "select * from v_daftarpemeriksapasien " & strKriteria & " order by [Nama Pemeriksa]"
    Call msubRecFO(rs, strSQL)

    With lvPemeriksa
        .ListItems.clear
        For i = 0 To rs.RecordCount - 1
            strKey = "key" & rs(0).value
            .ListItems.Add , strKey, rs(1).value
            rs.MoveNext
        Next

        .Top = fraDataDokterPerawat.Top + txtNamaPerawat.Top + txtNamaPerawat.Height
        .Left = fraDataDokterPerawat.Left + txtNamaPerawat.Left
        .Height = 1815
        .ColumnHeaders.Item(1).Width = lvPemeriksa.Width - 500

        If subJmlTotal = 0 Then Exit Sub
        For i = 1 To .ListItems.Count
            For j = 1 To subJmlTotal
                If .ListItems(i).Key = subKdPemeriksa(j) Then .ListItems(i).Checked = True
            Next j
        Next i
    End With
End Sub

Private Function sp_Take_TarifOA(f_KdAsal As String, f_HargaSatuan As Currency) As Currency
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 6, f_KdAsal)
        .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , CCur(f_HargaSatuan))
        .Parameters.Append .CreateParameter("TarifTotal", adCurrency, adParamOutput, , Null)

        .ActiveConnection = dbConn
        .CommandText = "Take_TarifOA"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam Pengambilan biaya tarif", vbExclamation, "Validasi"
            sp_Take_TarifOA = 0
        Else
            sp_Take_TarifOA = .Parameters("TarifTotal").value
            Call Add_HistoryLoginActivity("Take_TarifOA")
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub txttotbiaya_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

