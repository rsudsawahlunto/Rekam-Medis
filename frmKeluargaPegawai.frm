VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmKeluargaPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Keluarga Pegawai"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKeluargaPegawai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   10320
   Begin VB.TextBox txtNoUrut 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtIdPegawai 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5040
      TabIndex        =   32
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      TabIndex        =   18
      Top             =   1080
      Width           =   10215
      Begin MSDataGridLib.DataGrid dgDataPegawai 
         Height          =   2535
         Left            =   3360
         TabIndex        =   31
         Top             =   1440
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   0
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1680
         Width           =   8295
      End
      Begin VB.TextBox txtTahun 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6360
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtBulan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7080
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtHari 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7800
         MaxLength       =   2
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
      Begin VB.ComboBox cboJnsKelamin 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "frmKeluargaPegawai.frx":0CCA
         Left            =   3360
         List            =   "frmKeluargaPegawai.frx":0CCC
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtNamaKeluarga 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1080
         Width           =   3015
      End
      Begin MSDataListLib.DataCombo dcHubungan 
         Height          =   330
         Left            =   3360
         TabIndex        =   1
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox txtNamaPegawai 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         MaxLength       =   50
         TabIndex        =   0
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtCariNamaPegawai 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   12
         Top             =   5040
         Width           =   2835
      End
      Begin MSDataGridLib.DataGrid dgKeluargaPegawai 
         Height          =   2775
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4895
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox meTglLahir 
         Height          =   390
         Left            =   4680
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   688
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         HideSelection   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mm-yy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo dcPekerjaan 
         Height          =   330
         Left            =   5760
         TabIndex        =   2
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcPendidikan 
         Height          =   330
         Left            =   7920
         TabIndex        =   3
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   210
         Index           =   7
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Pendidikan"
         Height          =   210
         Index           =   6
         Left            =   7920
         TabIndex        =   29
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Pekerjaan"
         Height          =   210
         Index           =   5
         Left            =   5760
         TabIndex        =   28
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblumur 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   6360
         TabIndex        =   27
         Top             =   840
         Width           =   525
      End
      Begin VB.Label lblTglLhr 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Lahir"
         Height          =   210
         Left            =   4800
         TabIndex        =   26
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Hari"
         Height          =   210
         Left            =   7800
         TabIndex        =   25
         Top             =   840
         Width           =   300
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Bulan"
         Height          =   210
         Left            =   7080
         TabIndex        =   24
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Index           =   3
         Left            =   3360
         TabIndex        =   23
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nama Lengkap Keluarga"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1950
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Hubungan dgn Keluarga"
         Height          =   210
         Index           =   1
         Left            =   3360
         TabIndex        =   21
         Top             =   240
         Width           =   1965
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pegawai"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Nama Pegawai"
         Height          =   210
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   5040
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   8640
      TabIndex        =   16
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   5280
      TabIndex        =   15
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   6720
      Width           =   1575
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   33
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
      Left            =   8520
      Picture         =   "frmKeluargaPegawai.frx":0CCE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmKeluargaPegawai.frx":1A56
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmKeluargaPegawai.frx":4417
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmKeluargaPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vTampil As Boolean

Private Function sp_KeluargaPegawai(f_Status As String) As Boolean
    On Error GoTo errLoad
    sp_KeluargaPegawai = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtIdPegawai.Text)
        .Parameters.Append .CreateParameter("KdHubungan", adChar, adParamInput, 2, dcHubungan.BoundText)
        .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, IIf(txtNoUrut.Text = "", Null, txtNoUrut.Text))
        .Parameters.Append .CreateParameter("NamaLengkap", adVarChar, adParamInput, 30, txtNamaKeluarga.Text)
        .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 1, IIf(cboJnsKelamin.Text = "Laki-Laki", "L", "P"))
        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , IIf(meTglLahir.Text = "", Null, Format(meTglLahir.Text, "yyyy/MM/dd")))
        .Parameters.Append .CreateParameter("KdPekerjaan", adChar, adParamInput, 2, IIf(dcPekerjaan.Text = "", Null, dcPekerjaan.BoundText))
        .Parameters.Append .CreateParameter("KdPendidikan", adChar, adParamInput, 2, IIf(dcPendidikan.Text = "", Null, dcPendidikan.BoundText))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, IIf(txtketerangan.Text = "", Null, txtketerangan.Text))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_KeluargaPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_KeluargaPegawai = False
        Else
            Call Add_HistoryLoginActivity("AUD_KeluargaPegawai")
        End If
        Set dbcmd = Nothing
        Call deleteADOCommandParameters(dbcmd)
    End With
    Exit Function
errLoad:
    sp_KeluargaPegawai = False
    Call msubPesanError
End Function

Private Sub subLoadDcSource()
    On Error GoTo errLoad
    strSQL = "SELECT Hubungan, NamaHubungan FROM  HubunganKeluarga where StatusEnabled='1'"
    Call msubDcSource(dcHubungan, rs, strSQL)

    strSQL = "SELECT  KdPekerjaan, Pekerjaan FROM Pekerjaan where StatusEnabled='1'"
    Call msubDcSource(dcPekerjaan, rs, strSQL)

    strSQL = "SELECT KdPendidikan, Pendidikan FROM Pendidikan where StatusEnabled='1'"
    Call msubDcSource(dcPendidikan, rs, strSQL)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subKosong()
    vTampil = False
    dgDataPegawai.Visible = False
    txtNamaPegawai.Text = ""
    dcHubungan.Text = ""
    dcPekerjaan.Text = ""
    dcPendidikan.Text = ""
    txtNamaKeluarga.Text = ""
    cboJnsKelamin.Text = ""
    meTglLahir.Text = "__/__/____"
    txtTahun.Text = ""
    txtBulan.Text = ""
    txtHari.Text = ""
    txtketerangan.Text = ""
    txtNoUrut.Text = ""
    dgDataPegawai.Visible = False
End Sub

Private Sub subLoadDataPegawai()
    On Error GoTo hell_
    If vTampil = False Then Exit Sub
    dgDataPegawai.Visible = True
    strSQL = "SELECT IdPegawai, NamaLengkap FROM DataPegawai WHERE NamaLengkap LIKE '" & txtNamaPegawai.Text & "%' "
    Call msubRecFO(rs, strSQL)
    Set dgDataPegawai.DataSource = rs
    dgDataPegawai.Columns(0).Width = 0
    dgDataPegawai.Columns(1).Width = 4000
    dgDataPegawai.Top = 840
    dgDataPegawai.Left = 120
    Exit Sub
hell_:
    msubPesanError
End Sub

Private Sub subLoadGrid()
    On Error GoTo hell_
    Set rs = Nothing
    strSQL = "SELECT  NamaPegawai, NamaHubungan, NoUrut,  JenisKelamin, TglLahir, Pekerjaan, Pendidikan, NamaKeluarga, Keterangan, IdPegawai, KdHubungan " & _
    " FROM  V_KeluargaPegawai WHERE NamaPegawai LIKE '" & txtCariNamaPegawai.Text & "%' "
    Call msubRecFO(rs, strSQL)
    Set dgKeluargaPegawai.DataSource = rs
    With dgKeluargaPegawai
        .Columns("NamaPegawai").Width = 2500
        .Columns("NamaHubungan").Width = 1000
        .Columns("NamaHubungan").Caption = "Hubungan"
        .Columns("JenisKelamin").Width = 400
        .Columns("JenisKelamin").Caption = "JK"
        .Columns("TglLahir").Width = 1800
        .Columns("Pekerjaan").Width = 1500
        .Columns("Pendidikan").Width = 1500
        .Columns("NamaKeluarga").Width = 2500
        .Columns("Keterangan").Width = 3000
        .Columns("IdPegawai").Width = 0
        .Columns("KdHubungan").Width = 0
        .Columns("NoUrut").Width = 700
    End With
    Exit Sub
hell_:
    msubPesanError
End Sub

Private Sub cboJnsKelamin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then meTglLahir.SetFocus
End Sub

Private Sub cmdBatal_Click()
    Call subKosong
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errLoad

    If txtNamaPegawai.Text = "" Then
        MsgBox "Pilih data yang akan dihapus", vbExclamation, "validasi"
        Exit Sub
    End If
    If MsgBox("Apakah anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    If sp_KeluargaPegawai("D") = False Then Exit Sub
    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
    Call subLoadGrid

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    If Len(Trim(cboJnsKelamin.Text)) = 0 Then MsgBox "Jenis Kelamin tidak boleh kosong!!!"
    If Periksa("text", txtNamaPegawai, "Nama Pegawai tidak boleh kosong!! ") = False Then Exit Sub
    If Periksa("datacombo", dcHubungan, "hubungan pegawai dengan keluarga tidak boleh kosong!! ") = False Then Exit Sub
    If sp_KeluargaPegawai("A") = False Then Exit Sub
    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    Call subLoadGrid
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub dcHubungan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcHubungan.MatchedWithList = True Then dcPekerjaan.SetFocus
        strSQL = "SELECT Hubungan, NamaHubungan FROM  HubunganKeluarga where StatusEnabled='1' and (NamaHubungan LIKE '%" & dcHubungan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcHubungan.Text = ""
            Exit Sub
        End If
        dcHubungan.BoundText = rs(0).value
        dcHubungan.Text = rs(1).value
    End If
End Sub

Private Sub dcPekerjaan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcPekerjaan.MatchedWithList = True Then dcPendidikan.SetFocus
        strSQL = "SELECT  KdPekerjaan, Pekerjaan FROM Pekerjaan where StatusEnabled='1' and (Pekerjaan LIKE '%" & dcPekerjaan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcPekerjaan.Text = ""
            Exit Sub
        End If
        dcPekerjaan.BoundText = rs(0).value
        dcPekerjaan.Text = rs(1).value
    End If
End Sub

Private Sub dcPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcPendidikan.MatchedWithList = True Then txtNamaKeluarga.SetFocus
        strSQL = "SELECT KdPendidikan, Pendidikan FROM Pendidikan where StatusEnabled='1' and (Pendidikan LIKE '%" & dcPendidikan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcPendidikan.Text = ""
            Exit Sub
        End If
        dcPendidikan.BoundText = rs(0).value
        dcPendidikan.Text = rs(1).value
    End If
End Sub

Private Sub dgDataPegawai_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDataPegawai
    WheelHook.WheelHook dgDataPegawai
End Sub

Private Sub dgDataPegawai_DblClick()
    Call dgDataPegawai_KeyPress(13)
End Sub

Private Sub dgDataPegawai_KeyPress(KeyAscii As Integer)
    On Error GoTo hell_
    If KeyAscii = 13 Then
        txtIdPegawai.Text = dgDataPegawai.Columns(0).value
        txtNamaPegawai.Text = dgDataPegawai.Columns(1).value
        dgDataPegawai.Visible = False
        dcHubungan.SetFocus
    End If
    Exit Sub
hell_:
    msubPesanError
End Sub

Private Sub dgKeluargaPegawai_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKeluargaPegawai
    WheelHook.WheelHook dgKeluargaPegawai
End Sub

Private Sub dgKeluargaPegawai_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    With dgKeluargaPegawai
        txtNamaPegawai.Text = .Columns("NamaPegawai")
        dcHubungan.Text = .Columns("Hubungan")
        txtNamaKeluarga.Text = .Columns("NamaKeluarga")
        If .Columns("JK") = "L" Then
            cboJnsKelamin.Text = "Laki-Laki"
        Else
            cboJnsKelamin.Text = "Perempuan"
        End If
        meTglLahir.Text = .Columns("TglLahir")
        dcPekerjaan.Text = .Columns("Pekerjaan")
        dcPendidikan.Text = .Columns("Pendidikan")
        txtketerangan.Text = .Columns("Keterangan")
        txtIdPegawai.Text = .Columns("IdPegawai")
        txtNoUrut.Text = .Columns("NoUrut")
    End With
    dgDataPegawai.Visible = False
End Sub

Private Sub Form_Activate()
    cboJnsKelamin.AddItem "Laki-Laki"
    cboJnsKelamin.AddItem "Perempuan"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then dgDataPegawai.Visible = False
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call subLoadDcSource
    Call subLoadGrid

    Exit Sub
errLoad:
End Sub

Private Sub meTglLahir_KeyPress(KeyAscii As Integer)
    On Error GoTo errTglLahir
    If KeyAscii = 13 Then
        If meTglLahir.Text = "__/__/____" Then
            txtTahun.SetFocus
            Exit Sub
        End If
        If funcCekValidasiTgl("TglLahir", meTglLahir) = "NoErr" Then
            Call subYearOldCount(Format(meTglLahir.Text, "yyyy/mm/dd"))
            txtTahun.Text = YOC_intYear
            txtBulan.Text = YOC_intMonth
            txtHari.Text = YOC_intDay
            txtketerangan.SetFocus
        Else
            txtTahun.Text = ""
            txtBulan.Text = ""
            txtHari.Text = ""
        End If
    End If
    Call SetKeyPressToNumber(KeyAscii)
    Exit Sub
errTglLahir:
    If Err.Number = 5 Then Exit Sub
    MsgBox "Sudahkah anda mengganti Regional Setting komputer anda menjadi 'Indonesia'?" _
    & vbNewLine & "Bila sudah hubungi Administrator dan laporkan pesan kesalahan berikut:" _
    & vbNewLine & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
End Sub

Private Sub meTglLahir_LostFocus()
    On Error GoTo errTglLahir
    If meTglLahir.Text = "__/__/____" Then Exit Sub
    If funcCekValidasiTgl("TglLahir", meTglLahir) = "NoErr" Then
        Call subYearOldCount(Format(meTglLahir.Text, "yyyy/mm/dd"))
        txtTahun.Text = YOC_intYear
        txtBulan.Text = YOC_intMonth
        txtHari.Text = YOC_intDay
    Else
        txtTahun.Text = ""
        txtBulan.Text = ""
        txtHari.Text = ""
    End If
    Exit Sub
errTglLahir:
    MsgBox "Sudahkah anda mengganti Regional Setting komputer anda menjadi 'Indonesia'?" _
    & vbNewLine & "Bila sudah hubungi Administrator dan laporkan pesan kesalahan berikut:" _
    & vbNewLine & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
End Sub

Private Sub txtBulan_Change()
    Dim dTglLahir As Date
    If txtBulan.Text = "" And txtTahun.Text = "" Then txtHari.SetFocus: Exit Sub
    If txtBulan.Text = "" Then txtBulan.Text = 0
    If txtTahun.Text = "" And txtHari.Text = "" Then
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
    ElseIf txtTahun.Text <> "" And txtHari.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    ElseIf txtTahun.Text = "" And txtHari.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
    ElseIf txtTahun.Text <> "" And txtHari.Text = "" Then
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    End If
    meTglLahir.Text = dTglLahir
End Sub

Private Sub txtBulan_KeyPress(KeyAscii As Integer)
    Dim dTglLahir As Date
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If txtBulan.Text = "" And txtTahun.Text = "" Then txtHari.SetFocus: Exit Sub
        If txtBulan.Text = "" Then txtBulan.Text = 0
        If txtTahun.Text = "" And txtHari.Text = "" Then
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
        ElseIf txtTahun.Text <> "" And txtHari.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        ElseIf txtTahun.Text = "" And txtHari.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        ElseIf txtTahun.Text <> "" And txtHari.Text = "" Then
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        End If
        meTglLahir.Text = dTglLahir
        txtHari.SetFocus
    End If
End Sub

Private Sub txtCariNamaPegawai_Change()
    Call subLoadGrid
End Sub

Private Sub txtCariNamaPegawai_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtHari_Change()
    Dim dTglLahir As Date
    If txtHari.Text = "" And txtBulan.Text = "" And txtTahun.Text = "" Then txtketerangan.SetFocus: Exit Sub
    If txtHari.Text = "" Then txtHari.Text = 0
    If txtTahun.Text = "" And txtBulan.Text = "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
    ElseIf txtTahun.Text <> "" And txtBulan.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    ElseIf txtTahun.Text = "" And txtBulan.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
    ElseIf txtTahun.Text <> "" And txtBulan.Text = "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    End If
    meTglLahir.Text = dTglLahir
End Sub

Private Sub txtHari_KeyPress(KeyAscii As Integer)
    Dim dTglLahir As Date
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If txtHari.Text = "" And txtBulan.Text = "" And txtTahun.Text = "" Then txtketerangan.SetFocus: Exit Sub
        If txtHari.Text = "" Then txtHari.Text = 0
        If txtTahun.Text = "" And txtBulan.Text = "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        ElseIf txtTahun.Text <> "" And txtBulan.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        ElseIf txtTahun.Text = "" And txtBulan.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        ElseIf txtTahun.Text <> "" And txtBulan.Text = "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        End If
        meTglLahir.Text = dTglLahir
        txtketerangan.SetFocus
    End If
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaKeluarga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboJnsKelamin.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtNamaKeluarga_LostFocus()
    txtNamaKeluarga.Text = StrConv(txtNamaKeluarga.Text, vbProperCase)
End Sub

Private Sub txtNamaPegawai_Change()
    vTampil = True
    Call subLoadDataPegawai
End Sub

Private Sub txtNamaPegawai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If dgDataPegawai.Visible = False Then Exit Sub
        dgDataPegawai.SetFocus
    End If
End Sub

Private Sub txtNamaPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dgDataPegawai.Visible = True Then
            dgDataPegawai.SetFocus
        Else
            dcHubungan.SetFocus
        End If
    End If
    If KeyAscii = 27 Then dgDataPegawai.Visible = False
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtTahun_Change()
    Dim dTglLahir As Date
    If txtTahun = "" Then txtBulan.SetFocus: Exit Sub
    If txtBulan.Text = "" And txtHari.Text = "" Then
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), Date)
    ElseIf txtBulan.Text <> "" And txtHari.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    ElseIf txtBulan.Text = "" And txtHari.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    ElseIf txtBulan.Text <> "" And txtHari.Text = "" Then
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    End If
    meTglLahir.Text = dTglLahir
End Sub

Private Sub txtTahun_KeyPress(KeyAscii As Integer)
    Dim dTglLahir As Date
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If txtTahun = "" Then txtBulan.SetFocus: Exit Sub
        If txtBulan.Text = "" And txtHari.Text = "" Then
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), Date)
        ElseIf txtBulan.Text <> "" And txtHari.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        ElseIf txtBulan.Text = "" And txtHari.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        ElseIf txtBulan.Text <> "" And txtHari.Text = "" Then
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        End If
        meTglLahir.Text = dTglLahir
        txtBulan.SetFocus
    End If
End Sub

Private Sub txtTptLhr_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then txtketerangan.SetFocus
End Sub

