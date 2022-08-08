VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRegistrasiRJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Registrasi Rawat Jalan"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegistrasiRJ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   10335
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   40
      Top             =   4575
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   18177
            Text            =   "Daftar Baru (Ctrl+N)"
            TextSave        =   "Daftar Baru (Ctrl+N)"
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
   Begin VB.Frame fraDokter 
      Caption         =   "Data Dokter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2400
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   7815
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   1455
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   16
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   28
      Top             =   3840
      Width           =   10335
      Begin VB.CommandButton cmTutup2 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   8400
         TabIndex        =   39
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdRujukan 
         Caption         =   "&Data Rujukan"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   4800
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Lanjutkan"
         Height          =   375
         Left            =   6600
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraDataRegistrasiRJ 
      Caption         =   "Data Registrasi Rawat Jalan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   20
      Top             =   2040
      Width           =   10335
      Begin VB.TextBox txtTglMasuk 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtKelompok 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtRuangan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   7440
         TabIndex        =   11
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtKelas 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   5040
         TabIndex        =   10
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtJenisKelas 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2280
         TabIndex        =   9
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6120
         TabIndex        =   14
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox txtKdDokter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7440
         TabIndex        =   36
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSDataListLib.DataCombo dcRujukanAsal 
         Height          =   330
         Left            =   3000
         TabIndex        =   13
         Top             =   1320
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Rujukan Dari"
         Height          =   210
         Left            =   3000
         TabIndex        =   38
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelas Pelayanan"
         Height          =   210
         Left            =   2280
         TabIndex        =   37
         Top             =   360
         Width           =   1725
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kelompok Pasien"
         Height          =   210
         Left            =   240
         TabIndex        =   30
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pendaftaran"
         Height          =   210
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Pelayanan"
         Height          =   210
         Left            =   5040
         TabIndex        =   23
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Pemeriksa"
         Height          =   210
         Left            =   6120
         TabIndex        =   22
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Ruang Pemeriksaan"
         Height          =   210
         Left            =   7440
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   24
      Top             =   960
      Width           =   10335
      Begin VB.CheckBox chkDetailPasien 
         Caption         =   "Detail Pasien"
         Enabled         =   0   'False
         Height          =   255
         Left            =   8880
         TabIndex        =   7
         Top             =   630
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "Umur"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6240
         TabIndex        =   31
         Top             =   320
         Width           =   2535
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   6
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            MaxLength       =   6
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            MaxLength       =   6
            TabIndex        =   4
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Left            =   2280
            TabIndex        =   34
            Top             =   292
            Width           =   165
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Left            =   1440
            TabIndex        =   33
            Top             =   292
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Left            =   600
            TabIndex        =   32
            Top             =   292
            Width           =   285
         End
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5040
         MaxLength       =   9
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   5040
         TabIndex        =   35
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No. Registrasi"
         Height          =   210
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1800
         TabIndex        =   26
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   2760
         TabIndex        =   25
         Top             =   360
         Width           =   1020
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   41
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
      Picture         =   "frmRegistrasiRJ.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRegistrasiRJ.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRegistrasiRJ.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmRegistrasiRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilter As String
Dim intRowNow As Integer
Dim strSubInstalasi As String
Dim strNoAntrian As String

Private Sub chkDetailPasien_Click()
    If chkDetailPasien.value = 1 Then
        strPasien = "View"
        Load frmPasienBaru
        frmPasienBaru.Show
    Else
        Unload frmPasienBaru
        Unload frmDetailPasien
    End If
End Sub

Private Sub cmdRujukan_Click()
    With frmRujukan
        .Show
        .txtNoCM.Text = txtNoCM
        .txtNamaPasien.Text = txtNamaPasien.Text
        .txtJK.Text = txtJK.Text
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHr.Text
        .txtnopendaftaran.Text = txtnopendaftaran.Text
    End With
End Sub

Private Sub cmdSimpan_Click()
    If funcCekValidasi = False Then Exit Sub
    Call sp_RegistrasiRJ(dbcmd)
    Call subEnableButtonReg(True)
End Sub

Private Sub cmdTutup_Click()
    
    'If Periksa("datacombo", dcRujukanAsal, "Data rujukan asal kosong") = False Then Exit Sub
    'If Periksa("text", txtDokter, "Nama dokter kosong") = False Then Exit Sub
    
    If dcRujukanAsal.Text = "" Then
         MsgBox "Rujukan Asal Harus Diisi", vbExclamation, "Validasi"
         Exit Sub
    End If
    
    If txtDokter.Text = "" Then
         MsgBox "Dokter Harus Diisi", vbExclamation, "Validasi"
         Exit Sub
    End If
    
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan data registrasi Rawat Jalan?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
        Else
            Exit Sub
        End If
    End If

    Call subLoadFormTP
    Unload Me
End Sub

Private Sub subLoadFormTP()
    On Error GoTo hell
    mstrNoPen = txtnopendaftaran.Text
    mstrNoCM = txtNoCM.Text
    mstrKdSubInstalasi = frmDaftarAntrianPasien.dgDaftarAntrianPasien.Columns("KdSubInstalasi").value
    With frmTransaksiPasien
        .Show
        .txtnopendaftaran.Text = mstrNoPen
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = txtNamaPasien
        .txtSex.Text = txtJK.Text
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHr.Text
        .txtKls.Text = txtKelas.Text
        .txtJenisPasien.Text = txtKelompok.Text

        .txtTglDaftar.Text = TxtTglMasuk.Text
    End With
hell:
End Sub

Private Sub cmTutup2_Click()
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan data registrasi Rawat Jalan", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub dcRujukanAsal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcRujukanAsal.MatchedWithList = True Then txtDokter.SetFocus
        strSQL = "SELECT KdRujukanAsal,RujukanAsal FROM RujukanAsal where KdRujukanAsal<>'09' and StatusEnabled='1' and (RujukanAsal LIKE '%" & dcRujukanAsal.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcRujukanAsal.BoundText = rs(0).value
        dcRujukanAsal.Text = rs(1).value
    End If
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        txtDokter.Text = dgDokter.Columns("Nama Dokter").value
        txtKdDokter.Text = dgDokter.Columns("Kode Dokter").value
        If txtKdDokter.Text = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        fraDokter.Visible = False
        Me.Height = 5370
        Call centerForm(Me, MDIUtama)
    End If
End Sub

Private Sub dtpTglPendaftaran_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcRujukanAsal.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    Select Case KeyCode
        Case vbKeyN
            If strCtrlKey = 4 Then Unload Me: frmRegistrasiRJ.Show
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo Errload
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    strRegistrasi = "RJ"
    If mblnCariPasien = True Then frmCariPasien.Enabled = False

    strSQL = "SELECT KdRujukanAsal,RujukanAsal FROM RujukanAsal where KdRujukanAsal<>'09' and StatusEnabled='1'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcRujukanAsal.RowSource = rs
    dcRujukanAsal.ListField = rs.Fields(1).Name
    dcRujukanAsal.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
    Call subTampilData(txtnopendaftaran)
    Exit Sub
Errload:
    msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnCariPasien = True Then frmCariPasien.Enabled = True
    frmDaftarAntrianPasien.Enabled = True
End Sub

Private Sub txtDokter_Change()
    strFilter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
    txtKdDokter.Text = ""
    Call subLoadDokter
    fraDokter.Left = 2280
    fraDokter.Top = fraDataRegistrasiRJ.Top + txtDokter.Top + txtDokter.Height
    fraDokter.Visible = True
    Me.Height = 6200
    Call centerForm(Me, MDIUtama)
End Sub

Private Sub txtDokter_GotFocus()
    Call txtDokter_Change
End Sub

Private Sub txtDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If fraDokter.Visible = False Then Exit Sub
        dgDokter.SetFocus
    End If
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
    On Error GoTo Errload
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        dgDokter.SetFocus
    End If
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 27 Then
        fraDokter.Visible = False
        Me.Height = 5385
    End If
    Call SetKeyPressToChar(KeyAscii)
Errload:
End Sub

Private Sub txtNoCM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkDetailPasien.Enabled = True Then chkDetailPasien.SetFocus
    End If
End Sub

Private Sub txtNoPendaftaran_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then Call subTampilData(txtnopendaftaran)
End Sub

Public Sub subTampilData(strNoPenndaftaran As String)
    On Error GoTo Errload
    Call subClearData
    Call subEnableButtonReg(False)
    strSQL = "Select * from V_DaftarAntrianPasienMRS WHERE NoPendaftaran ='" & strNoPenndaftaran & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        mstrNoCM = ""
        mstrNoPen = ""
        chkDetailPasien.Enabled = False
        cmdSimpan.Enabled = False
        Exit Sub
    End If
    txtNoCM.Text = rs("NoCM")
    mstrNoCM = txtNoCM.Text
    txtNamaPasien.Text = rs.Fields("Nama Pasien").value
    If rs.Fields("JK").value = "P" Then
        txtJK.Text = "Perempuan"
    ElseIf rs.Fields("JK").value = "L" Then
        txtJK.Text = "Laki-laki"
    End If
    txtThn.Text = rs.Fields("UmurTahun").value
    txtBln.Text = rs.Fields("UmurBulan").value
    txtHr.Text = rs.Fields("UmurHari").value

    TxtTglMasuk.Text = rs("TglMasuk")
    txtJenisKelas.Text = ""
    txtKelas.Text = rs("Kelas")
    txtRuangan.Text = rs("Ruangan")
    dcRujukanAsal.BoundText = ""
    txtKelompok.Text = ""
    txtDokter.Text = ""

    mdTglMasuk = TxtTglMasuk.Text
    mstrKdKelas = rs("KdKelas")
    mstrKelas = rs("Kelas")

    Set rs = Nothing

    strSQL = "SELECT dbo.PasienDaftar.KdKelompokPasien, dbo.KelompokPasien.JenisPasien, dbo.DetailJenisJasaPelayanan.DetailJenisJasaPelayanan, dbo.JenisJasaPelayanan.JenisJasaPelayanan, dbo.PasienDaftar.NoPendaftaran" & _
    " FROM dbo.PasienDaftar INNER JOIN dbo.DetailJenisJasaPelayanan ON dbo.PasienDaftar.KdDetailJenisJasaPelayanan = dbo.DetailJenisJasaPelayanan.KdDetailJenisJasaPelayanan INNER JOIN dbo.KelompokPasien ON dbo.PasienDaftar.KdKelompokPasien = dbo.KelompokPasien.KdKelompokPasien INNER JOIN dbo.JenisJasaPelayanan ON dbo.DetailJenisJasaPelayanan.KdJenisJasaPelayanan = dbo.JenisJasaPelayanan.KdJenisJasaPelayanan" & _
    " WHERE (dbo.PasienDaftar.NoPendaftaran = '" & txtnopendaftaran & "')"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        txtJenisKelas.Text = rs("JenisJasaPelayanan")
        txtKelompok.Text = rs("JenisPasien")
    End If

    chkDetailPasien.Enabled = True
    strSQL = "SELECT KdRujukanAsal FROM RegistrasiRJ WHERE (NoPendaftaran = '" & strNoPenndaftaran & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then dcRujukanAsal.BoundText = rs(0) Else dcRujukanAsal.BoundText = "01"
    txtDokter.SetFocus

    Exit Sub
Errload:
    Call msubPesanError
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter()
    On Error Resume Next
    strSQL = "SELECT NamaDokter AS [Nama Dokter],KodeDokter AS [Kode Dokter],JK,Jabatan FROM V_DaftarDokter " & strFilter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDokter = rs.RecordCount
    Set dgDokter.DataSource = rs
    With dgDokter
        .Columns(0).Width = 3000 'nama dokter
        .Columns(1).Width = 0 'kode dokter
        .Columns(2).Width = 400
        .Columns(3).Width = 3300
    End With
End Sub

'untuk enable/disable button reg
Private Sub subEnableButtonReg(blnStatus As Boolean)
    cmdRujukan.Enabled = blnStatus
    cmdSimpan.Enabled = Not blnStatus
    txtDokter.Enabled = Not blnStatus
End Sub

'Store procedure untuk mengisi registrasi pasien
Private Sub sp_RegistrasiRJ(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuanganPasien)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(TxtTglMasuk, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, txtKdDokter.Text)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, noidpegawai)

        .ActiveConnection = dbConn
        .CommandText = "Add_RegistrasiPasienMasukRJ"
        .CommandType = adCmdStoredProc
        .Execute
        Call openConnection
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Pendaftaran Pasien ke Instalasi Rawat Jalan", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Add_RegistrasiPasienMasukRJ")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'untuk cek validasi
Private Function funcCekValidasi() As Boolean
    If txtNamaPasien.Text = "" Then
        MsgBox "No. CM Harus Diisi", vbExclamation, "Validasi"
        funcCekValidasi = False
        txtNoCM.SetFocus
        Exit Function
    End If
    If txtKdDokter.Text = "" Then
        MsgBox "Pilihan Dokter harus diisi sesuai data daftar dokter", vbExclamation, "Validasi"
        funcCekValidasi = False
        txtDokter.SetFocus
        Exit Function
    End If
    funcCekValidasi = True
End Function

'untuk membersihkan data pasien registrasi
Private Sub subClearData()
    txtNoCM.Text = ""
    txtNamaPasien.Text = ""
    txtJK.Text = ""
    txtThn.Text = ""
    txtBln.Text = ""
    txtHr.Text = ""
    dcRujukanAsal.Text = ""
    txtDokter.Text = ""
    txtKdDokter.Text = ""
End Sub

