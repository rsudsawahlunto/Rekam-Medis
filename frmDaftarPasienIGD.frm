VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDaftarPasienIGD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Gawat Darurat"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienIGD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   14790
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
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   7800
      Width           =   14775
      Begin VB.CommandButton cmdTP 
         Caption         =   "Pemeriksaan Pasien"
         Height          =   495
         Left            =   10920
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdDataPasien 
         Caption         =   "&Data Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9000
         TabIndex        =   7
         ToolTipText     =   "Perbaiki data pasien"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   450
         Width           =   2655
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   12840
         TabIndex        =   9
         Top             =   260
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan  Nama Pasien / No.CM"
         Height          =   210
         Left            =   1560
         TabIndex        =   12
         Top             =   195
         Width           =   2640
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Daftar Pasien Gawat Darurat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   14775
      Begin VB.Frame Frame4 
         Caption         =   "Status Keluar"
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
         Left            =   6720
         TabIndex        =   15
         Top             =   150
         Width           =   2055
         Begin VB.CheckBox chkBelum 
            Caption         =   "Belum"
            Height          =   255
            Left            =   120
            TabIndex        =   0
            Top             =   285
            Width           =   855
         End
         Begin VB.CheckBox chkSudah 
            Caption         =   "Sudah"
            Height          =   255
            Left            =   1080
            TabIndex        =   1
            Top             =   285
            Width           =   855
         End
      End
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
         Left            =   8880
         TabIndex        =   13
         Top             =   150
         Width           =   5775
         Begin VB.CommandButton cmdCari 
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
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   2
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   22675459
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   3
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   22675459
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   14
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgDaftarPasienGD 
         Height          =   5775
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   10186
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
   Begin VB.Image Image1 
      Height          =   930
      Left            =   4580
      Picture         =   "frmDaftarPasienIGD.frx":08CA
      Top             =   0
      Width           =   10200
   End
   Begin VB.Image Image2 
      Height          =   930
      Left            =   0
      Picture         =   "frmDaftarPasienIGD.frx":6012
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "frmDaftarPasienIGD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkBelum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAwal.SetFocus
End Sub

Private Sub chkSudah_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAwal.SetFocus
End Sub

Private Sub cmdcari_Click()
Dim dTglKeluar As Date
On Error GoTo hell
    Set rs = Nothing
    If (chkBelum.Value = 1) And (chkSudah.Value = 0) Then
       rs.Open "select * from V_DaftarPasienIGDAll where([Nama Pasien] like '%" & txtParameter.Text & "%' OR [No. CM] like '%" & txtParameter.Text & "%') AND TglMasuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND [Status Keluar]='Belum'", dbConn, adOpenStatic, adLockOptimistic
    ElseIf (chkSudah.Value = 1) And (chkBelum.Value = 0) Then
       rs.Open "select * from V_DaftarPasienIGDAll where([Nama Pasien] like '%" & txtParameter.Text & "%' OR [No. CM] like '%" & txtParameter.Text & "%') AND TglMasuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND [Status Keluar]='Sudah'", dbConn, adOpenStatic, adLockOptimistic
    Else
       rs.Open "select * from V_DaftarPasienIGDAll where([Nama Pasien] like '%" & txtParameter.Text & "%' OR [No. CM] like '%" & txtParameter.Text & "%') AND TglMasuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'", dbConn, adOpenStatic, adLockOptimistic
    End If
    Set dgDaftarPasienGD.DataSource = rs
    Call SetGridPasienGD
    If dgDaftarPasienGD.ApproxCount = 0 Then
       dtpAwal.SetFocus
    Else
       dgDaftarPasienGD.SetFocus
    End If
hell:
End Sub

Sub SetGridPasienGD()
 With dgDaftarPasienGD
  .Columns(0).Width = 1150
  .Columns(0).Caption = "No. Registrasi"
  .Columns(1).Width = 800
  .Columns(1).Caption = "No. CM"
  .Columns(1).Alignment = dbgCenter
  .Columns(2).Width = 2000
  .Columns(3).Width = 400
  .Columns(4).Width = 1400
  .Columns(5).Width = 1200
  .Columns(6).Width = 1590
  .Columns(6).Caption = "Tgl. Masuk"
  .Columns(7).Width = 1200
  .Columns(7).Alignment = dbgCenter
  .Columns(8).Caption = "Tgl. Keluar"
  .Columns(8).Width = 1590
  .Columns(9).Width = 1400
  .Columns(10).Width = 1500
  .Columns(11).Width = 2500
  .Columns(12).Width = 700
  .Columns(13).Width = 1800
  .Columns(14).Width = 1800
  .Columns(15).Width = 1800
 End With
End Sub

Private Sub cmdDataPasien_Click()
On Error GoTo hell
    strPasien = "View"
    mstrNoCM = dgDaftarPasienGD.Columns(1).Value
    frmPasienBaru.Show
hell:
End Sub

Private Sub cmdTP_Click()
On Error GoTo hell
    Call subLoadFormTP
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgDaftarPasienGD_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then cmdTP.SetFocus
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    dtpAwal.Value = Now
    dtpAkhir.Value = Now
    chkBelum.Value = 1
    chkSudah.Value = 1
    Call cmdcari_Click
End Sub

Private Sub txtParameter_Change()
 Call cmdcari_Click
 txtParameter.SetFocus
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtParameter_LostFocus()
    txtParameter.Text = StrConv(txtParameter.Text, vbProperCase)
End Sub

'Store procedure untuk mengupdate status pasien dari non aktif ke aktif
Private Sub sp_UpdateNonAktifKeAktif(ByVal adoCommand As ADODB.Command, mstrNoPenda As String)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPenda)
        
        .ActiveConnection = dbConn
        .CommandText = "Update_PasienNonAktifKeAktif"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam update status pasien dari nonaktif ke aktif", vbCritical, "Validasi"
        Else
            MsgBox "Update status pasien dari nonaktif ke aktif Sukses", vbExclamation, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

Private Sub subLoadFormTP()
On Error GoTo hell
    mstrNoPen = dgDaftarPasienGD.Columns(0).Value
    mstrNoCM = dgDaftarPasienGD.Columns(1).Value
    mstrNoPen = mstrNoPen
    mstrNoCM = mstrNoCM
    
    With frmTransaksiPasien
        .Show
        .txtNoPendaftaran.Text = mstrNoPen
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = dgDaftarPasienGD.Columns(2).Value
        If dgDaftarPasienGD.Columns(3).Value = "L" Then
            .txtSex.Text = "Laki-Laki"
        Else
            .txtSex.Text = "Perempuan"
        End If
        .txtThn.Text = dgDaftarPasienGD.Columns(4).Value
        .txtBln.Text = dgDaftarPasienGD.Columns(5).Value
        .txtHr.Text = dgDaftarPasienGD.Columns(16).Value
        .txtKls.Text = dgDaftarPasienGD.Columns("Kelas").Value
        .txtJenisPasien.Text = dgDaftarPasienGD.Columns("JenisPasien").Value
        .txtTglDaftar.Text = dgDaftarPasienGD.Columns(6).Value
         mdTglMasuk = dgDaftarPasienGD.Columns(6).Value
         mstrKelas = dgDaftarPasienGD.Columns("Kelas").Value
         mstrKdRuangan = dgDaftarPasienGD.Columns(19).Value
         mstrKdSubInstalasi = dgDaftarPasienGD.Columns(20).Value
    End With
    
    strSQL = "SELECT dbo.RegistrasiIGD.IdDokter, dbo.DataPegawai.NamaLengkap " & _
        " FROM dbo.RegistrasiIGD INNER JOIN dbo.DataPegawai ON dbo.RegistrasiIGD.IdDokter = dbo.DataPegawai.IdPegawai " & _
        " WHERE (dbo.RegistrasiIGD.NoPendaftaran = '" & txtNoPendaftaran.Text & "')"
    Call msubRecFO(rs, strSQL)
    
    mstrKdDokter = rs(0).Value
    mstrNamaDokter = rs(1).Value
    
hell:
End Sub


