VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDaftarPasienRJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Poliklinik"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienRJ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   14295
   Begin VB.Frame fraCari 
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
      Height          =   840
      Left            =   0
      TabIndex        =   8
      Top             =   7440
      Width           =   14295
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
         Left            =   8400
         TabIndex        =   5
         ToolTipText     =   "Perbaiki data pasien"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdTP 
         Caption         =   "Pemeriksaan Pasien"
         Height          =   495
         Left            =   10080
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   12150
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   440
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukan Nama Pasien /  No.CM / Ruangan"
         Height          =   210
         Left            =   1560
         TabIndex        =   11
         Top             =   195
         Width           =   3450
      End
   End
   Begin VB.Frame fraDaftar 
      Caption         =   "Daftar Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   14295
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
         Left            =   8400
         TabIndex        =   10
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
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   48037891
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   48037891
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   12
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgDaftarPasienRJ 
         Height          =   5415
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   9551
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image Image2 
      Height          =   930
      Left            =   4090
      Picture         =   "frmDaftarPasienRJ.frx":08CA
      Top             =   0
      Width           =   10200
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   -2280
      Picture         =   "frmDaftarPasienRJ.frx":6012
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "frmDaftarPasienRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dTglMasuk As Date

Public Sub cmdcari_Click()
On Error GoTo errLoad
    Set rs = Nothing
    rs.Open "select * from V_DaftarPasienLamaRJ where ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR Ruangan like '%" & txtParameter.Text & "%') and TglMasuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'", dbConn, adOpenStatic, adLockOptimistic
    Set dgDaftarPasienRJ.DataSource = rs
    Call SetGridPasienRJ
    If dgDaftarPasienRJ.ApproxCount > 0 Then
        dgDaftarPasienRJ.SetFocus
    Else
        dtpAwal.SetFocus
    End If
errLoad:
End Sub

Private Sub cmdDataPasien_Click()
On Error GoTo hell
    strPasien = "View"
    mstrNoCM = dgDaftarPasienRJ.Columns(1).Value
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

Private Sub dgDaftarPasienRJ_DblClick()
    Call cmdTP_Click
End Sub

Private Sub dgDaftarPasienRJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTP.SetFocus
End Sub

Private Sub dgDaftarPasienRJ_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo hell
    mstrKdRuangan = dgDaftarPasienRJ.Columns(18).Value
hell:
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
    Set rs = Nothing
    strQuery = "select * from V_DaftarPasienLamaRJ where TglMasuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'"
    rs.Open strQuery, dbConn, adOpenStatic, adLockOptimistic
    Set dgDaftarPasienRJ.DataSource = rs
    Call SetGridPasienRJ
    mblnForm = True
End Sub

Sub SetGridPasienRJ()
 With dgDaftarPasienRJ
  .Columns(0).Width = 1150 'NoPendaftaran
  .Columns(0).Caption = "No. Registrasi"
  .Columns(1).Width = 750 'NoCM
  .Columns(1).Alignment = dbgCenter
  .Columns(2).Width = 2000 '[Nama Pasien]
  .Columns(3).Width = 300 'JK
  .Columns(4).Width = 1400 'Umur
  .Columns(5).Width = 0 'Poliklinik
  .Columns(6).Width = 2200 'Ruangan
  .Columns(7).Width = 1200 'JenisPasien
  .Columns(8).Width = 0 'Kelas
  .Columns(9).Width = 1900 'TglMasuk
  .Columns(10).Width = 2600 ' [Dokter Pemeriksa]
  .Columns(11).Width = 0 ' [No. Urut]
  .Columns(12).Width = 0 'UmurTahun
  .Columns(13).Width = 0 'UmurBulan
  .Columns(14).Width = 0 'UmurHari
  .Columns(15).Width = 0 'KdJenisTarif
  .Columns(16).Width = 0 'KdKelas
  .Columns(17).Width = 0 'KdSubInstalasi
  .Columns(18).Width = 0 'KdRuangan
  .Columns(19).Width = 5500
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnForm = False
End Sub

Private Sub txtParameter_Change()
    Call cmdcari_Click
    txtParameter.SetFocus
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call cmdcari_Click
        If dgDaftarPasienRJ.ApproxCount > 0 Then
            dgDaftarPasienRJ.SetFocus
        Else
            txtParameter.SetFocus
        End If
    End If
End Sub

Private Sub subLoadFormTP()
On Error GoTo hell
    mstrNoPen = dgDaftarPasienRJ.Columns(0).Value
    mstrNoPen = dgDaftarPasienRJ.Columns(0).Value
    mstrNoCM = dgDaftarPasienRJ.Columns(1).Value
    mstrNoCM = dgDaftarPasienRJ.Columns(1).Value
    
    
    With frmTransaksiPasien
        .Show
        .txtNoPendaftaran.Text = dgDaftarPasienRJ.Columns(0).Value
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = dgDaftarPasienRJ.Columns(2).Value
        If dgDaftarPasienRJ.Columns(3).Value = "L" Then
            .txtSex.Text = "Laki-Laki"
        Else
            .txtSex.Text = "Perempuan"
        End If
        .txtThn.Text = dgDaftarPasienRJ.Columns(12).Value
        .txtBln.Text = dgDaftarPasienRJ.Columns(13).Value
        .txtHr.Text = dgDaftarPasienRJ.Columns(14).Value
        .txtKls.Text = dgDaftarPasienRJ.Columns("Kelas").Value
        .txtJenisPasien.Text = dgDaftarPasienRJ.Columns("JenisPasien").Value
        .txtTglDaftar.Text = dgDaftarPasienRJ.Columns(9).Value
         mdTglMasuk = dgDaftarPasienRJ.Columns(9).Value
         mstrKelas = dgDaftarPasienRJ.Columns(8).Value
         mstrKdRuangan = dgDaftarPasienRJ.Columns(18).Value
         mstrKdSubInstalasi = dgDaftarPasienRJ.Columns(17).Value
   End With
    
    strSQL = "SELECT dbo.RegistrasiRJ.IdDokter, dbo.DataPegawai.NamaLengkap " & _
        " FROM dbo.RegistrasiRJ INNER JOIN dbo.DataPegawai ON dbo.RegistrasiRJ.IdDokter = dbo.DataPegawai.IdPegawai " & _
        " WHERE (dbo.RegistrasiRJ.NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    
    mstrKdDokter = rs(0).Value
    mstrNamaDokter = rs(1).Value
    
hell:
End Sub
