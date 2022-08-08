VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDaftarPasienRI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Rawat Inap"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienRI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   13350
   Begin VB.Frame fraDokterP 
      Caption         =   "Setting Dokter Pemeriksa"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   360
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   12615
      Begin VB.Frame Frame5 
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
         Height          =   975
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   12135
         Begin VB.Frame Frame6 
            Caption         =   "Umur"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   8640
            TabIndex        =   35
            Top             =   240
            Width           =   2775
            Begin VB.TextBox txtHr 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   1920
               MaxLength       =   6
               TabIndex        =   38
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtBln 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   1080
               MaxLength       =   6
               TabIndex        =   37
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtThn 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   240
               MaxLength       =   6
               TabIndex        =   36
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "hr"
               Height          =   210
               Left            =   2400
               TabIndex        =   41
               Top             =   285
               Width           =   165
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "bln"
               Height          =   210
               Left            =   1560
               TabIndex        =   40
               Top             =   285
               Width           =   240
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "thn"
               Height          =   210
               Left            =   720
               TabIndex        =   39
               Top             =   285
               Width           =   285
            End
         End
         Begin VB.TextBox txtJK 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   6480
            MaxLength       =   9
            TabIndex        =   34
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtNoPendaftaran 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   240
            MaxLength       =   10
            TabIndex        =   33
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtNoCM 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   32
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtNamaPasien 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   3240
            MaxLength       =   50
            TabIndex        =   31
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblJnsKlm 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Kelamin"
            Height          =   210
            Left            =   6480
            TabIndex        =   45
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "No. Pendaftaran"
            Height          =   210
            Left            =   240
            TabIndex        =   44
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "No. CM"
            Height          =   210
            Left            =   1920
            TabIndex        =   43
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lblNamaPasien 
            AutoSize        =   -1  'True
            Caption         =   "Nama Pasien"
            Height          =   210
            Left            =   3240
            TabIndex        =   42
            Top             =   240
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdSimpanDokter 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   8760
         TabIndex        =   29
         Top             =   5400
         Width           =   1815
      End
      Begin VB.CommandButton cmdBatalDokter 
         Caption         =   "&Tutup"
         Height          =   375
         Left            =   10680
         TabIndex        =   28
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Caption         =   "Data Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   12135
         Begin VB.TextBox txtPrevDokter 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   5040
            TabIndex        =   23
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox txtTglPeriksa 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2760
            TabIndex        =   22
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtDokter 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   8520
            TabIndex        =   21
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox txtPoli 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   240
            TabIndex        =   20
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Dokter Pemeriksa Sebelumnya"
            Height          =   210
            Left            =   5040
            TabIndex        =   27
            Top             =   360
            Width           =   2475
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Tanggal Pemeriksaan"
            Height          =   210
            Left            =   2760
            TabIndex        =   26
            Top             =   360
            Width           =   1710
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Dokter Pemeriksa Sekarang"
            Height          =   210
            Left            =   8520
            TabIndex        =   25
            Top             =   360
            Width           =   2235
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ruang Pemeriksaan"
            Height          =   210
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   1575
         End
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
         Height          =   2655
         Left            =   240
         TabIndex        =   17
         Top             =   2520
         Width           =   12135
         Begin MSDataGridLib.DataGrid dgDokter 
            Height          =   2055
            Left            =   2040
            TabIndex        =   18
            Top             =   360
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   3625
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
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
   End
   Begin VB.Frame fraCari 
      Height          =   840
      Left            =   0
      TabIndex        =   9
      Top             =   7080
      Width           =   13335
      Begin VB.CommandButton cmdPasienDirujuk 
         Caption         =   "Konsul ke Unit Lain"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4200
         TabIndex        =   7
         Top             =   250
         Width           =   1575
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   420
         Width           =   2295
      End
      Begin VB.CommandButton cmdPasienPulang 
         Appearance      =   0  'Flat
         Caption         =   "Pasien Pu&lang"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   10560
         TabIndex        =   6
         Top             =   250
         Width           =   1335
      End
      Begin VB.CommandButton cmdTP 
         Caption         =   "Transaksi Pela&yanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5880
         TabIndex        =   2
         Top             =   250
         Width           =   1695
      End
      Begin VB.CommandButton cmdDiagnosaPasien 
         Appearance      =   0  'Flat
         Caption         =   "&Diagnosa Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6120
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdMasukKamar 
         Appearance      =   0  'Flat
         Caption         =   "&Masuk Kamar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7680
         TabIndex        =   4
         Top             =   250
         Width           =   1335
      End
      Begin VB.CommandButton cmdKeluarKamar 
         Appearance      =   0  'Flat
         Caption         =   "&Keluar Kamar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9120
         TabIndex        =   5
         Top             =   250
         Width           =   1335
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   12000
         TabIndex        =   8
         Top             =   250
         Width           =   1215
      End
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
         Left            =   2520
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Input Nama Pasien /  No.CM"
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   165
         Width           =   2310
      End
   End
   Begin VB.Frame fraDaftar 
      Caption         =   "Daftar Pasien Rawat Inap"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      TabIndex        =   10
      Top             =   1680
      Width           =   13335
      Begin MSDataGridLib.DataGrid dgDaftarPasienRI 
         Height          =   4815
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   8493
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
   Begin VB.Frame fraPilih 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   0
      TabIndex        =   13
      Top             =   960
      Width           =   13335
      Begin VB.OptionButton optPasNonAktif 
         Caption         =   "Daftar Pasien Non Aktif"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   15
         Top             =   200
         Width           =   3735
      End
      Begin VB.OptionButton optPasAktif 
         Caption         =   "Daftar Pasien Aktif"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   200
         Width           =   6615
      End
   End
   Begin VB.Image Image2 
      Height          =   930
      Left            =   3130
      Picture         =   "frmDaftarPasienRI.frx":08CA
      Top             =   0
      Width           =   10200
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   -2640
      Picture         =   "frmDaftarPasienRI.frx":6012
      Top             =   0
      Width           =   10200
   End
   Begin VB.Menu mnuinfo 
      Caption         =   "&Informasi"
      Begin VB.Menu mnurbs 
         Caption         =   "Rincian Biaya Sementara"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnusepcmp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuccmp 
         Caption         =   "Cetak Catatan Medik Pasien"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuSetting 
      Caption         =   "&Setting"
      Begin VB.Menu mnuSDokter 
         Caption         =   "Dokter Penanggung Jawab"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnusepubahdatapas 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSEditDPasien 
         Caption         =   "Ubah Data Pasien"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmDaftarPasienRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFilterDokter As String
Dim strKdDokter As String
Dim intJmlDokter As Integer
Dim dTglMasuk As Date

Private Sub cmdBatalDokter_Click()
    fraDokterP.Visible = False
    fraDokterP.Enabled = False
    fraPilih.Enabled = True
    fraDaftar.Enabled = True
    fraCari.Enabled = True
End Sub

Public Sub cmdcari_Click()
    If optPasAktif.Value = True Then
        Call optPasAktif_GotFocus
    Else
        Call optPasNonAktif_GotFocus
    End If
End Sub

Private Sub cmdDiagnosaPasien_Click()
On Error GoTo hell
If dgDaftarPasienRI.Columns(0).Value = "" Then
   Exit Sub
End If
    frmDaftarPasienRI.Enabled = False
    Call subLoadFormPeriksaDiagnosa
hell:
End Sub

Private Sub cmdKeluarKamar_Click()
On Error GoTo hell
If dgDaftarPasienRI.Columns(0).Value = "" Then
   Exit Sub
End If
    frmDaftarPasienRI.Enabled = False
    Call subLoadFormKeluarKam
hell:
End Sub

Private Sub cmdMasukKamar_Click()
On Error GoTo hell
If dgDaftarPasienRI.Columns(0).Value = "" Then
   Exit Sub
End If
    frmDaftarPasienRI.Enabled = False
    Call subLoadFormMasukKam
hell:
End Sub

Private Sub cmdPasienDirujuk_Click()
On Error GoTo hell
If dgDaftarPasienRI.Columns(0).Value = "" Then
   Exit Sub
End If
frmPasienRujukan.Show
With frmPasienRujukan
    .txtNoPendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
    .txtNoCM.Text = dgDaftarPasienRI.Columns(1).Value
    .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
    If dgDaftarPasienRI.Columns(3).Value = "P" Then
        .txtSex.Text = "Perempuan"
    Else
        .txtSex.Text = "Laki-Laki"
    End If
    .txtThn.Text = dgDaftarPasienRI.Columns(11).Value
    .txtBln.Text = dgDaftarPasienRI.Columns(12).Value
    .txtHari.Text = dgDaftarPasienRI.Columns(13).Value
    .dtpTglDirujuk.Value = Now
End With
frmDaftarPasienRI.Enabled = False
hell:
End Sub

Private Sub cmdPasienPulang_Click()
On Error GoTo hell
If dgDaftarPasienRI.Columns(0).Value = "" Then
   Exit Sub
End If
    frmDaftarPasienRI.Enabled = False
    Call subLoadFormPsnPulang
hell:
End Sub

Private Sub cmdSimpanDokter_Click()
    If strKdDokter = "" Then
        MsgBox "Pilih dulu dokternya", vbCritical
        txtDokter.SetFocus
        Exit Sub
    End If
    Call cmdBatalDokter_Click
    Call sp_UbahDokter(dbcmd)
    Call cmdcari_Click
End Sub

Private Sub cmdTP_Click()
On Error GoTo hell
    If dgDaftarPasienRI.Columns(0).Value = "" Then
       Exit Sub
    End If
    frmDaftarPasienRI.Enabled = False
    Call subLoadFormTP
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        txtDokter.Text = dgDokter.Columns(1).Value
        strKdDokter = dgDokter.Columns(0).Value
        If strKdDokter = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        cmdSimpanDokter.SetFocus
'        fraDokter.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    optPasAktif.Caption = "Daftar Pasien Aktif " & strNNamaRuangan
    optPasAktif.Value = True
    Set rs = Nothing
    Call cmdcari_Click
    mblnForm = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnForm = False
End Sub

Private Sub mnurbs_Click()
    If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
    frm_cetak_RincianBiaya.Show
End Sub

Private Sub mnuSDokter_Click()
    If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
    fraDokterP.Visible = True
    fraDokterP.Enabled = True
    fraPilih.Enabled = False
    fraDaftar.Enabled = False
    fraCari.Enabled = False
    txtNoPendaftaran.Text = dgDaftarPasienRI.Columns("No. Register").Value
    txtNoCM.Text = dgDaftarPasienRI.Columns("No. CM").Value
    txtNamaPasien.Text = dgDaftarPasienRI.Columns("Nama Pasien").Value
    If dgDaftarPasienRI.Columns("JK").Value = "P" Then
        txtJK.Text = "Perempuan"
    Else
        txtJK.Text = "Laki-Laki"
    End If
    txtPoli.Text = strNNamaRuangan
    txtThn.Text = dgDaftarPasienRI.Columns("UmurTahun").Value
    txtBln.Text = dgDaftarPasienRI.Columns("UmurBulan").Value
    txtHr.Text = dgDaftarPasienRI.Columns("UmurHari").Value
    dTglMasuk = dgDaftarPasienRI.Columns("TglMasuk").Value
    txtTglPeriksa.Text = Format(dTglMasuk, "dd MMMM yyyy hh:mm:ss")
    txtDokter.Text = ""
    txtDokter.SetFocus
    strSQL = "select DokterPenanggungJawab from V_DaftarPasienRIAktif where Ruangan='" & strNNamaRuangan & "' and NoPendaftaran='" & txtNoPendaftaran.Text & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    txtPrevDokter.Text = rs("DokterPenanggungJawab").Value
End Sub

Private Sub mnuSEditDPasien_Click()
On Error GoTo hell
    strPasien = "Lama"
    strNoCM = dgDaftarPasienRI.Columns(1).Value
    frmPasienBaru.Show
hell:
End Sub

Public Sub optPasAktif_GotFocus()
On Error GoTo hell
    Set rs = Nothing
    strQuery = "select NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,TglMasuk,NoKamar,NoBed,NoPakai,UmurTahun,UmurBulan,UmurHari,KdSubInstalasi,KdKelas from V_DaftarPasienRIAktif where Ruangan='" & strNNamaRuangan & "' and ([Nama Pasien] like '%" & txtParameter.Text & "%' or NoCM like '%" & txtParameter.Text & "%')"
    rs.Open strQuery, dbConn, adOpenStatic, adLockOptimistic
    Set dgDaftarPasienRI.DataSource = rs
    Call SetGridPasienRIAktif
    cmdKeluarKamar.Enabled = True
    cmdMasukKamar.Enabled = False
    cmdPasienPulang.Enabled = False
    cmdPasienDirujuk.Enabled = True
    mnuSDokter.Enabled = True
    mnurbs.Enabled = True
hell:
End Sub

Public Sub optPasNonAktif_GotFocus()
On Error GoTo hell
    Set rs = Nothing
    strQuery = "select NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,TglKeluar,Ruangan,UmurTahun,UmurBulan,UmurHari,TglMasuk,KdSubInstalasi,KdKelas from V_DaftarPasienRINonAktif where ([Nama Pasien] like '%" & txtParameter.Text & "%' or NoCM like '%" & txtParameter.Text & "%')"
    rs.Open strQuery, dbConn, adOpenStatic, adLockOptimistic
    Set dgDaftarPasienRI.DataSource = rs
    Call SetGridPasienRINonAktif
    cmdKeluarKamar.Enabled = False
    cmdMasukKamar.Enabled = True
    cmdPasienPulang.Enabled = True
    cmdPasienDirujuk.Enabled = False
    mnuSDokter.Enabled = False
    mnurbs.Enabled = False
hell:
End Sub

Private Sub txtDokter_Change()
    strFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
    strKdDokter = ""
    Call subLoadDokter
End Sub

Private Sub txtDokter_GotFocus()
'    fraDokter.Visible = True
    If txtDokter.Text = "" Then strFilterDokter = ""
    Call subLoadDokter
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        dgDokter.SetFocus
    End If
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtParameter_Change()
    Call cmdcari_Click
End Sub

'untuk set grid pasien ri aktif
Sub SetGridPasienRIAktif()
    With dgDaftarPasienRI
        .Columns(0).Width = 1150
        .Columns(0).Caption = "No. Register"
        .Columns(1).Width = 750
        .Columns(1).Caption = "No. CM"
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 2100
        .Columns(3).Width = 300
        .Columns(4).Width = 1500
        .Columns(5).Width = 1800
        .Columns(6).Width = 1600
        .Columns(7).Width = 1900
        .Columns(8).Width = 810
        .Columns(8).Alignment = dbgCenter
        .Columns(9).Width = 580
        .Columns(9).Alignment = dbgCenter
        .Columns(10).Width = 0
        .Columns(11).Width = 0
        .Columns(12).Width = 0
        .Columns(13).Width = 0
        .Columns(14).Width = 0
        .Columns(15).Width = 0
    End With
End Sub

'untuk set grid pasien ri non aktif
Sub SetGridPasienRINonAktif()
    With dgDaftarPasienRI
        .Columns(0).Width = 1150
        .Columns(0).Caption = "No. Register"
        .Columns(1).Width = 750
        .Columns(1).Caption = "No. CM"
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 2000
        .Columns(3).Width = 300
        .Columns(4).Width = 1500
        .Columns(5).Width = 1500
        .Columns(6).Width = 1600
        .Columns(6).Caption = "Jenis Pasien"
        .Columns(7).Width = 1900
        .Columns(7).Caption = "Tgl. Keluar"
        .Columns(8).Width = 1800
        .Columns(8).Caption = "Ruangan Asal"
        .Columns(9).Width = 0
        .Columns(10).Width = 0
        .Columns(11).Width = 0
        .Columns(12).Width = 0
        .Columns(13).Width = 0
        .Columns(14).Width = 0
    End With
End Sub

'untuk load data pasien di form transaksi pasien
Private Sub subLoadFormTP()
    On Error GoTo hell
    strNoPen = dgDaftarPasienRI.Columns(0).Value
    strNoCM = dgDaftarPasienRI.Columns(1).Value
    If optPasAktif.Value = True Then
        With frmTransaksiPasien
            .Show
            .txtNoPendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
            .txtNoCM.Text = strNoCM
            .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
            If dgDaftarPasienRI.Columns(3).Value = "P" Then
                .txtSex.Text = "Perempuan"
            Else
                .txtSex.Text = "Laki-Laki"
            End If
            .txtKls.Text = dgDaftarPasienRI.Columns("Kelas").Value
            .txtThn.Text = dgDaftarPasienRI.Columns(11).Value
            .txtBln.Text = dgDaftarPasienRI.Columns(12).Value
            .txtHr.Text = dgDaftarPasienRI.Columns(13).Value
            .txtJenisPasien.Text = dgDaftarPasienRI.Columns(6).Value
            .txtTglDaftar.Text = dgDaftarPasienRI.Columns(7).Value
            mdTglMasuk = dgDaftarPasienRI.Columns(7).Value
            mstrKdKelas = dgDaftarPasienRI.Columns(15).Value
        End With
    ElseIf optPasNonAktif.Value = True Then
        If dgDaftarPasienRI.Columns(8).Value <> mstrNamaRuangan Then
            MsgBox "Anda tidak berhak mengakses pasien dari ruangan lain", vbCritical, "Validasi"
            Me.Enabled = True
            Exit Sub
        End If
        With frmTransaksiPasien
            .Show
            .txtNoPendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
            .txtNoCM.Text = strNoCM
            .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
            .txtSex.Text = dgDaftarPasienRI.Columns(3).Value
            .txtThn.Text = dgDaftarPasienRI.Columns(9).Value
            .txtBln.Text = dgDaftarPasienRI.Columns(10).Value
            .txtHr.Text = dgDaftarPasienRI.Columns(11).Value
            .txtJenisPasien.Text = dgDaftarPasienRI.Columns(6).Value
            .txtTglDaftar.Text = dgDaftarPasienRI.Columns(12).Value
            mdTglMasuk = dgDaftarPasienRI.Columns(12).Value
            mstrKdKelas = dgDaftarPasienRI.Columns(14).Value
        End With
    End If
hell:
End Sub

'untuk load data pasien di form keluar kamar
Private Sub subLoadFormKeluarKam()
    On Error GoTo hell
    strNoPen = dgDaftarPasienRI.Columns(0).Value
    strNoCM = dgDaftarPasienRI.Columns(1).Value
    With frmKeluarKamar
        .Show
        .txtNoPendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
        .txtNoCM.Text = strNoCM
        .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
        If dgDaftarPasienRI.Columns(3).Value = "P" Then
             .txtSex.Text = "Perempuan"
        Else
             .txtSex.Text = "Laki-Laki"
        End If
        .txtThn.Text = dgDaftarPasienRI.Columns(11).Value
        .txtBln.Text = dgDaftarPasienRI.Columns(12).Value
        .txtHari.Text = dgDaftarPasienRI.Columns(13).Value
        .txtNoPemakaian.Text = dgDaftarPasienRI.Columns(10).Value
        .txtTglMasuk.Text = dgDaftarPasienRI.Columns(7).Value
    End With
hell:
End Sub

'untuk load data pasien di form masuk kamar
Private Sub subLoadFormMasukKam()
'    On Error GoTo hell
    strNoPen = dgDaftarPasienRI.Columns(0).Value
    strNoCM = dgDaftarPasienRI.Columns(1).Value
    strSQL = "SELECT * FROM PemakaianKamar WHERE NoCM='" & strNoCM _
        & "' AND StatusKeluar='T'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount <> 0 Then
        MsgBox "Pasien belum keluar kamar", vbCritical, "Validasi"
        Me.Enabled = True
        Exit Sub
    End If
    With frmMasukKamar
        .Show
        .txtNoPendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
        .txtNoCM.Text = strNoCM
        .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
        If dgDaftarPasienRI.Columns(3).Value = "P" Then
             .txtSex.Text = "Perempuan"
        Else
             .txtSex.Text = "Laki-Laki"
        End If
        .txtThn.Text = dgDaftarPasienRI.Columns(9).Value
        .txtBln.Text = dgDaftarPasienRI.Columns(10).Value
        .txtHari.Text = dgDaftarPasienRI.Columns(11).Value
    End With
'hell:
End Sub

'untuk load data pasien di form masuk kamar
Private Sub subLoadFormPsnPulang()
'    On Error GoTo hell
    strNoPen = dgDaftarPasienRI.Columns(0).Value
    strNoCM = dgDaftarPasienRI.Columns(1).Value
    strSQL = "SELECT * FROM PemakaianKamar WHERE NoCM='" & strNoCM _
        & "' AND StatusKeluar='T'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount <> 0 Then
        MsgBox "Pasien belum keluar kamar", vbCritical, "Validasi"
        Me.Enabled = True
        Exit Sub
    End If
    If dgDaftarPasienRI.Columns(8).Value <> mstrNamaRuangan Then
        MsgBox "Anda tidak berhak mengakses pasien dari ruangan lain", vbCritical, "Validasi"
        Me.Enabled = True
        Exit Sub
    End If
    With frmPasienPulang
        .Show
        .txtNoPendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
        .txtNoCM.Text = strNoCM
        .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
        If dgDaftarPasienRI.Columns(3).Value = "P" Then
             .txtSex.Text = "Perempuan"
        Else
             .txtSex.Text = "Laki-Laki"
        End If
        .txtThn.Text = dgDaftarPasienRI.Columns(9).Value
        .txtBln.Text = dgDaftarPasienRI.Columns(10).Value
        .txtHari.Text = dgDaftarPasienRI.Columns(11).Value
        .txtTglMasuk.Text = dgDaftarPasienRI.Columns(7).Value
    End With
'hell:
End Sub

'untuk load data pasien di form transaksi pasien
Private Sub subLoadFormPeriksaDiagnosa()
'    On Error GoTo hell
    strNoPen = dgDaftarPasienRI.Columns(0).Value
    strNoCM = dgDaftarPasienRI.Columns(1).Value
    If optPasAktif.Value = True Then
        With frmPeriksaDiagnosa
            .Show
            .txtNoPendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
            .txtNoCM.Text = strNoCM
            .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
            If dgDaftarPasienRI.Columns(3).Value = "P" Then
             .txtSex.Text = "Perempuan"
            Else
             .txtSex.Text = "Laki-Laki"
            End If
            .txtThn.Text = dgDaftarPasienRI.Columns(11).Value
            .txtBln.Text = dgDaftarPasienRI.Columns(12).Value
            .txtHari.Text = dgDaftarPasienRI.Columns(13).Value
            strKdSubInstalasi = frmDaftarPasienRI.dgDaftarPasienRI.Columns(14)
        End With
    ElseIf optPasNonAktif.Value = True Then
        If dgDaftarPasienRI.Columns(8).Value <> mstrNamaRuangan Then
            MsgBox "Anda tidak berhak mengakses pasien dari ruangan lain", vbCritical, "Validasi"
            Me.Enabled = True
            Exit Sub
        End If
        With frmPeriksaDiagnosa
            .Show
            .txtNoPendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
            .txtNoCM.Text = strNoCM
            .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
            If dgDaftarPasienRI.Columns(3).Value = "P" Then
             .txtSex.Text = "Perempuan"
            Else
             .txtSex.Text = "Laki-Laki"
            End If
            .txtThn.Text = dgDaftarPasienRI.Columns(9).Value
            .txtBln.Text = dgDaftarPasienRI.Columns(10).Value
            .txtHari.Text = dgDaftarPasienRI.Columns(11).Value
            strKdSubInstalasi = frmDaftarPasienRI.dgDaftarPasienRI.Columns(13)
        End With
    End If
'hell:
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
        .Columns(1).Width = 4000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
End Sub

'Store procedure untuk mengisi registrasi pasien
Private Sub sp_UbahDokter(ByVal adocommand As ADODB.Command)
    With adocommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, strKdDokter)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dTglMasuk, "yyyy/MM/dd hh:mm:ss"))
        
        .ActiveConnection = dbConn
        .CommandText = "Update_DokterPemeriksaRI"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penyimpanan Dokter Pemeriksa pasien", vbCritical, "Validasi"
        Else
            MsgBox "Penyimpanan Dokter Pemeriksa pasien sukses", vbInformation, "Validasi"
        End If
        Call deleteADOCommandParameters(adocommand)
        Set adocommand = Nothing
    End With
    Exit Sub
End Sub
