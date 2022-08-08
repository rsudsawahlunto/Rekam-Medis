VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInfoPasienRI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Rawat Inap"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInfoPasienRI.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   14415
   Begin VB.Frame Frame2 
      Caption         =   "Daftar Pasien Rawat Inap"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   0
      TabIndex        =   12
      Top             =   840
      Width           =   14415
      Begin VB.Frame Frame3 
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
         Left            =   8520
         TabIndex        =   13
         Top             =   140
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
            Format          =   47448067
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
            Format          =   47448067
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
      Begin MSDataGridLib.DataGrid dgDaftarPasienRI 
         Height          =   5535
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   14175
         _ExtentX        =   25003
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
      Begin VB.Frame Frame4 
         Caption         =   "Status Pulang"
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
         Left            =   6360
         TabIndex        =   15
         Top             =   140
         Width           =   2055
         Begin VB.CheckBox chkSudah 
            Caption         =   "Sudah"
            Height          =   255
            Left            =   1095
            TabIndex        =   1
            Top             =   285
            Width           =   855
         End
         Begin VB.CheckBox chkBelum 
            Caption         =   "Belum"
            Height          =   255
            Left            =   135
            TabIndex        =   0
            Top             =   285
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cari Data Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   10
      Top             =   7440
      Width           =   14415
      Begin VB.CommandButton cmdTP 
         Caption         =   "Pemeriksaan Pasien"
         Height          =   495
         Left            =   10560
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
         Left            =   8640
         TabIndex        =   7
         ToolTipText     =   "Perbaiki data pasien"
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   12480
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   460
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Masukkan Nama Pasien / No. CM / Ruangan "
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   210
         Width           =   4455
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   4920
      Picture         =   "frmInfoPasienRI.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   0
      Picture         =   "frmInfoPasienRI.frx":431A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmInfoPasienRI"
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

Public Sub cmdcari_Click()
 On Error GoTo hell
     Set rs = Nothing
    If (chkBelum.Value = 1) And (chkSudah.Value = 0) Then
       rs.Open "select * from V_DaftarPasienRIAll where([Nama Pasien] like '%" & txtParameter.Text & "%' OR [No. CM] like '%" & txtParameter.Text & "%' OR Ruangan like '%" & txtParameter.Text & "%') AND TglMasuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND [Status Pulang]='Belum'", dbConn, adOpenStatic, adLockOptimistic
    ElseIf (chkSudah.Value = 1) And (chkBelum.Value = 0) Then
       rs.Open "select * from V_DaftarPasienRIAll where([Nama Pasien] like '%" & txtParameter.Text & "%' OR [No. CM] like '%" & txtParameter.Text & "%' OR Ruangan like '%" & txtParameter.Text & "%') AND TglMasuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND [Status Pulang]='Sudah'", dbConn, adOpenStatic, adLockOptimistic
    Else
       rs.Open "select * from V_DaftarPasienRIAll where([Nama Pasien] like '%" & txtParameter.Text & "%' OR [No. CM] like '%" & txtParameter.Text & "%' OR Ruangan like '%" & txtParameter.Text & "%') AND TglMasuk between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'", dbConn, adOpenStatic, adLockOptimistic
    End If
    Set dgDaftarPasienRI.DataSource = rs
    Call SetdgDaftarPasienRI
    If dgDaftarPasienRI.ApproxCount = 0 Then
        dtpAwal.SetFocus
    Else
        dgDaftarPasienRI.SetFocus
    End If
hell:
End Sub

Private Sub cmdDataPasien_Click()
On Error GoTo hell
    strPasien = "LamaReg"
    mstrNoCM = dgDaftarPasienRI.Columns("No. CM").Value
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

Private Sub dgDaftarPasienRI_KeyPress(KeyAscii As Integer)
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
End Sub

Sub SetdgDaftarPasienRI()
    With dgDaftarPasienRI
        .Columns(0).Width = 1800
        .Columns(0).Caption = "Ruang Perawatan"
        .Columns(1).Width = 1200
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 800
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 2000
        .Columns(4).Width = 300
        .Columns(4).Alignment = dbgCenter
        .Columns(5).Width = 1500
        .Columns(6).Width = 1590
        .Columns(7).Width = 1000
        .Columns(7).Alignment = dbgCenter
        .Columns(8).Width = 600
        .Columns(8).Alignment = dbgCenter
        .Columns(9).Width = 1200
        .Columns(9).Alignment = dbgCenter
        .Columns(10).Width = 1590
        .Columns(11).Width = 1590
        .Columns(12).Width = 1500
        .Columns(13).Width = 2800
    End With
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtParameter_LostFocus()
    txtParameter.Text = StrConv(txtParameter.Text, vbProperCase)
End Sub

Private Sub subLoadFormTP()
On Error GoTo hell
    mstrNoPen = dgDaftarPasienRI.Columns(1).Value
    mstrNoCM = dgDaftarPasienRI.Columns(2).Value
    mstrNoPen = mstrNoPen
    mstrNoCM = mstrNoCM
    
    With frmTransaksiPasien
        .Show
        .txtNoPendaftaran.Text = mstrNoPen
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = dgDaftarPasienRI.Columns(3).Value
        If dgDaftarPasienRI.Columns(4).Value = "L" Then
            .txtSex.Text = "Laki-Laki"
        Else
            .txtSex.Text = "Perempuan"
        End If
        .txtThn.Text = dgDaftarPasienRI.Columns(14).Value
        .txtBln.Text = dgDaftarPasienRI.Columns(15).Value
        .txtHr.Text = dgDaftarPasienRI.Columns(16).Value
        .txtKls.Text = dgDaftarPasienRI.Columns("Kelas").Value
        .txtJenisPasien.Text = dgDaftarPasienRI.Columns("JenisPasien").Value
        .txtTglDaftar.Text = dgDaftarPasienRI.Columns(6).Value
         mdTglMasuk = dgDaftarPasienRI.Columns(6).Value
         mstrKelas = dgDaftarPasienRI.Columns("Kelas").Value
         mstrKdRuangan = dgDaftarPasienRI.Columns(19).Value
         mstrKdSubInstalasi = dgDaftarPasienRI.Columns(20).Value
   End With
hell:
End Sub

