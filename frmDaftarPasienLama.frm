VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDaftarPasienLama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Lama"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14700
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienLama.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   14700
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
      TabIndex        =   19
      Top             =   7440
      Width           =   14655
      Begin MSDataListLib.DataCombo dcStatusKeluar 
         Height          =   330
         Left            =   3600
         TabIndex        =   22
         Top             =   435
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   450
         Left            =   12960
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         TabIndex        =   9
         Top             =   440
         Width           =   2655
      End
      Begin VB.CommandButton cmdTP 
         Caption         =   "&Transaksi Pelayanan"
         Height          =   450
         Left            =   10920
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdBatalKeluar 
         Caption         =   "&Batal Keluar Kamar"
         Height          =   450
         Left            =   8880
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cara Keluar"
         Height          =   210
         Left            =   3600
         TabIndex        =   21
         Top             =   195
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Nama Pasien / No. CM"
         Height          =   210
         Left            =   720
         TabIndex        =   20
         Top             =   195
         Width           =   2640
      End
   End
   Begin VB.OptionButton OptSemua 
      Caption         =   "&Semua"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3960
      TabIndex        =   2
      Top             =   1560
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.Frame frameJudul 
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
      Left            =   -720
      TabIndex        =   13
      Top             =   1200
      Width           =   15375
      Begin VB.OptionButton optPulang 
         Caption         =   "Pulan&g"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3120
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optPindahan 
         Caption         =   "Pinda&han"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkRuangan 
         Caption         =   "Ruangan Perawatan"
         Height          =   255
         Left            =   6360
         TabIndex        =   3
         Top             =   120
         Value           =   1  'Checked
         Width           =   1935
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
         Left            =   9360
         TabIndex        =   14
         Top             =   120
         Width           =   5775
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   345
            Left            =   840
            TabIndex        =   5
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   136118275
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   345
            Left            =   3480
            TabIndex        =   6
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   136118275
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   15
            Top             =   300
            Width           =   255
         End
      End
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   360
         Left            =   6360
         TabIndex        =   4
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LblJumData 
         AutoSize        =   -1  'True
         Caption         =   "Data 0 / 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   840
         TabIndex        =   17
         Top             =   720
         Width           =   840
      End
   End
   Begin MSDataGridLib.DataGrid dgPasienLama 
      Height          =   5175
      Left            =   0
      TabIndex        =   8
      Top             =   2280
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   9128
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   8295
      Width           =   14700
      _ExtentX        =   25929
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   12912
            Text            =   "Cetak (F1)"
            TextSave        =   "Cetak (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   12912
            Text            =   "Refresh Data (F5)"
            TextSave        =   "Refresh Data (F5)"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   8599
            Text            =   "Cetak Daftar Pasien (F9)"
            TextSave        =   "Cetak Daftar Pasien (F9)"
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   18
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
      Left            =   12840
      Picture         =   "frmDaftarPasienLama.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPasienLama.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPasienLama.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmDaftarPasienLama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQLX As String

Private Sub chkRuangan_Click()
    If chkRuangan.value = Checked Then
        dcRuangan.Enabled = True
        dcRuangan.Text = ""
    Else
        dcRuangan.Enabled = False
        dcRuangan.Text = ""
    End If
End Sub

Private Sub cmdCari_Click()
    On Error GoTo hell
    Set rs = Nothing

    strSQLX = ""
    FlagStatusPulang = ""
    
    If dcRuangan.BoundText = "001" Then
        If optPindahan.value = True Then
            FlagStatusPulang = "1"
            If dtpAwal.Day <> dtpAkhir.Day Or dtpAwal.Month <> dtpAkhir.Month Or dtpAwal.Year <> dtpAkhir.Year Then
                rs.Open "select * from V_DaftarPasienLamaIGD where tglPulang is null and RuanganTujuan is not null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'", dbConn, adOpenStatic, adLockOptimistic
                strSQLX = "select * from V_DaftarPasienLamaIGD where tglPulang is null and RuanganTujuan is not null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'"
            Else
                rs.Open "select * from V_DaftarPasienLamaIGD where tglPulang is null and RuanganTujuan is not null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "')and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "'%", dbConn, adOpenStatic, adLockOptimistic
                strSQLX = "select * from V_DaftarPasienLamaIGD where tglPulang is null and RuanganTujuan is not null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "')and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'"
            End If
        End If
    
        If optPulang.value = True Then
            FlagStatusPulang = "2"
            If dtpAwal.Day <> dtpAkhir.Day Or dtpAwal.Month <> dtpAkhir.Month Or dtpAwal.Year <> dtpAkhir.Year Then
                rs.Open "select * from V_DaftarPasienLamaIGD where tglPulang is not null and RuanganTujuan is null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'", dbConn, adOpenStatic, adLockOptimistic
                strSQLX = "select * from V_DaftarPasienLamaIGD where tglPulang is not null and RuanganTujuan is null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'"
            Else
                rs.Open "select * from V_DaftarPasienLamaIGD where tglPulang is not null and RuanganTujuan is null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'", dbConn, adOpenStatic, adLockOptimistic
                strSQLX = "select * from V_DaftarPasienLamaIGD where tglPulang is not null and RuanganTujuan is null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'"
            End If
        End If
    
        If OptSemua.value = True Then
            FlagStatusPulang = "3"
            If dtpAwal.Day <> dtpAkhir.Day Or dtpAwal.Month <> dtpAkhir.Month Or dtpAwal.Year <> dtpAkhir.Year Then
                rs.Open "select * from V_DaftarPasienLamaIGD where ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'", dbConn, adOpenStatic, adLockOptimistic
                strSQLX = "select * from V_DaftarPasienLamaIGD where ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'"
            Else
                rs.Open "select * from V_DaftarPasienLamaIGD where ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'", dbConn, adOpenStatic, adLockOptimistic
                strSQLX = "select * from V_DaftarPasienLamaIGD where ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'"
            End If
        End If
    Else
        If optPindahan.value = True Then
            FlagStatusPulang = "1"
            If dtpAwal.Day <> dtpAkhir.Day Or dtpAwal.Month <> dtpAkhir.Month Or dtpAwal.Year <> dtpAkhir.Year Then
                rs.Open "select * from V_DaftarPasienLamaRI where tglPulang is null and RuanganTujuan is not null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'", dbConn, adOpenStatic, adLockOptimistic
                strSQLX = "select * from V_DaftarPasienLamaRI where tglPulang is null and RuanganTujuan is not null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'"
            Else
                rs.Open "select * from V_DaftarPasienLamaRI where tglPulang is null and RuanganTujuan is not null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "')and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "'%", dbConn, adOpenStatic, adLockOptimistic
                strSQLX = "select * from V_DaftarPasienLamaRI where tglPulang is null and RuanganTujuan is not null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "')and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'"
            End If
        End If
    
        If optPulang.value = True Then
            FlagStatusPulang = "2"
            If dtpAwal.Day <> dtpAkhir.Day Or dtpAwal.Month <> dtpAkhir.Month Or dtpAwal.Year <> dtpAkhir.Year Then
                rs.Open "select * from V_DaftarPasienLamaRI where tglPulang is not null and RuanganTujuan is null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'", dbConn, adOpenStatic, adLockOptimistic
                strSQLX = "select * from V_DaftarPasienLamaRI where tglPulang is not null and RuanganTujuan is null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'"
            Else
                rs.Open "select * from V_DaftarPasienLamaRI where tglPulang is not null and RuanganTujuan is null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'", dbConn, adOpenStatic, adLockOptimistic
                strSQLX = "select * from V_DaftarPasienLamaRI where tglPulang is not null and RuanganTujuan is null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'"
            End If
        End If
    
        If OptSemua.value = True Then
            FlagStatusPulang = "3"
            If dtpAwal.Day <> dtpAkhir.Day Or dtpAwal.Month <> dtpAkhir.Month Or dtpAwal.Year <> dtpAkhir.Year Then
                rs.Open "select * from V_DaftarPasienLamaRI where ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'", dbConn, adOpenStatic, adLockOptimistic
                strSQLX = "select * from V_DaftarPasienLamaRI where ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'"
            Else
                rs.Open "select * from V_DaftarPasienLamaRI where ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'", dbConn, adOpenStatic, adLockOptimistic
                strSQLX = "select * from V_DaftarPasienLamaRI where ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & dcRuangan.Text & "%' and [Cara Keluar] like '%" & dcStatusKeluar.Text & "%'"
            End If
        End If
    End If
    
    Set dgPasienLama.DataSource = rs
    Call SetGridPasienLamaRI

    If rs.EOF = False Then lblJumData.Caption = "Data 0/" & rs.RecordCount

    ' lock on
    If boolStafSIMRS = True Or strIDPegawaiAktif = "8888888888" Then
        If optPulang.value = True Then
            cmdTP.Enabled = True
            cmdBatalKeluar.Enabled = True
        Else
            cmdTP.Enabled = False
            cmdBatalKeluar.Enabled = False
        End If
    End If
    If dgPasienLama.ApproxCount = 0 Then dtpAwal.SetFocus Else dgPasienLama.SetFocus

    Exit Sub
hell:
End Sub

Private Sub cmdTP_Click()
    On Error GoTo errLoad

    If dgPasienLama.ApproxCount = 0 Then Exit Sub
    Call subLoadFormTP

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcRuangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcRuangan.MatchedWithList = True Then dtpAwal.SetFocus
        strSQL = "select kdruangan, namaruangan from V_LoginAplikasiRawatInap where StatusEnabled='1' and (namaruangan LIKE '%" & dcRuangan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcRuangan.Text = ""
            Exit Sub
        End If
        dcRuangan.BoundText = rs(0).value
        dcRuangan.Text = rs(1).value
    End If
End Sub

Private Sub dgPasienLama_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPasienLama
    WheelHook.WheelHook dgPasienLama
End Sub

Private Sub dgPasienLama_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    lblJumData.Caption = "Data " & dgPasienLama.Bookmark & " / " & dgPasienLama.ApproxCount
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    Select Case KeyCode
        Case vbKeyF5
            If dgPasienLama.ApproxCount = 0 Then Exit Sub
            Call cmdCari_Click
        Case vbKeyF9
            If dgPasienLama.ApproxCount = 0 Then Exit Sub
            frmCtkDaftarPasienLama.Show
        Case vbKeyF1
            If dgPasienLama.ApproxCount = 0 Then Exit Sub
            mdTglAwal = dtpAwal.value: mdTglAkhir = dtpAkhir.value
            frmCetakDaftarPasienLama.Show
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Set rs = Nothing
'    rs.Open "select * from V_LoginAplikasiRawatInap where StatusEnabled='1'", dbConn, adOpenStatic, adLockReadOnly
    rs.Open "select KdRuangan,NamaRuangan from Ruangan where KdInstalasi in ('01','03') and StatusEnabled='1'", dbConn, adOpenStatic, adLockReadOnly
    Set dcRuangan.RowSource = rs
    dcRuangan.ListField = rs(1).Name
    dcRuangan.BoundColumn = rs(0).Name
    Set rs = Nothing
    dtpAwal.value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.value = Now
    frameJudul.Caption = "Daftar Pasien Lama "
    optPindahan.value = False
    optPulang.value = False
    OptSemua.value = True
    Call cmdCari_Click
    Set rs = Nothing
    Call msubDcSource(dcStatusKeluar, rs, "select KdStatusKeluar,StatusKeluar from StatusKeluarKamar where StatusEnabled='1' order by StatusKeluar")
    
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub SetGridPasienLamaRI()
    With dgPasienLama

        .Columns(0).Width = 1200
        .Columns(0).Alignment = dbgCenter
        .Columns(0).Caption = "No. Pendaftaran"
        .Columns(1).Width = 1500
        .Columns(1).Alignment = dbgCenter
        .Columns(1).Caption = "No. Registrasi"
        .Columns(2).Width = 1800
        .Columns(3).Width = 300
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Width = 2200
        .Columns(5).Width = 1590
        .Columns(6).Width = 1590
        .Columns(7).Width = 1500
        .Columns(8).Width = 1800
        .Columns(9).Width = 1590
        .Columns(10).Width = 1500
        .Columns(11).Width = 1500
        .Columns(12).Width = 1200
        .Columns(13).Width = 1300
    End With
End Sub

Private Sub optPindahan_Click()
    optPulang.value = False
    OptSemua.value = False
    Call cmdCari_Click
End Sub

Private Sub optPindahan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkRuangan.SetFocus
End Sub

Private Sub optPulang_Click()
    optPindahan.value = False
    OptSemua.value = False
    Call cmdCari_Click
End Sub

Private Sub optPulang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkRuangan.SetFocus
End Sub

Private Sub OptSemua_Click()
    optPulang.value = False
    optPindahan.value = False
    Call cmdCari_Click
End Sub

Private Sub OptSemua_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkRuangan.SetFocus
End Sub

'untuk load data pasien di form transaksi pelayanan
Private Sub subLoadFormTP()
    On Error GoTo hell

    mstrNoPen = dgPasienLama.Columns("No. Pendaftaran").value
    mstrNoCM = dgPasienLama.Columns("No. Registrasi").value
    mstrKdRuangan = dgPasienLama.Columns("KdRuangan").value

    With frmTransaksiPasien
        .Show

        .txtNoPendaftaran.Text = dgPasienLama.Columns("No. Pendaftaran").value
        .txtNoCM.Text = dgPasienLama.Columns("No. Registrasi").value

        .txtNamaPasien.Text = dgPasienLama.Columns("Nama Pasien").value
        If dgPasienLama.Columns(3).value = "P" Then
            .txtSex.Text = "Perempuan"
        Else
            .txtSex.Text = "Laki-Laki"
        End If
        .txtKls.Text = dgPasienLama.Columns("Kelas").value
        .txtThn.Text = dgPasienLama.Columns("Thn").value
        .txtBln.Text = dgPasienLama.Columns("Bln").value
        .txtHr.Text = dgPasienLama.Columns("Hr").value

        .txtJenisPasien.Text = dgPasienLama.Columns("JenisPasien").value
        .txtTglDaftar.Text = dgPasienLama.Columns("TglPendaftaran").value
        mdTglMasuk = dgPasienLama.Columns("TglMasuk").value
        mstrKdKelas = dgPasienLama.Columns("KdKelas").value
        mstrKdSubInstalasi = dgPasienLama.Columns("KdSubInstalasi").value
    End With

    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call cmdCari_Click
        txtParameter.SetFocus
    End If
End Sub

