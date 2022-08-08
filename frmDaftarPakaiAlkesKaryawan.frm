VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDaftarPakaiAlkesKaryawan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Informasi Pemakaian Barang"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPakaiAlkesKaryawan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   14205
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   7440
      Width           =   14175
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         Height          =   465
         Left            =   12480
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtTotalBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya Pemakaian Bahan dan Obat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3585
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Daftar Pemakaian Bahan dan Obat Alkes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   14175
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
         Left            =   8280
         TabIndex        =   7
         Top             =   150
         Width           =   5775
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   615
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
            CustomFormat    =   "dd  MMMM, yyyy"
            Format          =   127664131
            UpDown          =   -1  'True
            CurrentDate     =   37967
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
            CustomFormat    =   "dd  MMMM, yyyy"
            Format          =   127664131
            UpDown          =   -1  'True
            CurrentDate     =   37967
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   8
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgDaftarPakaiAlkes 
         Height          =   5295
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   9340
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   10
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
      Left            =   12360
      Picture         =   "frmDaftarPakaiAlkesKaryawan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPakaiAlkesKaryawan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPakaiAlkesKaryawan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmDaftarPakaiAlkesKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCari_Click()
    On Error GoTo errLoad

    Dim rsa As New ADODB.recordset
    Dim strsqlku As String

    strsqlku = "select * from V_PemakaianBahanNObat where KdRuangan='" & mstrKdRuangan & "' and Tanggal BETWEEN '" & Format(dtpAwal.value, "yyyy-MM-dd 00:00:00") & "' AND '" & Format(dtpAkhir.value, "yyyy-MM-dd 23:59:59") & "'"
    rsa.Open strsqlku, dbConn, adOpenStatic, adLockOptimistic
    Set dgDaftarPakaiAlkes.DataSource = rsa
    Call SetGridDaftarAlkes

    strsqlku = "select sum(Total) as Total from V_PemakaianBahanNObat where KdRuangan='" & mstrKdRuangan & "' and Tanggal BETWEEN '" & Format(dtpAwal.value, "yyyy-MM-dd 00:00:00") & "' AND '" & Format(dtpAkhir.value, "yyyy-MM-dd 23:59:59") & "'"
    Set rsa = Nothing
    rsa.Open strsqlku, dbConn, adOpenStatic, adLockOptimistic
    txtTotalBiaya.Text = FormatCurrency(rsa!total, 0, vbUseDefault)
    Set rsa = Nothing

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgDaftarPakaiAlkes_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDaftarPakaiAlkes
    WheelHook.WheelHook dgDaftarPakaiAlkes
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

Private Sub Form_Activate()
    cmdCari_Click
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpAwal.value = Now
    dtpAkhir.value = Now
    Call cmdCari_Click

    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub SetGridDaftarAlkes()
    With dgDaftarPakaiAlkes
        .Columns("Ruangan").Width = 0
        .Columns("KdRuangan").Width = 0
        .Columns("Tanggal").Width = 1700
        .Columns("NamaBarang").Width = 2500
        .Columns("AsalBarang").Width = 1600
        .Columns("SatuanJml").Width = 900
        .Columns("JmlBarang").Width = 1000
        .Columns("HargaSatuan").Width = 1500
        .Columns("Total").Width = 1500
        .Columns("Keperluan").Width = 2000

        .Columns("JmlBarang").Alignment = dbgCenter
        .Columns("HargaSatuan").Alignment = dbgRight
        .Columns("Total").Alignment = dbgRight
    End With
End Sub
