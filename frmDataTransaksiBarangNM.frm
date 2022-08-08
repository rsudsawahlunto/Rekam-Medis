VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDataTransaksiBarangNM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Rekapitulasi Transaksi Barang Non Medis"
   ClientHeight    =   8445
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
   Icon            =   "frmDataTransaksiBarangNM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   14790
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   7440
      Width           =   14775
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   13080
         TabIndex        =   4
         Top             =   300
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
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
      Width           =   14775
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
         Left            =   8880
         TabIndex        =   7
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
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   60227587
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
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   60227587
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
      Begin MSDataGridLib.DataGrid dgDataTransaksiBarang 
         Height          =   5295
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   9340
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   3
         RowHeight       =   19
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
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
      Left            =   12960
      Picture         =   "frmDataTransaksiBarangNM.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDataTransaksiBarangNM.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataTransaksiBarangNM.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmDataTransaksiBarangNM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCari_Click()
    Set rs = Nothing
    If mstrKdKelompokBarang = "02" Then     'medis
        strSQL = "select * from V_DataTransaksiBarangM where KdRuangan='" & mstrKdRuangan & "' " _
              & " AND TglTransaksi BETWEEN '" & Format(dtpAwal.Value, "yyyy-MM-dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy-MM-dd 23:59:59") & "'"
        rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
        Set dgDataTransaksiBarang.DataSource = rs
        Set rs = Nothing
        Call subSetGrid
    ElseIf mstrKdKelompokBarang = "01" Then     'non medis
        strSQL = "select * from V_DataTransaksiBarangNM where KdRuangan='" & mstrKdRuangan & "' " _
              & " AND TglTransaksi BETWEEN '" & Format(dtpAwal.Value, "yyyy-MM-dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy-MM-dd 23:59:59") & "'"
        rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
        Set dgDataTransaksiBarang.DataSource = rs
        Set rs = Nothing
        Call subSetGridNM
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
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

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call openConnection
    dtpAwal.Value = Format(Now, "dd MMM yyyy 00:00")
    dtpAkhir.Value = Now
    
    Call cmdCari_Click
End Sub

Private Sub subSetGrid()
    With dgDataTransaksiBarang
        .Columns(0).Width = 1590        'TglTransaksi
        .Columns(1).Width = 0           'Ruangan
        .Columns(2).Width = 1500        'JenisBarang
        .Columns(3).Width = 2000        'NamaBarang
        .Columns(4).Width = 1200        'AsalBarang
        .Columns(5).Width = 1000        'JmlStokAwal
        .Columns(5).Alignment = dbgRight
        .Columns(5).WrapText = True
        .Columns(6).Width = 1000        'JmlTerima
        .Columns(6).Alignment = dbgRight
        .Columns(7).Width = 1000        'JmlKeluar
        .Columns(7).Alignment = dbgRight
        .Columns(8).Alignment = dbgRight
        .Columns(8).Width = 1000
        .Columns(8).WrapText = True
        .Columns(9).Alignment = dbgRight
        .Columns(9).Width = 1000
        .Columns(9).WrapText = True
        .Columns(10).Alignment = dbgRight
        .Columns(10).Width = 1000
        .Columns(10).WrapText = True
        .Columns(11).Alignment = dbgRight
        .Columns(11).Width = 1000
        .Columns(11).WrapText = True
        .Columns(12).Alignment = dbgRight
        .Columns(12).Width = 1000
        .Columns(12).WrapText = True
        .Columns(13).Alignment = dbgRight
        .Columns(13).Width = 1000
        .Columns(13).WrapText = True
        .Columns(14).Width = 1100
        .Columns(14).Alignment = dbgCenter
        .Columns(15).Width = 0
        .Columns(16).Width = 0
        .Columns(17).Width = 0
        .Columns(18).Width = 0
    End With
End Sub

Private Sub subSetGridNM()
With dgDataTransaksiBarang
        .Columns(0).Width = 1590        'TglTransaksi
        .Columns(1).Width = 0           'Ruangan
        .Columns(2).Width = 1500        'JenisBarang
        .Columns(3).Width = 2000        'NamaBarang
        .Columns(4).Width = 1200        'AsalBarang
        .Columns(5).Width = 1200        'Merk
        .Columns(6).Width = 1200        'Type
        .Columns(7).Width = 1200        'Bahan
        .Columns(8).Width = 1000        'JmlStokAwal
        .Columns(8).Alignment = dbgRight
        .Columns(8).WrapText = True
        .Columns(9).Width = 1000        'JmlTerima
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 1000       'JmlKeluar
        .Columns(10).Alignment = dbgRight
        .Columns(11).Alignment = dbgRight
        .Columns(11).Width = 1000
        .Columns(11).WrapText = True
        .Columns(12).Alignment = dbgRight
        .Columns(12).Width = 1000
        .Columns(12).WrapText = True
        .Columns(13).Alignment = dbgRight
        .Columns(13).Width = 1000
        .Columns(13).WrapText = True
        .Columns(14).Alignment = dbgRight
        .Columns(14).Width = 1000
        .Columns(14).WrapText = True
        .Columns(15).Alignment = dbgRight
        .Columns(15).Width = 1000
        .Columns(15).WrapText = True
        .Columns(16).Alignment = dbgRight
        .Columns(16).Width = 1000
        .Columns(16).WrapText = True
        .Columns(17).Width = 1100
        .Columns(17).Alignment = dbgCenter
        .Columns(18).Width = 0
        .Columns(19).Width = 0
        .Columns(20).Width = 0
        .Columns(21).Width = 0
        .Columns(22).Width = 0
'        .Columns(23).Width = 0
'        .Columns(24).Width = 0
    End With
End Sub
