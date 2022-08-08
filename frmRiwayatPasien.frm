VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRiwayatPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Pemeriksaan  Pasien"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRiwayatPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   14670
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   48
      Top             =   8550
      Width           =   14670
      _ExtentX        =   25876
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   25823
            Text            =   "Refresh Data (F5)"
            TextSave        =   "Refresh Data (F5)"
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
   Begin VB.Frame Frame1 
      Caption         =   "Pelayanan Pasien"
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
      TabIndex        =   46
      Top             =   2040
      Width           =   14655
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   12480
         TabIndex        =   36
         Top             =   6000
         Width           =   2055
      End
      Begin TabDlg.SSTab sstTP 
         Height          =   5535
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   9763
         _Version        =   393216
         Tabs            =   11
         Tab             =   5
         TabsPerRow      =   11
         TabHeight       =   970
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Pe&layanan Tindakan"
         TabPicture(0)   =   "frmRiwayatPasien.frx":0CCA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtTindakanTotal"
         Tab(0).Control(1)=   "cmdUbahPT"
         Tab(0).Control(2)=   "cmdHapusDataPT"
         Tab(0).Control(3)=   "cmdTambahPT"
         Tab(0).Control(4)=   "dgTindakan"
         Tab(0).Control(5)=   "Label1"
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Pemakaian &Obat && Alkes"
         TabPicture(1)   =   "frmRiwayatPasien.frx":0CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtAlkesTotal"
         Tab(1).Control(1)=   "cmdHapusDataPOA"
         Tab(1).Control(2)=   "cmdTambahPOA"
         Tab(1).Control(3)=   "dgObatAlkes"
         Tab(1).Control(4)=   "Label2"
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "Riwayat Catatan Klinis"
         TabPicture(2)   =   "frmRiwayatPasien.frx":0D02
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "dgCatatanKlinis"
         Tab(2).Control(1)=   "cmdCetakCatatanKlinis"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Riwayat Catatan Medis"
         TabPicture(3)   =   "frmRiwayatPasien.frx":0D1E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "dgCatatanMedis"
         Tab(3).Control(1)=   "cmdCetakCatatanMedis"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Riwayat &Diagnosa"
         TabPicture(4)   =   "frmRiwayatPasien.frx":0D3A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "cmdCetakDiagnosa"
         Tab(4).Control(1)=   "dgRiwayatDiagnosa"
         Tab(4).ControlCount=   2
         TabCaption(5)   =   "Riwayat Operasi"
         TabPicture(5)   =   "frmRiwayatPasien.frx":0D56
         Tab(5).ControlEnabled=   -1  'True
         Tab(5).Control(0)=   "dgRiwayatOperasi"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).Control(1)=   "cmdCetakOperasi"
         Tab(5).Control(1).Enabled=   0   'False
         Tab(5).ControlCount=   2
         TabCaption(6)   =   "Riwayat Konsul"
         TabPicture(6)   =   "frmRiwayatPasien.frx":0D72
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "cmdRiwayatKonsul"
         Tab(6).Control(1)=   "dgKonsul"
         Tab(6).ControlCount=   2
         TabCaption(7)   =   "Riwayat Kecelakaan"
         TabPicture(7)   =   "frmRiwayatPasien.frx":0D8E
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "cmdCetakKecelakaan"
         Tab(7).Control(1)=   "dgKecelakaan"
         Tab(7).ControlCount=   2
         TabCaption(8)   =   "Riwayat Peme&riksaan"
         TabPicture(8)   =   "frmRiwayatPasien.frx":0DAA
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "cmdCetakRiwayatPemeriksaan"
         Tab(8).Control(1)=   "dgRiwayatPemeriksaan"
         Tab(8).ControlCount=   2
         TabCaption(9)   =   "Riwayat Hasil Pemeriksaan"
         TabPicture(9)   =   "frmRiwayatPasien.frx":0DC6
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "cmdCetakHasilPemeriksaan"
         Tab(9).Control(1)=   "dgHasilPemeriksaan"
         Tab(9).ControlCount=   2
         TabCaption(10)  =   "Riwayat Kunjungan"
         TabPicture(10)  =   "frmRiwayatPasien.frx":0DE2
         Tab(10).ControlEnabled=   0   'False
         Tab(10).Control(0)=   "cmdRiwayatKunjungan"
         Tab(10).Control(1)=   "dgRiwayatKunjungan"
         Tab(10).ControlCount=   2
         Begin VB.CommandButton cmdRiwayatKunjungan 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -62400
            TabIndex        =   35
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetakCatatanMedis 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -62400
            TabIndex        =   21
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetakCatatanKlinis 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -62400
            TabIndex        =   19
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetakRiwayatPemeriksaan 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -62400
            TabIndex        =   31
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetakDiagnosa 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -62400
            TabIndex        =   23
            Top             =   5040
            Width           =   1575
         End
         Begin VB.TextBox txtAlkesTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   -71640
            TabIndex        =   15
            Top             =   4965
            Width           =   2415
         End
         Begin VB.TextBox txtTindakanTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   -72000
            TabIndex        =   10
            Top             =   4965
            Width           =   2415
         End
         Begin VB.CommandButton cmdCetakOperasi 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   12600
            TabIndex        =   25
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetakKecelakaan 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -62400
            TabIndex        =   29
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetakHasilPemeriksaan 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -62400
            TabIndex        =   33
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdUbahPT 
            Caption         =   "&Ubah Data"
            Height          =   375
            Left            =   -65760
            TabIndex        =   11
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusDataPT 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   -64080
            TabIndex        =   12
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahPT 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   -62400
            TabIndex        =   13
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusDataPOA 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   -64080
            TabIndex        =   16
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahPOA 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   -62400
            TabIndex        =   17
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdRiwayatKonsul 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -62400
            TabIndex        =   27
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin MSDataGridLib.DataGrid dgTindakan 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   9
            Top             =   765
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgObatAlkes 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   14
            Top             =   765
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgRiwayatDiagnosa 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   22
            Top             =   840
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgRiwayatPemeriksaan 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   30
            Top             =   840
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgCatatanKlinis 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   18
            Top             =   840
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgCatatanMedis 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   20
            Top             =   840
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgRiwayatOperasi 
            Height          =   4095
            Left            =   240
            TabIndex        =   24
            Top             =   840
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgKonsul 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   26
            Top             =   840
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgKecelakaan 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   28
            Top             =   840
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgHasilPemeriksaan 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   32
            Top             =   840
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   7223
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
         Begin MSDataGridLib.DataGrid dgRiwayatKunjungan 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   34
            Top             =   840
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   7223
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total Biaya Pemakaian Obat && Alkes"
            Height          =   210
            Left            =   -74760
            TabIndex        =   50
            Top             =   5025
            Width           =   2925
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Biaya Pelayanan Tindakan"
            Height          =   210
            Left            =   -74760
            TabIndex        =   49
            Top             =   5025
            Width           =   2550
         End
      End
   End
   Begin VB.Frame Frame3 
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
      TabIndex        =   37
      Top             =   960
      Width           =   14655
      Begin VB.TextBox txtAlamat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   8520
         TabIndex        =   6
         Top             =   600
         Width           =   4215
      End
      Begin VB.Frame Frame5 
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
         Height          =   580
         Left            =   6000
         TabIndex        =   38
         Top             =   360
         Width           =   2415
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   900
            MaxLength       =   6
            TabIndex        =   4
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   3
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Left            =   2130
            TabIndex        =   41
            Top             =   270
            Width           =   165
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Left            =   1350
            TabIndex        =   40
            Top             =   277
            Width           =   240
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Left            =   550
            TabIndex        =   39
            Top             =   277
            Width           =   285
         End
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4680
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtTglDaftar 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   12840
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Alamat"
         Height          =   210
         Left            =   8520
         TabIndex        =   47
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   1560
         TabIndex        =   44
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   4680
         TabIndex        =   43
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Membership"
         Height          =   210
         Left            =   12840
         TabIndex        =   42
         Top             =   360
         Width           =   1350
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   51
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
      Picture         =   "frmRiwayatPasien.frx":0DFE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatPasien.frx":1B86
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatPasien.frx":4547
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmRiwayatPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim subKdDokterTemp As String
Dim intJumlahPrint  As Integer

Private Sub cmdCetakCatatanMedis_Click()
    On Error GoTo hell
    If dgCatatanMedis.ApproxCount = 0 Then
        Exit Sub
    Else
        vLaporan = ""
        If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
        cmdCetakCatatanMedis.Enabled = False
        frmCetakCatatanMedis.Show
        cmdCetakCatatanMedis.Enabled = True
    End If
    Exit Sub
hell:
End Sub

Private Sub cmdCetakDiagnosa_Click()
    On Error GoTo jasmed
    mblnStatusCetakRD = True
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frm_cetak_info_diag_viewer.Show
jasmed:
End Sub

Private Sub cmdCetakHasilPemeriksaan_Click()
    On Error GoTo errLoad

    If dgHasilPemeriksaan.ApproxCount = 0 Then
        Exit Sub
    Else
        cmdCetakHasilPemeriksaan.Enabled = False
        strSQL = "SELECT * FROM V_RiwayatHasilPemeriksaan" & _
        " WHERE NoCM = '" & mstrNoCM & "'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount = 0 Then
            cmdCetakHasilPemeriksaan.Enabled = True
            Exit Sub
        End If

        mstrNoLabRad = dgHasilPemeriksaan.Columns("NoLab_Rad").value

        Select Case dgHasilPemeriksaan.Columns("KdInstalasi").value
            Case "09" 'lab pk
                vLaporan = ""
                If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
                Set frmcetakhasillab = Nothing
                frmcetakhasillab.Show

            Case "16" 'lab pa
                vLaporan = ""
                If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
                Set frmCetakHasilLabPA = Nothing
                frmCetakHasilLabPA.Show

            Case "10" 'radiologi
                vLaporan = ""
                If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
                Set frmCetakHasilRadiologi = Nothing
                frmCetakHasilRadiologi.Show
        End Select

        cmdCetakHasilPemeriksaan.Enabled = True

    End If

    Exit Sub
errLoad:
End Sub

Private Sub cmdRiwayatKonsul_Click()
    On Error GoTo errLoad

    If dgKonsul.ApproxCount = 0 Then
        Exit Sub
    Else
        If intJumlahPrint = 0 Then
            intJumlahPrint = 1
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmCetakStrukKonsul.Show
        Else
            intJumlahPrint = 0
        End If
    End If

    Exit Sub
errLoad:
End Sub

Private Sub cmdTambahPOA_Click()
    On Error GoTo errLoad

    strSQL = "SELECT dbo.RegistrasiRJ.IdDokter, dbo.DataPegawai.NamaLengkap " & _
    " FROM dbo.RegistrasiRJ INNER JOIN dbo.DataPegawai ON dbo.RegistrasiRJ.IdDokter = dbo.DataPegawai.IdPegawai " & _
    " WHERE (dbo.RegistrasiRJ.NoPendaftaran = '" & dgObatAlkes.Columns(1).value & "')"
    Call msubRecFO(rs, strSQL)

    If Not rs.EOF Then
        frmPemakaianObatAlkes.txtDokter.Text = rs(1).value
        mstrKdDokter = rs(0).value
        intJmlDokter = rs.RecordCount
        frmPemakaianObatAlkes.frameDokter.Visible = False
    End If

    frmPemakaianObatAlkes.Show

    Exit Sub
errLoad:
    Call msubPesanError
    frmPemakaianObatAlkes.Show
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgCatatanKlinis_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgCatatanKlinis
    WheelHook.WheelHook dgCatatanKlinis
End Sub

Private Sub dgCatatanKlinis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetakCatatanKlinis.SetFocus
End Sub

Private Sub dgCatatanMedis_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgCatatanMedis
    WheelHook.WheelHook dgCatatanMedis
End Sub

Private Sub dgCatatanMedis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetakCatatanMedis.SetFocus
End Sub

Private Sub dgHasilPemeriksaan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgHasilPemeriksaan
    WheelHook.WheelHook dgHasilPemeriksaan
End Sub

Private Sub dgHasilPemeriksaan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetakHasilPemeriksaan.SetFocus
End Sub

Private Sub dgKecelakaan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKecelakaan
    WheelHook.WheelHook dgKecelakaan
End Sub

Private Sub dgKecelakaan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetakKecelakaan.SetFocus
End Sub

Private Sub dgKonsul_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKonsul
    WheelHook.WheelHook dgKonsul
End Sub

Private Sub dgKonsul_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdRiwayatKonsul.SetFocus
End Sub

Private Sub dgObatAlkes_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgObatAlkes
    WheelHook.WheelHook dgObatAlkes
End Sub

Private Sub dgObatAlkes_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then cmdTambahPOA.SetFocus
End Sub

Private Sub dgRiwayatDiagnosa_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgRiwayatDiagnosa
    WheelHook.WheelHook dgRiwayatDiagnosa
End Sub

Private Sub dgRiwayatDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetakDiagnosa.SetFocus
End Sub

Private Sub dgRiwayatKunjungan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgRiwayatKunjungan
    WheelHook.WheelHook dgRiwayatKunjungan
End Sub

Private Sub dgRiwayatKunjungan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdRiwayatKunjungan.SetFocus
End Sub

Private Sub dgRiwayatOperasi_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgRiwayatOperasi
    WheelHook.WheelHook dgRiwayatOperasi
End Sub

Private Sub dgRiwayatOperasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetakOperasi.SetFocus
End Sub

Private Sub dgRiwayatPemeriksaan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgRiwayatPemeriksaan
    WheelHook.WheelHook dgRiwayatPemeriksaan
End Sub

Private Sub dgRiwayatPemeriksaan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCetakRiwayatPemeriksaan.SetFocus
End Sub

Private Sub dgTindakan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgTindakan
    WheelHook.WheelHook dgTindakan
End Sub

Private Sub dgTindakan_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then cmdTambahPT.SetFocus
End Sub

Private Sub Form_Activate()
    On Error GoTo errLoad
    Call subLoadRiwayatCatatanKlinis
    Call subLoadRiwayatCatatanMedis
    Call subLoadRiwayatDiagnosa
    Call subLoadRiwayatKecelakaan
    Call subLoadRiwayatOperasi
    Call subLoadRiwayatKonsul
    Call subLoadRiwayatPemeriksaan(False)
    Call subLoadRiwayatHasilPemeriksaan
    Call subLoadRiwayatKunjungan
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errLoad

    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    Select Case KeyCode
        Case vbKey1
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 0
        Case vbKey2
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 1
        Case vbKey3
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 2
        Case vbKey4
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 3
        Case vbKey5
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 4
        Case vbKey6
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 5
        Case vbKey7
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 6
        Case vbKey8
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 7
        Case vbKey9
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 8
        Case vbKey0
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 9
        Case vbKeyF5
            Call subLoadRiwayatCatatanKlinis
            Call subLoadRiwayatCatatanMedis
            Call subLoadRiwayatDiagnosa
            Call subLoadRiwayatKecelakaan
            Call subLoadRiwayatOperasi
            Call subLoadRiwayatKonsul
            Call subLoadRiwayatPemeriksaan(False)
            Call subLoadRiwayatHasilPemeriksaan
            Call subLoadRiwayatKunjungan
    End Select

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call subLoadRiwayatCatatanKlinis
    Call subLoadRiwayatCatatanMedis
    Call subLoadRiwayatDiagnosa
    Call subLoadRiwayatKecelakaan
    Call subLoadRiwayatOperasi
    Call subLoadRiwayatKonsul
    Call subLoadRiwayatPemeriksaan(False)
    Call subLoadRiwayatHasilPemeriksaan
    Call subLoadRiwayatKunjungan

    sstTP.TabVisible(0) = False
    sstTP.TabVisible(1) = False
    sstTP.TabsPerRow = 9
    sstTP.Tab = 2

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Untuk meload riwayat diagnosa yang sudah pernah didapat
Public Sub subLoadRiwayatDiagnosa()
    Set rs = Nothing
    strSQL = "Select * from V_DaftarDiagnosaPasien where nocm = '" & mstrNoCM & "'"
    rs.Open strSQL, dbConn, adOpenDynamic, adLockOptimistic
    Set dgRiwayatDiagnosa.DataSource = rs
    With dgRiwayatDiagnosa
        .Columns(0).Width = 0
        .Columns(1).Width = 0
        .Columns(2).Width = 1590
        .Columns(2).Caption = "Tgl. Periksa"
        .Columns(3).Width = 2000
        .Columns(3).Caption = "Jenis Diagnosa"
        .Columns(4).Width = 900
        .Columns(4).Caption = "Kode ICD"
        .Columns(5).Width = 4500
        .Columns(5).Caption = "Diagnosa ICD"
        .Columns(6).Width = 2200
        .Columns(7).Width = 2700
        .Columns(8).Width = 0
        .Columns(9).Width = 0
        .Columns(10).Width = 0
    End With

    Set rs = Nothing
End Sub

'Untuk meload riwayat pemeriksaan yang sudah pernah didapat
Public Sub subLoadRiwayatPemeriksaan(blnAll As Boolean)
    If blnAll = False Then
        strSQL = "Select * from V_RiwayatPemeriksaanPasien where nocm = '" & mstrNoCM & "' " 'AND KdRuangan='" & mstrKdRuangan & "'"
    Else
        strSQL = "Select * from V_RiwayatPemeriksaanPasien where nocm = '" & mstrNoCM & "'"
    End If
    msubRecFO rs, strSQL
    Set dgRiwayatPemeriksaan.DataSource = rs
    With dgRiwayatPemeriksaan
        .Columns(0).Width = 0 'nocm
        .Columns(1).Width = 0 ' nopendaftaran
        .Columns(2).Width = 2400
        .Columns(3).Width = 1590
        .Columns(4).Width = 4000
        .Columns(5).Width = 3000
        .Columns(6).Width = 2500
        .Columns("KdRuangan").Width = 0
        .Columns("NoLab_Rad").Width = 0
        .Columns("KdJnsPelayanan").Width = 0
        .Columns("KdPelayananRS").Width = 0
    End With
    Set rs = Nothing
End Sub

'Untuk meload riwayat kunjungan yang sudah pernah didapat
Public Sub subLoadRiwayatKunjungan()
    strSQL = "Select * from V_RiwayatKunjunganPasienAll where nocm = '" & mstrNoCM & "' ORDER BY TglMasuk DESC"
    msubRecFO rs, strSQL
    Set dgRiwayatKunjungan.DataSource = rs
    With dgRiwayatKunjungan
        .Columns(0).Width = 0 'nocm
        .Columns("RuanganPerawatan").Width = 3500
        .Columns("KasusPenyakit").Width = 2900
    End With
    Set rs = Nothing
End Sub

'Untuk meload riwayat hasil pemeriksaan yang sudah pernah didapat
Public Sub subLoadRiwayatHasilPemeriksaan()
    strSQL = "Select NoLab_Rad, [Ruang Pemeriksa], [Dokter Pemeriksa], TglPendaftaran, TglHasil, [Asal Rujukan], [Ruangan Perujuk], [Dokter Perujuk], KdInstalasi from V_RiwayatHasilPemeriksaan where nocm = '" & mstrNoCM & "'"
    msubRecFO rs, strSQL
    Set dgHasilPemeriksaan.DataSource = rs
    dgHasilPemeriksaan.Columns("KdInstalasi").Width = 0
End Sub

Private Sub sstTP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case sstTP.Tab
            Case 0 'pelayanan tindakan
                dgTindakan.SetFocus
            Case 1 'pemakaian obat alkes
                dgObatAlkes.SetFocus
            Case 2 'riwayat catatan klinis
                dgCatatanKlinis.SetFocus
            Case 3 'riwayat catatan medis
                dgCatatanMedis.SetFocus
            Case 4 'riwayat diagnosa
                dgRiwayatDiagnosa.SetFocus
            Case 5 'riwayat catatan operasi
                dgRiwayatOperasi.SetFocus
            Case 6 'riwayat konsul
                dgKonsul.SetFocus
            Case 7 'riwayat kecelakaan
                dgKecelakaan.SetFocus
            Case 8 'riwayat pemeriksaan
                dgRiwayatPemeriksaan.SetFocus
            Case 9 ' riwayat hasil pemeriksaan
                dgHasilPemeriksaan.SetFocus
            Case 10 ' riwayat kunjungan
                dgRiwayatKunjungan.SetFocus
        End Select
    End If
End Sub

'Untuk meload riwayat catatan klinis yang sudah pernah didapat
Public Sub subLoadRiwayatCatatanKlinis()
    strSQL = "SELECT * " & _
    " FROM V_RiwayatCatatanKlinisPasien" & _
    " WHERE (nocm = '" & mstrNoCM & "')"
    Call msubRecFO(rs, strSQL)

    Set dgCatatanKlinis.DataSource = rs
    With dgCatatanKlinis
        .Columns(0).Width = 0 'NoCM
        .Columns(1).Width = 0 'NoPendaftaran
        .Columns(2).Width = 1590 'TglPeriksa
        .Columns(3).Width = 1500 '[Ruang Pemeriksaan]
        .Columns(4).Width = 1300 'TekananDarah
        .Columns(5).Width = 1000 'Nadi
        .Columns(6).Width = 1000 'Pernapasan
        .Columns(7).Width = 1000 'Suhu
        .Columns(8).Width = 1500 'BeratTinggiBadan
        .Columns(9).Width = 1500 'Kesadaran
        .Columns(10).Width = 1500 'Keterangan
        .Columns(11).Width = 1500 'Pemeriksa
        .Columns(12).Width = 0 'KdRuangan
    End With
End Sub

'Untuk meload riwayat catatan medis yang sudah pernah didapat
Public Sub subLoadRiwayatCatatanMedis()
    strSQL = "SELECT *" & _
    " FROM V_RiwayatCatatanMedisPasien " & _
    " WHERE (nocm = '" & mstrNoCM & "')"
    Call msubRecFO(rs, strSQL)

    Set dgCatatanMedis.DataSource = rs
    With dgCatatanMedis
        .Columns(0).Width = 0 'NoCM
        .Columns(1).Width = 0 'NoPendaftaran
        .Columns(2).Width = 1590 'TglPeriksa
        .Columns(3).Width = 1500 'RuangPemeriksaan
        .Columns(4).Width = 1500 'KeluhanUtama
        .Columns(5).Width = 2500 'Diagnosa
        .Columns(6).Width = 2500 'Pengobatan
        .Columns(7).Width = 1500 'Keterangan
        .Columns(8).Width = 2500 '[Dokter Pemeriksa]
        .Columns(9).Width = 0 'KdRuangan
    End With
End Sub

'Untuk meload riwayat Kecelakaan yang sudah pernah didapat
Public Sub subLoadRiwayatKecelakaan()
    strSQL = "SELECT *" & _
    " FROM V_RiwayatKecelakanPasien " & _
    " WHERE (nocm = '" & mstrNoCM & "')"
    Call msubRecFO(rs, strSQL)

    Set dgKecelakaan.DataSource = rs
    With dgKecelakaan
        .Columns(0).Width = 0 'NoCM
        .Columns(1).Width = 0 'NoPendaftaran
        .Columns(2).Width = 1590 'TglPeriksa
        .Columns(3).Width = 1500 '[Ruangan Pemeriksa]
        .Columns(4).Width = 2500 '[Kasus Kecelakaan]
        .Columns(5).Width = 1590  'TglKecelakaan
        .Columns(6).Width = 2500 'TempatKecelakaan
        .Columns(7).Width = 1500 'Pemeriksa
        .Columns(8).Width = 0 'KdRuangan
    End With
End Sub

'Untuk meload riwayat konsul pasien
Public Sub subLoadRiwayatKonsul()
    strSQL = "SELECT * " & _
    " FROM V_RiwayatRujukanPasien " & _
    " where (nocm = '" & mstrNoCM & "')"
    Call msubRecFO(rs, strSQL)

    Set dgKonsul.DataSource = rs
    With dgKonsul
        .Columns(0).Width = 0 'NoCM
        .Columns(1).Width = 0 'NoPendaftaran
        .Columns(2).Width = 1590 'TglDirujuk
        .Columns(3).Width = 2500 '[Ruangan Perujuk]
        .Columns(4).Width = 2500 '[Ruangan Tujuan]
        .Columns(5).Width = 2500 '[Dokter Perujuk]
        .Columns(6).Width = 1700 'StatusPeriksa
        .Columns(7).Width = 0 'KdRuanganAsal
    End With
End Sub

'Untuk meload riwayat operasi yang sudah pernah didapat
Public Sub subLoadRiwayatOperasi()
'    strSQL = "SELECT * " & _
'    " FROM V_RiwayatOperasiPasien " & _
'    " WHERE (nocm = '" & mstrNoCM & "')"
'
'    Call msubRecFO(rs, strSQL)
'    Set dgRiwayatOperasi.DataSource = rs
'    With dgRiwayatOperasi
'        .Columns(0).Width = 0 'NoCM
'        .Columns(1).Width = 0 'NoPendaftaran
'        .Columns(2).Width = 1590 '[Tgl. Operasi]
'        .Columns(2).Caption = "Tgl. Operasi"
'        .Columns(3).Width = 1590 '[Tgl. Selesai]
'        .Columns(3).Caption = "Tgl. Selesai"
'        .Columns(4).Width = 2500 '[Jenis Operasi]
'        .Columns(5).Width = 2500 '[Tindakan Operasi]
'        .Columns(6).Width = 1900 '[Ruang Perujuk]
'        .Columns(7).Width = 2700 '[Dokter Penanggung Jawab]
'        .Columns(8).Width = 0 'KdRuangan
'    End With
    Dim i As Integer

    strSQL = "Select * From V_HasilTindakanMedisPasien Where (NoCm='" & mstrNoCM & "')"
    Call msubRecFO(rs, strSQL)
    Set dgRiwayatOperasi.DataSource = rs
    With dgRiwayatOperasi
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i
        .Columns("TglHasilPeriksa").Width = 1590
        .Columns("JenisTindakanMedis").Width = 1200
        .Columns("Kamar").Width = 0
        .Columns("DokterKepala").Width = 2700
        .Columns("JenisAnastesi").Width = 1200
        .Columns("KetAnastesi").Width = 1900
        .Columns("NamaPelayanan").Width = 1900
        .Columns("TindakanMedis").Width = 1200
        .Columns("KualitasHasil").Width = 1200
        .Columns("HasilPeriksa").Width = 1500
        .Columns("MemoHasilPeriksa").Width = 1500
    End With

End Sub

Public Sub txtNoCM_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    strSQL = "SELECT * " & _
    " FROM V_CariPasien WHERE [No. CM] = '" & txtNoCM.Text & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        mstrNoCM = ""
        Exit Sub
    End If

    mstrNoCM = txtNoCM.Text

    txtNamaPasien.Text = rs("Nama Lengkap")
    txtSex.Text = IIf(rs("JK") = "L", "Laki-Laki", "Perempuan")
    txtThn.Text = rs("UmurTahun")
    txtBln.Text = rs("UmurBulan")
    txtHr.Text = rs("UmurHari")
    txtAlamat.Text = IIf(IsNull(rs("Alamat")), "", rs("Alamat"))
    txtTglDaftar.Text = rs("TglDaftarMembership")

    Call subLoadRiwayatCatatanKlinis
    Call subLoadRiwayatCatatanMedis
    Call subLoadRiwayatDiagnosa
    Call subLoadRiwayatHasilPemeriksaan
    Call subLoadRiwayatKecelakaan
    Call subLoadRiwayatKonsul
    Call subLoadRiwayatOperasi
    Call subLoadRiwayatPemeriksaan(True)

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

