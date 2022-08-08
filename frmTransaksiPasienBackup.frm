VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTransaksiPasienBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pemeriksaan Pasien"
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
   Icon            =   "frmTransaksiPasienBackup.frx":0000
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
      TabIndex        =   61
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
      TabIndex        =   58
      Top             =   2040
      Width           =   14655
      Begin VB.TextBox txtGrandTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   45
         Top             =   6000
         Width           =   2415
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   12480
         TabIndex        =   46
         Top             =   6000
         Width           =   2055
      End
      Begin TabDlg.SSTab sstTP 
         Height          =   5535
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   9763
         _Version        =   393216
         Tabs            =   10
         Tab             =   1
         TabsPerRow      =   10
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
         TabPicture(0)   =   "frmTransaksiPasienBackup.frx":0CCA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtTindakanTotal"
         Tab(0).Control(1)=   "cmdUbahPT"
         Tab(0).Control(2)=   "cmdHapusDataPT"
         Tab(0).Control(3)=   "cmdTambahPT"
         Tab(0).Control(4)=   "dgTindakan"
         Tab(0).Control(5)=   "Label1"
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Pemakaian &Obat && Alkes"
         TabPicture(1)   =   "frmTransaksiPasienBackup.frx":0CE6
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "dgObatAlkes"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cmdTambahPOA"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmdHapusDataPOA"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "txtAlkesTotal"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "Riwayat Catatan Klinis"
         TabPicture(2)   =   "frmTransaksiPasienBackup.frx":0D02
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdHapusCatataKlinis"
         Tab(2).Control(1)=   "cmdTambahCatatanKlinis"
         Tab(2).Control(2)=   "cmdKehamilandanKB"
         Tab(2).Control(3)=   "dgCatatanKlinis"
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "Riwayat Catatan Medis"
         TabPicture(3)   =   "frmTransaksiPasienBackup.frx":0D1E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "dgCatatanMedis"
         Tab(3).Control(1)=   "cmdTambahCatatanMedis"
         Tab(3).Control(2)=   "cmdHapusCatatanMedis"
         Tab(3).Control(3)=   "cmdCetakCatatanMedis"
         Tab(3).ControlCount=   4
         TabCaption(4)   =   "Riwayat &Diagnosa"
         TabPicture(4)   =   "frmTransaksiPasienBackup.frx":0D3A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "cmdDelDiagnosa"
         Tab(4).Control(1)=   "cmdCetakDiagnosa"
         Tab(4).Control(2)=   "cmdTambahDiagnosa"
         Tab(4).Control(3)=   "dgRiwayatDiagnosa"
         Tab(4).ControlCount=   4
         TabCaption(5)   =   "Riwayat Operasi"
         TabPicture(5)   =   "frmTransaksiPasienBackup.frx":0D56
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "dgRiwayatOperasi"
         Tab(5).Control(1)=   "cmdTambahOperasi"
         Tab(5).ControlCount=   2
         TabCaption(6)   =   "Riwayat Konsul"
         TabPicture(6)   =   "frmTransaksiPasienBackup.frx":0D72
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "dgKonsul"
         Tab(6).Control(1)=   "Command1"
         Tab(6).Control(2)=   "cmdTambahKonsul"
         Tab(6).Control(3)=   "cmdHapusKonsul"
         Tab(6).ControlCount=   4
         TabCaption(7)   =   "Riwayat Kecelakaan"
         TabPicture(7)   =   "frmTransaksiPasienBackup.frx":0D8E
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "dgKecelakaan"
         Tab(7).Control(1)=   "cmdTambahKecelakaan"
         Tab(7).Control(2)=   "cmdHapusKecelakaan"
         Tab(7).ControlCount=   3
         TabCaption(8)   =   "Riwayat Peme&riksaan"
         TabPicture(8)   =   "frmTransaksiPasienBackup.frx":0DAA
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "dgRiwayatPemeriksaan"
         Tab(8).Control(1)=   "cmdCetakRP"
         Tab(8).Control(2)=   "chkRP"
         Tab(8).ControlCount=   3
         TabCaption(9)   =   "Riwayat Hasil Pemeriksaan"
         TabPicture(9)   =   "frmTransaksiPasienBackup.frx":0DC6
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "cmdCetakResume"
         Tab(9).Control(1)=   "cmdCetakHasilPemeriksaan"
         Tab(9).Control(2)=   "dgHasilPemeriksaan"
         Tab(9).ControlCount=   3
         Begin VB.CommandButton cmdCetakResume 
            Caption         =   "Cetak &Resume"
            Height          =   375
            Left            =   -64080
            TabIndex        =   66
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdDelDiagnosa 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   -64080
            TabIndex        =   30
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetakCatatanMedis 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -65760
            TabIndex        =   25
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusKonsul 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   -64080
            TabIndex        =   36
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahKonsul 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   -62400
            TabIndex        =   37
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusCatatanMedis 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   -64080
            TabIndex        =   26
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahCatatanMedis 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   -62400
            TabIndex        =   27
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusCatataKlinis 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   -64080
            TabIndex        =   22
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahCatatanKlinis 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   -62400
            TabIndex        =   23
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CheckBox chkRP 
            Caption         =   "Tampilkan Semua"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -74760
            TabIndex        =   63
            Top             =   4965
            Width           =   1815
         End
         Begin VB.CommandButton cmdCetakRP 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -62400
            TabIndex        =   42
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetakDiagnosa 
            Caption         =   "&Cetak"
            Height          =   345
            Left            =   -65745
            TabIndex        =   29
            Top             =   5070
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahDiagnosa 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   -62400
            TabIndex        =   31
            Top             =   5040
            Width           =   1575
         End
         Begin VB.TextBox txtAlkesTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   3360
            TabIndex        =   17
            Top             =   4965
            Width           =   2415
         End
         Begin VB.TextBox txtTindakanTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   -72000
            TabIndex        =   12
            Top             =   4965
            Width           =   2415
         End
         Begin VB.CommandButton cmdTambahOperasi 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   -62400
            TabIndex        =   33
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusKecelakaan 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   -64080
            TabIndex        =   39
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahKecelakaan 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   -62400
            TabIndex        =   40
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetakHasilPemeriksaan 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -62400
            TabIndex        =   44
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdUbahPT 
            Caption         =   "&Ubah Data"
            Height          =   375
            Left            =   -65760
            TabIndex        =   13
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusDataPT 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   -64080
            TabIndex        =   14
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahPT 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   -62400
            TabIndex        =   15
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusDataPOA 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   11040
            TabIndex        =   18
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahPOA 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   12600
            TabIndex        =   19
            Top             =   5040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -65760
            TabIndex        =   35
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdKehamilandanKB 
            Caption         =   "&Data Kehamilan dan KB"
            Height          =   375
            Left            =   -66480
            TabIndex        =   21
            Top             =   5040
            Width           =   2295
         End
         Begin MSDataGridLib.DataGrid dgTindakan 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   11
            Top             =   720
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
            Left            =   240
            TabIndex        =   16
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
            TabIndex        =   28
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
         Begin MSDataGridLib.DataGrid dgRiwayatPemeriksaan 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   41
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
            TabIndex        =   20
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
         Begin MSDataGridLib.DataGrid dgCatatanMedis 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   24
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
         Begin MSDataGridLib.DataGrid dgRiwayatOperasi 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   32
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
         Begin MSDataGridLib.DataGrid dgKonsul 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   34
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
         Begin MSDataGridLib.DataGrid dgKecelakaan 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   38
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
         Begin MSDataGridLib.DataGrid dgHasilPemeriksaan 
            Height          =   3975
            Left            =   -74760
            TabIndex        =   43
            Top             =   840
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   7011
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
            Left            =   240
            TabIndex        =   65
            Top             =   5025
            Width           =   2925
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Biaya Pelayanan Tindakan"
            Height          =   210
            Left            =   -74760
            TabIndex        =   64
            Top             =   5025
            Width           =   2550
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya Pelayanan"
         Height          =   210
         Left            =   360
         TabIndex        =   60
         Top             =   6000
         Width           =   1755
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
      TabIndex        =   47
      Top             =   960
      Width           =   14655
      Begin VB.TextBox txtKls 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   9840
         TabIndex        =   7
         Top             =   600
         Width           =   1455
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
         Left            =   7320
         TabIndex        =   48
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
            TabIndex        =   6
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
            TabIndex        =   5
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
            TabIndex        =   4
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Left            =   2130
            TabIndex        =   51
            Top             =   270
            Width           =   165
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Left            =   1350
            TabIndex        =   50
            Top             =   277
            Width           =   240
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Left            =   550
            TabIndex        =   49
            Top             =   277
            Width           =   285
         End
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   6000
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtJenisPasien 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   11400
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtTglDaftar 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   12840
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Pelayanan"
         Height          =   210
         Left            =   9840
         TabIndex        =   59
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   120
         TabIndex        =   57
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1560
         TabIndex        =   56
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   2640
         TabIndex        =   55
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   6000
         TabIndex        =   54
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pasien"
         Height          =   210
         Left            =   11400
         TabIndex        =   53
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pendaftaran"
         Height          =   210
         Left            =   12840
         TabIndex        =   52
         Top             =   360
         Width           =   1365
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   62
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
      Picture         =   "frmTransaksiPasienBackup.frx":0DE2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmTransaksiPasienBackup.frx":1B6A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTransaksiPasienBackup.frx":452B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmTransaksiPasienBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim subKdDokterTemp As String
Dim intJumlahPrint  As Integer
Dim bolStatusFIFO As Boolean
'Private Sub chkRD_Click()
'    If chkRD.Value = Checked Then
'        Call subLoadRiwayatDiagnosa(True)
'    Else
'        Call subLoadRiwayatDiagnosa(False)
'    End If
'End Sub

Private Sub chkRP_Click()
    If chkRP.value = 0 Then
        subLoadRiwayatPemeriksaan False
    Else
        subLoadRiwayatPemeriksaan True
    End If
End Sub

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
'    cmdCetakCatatanMedis.Enabled = True
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
            
        Case Else
            Call subLoadDiagramOdonto
    End Select

    cmdCetakHasilPemeriksaan.Enabled = True
    
    End If

Exit Sub
errLoad:
'    Call msubPesanError
End Sub

Private Sub cmdCetakResume_Click()
On Error Resume Next
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakDataRiwayatPemeriksaanPasien.Show
End Sub

Private Sub cmdDelDiagnosa_Click()
On Error GoTo errHapusData

Dim vbMsgboxRslt As VbMsgBoxResult
Dim dbcmd As New ADODB.Command

    If dgRiwayatDiagnosa.ApproxCount = 0 Then Exit Sub

    '***agar bisa dihapus ceunah
'    If dgRiwayatDiagnosa.Columns("Ruang Periksa").Value <> mstrNamaRuangan Then
'        MsgBox "Diagnosa yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
'        Exit Sub
'    End If
    'periksa dulu diagnosa utama atau bukan 2010-01-04
    If dgRiwayatDiagnosa.Columns("KdJenisDiagnosa").value = "05" Then
        'periksa dulu diagnosa dibawahnya jika tidak ada baru dieksekusi
        If rsdiagnosa.RecordCount = 1 Then
            MsgBox "Periksa terlebih dahulu " & vbCrLf & "Jenis Diagnosa yang akan dihapus", vbCritical + vbInformation
            Exit Sub
        End If
    End If
    If vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus data diagnosa ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub

    sp_DelDiagnosa dbcmd
    subLoadRiwayatDiagnosa (False)
    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"

Exit Sub
errHapusData:
    MsgBox "Pilih diagnosa yang akan dihapus." & vbCrLf & "Jika berlanjut hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusCatataKlinis_Click()
On Error GoTo errHapusData

Dim vbMsgboxRslt As VbMsgBoxResult
Dim dbcmd As New ADODB.Command
    
    If dgCatatanKlinis.ApproxCount = 0 Then Exit Sub
    
'    If dgCatatanKlinis.Columns("Ruang Pemeriksaan").value <> mstrNamaRuangan Then
'        MsgBox "Catatan klinis yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
'        Exit Sub
'    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus catatan klinis pasien '" _
        & dgCatatanKlinis.Columns("NoPendaftaran").value & "'" & vbNewLine _
        & "Dengan tanggal periksa '" & dgCatatanKlinis.Columns("TglPeriksa").value _
        & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    sp_DelBiayaCatatanKlinis dbcmd
    Call subLoadRiwayatCatatanKlinis
    MsgBox "Penghapusan data berhasil", vbInformation, "Informasi"

Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusCatatanMedis_Click()
On Error GoTo errHapusData

Dim vbMsgboxRslt As VbMsgBoxResult
Dim dbcmd As New ADODB.Command
    
    If dgCatatanMedis.ApproxCount = 0 Then Exit Sub
    
'    If dgCatatanMedis.Columns("RuangPemeriksaan").value <> mstrNamaRuangan Then
'        MsgBox "Catatan medis yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
'        Exit Sub
'    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus catatan medis pasien '" _
        & dgCatatanMedis.Columns("NoPendaftaran").value & "'" & vbNewLine _
        & "Dengan tanggal periksa '" & dgCatatanMedis.Columns("TglPeriksa").value _
        & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    sp_DelBiayaCatatanMedis dbcmd
    Call subLoadRiwayatCatatanMedis
    MsgBox "Penghapusan data berhasil", vbInformation, "Informasi"

Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusDataPOA_Click()
Dim adoCommand As New ADODB.Command
Dim vbMsgboxRslt As VbMsgBoxResult
    On Error GoTo errHapusData
    
    If dgObatAlkes.ApproxCount = 0 Then Exit Sub
    If dgObatAlkes.Columns("Status Bayar").value = "Sudah DiBayar" Then
        MsgBox "Pemakaian Obat dan Alkes yang sudah dibayar tidak dapat dihapus", vbCritical, "Validasi"
        Exit Sub
    ElseIf dgObatAlkes.Columns("Ruang Pelayanan").value <> mstrNamaRuanganPasien Then
        MsgBox "Pemakaian Obat dan Alkes yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
        Exit Sub
    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus pemakaian obat dan alkes '" _
        & dgObatAlkes.Columns("NamaBarang").value & "'" & vbNewLine _
        & "Dengan tanggal pelayanan '" & dgObatAlkes.Columns("TglPelayanan").value _
        & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, dgObatAlkes.Columns("KdBarang").value)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, dgObatAlkes.Columns("KdAsal").value)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuanganPasien)
        .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, dgObatAlkes.Columns("Sat").value)
        .Parameters.Append .CreateParameter("JmlBrg", adInteger, adParamInput, , dgObatAlkes.Columns("Jml").value)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dgObatAlkes.Columns("TglPelayanan").value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        
        .ActiveConnection = dbConn
        .CommandText = "Delete_PemakaianObatAlkes"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam penghapusan data pemakaian obat dan alkes", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            Exit Sub
        Else
            Call Add_HistoryLoginActivity("Delete_PemakaianObatAlkes")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Call subpemakaianobatalkes
    MsgBox "Penghapusan data berhasil", vbInformation, "Informasi"
    
    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusDataPT_Click()
Dim vbMsgboxRslt As VbMsgBoxResult
Dim dbcmd As New ADODB.Command
    On Error GoTo errHapusData
    
    If dgTindakan.ApproxCount = 0 Then Exit Sub
    
    If dgTindakan.Columns("Status Bayar").value = "Sudah DiBayar" Then
        MsgBox "Pelayanan yang sudah dibayar tidak dapat dihapus", vbCritical, "Validasi"
        Exit Sub
'    ElseIf dgTindakan.Columns("Ruang Pelayanan").Value <> mstrNamaRuanganPasien Then
'        MsgBox "Pelayanan yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
'        Exit Sub
    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus pelayanan '" _
        & dgTindakan.Columns("NamaPelayanan").value & "'" & vbNewLine _
        & "Dengan tanggal pelayanan '" & dgTindakan.Columns("TglPelayanan").value _
        & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub
    
    sp_DelBiayaPelayanan dbcmd
    Call subLoadPelayananDidapat
    MsgBox "Penghapusan data berhasil", vbInformation, "Informasi"
    
    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusKecelakaan_Click()
On Error GoTo errHapusData

Dim vbMsgboxRslt As VbMsgBoxResult
Dim dbcmd As New ADODB.Command
    
    If dgKecelakaan.ApproxCount = 0 Then Exit Sub
    
    If dgKecelakaan.Columns("Ruangan Pemeriksa").value <> mstrNamaRuangan Then
        MsgBox "Catatan klinis yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
        Exit Sub
    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus data kecelakaan '" _
        & dgKecelakaan.Columns("NoPendaftaran").value & "'" & vbNewLine _
        & "Dengan tanggal periksa '" & dgKecelakaan.Columns("TglPeriksa").value _
        & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    sp_DelKecelakaan dbcmd
    Call subLoadRiwayatKecelakaan
    MsgBox "Penghapusan data berhasil", vbInformation, "Informasi"

Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusKonsul_Click()
On Error GoTo errHapusData

Dim vbMsgboxRslt As VbMsgBoxResult
Dim dbcmd As New ADODB.Command
    
    If dgKonsul.ApproxCount = 0 Then Exit Sub
    
'    If dgKonsul.Columns("Ruangan Perujuk").Value <> mstrNamaRuangan Then
'        MsgBox "Catatan medis yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
'        Exit Sub
'    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus catatan medis pasien '" _
        & dgKonsul.Columns("NoPendaftaran").value & "'" & vbNewLine _
        & "Dengan tanggal periksa '" & dgKonsul.Columns("TglDirujuk").value _
        & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    sp_DelKonsul dbcmd
    Call subLoadRiwayatKonsul
    MsgBox "Penghapusan data berhasil", vbInformation, "Informasi"

Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

'Private Sub cmdICD9_Click()
'On Error GoTo hell
'
'    If dgRiwayatDiagnosa.ApproxCount = 0 Then Exit Sub
'    If dgRiwayatDiagnosa.Columns(0) <> txtNoPendaftaran.Text Then
'        MsgBox "No Pendaftaran tidak sama, mohon isi diagnosanya [ICD 10] dahulu", vbExclamation, "Validasi"
'        cmdTambahDiagnosa.SetFocus
'        Exit Sub
'    End If
'    Set rs = Nothing
'    rs.Open "Select KdJenisDiagnosa From JenisDiagnosa Where KdJenisDiagnosa = '" & dgRiwayatDiagnosa.Columns(14) & "'", dbConn, adOpenForwardOnly, adLockReadOnly
'    If rs.EOF = True Then
'        MsgBox dgRiwayatDiagnosa.Columns(3) & " tidak terdapat di ruangan " & mstrNamaRuangan, vbExclamation, "Validasi"
'        cmdTambahDiagnosa.SetFocus
'        Exit Sub
'    End If
''    Me.Enabled = False
'    mstrKdDiagnosa = ""
'    mstrKdDiagnosa = dgRiwayatDiagnosa.Columns(4)
'    mstrKdJenisDiagnosaTindakan = ""
'    mstrKdJenisDiagnosaTindakan = dgRiwayatDiagnosa.Columns(15)
'    bolEditDiagnosa = True
'    With frmPeriksaDiagnosa
'        .Show
'        .txtNoPendaftaran = txtNoPendaftaran.Text
'        .txtNoCM = txtNoCM.Text
'        .txtNamaPasien = txtNamaPasien.Text
'        If Left(.txtSex, 1) = "P" Then
'            .txtSex.Text = "Perempuan"
'        Else
'            .txtSex.Text = "Laki-laki"
'        End If
'        .txtThn = txtThn.Text
'        .txtBln = txtBln.Text
'        .txtHari = txtHr.Text
'
'        subKdDokterTemp = mstrKdDokter
'        .txtDokter = mstrNamaDokter
'        mstrKdDokter = subKdDokterTemp
'        .fraDokter.Visible = False
'
'        .dtpTglPeriksa.Value = dgRiwayatDiagnosa.Columns(2)
'        .dcJenisDiagnosa.BoundText = dgRiwayatDiagnosa.Columns(14)
'        .dcJenisDiagnosaTindakan.BoundText = dgRiwayatDiagnosa.Columns(15)
'
'        .dcJenisDiagnosa.Enabled = False
'        .lvwDiagnosa.Enabled = False
'        .txtNamaDiagnosa.Enabled = False
'        .txtDokter.Enabled = False
'        .dtpTglPeriksa.Enabled = False
'        .chkICD9.Value = Checked
'
'    End With
'Exit Sub
'hell:
'    Call msubPesanError
'    frmPeriksaDiagnosa.Show
'
'
'End Sub

Private Sub cmdKehamilandanKB_Click()
    frmDataKehamilandanKB.Show
End Sub

Private Sub cmdTambahCatatanKlinis_Click()
    If txtNoCM.Text = "" Then Exit Sub
    frmCatatanKlinisPasien.Show
    frmTransaksiPasien.Enabled = False
End Sub

Private Sub cmdTambahCatatanMedis_Click()
On Error GoTo errLoad
    If txtNoCM.Text = "" Then Exit Sub
    frmTransaksiPasien.Enabled = False
    
    With frmCatatanMedikPasien
        subKdDokterTemp = mstrKdDokter
        .txtDokter = mstrNamaDokter
        mstrKdDokter = subKdDokterTemp
        .fraDokter.Visible = False
        .Show
    End With
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTambahDiagnosa_Click()
On Error GoTo errLoad
    
    Me.Enabled = False
    With frmPeriksaDiagnosa
        .Show
        .txtNamaFormPengirim.Text = Me.Name
        .txtNoPendaftaran = txtNoPendaftaran.Text
        .txtNoCM = txtNoCM.Text
        .txtNamaPasien = txtNamaPasien.Text
        If Left(.txtSex, 1) = "P" Then
            .txtSex.Text = "Perempuan"
        Else
            .txtSex.Text = "Laki-laki"
        End If
        .txtThn = txtThn.Text
        .txtBln = txtBln.Text
        .txtHari = txtHr.Text
    
        subKdDokterTemp = mstrKdDokter
        .txtDokter = mstrNamaDokter
        mstrKdDokter = subKdDokterTemp
        .fraDokter.Visible = False
        .Show
    End With
    
    frmPeriksaDiagnosa.Show
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTambahKecelakaan_Click()
    If txtNoCM.Text = "" Then Exit Sub
    frmPasienGDKecelakaan.Show
    frmTransaksiPasien.Enabled = False
End Sub

Private Sub cmdTambahKonsul_Click()
On Error GoTo errLoad
    
    If txtNoCM.Text = "" Then Exit Sub
    
    frmPasienRujukan.Show
    With frmPasienRujukan
        .txtNoPendaftaran.Text = txtNoPendaftaran.Text
        .txtNoCM.Text = txtNoCM.Text
        .txtNamaPasien.Text = txtNamaPasien.Text
        .txtSex.Text = txtSex.Text
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHari.Text = txtHr.Text
'        .dtpTglDirujuk.value = Now
        strSQL = "SELECT KdSubInstalasi, IdDokter, Dokter FROM V_DokterPasien WHERE NoPendaftaran = '" & txtNoPendaftaran.Text & "'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            mstrKdSubInstalasi = rs("KdSubInstalasi").value
            mstrKdDokter = rs("IdDokter").value
            intJmlDokter = rs.RecordCount
            frmPasienRujukan.txtDokter.Text = rs("Dokter").value
            frmPasienRujukan.fraDokter.Visible = False
        End If
    End With
    
    Me.Enabled = False
    frmPasienRujukan.Show
    
Exit Sub
errLoad:
    Call msubPesanError
    frmPasienRujukan.Show
End Sub

Private Sub cmdTambahOperasi_Click()
    On Error GoTo errTO
    Me.Enabled = False
    
    With frmTindakanOperasi
        .Show
        .txtNoPendaftaran.Text = mstrNoPen
        .txtNoIBS.Text = mstrNoIBS
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = txtNamaPasien.Text
        If Left(txtSex.Text, 1) = "P" Then
           .txtJK.Text = "Perempuan"
        Else
           .txtJK.Text = "Laki-Laki"
        End If
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHr.Text
        .txtJenisOperasi.Text = mstrJenisOperasi
    End With
    Exit Sub
errTO:
    Call msubPesanError
    Me.Enabled = True
End Sub

Private Sub cmdTambahPOA_Click()
On Error GoTo errLoad
    
    strSQL = "SELECT dbo.RegistrasiRJ.IdDokter, dbo.DataPegawai.NamaLengkap " & _
        " FROM dbo.RegistrasiRJ INNER JOIN dbo.DataPegawai ON dbo.RegistrasiRJ.IdDokter = dbo.DataPegawai.IdPegawai " & _
        " WHERE (dbo.RegistrasiRJ.NoPendaftaran = '" & txtNoPendaftaran.Text & "')"
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

Private Sub cmdTambahPT_Click()
On Error GoTo errLoad
    
    strSQL = "SELECT dbo.RegistrasiRJ.IdDokter, dbo.DataPegawai.NamaLengkap " & _
        " FROM dbo.RegistrasiRJ INNER JOIN dbo.DataPegawai ON dbo.RegistrasiRJ.IdDokter = dbo.DataPegawai.IdPegawai " & _
        " WHERE (dbo.RegistrasiRJ.NoPendaftaran = '" & txtNoPendaftaran.Text & "')"
    Call msubRecFO(rs, strSQL)
    
    If Not rs.EOF Then
        frmTindakan.txtDokter.Text = rs(1).value
        mstrKdDokter = rs(0).value
        intJmlDokter = rs.RecordCount
        frmTindakan.fraDokter.Visible = False
    End If
    
    frmTindakan.Show
    
Exit Sub
errLoad:
    Call msubPesanError
    frmTindakan.Show
End Sub

Private Sub cmdTutup_Click()
  Unload Me
End Sub

Private Sub cmdUbahPT_Click()
On Error GoTo errLoad

    If dgTindakan.ApproxCount = 0 Then Exit Sub
    If dgTindakan.Columns("Status Bayar").value = "Sudah DiBayar" Then
        MsgBox "Pelayanan yang sudah dibayar tidak dapat diubah", vbCritical, "Validasi"
        Exit Sub
'    ElseIf dgTindakan.Columns("Ruang Pelayanan").Value <> mstrNamaRuanganPasien Then
'        MsgBox "Pelayanan yang didapat di ruangan lain tidak dapat diubah di ruangan ini", vbCritical
'        Exit Sub
    End If
    With frmUpdateBiayaPelayanan
        .txtNoPendaftaran = txtNoPendaftaran.Text
        strKodePelayananRS = dgTindakan.Columns(12).value
        Call .txtNoPendaftaran_KeyPress(13)
    End With

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Command1_Click()
On Error GoTo errLoad

    If dgKonsul.ApproxCount = 0 Then
    Exit Sub
    Else
    If intJumlahPrint = 0 Then
        intJumlahPrint = 1
        mstrNoPen = txtNoPendaftaran.Text
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

Private Sub dgCatatanKlinis_Click()
'WheelHook.WheelUnHook
'        Set MyProperty = dgCatatanKlinis
'        WheelHook.WheelHook dgCatatanKlinis
End Sub

Private Sub dgCatatanKlinis_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then cmdTambahCatatanKlinis.SetFocus
End Sub

Private Sub dgCatatanMedis_Click()
'WheelHook.WheelUnHook
'        Set MyProperty = dgCatatanMedis
'        WheelHook.WheelHook dgCatatanMedis
End Sub

Private Sub dgCatatanMedis_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then cmdTambahCatatanMedis.SetFocus
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

Private Sub dgKonsul_Click()
'WheelHook.WheelUnHook
'        Set MyProperty = dgKonsul
'        WheelHook.WheelHook dgKonsul
End Sub

Private Sub dgKonsul_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then cmdTambahKonsul.SetFocus
End Sub

Private Sub dgObatAlkes_Click()
'WheelHook.WheelUnHook
'        Set MyProperty = dgObatAlkes
'        WheelHook.WheelHook dgObatAlkes
End Sub

Private Sub dgObatAlkes_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then cmdTambahPOA.SetFocus
End Sub

Private Sub dgRiwayatDiagnosa_Click()
'WheelHook.WheelUnHook
'        Set MyProperty = dgRiwayatDiagnosa
'        WheelHook.WheelHook dgRiwayatDiagnosa
End Sub

Private Sub dgRiwayatDiagnosa_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then cmdTambahDiagnosa.SetFocus
End Sub

Private Sub dgRiwayatOperasi_Click()
WheelHook.WheelUnHook
        Set MyProperty = dgRiwayatOperasi
        WheelHook.WheelHook dgRiwayatOperasi
End Sub

'Private Sub dgRiwayatMedik_KeyPress(KeyAscii As Integer)
'On Error Resume Next
'    If KeyAscii = 13 Then cmdTambahCM.SetFocus
'End Sub

Private Sub dgRiwayatOperasi_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then cmdTambahOperasi.SetFocus
End Sub

Private Sub dgRiwayatPemeriksaan_Click()
WheelHook.WheelUnHook
        Set MyProperty = dgRiwayatPemeriksaan
        WheelHook.WheelHook dgRiwayatPemeriksaan
End Sub

Private Sub dgRiwayatPemeriksaan_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then cmdCetakRP.SetFocus
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
'            Call subLoadPelayananDidapat
'            Call subPemakaianObatAlkes
'            Call subLoadRiwayatCatatanKlinis
'            Call subLoadRiwayatCatatanMedis
            Call subLoadRiwayatDiagnosa(False)
'            Call subLoadRiwayatKecelakaan
            Call subLoadRiwayatOperasi
'           Call subLoadRiwayatKonsul
'            Call subLoadRiwayatPemeriksaan(False)
'            Call subLoadRiwayatHasilPemeriksaan
    End Select
     

Exit Sub
errLoad:
    Call msubPesanError
End Sub
'--------dirubah tgl 2009-09-08
Private Sub Form_Load()
On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call subLoadPelayananDidapat
    Call subpemakaianobatalkes
    Call subLoadRiwayatCatatanKlinis
    Call subLoadRiwayatCatatanMedis
    Call subLoadRiwayatDiagnosa(False)
    Call subLoadRiwayatKecelakaan
    Call subLoadRiwayatOperasi
    Call subLoadRiwayatKonsul
    Call subLoadRiwayatPemeriksaan(False)
    Call subLoadRiwayatHasilPemeriksaan
        
'    'hanya untuk rekam medis
'    sstTP.TabVisible(0) = False
'    sstTP.TabVisible(1) = False

    cmdHapusKecelakaan.Visible = False
    cmdTambahKecelakaan.Visible = False
'    cmdHapusKonsul.Visible = False
'    cmdTambahKonsul.Visible = False
    cmdTambahOperasi.Visible = False
    
'    sstTP.TabsPerRow = 8
    sstTP.Tab = 0

'    If mblnAdmin = True Then
'        cmdHapusDataPT.Enabled = True
'        cmdHapusDataPOA.Enabled = True
'        cmdHapusCatataKlinis.Enabled = True
'        cmdHapusCatatanMedis.Enabled = True
'        cmdHapusKonsul.Visible = True
'    Else
'        cmdHapusDataPT.Enabled = False
'        cmdHapusDataPOA.Enabled = False
'        cmdHapusCatataKlinis.Enabled = False
'        cmdHapusCatatanMedis.Enabled = False
'        cmdHapusKonsul.Visible = True
'    End If
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Public Sub subLoadRiwayatDiagnosa(blnAll As Boolean)
On Error GoTo hell
    If blnAll = False Then
        strSQL = "Select * from V_DaftarDiagnosaPasien where nocm = '" & mstrNoCM & "' AND NoPendaftaran = '" & mstrNoPen & "'"
    Else
        strSQL = "Select * from V_DaftarDiagnosaPasien where nocm = '" & mstrNoCM & "'"
    End If
    'Call msubRecFO(rsDiagnosa, strSQL)
    Set rsdiagnosa = Nothing
    rsdiagnosa.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgRiwayatDiagnosa.DataSource = rsdiagnosa
    With dgRiwayatDiagnosa
        .Columns(0).Width = 1500 'NoPendaftaran
        .Columns(1).Width = 0 'NoCM
        .Columns(2).Width = 1590 'TglPeriksa
        .Columns(3).Width = 2000 'JenisDiagnosa
        .Columns(4).Width = 1100 'KdDiagnosa ICD 10
        .Columns(4).Caption = "Kode ICD 10"
        .Columns(5).Width = 2700 'Diagnosa ICD 10
        .Columns(5).Caption = "Diagnosa ICD 10"
        .Columns(6).Width = 2500 'JenisDiagnosa
        .Columns(6).Caption = "JenisDiagnosaTindakan"
        .Columns(7).Width = 1000 'KdDiagnosaTindakan ICD 9
        .Columns(7).Caption = "Kode ICD 9"
        .Columns(8).Width = 2700 'DiagnosaTindakan ICD 9
        .Columns(8).Caption = "Diagnosa Tindakan ICD 9"
        .Columns(9).Width = 2000 '[Ruang Periksa]
        .Columns(10).Width = 2700 '[Dokter Pemeriksa]
        .Columns(11).Width = 0 '[Nama Pasien]
        .Columns(12).Width = 0 'JK
        .Columns(13).Width = 0 'Umur
        .Columns(14).Width = 0 'KdJnsDiagnosa
        .Columns(15).Width = 0 'KdJnsDiagnosaTindakan
        .Columns(16).Width = 0 'IdDokterPemeriksa
    End With
    mstrDiagnosaUtama = ""
    If rsdiagnosa.RecordCount > 0 Then
        rsdiagnosa.MoveFirst
        Do While Not rsdiagnosa.EOF
            If rsdiagnosa("kdJenisDiagnosa").value = "05" Then
                mstrDiagnosaUtama = rsdiagnosa("kdDiagnosa").value
                mstrTglDiagnosaUtama = rsdiagnosa("tglPeriksa").value
            End If
            rsdiagnosa.MoveNext
        Loop
    End If
Exit Sub
hell:
    Call msubPesanError
End Sub
'Untuk meload riwayat pemeriksaan yang sudah pernah didapat
Public Sub subLoadRiwayatPemeriksaan(blnAll As Boolean)
On Error GoTo errLoad
    If blnAll = False Then
        strSQL = "Select * from V_RiwayatPemeriksaanPasien where nocm = '" & mstrNoCM & "' AND KdRuangan='" & mstrKdRuangan & "'"
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
Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Untuk meload riwayat hasil pemeriksaan yang sudah pernah didapat
Public Sub subLoadRiwayatHasilPemeriksaan()
On Error GoTo errLoad
    strSQL = "Select NoLab_Rad, [Ruang Pemeriksa], [Dokter Pemeriksa], TglPendaftaran, TglHasil, [Asal Rujukan], [Ruangan Perujuk], [Dokter Perujuk], KdInstalasi from V_RiwayatHasilPemeriksaan where nocm = '" & mstrNoCM & "'"
    msubRecFO rs, strSQL
    Set dgHasilPemeriksaan.DataSource = rs
    dgHasilPemeriksaan.Columns("KdInstalasi").Width = 0
Exit Sub
errLoad:
    Call msubPesanError
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
'            Case 10 ' pemakaian bahan
'                dgPemakaianBahan.SetFocus
        End Select
    End If
End Sub

Private Sub txtNoPendaftaran_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then sstTP.SetFocus
End Sub

'Untuk meload pelayanan yang sudah pernah didapat
Public Sub subLoadPelayananDidapat()
On Error GoTo errLoad
strSQL = "SELECT TglPelayanan,JenisPelayanan,NamaPelayanan,NamaRuangan AS [Ruang Pelayanan]," _
        & "Kelas,JenisTarif,CITO,JmlPelayanan as Jml,Total as Tarif,BiayaTotal," _
        & "DokterPemeriksa,[Status Bayar],KdPelayananRS,Operator FROM V_BiayaPelayananTindakan WHERE " _
        & "NoPendaftaran='" & mstrNoPen & "' ORDER BY TglPelayanan"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgTindakan.DataSource = rs
    With dgTindakan
        .Columns(0).Width = 1600
        .Columns(1).Width = 2000
        .Columns(2).Width = 2000
        .Columns(3).Width = 1600
        .Columns(4).Width = 900
        .Columns(5).Width = 1000
        .Columns(6).Width = 500
        .Columns(7).Width = 400
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 900
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 1000
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 2400
        .Columns(11).Width = 1200
        .Columns(12).Width = 0
        .Columns(13).Width = 2000
        
        .Columns("Tarif").NumberFormat = "#,###"
        .Columns("BiayaTotal").NumberFormat = "#,###"
    End With
    
    strSQL = "SELECT sum(BiayaTotal) as TotalBayar FROM V_BiayaPelayananTindakan " _
        & "WHERE NoPendaftaran='" _
        & mstrNoPen & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount <> 0 Then
        txtTindakanTotal.Text = FormatCurrency(rs.Fields(0).value, 2)
    Else
        txtTindakanTotal.Text = FormatCurrency(0, 2)
    End If
    If txtAlkesTotal.Text = "" Then
        txtAlkesTotal.Text = 0
        txtAlkesTotal.Text = FormatCurrency(txtAlkesTotal.Text, 2)
    End If
    If txtTindakanTotal.Text = "" Then txtTindakanTotal.Text = 0
    If txtAlkesTotal.Text = "" Then txtAlkesTotal.Text = 0
    txtGrandTotal.Text = CCur(txtTindakanTotal.Text) + CCur(txtAlkesTotal)
    txtGrandTotal.Text = FormatCurrency(txtGrandTotal.Text, 2)
Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Untuk meload pemakaian alkes yang sudah pernah didapat
Public Sub subpemakaianobatalkes()
On Error GoTo errLoad
    strSQL = "SELECT TglPelayanan,[Detail Jenis Brg],NamaBarang," _
        & "NamaRuangan AS [Ruang Pelayanan],Kelas,JenisTarif,SatuanJml as Sat," _
        & "JmlBarang as Jml,HargaSatuan as Tarif,BiayaTotal,DokterPemeriksa," _
        & "[Status Bayar],KdBarang,KdAsal,Operator " _
        & "FROM V_BiayaPemakaianObatAlkes WHERE NoPendaftaran='" _
        & mstrNoPen & "' ORDER BY TglPelayanan"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgObatAlkes.DataSource = rs
    With dgObatAlkes
        .Columns(0).Width = 1600
        .Columns(1).Width = 2000
        .Columns(2).Width = 2000
        .Columns(3).Width = 1600
        .Columns(4).Width = 900
        .Columns(5).Width = 1000
        .Columns(6).Width = 400
        .Columns(7).Width = 400
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 900
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 1000
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 2400
        .Columns(11).Width = 1200
        .Columns(12).Width = 0
        .Columns(13).Width = 0
        .Columns(14).Width = 2000
    End With
    
    strSQL = "SELECT sum(BiayaTotal) as TotalBayar FROM V_BiayaPemakaianObatAlkes " _
        & "WHERE NoPendaftaran='" _
        & mstrNoPen & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount <> 0 Then
        txtAlkesTotal.Text = FormatCurrency(rs.Fields(0).value, 2)
        If IsNull(rs.Fields(0)) = True Then txtAlkesTotal.Text = FormatCurrency(0, 2)
    Else
        txtAlkesTotal.Text = FormatCurrency(0, 2)
    End If
    If txtTindakanTotal.Text = "" Then
        txtTindakanTotal.Text = 0
        txtTindakanTotal.Text = FormatCurrency(txtTindakanTotal.Text, 2)
    End If
    If txtTindakanTotal.Text = "" Then txtTindakanTotal.Text = 0
    If txtAlkesTotal.Text = "" Then txtAlkesTotal.Text = 0
    txtGrandTotal.Text = CCur(txtTindakanTotal.Text) + CCur(txtAlkesTotal)
    txtGrandTotal.Text = FormatCurrency(txtGrandTotal.Text, 2)
Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Store procedure untuk menghapus biaya pelayanan pasien
Private Sub sp_DelBiayaPelayanan(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuanganPasien)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, dgTindakan.Columns("KdPelayananRS").value)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dgTindakan.Columns("TglPelayanan").value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        
        .ActiveConnection = dbConn
        .CommandText = "Delete_BiayaPelayanan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Penghapusan Biaya Pelayanan Pasien", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_BiayaPelayanan")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'Store procedure untuk menghapus diagnosa
Private Sub sp_DelDiagnosa(ByVal adoCommand As ADODB.Command)
Dim rsNew As New ADODB.recordset
    With adoCommand
    
        strSQL = "SELECT * FROM PeriksaDiagnosa WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "' AND KdDiagnosa='" & dgRiwayatDiagnosa.Columns("Kode ICD 10").value & "' AND KdRuangan = '" & mstrKdRuangan & "' AND TglPeriksa='" & Format(dgRiwayatDiagnosa.Columns("TglPeriksa").value, "yyyy/MM/dd HH:mm:ss") & "'"
        Set rsNew = Nothing
        rsNew.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdDiagnosa", adVarChar, adParamInput, 7, dgRiwayatDiagnosa.Columns("Kode ICD 10").value)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dgRiwayatDiagnosa.Columns("TglPeriksa").value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, rsNew("KdSubInstalasi").value)
        .Parameters.Append .CreateParameter("StatusKasus", adChar, adParamInput, 4, rsNew("StatusKasus").value)
        .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 6, txtNoCM.Text)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        
        .ActiveConnection = dbConn
        .CommandText = "Delete_Diagnosa"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Penghapusan Diagnosa Pasien", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_Diagnosa")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'Untuk meload riwayat catatan klinis yang sudah pernah didapat
Public Sub subLoadRiwayatCatatanKlinis()
On Error GoTo errLoad
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
Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Untuk meload riwayat catatan medis yang sudah pernah didapat
Public Sub subLoadRiwayatCatatanMedis()
On Error GoTo errLoad
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
Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Untuk meload riwayat Kecelakaan yang sudah pernah didapat
Public Sub subLoadRiwayatKecelakaan()
On Error GoTo errLoad
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
Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Untuk meload riwayat konsul pasien
Public Sub subLoadRiwayatKonsul()
On Error GoTo errLoad
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
Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Untuk meload riwayat operasi yang sudah pernah didapat
Public Sub subLoadRiwayatOperasi()
On Error GoTo errLoad
    strSQL = "SELECT * " & _
        " FROM V_RiwayatOperasiPasien " & _
        " WHERE (nocm = '" & mstrNoCM & "')"
    Call msubRecFO(rs, strSQL)
    Set dgRiwayatOperasi.DataSource = rs
    With dgRiwayatOperasi
        .Columns(0).Width = 0 'NoCM
        .Columns(1).Width = 0 'NoPendaftaran
        .Columns(2).Width = 1590 '[Tgl. Operasi]
        .Columns(2).Caption = "Tgl. Operasi"
        .Columns(3).Width = 1590 '[Tgl. Selesai]
        .Columns(3).Caption = "Tgl. Selesai"
        .Columns(4).Width = 2500 '[Jenis Operasi]
        .Columns(5).Width = 2500 '[Tindakan Operasi]
        .Columns(6).Width = 1900 '[Ruang Perujuk]
        .Columns(7).Width = 2700 '[Dokter Penanggung Jawab]
        .Columns(8).Width = 0 'KdRuangan
    End With
Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Store procedure untuk menghapus catatan klinis
Private Sub sp_DelBiayaCatatanKlinis(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dgCatatanKlinis.Columns("TglPeriksa").value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        
        .ActiveConnection = dbConn
        .CommandText = "Delete_CatatanKlinis"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penghapusan catatan klinis", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_CatatanKlinis")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'Store procedure untuk menghapus catatan medis
Private Sub sp_DelBiayaCatatanMedis(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dgCatatanMedis.Columns("TglPeriksa").value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        
        .ActiveConnection = dbConn
        .CommandText = "Delete_CatatanMedis"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penghapusan catatan medis", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_CatatanMedis")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'Store procedure untuk menghapus data kecelakaan
Private Sub sp_DelKecelakaan(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        
        .ActiveConnection = dbConn
        .CommandText = "Delete_KasusKecelakaan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penghapusan data kecelakaan", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_KasusKecelakaan")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'Store procedure untuk menghapus data konsul
Private Sub sp_DelKonsul(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, dgKonsul.Columns("KdRuanganAsal"))
        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , Format(dgKonsul.Columns("TglDirujuk").value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        
        .ActiveConnection = dbConn
        .CommandText = "Delete_PasienRujukan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penghapusan data kecelakaan", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_PasienRujukan")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

Private Sub subLoadDiagramOdonto()
    On Error GoTo hell
    Dim blnSudahAda As Boolean
    Dim strTglPeriksa As String
    Dim i As Integer
    
    If dgHasilPemeriksaan.ApproxCount = 0 Then Exit Sub
    
    strSQL = "select NoPendaftaran,TglPeriksa from DetailCatatanOdonto where NoPendaftaran='" & mstrNoPen & "'"
    Call msubRecFO(rs, strSQL)
    
    With frmDiagramOdonto
        .Show
        
        For i = 0 To 14
            .optAksi(i).Visible = False
        Next i
        .txtKeterangan.Visible = False
        .Label3.Visible = False
        .lblBelumErupsi.Visible = False
        .lblErupsiSebagian.Visible = False
        .lblAnomaliBentuk.Visible = False
        .lblCalculus.Visible = False
        .picKaries.Visible = False
        .picNonVital.Visible = False
        .picTLogam.Visible = False
        .picTNonLogam.Visible = False
        .picMLogam.Visible = False
        .picMNonLogam.Visible = False
        .picSisaAkar.Visible = False
        .picGigiHilang.Visible = False
        .picJembatan.Visible = False
        .picGigiTiruanLepas.Visible = False
        .cmdSimpan.Visible = False
        .dtpTglPeriksa.Enabled = False
        
        .Frame2.Height = 800
        .Frame4.Top = .Frame2.Top + .Frame2.Height
        .Height = 8300

        
        
        .txtNoPendaftaran.Text = mstrNoPen
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = txtNamaPasien.Text
        If txtSex.Text = "L" Then
            .txtSex.Text = "Laki-Laki"
        Else
            .txtSex.Text = "Perempuan"
        End If
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHr.Text
        .txtKls.Text = txtKls.Text
        .txtJenisPasien.Text = txtJenisPasien.Text
        .txtTglDaftar.Text = dgHasilPemeriksaan.Columns("TglHasil")
        
        strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            mstrKdJenisPasien = rs("KdKelompokPasien").value
            mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
        End If
        .subLoadDetailCatatanOdonto
    End With
Exit Sub
hell:
    Call msubPesanError
End Sub

