VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTransaksiPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaksi Pelayanan Pasien"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTransaksiPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   14715
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   76
      Top             =   8550
      Width           =   14715
      _ExtentX        =   25956
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   25903
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
   Begin VB.Frame Frame4 
      Caption         =   "Rekapitulasi Tagihan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   13200
      TabIndex        =   63
      Top             =   3360
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox TxtTotalBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2760
         TabIndex        =   64
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtTAsuransi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2760
         TabIndex        =   65
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtTRS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2760
         TabIndex        =   66
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtPembebasan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2760
         TabIndex        =   67
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   73
         Top             =   315
         Width           =   2130
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Tanggungan Penjamin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   72
         Top             =   795
         Width           =   2115
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Tanggungan Rumah Sakit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   71
         Top             =   1275
         Width           =   2445
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Pembebasan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   70
         Top             =   1755
         Width           =   1230
      End
      Begin VB.Label lblTotalTagihan 
         AutoSize        =   -1  'True
         Caption         =   "Rp. 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1680
         TabIndex        =   69
         Top             =   2160
         Width           =   1125
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   240
         TabIndex        =   68
         Top             =   2160
         Width           =   1410
      End
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
      Height          =   6615
      Left            =   0
      TabIndex        =   60
      Top             =   2280
      Width           =   14655
      Begin VB.CommandButton cmdBayar 
         Caption         =   "&Bayar"
         Height          =   375
         Left            =   10440
         TabIndex        =   47
         Top             =   5880
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkBayarKarcis 
         Caption         =   "Bayar Karcis"
         Height          =   255
         Left            =   8880
         TabIndex        =   46
         Top             =   5940
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   12480
         TabIndex        =   48
         Top             =   5880
         Width           =   2055
      End
      Begin VB.TextBox txtGrandTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   45
         Top             =   5880
         Width           =   2415
      End
      Begin TabDlg.SSTab sstTP 
         Height          =   5535
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   9763
         _Version        =   393216
         Tabs            =   10
         Tab             =   4
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
         TabPicture(0)   =   "frmTransaksiPasien.frx":0CCA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "dgTindakan"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdTambahPT"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmdHapusDataPT"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtTindakanTotal"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdUbahPT"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Pemakaian &Obat && Alkes"
         TabPicture(1)   =   "frmTransaksiPasien.frx":0CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdUbahOA"
         Tab(1).Control(1)=   "txtAlkesTotal"
         Tab(1).Control(2)=   "cmdHapusDataPOA"
         Tab(1).Control(3)=   "cmdTambahPOA"
         Tab(1).Control(4)=   "dgObatAlkes"
         Tab(1).Control(5)=   "Label2"
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "Riwayat Catatan Klinis"
         TabPicture(2)   =   "frmTransaksiPasien.frx":0D02
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdKehamilandanKB"
         Tab(2).Control(1)=   "cmdTambahCatatanKlinis"
         Tab(2).Control(2)=   "cmdHapusCatataKlinis"
         Tab(2).Control(3)=   "dgCatatanKlinis"
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "Riwayat Catatan Medis"
         TabPicture(3)   =   "frmTransaksiPasien.frx":0D1E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "cmdTambahCatatanMedis"
         Tab(3).Control(1)=   "cmdHapusCatatanMedis"
         Tab(3).Control(2)=   "cmdCetakCatatanMedis"
         Tab(3).Control(3)=   "dgCatatanMedis"
         Tab(3).ControlCount=   4
         TabCaption(4)   =   "Riwayat &Diagnosa"
         TabPicture(4)   =   "frmTransaksiPasien.frx":0D3A
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "dgRiwayatDiagnosa"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "cmdCetakDiagnosa"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "cmdTambahDiagnosa"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).Control(3)=   "cmdDelDiagnosa"
         Tab(4).Control(3).Enabled=   0   'False
         Tab(4).Control(4)=   "cmdICD9"
         Tab(4).Control(4).Enabled=   0   'False
         Tab(4).Control(5)=   "chkTampil"
         Tab(4).Control(5).Enabled=   0   'False
         Tab(4).ControlCount=   6
         TabCaption(5)   =   "Riwayat Tindakan Medis"
         TabPicture(5)   =   "frmTransaksiPasien.frx":0D56
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "dgRiwayatOperasi"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).Control(1)=   "cmdTambahTM"
         Tab(5).Control(1).Enabled=   0   'False
         Tab(5).Control(2)=   "cmdHapusTM"
         Tab(5).Control(2).Enabled=   0   'False
         Tab(5).ControlCount=   3
         TabCaption(6)   =   "Riwayat Konsul"
         TabPicture(6)   =   "frmTransaksiPasien.frx":0D72
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "dgKonsul"
         Tab(6).Control(0).Enabled=   0   'False
         Tab(6).Control(1)=   "cmdHapusKonsul"
         Tab(6).Control(1).Enabled=   0   'False
         Tab(6).Control(2)=   "cmdTambahKonsul"
         Tab(6).Control(2).Enabled=   0   'False
         Tab(6).ControlCount=   3
         TabCaption(7)   =   "Riwayat Kecelakaan"
         TabPicture(7)   =   "frmTransaksiPasien.frx":0D8E
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "dgKecelakaan"
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "Riwayat Peme&riksaan"
         TabPicture(8)   =   "frmTransaksiPasien.frx":0DAA
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "cmdCetakRP"
         Tab(8).Control(1)=   "chkRP"
         Tab(8).Control(2)=   "dgRiwayatPemeriksaan"
         Tab(8).ControlCount=   3
         TabCaption(9)   =   "Riwayat Hasil Pemeriksaan"
         TabPicture(9)   =   "frmTransaksiPasien.frx":0DC6
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "cmdCetakResume"
         Tab(9).Control(1)=   "cmdCetakHasilPemeriksaan"
         Tab(9).Control(2)=   "dgHasilPemeriksaan"
         Tab(9).ControlCount=   3
         Begin VB.CommandButton cmdCetakResume 
            Caption         =   "Cetak Resu&me"
            Height          =   375
            Left            =   -64080
            TabIndex        =   80
            Top             =   4965
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusTM 
            Caption         =   "&Hapus Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -64080
            TabIndex        =   79
            Top             =   4845
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahTM 
            Caption         =   "&Tambah Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -62400
            TabIndex        =   78
            Top             =   4845
            Width           =   1575
         End
         Begin VB.CheckBox chkTampil 
            Caption         =   "Semua Diagnosa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   4845
            Width           =   3015
         End
         Begin VB.CommandButton cmdICD9 
            Caption         =   "&Edit Diagnosa Tindakan [ICD 9]"
            Enabled         =   0   'False
            Height          =   375
            Left            =   7680
            TabIndex        =   32
            Top             =   4845
            Width           =   3135
         End
         Begin VB.CommandButton cmdKehamilandanKB 
            Caption         =   "&Data Kehamilan dan KB"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -66480
            TabIndex        =   22
            Top             =   4845
            Width           =   2295
         End
         Begin VB.CommandButton cmdUbahOA 
            Caption         =   "&Ubah Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -65760
            TabIndex        =   18
            Top             =   4845
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetakHasilPemeriksaan 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -62400
            TabIndex        =   44
            Top             =   4965
            Width           =   1575
         End
         Begin VB.CommandButton cmdUbahPT 
            Caption         =   "&Ubah Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -65760
            TabIndex        =   13
            Top             =   4965
            Width           =   1575
         End
         Begin VB.TextBox txtTindakanTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   -72000
            TabIndex        =   12
            Top             =   5085
            Width           =   2415
         End
         Begin VB.TextBox txtAlkesTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   -71640
            TabIndex        =   17
            Top             =   4965
            Width           =   2415
         End
         Begin VB.CommandButton cmdHapusDataPT 
            Caption         =   "&Hapus Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -64080
            TabIndex        =   14
            Top             =   4965
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahPT 
            Caption         =   "&Tambah Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -62400
            TabIndex        =   15
            Top             =   4965
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusDataPOA 
            Caption         =   "&Hapus Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -64080
            TabIndex        =   19
            Top             =   4845
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahPOA 
            Caption         =   "&Tambah Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -62400
            TabIndex        =   20
            Top             =   4845
            Width           =   1575
         End
         Begin VB.CommandButton cmdDelDiagnosa 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   10920
            TabIndex        =   33
            Top             =   4845
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahDiagnosa 
            Caption         =   "&Tambah Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   12600
            TabIndex        =   34
            Top             =   4845
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetakDiagnosa 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   6000
            TabIndex        =   31
            Top             =   4845
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetakRP 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -62400
            TabIndex        =   42
            Top             =   4965
            Visible         =   0   'False
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
            TabIndex        =   41
            Top             =   4845
            Width           =   1815
         End
         Begin VB.CommandButton cmdTambahCatatanKlinis 
            Caption         =   "&Tambah Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -62400
            TabIndex        =   24
            Top             =   4845
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusCatataKlinis 
            Caption         =   "&Hapus Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -64080
            TabIndex        =   23
            Top             =   4845
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahCatatanMedis 
            Caption         =   "&Tambah Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -62400
            TabIndex        =   28
            Top             =   4845
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusCatatanMedis 
            Caption         =   "&Hapus Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -64080
            TabIndex        =   27
            Top             =   4845
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahKonsul 
            Caption         =   "&Tambah Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -62400
            TabIndex        =   38
            Top             =   4845
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusKonsul 
            Caption         =   "&Hapus Data"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -64080
            TabIndex        =   37
            Top             =   4845
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetakCatatanMedis 
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   -65760
            TabIndex        =   26
            Top             =   4845
            Width           =   1575
         End
         Begin MSDataGridLib.DataGrid dgTindakan 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   11
            Top             =   645
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
            TabIndex        =   16
            Top             =   645
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
            Left            =   240
            TabIndex        =   29
            Top             =   645
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
            TabIndex        =   40
            Top             =   645
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
            TabIndex        =   21
            Top             =   645
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
            TabIndex        =   25
            Top             =   690
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
            TabIndex        =   35
            Top             =   645
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
            TabIndex        =   36
            Top             =   645
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
            TabIndex        =   39
            Top             =   645
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
            TabIndex        =   43
            Top             =   645
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Biaya Pelayanan Tindakan"
            Height          =   210
            Left            =   -74760
            TabIndex        =   75
            Top             =   5145
            Width           =   2550
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total Biaya Pemakaian Obat && Alkes"
            Height          =   210
            Left            =   -74760
            TabIndex        =   74
            Top             =   5025
            Width           =   2925
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya Pelayanan"
         Height          =   210
         Left            =   360
         TabIndex        =   61
         Top             =   5955
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
      TabIndex        =   49
      Top             =   1080
      Width           =   14655
      Begin VB.TextBox txtKls 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   9480
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
         Left            =   6960
         TabIndex        =   50
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
            TabIndex        =   53
            Top             =   270
            Width           =   165
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Left            =   1350
            TabIndex        =   52
            Top             =   277
            Width           =   240
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Left            =   550
            TabIndex        =   51
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
         Width           =   1215
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1395
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2760
         TabIndex        =   2
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5640
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtJenisPasien 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   11040
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtTglDaftar 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   12960
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Pelayanan"
         Height          =   210
         Left            =   9480
         TabIndex        =   62
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No Pendaftaran"
         Height          =   210
         Left            =   75
         TabIndex        =   59
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1440
         TabIndex        =   58
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   2760
         TabIndex        =   57
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   5640
         TabIndex        =   56
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pasien"
         Height          =   210
         Left            =   11040
         TabIndex        =   55
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pendaftaran"
         Height          =   210
         Left            =   12960
         TabIndex        =   54
         Top             =   360
         Width           =   1365
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   77
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
      Picture         =   "frmTransaksiPasien.frx":0DE2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image4 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTransaksiPasien.frx":1B6A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13455
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmTransaksiPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bolStatusFIFO As Boolean
Dim vbMsgboxRslt As String

Private Sub chkRP_Click()
    If chkRP.value = 0 Then
        subLoadRiwayatPemeriksaan False
    Else
        subLoadRiwayatPemeriksaan True
    End If
End Sub

'Menampilkan Riwayat Diagnosa Sebelumnya
Private Sub chkTampil_Click()
    If chkTampil.value = 1 Then
        Call subLoadRiwayatDiagnosa(True)
    End If

    If chkTampil.value = 0 Then
        Call subLoadRiwayatDiagnosa(False)
    Else
        Call subLoadRiwayatDiagnosa(True)
    End If
End Sub

Private Sub cmdCetakCatatanMedis_Click()
    On Error GoTo hell
    If dgCatatanMedis.ApproxCount = 0 Then Exit Sub
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    mstrNoCM = txtNoCM.Text
    frmCetakCatatanMedis.Show
    Exit Sub
hell:
End Sub

Private Sub cmdCetakDiagnosa_Click()
    On Error GoTo hell
    If dgRiwayatDiagnosa.ApproxCount = 0 Then Exit Sub
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    mblnStatusCetakRD = True
    frm_cetak_info_diag_viewer.Show
    Exit Sub
hell:
End Sub

Private Sub cmdCetakHasilPemeriksaan_Click()
On Error Resume Next
'On Error GoTo errLoad
    Dim pesan As VbMsgBoxResult

    If dgHasilPemeriksaan.ApproxCount = 0 Then Exit Sub
    
'    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
'    vLaporan = ""
'    If pesan = vbYes Then vLaporan = "Print"
    

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
            strSQL = "select NoVerifikasi from HasilPemeriksaan where NoLab_Rad='" & dgHasilPemeriksaan.Columns("NoLab_Rad") & "'" 'and NoPendaftaran='" & txtnopendaftaran.Text & "'"
            Call msubRecFO(rs, strSQL)
'            If IsNull(rs(0)) Then
            If rs.EOF = True Or IsNull(rs(0)) Then
                MsgBox "Data hasil belum di verifikasi..", vbCritical, "Validasi": GoTo lanjut
            End If
            
            Set frmcetakhasillab = Nothing
            strQuery = "sELECT * from V_CetakHasilLaboratoriumPK WHERE NoLaboratorium = '" & mstrNoLabRad & "'"
            Call msubRecFO(dbRst, strQuery)
            If dbRst.EOF = False Then
                vLaporan = ""
                If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
                frmcetakhasillab.Show
                
            Else
               MsgBox "Tidak ada data yang di tampilkan.", vbInformation, "Informasi"
            End If

        Case "16" 'lab pa
            strSQL = "select NoVerifikasi from HasilPemeriksaan where NoLab_Rad='" & dgHasilPemeriksaan.Columns("NoLab_Rad") & "'" 'and NoPendaftaran='" & txtnopendaftaran.Text & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = True Or IsNull(rs(0)) Then
                MsgBox "Data hasil belum di verifikasi..", vbCritical, "Validasi": GoTo lanjut
            End If
           
           Set frmCetakHasilLabPA = Nothing
            strQuery = "SELECT * " & _
                     " from V_CetakHasilPeriksaLaboratoryPA " & _
                     " WHERE NoLaboratorium = '" & mstrNoLabRad & "'"
            Call msubRecFO(dbRst, strQuery)
            If dbRst.EOF = False Then
                vLaporan = ""
                If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
                frmCetakHasilLabPA.Show
            Else
               MsgBox "Tidak ada data yang di tampilkan.", vbInformation, "Informasi"
            End If
            
'            frmCetakHasilLabPA.Show

        Case "10" 'radiologi
             strSQL = "select NoVerifikasi from HasilPemeriksaan where NoLab_Rad='" & dgHasilPemeriksaan.Columns("NoLab_Rad") & "'" 'and NoPendaftaran='" & txtnopendaftaran.Text & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = True Or IsNull(rs(0)) Then
                MsgBox "Data hasil belum di verifikasi..", vbCritical, "Validasi": GoTo lanjut
            End If
       
            Set frmCetakHasilRadiologi = Nothing
            strSQL = "SELECT distinct NoRadiology,NoPendaftaran,NoCM,NamaPasien,TglHasil,Umur,AlamatLengkap,RuanganPerujuk,AsalPasien, " & _
                        " JenisKelamin,DokterPerujuk,NamaDetailPeriksa,NamaPelayanan,MemoHasilPeriksa,Catatan  " & _
                        " from V_CetakHasilPemeriksaanRadiology WHERE NoRadiology = '" & mstrNoLabRad & "'"
            Call msubRecFO(dbRst, strSQL)
            If dbRst.EOF = False Then
                vLaporan = ""
                If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
                frmCetakHasilRadiologi.Show
            Else
               MsgBox "Tidak ada data yang di tampilkan.", vbInformation, "Informasi"
            End If
            
        Case Else
            Call subLoadDiagramOdonto
    End Select
lanjut:
    cmdCetakHasilPemeriksaan.Enabled = True

'Exit Sub
'errLoad:
'    Call msubPesanError
'    cmdCetakHasilPemeriksaan.Enabled = True
End Sub

Private Sub cmdCetakResume_Click()
    On Error Resume Next
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakDataRiwayatPemeriksaanPasien.Show
End Sub

Private Sub cmdDelDiagnosa_Click()
    Dim vbMsgboxRslt As VbMsgBoxResult
    Dim dbcmd As New ADODB.Command
    On Error GoTo errHapusData

    If dgRiwayatDiagnosa.ApproxCount = 0 Then Exit Sub

'    If dgRiwayatDiagnosa.Columns("Ruang Periksa").value <> mstrNamaRuangan Then
'        MsgBox "Diagnosa yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
'        Exit Sub
'    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus diagnosa '" _
    & dgRiwayatDiagnosa.Columns("Diagnosa ICD 10").value & "'" & vbNewLine _
    & "Dengan tanggal pelayanan '" & dgRiwayatDiagnosa.Columns("TglPeriksa").value _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    'diagnosa utama hanya bisa di-replace
'    If dgRiwayatDiagnosa.Columns(14).value <> "05" Then
        sp_DelDiagnosa dbcmd
        subLoadRiwayatDiagnosa (False)
        MsgBox "Data diagnosa berhasil dihapus ", vbInformation, "Informasi"
'    Else
'        MsgBox "Diagnosa Utama hanya bisa diganti (replace)", vbInformation, "Informasi"
'        Exit Sub
'    End If
    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusCatataKlinis_Click()
    On Error GoTo errHapusData

    Dim vbMsgboxRslt As VbMsgBoxResult
    Dim dbcmd As New ADODB.Command

    If dgCatatanKlinis.ApproxCount = 0 Then Exit Sub

    If dgCatatanKlinis.Columns("Ruang Pemeriksaan").value <> mstrNamaRuangan Then
        MsgBox "Catatan klinis yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
        Exit Sub
    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus catatan klinis pasien '" _
    & dgCatatanKlinis.Columns("NoPendaftaran").value & "'" & vbNewLine _
    & "Dengan tanggal periksa '" & dgCatatanKlinis.Columns("TglPeriksa").value _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    sp_DelBiayaCatatanKlinis dbcmd
    Call subLoadRiwayatCatatanKlinis
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"

    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusCatatanMedis_Click()
    On Error GoTo errHapusData

    Dim vbMsgboxRslt As VbMsgBoxResult
    Dim dbcmd As New ADODB.Command

    If dgCatatanMedis.ApproxCount = 0 Then Exit Sub

    If dgCatatanMedis.Columns("RuangPemeriksaan").value <> mstrNamaRuangan Then
        MsgBox "Catatan medis yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
        Exit Sub
    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus catatan medis pasien '" _
    & dgCatatanMedis.Columns("NoPendaftaran").value & "'" & vbNewLine _
    & "Dengan tanggal periksa '" & dgCatatanMedis.Columns("TglPeriksa").value _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    sp_DelBiayaCatatanMedis dbcmd
    Call subLoadRiwayatCatatanMedis
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"

    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusDataPOA_Click()
    Dim adoCommand As New ADODB.Command
    Dim i As Integer
    On Error GoTo errHapusData

    If dgObatAlkes.ApproxCount = 0 Then Exit Sub
    If dgObatAlkes.Columns("Status Bayar").value = "Sudah DiBayar" Then
        MsgBox "Pemakaian Obat dan Alkes yang sudah dibayar tidak dapat dihapus", vbCritical, "Validasi"
        Exit Sub
    ElseIf dgObatAlkes.Columns("Ruang Pelayanan").value <> mstrNamaRuangan Then
        MsgBox "Pemakaian Obat dan Alkes yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
        Exit Sub
    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus pemakaian obat dan alkes '" _
    & dgObatAlkes.Columns("NamaBarang").value & "'" & vbNewLine _
    & "Dengan tanggal pelayanan '" & dgObatAlkes.Columns("TglPelayanan").value _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    If bolStatusFIFO = True Then
        strSQL = "SELECT * FROM V_BiayaPemakaianObatAlkes WHERE NoPendaftaran='" & mstrNoPen & "' and KdBarang='" & dgObatAlkes.Columns("KdBarang") & "' " _
        & "and KdRuangan='" & mstrKdRuangan & "' and KdAsal='" & dgObatAlkes.Columns("KdAsal") & "' and SatuanJml='" & dgObatAlkes.Columns("Sat") & "' " _
        & "and tglPelayanan='" & Format(dgObatAlkes.Columns("TglPelayanan").value, "yyyy/MM/dd HH:mm:ss") & "' Order by NoTerima Desc"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.EOF = False Then
            rs.MoveFirst
            For i = 1 To rs.RecordCount
                Set dbcmd = New ADODB.Command
                With dbcmd
                    .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
                    .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, rs("KdBarang"))
                    .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, rs("KdAsal"))
                    .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
                    .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, rs("SatuanJml"))
                    .Parameters.Append .CreateParameter("JmlBrg", adInteger, adParamInput, , rs("JmlBarang"))
                    .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
                    .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(rs("TglPelayanan"), "yyyy/MM/dd HH:mm:ss"))
                    .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
                    .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, rs("NoTerima"))

                    .ActiveConnection = dbConn
                    .CommandText = "dbo.Delete_PemakaianObatAlkes"
                    .CommandType = adCmdStoredProc
                    .Execute

                    If Not (.Parameters("RETURN_VALUE").value = 0) Then
                        MsgBox "Ada Kesalahan dalam penghapusan data pemakaian obat dan alkes", vbCritical, "Validasi"
                        Exit Sub
                    End If
                    Call deleteADOCommandParameters(dbcmd)
                    Set dbcmd = Nothing
                End With

                rs.MoveNext
            Next i
            Call Add_HistoryLoginActivity("Delete_PemakaianObatAlkes")
        End If
    Else
        Set dbcmd = New ADODB.Command
        With dbcmd
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
            .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, dgObatAlkes.Columns("KdBarang").value)
            .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, dgObatAlkes.Columns("KdAsal").value)
            .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
            .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, dgObatAlkes.Columns("Sat").value)
            .Parameters.Append .CreateParameter("JmlBrg", adInteger, adParamInput, , dgObatAlkes.Columns("Jml").value)
            .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
            .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dgObatAlkes.Columns("TglPelayanan").value, "yyyy/MM/dd HH:mm:ss"))
            .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
            .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, dgObatAlkes.Columns("NoTerima").value)

            .ActiveConnection = dbConn
            .CommandText = "dbo.Delete_PemakaianObatAlkes"
            .CommandType = adCmdStoredProc
            .Execute

            If Not (.Parameters("RETURN_VALUE").value = 0) Then
                MsgBox "Ada Kesalahan dalam penghapusan data pemakaian obat dan alkes", vbCritical, "Validasi"
                Exit Sub
            Else
                Call Add_HistoryLoginActivity("Delete_PemakaianObatAlkes")
            End If
            Call deleteADOCommandParameters(dbcmd)
            Set dbcmd = Nothing
        End With
    End If
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
        MsgBox "Pelayanan yang sudah dibayar tidak dapat dihapus ", vbCritical, "Validasi"
        Exit Sub
    ElseIf dgTindakan.Columns("Ruang Pelayanan").value <> mstrNamaRuangan Then
        MsgBox "Pelayanan yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
        Exit Sub
    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus pelayanan '" _
    & dgTindakan.Columns("NamaPelayanan").value & "'" & vbNewLine _
    & "Dengan tanggal pelayanan '" & dgTindakan.Columns("TglPelayanan").value _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    sp_DelBiayaPelayanan dbcmd
    Call subLoadPelayananDidapat
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"

    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusKonsul_Click()
    On Error GoTo errHapusData

    Dim vbMsgboxRslt As VbMsgBoxResult
    Dim dbcmd As New ADODB.Command

    If dgKonsul.ApproxCount = 0 Then Exit Sub

    If dgKonsul.Columns("Ruangan Perujuk").value <> mstrNamaRuangan Then
        MsgBox "Catatan medis yang didapat di ruangan lain tidak dapat dihapus di ruangan ini", vbCritical, "Validasi"
        Exit Sub
    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus catatan medis pasien '" _
    & dgKonsul.Columns("NoPendaftaran").value & "'" & vbNewLine _
    & "Dengan tanggal periksa '" & dgKonsul.Columns("TglDirujuk").value _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    sp_DelKonsul dbcmd
    Call subLoadRiwayatKonsul
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"

    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusTM_Click()
    On Error GoTo errHapusData

    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus tindakan medis pasien '" _
    & dgRiwayatOperasi.Columns("NoHasilPeriksa").value & "'" & vbNewLine _
    & "Dengan tanggal periksa '" & dgRiwayatOperasi.Columns("TglMulaiPeriksa").value _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    dbConn.Execute "DELETE FROM DetailHasilTindakanMedisPasien WHERE NoHasilPeriksa = '" & dgRiwayatOperasi.Columns("NoHasilPeriksa") & "' and NoPendaftaran = '" & dgRiwayatOperasi.Columns("NoPendaftaran") & "'"
    dbConn.Execute "DELETE FROM HasilTindakanMedis WHERE NoHasilPeriksa = '" & dgRiwayatOperasi.Columns("NoHasilPeriksa") & "' and NoPendaftaran = '" & dgRiwayatOperasi.Columns("NoPendaftaran") & "'"

    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"
    Call subLoadRiwayatOperasi
    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdICD9_Click()
    On Error GoTo hell
    Dim X As Integer

    If dgRiwayatDiagnosa.ApproxCount = 0 Then Exit Sub
    If dgRiwayatDiagnosa.Columns(0) <> txtNoPendaftaran.Text Then
        MsgBox "No Pendaftaran tidak sama, mohon isi diagnosanya [ICD 10] dahulu", vbExclamation, "Validasi"
        cmdTambahDiagnosa.SetFocus
        Exit Sub
    End If
    Me.Enabled = False
    mstrKdDiagnosa = ""
    mstrKdDiagnosa = dgRiwayatDiagnosa.Columns(4)
    mstrKdJenisDiagnosaTindakan = ""
    mstrKdJenisDiagnosaTindakan = dgRiwayatDiagnosa.Columns(15)
    bolEditDiagnosa = True
    With frmPeriksaDiagnosa
        .Show
        .txtNoPendaftaran = txtNoPendaftaran.Text
        .txtNamaFormPengirim.Text = Me.Name
        .txtNoCM = txtNoCM.Text
        .txtNamaPasien = txtNamaPasien.Text
        If Left(.txtSex, 1) = "P" Then
            .txtSex.Text = "Perempuan"
        Else
            .txtSex.Text = "Laki-laki"
        End If
        .txtThn = txtThn.Text
        .txtBln = txtBln.Text
        .txthari = txtHr.Text

        .txtDokter.Text = dgRiwayatDiagnosa.Columns(10)
        mstrKdDokter = dgRiwayatDiagnosa.Columns(16)
        intJmlDokter = 1
        .fraDokter.Visible = False

        .dtpTglPeriksa.value = dgRiwayatDiagnosa.Columns(2)
        .dcJenisDiagnosa.BoundText = dgRiwayatDiagnosa.Columns(14)
        .dcJenisDiagnosaTindakan.BoundText = dgRiwayatDiagnosa.Columns(15)

        .dcJenisDiagnosa.Enabled = False
        .lvwDiagnosa.Enabled = False
        .txtNamaDiagnosa.Enabled = False
        .txtDokter.Enabled = False
        .dtpTglPeriksa.Enabled = False
        .chkICD9.value = Checked

        .Show
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdKehamilandanKB_Click()
    frmDataKehamilandanKB.Show
End Sub

Private Sub cmdTambahCatatanKlinis_Click()
    If txtNoCM.Text = 0 Then Exit Sub
    frmCatatanKlinisPasien.Show
    frmTransaksiPasien.Enabled = False
End Sub

Private Sub cmdTambahCatatanMedis_Click()
    On Error GoTo hell
    If txtNoCM.Text = "" Then Exit Sub
    frmTransaksiPasien.Enabled = False

    With frmCatatanMedikPasien
        strSQL = "SELECT dbo.RegistrasiRJ.IdDokter, dbo.DataPegawai.NamaLengkap " & _
        " FROM dbo.RegistrasiRJ INNER JOIN dbo.DataPegawai ON dbo.RegistrasiRJ.IdDokter = dbo.DataPegawai.IdPegawai " & _
        " WHERE (dbo.RegistrasiRJ.NoPendaftaran = '" & txtNoPendaftaran.Text & "')"
        Call msubRecFO(rs, strSQL)

        If Not rs.EOF Then
            .txtDokter.Text = rs(1).value
            mstrKdDokter = rs(0).value
            intJmlDokter = rs.RecordCount
            .fraDokter.Visible = False
        End If
        .Show
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTambahDiagnosa_Click()
    On Error GoTo errLoad
    Me.Enabled = False
    With frmPeriksaDiagnosa
        .Show
        .txtNoPendaftaran = txtNoPendaftaran.Text
        .txtNamaFormPengirim.Text = Me.Name
        .txtNoCM = txtNoCM.Text
        .txtNamaPasien = txtNamaPasien.Text
        .txtSex.Text = txtSex.Text
'        If Left(.txtSex, 1) = "P" Then
'            .txtSex.Text = "Perempuan"
'        Else
'            .txtSex.Text = "Laki-laki"
'        End If
        .txtThn = txtThn.Text
        .txtBln = txtBln.Text
        .txthari = txtHr.Text

        strSQL = "SELECT dbo.RegistrasiRJ.IdDokter, dbo.DataPegawai.NamaLengkap " & _
        " FROM dbo.RegistrasiRJ INNER JOIN dbo.DataPegawai ON dbo.RegistrasiRJ.IdDokter = dbo.DataPegawai.IdPegawai " & _
        " WHERE (dbo.RegistrasiRJ.NoPendaftaran = '" & txtNoPendaftaran.Text & "')"
        Call msubRecFO(rs, strSQL)

        If Not rs.EOF Then
            .txtDokter.Text = rs(1).value
            mstrKdDokter = rs(0).value
            intJmlDokter = rs.RecordCount
            .fraDokter.Visible = False
        End If

    End With
    Exit Sub
errLoad:
    Me.Enabled = True
    Call msubPesanError
    frmPeriksaDiagnosa.Show
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
        .txthari.Text = txtHr.Text
        .dtpTglDirujuk.value = Now
        strSQL = "SELECT KdSubInstalasi, IdDokter, Dokter FROM V_DokterPasien WHERE NoPendaftaran = '" & txtNoPendaftaran.Text & "'"
        Call msubRecFO(dbRst, strSQL)
        If dbRst.EOF = False Then
            mstrKdSubInstalasi = dbRst("KdSubInstalasi").value
            frmPasienRujukan.txtDokter.Text = dbRst("Dokter").value
            mstrKdDokter = dbRst("IdDokter").value
            intJmlDokter = dbRst.RecordCount
            frmPasienRujukan.fraDokter.Visible = False
        Else
            mstrKdDokter = ""
            intJmlDokter = 0
        End If
    End With

    Me.Enabled = False
    frmPasienRujukan.Show

    Exit Sub
errLoad:
    Call msubPesanError
    frmPasienRujukan.Show
End Sub

Private Sub cmdTambahPOA_Click()
    On Error GoTo errLoad

    frmPemakaianObatAlkes.Show

    Exit Sub
errLoad:
    Call msubPesanError
    frmPemakaianObatAlkes.Show
End Sub

Private Sub cmdTambahPT_Click()
    On Error GoTo errLoad
    Dim tempKodeDokter As String

    strSQL = "SELECT dbo.RegistrasiRJ.IdDokter, dbo.DataPegawai.NamaLengkap " & _
    " FROM dbo.RegistrasiRJ INNER JOIN dbo.DataPegawai ON dbo.RegistrasiRJ.IdDokter = dbo.DataPegawai.IdPegawai " & _
    " WHERE (dbo.RegistrasiRJ.NoPendaftaran = '" & txtNoPendaftaran.Text & "')"
    Call msubRecFO(rs, strSQL)

    If Not rs.EOF Then
        tempKodeDokter = rs(0).value
        intJmlDokter = rs.RecordCount

        frmTindakan.txtDokter.Text = rs(1).value
        mstrKdDokter = tempKodeDokter
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

Private Sub cmdUbahOA_Click()
    On Error GoTo errLoad

    If dgObatAlkes.ApproxCount = 0 Then Exit Sub
    If dgObatAlkes.Columns("Status Bayar").value = "Sudah DiBayar" Then
        MsgBox "Pelayanan yang sudah dibayar tidak dapat diubah", vbCritical, "Validasi"
        Exit Sub
    ElseIf dgObatAlkes.Columns("Ruang Pelayanan").value <> mstrNamaRuangan Then
        MsgBox "Pelayanan yang didapat di ruangan lain tidak dapat diubah di ruangan ini", vbCritical
        Exit Sub
    End If
    With frmUpdateBiayaPelayanan
        .txtNoPendaftaran = txtNoPendaftaran.Text
        strKodePelayananRS = dgObatAlkes.Columns(12).value
        Call .txtNoPendaftaran_KeyPress(13)
    End With

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdUbahPT_Click()
    On Error GoTo errLoad

    If dgTindakan.ApproxCount = 0 Then Exit Sub
    If dgTindakan.Columns("Status Bayar").value = "Sudah DiBayar" Then
        MsgBox "Pelayanan yang sudah dibayar tidak dapat diubah", vbCritical, "Validasi"
        Exit Sub
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

Private Sub dgCatatanKlinis_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgCatatanKlinis
    WheelHook.WheelHook dgCatatanKlinis
End Sub

Private Sub dgCatatanKlinis_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then cmdTambahCatatanKlinis.SetFocus
End Sub

Private Sub dgCatatanMedis_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgCatatanMedis
    WheelHook.WheelHook dgCatatanMedis
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
    On Error Resume Next
    If KeyAscii = 13 Then cmdCetakHasilPemeriksaan.SetFocus
End Sub

Private Sub dgHasilPemeriksaan2_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then cmdCetakHasilPemeriksaan.SetFocus
End Sub

Private Sub dgKecelakaan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKecelakaan
    WheelHook.WheelHook dgKecelakaan
End Sub

Private Sub dgKonsul_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKonsul
    WheelHook.WheelHook dgKonsul
End Sub

Private Sub dgKonsul_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then cmdTambahKonsul.SetFocus
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
    On Error Resume Next
    If KeyAscii = 13 Then cmdTambahDiagnosa.SetFocus
End Sub

Private Sub dgRiwayatOperasi_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgRiwayatOperasi
    WheelHook.WheelHook dgRiwayatOperasi
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
    End Select

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
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

    sstTP.Tab = 0
'    If mblnAdmin = True Then
'        cmdHapusDataPT.Enabled = True
'        cmdHapusDataPOA.Enabled = True
'        cmdHapusCatataKlinis.Enabled = True
'        cmdHapusCatatanMedis.Enabled = True
'        cmdHapusKonsul.Enabled = True
'        cmdUbahPT.Enabled = True
'    Else
'        cmdHapusDataPT.Enabled = False
'        cmdHapusDataPOA.Enabled = False
'        cmdHapusCatataKlinis.Enabled = False
'        cmdHapusCatatanMedis.Enabled = False
'        cmdHapusKonsul.Enabled = False
'        cmdUbahPT.Enabled = False
'    End If

    strSQL = "SELECT * FROM StatusObject WHERE KdAplikasi='006' AND NamaForm='frmTransaksiPasien' AND NamaObject='chkBayarKarcis' AND StatusEnable='T'"
    msubRecFO rs, strSQL
    If rs.RecordCount = 1 Then
        chkBayarKarcis.Enabled = False
    Else
        chkBayarKarcis.Enabled = True
    End If
    strSQL = "SELECT * FROM StatusObject WHERE KdAplikasi='006' AND NamaForm='frmTransaksiPasien' AND NamaObject='cmdBayar' AND StatusEnable='T'"
    msubRecFO rs, strSQL
    If rs.RecordCount = 1 Then
        cmdBayar.Enabled = False
    Else
        cmdBayar.Enabled = True
    End If
    Call PlayFlashMovie(Me)
    Exit Sub
errLoad:
    msubPesanError
End Sub

'Untuk meload riwayat diagnosa yang sudah pernah didapat
Public Sub subLoadRiwayatDiagnosa(blnAll As Boolean)
    If blnAll = False Then
        strSQL = "Select * from V_DaftarDiagnosaPasien where nocm = '" & Right(mstrNoCM, 6) & "' AND NoPendaftaran = '" & mstrNoPen & "'"
    Else
        strSQL = "Select * from V_DaftarDiagnosaPasien where nocm = '" & Right(mstrNoCM, 6) & "'"
    End If
    Call msubRecFO(rs, strSQL)
    Set dgRiwayatDiagnosa.DataSource = rs
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
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat pemeriksaan yang sudah pernah didapat
Public Sub subLoadRiwayatPemeriksaan(blnAll As Boolean)
    On Error GoTo hell
    If blnAll = False Then
        strSQL = "Select * from V_RiwayatPemeriksaanPasien where NoPendaftaran = '" & mstrNoPen & "'" 'AND KdRuangan='" & mstrKdRuanganPasien & "'" 'mstrKdRuangan
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
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat hasil pemeriksaan yang sudah pernah didapat
Public Sub subLoadRiwayatHasilPemeriksaan()
    On Error GoTo hell
    strSQL = "Select NoLab_Rad, [Ruang Pemeriksa], [Dokter Pemeriksa], TglPendaftaran, TglHasil, [Asal Rujukan], [Ruangan Perujuk], [Dokter Perujuk], KdInstalasi from V_RiwayatHasilPemeriksaan where nocm = '" & mstrNoCM & "'"
    msubRecFO rs, strSQL
    Set dgHasilPemeriksaan.DataSource = rs
    dgHasilPemeriksaan.Columns("KdInstalasi").Width = 0
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Untuk meload pelayanan yang sudah pernah didapat
Public Sub subLoadPelayananDidapat()
    On Error GoTo hell

    strSQL = "SELECT TglPelayanan,JenisPelayanan,NamaPelayanan,NamaRuangan AS [Ruang Pelayanan]," _
    & "Kelas,JenisTarif,CITO,JmlPelayanan as Jml,Total as Tarif,BiayaTotal," _
    & "DokterPemeriksa,[Status Bayar],KdPelayananRS, KdRuangan,Operator FROM V_BiayaPelayananTindakan WHERE " _
    & "NoPendaftaran='" & mstrNoPen & "' ORDER BY TglPelayanan"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgTindakan.DataSource = rs
    With dgTindakan
        .Columns(0).Width = 1590
        .Columns(1).Width = 2600
        .Columns(2).Width = 2400
        .Columns(3).Width = 1800
        .Columns(4).Width = 1200
        .Columns(5).Width = 1000
        .Columns(6).Width = 500
        .Columns(7).Width = 400
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 1000
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 1100
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 2400
        .Columns(11).Width = 1300
        .Columns(12).Width = 0 'KdPelayananRS
        .Columns(13).Width = 0 'KdRuangan
        .Columns(14).Width = 2000

        .Columns(8).NumberFormat = "#,###"
        .Columns(9).NumberFormat = "#,###"
    End With

    strSQL = "SELECT sum(BiayaTotal) as TotalBayar FROM V_BiayaPelayananTindakan " _
    & "WHERE NoPendaftaran='" _
    & mstrNoPen & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount <> 0 Then
        txtTindakanTotal.Text = FormatCurrency(rs.Fields(0).value, 2)
        If IsNull(rs.Fields(0)) = True Then txtTindakanTotal.Text = FormatCurrency(0, 2)
    Else
        txtTindakanTotal.Text = FormatCurrency(0, 2)
    End If
    If txtAlkesTotal.Text = "" Then
        txtAlkesTotal.Text = 0
        txtAlkesTotal.Text = FormatCurrency(txtAlkesTotal.Text, 2)
    End If
    txtGrandTotal.Text = CCur(txtTindakanTotal.Text) + CCur(txtAlkesTotal)
    txtGrandTotal.Text = FormatCurrency(txtGrandTotal.Text, 2)
    Exit Sub
hell:
    Call msubPesanError
End Sub

Public Sub subpemakaianobatalkes()
    On Error GoTo hell

    strSQL = "SELECT TglPelayanan,[Detail Jenis Brg],NamaBarang," _
    & "NamaRuangan AS [Ruang Pelayanan],Kelas,JenisTarif,SatuanJml as Sat," _
    & "JmlBarang as Jml,HargaSatuan as Tarif,BiayaTotal,DokterPemeriksa," _
    & "[Status Bayar],KdBarang,KdAsal,Operator,KdRuangan, NoTerima " _
    & "FROM V_BiayaPemakaianObatAlkes WHERE NoPendaftaran='" _
    & mstrNoPen & "' ORDER BY TglPelayanan"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgObatAlkes.DataSource = rs
    With dgObatAlkes
        .Columns(0).Width = 1590
        .Columns(1).Width = 2000
        .Columns(2).Width = 2000
        .Columns(3).Width = 1800
        .Columns(4).Width = 1200
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
        .Columns(15).Width = 0
        .Columns(16).Width = 0
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
    txtGrandTotal.Text = CCur(txtTindakanTotal.Text) + CCur(txtAlkesTotal)
    txtGrandTotal.Text = FormatCurrency(txtGrandTotal.Text, 2)
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Store procedure untuk menghapus biaya pelayanan pasien
Private Sub sp_DelBiayaPelayanan(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, dgTindakan.Columns("KdPelayananRS").value)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dgTindakan.Columns("TglPelayanan").value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Delete_BiayaPelayananNew"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Penghapusan Biaya Pelayanan Pasien", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_BiayaPelayananNew")
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
'        strSQL = "SELECT * FROM PeriksaDiagnosa WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "' AND KdRuangan='" & mstrKdRuangan & "' AND KdDiagnosa='" & dgRiwayatDiagnosa.Columns("Kode ICD 10").value & "' AND TglPeriksa='" & Format(dgRiwayatDiagnosa.Columns("TglPeriksa").value, "yyyy/MM/dd HH:mm:ss") & "'"
        strSQL = "SELECT * FROM PeriksaDiagnosa WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "' AND KdDiagnosa='" & dgRiwayatDiagnosa.Columns("Kode ICD 10").value & "' AND TglPeriksa='" & Format(dgRiwayatDiagnosa.Columns("TglPeriksa").value, "yyyy/MM/dd HH:mm:ss") & "'"
        Set rsNew = Nothing
        rsNew.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, rsNew("KdRuangan").value)
        .Parameters.Append .CreateParameter("KdDiagnosa", adVarChar, adParamInput, 7, dgRiwayatDiagnosa.Columns("Kode ICD 10").value)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dgRiwayatDiagnosa.Columns("TglPeriksa").value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, rsNew("KdSubInstalasi").value)
        .Parameters.Append .CreateParameter("StatusKasus", adChar, adParamInput, 4, rsNew("StatusKasus").value)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, Right(txtNoCM.Text, 6))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Delete_Diagnosa"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            Call Add_HistoryLoginActivity("Delete_Diagnosa")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'Store procedure untuk menghapus catatan medis
Private Sub sp_DelCM(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dgCatatanMedis.Columns("Tgl. Periksa").value, "yyyy/MM/dd HH:mm:ss"))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Delete_CatatanMedis"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Penghapusan Catatan Medis Pasien", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_CatatanMedis")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If mblnFormDaftarAntrian = True Then Call frmDaftarAntrianPasien.cmdCari_Click
End Sub

Private Sub sstTP_GotFocus()
'    If txtSex.Text = "Laki-Laki" Then
'        cmdKehamilandanKB.Enabled = False
'    Else
'        cmdKehamilandanKB.Enabled = True
'    End If
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
        End Select
    End If
End Sub

Private Sub txtBln_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtGrandTotal_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)

End Sub

Private Sub txtHr_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtJenisPasien_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)

End Sub

Private Sub txtKls_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)

End Sub

Private Sub txtNamaPasien_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)

End Sub

Private Sub txtNoCM_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtNoPendaftaran_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then sstTP.SetFocus: sstTP.Tab = 0
    Call SetKeyPressToNumber(KeyAscii)
End Sub

'Untuk meload riwayat catatan klinis yang sudah pernah didapat
Public Sub subLoadRiwayatCatatanKlinis()
    On Error GoTo hell
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
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat catatan medis yang sudah pernah didapat
Public Sub subLoadRiwayatCatatanMedis()
    On Error GoTo hell
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
        .Columns(5).Width = 2500 'Pengobatan
        .Columns(6).Width = 1500 'Keterangan
        .Columns(7).Width = 2500 '[Dokter Pemeriksa]
        .Columns(8).Width = 0 'KdRuangan
    End With

    Exit Sub
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat Kecelakaan yang sudah pernah didapat
Public Sub subLoadRiwayatKecelakaan()
    On Error GoTo hell
    strSQL = "SELECT *" & _
    " FROM V_RiwayatKecelakanPasien " & _
    " WHERE (nocm = '" & Right(mstrNoCM, 6) & "')"
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
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat konsul pasien
Public Sub subLoadRiwayatKonsul()
    On Error GoTo hell
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
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat operasi yang sudah pernah didapat
Public Sub subLoadRiwayatOperasi()
    On Error GoTo hell
    strSQL = " SELECT     TOP (200) NoHasilPeriksa, NoCM, NoPendaftaran, KasusPenyakit, JenisTindakanMedis, TglMulaiPeriksa, TglAkhirPeriksa, TglHasilPeriksa" & _
    " FROM   V_HasilTindakanMedis " & _
    " WHERE (nocm = '" & mstrNoCM & "')"
    Call msubRecFO(rs, strSQL)
    Set dgRiwayatOperasi.DataSource = rs
    With dgRiwayatOperasi
        .Columns(0).Width = 1200 'NoHasilPeriksa
        .Columns(1).Width = 800 'NoCM
        .Columns(2).Width = 1250 'NoPendaftaran
        .Columns(3).Width = 2000 'kasus penyakt
        .Columns(4).Width = 2000 'jenis tindakan medis
        .Columns(5).Width = 2000 'tgl mulai
        .Columns(6).Width = 2000 'tgl akhir
        .Columns(7).Width = 2000 'tgl hasil
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Store procedure untuk menghapus catatan klinis
Private Sub sp_DelBiayaCatatanKlinis(ByVal adoCommand As ADODB.Command)
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dgCatatanKlinis.Columns("TglPeriksa").value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Delete_CatatanKlinis"
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
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dgCatatanMedis.Columns("TglPeriksa").value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Delete_CatatanMedis"
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
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dgCatatanMedis.Columns("TglPeriksa").value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo."
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penghapusan data kecelakaan", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'Store procedure untuk menghapus data konsul
Private Sub sp_DelKonsul(ByVal adoCommand As ADODB.Command)
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , Format(dgKonsul.Columns("TglDirujuk").value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Delete_PasienRujukan"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penghapusan data konsul", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_PasienRujukan")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

Private Sub txtSex_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtTglDaftar_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtThn_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtTindakanTotal_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
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

