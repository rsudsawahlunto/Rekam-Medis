VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPasienBaru 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pasien Baru"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPasienBaru.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   8865
   Begin VB.CheckBox chkDetail 
      Caption         =   "Detail"
      Height          =   210
      Left            =   120
      TabIndex        =   70
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   0
      TabIndex        =   61
      Top             =   7320
      Width           =   8775
      Begin VB.TextBox txtIbuKandung 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   6600
         MaxLength       =   50
         TabIndex        =   68
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtKepalaKeluarga 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   66
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtNoKK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   64
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtNamaPanggilan 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   240
         MaxLength       =   50
         TabIndex        =   62
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nama Ibu Kandung"
         Height          =   210
         Left            =   6600
         TabIndex        =   69
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Nama Kepala Keluarga"
         Height          =   210
         Left            =   4440
         TabIndex        =   67
         Top             =   360
         Width           =   1785
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "No. KK"
         Height          =   210
         Left            =   2400
         TabIndex        =   65
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nama Panggilan"
         Height          =   210
         Left            =   240
         TabIndex        =   63
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.TextBox txtNoReservasi 
      Height          =   375
      Left            =   2400
      TabIndex        =   60
      Text            =   "txtNoReservasi"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtFormPengirim 
      Height          =   375
      Left            =   600
      TabIndex        =   59
      Text            =   "txtFormPengirim"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtKdAntrian 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   7560
      MaxLength       =   15
      TabIndex        =   57
      Top             =   840
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dgDaftarBayiRSUD 
      Height          =   2775
      Left            =   4320
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
   Begin MSDataGridLib.DataGrid dgKeluargaPegawai 
      Height          =   2775
      Left            =   4200
      TabIndex        =   54
      Top             =   960
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
   Begin VB.ComboBox cbojnsPrinter 
      Height          =   330
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   0
      TabIndex        =   51
      Top             =   8520
      Width           =   8775
      Begin VB.CommandButton cmdRegMRS 
         Caption         =   "&Registrasi Pasien"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2160
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdDetailPasien 
         Caption         =   "&Detail Pasien"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4080
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   7200
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   5760
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraAlamatPas 
      Caption         =   "Alamat Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   41
      Top             =   4560
      Width           =   8775
      Begin VB.TextBox txtTelepon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5880
         MaxLength       =   15
         TabIndex        =   18
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtKodePos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7680
         MaxLength       =   5
         TabIndex        =   23
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtAlamat 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         MaxLength       =   100
         TabIndex        =   16
         Top             =   480
         Width           =   4575
      End
      Begin MSDataListLib.DataCombo dcKota 
         Height          =   390
         Left            =   4200
         TabIndex        =   20
         Top             =   1200
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcKecamatan 
         Height          =   390
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcKelurahan 
         Height          =   390
         Left            =   4200
         TabIndex        =   22
         Top             =   1920
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcPropinsi 
         Height          =   390
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         ListField       =   "k"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox meRTRW 
         Height          =   390
         Left            =   4920
         TabIndex        =   17
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Telepon"
         Height          =   210
         Left            =   5880
         TabIndex        =   49
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Kode Pos"
         Height          =   210
         Left            =   7680
         TabIndex        =   48
         Top             =   1680
         Width           =   765
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "RT/RW"
         Height          =   210
         Left            =   4920
         TabIndex        =   47
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Kelurahan (Desa)"
         Height          =   210
         Left            =   4200
         TabIndex        =   46
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Kecamatan"
         Height          =   210
         Left            =   240
         TabIndex        =   45
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Propinsi"
         Height          =   210
         Left            =   240
         TabIndex        =   44
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Kota (Kabupaten)"
         Height          =   210
         Left            =   4200
         TabIndex        =   43
         Top             =   960
         Width           =   1470
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Alamat Lengkap"
         Height          =   210
         Left            =   240
         TabIndex        =   42
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.Frame fraPasien 
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
      Height          =   1935
      Left            =   0
      TabIndex        =   32
      Top             =   2520
      Width           =   8775
      Begin VB.CheckBox chkCariBayi 
         Caption         =   "Cari Bayi Dari VK"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   350
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid dgPegawaiRSUD 
         Height          =   855
         Left            =   1680
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1508
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         HeadLines       =   0
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
      Begin VB.ComboBox cboNamaDepan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "frmPasienBaru.frx":0CCA
         Left            =   240
         List            =   "frmPasienBaru.frx":0CE3
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNoIdentitas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5760
         MaxLength       =   20
         TabIndex        =   9
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtHari 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7920
         MaxLength       =   2
         TabIndex        =   15
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtBulan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7200
         MaxLength       =   2
         TabIndex        =   14
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   7
         Top             =   600
         Width           =   3975
      End
      Begin VB.ComboBox cboJnsKelaminPasien 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmPasienBaru.frx":0D0E
         Left            =   240
         List            =   "frmPasienBaru.frx":0D18
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtTempatLahir 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1680
         MaxLength       =   25
         TabIndex        =   11
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtTahun 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   13
         Top             =   1320
         Width           =   615
      End
      Begin MSMask.MaskEdBox meTglLahir 
         Height          =   390
         Left            =   4800
         TabIndex        =   12
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   688
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         HideSelection   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mm-yy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Nama Depan"
         Height          =   210
         Left            =   240
         TabIndex        =   50
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Bulan"
         Height          =   210
         Left            =   7200
         TabIndex        =   40
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Hari"
         Height          =   210
         Left            =   7920
         TabIndex        =   39
         Top             =   1080
         Width           =   300
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Nama Lengkap"
         Height          =   210
         Left            =   1680
         TabIndex        =   38
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   240
         TabIndex        =   37
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label lblTmpLhr 
         AutoSize        =   -1  'True
         Caption         =   "Tempat Lahir"
         Height          =   210
         Left            =   1680
         TabIndex        =   36
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label lblTglLhr 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Lahir"
         Height          =   210
         Left            =   4920
         TabIndex        =   35
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label lblumur 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   6480
         TabIndex        =   34
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label lblGolDrh 
         AutoSize        =   -1  'True
         Caption         =   "No. Identitas (KTP/SIM/...)"
         Height          =   210
         Left            =   5760
         TabIndex        =   33
         Top             =   360
         Width           =   2235
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   29
      Top             =   1560
      Width           =   8775
      Begin VB.TextBox txtNoCMTerm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3120
         MaxLength       =   12
         TabIndex        =   71
         Top             =   480
         Width           =   2415
      End
      Begin VB.CheckBox chkNoCM 
         Caption         =   "No. CM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   56
         Top             =   220
         Width           =   1095
      End
      Begin VB.CheckBox chkPegawaiToPasien 
         Caption         =   "Pegawai RS / Keluarga Pegawai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
      Begin VB.ComboBox cboPegawaiToPasien 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "frmPasienBaru.frx":0D32
         Left            =   5640
         List            =   "frmPasienBaru.frx":0D34
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3120
         MaxLength       =   12
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtpTglPendaftaran 
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dddd,dd MMMM yyyy HH:mm"
         Format          =   138936323
         UpDown          =   -1  'True
         CurrentDate     =   38061
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   3120
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pendaftaran"
         Height          =   210
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1365
      End
   End
   Begin MSComctlLib.StatusBar stbInformasi 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   52
      Top             =   9375
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5159
            Text            =   "Cetak Kartu (F1)"
            TextSave        =   "Cetak Kartu (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5159
            Text            =   "Pasien Lama Ctrl+L"
            TextSave        =   "Pasien Lama Ctrl+L"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5159
            Text            =   "Cari Pasien (F3)"
            TextSave        =   "Cari Pasien (F3)"
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
      TabIndex        =   55
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Kode Antrian"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   58
      Top             =   840
      Width           =   3105
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6960
      Picture         =   "frmPasienBaru.frx":0D36
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPasienBaru.frx":1ABE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPasienBaru.frx":311C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmPasienBaru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Barcode39 As clsBarCode39
Dim subPrinterZebra As Printer
Dim X As String
Dim sIdPegawai As String
Dim sNoUrut As String
Dim sNoUrutBayiLahir As String
Dim sKdHubungan As String
Dim vTampil As Boolean
Dim vcektanggaLahir As Boolean
Dim HitungTanggalFromDatePicker As Boolean
Dim j As Integer

Dim strTutup As Boolean
Dim varPropinsi As String
Dim varKota As String
Dim varKecamatan As String
Dim varKelurahan As String

Dim varkdPropinsi As String
Dim varkdKota As String
Dim varkdKecamatan As String
Dim varkdKelurahan As String
Dim ValNoCM As Boolean

Private Function SimpanConvertPegawaiToPasien() As Boolean
    On Error GoTo hell_
    SimpanConvertPegawaiToPasien = True
    strSQL = "insert into ConvertPegawaiToPasien values('" & sIdPegawai & "','" & txtNoCM & "','" & sKdHubungan & "','" & sNoUrut & "')"
    dbConn.Execute strSQL
    Exit Function
hell_:
    SimpanConvertPegawaiToPasien = False
    msubPesanError
End Function

Private Function ConvertPasienToBayiLahir() As Boolean
    On Error GoTo hell_
    ConvertPasienToBayiLahir = True
    strSQL = "insert into ConvertPasienToBayiLahir values('" & txtNoCM & "','" & sNoUrutBayiLahir & "')"
    dbConn.Execute strSQL
    Exit Function
hell_:
    ConvertPasienToBayiLahir = False
    msubPesanError
End Function

Private Sub chkDetail_Click()
 If chkDetail.value = 1 Then
 
    txtNamaPanggilan.Enabled = True
    txtNoKK.Enabled = True
    txtKepalaKeluarga.Enabled = True
    txtIbuKandung.Enabled = True
 
 Else
    txtNamaPanggilan.Text = ""
    txtNoKK.Text = ""
    txtKepalaKeluarga.Text = ""
    txtIbuKandung.Text = ""
    
    txtNamaPanggilan.Enabled = False
    txtNoKK.Enabled = False
    txtKepalaKeluarga.Enabled = False
    txtIbuKandung.Enabled = False
 End If
 
 
End Sub

Private Sub chkDetail_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
       If chkDetail.value = 1 Then
        txtNamaPanggilan.SetFocus
       Else
         If cmdSimpan.Enabled = True Then
            cmdSimpan.SetFocus
         Else
            cmdTutup.SetFocus
         End If
        End If
    End If
   ' Call SetKeyPressToNumber(KeyAscii)
End Sub

'Private Sub cboNamaDepan_LostFocus()
'    Call cboNamaDepan_Change
'End Sub


Private Sub dcKelurahan_Click(Area As Integer)
If dcKelurahan.Text = "" Then Exit Sub
If varkdKelurahan <> dcKelurahan.BoundText Then
    txtKodePos = ""
    CekPilihanWilayah "dcKelurahan", "Click"
  End If
End Sub

Private Sub dgDaftarBayiRSUD_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDaftarBayiRSUD
    WheelHook.WheelHook dgDaftarBayiRSUD
End Sub

Private Sub dgKeluargaPegawai_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKeluargaPegawai
    WheelHook.WheelHook dgKeluargaPegawai
End Sub

Private Sub dgPegawaiRSUD_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPegawaiRSUD
    WheelHook.WheelHook dgPegawaiRSUD
End Sub

Private Sub Form_Activate()
    If bolAntrian = True Then
        txtKdAntrian.Enabled = True
        txtKdAntrian.SetFocus

    Else
        txtKdAntrian.Enabled = False
        cboNamaDepan.SetFocus

    End If
'    HitungTanggalFromDatePicker = False
End Sub

Private Sub meTglLahir_Change()
'If txtFormPengirim.Text <> "frmDaftarReservasiPasien" Then
'   If vcektanggaLahir = False Then Exit Sub
'End If
'    If meTglLahir.Text = "__/__/____" Then
'        txtTahun.SetFocus
'        Exit Sub
'    End If
'    If funcCekValidasiTgl("TglLahir", meTglLahir) = "NoErr" Then
'        Call subYearOldCount(Format(meTglLahir.Text, "yyyy/mm/dd"))
'        txtTahun.Text = YOC_intYear
'        txtBulan.Text = YOC_intMonth
'        txtHari.Text = YOC_intDay
'
'        If strPasien = "Lama" Or strPasien = "View" Or strPasien = "LamaReg" Then Exit Sub
'        txtAlamat.SetFocus
'        If (CCur(txtTahun.Text) = 0 And CCur(txtBulan.Text) = 0 And CCur(txtHari.Text) <= 28) Then
'            cboNamaDepan.Text = "Bayi"
'        ElseIf (CCur(txtTahun.Text) < 14 And CCur(txtBulan.Text) <= 12 And CCur(txtHari.Text) <= 31) Or (CCur(txtTahun.Text) = 14 And CCur(txtBulan.Text) = 0 And CCur(txtHari.Text) <= 2) Then
'            cboNamaDepan.Text = "An."
'        End If
'    Else
'        txtTahun.Text = ""
'        txtBulan.Text = ""
'        txtHari.Text = ""
'    End If

End Sub


Private Sub txtIbuKandung_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
  '  If KeyAscii = 13 Then txtAlamat.SetFocus
     If KeyAscii = 13 Then
        If cmdSimpan.Enabled = True Then cmdSimpan.SetFocus Else cmdTutup.SetFocus
    End If
    
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtKdAntrian_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglPendaftaran.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtKdAntrian_LostFocus()
    On Error GoTo hell_
    If AntrianForDataPasien = False Then
        If txtKdAntrian.Text = "" Then MsgBox "Isi Kode Antrian Pasien!!", vbInformation, "Validasi": txtKdAntrian.SetFocus: Exit Sub
        If Update_AntrianPasienRegistrasi(txtKdAntrian.Text, 0, 0, 0, 0, 0, "PROSES") = False Then Exit Sub
    Else
    
    End If
    
    Exit Sub
hell_:
    msubPesanError
End Sub

Private Sub cboJnsKelaminPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTempatLahir.SetFocus
End Sub

Private Sub subLoadBayi()
    On Error GoTo hell_
    Set rs = Nothing
    strSQL = "SELECT  NamaIbuBayi, JKBayi, TglLahir, NoUrutBayiLahir,Tahun,Bulan,Hari From [dbo].[V_DaftarNamaIbuBersalin] " & _
    " WHERE NamaIbuBayi LIKE '%" & txtNamaPasien.Text & "%'"
    Call msubRecFO(rs, strSQL)
    Set dgDaftarBayiRSUD.DataSource = rs
    With dgDaftarBayiRSUD
        .Columns("NamaIbuBayi").Width = 2000
        .Columns("JKBayi").Width = 600
        .Columns("JKBayi").Caption = "SEX"
        .Columns("TglLahir").Width = 1700
        .Columns("NoUrutBayiLahir").Width = 1300
        .Columns("Tahun").Width = 0
        .Columns("Bulan").Width = 0
        .Columns("Hari").Width = 0
        .Top = 2880
        .Left = 1680
    End With
    Exit Sub
hell_:
    msubPesanError
End Sub

Private Sub cboNamaDepan_Change()

    If cboNamaDepan.Text = "Bayi" Then
        cboJnsKelaminPasien.Enabled = True
        If MsgBox("Ingin Mendaftarkan Bayi Yang Lahir di RSUD [VK Bersalin]???", vbYesNo, "Konfirmasi") = vbNo Then
            chkCariBayi.Visible = False
            chkCariBayi.value = vbUnchecked
            If (txtNamaPasien.Enabled = True And txtNamaPasien.Visible = True) Then
                txtNamaPasien.SetFocus
            End If
        Else
            chkCariBayi.Visible = True
            chkCariBayi.value = vbChecked
            If (txtNamaPasien.Enabled = True And txtNamaPasien.Visible = True) Then
                txtNamaPasien.SetFocus
            End If
        End If
    ElseIf cboNamaDepan.Text = "Tn." Or cboNamaDepan.Text = "Sdr." Then
        cboJnsKelaminPasien.Text = "Laki-Laki"
        chkCariBayi.Visible = False
        chkCariBayi.value = vbUnchecked
        lblNamaPasien.Caption = "Nama Lengkap"
        dgDaftarBayiRSUD.Visible = False
        cboJnsKelaminPasien.Enabled = False
    ElseIf cboNamaDepan.Text = "Ny." Or cboNamaDepan.Text = "Nn." Or cboNamaDepan = "Sdri." Then
        cboJnsKelaminPasien.Text = "Perempuan"
        chkCariBayi.Visible = False
        chkCariBayi.value = vbUnchecked
        lblNamaPasien.Caption = "Nama Lengkap"
        dgDaftarBayiRSUD.Visible = False
        cboJnsKelaminPasien.Enabled = False
    Else
        cboJnsKelaminPasien.Enabled = True
        chkCariBayi.Visible = False
        chkCariBayi.value = vbUnchecked
        lblNamaPasien.Caption = "Nama Lengkap"
        dgDaftarBayiRSUD.Visible = False
    End If
End Sub

Private Sub cboNamaDepan_Click()
    If txtFormPengirim = "frmDaftarPasienAsuransi" Then
        Call cboNamaDepan_Change
    Else
        If boltampil = True Then Call cboNamaDepan_Change
    End If
    
    If bolDetailPasien = True Then
        Call cboNamaDepan_Change
    End If
    
End Sub

Private Sub cboNamaDepan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaPasien.SetFocus
End Sub

Private Sub cboPegawaiToPasien_Change()
    If cboPegawaiToPasien.Text = "Pegawai RSUD" Then
        cboNamaDepan.SetFocus
    Else
        cboNamaDepan.SetFocus
    End If
End Sub

Private Sub cboPegawaiToPasien_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cboNamaDepan.SetFocus
End Sub

Private Sub cboPegawaiToPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboNamaDepan.SetFocus
End Sub

Private Sub chkCariBayi_Click()
    On Error Resume Next
    If chkCariBayi.value = vbChecked Then
        lblNamaPasien.Caption = "Cari Nama Ibu Bayi"
        vTampil = True
    Else
        lblNamaPasien.Caption = "Nama Lengkap"
    End If
    txtNamaPasien.SetFocus
End Sub

Private Sub chkNoCM_Click()
    If chkNoCM.value = Checked Then
        txtNoCM.Enabled = True
        txtNoCM.SetFocus
    Else
        txtNoCM.Enabled = False
    End If
End Sub

Private Sub chkPegawaiToPasien_Click()
    If chkPegawaiToPasien.value = vbChecked Then
        cboPegawaiToPasien.Visible = True
        cboPegawaiToPasien.AddItem "Pegawai RSUD"
        cboPegawaiToPasien.AddItem "Keluarga Pegawai"
        vTampil = True
        cboPegawaiToPasien.SetFocus
    Else
        Call subClearData
        cboPegawaiToPasien.Visible = False
        dgPegawaiRSUD.Visible = False
        cboPegawaiToPasien.Clear
    End If
End Sub

Private Sub cmdDetailPasien_Click()
    On Error GoTo hell_
'    Load frmDetailPasien
    With frmDetailPasien
        .Show
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = txtNamaPasien.Text
        .txtJK.Text = cboJnsKelaminPasien.Text
        .txtThn.Text = txtTahun.Text
        .txtBln.Text = txtBulan.Text
        .txtHr.Text = txthari.Text
        .txtNoKK.Text = txtNoKK.Text
        .txtKepalaKeluarga.Text = txtKepalaKeluarga.Text
        .TxtIbu.Text = txtIbuKandung.Text
        .txtFormPengirim.Text = Me.Name
    End With
    Exit Sub
hell_:
    msubPesanError
End Sub

Private Sub cmdRegMRS_Click()
    On Error GoTo hell_
'    With frmRegistrasiAll
'        .Show
'        .txtKdAntrian = txtKdAntrian.Text
'        .txtnocm.Text = ""
'        .txtnocm.Text = txtnocm.Text
'        .txtNamaPasien.Text = txtNamaPasien.Text
'        .cboJK.Text = cboJnsKelaminPasien.Text
'        .txtThn.Text = txtTahun.Text
'        .txtBln.Text = txtBulan.Text
'        .txtHr.Text = txtHari.Text
'        If txtnocm.Text = "" Then .txtnocm.Text = mstrNoCM
'        If .dcInstalasi <> "" Then .dcSubInstalasi.SetFocus
'    End With
If txtFormPengirim.Text = Me.Name Then
    If MsgBox("Pasien mau ke RAWAT JALAN atau PENUNJANG ?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
        With frmRegistrasiRJPenunjang
            .Show
            .txtKdAntrian = txtKdAntrian.Text
            .txtFormPengirim = txtFormPengirim.Text
            .txtNoCM.Text = ""
            .txtNoCM.Text = txtNoCM.Text
            .txtNamaPasien.Text = txtNamaPasien.Text
            .cboJK.Text = cboJnsKelaminPasien.Text
            .txtThn.Text = txtTahun.Text
            .txtBln.Text = txtBulan.Text
            .txtHr.Text = txthari.Text
            If txtNoCM.Text = "" Then .txtNoCM.Text = mstrNoCM
        End With
    Else
        With frmRegistrasiAll
            .Show
            .txtKdAntrian = txtKdAntrian.Text
            .txtNoCM.Text = ""
            .txtNoCM.Text = txtNoCM.Text
            .txtNamaPasien.Text = txtNamaPasien.Text
            .cboJK.Text = cboJnsKelaminPasien.Text
            .txtThn.Text = txtTahun.Text
            .txtBln.Text = txtBulan.Text
            .txtHr.Text = txthari.Text
            If txtNoCM.Text = "" Then .txtNoCM.Text = mstrNoCM
            If .dcInstalasi <> "" Then .dcSubInstalasi.SetFocus
            
        End With
    End If
End If
    Unload Me
    Unload frmDetailPasien
    Exit Sub
hell_:
    msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    
    On Error GoTo errLoad
    
  If bolDetailPasien = False Then
    If dcKecamatan.Text <> "" Then
        If Periksa("datacombo", dcKecamatan, "Kecamatan Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcKelurahan.Text <> "" Then
        If Periksa("datacombo", dcKelurahan, "Kelurahan Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcKota.Text <> "" Then
        If Periksa("datacombo", dcKota, "Kota Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcPropinsi.Text <> "" Then
    If Periksa("datacombo", dcPropinsi, "Provinsi Tidak Terdaftar") = False Then Exit Sub
    End If
    
  Else
  
     If dcPropinsi.Text = "" Then
         MsgBox "Pilih Propinsi"
     ElseIf dcKota.Text = "" Then
        MsgBox "Pilih Kota / Kabupaten"
     ElseIf dcKecamatan.Text = "" Then
        MsgBox "Pilih Kecamatan"
     ElseIf dcKelurahan.Text = "" Then
        MsgBox "Pilih Kelurahan"
     End If
        
  End If
  
    If funcCekValidasi = False Then Exit Sub
    If txtTahun.Text = "" Then txtTahun.Text = 0
    If txtBulan.Text = "" Then txtBulan.Text = 0
    If txthari.Text = "" Then txthari.Text = 0

    Call sp_IdentitasPasien(dbcmd)
    Call sp_Add_NoCMTerm(dbcmd)
    
    If ValNoCM = False Then Exit Sub
    
    Call sp_DetailPasien(dbcmd)
    
    If txtFormPengirim.Text = "frmDaftarPasienAsuransi" Then
        dbConn.Execute "Update DataPesertaAsuransi set NoCM = '" & txtNoCM.Text & "' where IdPeserta = '" & frmDaftarPasienAsuransi.dgDaftarPasienAsuransi.Columns(1) & "'"
    End If
    
    If txtFormPengirim.Text = "frmDaftarReservasiPasien" Then
        dbConn.Execute "Update ReservasiPasien Set NoCM = '" & mstrNoCM & "' where Noreservasi ='" & txtNoReservasi.Text & "'"
    End If
    
    chkPegawaiToPasien.Enabled = False
    
    If chkPegawaiToPasien.value = vbChecked And cboPegawaiToPasien.Text = "Keluarga Pegawai" Then
        If SimpanConvertPegawaiToPasien = False Then Exit Sub
    End If
    
'    If chkCariBayi.value = vbChecked Then
'        If ConvertPasienToBayiLahir() = False Then Exit Sub
'    End If

    Call subEnableButtonReg(True)
    If strPasien = "Lama" Then
        If blnCariPasien = True Then
            Call frmCariPasien.cmdsearch_Click
            Me.ZOrder 0
        End If
    ElseIf strPasien = "LamaReg" Then
        frmDaftarPasienRJRIIGD.cmdCari_Click
        cmdTutup_Click
    Else
        If cmdRegMRS.Visible And cmdRegMRS.Enabled = True Then
            cmdRegMRS.SetFocus
        Else
            cmdTutup.SetFocus
        End If
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub
Private Sub sp_DetailPasien(ByVal adoCommand As ADODB.Command)

Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("NamaKeluarga", adVarChar, adParamInput, 50, "")
        .Parameters.Append .CreateParameter("WargaNegara", adChar, adParamInput, 1, "")
        .Parameters.Append .CreateParameter("GolDarah", adVarChar, adParamInput, 2, "")
        .Parameters.Append .CreateParameter("Rhesus", adChar, adParamInput, 1, "")
        .Parameters.Append .CreateParameter("StatusNikah", adVarChar, adParamInput, 10, "")
        .Parameters.Append .CreateParameter("Pekerjaan", adVarChar, adParamInput, 30, "")
        .Parameters.Append .CreateParameter("Agama", adVarChar, adParamInput, 20, "")
        .Parameters.Append .CreateParameter("Suku", adVarChar, adParamInput, 20, "")
        .Parameters.Append .CreateParameter("Pendidikan", adVarChar, adParamInput, 25, "")
        .Parameters.Append .CreateParameter("NamaAyah", adVarChar, adParamInput, 30, "")
        .Parameters.Append .CreateParameter("NamaIbu", adVarChar, adParamInput, 30, IIf(txtIbuKandung.Text = "", Null, txtIbuKandung.Text))
        .Parameters.Append .CreateParameter("NamaIstriSuami", adVarChar, adParamInput, 30, "")
        .Parameters.Append .CreateParameter("NoKK", adVarChar, adParamInput, 30, IIf(txtNoKK.Text = "", Null, txtNoKK.Text))
        .Parameters.Append .CreateParameter("NamaKepalaKeluarga", adVarChar, adParamInput, 30, IIf(txtKepalaKeluarga.Text = "", Null, txtKepalaKeluarga.Text))

        .ActiveConnection = dbConn
        .CommandText = "dbo.AU_DetailPasien "
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data Detail Pasien", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("AU_DetailPasien")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

Private Sub cmdTambah_Click()
    dgDaftarBayiRSUD.Visible = False
    chkCariBayi.value = vbUnchecked
    chkCariBayi.Visible = False
    Call subClearData
    Call subEnableButtonReg(False)
End Sub

Private Sub cmdTutup_Click()
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan data pasien baru ", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    strTutup = True
    strPasien = "View"
    bolDetailPasien = False
    Unload Me
End Sub

Private Sub dcKecamatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        j = 3
        dcKelurahan.Enabled = True
        Call subLoadDataWilayah("kecamatan")
        If dcKelurahan.Enabled = True Then
            dcKelurahan.SetFocus
        Else

        End If
    End If
    
End Sub

Private Sub dcKecamatan_LostFocus()
    dcKecamatan = Trim(StrConv(dcKecamatan, vbProperCase))
End Sub

Private Sub dcKecamatan_Click(Area As Integer)
   If dcKecamatan.Text = "" Then Exit Sub
   
   If varkdKecamatan <> dcKecamatan.BoundText Then
    dcKelurahan.Text = ""
    txtKodePos = ""
    CekPilihanWilayah "dcKecamatan", "Click"
   End If
End Sub

Private Sub dcKelurahan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        j = 4
        Call subLoadDataWilayah("desa")
        txtKodePos.SetFocus
    End If
End Sub

Private Sub dcKelurahan_LostFocus()
    dcKelurahan = Trim(StrConv(dcKelurahan, vbProperCase))
End Sub

Private Sub dcKota_Click(Area As Integer)
    If dcKota.Text = "" Then Exit Sub
    
    If varkdKota <> dcKota.BoundText Then
        dcKecamatan.Text = ""
        dcKelurahan.Text = ""
        txtKodePos = ""
        CekPilihanWilayah "dcKota", "Click"
        CekPilihanWilayah "dcKecamatan", "Click"
    End If
End Sub

Private Sub CekPilihanWilayah(strItem As String, Optional strEvent As String)
    Dim X As Integer
    Dim Y

    X = 0
    Select Case strItem
        Case "dcPropinsi"
            Set dcKota.RowSource = Nothing
            Set dcKecamatan.RowSource = Nothing
            Set dcKelurahan.RowSource = Nothing
            dcKota.Text = ""
            dcKecamatan.Text = ""
            dcKelurahan.Text = ""
            txtKodePos = ""
            Select Case strEvent
                Case "Click"
                    subDcSource "Kota", " where kdPropinsi = '" & dcPropinsi.BoundText & "' order by NamaKotaKabupaten"
                Case "KeyPress"
                    If dcPropinsi.MatchedWithList = False Then
                        MsgBox "Pilih Propinsi"
                        X = 1
                        GoTo kosong
                        dcPropinsi.SetFocus
                    Else
                        subDcSource "Kota", " where kdPropinsi = '" & dcPropinsi.BoundText & "' order by NamaKotaKabupaten"
                        dcKota.SetFocus
                    End If
                Case "LostFocus"
                    If dcPropinsi.MatchedWithList = False Then
                        MsgBox "Pilih Propinsi"
                        X = 1
                        GoTo kosong
                        dcPropinsi.SetFocus
                    Else
                        subDcSource "Kota", " where kdPropinsi = '" & dcPropinsi.BoundText & "' order by NamaKotaKabupaten"
                        dcKota.SetFocus
                    End If
            End Select
            varkdPropinsi = dcPropinsi.BoundText
        Case "dcKota"
            Set dcKecamatan.RowSource = Nothing
            Set dcKelurahan.RowSource = Nothing
            dcKecamatan.Text = ""
            dcKelurahan.Text = ""
            txtKodePos = ""
            If dcPropinsi.MatchedWithList = True Then
                Select Case strEvent
                    Case "Click"
                        If dcKota.Text = "" Then Exit Sub
                        subDcSource "Kecamatan", " where kdKotaKabupaten = '" & dcKota.BoundText & "' order by NamaKecamatan"
                    Case "KeyPress"
                        If dcKota.MatchedWithList = False Then
                           MsgBox "Pilih Kota"
                            X = 2
                            GoTo kosong
                            dcKota.SetFocus
                        Else
                            subDcSource "Kecamatan", " where kdKotaKabupaten = '" & dcKota.BoundText & "' order by NamaKecamatan"
                            dcKecamatan.SetFocus
                        End If
                    Case "LostFocus"
                        If dcKota.MatchedWithList = False Then
                            MsgBox "Pilih Kota"
                            X = 2
                            GoTo kosong
                            dcKota.SetFocus
                        Else
                            subDcSource "Kecamatan", " where kdKotaKabupaten = '" & dcKota.BoundText & "' order by NamaKecamatan"
                            dcKecamatan.SetFocus
                        End If
                End Select
                varkdKota = dcKota.BoundText
            End If
        Case "dcKecamatan"
            Set dcKelurahan.RowSource = Nothing
            dcKelurahan.Text = ""
            txtKodePos = ""
            If dcKota.MatchedWithList = True Then
                Select Case strEvent
                    Case "Click"
                        If dcKecamatan.Text = "" Then Exit Sub
                        subDcSource "Kelurahan", " where kdkecamatan = '" & dcKecamatan.BoundText & "' order by NamaKelurahan"
                    Case "KeyPress"
                        If dcKecamatan.MatchedWithList = False Then
                            MsgBox "Pilih Kecamatan"
                            X = 3
                            GoTo kosong
                            dcKecamatan.SetFocus
                        Else
                            subDcSource "Kelurahan", " where kdkecamatan = '" & dcKecamatan.BoundText & "' order by NamaKelurahan"
                            dcKelurahan.SetFocus
                        End If
                    Case "LostFocus"
                        If dcKecamatan.MatchedWithList = False Then
                            MsgBox "Pilih Kecamatan"
                            X = 3
                            GoTo kosong
                            dcKecamatan.SetFocus
                        Else
                            subDcSource "Kelurahan", " where kdkecamatan = '" & dcKecamatan.BoundText & "' order by NamaKelurahan"
                            dcKelurahan.SetFocus
                        End If
                End Select
                varkdKecamatan = dcKecamatan.BoundText
            End If
        Case "dcKelurahan"
            txtKodePos = ""
            If dcKecamatan.MatchedWithList = True Then
                Select Case strEvent
                    Case "KeyPress"
                        If dcKelurahan.MatchedWithList = False Then
                            MsgBox "Pilih Desa/Kelurahan"
                            X = 4
                            GoTo kosong
                            dcKelurahan.Text = ""
                            dcKelurahan.SetFocus
                        Else
                            txtKodePos.SetFocus
                        End If
                    Case "LostFocus"
                        If dcKelurahan.MatchedWithList = False Then
                            MsgBox "Pilih Desa/Kelurahan"
                            X = 4
                            GoTo kosong
                            dcKelurahan.SetFocus
                        End If
                End Select
                varkdKelurahan = dcKelurahan.BoundText
            End If
    End Select

    Exit Sub

kosong:
    Y = MsgBox("Mulai lagi dari awal", vbYesNo, "Wilayah") ' vbYesNoCancel
    Select Case Y
        Case vbYes
            dcPropinsi.Text = ""
            dcKota.Text = ""
            dcKecamatan.Text = ""
            dcKelurahan.Text = ""
            dcPropinsi.SetFocus
        Case vbNo
            Exit Sub

    End Select
End Sub


Private Sub CekPilihanWilayah2(strItem As String, Optional strEvent As String)
    Dim X As Integer
    Dim Y

    X = 0
    Select Case strItem
        Case "dcPropinsi"
'            Set dcKota.RowSource = Nothing
'            Set dcKecamatan.RowSource = Nothing
'            Set dcKelurahan.RowSource = Nothing
'            dcKota.Text = ""
'            dcKecamatan.Text = ""
'            dcKelurahan.Text = ""
'            txtKodePos = ""
            Select Case strEvent
                Case "Click"
                    subDcSource "Kota", " where kdPropinsi = '" & dcPropinsi.BoundText & "' order by NamaKotaKabupaten"
                Case "KeyPress"
                    If dcPropinsi.MatchedWithList = False Then
                        MsgBox "Pilih Propinsi"
                        X = 1
                        GoTo kosong
                        dcPropinsi.SetFocus
                    Else
                        subDcSource "Kota", " where kdPropinsi = '" & dcPropinsi.BoundText & "' order by NamaKotaKabupaten"
                        dcKota.SetFocus
                    End If
                Case "LostFocus"
                    If dcPropinsi.MatchedWithList = False Then
                        MsgBox "Pilih Propinsi"
                        X = 1
                        GoTo kosong
                        dcPropinsi.SetFocus
                    Else
                        subDcSource "Kota", " where kdPropinsi = '" & dcPropinsi.BoundText & "' order by NamaKotaKabupaten"
                        dcKota.SetFocus
                    End If
            End Select
        Case "dcKota"
'            Set dcKecamatan.RowSource = Nothing
'            Set dcKelurahan.RowSource = Nothing
'            dcKecamatan.Text = ""
'            dcKelurahan.Text = ""
'            txtKodePos = ""
            If dcPropinsi.MatchedWithList = True Then
                Select Case strEvent
                    Case "Click"
                        If dcKota.Text = "" Then Exit Sub
                        subDcSource "Kecamatan", " where kdKotaKabupaten = '" & dcKota.BoundText & "' order by NamaKecamatan"
                    Case "KeyPress"
                        If dcKota.MatchedWithList = False Then
                           MsgBox "Pilih Kota"
                            X = 2
                            GoTo kosong
                            dcKota.SetFocus
                        Else
                            subDcSource "Kecamatan", " where kdKotaKabupaten = '" & dcKota.BoundText & "' order by NamaKecamatan"
                            dcKecamatan.SetFocus
                        End If
                    Case "LostFocus"
                        If dcKota.MatchedWithList = False Then
                            MsgBox "Pilih Kota"
                            X = 2
                            GoTo kosong
                            dcKota.SetFocus
                        Else
                            subDcSource "Kecamatan", " where kdKotaKabupaten = '" & dcKota.BoundText & "' order by NamaKecamatan"
                            dcKecamatan.SetFocus
                        End If
                End Select
            End If
        Case "dcKecamatan"
            Set dcKelurahan.RowSource = Nothing
'            dcKelurahan.Text = ""
'            txtKodePos = ""
            If dcKota.MatchedWithList = True Then
                Select Case strEvent
                    Case "Click"
                        If dcKecamatan.Text = "" Then Exit Sub
                        subDcSource "Kelurahan", " where kdkecamatan = '" & dcKecamatan.BoundText & "' order by NamaKelurahan"
                    Case "KeyPress"
                        If dcKecamatan.MatchedWithList = False Then
                            MsgBox "Pilih Kecamatan"
                            X = 3
                            GoTo kosong
                            dcKecamatan.SetFocus
                        Else
                            subDcSource "Kelurahan", " where kdkecamatan = '" & dcKecamatan.BoundText & "' order by NamaKelurahan"
                            dcKelurahan.SetFocus
                        End If
                    Case "LostFocus"
                        If dcKecamatan.MatchedWithList = False Then
                            MsgBox "Pilih Kecamatan"
                            X = 3
                            GoTo kosong
                            dcKecamatan.SetFocus
                        Else
                            subDcSource "Kelurahan", " where kdkecamatan = '" & dcKecamatan.BoundText & "' order by NamaKelurahan"
                            dcKelurahan.SetFocus
                        End If
                End Select
            End If
        Case "dcKelurahan"
'            txtKodePos = ""
            If dcKecamatan.MatchedWithList = True Then
                Select Case strEvent
                    Case "KeyPress"
                        If dcKelurahan.MatchedWithList = False Then
                            MsgBox "Pilih Desa/Kelurahan"
                            X = 4
                            GoTo kosong
                            dcKelurahan.Text = ""
                            dcKelurahan.SetFocus
                        Else
                            txtKodePos.SetFocus
                        End If
                    Case "LostFocus"
                        If dcKelurahan.MatchedWithList = False Then
                            MsgBox "Pilih Desa/Kelurahan"
                            X = 4
                            GoTo kosong
                            dcKelurahan.SetFocus
                        End If
                End Select
            End If
    End Select

    Exit Sub

kosong:
    Y = MsgBox("Mulai lagi dari awal", vbYesNo, "Wilayah") ' vbYesNoCancel
    Select Case Y
        Case vbYes
            dcPropinsi.Text = ""
            dcKota.Text = ""
            dcKecamatan.Text = ""
            dcKelurahan.Text = ""
            dcPropinsi.SetFocus
        Case vbNo
            Exit Sub

    End Select
End Sub

Private Sub dcKota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        j = 2
        dcKecamatan.Enabled = True
        dcKelurahan.Enabled = True
        Call subLoadDataWilayah("kota")
        CekPilihanWilayah "dcKota", "Click"
        If dcKecamatan.Enabled = True Then
            dcKecamatan.SetFocus
        End If
    End If
End Sub

Private Sub dcKota_LostFocus()
    dcKota = Trim(StrConv(dcKota, vbProperCase))
End Sub

Private Sub dcPropinsi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        j = 1
        dcKota.Enabled = True
        dcKecamatan.Enabled = True
        dcKelurahan.Enabled = True
        Call subLoadDataWilayah("propinsi")
        If dcKota.Enabled = True Then
            dcKota.SetFocus
        End If
    End If
End Sub

Private Sub dcPropinsi_Click(Area As Integer)
'    If dcPropinsi.DataField = "" Then Exit Sub
    If varkdPropinsi <> dcPropinsi.BoundText Then
        CekPilihanWilayah "dcPropinsi", "Click"
        dcKota.Text = ""
        dcKecamatan.Text = ""
        dcKelurahan.Text = ""
        txtKodePos = ""
    End If
End Sub


Private Sub dcPropinsi_LostFocus()
    dcPropinsi = Trim(StrConv(dcPropinsi, vbProperCase))
'    CekPilihanWilayah "dcPropinsi", "LostFocus"
End Sub

Private Sub dgDaftarBayiRSUD_DblClick()
    Call dgDaftarBayiRSUD_KeyPress(13)
End Sub

Private Sub dgDaftarBayiRSUD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        vTampil = False
        
        With dgDaftarBayiRSUD
        If (.row <> -1) Then
            txtNamaPasien.Text = .Columns("NamaIbuBayi")
            txtTahun.Text = .Columns("Tahun")
            txtBulan.Text = .Columns("Bulan")
            txthari.Text = .Columns("Hari")
            sNoUrutBayiLahir = .Columns("NoUrutBayiLahir")
            If .Columns("SEX").value = "L" Then
                cboJnsKelaminPasien.Text = "Laki-Laki"
            Else
                cboJnsKelaminPasien.Text = "Perempuan"
            End If
            dgDaftarBayiRSUD.Visible = False
        Else
            MsgBox "Data Bayi tidak ditemukan"
            txtNamaPasien.Text = ""
        End If
        End With
        dgDaftarBayiRSUD.Visible = False
        txtNoIdentitas.SetFocus
        vTampil = True
        dgDaftarBayiRSUD.Visible = False
    
    End If
End Sub

Private Sub dgKeluargaPegawai_DblClick()
    Call dgKeluargaPegawai_KeyPress(13)
End Sub

Private Sub dgKeluargaPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        vTampil = False
        vcektanggaLahir = True
        With dgKeluargaPegawai
            txtNamaPasien.Text = .Columns("NamaKeluarga")
            meTglLahir.Text = .Columns("TglLahir")
            sNoUrut = .Columns("NoUrut")
            sIdPegawai = .Columns("IdPegawai")
            sKdHubungan = .Columns("KdHubungan")
            If .Columns(1).value = "L" Then
                cboJnsKelaminPasien.Text = "Laki-Laki"
            Else
                cboJnsKelaminPasien.Text = "Perempuan"
            End If
        End With
        dgKeluargaPegawai.Visible = False
        txtNoIdentitas.SetFocus
'        meTglLahir_Change
        vcektanggaLahir = False
        vTampil = True
    End If
End Sub

Private Sub dgPegawaiRSUD_DblClick()
    Call dgPegawaiRSUD_KeyPress(13)
End Sub

Private Sub dgPegawaiRSUD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    vTampil = False
    vcektanggaLahir = True
        With dgPegawaiRSUD
            txtNamaPasien.Text = .Columns("NamaLengkap")
            txtTempatLahir.Text = .Columns("TempatLahir")
            If .Columns("TglLahir") = "" Then
                meTglLahir.Text = "__/__/____"
            Else
                meTglLahir.Text = .Columns("TglLahir")
            End If
            If .Columns(1).value = "L" Then
                cboJnsKelaminPasien.Text = "Laki-Laki"
            Else
                cboJnsKelaminPasien.Text = "Perempuan"
            End If
            
            If .Columns("KdTitle") = "01" Then
               
            ElseIf .Columns("KdTitle") = "02" Then
                cboNamaDepan.Text = "Nn."
                cboNamaDepan.Enabled = False
            ElseIf .Columns("KdTitle") = "03" Then
                cboNamaDepan.Text = "Ny."
                cboNamaDepan.Enabled = False
            ElseIf .Columns("KdTitle") = "04" Then
                cboNamaDepan.Text = "Tn."
                cboNamaDepan.Enabled = False
            End If
            
            If .Columns("KdTitle") = "" Then
               If .Columns(1).value = "L" Then
                   cboNamaDepan.Text = "Tn."
                   cboNamaDepan.Enabled = False
               Else
                   cboNamaDepan.Text = "Tn."
                   cboNamaDepan.Enabled = False
               End If
            End If
                              
        End With
        dgPegawaiRSUD.Visible = False
        txtNoIdentitas.SetFocus
'        Call meTglLahir_LostFocus
    vcektanggaLahir = False
    vTampil = True
    End If
End Sub

Private Sub dtpTglPendaftaran_Change()
    dtpTglPendaftaran.MaxDate = Now
End Sub

Private Sub dtpTglPendaftaran_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cboNamaDepan.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim subCtrlKey As String
    On Error Resume Next
    subCtrlKey = (Shift + vbCtrlMask)
    Select Case KeyCode
        Case vbKeyF1
            If txtNoCM.Text = "" Then Exit Sub
            mstrNoCM = txtNoCM.Text
            
            strSQL = "SELECT NoCM, NamaLengkap, Alamat, JenisKelamin, Propinsi, Kota, Kecamatan, Kelurahan, TglDaftarMembership, Umur,RTRW" & _
            " FROM   V_CetakKartuKuningPasien" & _
            " WHERE NoCM = '" & mstrNoCM & "'"
        
            Set rs = Nothing
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
'                frmCetakKartuKuning.Show
                Call subPrintRegistrasiBarcode
            Else
                MsgBox "Data pasien harus dilengkapi", vbInformation
            End If
        Case vbKeyL
            If subCtrlKey = 4 Then
                Unload Me
                frmRegistrasiAll.Show
            End If
        Case vbKeyF3
            Unload Me
            frmCariPasien.Show
        Case 27
            dgKeluargaPegawai.Visible = False
            dgPegawaiRSUD.Visible = False
            dgDaftarBayiRSUD.Visible = False
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglPendaftaran.MaxDate = Now
    dtpTglPendaftaran.value = Now
    j = 0
    
    strBanyakNoCM = 0
    strSQL = "Select value from SettingGlobal where Prefix = 'LenNoCM'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        strBanyakNoCM = rs(0).value
'    Else
'        strBanyakNoCM = "6"
    End If
    
    txtNoCM.MaxLength = strBanyakNoCM
    
    txtNamaPanggilan.Enabled = False
    txtNoKK.Enabled = False
    txtKepalaKeluarga.Enabled = False
    txtIbuKandung.Enabled = False
    
    subDcSource "Propinsi"
    If strPasien = "Lama" Then
        Call subEnableButtonReg(True)
        Call subVisibleButtonReg(True)
        cmdSimpan.Enabled = True
    ElseIf strPasien = "View" Then
        Call subEnableButtonReg(True)
        Call subVisibleButtonReg(False)
        cmdSimpan.Enabled = True
        cmdRegMRS.Visible = False
    ElseIf strPasien = "LamaReg" Then
        Call subEnableButtonReg(True)
        Call subVisibleButtonReg(False)
        cmdSimpan.Visible = True
        cmdSimpan.Enabled = True
        cmdRegMRS.Visible = False
    ElseIf strPasien = "Reservasi" Then
        Call subEnableButtonReg(True)
        Call subVisibleButtonReg(True)
        cmdSimpan.Enabled = True
        cmdRegMRS.Visible = False
    ElseIf strPasien = "Asuransi" Then
        Call subEnableButtonReg(True)
        Call subVisibleButtonReg(True)
        cmdSimpan.Enabled = True
        cmdRegMRS.Visible = False

    ElseIf strPasien = "LamaReg" Then


    End If

    If LCase(strPasien) = "baru" Then
        Call cmdTambah_Click
    ElseIf strPasien = "Reservasi" Then
        Call cmdTambah_Click
    Else
        Call subLoadDataPasien(mstrNoCM)
    End If
 

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If strTutup = True Then strTutup = False Else Call cmdTutup_Click
    If strPasien = "View" Then
        If strRegistrasi = "RJ" Then
        ElseIf strRegistrasi = "DaftarPasienRIRJIGD" Then
            Call frmDaftarPasienRJRIIGD.cmdCari_Click
        ElseIf strRegistrasi = "PasienLama" Then
            Call frmRegistrasiAll.CariData
        End If
    End If
    If bolStatusFrmAntrian = True Then
        frmDaftarAntrianRegistrasi.cmdCari_Click: bolStatusFrmAntrian = False
        frmDaftarAntrianRegistrasi.Enabled = True
    End If
End Sub

Private Sub meRTRW_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTelepon.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub meRTRW_LostFocus()
    If meRTRW.Text <> "__/__" Then
        If InStr(1, meRTRW.Text, "_") <> 0 Then
            MsgBox "Format RT/RW adalah 00/00, harap isi RT/RW dengan benar", vbCritical, "Validasi"
            meRTRW.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub meTglLahir_KeyPress(KeyAscii As Integer)
    On Error GoTo errTglLahir
    If KeyAscii = 13 Then
'         HitungTanggalFromDatePicker = True
        If meTglLahir.Text = "__/__/____" Then
            txtTahun.SetFocus
            Exit Sub
        End If
        If funcCekValidasiTgl("TglLahir", meTglLahir) = "NoErr" Then
            Call subYearOldCount(Format(meTglLahir.Text, "yyyy/mm/dd"))
            txtTahun.Text = YOC_intYear
            txtBulan.Text = YOC_intMonth
            txthari.Text = YOC_intDay
            
            
'            strSQL12 = "Select * from Pasien where NamaLengkap like '" & txtNamaPasien.Text & "%' And TglLahir  ='" & Format(meTglLahir, "yyyy/mm/dd") & "'"
'            Call msubRecFO(dbRst2, strSQL12)
'            If dbRst2.EOF = False Then
'                If MsgBox("Pasien dengan Nama " & txtNamaPasien.Text & " dengan tanggal lahir " & meTglLahir & " sudah ada" & vbNewLine & " Apakah anda akan melanjutkan mendaftarkan pasien ? ", vbInformation + vbYesNo) = vbNo Then Exit Sub
'            End If
            
            If strPasien = "Lama" Or strPasien = "View" Or strPasien = "LamaReg" Then Exit Sub
            'txtNamaPanggilan.SetFocus
            txtAlamat.SetFocus
'            If (CCur(txtTahun.Text) = 0 And CCur(txtBulan.Text) = 0 And CCur(txtHari.Text) <= 28) Then
'                cboNamaDepan.Text = "Bayi"
'            ElseIf (CCur(txtTahun.Text) < 14 And CCur(txtBulan.Text) <= 12 And CCur(txtHari.Text) <= 31) Or (CCur(txtTahun.Text) = 14 And CCur(txtBulan.Text) = 0 And CCur(txtHari.Text) <= 2) Then
'                cboNamaDepan.Text = "An."
'            End If
        Else
            txtTahun.Text = ""
            txtBulan.Text = ""
            txthari.Text = ""
        End If
    End If
    Call SetKeyPressToNumber(KeyAscii)
    Exit Sub
errTglLahir:
    If Err.Number = 5 Then Exit Sub
    MsgBox "Sudahkah anda mengganti Regional Setting komputer anda menjadi 'Indonesia'?" _
    & vbNewLine & "Bila sudah hubungi Administrator dan laporkan pesan kesalahan berikut:" _
    & vbNewLine & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
End Sub

Private Sub meTglLahir_LostFocus()
    On Error GoTo errTglLahir
    If meTglLahir.Text = "__/__/____" Then Exit Sub
    If funcCekValidasiTgl("TglLahir", meTglLahir) = "NoErr" Then
        Call subYearOldCount(Format(meTglLahir.Text, "yyyy/mm/dd"))
        txtTahun.Text = YOC_intYear
        txtBulan.Text = YOC_intMonth
        txthari.Text = YOC_intDay
            
        strSQL12 = "Select * from Pasien where NamaLengkap like '%" & txtNamaPasien.Text & "%' And TglLahir  ='" & Format(meTglLahir, "yyyy/mm/dd") & "'"
        Call msubRecFO(dbRst2, strSQL12)
        If dbRst2.EOF = False Then
        If MsgBox("Pasien dengan Nama " & txtNamaPasien.Text & " dengan tanggal lahir " & meTglLahir & " sudah ada" & vbNewLine & " Apakah anda akan melanjutkan mendaftarkan pasien ? ", vbInformation + vbYesNo) = vbNo Then Call subClearData: Exit Sub
        End If
    Else
        txtTahun.Text = ""
        txtBulan.Text = ""
        txthari.Text = ""
    End If

    Exit Sub
errTglLahir:
    MsgBox "Sudahkah anda mengganti Regional Setting komputer anda menjadi 'Indonesia'?" _
    & vbNewLine & "Bila sudah hubungi Administrator dan laporkan pesan kesalahan berikut:" _
    & vbNewLine & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
End Sub

Private Sub txtAlamat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then meRTRW.SetFocus
End Sub

Private Sub txtAlamat_LostFocus()
    txtAlamat = StrConv(txtAlamat, vbUpperCase)
End Sub

Private Sub txtBulan_Change()
    Dim dTglLahir As Date
'    If (HitungTanggalFromDatePicker = True) Then Exit Sub
    If txtBulan.Text = "" And txtTahun.Text = "" Then txthari.SetFocus: Exit Sub
    If txtBulan.Text = "" Then txtBulan.Text = 0
    If txtTahun.Text = "" And txthari.Text = "" Then
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
    ElseIf txtTahun.Text <> "" And txthari.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txthari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    ElseIf txtTahun.Text = "" And txthari.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txthari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
    ElseIf txtTahun.Text <> "" And txthari.Text = "" Then
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    End If
        meTglLahir.Text = dTglLahir
    
End Sub

Private Sub txtBulan_KeyPress(KeyAscii As Integer)
    Dim dTglLahir As Date
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If txtBulan.Text = "" And txtTahun.Text = "" Then txthari.SetFocus: Exit Sub
        If txtBulan.Text = "" Then txtBulan.Text = 0
        If txtTahun.Text = "" And txthari.Text = "" Then
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
        ElseIf txtTahun.Text <> "" And txthari.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txthari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        ElseIf txtTahun.Text = "" And txthari.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txthari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        ElseIf txtTahun.Text <> "" And txthari.Text = "" Then
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        End If
        meTglLahir.Text = dTglLahir
        txthari.SetFocus
    End If
End Sub

Private Sub txtHari_Change()
    Dim dTglLahir As Date
'    If (HitungTanggalFromDatePicker = True) Then Exit Sub
     If txthari.Text = "" And txtBulan.Text = "" And txtTahun.Text = "" Then txtAlamat.SetFocus: Exit Sub
    If txthari.Text = "" Then txthari.Text = 0
    If txtTahun.Text = "" And txtBulan.Text = "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txthari.Text), Date)
    ElseIf txtTahun.Text <> "" And txtBulan.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txthari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    ElseIf txtTahun.Text = "" And txtBulan.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txthari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
    ElseIf txtTahun.Text <> "" And txtBulan.Text = "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txthari.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    End If
        meTglLahir.Text = dTglLahir
End Sub

Private Sub txtHari_KeyPress(KeyAscii As Integer)
    Dim dTglLahir As Date
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If txthari.Text = "" And txtBulan.Text = "" And txtTahun.Text = "" Then txtAlamat.SetFocus: Exit Sub
        If txthari.Text = "" Then txthari.Text = 0
        If txtTahun.Text = "" And txtBulan.Text = "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txthari.Text), Date)
        ElseIf txtTahun.Text <> "" And txtBulan.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txthari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        ElseIf txtTahun.Text = "" And txtBulan.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txthari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        ElseIf txtTahun.Text <> "" And txtBulan.Text = "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txthari.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        End If
        meTglLahir.Text = dTglLahir
      '  txtNamaPanggilan.SetFocus
       txtAlamat.SetFocus
    End If
End Sub

Private Sub txtKodePos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        Call subLoadDataWilayah("kodepos")
        'If cmdSimpan.Enabled = True Then cmdSimpan.SetFocus Else cmdTutup.SetFocus
        If cmdSimpan.Enabled = True Then chkDetail.SetFocus Else cmdTutup.SetFocus
    End If
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtNamaPanggilan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNoKK.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub
Private Sub txtNoKK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtKepalaKeluarga.SetFocus
End Sub
Private Sub txtKepalaKeluarga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtIbuKandung.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtNamaPasien_Change()
    On Error Resume Next
    If vTampil = False Then Exit Sub
    If chkPegawaiToPasien.value = vbChecked And cboPegawaiToPasien.Text = "Pegawai RSUD" Then
        dgPegawaiRSUD.Visible = True
        strSQL = "SELECT NamaLengkap, JenisKelamin, TempatLahir, TglLahir,KdTitle FROM DataPegawai WHERE NamaLengkap like'" & txtNamaPasien.Text & "%'ORDER BY NamaLengkap"
        Call msubRecFO(rs, strSQL)
        Set dgPegawaiRSUD.DataSource = rs
        With dgPegawaiRSUD
            .Columns(0).Width = 3300
            .Columns(1).Width = 0
            .Columns(2).Width = 0
            .Columns(3).Width = 0
            .Columns(4).Width = 0
        End With
    End If
    If chkPegawaiToPasien.value = vbChecked And cboPegawaiToPasien.Text = "Keluarga Pegawai" Then
        dgKeluargaPegawai.Visible = True
        strSQL = "SELECT NamaKeluarga, JenisKelamin, TglLahir, NamaHubungan, NamaPegawai, NoUrut,IdPegawai,KdHubungan FROM    V_KeluargaPegawai" & _
        " WHERE NamaKeluarga like'" & txtNamaPasien.Text & "%'ORDER BY NamaKeluarga"
        Call msubRecFO(rs, strSQL)
        Set dgKeluargaPegawai.DataSource = rs
        With dgKeluargaPegawai
            .Columns(0).Width = 2500
            .Columns(1).Width = 400
            .Columns(1).Caption = "JK"
            .Columns(2).Width = 1200
            .Columns(3).Width = 1200
            .Columns(3).Caption = "Hubungan"
            .Columns(4).Width = 2000
            .Columns(5).Width = 600
            .Columns(6).Width = 0
            .Columns(7).Width = 0
        End With
    End If
    If chkCariBayi.value = vbChecked Then
        dgDaftarBayiRSUD.Visible = True
        Set rs = Nothing
        strSQL = "SELECT  NamaIbuBayi, JKBayi, TglLahir, NoUrutBayiLahir,Tahun,Bulan,Hari From [dbo].[V_DaftarNamaIbuBersalin] " & _
        " WHERE NamaIbuBayi LIKE '%" & txtNamaPasien.Text & "%'"
        Call msubRecFO(rs, strSQL)
        Set dgDaftarBayiRSUD.DataSource = rs
        With dgDaftarBayiRSUD
            .Columns("NamaIbuBayi").Width = 2000
            .Columns("JKBayi").Width = 600
            .Columns("JKBayi").Caption = "SEX"
            .Columns("TglLahir").Width = 1700
            .Columns("NoUrutBayiLahir").Width = 1300
            .Columns("Tahun").Width = 0
            .Columns("Bulan").Width = 0
            .Columns("Hari").Width = 0
            .Top = 3720
            .Left = 1680
        End With
'        Call subLoadBayi
    End If
End Sub

Private Sub txtNamaPasien_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If dgPegawaiRSUD.Visible = False Then Exit Sub
        dgPegawaiRSUD.SetFocus
    End If
End Sub

Private Sub txtNamaPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dgPegawaiRSUD.Visible = True Then
            dgPegawaiRSUD.SetFocus
        Else
            txtNoIdentitas.SetFocus
        End If
    End If
    If KeyAscii = 27 Then dgPegawaiRSUD.Visible = False
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtNamaPasien_LostFocus()
    txtNamaPasien = StrConv(txtNamaPasien, vbProperCase)
End Sub

Private Sub txtNoCM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        strSQL = "SELECT NoCM, Title + ' ' + [Nama Lengkap] AS NamaPasien FROM V_CariPasien WHERE ([No. CM] = '" & txtNoCM.Text & "' )"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            MsgBox "No. CM tersebut sudah dipakai " & rs("NamaPasien").value & "", vbExclamation, "Validasi"
            Exit Sub
        Else
            cboNamaDepan.SetFocus
        End If
    End If
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = 13 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    If KeyAscii = Asc(",") Then Exit Sub
    If KeyAscii = Asc(".") Then Exit Sub
End Sub

Private Sub txtNoCM_LostFocus()
    Call txtNoCM_KeyPress(13)
End Sub

Private Sub txtNoIdentitas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboJnsKelaminPasien.Enabled = False Then
            txtTempatLahir.SetFocus
        Else
            cboJnsKelaminPasien.SetFocus
        End If
    End If
End Sub

Private Sub txtNoIdentitas_LostFocus()
    txtNoIdentitas = StrConv(txtNoIdentitas, vbProperCase)
End Sub

Private Sub txtTahun_Change()
    Dim dTglLahir As Date
'    If (HitungTanggalFromDatePicker = True) Then Exit Sub
    If txtTahun = "" Then txtBulan.SetFocus: Exit Sub
    If txtBulan.Text = "" And txthari.Text = "" Then
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), Date)
    ElseIf txtBulan.Text <> "" And txthari.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txthari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    ElseIf txtBulan.Text = "" And txthari.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txthari.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    ElseIf txtBulan.Text <> "" And txthari.Text = "" Then
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    End If
    meTglLahir.Text = dTglLahir

End Sub

Private Sub txtTahun_KeyPress(KeyAscii As Integer)
    Dim dTglLahir As Date
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If txtTahun = "" Then txtBulan.SetFocus: Exit Sub
        If txtBulan.Text = "" And txthari.Text = "" Then
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), Date)
        ElseIf txtBulan.Text <> "" And txthari.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txthari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        ElseIf txtBulan.Text = "" And txthari.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txthari.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        ElseIf txtBulan.Text <> "" And txthari.Text = "" Then
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        End If
        meTglLahir.Text = dTglLahir
        txtBulan.SetFocus
    End If
End Sub

Private Sub txtTelepon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcPropinsi.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtTempatLahir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then meTglLahir.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtTempatLahir_LostFocus()
    txtTempatLahir = StrConv(txtTempatLahir, vbProperCase)
End Sub

'untuk cek validasi
Private Function funcCekValidasi() As Boolean
    If cboNamaDepan.Text = "" Then
        MsgBox "Titel Pasien harus diisi", vbExclamation, "Validasi"
        funcCekValidasi = False
        cboNamaDepan.SetFocus
        Exit Function
    End If
    If txtNamaPasien.Text = "" Then
        MsgBox "Nama Pasien harus diisi", vbExclamation, "Validasi"
        funcCekValidasi = False
        txtNamaPasien.SetFocus
        Exit Function
    End If
    If meTglLahir.Text = "__/__/____" Then
        MsgBox "Tanggal Lahir Pasien harus diisi", vbExclamation, "Validasi"
        funcCekValidasi = False
        meTglLahir.SetFocus
        Exit Function
    End If
    If cboJnsKelaminPasien.Text = "" Then
        MsgBox "Jenis Kelamin Pasien harus diisi", vbExclamation, "Validasi"
        funcCekValidasi = False
        cboJnsKelaminPasien.SetFocus
        Exit Function
    End If
    If dcPropinsi.Text = "" Then
        MsgBox "Propinsi Harus diisi", vbExclamation, "Validasi"
        funcCekValidasi = False
        dcPropinsi.SetFocus
        Exit Function
    End If
    If dcKota.Text = "" Then
        MsgBox "Kota harus diisi", vbExclamation, "Validasi"
        funcCekValidasi = False
        dcKota.SetFocus
        Exit Function
    End If
    If dcKecamatan.Text = "" Then
        MsgBox "Kecamatan harus diisi", vbExclamation, "Validasi"
        funcCekValidasi = False
        dcKecamatan.SetFocus
        Exit Function
    End If
    If dcKelurahan.Text = "" Then
        MsgBox "Kelurahan harus diisi", vbExclamation, "Validasi"
        funcCekValidasi = False
        dcKelurahan.SetFocus
        Exit Function
    End If
    
    funcCekValidasi = True
End Function

'Store procedure untuk mengisi identitas pasien

Private Sub sp_IdentitasPasien(ByVal adoCommand As ADODB.Command)
On Error GoTo errLoad
    Dim strLokal As String
   
    
    Const angka = "0123456789"
   ' Const huruf = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ "
    
        
    If InStr(angka, (strBanyakNoCM)) = 0 Then
       MsgBox "Ada kesalahan dalam inputan banyaknya NoCM pada Sistem Administrator", vbCritical, "Validasi"
       ValNoCM = False
       Exit Sub
    End If
    
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        
'        If rs.EOF Then
'          .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, IIf(Trim(txtNoCM.Text) = "", Null, Trim(txtNoCM.Text)))
'        Else
           .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, IIf(strBanyakNoCM = 0, 12, strBanyakNoCM), IIf(Trim(txtNoCM.Text) = "", Null, Trim(txtNoCM.Text)))
'
'        End If
        If txtNoIdentitas.Text <> "" Then
            .Parameters.Append .CreateParameter("NoIdentitas", adVarChar, adParamInput, 20, Trim(txtNoIdentitas.Text))
        Else
            .Parameters.Append .CreateParameter("NoIdentitas", adVarChar, adParamInput, 20, Null)
        End If
        .Parameters.Append .CreateParameter("TglDaftarMembership", adDate, adParamInput, , Format(dtpTglPendaftaran.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TitlePasien", adVarChar, adParamInput, 4, cboNamaDepan.Text)
        .Parameters.Append .CreateParameter("NamaLengkap", adVarChar, adParamInput, 50, Trim(txtNamaPasien.Text))
        .Parameters.Append .CreateParameter("NamaPanggilan", adVarChar, adParamInput, 50, Trim(txtNamaPanggilan.Text))

        If txtTempatLahir.Text <> "" Then
            .Parameters.Append .CreateParameter("TempatLahir", adVarChar, adParamInput, 25, Trim(txtTempatLahir.Text))
        Else
            .Parameters.Append .CreateParameter("TempatLahir", adVarChar, adParamInput, 25, Null)
        End If
        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(meTglLahir.Text, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 1, Left(cboJnsKelaminPasien.Text, 1))
        If txtAlamat.Text <> "" Then
            .Parameters.Append .CreateParameter("Alamat", adVarChar, adParamInput, 100, Trim(txtAlamat.Text))
        Else
            .Parameters.Append .CreateParameter("Alamat", adVarChar, adParamInput, 100, Null)
        End If
        If txtTelepon.Text <> "" Then
            .Parameters.Append .CreateParameter("Telepon", adVarChar, adParamInput, 15, Trim(txtTelepon.Text))
        Else
            .Parameters.Append .CreateParameter("Telepon", adVarChar, adParamInput, 15, Null)
        End If
        If dcPropinsi.Text <> "" Then
            .Parameters.Append .CreateParameter("Propinsi", adVarChar, adParamInput, 30, Trim(dcPropinsi.Text))
        Else
            .Parameters.Append .CreateParameter("Propinsi", adVarChar, adParamInput, 30, Null)
        End If
        If dcKota.Text <> "" Then
            .Parameters.Append .CreateParameter("Kota", adVarChar, adParamInput, 50, Trim(dcKota.Text))
        Else
            .Parameters.Append .CreateParameter("Kota", adVarChar, adParamInput, 50, Null)
        End If
        If dcKecamatan.Text <> "" Then
            .Parameters.Append .CreateParameter("Kecamatan", adVarChar, adParamInput, 50, Trim(dcKecamatan.Text))
        Else
            .Parameters.Append .CreateParameter("Kecamatan", adVarChar, adParamInput, 50, Null)
        End If
        If dcKelurahan.Text <> "" Then
            .Parameters.Append .CreateParameter("Kelurahan", adVarChar, adParamInput, 50, Trim(dcKelurahan.Text))
        Else
            .Parameters.Append .CreateParameter("Kelurahan", adVarChar, adParamInput, 50, Null)
        End If
        If meRTRW.Text <> "__/__" Then
            .Parameters.Append .CreateParameter("RTRW", adVarChar, adParamInput, 5, meRTRW.Text)
        Else
            .Parameters.Append .CreateParameter("RTRW", adVarChar, adParamInput, 5, Null)
        End If
        If txtKodePos.Text <> "" Then
            .Parameters.Append .CreateParameter("KodePos", adChar, adParamInput, 5, Trim(txtKodePos.Text))
        Else
            .Parameters.Append .CreateParameter("KodePos", adChar, adParamInput, 5, Null)
        End If
        .Parameters.Append .CreateParameter("OutputNoCM", adVarChar, adParamOutput, 12, Null)
'        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "AU_Pasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data Pasien", vbCritical, "Validasi"
        Else
            MsgBox "Data pasien baru berhasil disimpan..", vbInformation, "Informasi"
            ValNoCM = True

            If Trim(.Parameters("OutputNoCM").value) = "CMMAX" Then
                MsgBox "NoCM Tidak dapat melebihi 12 Digit" & vbNewLine _
                & "Hubungi administrator dan laporkan pesan tersebut" & vbNewLine _
                & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
                Exit Sub
            End If
            
            If Not IsNull(.Parameters("OutputNoCM").value) Then mstrNoCM = .Parameters("OutputNoCM").value
            If Len(mstrNoCM) = 0 Then
                strLokal = "SELECT NoCM from Pasien where namaLengkap = '" & Trim(txtNamaPasien.Text) & "' and TglLahir = '" & Format(meTglLahir.Text, "yyyy/MM/dd HH:mm:ss") & "' and TglDaftarMembership = '" & Format(dtpTglPendaftaran.value, "yyyy/MM/dd HH:mm:ss") & "'"
                Call msubRecFO(rs, strLokal)
                txtNoCM.Text = rs("NoCM").value
            Else
                txtNoCM.Text = mstrNoCM
            End If
            Call Add_HistoryLoginActivity("AU_Pasien")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub sp_Add_NoCMTerm(ByVal adoCommand As ADODB.Command)
On Error GoTo errLoad
With adoCommand
    .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
    .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
    .Parameters.Append .CreateParameter("OutputNoCMTerm", adVarChar, adParamOutput, 12)
    .ActiveConnection = dbConn
    .CommandText = "Add_NoCMTerm"
    .CommandType = adCmdStoredProc
    .Execute

If Not (.Parameters("RETURN_VALUE").value = 0) Then
    MsgBox "Ada kesalahan dalam pemasukan data Pasien", vbCritical, "Validasi"
Else
    txtnocmterm.Text = .Parameters("OutputNoCMTerm").value
End If
End With
    Call deleteADOCommandParameters(adoCommand)
    Set adoCommand = Nothing
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

'untuk membersihkan data pasien
Private Sub subClearData()
    On Error Resume Next
    txtNoCM.Text = ""
    cboNamaDepan.ListIndex = -1
    txtNamaPasien.Text = ""
    txtNoIdentitas.Text = ""
    cboJnsKelaminPasien.ListIndex = -1
    txtTempatLahir.Text = ""
    meTglLahir.Text = "__/__/____"
    txtTahun.Text = ""
    txtBulan.Text = ""
    txthari.Text = ""
    txtAlamat.Text = ""
    meRTRW.Text = "__/__"
    txtTelepon.Text = ""

    dcPropinsi.Enabled = True
    dcKota.Enabled = True
    dcKecamatan.Enabled = True
    dcKelurahan.Enabled = True

    dcPropinsi.Text = ""
    dcKota.Text = ""
    dcKecamatan.Text = ""
    dcKelurahan.Text = ""
    txtKodePos.Text = ""
    txtNamaPanggilan.Text = ""
    txtNoKK.Text = ""
    txtKepalaKeluarga.Text = ""
    txtIbuKandung.Text = ""
    txtFormPengirim.Text = Me.Name
    txtNoReservasi.Text = ""
    
    j = 0
    dtpTglPendaftaran.value = Now
    dtpTglPendaftaran.SetFocus
    chkPegawaiToPasien.Enabled = True
End Sub

'untuk enable/disable button reg
Private Sub subEnableButtonReg(blnStatus As Boolean)
    cmdDetailPasien.Enabled = blnStatus
    cmdRegMRS.Enabled = blnStatus
    cmdSimpan.Enabled = Not blnStatus
End Sub

'untuk enable/disable button reg
Private Sub subVisibleButtonReg(blnStatus As Boolean)
    cmdtambah.Visible = blnStatus
    cmdRegMRS.Enabled = blnStatus
End Sub

'untuk load data pasien yg sudah pernah didaftarkan
Public Sub subLoadDataPasien(strInput As String)
    On Error GoTo errLoad

    Dim strSQLLoadPasien As String
    Dim rsLoadPasien As New ADODB.recordset
    strSQLLoadPasien = "SELECT * FROM Pasien WHERE NoCM = '" & strInput & "'"
    Set rsLoadPasien = Nothing
    rsLoadPasien.Open strSQLLoadPasien, dbConn, adOpenForwardOnly, adLockReadOnly
    If rsLoadPasien.EOF = True Then Exit Sub

    txtNoCM.Text = mstrNoCM
    dtpTglPendaftaran.MaxDate = Now
'    dtpTglPendaftaran.value = rsLoadPasien.Fields("TglDaftarMembership").value
    cboNamaDepan.Text = rsLoadPasien.Fields("Title").value
    txtNamaPasien.Text = rsLoadPasien.Fields("NamaLengkap").value
    If Not IsNull(rsLoadPasien.Fields("NoIdentitas").value) Then txtNoIdentitas.Text = rsLoadPasien.Fields("NoIdentitas").value
    If rsLoadPasien.Fields("JenisKelamin").value = "L" Then
        cboJnsKelaminPasien.ListIndex = 0
    ElseIf rsLoadPasien.Fields("JenisKelamin").value = "P" Then
        cboJnsKelaminPasien.ListIndex = 1
    End If
    If Not IsNull(rsLoadPasien.Fields("TempatLahir").value) Then txtTempatLahir.Text = rsLoadPasien.Fields("TempatLahir").value
    meTglLahir.Text = Format(rsLoadPasien.Fields("TglLahir").value, "dd/MM/yyyy")
    If Not IsNull(rsLoadPasien.Fields("Alamat").value) Then txtAlamat.Text = rsLoadPasien.Fields("Alamat").value
    If Not IsNull(rsLoadPasien.Fields("RTRW").value) Then
        If Len(rsLoadPasien.Fields("RTRW").value) = 5 And InStr(1, rsLoadPasien.Fields("RTRW").value, "/") = 3 Then
            meRTRW.Text = rsLoadPasien.Fields("RTRW").value
        Else
            If InStr(1, rsLoadPasien.Fields("RTRW").value, "/") = 0 Then
                meRTRW.Text = Format(Left(rsLoadPasien.Fields("RTRW").value, Len(rsLoadPasien.Fields("RTRW").value) / 2), "00") & "/" & Format(Right(rsLoadPasien.Fields("RTRW").value, Len(rsLoadPasien.Fields("RTRW").value) / 2), "00")
            Else
                meRTRW.Text = Format(Left(rsLoadPasien.Fields("RTRW").value, InStr(1, rsLoadPasien.Fields("RTRW").value, "/") - 1), "00") & "/" & Format(Right(rsLoadPasien.Fields("RTRW").value, Len(rsLoadPasien.Fields("RTRW").value) - InStr(1, rsLoadPasien.Fields("RTRW").value, "/")), "00")
            End If
        End If
    End If
    
    
    If Not IsNull(rsLoadPasien.Fields("namapanggilan").value) Then txtNamaPanggilan.Text = rsLoadPasien.Fields("namapanggilan").value
    
    
    strSQL2 = "Select * from DetailPasien where NoCM = '" & strInput & "'"
    Call msubRecFO(rs2, strSQL2)
    If (Not rs2.EOF) Then
        If Not IsNull(rs2.Fields("NoKK").value) Then txtNoKK.Text = rs2.Fields("NoKK").value
        If Not IsNull(rs2.Fields("NamaKepalaKeluarga").value) Then txtKepalaKeluarga.Text = rs2.Fields("NamaKepalaKeluarga").value
        If Not IsNull(rs2.Fields("NamaIbu").value) Then txtIbuKandung.Text = rs2.Fields("NamaIbu").value
    End If
    'KINGKONG
    If Not IsNull(rsLoadPasien.Fields("Telepon").value) Then txtTelepon.Text = rsLoadPasien.Fields("Telepon").value
    If Not IsNull(rsLoadPasien.Fields("Propinsi").value) Then dcPropinsi.Text = rsLoadPasien.Fields("Propinsi").value
    Call subLoadDataWilayah("propinsi")
    
    If Not IsNull(rsLoadPasien.Fields("Kota").value) Then dcKota.Text = rsLoadPasien.Fields("Kota").value
     Call subLoadDataWilayah("kota")
    If Not IsNull(rsLoadPasien.Fields("Kecamatan").value) Then dcKecamatan.Text = rsLoadPasien.Fields("Kecamatan").value
     Call subLoadDataWilayah("kecamatan")
    If Not IsNull(rsLoadPasien.Fields("Kelurahan").value) Then dcKelurahan.Text = rsLoadPasien.Fields("Kelurahan").value
     Call subLoadDataWilayah("desa")
    If Not IsNull(rsLoadPasien.Fields("KodePos").value) Then txtKodePos.Text = rsLoadPasien.Fields("KodePos").value
  
  
  If bolDetailPasien = True Then
     strSQL = "Select * from Pasien where NoCM = '" & strInput & "'"
    Call msubRecFO(rs, strSQL)
    If Not IsNull(rs.Fields("Propinsi").value) Then dcPropinsi.Text = rs.Fields("Propinsi").value
    If Not IsNull(rsLoadPasien.Fields("Kota").value) Then dcKota.Text = rsLoadPasien.Fields("Kota").value
    If Not IsNull(rsLoadPasien.Fields("Kecamatan").value) Then dcKecamatan.Text = rsLoadPasien.Fields("Kecamatan").value
    If Not IsNull(rsLoadPasien.Fields("Kelurahan").value) Then dcKelurahan.Text = rsLoadPasien.Fields("Kelurahan").value
  Else
      CekPilihanWilayah2 "dcPropinsi", "Click"
      CekPilihanWilayah2 "dcKota", "Click"
      CekPilihanWilayah2 "dcKecamatan", "Click"
      CekPilihanWilayah2 "dcKelurahan", "Click"
  
  End If
    
    Call meTglLahir_KeyPress(13)
    Set rsLoadPasien = Nothing

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDataWilayah(strPencarian As String)
    'On Error GoTo errLoad
    On Error Resume Next
    Dim strTempSql As String

    Select Case strPencarian
        Case "propinsi"
            If Len(Trim(dcPropinsi.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaPropinsi LIKE '%" & dcPropinsi.Text & "%')and statusenabled=1"

        Case "kota"
            If Len(Trim(dcKota.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaPropinsi LIKE '%" & dcPropinsi.Text & "%') and (NamaKotaKabupaten LIKE '%" & dcKota.Text & "%')"

        Case "kecamatan"
            If Len(Trim(dcKecamatan.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaPropinsi LIKE '%" & dcPropinsi.Text & "%') and (NamaKotaKabupaten LIKE '%" & dcKota.Text & "%') and (NamaKecamatan LIKE '%" & dcKecamatan.Text & "%')"
        Case "desa"
            If Len(Trim(dcKelurahan.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaPropinsi LIKE '%" & dcPropinsi.Text & "%') and (NamaKotaKabupaten LIKE '%" & dcKota.Text & "%') and (NamaKecamatan LIKE '%" & dcKecamatan.Text & "%') and (NamaKelurahan LIKE '%" & dcKelurahan.Text & "%')"

        Case "kodepos"
            If Len(Trim(txtKodePos.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaPropinsi LIKE '%" & dcPropinsi.Text & "%') and (NamaKotaKabupaten LIKE '%" & dcKota.Text & "%') and (NamaKecamatan LIKE '%" & dcKecamatan.Text & "%') and (NamaKelurahan LIKE '%" & dcKelurahan.Text & "%') and (KodePos LIKE '%" & txtKodePos.Text & "%')"

    End Select

    strSQL = "SELECT DISTINCT ISNULL(NamaPropinsi, '') AS NamaPropinsi, ISNULL(NamaKotaKabupaten, '') AS NamaKotaKabupaten, ISNULL(NamaKecamatan, '')  AS NamaKecamatan, ISNULL(NamaKelurahan, '') AS NamaKelurahan, ISNULL(KodePos, '') AS KodePos" & _
    " FROM V_Wilayah" & _
    " " & strTempSql

    Call msubRecFO(rs, strSQL)
    If rs.EOF Then
'       MsgBox "Data Wilayah Tidak Sesuai, Harap Cek Data Wilayah", vbInformation, "Validasi"
       MsgBox "Data Kodepos Tidak Sesuai", vbInformation, "Validasi"
'        dcPropinsi.BoundText = ""
'        dcKota.BoundText = ""
'        dcKecamatan.BoundText = ""
'        dcKelurahan.BoundText = ""
        txtKodePos.Text = ""
    End If
    If j = 1 Then
        If rs(1).value = "" Then
            MsgBox "Data Kota/Kabupaten Belum Ada", vbInformation, "Validasi"
'            dcKota.Enabled = False
'            dcKecamatan.Enabled = False
'            dcKelurahan.Enabled = False
        Else

        End If

    ElseIf j = 2 Then
        If rs(2).value = "" Then
            MsgBox "Data Kecamatan Belum Ada", vbInformation, "Validasi"
'            dcKecamatan.Enabled = False
'            dcKelurahan.Enabled = False
        Else

        End If

    ElseIf j = 3 Then
        If rs(3).value = "" Then
            MsgBox "Data Kelurahan Belum Ada", vbInformation, "Validasi"
'            dcKelurahan.Enabled = False
        Else

        End If

    Else
        dcPropinsi.Text = rs("NamaPropinsi")
        dcKota.Text = rs("NamaKotaKabupaten")
        dcKecamatan.Text = rs("NamaKecamatan")
        dcKelurahan.Text = rs("NamaKelurahan")
        txtKodePos.Text = rs("KodePos")
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subDcSource(varstrPilihan As String, Optional varStrSQL As String)
    Select Case varstrPilihan

        Case "Propinsi"
            strSQL = "SELECT DISTINCT KdPropinsi, NamaPropinsi AS alias FROM V_Wilayah where StatusEnabled=1 order by NamaPropinsi"
            Set rsPropinsi = Nothing
            rsPropinsi.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dcPropinsi.RowSource = rsPropinsi
            dcPropinsi.BoundColumn = rsPropinsi(0).Name
            dcPropinsi.ListField = rsPropinsi(1).Name
        Case "Kota"
            strSQL = "SELECT DISTINCT KdKotaKabupaten, NamaKotaKabupaten AS alias FROM V_Wilayah " & varStrSQL & ""
            Set rsKota = Nothing
            rsKota.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dcKota.RowSource = rsKota
            dcKota.BoundColumn = rsKota(0).Name
            dcKota.ListField = rsKota(1).Name
        Case "Kecamatan"
            strSQL = "SELECT DISTINCT KdKecamatan, NamaKecamatan AS alias FROM V_Wilayah " & varStrSQL & ""
            Set rsKecamatan = Nothing
            rsKecamatan.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dcKecamatan.RowSource = rsKecamatan
            dcKecamatan.BoundColumn = rsKecamatan(0).Name
            dcKecamatan.ListField = rsKecamatan(1).Name
        Case "Kelurahan"
            strSQL = "SELECT DISTINCT KdKelurahan, NamaKelurahan AS alias FROM V_Wilayah " & varStrSQL & ""
            Set rsKelurahan = Nothing
            rsKelurahan.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dcKelurahan.RowSource = rsKelurahan
            dcKelurahan.BoundColumn = rsKelurahan(0).Name
            dcKelurahan.ListField = rsKelurahan(1).Name
    End Select

    Exit Sub
End Sub

'@Dimas 2014-05-10
Private Sub subPrintRegistrasiBarcode()

On Error GoTo errLoad
Dim PosAwal, PosTamb, Hal As Double
Dim mstrNoCMBar As String
Dim tmpXY As String
Dim z As String
    
'    If dgpasien.ApproxCount = 0 Then Exit Sub
    Call msubRecFO(rs, "SELECT NamaPrinterBarcode FROM MasterDataPendukung")
    
    If IsNull(rs("NamaPrinterBarcode")) Then
        MsgBox "Nama printer barcode kosong", vbExclamation, "Informasi"
        Exit Sub
    End If
    
    
    cbojnsPrinter.Clear
    z = "Zebra P330i Card Printer USB"
    For Each subPrinterZebra In Printers
        cbojnsPrinter.AddItem subPrinterZebra.DeviceName
        If Right(subPrinterZebra.DeviceName, Len(rs("NamaPrinterBarcode"))) = rs("NamaPrinterBarcode") Then
            X = rs("NamaPrinterBarcode")
            Exit For
        ElseIf Right(subPrinterZebra.DeviceName, Len(rs("NamaPrinterBarcode"))) = z Then
            X = z
            Exit For
        End If
    Next
    
    If X = "" Then Exit Sub
    
    tmpXY = X
    
    If subPrinterZebra.DeviceName = X Then
        Set Printer = subPrinterZebra
    ElseIf subPrinterZebra.DeviceName = tmpXY Then
        Set Printer = subPrinterZebra
    Else
        MsgBox "Printer barcode tidak terdeteksi, harap periksa lagi", vbInformation, "Informasi"
        Exit Sub
    End If
    
    Set Barcode39 = New clsBarCode39
    PosAwal = 100 'pos awal ???
    PosTamb = 0
    Hal = 1
    
    Printer.CurrentY = PosTamb
    
    'print nama rs
    Printer.Print ""
    Printer.CurrentX = 100
    Printer.FontName = "Tahoma"
    Printer.Font.Size = 12
    Printer.Font.Bold = True
    Printer.Print ""
    
    'print jalan rs
    Printer.CurrentX = 100
    Printer.FontName = "Tahoma"
    Printer.Font.Size = 9
    Printer.Font.Bold = False
    Printer.Print ""
    
    'print telp
    Printer.CurrentX = 100
    Printer.FontName = "Tahoma"
    Printer.Font.Size = 12
    Printer.Font.Bold = False
    Printer.Print ""
    
    Printer.CurrentX = 100
    Printer.FontName = "Tahoma"
    Printer.Font.Size = 10
    Printer.Font.Bold = False
    Printer.Print ""
    
'    Printer.CurrentX = 100
'    Printer.FontName = "Tahoma"
'    Printer.Font.Size = 9
'    Printer.Font.Bold = False
'    Printer.Print ""
    
'    Printer.CurrentX = 100
'    Printer.FontName = "Tahoma"
'    Printer.Font.Size = 9
'    Printer.Font.Bold = False
'    Printer.Print ""
'
'     Printer.Print ""
'    mstrNoCMBar = Left(dgpasien.Columns("No. CM"), 2) & "-" & Mid(dgpasien.Columns("No. CM"), 3, 2) & "-" & Right(dgpasien.Columns("No. CM"), 2)
'    mstrNoCMBar = dgpasien.Columns("No. CM").value
    Printer.CurrentX = 550
    Printer.Font.Name = "Tahoma"
    Printer.Font.Size = 14
    Printer.FontBold = True
    Printer.Print mstrNoCM
    
   
    
    Printer.CurrentX = 300
    Printer.Font.Name = "Tahoma"
    Printer.Font.Size = 12
    Printer.FontBold = False
    PosTamb = Printer.CurrentY
    
    With Barcode39
        .CurrentX = 275  '400 - 150
        .CurrentY = 1450 'sip ' jarak barcode dari atas ke bawah makin dikit makin ke atas
        
        .NarrowX = 12 'Val(txtNarrowX.Text)
        .BarcodeHeight = 400 'Val(txtHeight.Text)
        .ShowBox = 0
        .Barcode = mstrNoCM
        If .ErrNumber <> 0 Then
            MsgBox "Error: It contain invalid barcode charater", vbOKOnly + vbCritical, "Error"
            Exit Sub
        End If
        .Draw Printer
        
    End With
     
    Printer.CurrentX = 500
    Printer.Font.Name = "Tahoma"
    Printer.Font.Size = 5
    
     Printer.Print ""
    'print NamaPasien
    Printer.CurrentX = 300
    Printer.Font.Name = "Tahoma"
    Printer.Font.Size = 8
    Printer.Font.Bold = True
    Printer.Print UCase(txtNamaPasien.Text)
    
    'print Alamat Pasien
    Printer.CurrentX = 300
    Printer.Font.Name = "Tahoma"
    Printer.Font.Size = 5
    'Printer.Font.Bold = True
    Printer.Print txtAlamat.Text & " " & "KEC. " & dcKecamatan.Text
'    Printer.Print dgpasien.Columns("Kelurahan").value & " " & "RT/RW : " & dgpasien.Columns("RT/RW").value & " " & "KEC. " & dgpasien.Columns("Kecamatan").value
    
    'print Kota Pasien
    Printer.CurrentX = 300
    Printer.Font.Name = "Tahoma"
    Printer.Font.Size = 5
    'Printer.Font.Bold = True
    Printer.Print dcKota.Text

    Printer.EndDoc
Exit Sub
errLoad:
    Call msubPesanError
    Printer.KillDoc
End Sub



