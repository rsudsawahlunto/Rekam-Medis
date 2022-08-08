VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmMasterUmum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Sosial Pasien"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMasterData_Umum.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   8415
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   120
      TabIndex        =   87
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   7200
      Width           =   1575
   End
   Begin TabDlg.SSTab sstMasterUmum 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   7
      Tab             =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Agama"
      TabPicture(0)   =   "frmMasterData_Umum.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Suku"
      TabPicture(1)   =   "frmMasterData_Umum.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Pekerjaan"
      TabPicture(2)   =   "frmMasterData_Umum.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Pendidikan"
      TabPicture(3)   =   "frmMasterData_Umum.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Kelompok Umur"
      TabPicture(4)   =   "frmMasterData_Umum.frx":0D3A
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "dgKelompokUmur"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "txtKeterangan"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "txtKelompokUmur"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "txtKdKelompokUmur"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "txtNamaExternal4"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "txtKodeExternal4"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "CheckStatusEnbl4"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "Hubungan Keluarga"
      TabPicture(5)   =   "frmMasterData_Umum.frx":0D56
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame3"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Golongan Darah"
      TabPicture(6)   =   "frmMasterData_Umum.frx":0D72
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame4"
      Tab(6).ControlCount=   1
      Begin VB.CheckBox CheckStatusEnbl4 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         Height          =   255
         Left            =   6600
         TabIndex        =   35
         Top             =   3240
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtKodeExternal4 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   33
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtNamaExternal4 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   34
         Top             =   3120
         Width           =   4335
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   -74880
         TabIndex        =   64
         Top             =   720
         Width           =   8055
         Begin VB.CheckBox CheckStatusEnbl6 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   6600
            TabIndex        =   47
            Top             =   1560
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtKodeExternal6 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   45
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtNamaExternal6 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   46
            Top             =   1440
            Width           =   4695
         End
         Begin VB.TextBox txtKdGolonganDarah 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   43
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtGolonganDarah 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   44
            Top             =   720
            Width           =   615
         End
         Begin MSDataGridLib.DataGrid dgGolonganDarah 
            Height          =   3135
            Left            =   120
            TabIndex        =   48
            Top             =   1920
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   5530
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
         Begin VB.Label Label30 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   86
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label29 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   85
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Kode Gol. Darah"
            Height          =   210
            Left            =   240
            TabIndex        =   66
            Top             =   360
            Width           =   1320
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Gol. Darah"
            Height          =   210
            Left            =   240
            TabIndex        =   65
            Top             =   720
            Width           =   840
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   -74880
         TabIndex        =   61
         Top             =   720
         Width           =   8055
         Begin VB.TextBox txtNamaExternal5 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   40
            Top             =   1440
            Width           =   4695
         End
         Begin VB.TextBox txtKodeExternal5 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   39
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CheckBox CheckStatusEnbl5 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   6600
            TabIndex        =   41
            Top             =   1560
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtKdHubungan 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   37
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtHubungan 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   38
            Top             =   720
            Width           =   3255
         End
         Begin MSDataGridLib.DataGrid dgHubunganKeluarga 
            Height          =   3135
            Left            =   120
            TabIndex        =   42
            Top             =   1920
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   5530
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
         Begin VB.Label Label28 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   78
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label27 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Kode Hubungan "
            Height          =   210
            Left            =   240
            TabIndex        =   63
            Top             =   360
            Width           =   1380
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Hubungan"
            Height          =   210
            Left            =   240
            TabIndex        =   62
            Top             =   720
            Width           =   840
         End
      End
      Begin VB.TextBox txtKdKelompokUmur 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   30
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtKelompokUmur 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   31
         Top             =   1320
         Width           =   4335
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   32
         Top             =   2400
         Width           =   4335
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   -74880
         TabIndex        =   58
         Top             =   720
         Width           =   8055
         Begin VB.TextBox txtNamaExternal 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   4
            Top             =   1440
            Width           =   4695
         End
         Begin VB.TextBox txtKodeExternal 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   3
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CheckBox CheckStatusEnbl 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   6600
            TabIndex        =   5
            Top             =   1560
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtAgama 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   2
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtKdAgama 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   1
            Top             =   360
            Width           =   615
         End
         Begin MSDataGridLib.DataGrid dgAgama 
            Height          =   3135
            Left            =   120
            TabIndex        =   6
            Top             =   1920
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   5530
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
         Begin VB.Label Label18 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Agama"
            Height          =   210
            Left            =   240
            TabIndex        =   60
            Top             =   720
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
            Height          =   210
            Left            =   240
            TabIndex        =   59
            Top             =   360
            Width           =   420
         End
      End
      Begin VB.Frame Frame7 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   -74880
         TabIndex        =   55
         Top             =   720
         Width           =   8055
         Begin VB.CheckBox CheckStatusEnbl1 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   6600
            TabIndex        =   15
            Top             =   1560
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtKodeExternal1 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   13
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtNamaExternal1 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   14
            Top             =   1440
            Width           =   4695
         End
         Begin VB.TextBox txtSuku 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   12
            Top             =   720
            Width           =   3615
         End
         Begin VB.TextBox txtKdSuku 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   11
            Top             =   360
            Width           =   615
         End
         Begin MSDataGridLib.DataGrid dgSuku 
            Height          =   3135
            Left            =   120
            TabIndex        =   16
            Top             =   1920
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   5530
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
         Begin VB.Label Label20 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   72
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label19 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Suku"
            Height          =   210
            Left            =   240
            TabIndex        =   57
            Top             =   720
            Width           =   405
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Kode Suku"
            Height          =   210
            Left            =   240
            TabIndex        =   56
            Top             =   360
            Width           =   885
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   -74880
         TabIndex        =   52
         Top             =   720
         Width           =   8055
         Begin VB.TextBox txtNamaExternal2 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   20
            Top             =   1440
            Width           =   4695
         End
         Begin VB.TextBox txtKodeExternal2 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   19
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CheckBox CheckStatusEnbl2 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   6600
            TabIndex        =   21
            Top             =   1560
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtPekerjaan 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   18
            Top             =   720
            Width           =   4215
         End
         Begin VB.TextBox txtKdPekerjaan 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   17
            Top             =   360
            Width           =   615
         End
         Begin MSDataGridLib.DataGrid dgPekerjaan 
            Height          =   3135
            Left            =   120
            TabIndex        =   22
            Top             =   1920
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   5530
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
         Begin VB.Label Label22 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   74
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pekerjaan"
            Height          =   210
            Left            =   240
            TabIndex        =   54
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Kode Pekerjaan"
            Height          =   210
            Left            =   240
            TabIndex        =   53
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   -74880
         TabIndex        =   49
         Top             =   720
         Width           =   8055
         Begin VB.CheckBox CheckStatusEnbl3 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   6600
            TabIndex        =   28
            Top             =   1560
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtKodeExternal3 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   26
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtNamaExternal3 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   27
            Top             =   1440
            Width           =   4695
         End
         Begin VB.TextBox txtNoUrut 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   6480
            MaxLength       =   2
            TabIndex        =   25
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtKdPendidikan 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   23
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtPendidikan 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   25
            TabIndex        =   24
            Top             =   720
            Width           =   3495
         End
         Begin MSDataGridLib.DataGrid dgPendidikan 
            Height          =   3135
            Left            =   120
            TabIndex        =   29
            Top             =   1920
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   5530
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
         Begin VB.Label Label24 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   76
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "No. Urut"
            Height          =   210
            Left            =   5640
            TabIndex        =   68
            Top             =   840
            Width           =   705
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Kode Pendidikan"
            Height          =   210
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   1350
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Pendidikan"
            Height          =   210
            Left            =   240
            TabIndex        =   50
            Top             =   720
            Width           =   870
         End
      End
      Begin MSDataGridLib.DataGrid dgKelompokUmur 
         Height          =   2175
         Left            =   240
         TabIndex        =   36
         Top             =   3600
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   3836
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
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   120
         TabIndex        =   79
         Top             =   720
         Width           =   8055
         Begin VB.TextBox txtUmurMax 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   5040
            MaxLength       =   20
            TabIndex        =   93
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtUmurMin 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2040
            MaxLength       =   20
            TabIndex        =   92
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtRangeUmur 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   89
            Top             =   960
            Width           =   4335
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Umur Min."
            Height          =   210
            Left            =   120
            TabIndex        =   91
            Top             =   1320
            Width           =   825
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Umur Max."
            Height          =   210
            Left            =   4080
            TabIndex        =   90
            Top             =   1320
            Width           =   870
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Range Umur"
            Height          =   210
            Left            =   120
            TabIndex        =   88
            Top             =   960
            Width           =   1005
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Keterangan"
            Height          =   210
            Left            =   120
            TabIndex        =   84
            Top             =   1680
            Width           =   945
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Kelompok Umur"
            Height          =   210
            Left            =   120
            TabIndex        =   83
            Top             =   600
            Width           =   1290
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Kode Kelompok Umur"
            Height          =   210
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Width           =   1770
         End
         Begin VB.Label Label25 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label26 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   2040
            Width           =   1215
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   67
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
      Left            =   6600
      Picture         =   "frmMasterData_Umum.frx":0D8E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterData_Umum.frx":1B16
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterData_Umum.frx":3174
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmMasterUmum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoCommand As New ADODB.Command

Private Sub cmdBatal_Click()
    On Error GoTo errLoad

    Call clear
    Call subLoadGridSource
    Call sstMasterUmum_KeyPress(13)

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    Select Case sstMasterUmum.Tab
        Case 0
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmagama.Show
        Case 1
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmsuku.Show
        Case 2
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmpekerjaan.Show
        Case 3
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmpendidikan.Show
        Case 4
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmkelompokumur.Show
        Case 5
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmhubungankeluarga.Show
        Case 6
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmgolongandarah.Show
    End Select
hell:
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errLoad
    If MsgBox("Yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Select Case sstMasterUmum.Tab
        Case 0 'Agama
            If Periksa("text", txtAgama, "Silahkan isi Nama Agama") = False Then Exit Sub
            If sp_Agama("D") = False Then Exit Sub
        Case 1 'Suku
            If Periksa("text", txtSuku, "Silahkan isi Nama Suku") = False Then Exit Sub
            If sp_Suku("D") = False Then Exit Sub
        Case 2 'Pekerjaan
            If Periksa("text", txtPekerjaan, "Silahkan isi Nama Pekerjaan") = False Then Exit Sub
            If sp_Pekerjaan("D") = False Then Exit Sub
        Case 3 'Pendidikan
            If Periksa("text", txtPendidikan, "Silahkan isi Nama Pendidikan") = False Then Exit Sub
            If sp_Pendidikan("D") = False Then Exit Sub
        Case 4 'Kelompok Umur
            If Periksa("text", txtKelompokUmur, "Silahkan isi Nama Kelompok") = False Then Exit Sub
            If sp_KelompokUmur("D") = False Then Exit Sub
        Case 5 'Hubungan Keluarga
            If Periksa("text", txtHubungan, "Silahkan isi Nama Hubungan") = False Then Exit Sub
            If sp_HubunganKeluarga("D") = False Then Exit Sub
        Case 6 'Golongan Darah
            If Periksa("text", txtGolonganDarah, "Silahkan isi Golongan") = False Then Exit Sub
            If sp_GolonganDarah("D") = False Then Exit Sub
    End Select

    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function sp_Agama(f_Status As String) As Boolean
    sp_Agama = True

    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdAgama", adChar, adParamInput, 2, txtKdAgama.Text)
        .Parameters.Append .CreateParameter("Agama", adVarChar, adParamInput, 20, Trim(txtAgama.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 20, txtNamaExternal.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_Agama"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data Agama", vbCritical, "Validasi"
            sp_Agama = False
        End If

        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_Suku(f_Status As String) As Boolean
    sp_Suku = True

    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdSuku", adChar, adParamInput, 2, txtKdSuku.Text)
        .Parameters.Append .CreateParameter("Suku", adVarChar, adParamInput, 20, Trim(txtSuku.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal1.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 20, txtNamaExternal1.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl1.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_Suku"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data Suku", vbCritical, "Validasi"
            sp_Suku = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_Pekerjaan(f_Status As String) As Boolean
    sp_Pekerjaan = True

    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdPekerjaan", adChar, adParamInput, 2, txtKdPekerjaan.Text)
        .Parameters.Append .CreateParameter("Pekerjaan", adVarChar, adParamInput, 30, Trim(txtPekerjaan.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal2.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, txtNamaExternal2.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl2.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_Pekerjaan"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Pekerjaan", vbCritical, "Validasi"
            sp_Pekerjaan = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_Pendidikan(f_Status As String) As Boolean
    sp_Pendidikan = True

    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdPendidikan", adChar, adParamInput, 2, txtKdPendidikan.Text)
        .Parameters.Append .CreateParameter("Pendidikan", adVarChar, adParamInput, 25, Trim(txtPendidikan.Text))
        .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, IIf(txtNoUrut.Text = "", Null, txtNoUrut.Text))
        .Parameters.Append .CreateParameter("KdJenisPendidikan", adVarChar, adParamInput, 3, Null)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal3.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 25, txtNamaExternal3.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl3.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_Pendidikan"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Pendidikan", vbCritical, "Validasi"
            sp_Pendidikan = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_KelompokUmur(f_Status As String) As Boolean
    sp_KelompokUmur = True

    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKelompokUmur", adChar, adParamInput, 3, txtKdKelompokUmur.Text)
        .Parameters.Append .CreateParameter("KelompokUmur", adVarChar, adParamInput, 30, Trim(txtKelompokUmur.Text))
        .Parameters.Append .CreateParameter("RangeUmur", adVarChar, adParamInput, 30, IIf(txtRangeUmur.Text = "", Null, txtRangeUmur.Text))
        .Parameters.Append .CreateParameter("UmurMin", adInteger, adParamInput, , IIf(txtUmurMin.Text = "", Null, txtUmurMin.Text))
        .Parameters.Append .CreateParameter("UmurMax", adInteger, adParamInput, , IIf(txtUmurMax.Text = "", Null, txtUmurMax.Text))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 50, IIf(txtketerangan.Text = "", Null, Trim(txtketerangan.Text)))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, IIf(txtKodeExternal4.Text = "", Null, txtKodeExternal4.Text))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, IIf(txtNamaExternal4.Text = "", Null, txtNamaExternal4.Text))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl4.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_KelompokUmur"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Kelompok Umur", vbCritical, "Validasi"
            sp_KelompokUmur = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_HubunganKeluarga(f_Status As String) As Boolean
    sp_HubunganKeluarga = True

    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdHubungan", adChar, adParamInput, 2, txtKdHubungan.Text)
        .Parameters.Append .CreateParameter("NamaHubungan", adVarChar, adParamInput, 50, Trim(txtHubungan.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal5.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNamaExternal5.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl5.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_HubunganKeluarga"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Hubungan Keluarga", vbCritical, "Validasi"
            sp_HubunganKeluarga = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_GolonganDarah(f_Status As String) As Boolean
    sp_GolonganDarah = True

    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdGolonganDarah", adChar, adParamInput, 2, txtKdGolonganDarah.Text)
        .Parameters.Append .CreateParameter("GolonganDarah", adVarChar, adParamInput, 2, Trim(txtGolonganDarah.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal6.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 2, txtNamaExternal6.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl6.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_GolonganDarah"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Golongan Darah", vbCritical, "Validasi"
            sp_GolonganDarah = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    Select Case sstMasterUmum.Tab
        Case 0 'Agama
            If Periksa("text", txtAgama, "Silahkan isi Nama Agama") = False Then Exit Sub
            If sp_Agama("A") = False Then Exit Sub
        Case 1 'Suku
            If Periksa("text", txtSuku, "Silahkan isi Nama Suku") = False Then Exit Sub
            If sp_Suku("A") = False Then Exit Sub
        Case 2 'Pekerjaan
            If Periksa("text", txtPekerjaan, "Silahkan isi Nama Pekerjaan") = False Then Exit Sub
            If sp_Pekerjaan("A") = False Then Exit Sub
        Case 3 'Pendidikan
            If Periksa("text", txtPendidikan, "Silahkan isi Nama Pendidikan") = False Then Exit Sub
            If sp_Pendidikan("A") = False Then Exit Sub
        Case 4 'Kelompok Umur
            If Periksa("text", txtKelompokUmur, "Silahkan isi Nama Kelompok") = False Then Exit Sub
            If sp_KelompokUmur("A") = False Then Exit Sub
        Case 5 'Hubungan Keluarga
            If Periksa("text", txtHubungan, "Silahkan isi Nama Hubungan") = False Then Exit Sub
            If sp_HubunganKeluarga("A") = False Then Exit Sub
        Case 6 'Golongan Darah
            If Periksa("text", txtGolonganDarah, "Silahkan isi Golongan") = False Then Exit Sub
            If sp_GolonganDarah("A") = False Then Exit Sub
    End Select

    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgAgama_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgAgama
    WheelHook.WheelHook dgAgama
End Sub

Private Sub dgAgama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdHapus.SetFocus
    On Error Resume Next
    If dgAgama.ApproxCount = 0 Then Exit Sub
    txtKdAgama.Text = dgAgama.Columns(0).value
    txtAgama.Text = dgAgama.Columns(1).value
    txtKodeExternal.Text = dgAgama.Columns(2).value
    txtNamaExternal.Text = dgAgama.Columns(3).value
    If dgAgama.Columns(4) = "" Then
        CheckStatusEnbl.value = 0
    ElseIf dgAgama.Columns(4) = 0 Then
        CheckStatusEnbl.value = 0
    ElseIf dgAgama.Columns(4) = 1 Then
        CheckStatusEnbl.value = 1
    End If
    txtKdAgama.Enabled = False
End Sub

Private Sub dgAgama_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgAgama.ApproxCount = 0 Then Exit Sub
    txtKdAgama.Text = dgAgama.Columns(0).value
    txtAgama.Text = dgAgama.Columns(1).value
    txtKodeExternal.Text = dgAgama.Columns(2).value
    txtNamaExternal.Text = dgAgama.Columns(3).value
    If dgAgama.Columns(4) = "" Then
        CheckStatusEnbl.value = 0
    ElseIf dgAgama.Columns(4) = 0 Then
        CheckStatusEnbl.value = 0
    ElseIf dgAgama.Columns(4) = 1 Then
        CheckStatusEnbl.value = 1
    End If
    txtKdAgama.Enabled = False
End Sub

Private Sub dgGolonganDarah_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgGolonganDarah
    WheelHook.WheelHook dgGolonganDarah
End Sub

Private Sub dgGolonganDarah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdHapus.SetFocus
    On Error Resume Next
    If dgGolonganDarah.ApproxCount = 0 Then Exit Sub
    txtKdGolonganDarah.Text = dgGolonganDarah.Columns(0).value
    txtGolonganDarah.Text = dgGolonganDarah.Columns(1).value
    txtKodeExternal6.Text = dgGolonganDarah.Columns(2).value
    txtNamaExternal6.Text = dgGolonganDarah.Columns(3).value
    If dgGolonganDarah.Columns(4) = "" Then
        CheckStatusEnbl6.value = 0
    ElseIf dgGolonganDarah.Columns(4) = 0 Then
        CheckStatusEnbl6.value = 0
    ElseIf dgGolonganDarah.Columns(4) = 1 Then
        CheckStatusEnbl6.value = 1
    End If
End Sub

Private Sub dgHubunganKeluarga_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgHubunganKeluarga
    WheelHook.WheelHook dgHubunganKeluarga
End Sub

Private Sub dgHubunganKeluarga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdHapus.SetFocus
    On Error Resume Next
    If dgHubunganKeluarga.ApproxCount = 0 Then Exit Sub
    txtKdHubungan.Text = dgHubunganKeluarga.Columns(0).value
    txtHubungan.Text = dgHubunganKeluarga.Columns(1).value
    txtKodeExternal5.Text = dgHubunganKeluarga.Columns(2).value
    txtNamaExternal5.Text = dgHubunganKeluarga.Columns(3).value
    If dgHubunganKeluarga.Columns(4) = "" Then
        CheckStatusEnbl5.value = 0
    ElseIf dgHubunganKeluarga.Columns(4) = 0 Then
        CheckStatusEnbl5.value = 0
    ElseIf dgHubunganKeluarga.Columns(4) = 1 Then
        CheckStatusEnbl5.value = 1
    End If
End Sub

Private Sub dgKelompokUmur_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKelompokUmur
    WheelHook.WheelHook dgKelompokUmur
End Sub

Private Sub dgKelompokUmur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdHapus.SetFocus
    On Error Resume Next
    If dgKelompokUmur.ApproxCount = 0 Then Exit Sub
    txtKdKelompokUmur.Text = dgKelompokUmur.Columns(0).value
    txtKelompokUmur.Text = dgKelompokUmur.Columns(1).value
    txtKodeExternal4.Text = dgKelompokUmur.Columns(3).value
    txtNamaExternal4.Text = dgKelompokUmur.Columns(4).value
    If dgKelompokUmur.Columns(5) = "" Then
        CheckStatusEnbl4.value = 0
    ElseIf dgKelompokUmur.Columns(5) = 0 Then
        CheckStatusEnbl4.value = 0
    ElseIf dgKelompokUmur.Columns(5) = 1 Then
        CheckStatusEnbl4.value = 1
    End If
End Sub

Private Sub dgPekerjaan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPekerjaan
    WheelHook.WheelHook dgPekerjaan
End Sub

Private Sub dgPekerjaan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdHapus.SetFocus
    On Error Resume Next
    If dgPekerjaan.ApproxCount = 0 Then Exit Sub
    txtKdPekerjaan.Text = dgPekerjaan.Columns(0).value
    txtPekerjaan.Text = dgPekerjaan.Columns(1).value
    txtKodeExternal2.Text = dgPekerjaan.Columns(3).value
    txtNamaExternal2.Text = dgPekerjaan.Columns(4).value
    If dgPekerjaan.Columns(5) = "" Then
        CheckStatusEnbl2.value = 0
    ElseIf dgPekerjaan.Columns(5) = 0 Then
        CheckStatusEnbl2.value = 0
    ElseIf dgPekerjaan.Columns(5) = 1 Then
        CheckStatusEnbl2.value = 1
    End If
    txtKdPekerjaan.Enabled = False
End Sub

Private Sub dgPendidikan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPendidikan
    WheelHook.WheelHook dgPendidikan
End Sub

Private Sub dgPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdHapus.SetFocus
    On Error Resume Next
    If dgPendidikan.ApproxCount = 0 Then Exit Sub
    txtKdPendidikan.Text = dgPendidikan.Columns(0).value
    txtPendidikan.Text = dgPendidikan.Columns(1).value
    txtNoUrut.Text = dgPendidikan.Columns(2).value
    txtKodeExternal3.Text = dgPendidikan.Columns(5).value
    txtNamaExternal3.Text = dgPendidikan.Columns(6).value
    If dgPendidikan.Columns(7) = "" Then
        CheckStatusEnbl3.value = 0
    ElseIf dgPendidikan.Columns(7) = 0 Then
        CheckStatusEnbl3.value = 0
    ElseIf dgPendidikan.Columns(7) = 1 Then
        CheckStatusEnbl3.value = 1
    End If
End Sub

Private Sub dgSuku_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgSuku
    WheelHook.WheelHook dgSuku
End Sub

Private Sub dgSuku_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdHapus.SetFocus
    On Error Resume Next
    If dgSuku.ApproxCount = 0 Then Exit Sub
    txtKdSuku.Text = dgSuku.Columns(0).value
    txtSuku.Text = dgSuku.Columns(1).value
    txtKodeExternal1.Text = dgSuku.Columns(3).value
    txtNamaExternal1.Text = dgSuku.Columns(4).value
    If dgSuku.Columns(5) = "" Then
        CheckStatusEnbl1.value = 0
    ElseIf dgSuku.Columns(5) = 0 Then
        CheckStatusEnbl1.value = 0
    ElseIf dgSuku.Columns(5) = 1 Then
        CheckStatusEnbl1.value = 1
    End If
    txtKdSuku.Enabled = False
End Sub

Private Sub dgSuku_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgSuku.ApproxCount = 0 Then Exit Sub
    txtKdSuku.Text = dgSuku.Columns(0).value
    txtSuku.Text = dgSuku.Columns(1).value
    txtKodeExternal1.Text = dgSuku.Columns(3).value
    txtNamaExternal1.Text = dgSuku.Columns(4).value
    If dgSuku.Columns(5) = "" Then
        CheckStatusEnbl1.value = 0
    ElseIf dgSuku.Columns(5) = 0 Then
        CheckStatusEnbl1.value = 0
    ElseIf dgSuku.Columns(5) = 1 Then
        CheckStatusEnbl1.value = 1
    End If
    txtKdSuku.Enabled = False
End Sub

Private Sub dgPekerjaan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgPekerjaan.ApproxCount = 0 Then Exit Sub
    txtKdPekerjaan.Text = dgPekerjaan.Columns(0).value
    txtPekerjaan.Text = dgPekerjaan.Columns(1).value
    txtKodeExternal2.Text = dgPekerjaan.Columns(3).value
    txtNamaExternal2.Text = dgPekerjaan.Columns(4).value
    If dgPekerjaan.Columns(5) = "" Then
        CheckStatusEnbl2.value = 0
    ElseIf dgPekerjaan.Columns(5) = 0 Then
        CheckStatusEnbl2.value = 0
    ElseIf dgPekerjaan.Columns(5) = 1 Then
        CheckStatusEnbl2.value = 1
    End If
    txtKdPekerjaan.Enabled = False
End Sub

Private Sub dgPendidikan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgPendidikan.ApproxCount = 0 Then Exit Sub
    txtKdPendidikan.Text = dgPendidikan.Columns(0).value
    txtPendidikan.Text = dgPendidikan.Columns(1).value
    txtNoUrut.Text = dgPendidikan.Columns(2).value
    txtKodeExternal3.Text = dgPendidikan.Columns(5).value
    txtNamaExternal3.Text = dgPendidikan.Columns(6).value
    If dgPendidikan.Columns(7) = "" Then
        CheckStatusEnbl3.value = 0
    ElseIf dgPendidikan.Columns(7) = 0 Then
        CheckStatusEnbl3.value = 0
    ElseIf dgPendidikan.Columns(7) = 1 Then
        CheckStatusEnbl3.value = 1
    End If
End Sub

Private Sub dgKelompokUmur_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgKelompokUmur.ApproxCount = 0 Then Exit Sub
    txtKdKelompokUmur.Text = dgKelompokUmur.Columns(0).value
    txtKelompokUmur.Text = dgKelompokUmur.Columns(1).value
    txtRangeUmur.Text = dgKelompokUmur.Columns(2).value
    txtUmurMin.Text = dgKelompokUmur.Columns(3).value
    txtUmurMax.Text = dgKelompokUmur.Columns(4).value
    txtketerangan.Text = dgKelompokUmur.Columns(5).value
    txtKodeExternal4.Text = dgKelompokUmur.Columns(6).value
    txtNamaExternal4.Text = dgKelompokUmur.Columns(7).value
    If dgKelompokUmur.Columns(8) = "" Then
        CheckStatusEnbl4.value = 0
    ElseIf dgKelompokUmur.Columns(8) = 0 Then
        CheckStatusEnbl4.value = 0
    ElseIf dgKelompokUmur.Columns(8) = 1 Then
        CheckStatusEnbl4.value = 1
    End If
End Sub

Private Sub dgHubunganKeluarga_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgHubunganKeluarga.ApproxCount = 0 Then Exit Sub
    txtKdHubungan.Text = dgHubunganKeluarga.Columns(0).value
    txtHubungan.Text = dgHubunganKeluarga.Columns(1).value
    txtKodeExternal5.Text = dgHubunganKeluarga.Columns(2).value
    txtNamaExternal5.Text = dgHubunganKeluarga.Columns(3).value
    If dgHubunganKeluarga.Columns(4) = "" Then
        CheckStatusEnbl5.value = 0
    ElseIf dgHubunganKeluarga.Columns(4) = 0 Then
        CheckStatusEnbl5.value = 0
    ElseIf dgHubunganKeluarga.Columns(4) = 1 Then
        CheckStatusEnbl5.value = 1
    End If
End Sub

Private Sub dgGolonganDarah_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgGolonganDarah.ApproxCount = 0 Then Exit Sub
    txtKdGolonganDarah.Text = dgGolonganDarah.Columns(0).value
    txtGolonganDarah.Text = dgGolonganDarah.Columns(1).value
    txtKodeExternal6.Text = dgGolonganDarah.Columns(2).value
    txtNamaExternal6.Text = dgGolonganDarah.Columns(3).value
    If dgGolonganDarah.Columns(4) = "" Then
        CheckStatusEnbl6.value = 0
    ElseIf dgGolonganDarah.Columns(4) = 0 Then
        CheckStatusEnbl6.value = 0
    ElseIf dgGolonganDarah.Columns(4) = 1 Then
        CheckStatusEnbl6.value = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    sstMasterUmum.Tab = 0
    Call cmdBatal_Click
End Sub

Private Sub subLoadGridSource()
    On Error GoTo errLoad
    Select Case sstMasterUmum.Tab
        Case 0 'Agama
            strSQL = "select * from Agama"
            Call msubRecFO(rs, strSQL)
            Set dgAgama.DataSource = rs
            dgAgama.Columns(0).DataField = rs(0).Name
            dgAgama.Columns(1).DataField = rs(1).Name
            dgAgama.Columns(1).Width = 2000

        Case 1 'Suku
            strSQL = "select * from Suku"
            Call msubRecFO(rs, strSQL)
            Set dgSuku.DataSource = rs
            dgSuku.Columns(0).DataField = rs(0).Name
            dgSuku.Columns(1).DataField = rs(1).Name
            dgSuku.Columns(1).Width = 3000

        Case 2  'Pekerjaan
            strSQL = "select * from Pekerjaan"
            Call msubRecFO(rs, strSQL)
            Set dgPekerjaan.DataSource = rs
            dgPekerjaan.Columns(0).DataField = rs(0).Name
            dgPekerjaan.Columns(1).DataField = rs(1).Name
            dgPekerjaan.Columns(1).Width = 4000

        Case 3 'Pendidikan
            strSQL = "select * from Pendidikan"
            Call msubRecFO(rs, strSQL)
            Set dgPendidikan.DataSource = rs
            dgPendidikan.Columns(0).DataField = rs(0).Name
            dgPendidikan.Columns(1).DataField = rs(1).Name
            dgPendidikan.Columns(1).Width = 3500
            dgPendidikan.Columns(2).DataField = rs(2).Name

        Case 4 'Kelompok Umur
            strSQL = "select * from KelompokUmur"
            Call msubRecFO(rs, strSQL)
            Set dgKelompokUmur.DataSource = rs
            dgKelompokUmur.Columns(0).DataField = rs(0).Name
            dgKelompokUmur.Columns(1).DataField = rs(1).Name
            dgKelompokUmur.Columns(1).Width = 2000
            dgKelompokUmur.Columns(2).Width = 3000

        Case 5 'Hubungan Keluarga
            strSQL = "select * from HubunganKeluarga"
            Call msubRecFO(rs, strSQL)
            Set dgHubunganKeluarga.DataSource = rs
            dgHubunganKeluarga.Columns(0).DataField = rs(0).Name
            dgHubunganKeluarga.Columns(1).DataField = rs(1).Name
            dgHubunganKeluarga.Columns(0).Width = 1000
            dgHubunganKeluarga.Columns(1).Width = 3000

        Case 6 'Golongan Darah
            strSQL = "select * from GolonganDarah"
            Call msubRecFO(rs, strSQL)
            Set dgGolonganDarah.DataSource = rs
            dgGolonganDarah.Columns(0).DataField = rs(0).Name
            dgGolonganDarah.Columns(1).DataField = rs(1).Name
            dgGolonganDarah.Columns(1).Width = 4000
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Sub clear()
    On Error Resume Next
    Select Case sstMasterUmum.Tab
        Case 0 'Agama
            txtKdAgama.Text = ""
            txtAgama.Text = ""
            txtKodeExternal.Text = ""
            txtNamaExternal.Text = ""
            CheckStatusEnbl.value = 1

        Case 1 'Suku
            txtKdSuku.Text = ""
            txtSuku.Text = ""
            txtKodeExternal1.Text = ""
            txtNamaExternal1.Text = ""
            CheckStatusEnbl1.value = 1

        Case 2 'Pekerjaan
            txtKdPekerjaan.Text = ""
            txtPekerjaan.Text = ""
            txtKodeExternal2.Text = ""
            txtNamaExternal2.Text = ""
            CheckStatusEnbl2.value = 1

        Case 3 'Pendidikan
            txtKdPendidikan.Text = ""
            txtPendidikan.Text = ""
            txtNoUrut.Text = ""
            txtKodeExternal3.Text = ""
            txtNamaExternal3.Text = ""
            CheckStatusEnbl3.value = 1

        Case 4 'Kelompok Umur
            txtKdKelompokUmur.Text = ""
            txtKelompokUmur.Text = ""
            txtRangeUmur.Text = ""
            txtUmurMin.Text = ""
            txtUmurMax.Text = ""
            txtketerangan.Text = ""
            txtKodeExternal4.Text = ""
            txtNamaExternal4.Text = ""
            CheckStatusEnbl4.value = 1

        Case 5 'Hubungan Keluarga
            txtKdHubungan.Text = ""
            txtHubungan.Text = ""
            txtKodeExternal5.Text = ""
            txtNamaExternal5.Text = ""
            CheckStatusEnbl5.value = 1

        Case 6 'Golongan Darah
            txtKdGolonganDarah.Text = ""
            txtGolonganDarah.Text = ""
            txtKodeExternal6.Text = ""
            txtNamaExternal6.Text = ""
            CheckStatusEnbl6.value = 1
    End Select
End Sub

Sub tampildata()
    On Error GoTo errLoad
    Select Case sstMasterUmum.Tab
        Case 0 'Agama
            Set rs = Nothing
            strSQL = "select * from Agama"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgAgama.DataSource = rs
            dgAgama.Columns(0).DataField = rs(0).Name
            dgAgama.Columns(1).DataField = rs(1).Name
            Set rs = Nothing

        Case 1  'Suku
            Set rs = Nothing
            strSQL = "select * from Suku "
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgSuku.DataSource = rs
            dgSuku.Columns(0).DataField = rs(0).Name
            dgSuku.Columns(1).DataField = rs(1).Name
            Set rs = Nothing

        Case 2  'Pekerjaan
            Set rs = Nothing
            strSQL = "select * from Pekerjaan"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgPekerjaan.DataSource = rs
            dgPekerjaan.Columns(0).DataField = rs(0).Name
            dgPekerjaan.Columns(1).DataField = rs(1).Name
            Set rs = Nothing

        Case 3 'Pendidikan
            Set rs = Nothing
            strSQL = "select * from Pendidikan"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgPendidikan.DataSource = rs
            dgPendidikan.Columns(0).DataField = rs(0).Name
            dgPendidikan.Columns(1).DataField = rs(1).Name
            dgPendidikan.ReBind
            Set rs = Nothing

        Case 4  'Kelompok Umur
            Set rs = Nothing
            strSQL = "select * from KelompokUmur "
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgKelompokUmur.DataSource = rs
            dgKelompokUmur.Columns(0).DataField = rs(0).Name
            dgKelompokUmur.Columns(1).DataField = rs(1).Name
            dgKelompokUmur.ReBind
            Set rs = Nothing

        Case 5  'Hubungan Keluarga
            Set rs = Nothing
            strSQL = "select * from HubunganKeluarga"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgHubunganKeluarga.DataSource = rs
            dgHubunganKeluarga.Columns(0).DataField = rs(0).Name
            dgHubunganKeluarga.Columns(1).DataField = rs(1).Name
            dgHubunganKeluarga.ReBind
            Set rs = Nothing

        Case 6  'Golongan Darah
            Set rs = Nothing
            strSQL = "select * from GolonganDarah"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgGolonganDarah.DataSource = rs
            dgGolonganDarah.Columns(0).DataField = rs(0).Name
            dgGolonganDarah.Columns(1).DataField = rs(1).Name
            dgGolonganDarah.ReBind
            Set rs = Nothing
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub sstMasterUmum_Click(PreviousTab As Integer)
    Call clear
    Call subLoadGridSource
End Sub

Private Sub sstMasterUmum_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    If KeyAscii = 13 Then
        Select Case sstMasterUmum.Tab
            Case 0 'Agama
                txtAgama.SetFocus
            Case 1 'Suku
                txtSuku.SetFocus
            Case 2 'Pekerjaan
                txtPekerjaan.SetFocus
            Case 3 'Pendidikan
                txtPendidikan.SetFocus
            Case 4 'Kelompok Umur
                txtKelompokUmur.SetFocus
            Case 5 'Hubungan Keluarga
                txtHubungan.SetFocus
            Case 6 'Golongan Darah
                txtGolonganDarah.SetFocus
        End Select
    End If
errLoad:
End Sub

Private Sub txtAgama_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            txtKodeExternal.SetFocus
    End Select
End Sub

Private Sub txtAgama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtGolonganDarah_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            txtKodeExternal6.SetFocus
    End Select
End Sub

Private Sub txtGolonganDarah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal6.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtHubungan_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            txtKodeExternal5.SetFocus
    End Select
End Sub

Private Sub txtHubungan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal5.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtKelompokUmur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRangeUmur.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtKeterangan_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            txtKodeExternal4.SetFocus
    End Select
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal4.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtPekerjaan_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            txtKodeExternal2.SetFocus
    End Select
End Sub

Private Sub txtPekerjaan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal2.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtPendidikan_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgPendidikan.SetFocus
    End Select
End Sub

Private Sub txtPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNoUrut.SetFocus
End Sub

Private Sub txtNoUrut_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal3.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtRangeUmur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtUmurMin.SetFocus
End Sub

Private Sub txtSuku_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            txtKodeExternal1.SetFocus
    End Select
End Sub

Private Sub txtSuku_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal1.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtKodeExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal.SetFocus
End Sub

Private Sub txtNamaExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl.SetFocus
End Sub

Private Sub CheckStatusEnbl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKodeExternal1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal1.SetFocus
End Sub

Private Sub txtNamaExternal1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl1.SetFocus
End Sub

Private Sub CheckStatusEnbl1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKodeExternal2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal2.SetFocus
End Sub

Private Sub txtNamaExternal2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl2.SetFocus
End Sub

Private Sub CheckStatusEnbl2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKodeExternal3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal3.SetFocus
End Sub

Private Sub txtNamaExternal3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl3.SetFocus
End Sub

Private Sub CheckStatusEnbl3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKodeExternal4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal4.SetFocus
End Sub

Private Sub txtNamaExternal4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl4.SetFocus
End Sub

Private Sub CheckStatusEnbl4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKodeExternal5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal5.SetFocus
End Sub

Private Sub txtNamaExternal5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl5.SetFocus
End Sub

Private Sub CheckStatusEnbl5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKodeExternal6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal6.SetFocus
End Sub

Private Sub txtNamaExternal6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl6.SetFocus
End Sub

Private Sub CheckStatusEnbl6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtUmurMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtketerangan.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtUmurMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtUmurMax.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub
