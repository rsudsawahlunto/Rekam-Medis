VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmMasterDaftarKontrolPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Kontrol Pasien"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   Icon            =   "frmMasterDaftarKontrolPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   9855
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "&Hapus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   10
      Top             =   6960
      Width           =   1455
   End
   Begin TabDlg.SSTab sstDataPenunjang 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
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
      TabCaption(0)   =   "Cara Masuk"
      TabPicture(0)   =   "frmMasterDaftarKontrolPasien.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "InTake Pasien"
      TabPicture(1)   =   "frmMasterDaftarKontrolPasien.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Output Pasien"
      TabPicture(2)   =   "frmMasterDaftarKontrolPasien.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Status Periksa Pasien"
      TabPicture(3)   =   "frmMasterDaftarKontrolPasien.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Triase"
      TabPicture(4)   =   "frmMasterDaftarKontrolPasien.frx":0D3A
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   5175
         Left            =   240
         TabIndex        =   49
         Top             =   480
         Width           =   9375
         Begin VB.CheckBox CheckStatusEnbl4 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7920
            TabIndex        =   34
            Top             =   1800
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtNamaExternal4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   33
            Top             =   1800
            Width           =   5775
         End
         Begin VB.TextBox txtKodeExternal4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   32
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtKdTriase 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   30
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtTriase 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   31
            Top             =   840
            Width           =   7095
         End
         Begin MSDataGridLib.DataGrid dgTriase 
            Height          =   2775
            Left            =   120
            TabIndex        =   35
            Top             =   2280
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   4895
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
         Begin VB.Label Label21 
            Caption         =   "Nama External"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   62
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "Kode External"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   61
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   600
            TabIndex        =   51
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Triase"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   600
            TabIndex        =   50
            Top             =   840
            Width           =   480
         End
      End
      Begin VB.Frame Frame4 
         Height          =   5175
         Left            =   -74760
         TabIndex        =   45
         Top             =   480
         Width           =   9375
         Begin VB.CheckBox CheckStatusEnbl3 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7920
            TabIndex        =   28
            Top             =   1800
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtNamaExternal3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            MaxLength       =   30
            TabIndex        =   27
            Top             =   1800
            Width           =   6135
         End
         Begin VB.TextBox txtKodeExternal3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            MaxLength       =   30
            TabIndex        =   26
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtSingkatan 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8520
            MaxLength       =   1
            TabIndex        =   25
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtKdStatusPeriksa 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   23
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtStatusPeriksa 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   24
            Top             =   840
            Width           =   5535
         End
         Begin MSDataGridLib.DataGrid dgStatusPeriksa 
            Height          =   2775
            Left            =   120
            TabIndex        =   29
            Top             =   2280
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   4895
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
         Begin VB.Label Label19 
            Caption         =   "Nama External"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "Kode External"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Singkatan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7560
            TabIndex        =   48
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   47
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Status Periksa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   46
            Top             =   840
            Width           =   1140
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5175
         Left            =   -74760
         TabIndex        =   42
         Top             =   480
         Width           =   9375
         Begin VB.CheckBox CheckStatusEnbl 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7920
            TabIndex        =   5
            Top             =   1800
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtNamaExternal 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   4
            Top             =   1800
            Width           =   5775
         End
         Begin VB.TextBox txtKodeExternal 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   3
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtCaraMasuk 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   2
            Top             =   840
            Width           =   7095
         End
         Begin VB.TextBox txtKdCaraMasuk 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   1
            Top             =   360
            Width           =   735
         End
         Begin MSDataGridLib.DataGrid dgCaraMasuk 
            Height          =   2775
            Left            =   120
            TabIndex        =   6
            Top             =   2280
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   4895
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
         Begin VB.Label Label13 
            Caption         =   "Nama External"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   54
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Kode External"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   53
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cara Masuk"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   600
            TabIndex        =   44
            Top             =   840
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   600
            TabIndex        =   43
            Top             =   360
            Width           =   420
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5175
         Left            =   -74760
         TabIndex        =   39
         Top             =   480
         Width           =   9375
         Begin VB.CheckBox CheckStatusEnbl1 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7920
            TabIndex        =   15
            Top             =   1800
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtNamaExternal1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   14
            Top             =   1800
            Width           =   5775
         End
         Begin VB.TextBox txtKodeExternal1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   13
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtKdIntakePasien 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   11
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtInTakePasien 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   12
            Top             =   840
            Width           =   7095
         End
         Begin MSDataGridLib.DataGrid dgInTakePasien 
            Height          =   2775
            Left            =   120
            TabIndex        =   16
            Top             =   2280
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   4895
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
         Begin VB.Label Label15 
            Caption         =   "Nama External"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   56
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Kode External"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   55
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   600
            TabIndex        =   41
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "InTake Pasien"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   600
            TabIndex        =   40
            Top             =   840
            Width           =   1140
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5175
         Left            =   -74760
         TabIndex        =   36
         Top             =   480
         Width           =   9375
         Begin VB.CheckBox CheckStatusEnbl2 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7920
            TabIndex        =   21
            Top             =   1800
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtNamaExternal2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   20
            Top             =   1800
            Width           =   5775
         End
         Begin VB.TextBox txtKodeExternal2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   19
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtKdOutputPasien 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   17
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtOutputPasien 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   18
            Top             =   840
            Width           =   7095
         End
         Begin MSDataGridLib.DataGrid dgOutputPasien 
            Height          =   2775
            Left            =   120
            TabIndex        =   22
            Top             =   2280
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   4895
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
         Begin VB.Label Label17 
            Caption         =   "Nama External"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   58
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Kode External"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   57
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   600
            TabIndex        =   38
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Output Pasien"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   600
            TabIndex        =   37
            Top             =   840
            Width           =   1170
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   52
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
      Left            =   8040
      Picture         =   "frmMasterDaftarKontrolPasien.frx":0D56
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterDaftarKontrolPasien.frx":1ADE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterDaftarKontrolPasien.frx":313C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmMasterDaftarKontrolPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoCommand As New ADODB.Command

Private Sub CheckStatusEnbl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub CheckStatusEnbl1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub CheckStatusEnbl2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub CheckStatusEnbl3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub CheckStatusEnbl4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub cmdBatal_Click()
    On Error GoTo errLoad
    Call clear
    Call subLoadGridSource
    Call sstDataPenunjang_KeyPress(13)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errLoad

    If MsgBox("Apakah anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub

    Select Case sstDataPenunjang.Tab
        Case 0 'Cara Masuk
            If Periksa("text", txtCaraMasuk, "Nama cara masuk kosong") = False Then Exit Sub
            If sp_CaraMasuk("D") = False Then Exit Sub
        Case 1 'InTake Pasien
            If Periksa("text", txtInTakePasien, "Nama InTake pasien kosong") = False Then Exit Sub
            If sp_IntakePasien("D") = False Then Exit Sub
        Case 2 'Output Pasien
            If Periksa("text", txtOutputPasien, "Nama Ouput pasien kosong") = False Then Exit Sub
            If sp_OutputPasien("D") = False Then Exit Sub
        Case 3 'Status Periksa Pasien
            If Periksa("text", txtStatusPeriksa, "Nama status periksa pasien kosong") = False Then Exit Sub
            If Periksa("text", txtSingkatan, "Singkatan status periksa pasien kosong") = False Then Exit Sub
            If sp_StatusPeriksaPasien("D") = False Then Exit Sub
        Case 4 'Triase
            If Periksa("text", txtTriase, "Nama triase pasien kosong") = False Then Exit Sub
            If sp_Triase("D") = False Then Exit Sub
    End Select

    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
    Call cmdBatal_Click
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function sp_OutputPasien(f_Status As String) As Boolean
    sp_OutputPasien = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdOutput", adChar, adParamInput, 2, txtKdOutputPasien.Text)
        .Parameters.Append .CreateParameter("NamaOutput", adVarChar, adParamInput, 30, Trim(txtOutputPasien.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal2.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, txtNamaExternal2.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl2.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_OutputPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Output Pasien", vbCritical
            sp_OutputPasien = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_IntakePasien(f_Status As String) As Boolean
    sp_IntakePasien = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdInTake", adChar, adParamInput, 2, txtKdIntakePasien.Text)
        .Parameters.Append .CreateParameter("NamaInTake", adVarChar, adParamInput, 30, Trim(txtInTakePasien.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal1.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, txtNamaExternal1.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl1.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_InTakePasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan InTake Pasien", vbCritical
            sp_IntakePasien = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_CaraMasuk(f_Status As String) As Boolean
    sp_CaraMasuk = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdCaraMasuk", adChar, adParamInput, 2, txtKdCaraMasuk.Text)
        .Parameters.Append .CreateParameter("CaraMasuk", adVarChar, adParamInput, 30, Trim(txtCaraMasuk.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, txtNamaExternal.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_CaraMasuk"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Cara Masuk Pasien", vbCritical
            sp_CaraMasuk = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_StatusPeriksaPasien(f_Status As String) As Boolean
    sp_StatusPeriksaPasien = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdStatusPeriksa", adChar, adParamInput, 2, txtKdStatusPeriksa.Text)
        .Parameters.Append .CreateParameter("StatusPeriksa", adVarChar, adParamInput, 30, Trim(txtStatusPeriksa.Text))
        .Parameters.Append .CreateParameter("Singkatan", adChar, adParamInput, 1, txtSingkatan.Text)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal3.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, txtNamaExternal3.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl3.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_StatusPeriksaPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Status Periksa Pasien", vbCritical
            sp_StatusPeriksaPasien = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_Triase(f_Status As String) As Boolean
    sp_Triase = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdTriase", adChar, adParamInput, 2, txtKdTriase.Text)
        .Parameters.Append .CreateParameter("NamaTriase", adVarChar, adParamInput, 30, Trim(txtTriase.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal4.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, txtNamaExternal4.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl4.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_Triase"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Triase Pasien", vbCritical
            sp_Triase = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    Select Case sstDataPenunjang.Tab
        Case 0 'Cara Masuk
            If Periksa("text", txtCaraMasuk, "Nama cara masuk kosong") = False Then Exit Sub
            If sp_CaraMasuk("A") = False Then Exit Sub
        Case 1 'InTake Pasien
            If Periksa("text", txtInTakePasien, "Nama InTake pasien kosong") = False Then Exit Sub
            If sp_IntakePasien("A") = False Then Exit Sub
        Case 2 'Output Pasien
            If Periksa("text", txtOutputPasien, "Nama Ouput pasien kosong") = False Then Exit Sub
            If sp_OutputPasien("A") = False Then Exit Sub
        Case 3 'Status Periksa Pasien
            If Periksa("text", txtStatusPeriksa, "Nama status periksa pasien kosong") = False Then Exit Sub
            If Periksa("text", txtSingkatan, "Singkatan status periksa pasien kosong") = False Then Exit Sub
            If sp_StatusPeriksaPasien("A") = False Then Exit Sub
        Case 4 'Triase
            If Periksa("text", txtTriase, "Nama triase pasien kosong") = False Then Exit Sub
            If sp_Triase("A") = False Then Exit Sub
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

Private Sub dgCaraMasuk_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgCaraMasuk
    WheelHook.WheelHook dgCaraMasuk
End Sub

Private Sub dgCaraMasuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtCaraMasuk.SetFocus
End Sub

Private Sub dgCaraMasuk_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgCaraMasuk.ApproxCount = 0 Then Exit Sub
    txtKdCaraMasuk.Text = dgCaraMasuk.Columns(0).value
    txtCaraMasuk.Text = dgCaraMasuk.Columns(1).value
    txtKodeExternal.Text = dgCaraMasuk.Columns(3).value
    txtNamaExternal.Text = dgCaraMasuk.Columns(4).value
    If dgCaraMasuk.Columns(5) = "" Then
        CheckStatusEnbl.value = 0
    ElseIf dgCaraMasuk.Columns(5) = 0 Then
        CheckStatusEnbl.value = 0
    ElseIf dgCaraMasuk.Columns(5) = 1 Then
        CheckStatusEnbl.value = 1
    End If
End Sub

Private Sub dgInTakePasien_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgInTakePasien
    WheelHook.WheelHook dgInTakePasien
End Sub

Private Sub dgInTakePasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtInTakePasien.SetFocus
End Sub

Private Sub dgInTakePasien_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgInTakePasien.ApproxCount = 0 Then Exit Sub
    txtKdIntakePasien.Text = dgInTakePasien.Columns(0).value
    txtInTakePasien.Text = dgInTakePasien.Columns(1).value
    txtKodeExternal1.Text = dgInTakePasien.Columns(2).value
    txtNamaExternal1.Text = dgInTakePasien.Columns(3).value
    If dgInTakePasien.Columns(4) = "" Then
        CheckStatusEnbl1.value = 0
    ElseIf dgInTakePasien.Columns(4) = 0 Then
        CheckStatusEnbl1.value = 0
    ElseIf dgInTakePasien.Columns(4) = 1 Then
        CheckStatusEnbl1.value = 1
    End If
End Sub

Private Sub dgOutputPasien_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgOutputPasien
    WheelHook.WheelHook dgOutputPasien
End Sub

Private Sub dgOutputPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtOutputPasien.SetFocus
End Sub

Private Sub dgOutputPasien_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgOutputPasien.ApproxCount = 0 Then Exit Sub
    txtKdOutputPasien.Text = dgOutputPasien.Columns(0).value
    txtOutputPasien.Text = dgOutputPasien.Columns(1).value
    txtKodeExternal2.Text = dgOutputPasien.Columns(2).value
    txtNamaExternal2.Text = dgOutputPasien.Columns(3).value
    If dgOutputPasien.Columns(4) = "" Then
        CheckStatusEnbl2.value = 0
    ElseIf dgOutputPasien.Columns(4) = 0 Then
        CheckStatusEnbl2.value = 0
    ElseIf dgOutputPasien.Columns(4) = 1 Then
        CheckStatusEnbl2.value = 1
    End If
End Sub

Private Sub dgStatusPeriksa_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgStatusPeriksa
    WheelHook.WheelHook dgStatusPeriksa
End Sub

Private Sub dgStatusPeriksa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtStatusPeriksa.SetFocus
End Sub

Private Sub dgStatusPeriksa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgStatusPeriksa.ApproxCount = 0 Then Exit Sub
    txtKdStatusPeriksa.Text = dgStatusPeriksa.Columns(0).value
    txtStatusPeriksa.Text = dgStatusPeriksa.Columns(1).value
    txtSingkatan.Text = dgStatusPeriksa.Columns(2).value
    txtKodeExternal3.Text = dgStatusPeriksa.Columns(3).value
    txtNamaExternal3.Text = dgStatusPeriksa.Columns(4).value
    If dgStatusPeriksa.Columns(5) = "" Then
        CheckStatusEnbl3.value = 0
    ElseIf dgStatusPeriksa.Columns(5) = 0 Then
        CheckStatusEnbl3.value = 0
    ElseIf dgStatusPeriksa.Columns(5) = 1 Then
        CheckStatusEnbl3.value = 1
    End If
End Sub

Private Sub dgTriase_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgTriase
    WheelHook.WheelHook dgTriase
End Sub

Private Sub dgTriase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTriase.SetFocus
End Sub

Private Sub dgTriase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgTriase.ApproxCount = 0 Then Exit Sub
    txtKdTriase.Text = dgTriase.Columns(0).value
    txtTriase.Text = dgTriase.Columns(1).value
    txtKodeExternal4.Text = dgTriase.Columns(2).value
    txtNamaExternal4.Text = dgTriase.Columns(3).value
    If dgTriase.Columns(4) = "" Then
        CheckStatusEnbl4.value = 0
    ElseIf dgTriase.Columns(4) = 0 Then
        CheckStatusEnbl4.value = 0
    ElseIf dgTriase.Columns(4) = 1 Then
        CheckStatusEnbl4.value = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    sstDataPenunjang.Tab = 0
    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()
    On Error GoTo errLoad

    Select Case sstDataPenunjang.Tab
        Case 0 'Cara Masuk
            strSQL = "select * from CaraMasuk"
            Call msubRecFO(rs, strSQL)
            Set dgCaraMasuk.DataSource = rs
            dgCaraMasuk.Columns(0).DataField = rs(0).Name
            dgCaraMasuk.Columns(1).DataField = rs(1).Name
            dgCaraMasuk.Columns(1).Width = 3000

        Case 1 'InTake Pasien
            strSQL = "select * from InTakePasien"
            Call msubRecFO(rs, strSQL)
            Set dgInTakePasien.DataSource = rs
            dgInTakePasien.Columns(0).DataField = rs(0).Name
            dgInTakePasien.Columns(1).DataField = rs(1).Name
            dgInTakePasien.Columns(1).Width = 3000

        Case 2  'Output Pasien
            strSQL = "select * from OutputPasien"
            Call msubRecFO(rs, strSQL)
            Set dgOutputPasien.DataSource = rs
            dgOutputPasien.Columns(0).DataField = rs(0).Name
            dgOutputPasien.Columns(1).DataField = rs(1).Name
            dgOutputPasien.Columns(1).Width = 3000

        Case 3  'Status Periksa Pasien
            strSQL = "select * from StatusPeriksaPasien"
            Call msubRecFO(rs, strSQL)
            Set dgStatusPeriksa.DataSource = rs
            dgStatusPeriksa.Columns(0).DataField = rs(0).Name
            dgStatusPeriksa.Columns(1).DataField = rs(1).Name
            dgStatusPeriksa.Columns(1).Width = 3000
            dgStatusPeriksa.Columns(2).DataField = rs(2).Name

        Case 4  'Triase
            strSQL = "select * from Triase"
            Call msubRecFO(rs, strSQL)
            Set dgTriase.DataSource = rs
            dgTriase.Columns(0).DataField = rs(0).Name
            dgTriase.Columns(1).DataField = rs(1).Name
            dgTriase.Columns(1).Width = 3000
    End Select

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub clear()
    On Error Resume Next
    Select Case sstDataPenunjang.Tab
        Case 0 'Cara Masuk
            txtKdCaraMasuk.Text = ""
            txtCaraMasuk.Text = ""
            txtKodeExternal.Text = ""
            txtNamaExternal.Text = ""
            CheckStatusEnbl.value = 1

        Case 1 'InTake Pasien
            txtKdIntakePasien.Text = ""
            txtInTakePasien.Text = ""
            txtKodeExternal1.Text = ""
            txtNamaExternal1.Text = ""
            CheckStatusEnbl1.value = 1

        Case 2 'Output Pasien
            txtKdOutputPasien.Text = ""
            txtOutputPasien.Text = ""
            txtKodeExternal2.Text = ""
            txtNamaExternal2.Text = ""
            CheckStatusEnbl2.value = 1

        Case 3 'Status Periksa Pasien
            txtKdStatusPeriksa.Text = ""
            txtStatusPeriksa.Text = ""
            txtSingkatan.Text = ""
            txtKodeExternal3.Text = ""
            txtNamaExternal3.Text = ""
            CheckStatusEnbl3.value = 1

        Case 4 'Triase
            txtKdTriase.Text = ""
            txtTriase.Text = ""
            txtKodeExternal4.Text = ""
            txtNamaExternal4.Text = ""
            CheckStatusEnbl4.value = 1
    End Select
End Sub

Private Sub sstDataPenunjang_Click(PreviousTab As Integer)
    Call clear
    Call subLoadGridSource
End Sub

Private Sub sstDataPenunjang_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    If KeyAscii = 13 Then
        Select Case sstDataPenunjang.Tab
            Case 0
                txtCaraMasuk.SetFocus
            Case 1
                txtInTakePasien.SetFocus
            Case 2
                txtOutputPasien.SetFocus
            Case 3
                txtStatusPeriksa.SetFocus
            Case 4
                txtTriase.SetFocus
        End Select
    End If
errLoad:
End Sub

Private Sub txtCaraMasuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal.SetFocus
End Sub

Private Sub txtCaraMasuk_LostFocus()
    txtCaraMasuk.Text = Trim(StrConv(txtCaraMasuk.Text, vbProperCase))
End Sub

Private Sub txtInTakePasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal1.SetFocus
End Sub

Private Sub txtInTakePasien_LostFocus()
    txtInTakePasien.Text = Trim(StrConv(txtInTakePasien.Text, vbProperCase))
End Sub

Private Sub txtKodeExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal.SetFocus
End Sub

Private Sub txtKodeExternal1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal1.SetFocus
End Sub

Private Sub txtKodeExternal2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal2.SetFocus
End Sub

Private Sub txtKodeExternal3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal3.SetFocus
End Sub

Private Sub txtKodeExternal4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal4.SetFocus
End Sub

Private Sub txtNamaExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl.SetFocus
End Sub

Private Sub txtNamaExternal1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl1.SetFocus
End Sub

Private Sub txtNamaExternal2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl2.SetFocus
End Sub

Private Sub txtNamaExternal3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl3.SetFocus
End Sub

Private Sub txtNamaExternal4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl4.SetFocus
End Sub

Private Sub txtOutputPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal2.SetFocus
End Sub

Private Sub txtOutputPasien_LostFocus()
    txtOutputPasien.Text = Trim(StrConv(txtOutputPasien.Text, vbProperCase))
End Sub

Private Sub txtSingkatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal3.SetFocus
End Sub

Private Sub txtSingkatan_LostFocus()
    txtSingkatan.Text = StrConv(txtSingkatan.Text, vbUpperCase)
End Sub

Private Sub txtStatusPeriksa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtSingkatan.SetFocus
End Sub

Private Sub txtStatusPeriksa_LostFocus()
    txtStatusPeriksa.Text = Trim(StrConv(txtStatusPeriksa.Text, vbProperCase))
End Sub

Private Sub txtTriase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal4.SetFocus
End Sub

Private Sub txtTriase_LostFocus()
    txtTriase.Text = Trim(StrConv(txtTriase.Text, vbProperCase))
End Sub

