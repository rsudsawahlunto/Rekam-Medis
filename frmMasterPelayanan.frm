VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmMasterPelayanan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Setting Pelayanan"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMasterPelayanan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   7965
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   6840
      Width           =   1575
   End
   Begin TabDlg.SSTab sstDataPendukung 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   8
      Tab             =   7
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
      TabCaption(0)   =   "Status Pulang"
      TabPicture(0)   =   "frmMasterPelayanan.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame8"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Kondisi Pulang"
      TabPicture(1)   =   "frmMasterPelayanan.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Kesadaran"
      TabPicture(2)   =   "frmMasterPelayanan.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame10"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Jenis Diagnosa"
      TabPicture(3)   =   "frmMasterPelayanan.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame12"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Status Keluar"
      TabPicture(4)   =   "frmMasterPelayanan.frx":0D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame14"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Rujukan Asal"
      TabPicture(5)   =   "frmMasterPelayanan.frx":0D56
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame13"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Detail Rujukan Asal"
      TabPicture(6)   =   "frmMasterPelayanan.frx":0D72
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame1"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Kondisi Keluar"
      TabPicture(7)   =   "frmMasterPelayanan.frx":0D8E
      Tab(7).ControlEnabled=   -1  'True
      Tab(7).Control(0)=   "Frame2"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).ControlCount=   1
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   240
         TabIndex        =   81
         Top             =   840
         Width           =   7335
         Begin VB.TextBox txtKodeExternal7 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   52
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox txtNamaExternal7 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   54
            Top             =   1560
            Width           =   5055
         End
         Begin VB.CheckBox CheckStatusEnbl7 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   5880
            TabIndex        =   53
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtKondisiKeluar 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   51
            Top             =   840
            Width           =   5055
         End
         Begin VB.TextBox txtKdKondisiKeluar 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   50
            Top             =   480
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid dgKondisiKeluar 
            Height          =   2535
            Left            =   120
            TabIndex        =   55
            Top             =   2040
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4471
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
         Begin VB.Label Label35 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   99
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label34 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   98
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Kondisi Keluar"
            Height          =   210
            Left            =   240
            TabIndex        =   83
            Top             =   840
            Width           =   1110
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Kode Kondisi"
            Height          =   210
            Left            =   240
            TabIndex        =   82
            Top             =   480
            Width           =   1035
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   -74760
         TabIndex        =   77
         Top             =   840
         Width           =   7335
         Begin VB.CheckBox CheckStatusEnbl6 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   5880
            TabIndex        =   47
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtNamaExternal6 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   48
            Top             =   1560
            Width           =   5535
         End
         Begin VB.TextBox txtKodeExternal6 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   46
            Top             =   1200
            Width           =   1815
         End
         Begin MSDataListLib.DataCombo dcRujukanAsal 
            Height          =   330
            Left            =   4320
            TabIndex        =   44
            Top             =   360
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.TextBox txtDetailRujukanAsal 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1680
            MaxLength       =   100
            TabIndex        =   45
            Top             =   840
            Width           =   3735
         End
         Begin VB.TextBox txtKDDetailRujukanAsal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1680
            MaxLength       =   8
            TabIndex        =   43
            Text            =   "12345678"
            Top             =   360
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid dgDetailRujukanAsal 
            Height          =   2535
            Left            =   120
            TabIndex        =   49
            Top             =   2040
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4471
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
         Begin VB.Label Label33 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   97
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label32 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   96
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Rujukan Asal"
            Height          =   210
            Left            =   3120
            TabIndex        =   80
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Rujukan Asal"
            Height          =   210
            Left            =   240
            TabIndex        =   79
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Kode "
            Height          =   210
            Left            =   240
            TabIndex        =   78
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Frame14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   -74760
         TabIndex        =   74
         Top             =   840
         Width           =   7335
         Begin VB.CheckBox CheckStatusEnbl4 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   5880
            TabIndex        =   33
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtNamaExternal4 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   34
            Top             =   1560
            Width           =   5055
         End
         Begin VB.TextBox txtKodeExternal4 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   32
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox txtKdStatusKeluar 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   30
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtStatusKeluar 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   31
            Text            =   "50"
            Top             =   840
            Width           =   5055
         End
         Begin MSDataGridLib.DataGrid dgStatusKeluar 
            Height          =   2535
            Left            =   120
            TabIndex        =   35
            Top             =   2040
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4471
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
         Begin VB.Label Label29 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   93
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label28 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Kode Status Keluar "
            Height          =   210
            Left            =   240
            TabIndex        =   76
            Top             =   480
            Width           =   1620
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Status Keluar"
            Height          =   210
            Left            =   240
            TabIndex        =   75
            Top             =   840
            Width           =   1080
         End
      End
      Begin VB.Frame Frame13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   -74760
         TabIndex        =   70
         Top             =   840
         Width           =   7335
         Begin VB.TextBox txtKodeExternal5 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            TabIndex        =   39
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox txtNamaExternal5 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            TabIndex        =   41
            Top             =   1560
            Width           =   5175
         End
         Begin VB.CheckBox CheckStatusEnbl5 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   5880
            TabIndex        =   40
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtKdRujukanAsal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   36
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtRujukanAsal 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   37
            Text            =   "30"
            Top             =   840
            Width           =   3135
         End
         Begin VB.TextBox txtSingkatanRujukan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6480
            MaxLength       =   5
            TabIndex        =   38
            Top             =   840
            Width           =   735
         End
         Begin MSDataGridLib.DataGrid dgRujukanAsal 
            Height          =   2535
            Left            =   120
            TabIndex        =   42
            Top             =   2040
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4471
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
         Begin VB.Label Label31 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   95
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label30 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   94
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Kode Rujukan Asal"
            Height          =   210
            Left            =   240
            TabIndex        =   73
            Top             =   480
            Width           =   1500
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Rujukan Asal"
            Height          =   210
            Left            =   240
            TabIndex        =   72
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Singkatan"
            Height          =   210
            Left            =   5520
            TabIndex        =   71
            Top             =   840
            Width           =   795
         End
      End
      Begin VB.Frame Frame8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4755
         Left            =   -74760
         TabIndex        =   65
         Top             =   780
         Width           =   7335
         Begin VB.CheckBox CheckStatusEnbl 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   5880
            TabIndex        =   4
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtNamaExternal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2280
            TabIndex        =   5
            Top             =   1560
            Width           =   4935
         End
         Begin VB.TextBox txtKodeExternal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2280
            TabIndex        =   3
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox txtKdStatusPulang 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   1
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtStatusPulang 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   2
            Text            =   "50"
            Top             =   840
            Width           =   4935
         End
         Begin MSDataGridLib.DataGrid dgStatusPulang 
            Height          =   2535
            Left            =   120
            TabIndex        =   6
            Top             =   2040
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4471
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
         Begin VB.Label Label9 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   85
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   84
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Kode Status Pulang"
            Height          =   210
            Left            =   240
            TabIndex        =   67
            Top             =   480
            Width           =   1605
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Status Pulang"
            Height          =   210
            Left            =   240
            TabIndex        =   66
            Top             =   840
            Width           =   1125
         End
      End
      Begin VB.Frame Frame9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4755
         Left            =   -74760
         TabIndex        =   62
         Top             =   780
         Width           =   7335
         Begin VB.TextBox txtKodeExternal1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2280
            TabIndex        =   14
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox txtNamaExternal1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2280
            TabIndex        =   16
            Top             =   1560
            Width           =   4935
         End
         Begin VB.CheckBox CheckStatusEnbl1 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   5880
            TabIndex        =   15
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtSingkatan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6480
            MaxLength       =   5
            TabIndex        =   13
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox txtKondisiPulang 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   12
            Text            =   "50"
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox txtKdKondisiPulang 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   11
            Top             =   480
            Width           =   1215
         End
         Begin MSDataGridLib.DataGrid dgKondisiPulang 
            Height          =   2535
            Left            =   120
            TabIndex        =   17
            Top             =   2040
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4471
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
         Begin VB.Label Label13 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   87
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   86
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Singkatan"
            Height          =   210
            Left            =   5520
            TabIndex        =   69
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Kondisi Pulang"
            Height          =   210
            Left            =   240
            TabIndex        =   64
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Kode Kondisi Pulang"
            Height          =   210
            Left            =   240
            TabIndex        =   63
            Top             =   480
            Width           =   1635
         End
      End
      Begin VB.Frame Frame10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4755
         Left            =   -74760
         TabIndex        =   59
         Top             =   780
         Width           =   7335
         Begin VB.CheckBox CheckStatusEnbl2 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   5880
            TabIndex        =   21
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtNamaExternal2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2280
            TabIndex        =   22
            Top             =   1560
            Width           =   4935
         End
         Begin VB.TextBox txtKodeExternal2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2280
            TabIndex        =   20
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox txtKesadaran 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   19
            Text            =   "50"
            Top             =   840
            Width           =   4935
         End
         Begin VB.TextBox txtKdKesadaran 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   18
            Top             =   480
            Width           =   1215
         End
         Begin MSDataGridLib.DataGrid dgKesadaran 
            Height          =   2535
            Left            =   120
            TabIndex        =   23
            Top             =   2040
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4471
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
         Begin VB.Label Label25 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   89
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   88
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Kesadaran"
            Height          =   210
            Left            =   240
            TabIndex        =   61
            Top             =   840
            Width           =   825
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Kode Kesadaran"
            Height          =   210
            Left            =   240
            TabIndex        =   60
            Top             =   480
            Width           =   1305
         End
      End
      Begin VB.Frame Frame12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4755
         Left            =   -74760
         TabIndex        =   56
         Top             =   780
         Width           =   7335
         Begin VB.TextBox txtKodeExternal3 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   26
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox txtNamaExternal3 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   28
            Top             =   1560
            Width           =   5055
         End
         Begin VB.CheckBox CheckStatusEnbl3 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   5880
            TabIndex        =   27
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtJenisDiagnosa 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2160
            MaxLength       =   30
            TabIndex        =   25
            Text            =   "30"
            Top             =   840
            Width           =   5055
         End
         Begin VB.TextBox txtKdJenisDiagnosa 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   24
            Top             =   480
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid dgJenisDiagnosa 
            Height          =   2535
            Left            =   120
            TabIndex        =   29
            Top             =   2040
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4471
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
         Begin VB.Label Label27 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   91
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label26 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   90
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Diagnosa"
            Height          =   210
            Left            =   240
            TabIndex        =   58
            Top             =   840
            Width           =   1170
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Kode Jenis Diagnosa"
            Height          =   210
            Left            =   240
            TabIndex        =   57
            Top             =   480
            Width           =   1650
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   68
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
      Left            =   6120
      Picture         =   "frmMasterPelayanan.frx":0DAA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterPelayanan.frx":1B32
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterPelayanan.frx":3190
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmMasterPelayanan"
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
    Call sstDataPendukung_KeyPress(13)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errLoad
    If MsgBox("Yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Select Case sstDataPendukung.Tab
        Case 0 'Status Pulang
            If txtKdStatusPulang.Text = "" Then Exit Sub
            If sp_StatusPulang("D") = False Then Exit Sub
        Case 1 'Kondisi Pulang
            If txtKdKondisiPulang.Text = "" Then Exit Sub
            If sp_KondisiPulang("D") = False Then Exit Sub
        Case 2 'Kesadaran
            If txtKdKesadaran.Text = "" Then Exit Sub
            If sp_Kesadaran("D") = False Then Exit Sub
        Case 3 'Jenis Diagnosa
            If txtKdJenisDiagnosa.Text = "" Then Exit Sub
            If sp_JenisDiagnosa("D") = False Then Exit Sub
        Case 4 'Status Keluar Kamar
            If txtKdStatusKeluar.Text = "" Then Exit Sub
            If sp_StatusKeluarKamar("D") = False Then Exit Sub
        Case 5 'Rujukan Asal
            If txtKdRujukanAsal.Text = "" Then Exit Sub
            If sp_RujukanAsal("D") = False Then Exit Sub
        Case 6 'Detail Rujukan Asal
            If txtKDDetailRujukanAsal.Text = "" Then Exit Sub
            If sp_DetailRujukanAsal("D") = False Then Exit Sub

        Case 7 'Kondisi Keluar
            If txtKdKondisiKeluar.Text = "" Then Exit Sub
            If sp_KondisiKeluar("D") = False Then Exit Sub
    End Select
    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
    Call cmdBatal_Click
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function sp_StatusPulang(f_Status As String) As Boolean
    sp_StatusPulang = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdStatusPulang", adChar, adParamInput, 2, txtKdStatusPulang.Text)
        .Parameters.Append .CreateParameter("StatusPulang", adVarChar, adParamInput, 50, Trim(txtStatusPulang.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNamaExternal.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_StatusPulang"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Status Pulang", vbCritical
            sp_StatusPulang = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_KondisiPulang(f_Status As String) As Boolean
    sp_KondisiPulang = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKondisiPulang", adChar, adParamInput, 2, txtKdKondisiPulang.Text)
        .Parameters.Append .CreateParameter("KondisiPulang", adVarChar, adParamInput, 50, Trim(txtKondisiPulang.Text))
        .Parameters.Append .CreateParameter("Singkatan", adVarChar, adParamInput, 5, IIf(txtSingkatan.Text = "", Null, txtSingkatan.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal1.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNamaExternal1.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl1.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_KondisiPulang"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Kondisi Pulang", vbCritical
            sp_KondisiPulang = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_Kesadaran(f_Status As String) As Boolean
    sp_Kesadaran = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKesadaran", adChar, adParamInput, 2, txtKdKesadaran.Text)
        .Parameters.Append .CreateParameter("NamaKesadaran", adVarChar, adParamInput, 50, Trim(txtKesadaran.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal2.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNamaExternal2.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl2.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_Kesadaran"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Kesadaran", vbCritical
            sp_Kesadaran = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_JenisDiagnosa(f_Status As String) As Boolean
    sp_JenisDiagnosa = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdJenisDiagnosa", adChar, adParamInput, 2, txtKdJenisDiagnosa.Text)
        .Parameters.Append .CreateParameter("JenisDiagnosa", adVarChar, adParamInput, 30, Trim(txtJenisDiagnosa.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal3.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, txtNamaExternal3.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl3.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_JenisDiagnosa"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Jenis Diagnosa", vbCritical
            sp_JenisDiagnosa = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_RujukanAsal(f_Status As String) As Boolean
    sp_RujukanAsal = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdRujukanAsal", adChar, adParamInput, 2, txtKdRujukanAsal.Text)
        .Parameters.Append .CreateParameter("RujukanAsal", adVarChar, adParamInput, 30, Trim(txtRujukanAsal.Text))
        .Parameters.Append .CreateParameter("Singkatan", adVarChar, adParamInput, 5, txtSingkatanRujukan.Text)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal5.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, txtNamaExternal5.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl5.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_RujukanAsal"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Rujukan Asal", vbCritical
            sp_RujukanAsal = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_DetailRujukanAsal(f_Status As String) As Boolean
    sp_DetailRujukanAsal = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdDetailRujukanAsal", adChar, adParamInput, 8, IIf(txtKDDetailRujukanAsal.Text = "", Null, txtKDDetailRujukanAsal.Text))
        .Parameters.Append .CreateParameter("DetailRujukanAsal", adVarChar, adParamInput, 100, Trim(txtDetailRujukanAsal.Text))
        .Parameters.Append .CreateParameter("KdRujukanAsal", adChar, adParamInput, 2, dcRujukanAsal.BoundText)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal6.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 100, txtNamaExternal6.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl6.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_DetailRujukanAsal"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Detail Rujukan Asal", vbCritical
            sp_DetailRujukanAsal = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_StatusKeluarKamar(f_Status As String) As Boolean
    sp_StatusKeluarKamar = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdStatusKeluar", adChar, adParamInput, 2, txtKdStatusKeluar.Text)
        .Parameters.Append .CreateParameter("StatusKeluar", adVarChar, adParamInput, 50, Trim(txtStatusKeluar.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal4.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNamaExternal4.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl4.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_StatusKeluarKamar"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Status Keluar", vbCritical
            sp_StatusKeluarKamar = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_KondisiKeluar(f_Status As String) As Boolean
    sp_KondisiKeluar = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKondisiKeluar", adChar, adParamInput, 2, txtKdKondisiKeluar.Text)
        .Parameters.Append .CreateParameter("KondisiKeluar", adVarChar, adParamInput, 50, Trim(txtKondisiKeluar.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal7.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNamaExternal7.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl7.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_KondisiKeluar"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam proses data kondisi keluar", vbCritical
            sp_KondisiKeluar = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    Select Case sstDataPendukung.Tab
        Case 0  'Status Pulang
            If Periksa("text", txtStatusPulang, "Nama status pulang kosong") = False Then Exit Sub
            If sp_StatusPulang("A") = False Then Exit Sub
        Case 1  'Kondisi Pulang
            If Periksa("text", txtKondisiPulang, "Nama kondisi pulang kosong") = False Then Exit Sub
            If sp_KondisiPulang("A") = False Then Exit Sub
        Case 2  'Kesadaran
            If Periksa("text", txtKesadaran, "Nama Kesadaran kosong") = False Then Exit Sub
            If sp_Kesadaran("A") = False Then Exit Sub
        Case 3   'Jenis Diagnosa
            If Periksa("text", txtJenisDiagnosa, "Nama jenis diagnosa kosong") = False Then Exit Sub
            If sp_JenisDiagnosa("A") = False Then Exit Sub
        Case 4  'Status Keluar Kamar
            If Periksa("text", txtStatusKeluar, "Nama status keluar kamar kosong") = False Then Exit Sub
            If sp_StatusKeluarKamar("A") = False Then Exit Sub
        Case 5  'Rujukan Asal
            If Periksa("text", txtRujukanAsal, "Nama rujukan asal kosong") = False Then Exit Sub
            If sp_RujukanAsal("A") = False Then Exit Sub
        Case 6  'Detail Rujukan Asal
            If Periksa("text", txtDetailRujukanAsal, "Detail rujukan asal kosong") = False Then Exit Sub
            If Periksa("datacombo", dcRujukanAsal, "Rujukan asal kosong") = False Then Exit Sub
            If sp_DetailRujukanAsal("A") = False Then Exit Sub
        Case 7  'Kondisi Keluar Kamar
            If Periksa("text", txtKondisiKeluar, "Nama kondisi keluar kamar kosong") = False Then Exit Sub
            If sp_KondisiKeluar("A") = False Then Exit Sub
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

Private Sub dcRujukanAsal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcRujukanAsal.MatchedWithList = True Then txtDetailRujukanAsal.SetFocus
        strSQL = "SELECT KdRujukanAsal, RujukanAsal FROM dbo.RujukanAsal where StatusEnabled='1' and (RujukanAsal LIKE '%" & dcRujukanAsal.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcRujukanAsal.Text = ""
            Exit Sub
        End If
        dcRujukanAsal.BoundText = rs(0).value
        dcRujukanAsal.Text = rs(1).value
    End If
End Sub

Private Sub dgDetailRujukanAsal_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDetailRujukanAsal
    WheelHook.WheelHook dgDetailRujukanAsal
End Sub

Private Sub dgDetailRujukanAsal_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    With dgDetailRujukanAsal
        txtKDDetailRujukanAsal.Text = .Columns(0)
        dcRujukanAsal.Text = .Columns(1)
        txtDetailRujukanAsal.Text = .Columns(2)
        txtKodeExternal6.Text = .Columns(3).value
        txtNamaExternal6.Text = .Columns(4).value
        If .Columns(5) = "" Then
            CheckStatusEnbl6.value = 0
        ElseIf .Columns(5) = 0 Then
            CheckStatusEnbl6.value = 0
        ElseIf .Columns(5) = 1 Then
            CheckStatusEnbl6.value = 1
        End If
    End With
End Sub

Private Sub dgJenisDiagnosa_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgJenisDiagnosa
    WheelHook.WheelHook dgJenisDiagnosa
End Sub

Private Sub dgJenisDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJenisDiagnosa.SetFocus
End Sub

Private Sub dgKesadaran_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKesadaran
    WheelHook.WheelHook dgKesadaran
End Sub

Private Sub dgKesadaran_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKesadaran.SetFocus
End Sub

Private Sub dgKondisiKeluar_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKondisiKeluar
    WheelHook.WheelHook dgKondisiKeluar
End Sub

Private Sub dgKondisiPulang_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKondisiPulang
    WheelHook.WheelHook dgKondisiPulang
End Sub

Private Sub dgKondisiPulang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKondisiPulang.SetFocus
End Sub

Private Sub dgRujukanAsal_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgRujukanAsal
    WheelHook.WheelHook dgRujukanAsal
End Sub

Private Sub dgRujukanAsal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRujukanAsal.SetFocus
End Sub

Private Sub dgStatusKeluar_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgStatusKeluar
    WheelHook.WheelHook dgStatusKeluar
End Sub

Private Sub dgStatusKeluar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtStatusKeluar.SetFocus
End Sub

Private Sub dgStatusPulang_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgStatusPulang
    WheelHook.WheelHook dgStatusPulang
End Sub

Private Sub dgStatusPulang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtStatusPulang.SetFocus
End Sub

Private Sub dgStatusPulang_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgStatusPulang.ApproxCount = 0 Then Exit Sub
    txtKdStatusPulang.Text = dgStatusPulang.Columns(0).value
    txtStatusPulang.Text = dgStatusPulang.Columns(1).value
    txtKodeExternal.Text = dgStatusPulang.Columns(3).value
    txtNamaExternal.Text = dgStatusPulang.Columns(4).value
    If dgStatusPulang.Columns(5) = "" Then
        CheckStatusEnbl.value = 0
    ElseIf dgStatusPulang.Columns(5) = 0 Then
        CheckStatusEnbl.value = 0
    ElseIf dgStatusPulang.Columns(5) = 1 Then
        CheckStatusEnbl.value = 1
    End If
End Sub

Private Sub dgKondisiPulang_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgKondisiPulang.ApproxCount = 0 Then Exit Sub
    txtKdKondisiPulang.Text = dgKondisiPulang.Columns(0).value
    txtKondisiPulang.Text = dgKondisiPulang.Columns(1).value
    txtSingkatan.Text = dgKondisiPulang.Columns(2).value
    txtKodeExternal1.Text = dgKondisiPulang.Columns(4).value
    txtNamaExternal1.Text = dgKondisiPulang.Columns(5).value
    If dgKondisiPulang.Columns(6) = "" Then
        CheckStatusEnbl1.value = 0
    ElseIf dgKondisiPulang.Columns(6) = 0 Then
        CheckStatusEnbl1.value = 0
    ElseIf dgKondisiPulang.Columns(6) = 1 Then
        CheckStatusEnbl1.value = 1
    End If
End Sub

Private Sub dgKesadaran_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgKesadaran.ApproxCount = 0 Then Exit Sub
    txtKdKesadaran.Text = dgKesadaran.Columns(0).value
    txtKesadaran.Text = dgKesadaran.Columns(1).value
    txtKodeExternal2.Text = dgKesadaran.Columns(2).value
    txtNamaExternal2.Text = dgKesadaran.Columns(3).value
    If dgKesadaran.Columns(4) = "" Then
        CheckStatusEnbl2.value = 0
    ElseIf dgKesadaran.Columns(4) = 0 Then
        CheckStatusEnbl2.value = 0
    ElseIf dgKesadaran.Columns(4) = 1 Then
        CheckStatusEnbl2.value = 1
    End If
End Sub

Private Sub dgJenisDiagnosa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgJenisDiagnosa.ApproxCount = 0 Then Exit Sub
    txtKdJenisDiagnosa.Text = dgJenisDiagnosa.Columns(0).value
    txtJenisDiagnosa.Text = dgJenisDiagnosa.Columns(1).value
    txtKodeExternal3.Text = dgJenisDiagnosa.Columns(3).value
    txtNamaExternal3.Text = dgJenisDiagnosa.Columns(4).value
    If dgJenisDiagnosa.Columns(5) = "" Then
        CheckStatusEnbl3.value = 0
    ElseIf dgJenisDiagnosa.Columns(5) = 0 Then
        CheckStatusEnbl3.value = 0
    ElseIf dgJenisDiagnosa.Columns(5) = 1 Then
        CheckStatusEnbl3.value = 1
    End If
End Sub

Private Sub dgRujukanAsal_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgRujukanAsal.ApproxCount = 0 Then Exit Sub
    txtKdRujukanAsal.Text = dgRujukanAsal.Columns(0).value
    txtRujukanAsal.Text = dgRujukanAsal.Columns(1).value
    txtSingkatanRujukan.Text = dgRujukanAsal.Columns(2).value
    txtKodeExternal5.Text = dgRujukanAsal.Columns(4).value
    txtNamaExternal5.Text = dgRujukanAsal.Columns(5).value
    If dgRujukanAsal.Columns(6) = "" Then
        CheckStatusEnbl5.value = 0
    ElseIf dgRujukanAsal.Columns(6) = 0 Then
        CheckStatusEnbl5.value = 0
    ElseIf dgRujukanAsal.Columns(6) = 1 Then
        CheckStatusEnbl5.value = 1
    End If
End Sub

Private Sub dgStatusKeluar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgStatusKeluar.ApproxCount = 0 Then Exit Sub
    txtKdStatusKeluar.Text = dgStatusKeluar.Columns(0).value
    txtStatusKeluar.Text = dgStatusKeluar.Columns(1).value
    txtKodeExternal4.Text = dgStatusKeluar.Columns(3).value
    txtNamaExternal4.Text = dgStatusKeluar.Columns(4).value
    If dgStatusKeluar.Columns(5) = "" Then
        CheckStatusEnbl4.value = 0
    ElseIf dgStatusKeluar.Columns(5) = 0 Then
        CheckStatusEnbl4.value = 0
    ElseIf dgStatusKeluar.Columns(5) = 1 Then
        CheckStatusEnbl4.value = 1
    End If
End Sub

Private Sub dgKondisiKeluar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgKondisiKeluar.ApproxCount = 0 Then Exit Sub
    txtKdKondisiKeluar.Text = dgKondisiKeluar.Columns(0).value
    txtKondisiKeluar.Text = dgKondisiKeluar.Columns(1).value
    txtKodeExternal7.Text = dgKondisiKeluar.Columns(3).value
    txtNamaExternal7.Text = dgKondisiKeluar.Columns(4).value
    If dgKondisiKeluar.Columns(5) = "" Then
        CheckStatusEnbl7.value = 0
    ElseIf dgKondisiKeluar.Columns(5) = 0 Then
        CheckStatusEnbl7.value = 0
    ElseIf dgKondisiKeluar.Columns(5) = 1 Then
        CheckStatusEnbl7.value = 1
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    sstDataPendukung.Tab = 0
    Call subLoadDcSource
    Call subLoadGridSource
    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()
    On Error GoTo errLoad
    Select Case sstDataPendukung.Tab
        Case 0 'Status Pulang
            strSQL = "select * from StatusPulang"
            Call msubRecFO(rs, strSQL)
            Set dgStatusPulang.DataSource = rs
            dgStatusPulang.Columns(0).Width = 1000
            dgStatusPulang.Columns(1).Width = 4000
            dgStatusPulang.Columns(0).Caption = "Kd. Status Pulang"
            dgStatusPulang.Columns(1).Caption = "Status Pulang"

        Case 1 'Kondisi Pulang
            strSQL = "select * from KondisiPulang"
            Call msubRecFO(rs, strSQL)
            Set dgKondisiPulang.DataSource = rs
            dgKondisiPulang.Columns(0).Width = 1000
            dgKondisiPulang.Columns(0).Caption = "Kd. Kondisi Pulang"
            dgKondisiPulang.Columns(1).Width = 3000
            dgKondisiPulang.Columns(2).Width = 1000

        Case 2  'Kesadaran
            strSQL = "select * from Kesadaran"
            Call msubRecFO(rs, strSQL)
            Set dgKesadaran.DataSource = rs
            dgKesadaran.Columns(0).Width = 1000
            dgKesadaran.Columns(0).Caption = "Kd. Kesadaran"
            dgKesadaran.Columns(1).Width = 4000

        Case 3 'Jenis Diagnosa
            strSQL = "select * from JenisDiagnosa"
            Call msubRecFO(rs, strSQL)
            Set dgJenisDiagnosa.DataSource = rs
            dgJenisDiagnosa.Columns(0).Width = 1000
            dgJenisDiagnosa.Columns(0).Caption = "Kd. Jenis Kesadaran"
            dgJenisDiagnosa.Columns(1).Width = 4000

        Case 4 'Status Keluar Kamar
            strSQL = "select * from StatusKeluarKamar"
            Call msubRecFO(rs, strSQL)
            Set dgStatusKeluar.DataSource = rs
            dgStatusKeluar.Columns(0).Width = 1500
            dgStatusKeluar.Columns(0).Caption = "Kd. Status Keluar Kamar"
            dgStatusKeluar.Columns(1).Width = 4000
        Case 5 'Rujukan Asal
            strSQL = "select * from RujukanAsal"
            Call msubRecFO(rs, strSQL)
            Set dgRujukanAsal.DataSource = rs
            dgRujukanAsal.Columns(0).Width = 1100
            dgRujukanAsal.Columns(0).Caption = "Kd. Rujukan Asal"
            dgRujukanAsal.Columns(1).Width = 4000
            dgRujukanAsal.Columns(1).Width = 2000
        Case 6 'Detail Rujukan Asal
            strSQL = "SELECT  dbo.DetailRujukanAsal.KdDetailRujukanAsal,dbo.RujukanAsal.RujukanAsal, dbo.DetailRujukanAsal.DetailRujukanAsal," & _
            "  dbo.DetailRujukanAsal.KodeExternal,dbo.DetailRujukanAsal.NamaExternal,dbo.DetailRujukanAsal.StatusEnabled FROM  dbo.RujukanAsal INNER JOIN" & _
            " dbo.DetailRujukanAsal ON dbo.RujukanAsal.KdRujukanAsal = dbo.DetailRujukanAsal.KdRujukanAsal"
            Call msubRecFO(rs, strSQL)
            Set dgDetailRujukanAsal.DataSource = rs
            With dgDetailRujukanAsal
                .Columns(0).Width = 1150
                .Columns(0).Caption = "Kd. Detail Rujukan Asal"
                .Columns(1).Width = 2200
                .Columns(2).Width = 3300
            End With
        Case 7 'Kondisi Keluar Kamar
            strSQL = "select * from KondisiKeluar"
            Call msubRecFO(rs, strSQL)
            Set dgKondisiKeluar.DataSource = rs
            dgKondisiKeluar.Columns(0).Width = 1000
            dgKondisiKeluar.Columns(0).Caption = "Kd. Kondisi Keluar"
            dgKondisiKeluar.Columns(1).Width = 4000
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad
    strSQL = "SELECT     KdRujukanAsal, RujukanAsal FROM dbo.RujukanAsal where StatusEnabled='1' "
    Call msubDcSource(dcRujukanAsal, rs, strSQL)
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Sub clear()
    On Error Resume Next
    Select Case sstDataPendukung.Tab
        Case 0 'Status Pulang
            txtKdStatusPulang.Text = ""
            txtStatusPulang.Text = ""
            txtStatusPulang.SetFocus
            txtKodeExternal.Text = ""
            txtNamaExternal.Text = ""
            CheckStatusEnbl.value = 1

        Case 1 'Kondisi Pulang
            txtKdKondisiPulang.Text = ""
            txtKondisiPulang.Text = ""
            txtSingkatan.Text = ""
            txtKondisiPulang.SetFocus
            txtKodeExternal1.Text = ""
            txtNamaExternal1.Text = ""
            CheckStatusEnbl1.value = 1

        Case 2 'Kesadaran
            txtKdKesadaran.Text = ""
            txtKesadaran.Text = ""
            txtKesadaran.SetFocus
            txtKodeExternal2.Text = ""
            txtNamaExternal2.Text = ""
            CheckStatusEnbl2.value = 1

        Case 3 'Jenis Diagnosa
            txtKdJenisDiagnosa.Text = ""
            txtJenisDiagnosa.Text = ""
            txtJenisDiagnosa.SetFocus
            txtKodeExternal3.Text = ""
            txtNamaExternal3.Text = ""
            CheckStatusEnbl3.value = 1

        Case 4 'Status Keluar Kamar
            txtKdStatusKeluar.Text = ""
            txtStatusKeluar.Text = ""
            txtStatusKeluar.SetFocus
            txtKodeExternal4.Text = ""
            txtNamaExternal4.Text = ""
            CheckStatusEnbl4.value = 1

        Case 5 'Rujukan Asal
            txtKdRujukanAsal.Text = ""
            txtRujukanAsal.Text = ""
            txtSingkatanRujukan.Text = ""
            txtRujukanAsal.SetFocus
            txtKodeExternal5.Text = ""
            txtNamaExternal5.Text = ""
            CheckStatusEnbl5.value = 1

        Case 6 'Detail Rujukan Asal
            txtKDDetailRujukanAsal.Text = ""
            dcRujukanAsal.Text = ""
            txtDetailRujukanAsal.Text = ""
            txtDetailRujukanAsal.SetFocus
            txtKodeExternal6.Text = ""
            txtNamaExternal6.Text = ""
            CheckStatusEnbl6.value = 1

        Case 4 'Kondisi Keluar Kamar
            txtKdKondisiKeluar.Text = ""
            txtKondisiKeluar.Text = ""
            txtKondisiKeluar.SetFocus
            txtKodeExternal7.Text = ""
            txtNamaExternal7.Text = ""
            CheckStatusEnbl7.value = 1
    End Select
End Sub

Private Sub sstDataPendukung_Click(PreviousTab As Integer)
    Call clear
    Call subLoadGridSource

End Sub

Private Sub sstDataPendukung_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    If KeyAscii = 13 Then
        Select Case sstDataPendukung.Tab
            Case 0 'Status Pulang
                txtStatusPulang.SetFocus
            Case 1 'Kondisi Pulang
                txtKondisiPulang.SetFocus
            Case 2 'Kesadaran
                txtKesadaran.SetFocus
            Case 3 'Jenis Diagnosa
                txtJenisDiagnosa.SetFocus
            Case 4 'Status Keluar Kamar
                txtStatusKeluar.SetFocus
            Case 5 'Rujukan Asal
                txtRujukanAsal.SetFocus
            Case 6 'Detail Rujukan Asal
                txtDetailRujukanAsal.SetFocus
            Case 7 'Kondisi Keluar Kamar
                txtKondisiKeluar.SetFocus
        End Select
    End If
errLoad:
End Sub

Private Sub txtDetailRujukanAsal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal6.SetFocus
End Sub

Private Sub txtJenisDiagnosa_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            txtKodeExternal3.SetFocus
    End Select
End Sub

Private Sub txtJenisDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal3.SetFocus
End Sub

Private Sub txtJenisDiagnosa_LostFocus()
    txtJenisDiagnosa.Text = Trim(StrConv(txtJenisDiagnosa.Text, vbProperCase))
End Sub

Private Sub txtKesadaran_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            txtKodeExternal2.SetFocus
    End Select
End Sub

Private Sub txtKesadaran_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal2.SetFocus
End Sub

Private Sub txtKesadaran_LostFocus()
    txtKesadaran.Text = Trim(StrConv(txtKesadaran.Text, vbProperCase))
End Sub

Private Sub txtKondisiKeluar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal7.SetFocus
End Sub

Private Sub txtKondisiPulang_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            txtSingkatan.SetFocus
    End Select
End Sub

Private Sub txtKondisiPulang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtSingkatan.SetFocus
End Sub

Private Sub txtKondisiPulang_LostFocus()
    txtKondisiPulang.Text = Trim(StrConv(txtKondisiPulang.Text, vbProperCase))
End Sub

Private Sub txtRujukanAsal_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            txtSingkatanRujukan.SetFocus
    End Select
End Sub

Private Sub txtRujukanAsal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtSingkatanRujukan.SetFocus
End Sub

Private Sub txtRujukanAsal_LostFocus()
    txtRujukanAsal.Text = Trim(StrConv(txtRujukanAsal.Text, vbProperCase))
End Sub

Private Sub txtSingkatan_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            txtKodeExternal1.SetFocus
    End Select
End Sub

Private Sub txtSingkatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal1.SetFocus
End Sub

Private Sub txtSingkatanRujukan_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            txtKodeExternal5.SetFocus
    End Select
End Sub

Private Sub txtSingkatanRujukan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal5.SetFocus
End Sub

Private Sub txtStatusKeluar_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            txtKodeExternal4.SetFocus
    End Select
End Sub

Private Sub txtStatusKeluar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal4.SetFocus
End Sub

Private Sub txtStatusKeluar_LostFocus()
    txtStatusKeluar.Text = Trim(StrConv(txtStatusKeluar.Text, vbProperCase))
End Sub

Private Sub txtStatusPulang_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            txtKodeExternal.SetFocus
    End Select
End Sub

Private Sub txtStatusPulang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal.SetFocus
End Sub

Private Sub txtStatusPulang_LostFocus()
    txtStatusPulang.Text = Trim(StrConv(txtStatusPulang.Text, vbProperCase))
End Sub

Private Sub txtKodeExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal.SetFocus
End Sub

Private Sub CheckStatusEnbl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl.SetFocus
End Sub

Private Sub txtKodeExternal1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal1.SetFocus
End Sub

Private Sub CheckStatusEnbl1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaExternal1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl1.SetFocus
End Sub

Private Sub txtKodeExternal2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal2.SetFocus
End Sub

Private Sub CheckStatusEnbl2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaExternal2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl2.SetFocus
End Sub

Private Sub txtKodeExternal3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal3.SetFocus
End Sub

Private Sub CheckStatusEnbl3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaExternal3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl3.SetFocus
End Sub

Private Sub txtKodeExternal4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal4.SetFocus
End Sub

Private Sub CheckStatusEnbl4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaExternal4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl4.SetFocus
End Sub

Private Sub txtKodeExternal5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal5.SetFocus
End Sub

Private Sub CheckStatusEnbl5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaExternal5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl5.SetFocus
End Sub

Private Sub txtKodeExternal6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal6.SetFocus
End Sub

Private Sub CheckStatusEnbl6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaExternal6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl6.SetFocus
End Sub

Private Sub txtKodeExternal7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal7.SetFocus
End Sub

Private Sub CheckStatusEnbl7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaExternal7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl7.SetFocus
End Sub

