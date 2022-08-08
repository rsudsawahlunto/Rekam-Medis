VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmMasterDataPendukung 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Setting Pelayanan"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMasterDataPendukung.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8175
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   6000
      Width           =   1215
   End
   Begin TabDlg.SSTab sstDataPendukung 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   7
      Tab             =   5
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   -2147483638
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
      TabPicture(0)   =   "frmMasterDataPendukung.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame8"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Kondisi Pulang"
      TabPicture(1)   =   "frmMasterDataPendukung.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Kesadaran"
      TabPicture(2)   =   "frmMasterDataPendukung.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame10"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Jenis Diagnosa"
      TabPicture(3)   =   "frmMasterDataPendukung.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame12"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Status Keluar"
      TabPicture(4)   =   "frmMasterDataPendukung.frx":0D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame14"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Rujukan Asal"
      TabPicture(5)   =   "frmMasterDataPendukung.frx":0D56
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Frame13"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Detail Rujukan Asal"
      TabPicture(6)   =   "frmMasterDataPendukung.frx":0D72
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame1"
      Tab(6).ControlCount=   1
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
         Height          =   3735
         Left            =   -74760
         TabIndex        =   49
         Top             =   840
         Width           =   7335
         Begin MSDataListLib.DataCombo dcRujukanAsal 
            Height          =   330
            Left            =   4320
            TabIndex        =   26
            Top             =   360
            Width           =   2775
            _ExtentX        =   4895
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
            TabIndex        =   27
            Top             =   840
            Width           =   5415
         End
         Begin VB.TextBox txtKDDetailRujukanAsal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1680
            MaxLength       =   8
            TabIndex        =   25
            Top             =   360
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid dgDetailRujukanAsal 
            Height          =   2295
            Left            =   120
            TabIndex        =   28
            Top             =   1320
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4048
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Rujukan Asal"
            Height          =   210
            Left            =   3240
            TabIndex        =   52
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Rujukan Asal"
            Height          =   210
            Left            =   240
            TabIndex        =   51
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Kode "
            Height          =   210
            Left            =   240
            TabIndex        =   50
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
         Height          =   3735
         Left            =   -74760
         TabIndex        =   46
         Top             =   840
         Width           =   7335
         Begin VB.TextBox txtKdStatusKeluar 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   18
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtStatusKeluar 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   19
            Text            =   "50"
            Top             =   840
            Width           =   4935
         End
         Begin MSDataGridLib.DataGrid dgStatusKeluar 
            Height          =   2295
            Left            =   120
            TabIndex        =   20
            Top             =   1320
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            FormatLocked    =   -1  'True
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
               Caption         =   "Kode Status Keluar"
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
               Caption         =   "Status Keluar"
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
            AutoSize        =   -1  'True
            Caption         =   "Kode Status Keluar "
            Height          =   210
            Left            =   240
            TabIndex        =   48
            Top             =   480
            Width           =   1620
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Status Keluar"
            Height          =   210
            Left            =   240
            TabIndex        =   47
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
         Height          =   3735
         Left            =   240
         TabIndex        =   42
         Top             =   840
         Width           =   7335
         Begin VB.TextBox txtKdRujukanAsal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   21
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtRujukanAsal 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   22
            Text            =   "30"
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox txtSingkatanRujukan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6480
            MaxLength       =   5
            TabIndex        =   23
            Top             =   840
            Width           =   735
         End
         Begin MSDataGridLib.DataGrid dgRujukanAsal 
            Height          =   2295
            Left            =   120
            TabIndex        =   24
            Top             =   1320
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            FormatLocked    =   -1  'True
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   "Kode"
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
               Caption         =   "Rujukan Asal"
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
            BeginProperty Column02 
               DataField       =   ""
               Caption         =   "Singkatan"
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
               BeginProperty Column02 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Kode Rujukan Asal"
            Height          =   210
            Left            =   240
            TabIndex        =   45
            Top             =   480
            Width           =   1500
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Rujukan Asal"
            Height          =   210
            Left            =   240
            TabIndex        =   44
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Singkatan"
            Height          =   210
            Left            =   5520
            TabIndex        =   43
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
         Height          =   3855
         Left            =   -74760
         TabIndex        =   38
         Top             =   780
         Width           =   7335
         Begin VB.TextBox txtKdStatusPulang 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   1
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtStatusPulang 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   2
            Top             =   840
            Width           =   4815
         End
         Begin MSDataGridLib.DataGrid dgStatusPulang 
            Height          =   2415
            Left            =   120
            TabIndex        =   3
            Top             =   1320
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4260
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            FormatLocked    =   -1  'True
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
               Caption         =   "Kode Status Pulang"
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
               Caption         =   "Status Pulang"
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
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Kode Status Pulang"
            Height          =   210
            Left            =   240
            TabIndex        =   40
            Top             =   480
            Width           =   1605
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Status Pulang"
            Height          =   210
            Left            =   240
            TabIndex        =   39
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
         Height          =   3855
         Left            =   -74760
         TabIndex        =   35
         Top             =   780
         Width           =   7335
         Begin VB.TextBox txtSingkatan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6600
            MaxLength       =   5
            TabIndex        =   10
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtKondisiPulang 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   9
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox txtKdKondisiPulang 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   8
            Top             =   480
            Width           =   735
         End
         Begin MSDataGridLib.DataGrid dgKondisiPulang 
            Height          =   2415
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4260
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            FormatLocked    =   -1  'True
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   "Kode"
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
               Caption         =   "Kondisi"
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
            BeginProperty Column02 
               DataField       =   ""
               Caption         =   "Singkatan"
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
               BeginProperty Column02 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Singkatan"
            Height          =   210
            Left            =   5640
            TabIndex        =   41
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Kondisi Pulang"
            Height          =   210
            Left            =   240
            TabIndex        =   37
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Kode Kondisi Pulang"
            Height          =   210
            Left            =   240
            TabIndex        =   36
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
         Height          =   3855
         Left            =   -74760
         TabIndex        =   32
         Top             =   780
         Width           =   7335
         Begin VB.TextBox txtKesadaran 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   13
            Top             =   840
            Width           =   5295
         End
         Begin VB.TextBox txtKdKesadaran 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   12
            Top             =   480
            Width           =   735
         End
         Begin MSDataGridLib.DataGrid dgKesadaran 
            Height          =   2415
            Left            =   120
            TabIndex        =   14
            Top             =   1320
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4260
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            FormatLocked    =   -1  'True
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
               Caption         =   "Kode"
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
               Caption         =   "Kesadaran"
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
            AutoSize        =   -1  'True
            Caption         =   "Kesadaran"
            Height          =   210
            Left            =   240
            TabIndex        =   34
            Top             =   840
            Width           =   825
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Kode Kesadaran"
            Height          =   210
            Left            =   240
            TabIndex        =   33
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
         Height          =   3855
         Left            =   -74760
         TabIndex        =   29
         Top             =   780
         Width           =   7335
         Begin VB.TextBox txtJenisDiagnosa 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2160
            MaxLength       =   30
            TabIndex        =   16
            Top             =   840
            Width           =   4935
         End
         Begin VB.TextBox txtKdJenisDiagnosa 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   15
            Top             =   480
            Width           =   735
         End
         Begin MSDataGridLib.DataGrid dgJenisDiagnosa 
            Height          =   2415
            Left            =   120
            TabIndex        =   17
            Top             =   1320
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4260
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            FormatLocked    =   -1  'True
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
               Caption         =   "Kode Jenis Diagnosa"
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
               Caption         =   "Jenis Diagnosa"
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
            AutoSize        =   -1  'True
            Caption         =   "Jenis Diagnosa"
            Height          =   210
            Left            =   240
            TabIndex        =   31
            Top             =   840
            Width           =   1170
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Kode Jenis Diagnosa"
            Height          =   210
            Left            =   240
            TabIndex        =   30
            Top             =   480
            Width           =   1650
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   53
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
      Left            =   6480
      Picture         =   "frmMasterDataPendukung.frx":0D8E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterDataPendukung.frx":1B16
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterDataPendukung.frx":3174
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmMasterDataPendukung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    End Select

    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
    Call cmdBatal_Click

    Exit Sub
errLoad:
    MsgBox "Panghapusan data gagal", vbCritical, "Informasi"
End Sub

Private Function sp_StatusPulang(f_Status As String) As Boolean
    sp_StatusPulang = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdStatusPulang", adChar, adParamInput, 2, txtKdStatusPulang.Text)
        .Parameters.Append .CreateParameter("StatusPulang", adVarChar, adParamInput, 50, Trim(txtStatusPulang.Text))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_StatusPulang"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Status Pulang", vbCritical
            sp_StatusPulang = False
        Else
            Call Add_HistoryLoginActivity("AUD_StatusPulang")
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
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_KondisiPulang"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Kondisi Pulang", vbCritical
            sp_KondisiPulang = False
        Else
            Call Add_HistoryLoginActivity("AUD_KondisiPulang")
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
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_Kesadaran"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Kesadaran", vbCritical
            sp_Kesadaran = False
        Else
            Call Add_HistoryLoginActivity("AUD_Kesadaran")
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
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_JenisDiagnosa"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Jenis Diagnosa", vbCritical
            sp_JenisDiagnosa = False
        Else
            Call Add_HistoryLoginActivity("AUD_JenisDiagnosa")
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
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_RujukanAsal"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Rujukan Asal", vbCritical
            sp_RujukanAsal = False
        Else
            Call Add_HistoryLoginActivity("AUD_RujukanAsal")
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
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_DetailRujukanAsal"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Detail Rujukan Asal", vbCritical
            subSimpanDetailRujukanAsal = False
        Else
            Call Add_HistoryLoginActivity("AUD_DetailRujukanAsal")
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
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_StatusKeluarKamar"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Status Keluar", vbCritical
            sp_StatusKeluarKamar = False
        Else
            Call Add_HistoryLoginActivity("AUD_StatusKeluarKamar")
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
        strSQL = "SELECT KdRujukanAsal, RujukanAsal FROM dbo.RujukanAsal where (RujukanAsal LIKE '%" & dcRujukanAsal.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcRujukanAsal.Text = ""
            Exit Sub
        End If
        dcRujukanAsal.BoundText = rs(0).value
        dcRujukanAsal.Text = rs(1).value
    End If
End Sub

Private Sub dgDetailRujukanAsal_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    With dgDetailRujukanAsal
        txtKDDetailRujukanAsal.Text = .Columns(0)
        dcRujukanAsal.Text = .Columns(1)
        txtDetailRujukanAsal.Text = .Columns(2)
    End With
End Sub

Private Sub dgJenisDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJenisDiagnosa.SetFocus
End Sub

Private Sub dgKesadaran_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKesadaran.SetFocus
End Sub

Private Sub dgKondisiPulang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKondisiPulang.SetFocus
End Sub

Private Sub dgRujukanAsal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRujukanAsal.SetFocus
End Sub

Private Sub dgStatusKeluar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtStatusKeluar.SetFocus
End Sub

Private Sub dgStatusPulang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtStatusPulang.SetFocus
End Sub

Private Sub dgStatusPulang_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgStatusPulang.ApproxCount = 0 Then Exit Sub
    txtKdStatusPulang.Text = dgStatusPulang.Columns(0).value
    txtStatusPulang.Text = dgStatusPulang.Columns(1).value
End Sub

Private Sub dgKondisiPulang_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgKondisiPulang.ApproxCount = 0 Then Exit Sub
    txtKdKondisiPulang.Text = dgKondisiPulang.Columns(0).value
    txtKondisiPulang.Text = dgKondisiPulang.Columns(1).value
    txtSingkatan.Text = dgKondisiPulang.Columns(2).value
End Sub

Private Sub dgKesadaran_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgKesadaran.ApproxCount = 0 Then Exit Sub
    txtKdKesadaran.Text = dgKesadaran.Columns(0).value
    txtKesadaran.Text = dgKesadaran.Columns(1).value
End Sub

Private Sub dgJenisDiagnosa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgJenisDiagnosa.ApproxCount = 0 Then Exit Sub
    txtKdJenisDiagnosa.Text = dgJenisDiagnosa.Columns(0).value
    txtJenisDiagnosa.Text = dgJenisDiagnosa.Columns(1).value
End Sub

Private Sub dgRujukanAsal_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgRujukanAsal.ApproxCount = 0 Then Exit Sub
    txtKdRujukanAsal.Text = dgRujukanAsal.Columns(0).value
    txtRujukanAsal.Text = dgRujukanAsal.Columns(1).value
    txtSingkatanRujukan.Text = dgRujukanAsal.Columns(2).value
End Sub

Private Sub dgStatusKeluar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgStatusKeluar.ApproxCount = 0 Then Exit Sub
    txtKdStatusKeluar.Text = dgStatusKeluar.Columns(0).value
    txtStatusKeluar.Text = dgStatusKeluar.Columns(1).value
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    sstDataPendukung.Tab = 0
    Call subLoadDcSource
    Call subLoadGridSource

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
            dgStatusPulang.Columns(0).DataField = rs(0).Name
            dgStatusPulang.Columns(1).DataField = rs(1).Name
            dgStatusPulang.Columns(0).Width = 2000
            dgStatusPulang.Columns(1).Width = 4500
            dgStatusPulang.ReBind

        Case 1 'Kondisi Pulang
            strSQL = "select * from KondisiPulang"
            Call msubRecFO(rs, strSQL)
            Set dgKondisiPulang.DataSource = rs
            dgKondisiPulang.Columns(0).DataField = rs(0).Name
            dgKondisiPulang.Columns(1).DataField = rs(1).Name
            dgKondisiPulang.Columns(2).DataField = rs(2).Name
            dgKondisiPulang.Columns(0).Width = 1500
            dgKondisiPulang.Columns(1).Width = 4000
            dgKondisiPulang.Columns(2).Width = 1000
            dgKondisiPulang.ReBind

        Case 2  'Kesadaran
            strSQL = "select * from Kesadaran"
            Call msubRecFO(rs, strSQL)
            Set dgKesadaran.DataSource = rs
            dgKesadaran.Columns(0).DataField = rs(0).Name
            dgKesadaran.Columns(1).DataField = rs(1).Name
            dgKesadaran.Columns(0).Width = 2000
            dgKesadaran.Columns(1).Width = 4500
            dgKesadaran.ReBind

        Case 3 'Jenis Diagnosa
            strSQL = "select * from JenisDiagnosa"
            Call msubRecFO(rs, strSQL)
            Set dgJenisDiagnosa.DataSource = rs
            dgJenisDiagnosa.Columns(0).DataField = rs(0).Name
            dgJenisDiagnosa.Columns(1).DataField = rs(1).Name
            dgJenisDiagnosa.Columns(0).Width = 2000
            dgJenisDiagnosa.Columns(1).Width = 4500
            dgJenisDiagnosa.ReBind

        Case 4 'Status Keluar Kamar
            strSQL = "select * from StatusKeluarKamar"
            Call msubRecFO(rs, strSQL)
            Set dgStatusKeluar.DataSource = rs
            dgStatusKeluar.Columns(0).DataField = rs(0).Name
            dgStatusKeluar.Columns(1).DataField = rs(1).Name
            dgStatusKeluar.Columns(0).Width = 2000
            dgStatusKeluar.Columns(1).Width = 4500
            dgStatusKeluar.ReBind
        Case 5 'Rujukan Asal
            strSQL = "select * from RujukanAsal"
            Call msubRecFO(rs, strSQL)
            Set dgRujukanAsal.DataSource = rs
            dgRujukanAsal.Columns(0).DataField = rs(0).Name
            dgRujukanAsal.Columns(1).DataField = rs(1).Name
            dgRujukanAsal.Columns(2).DataField = rs(2).Name
            dgRujukanAsal.Columns(0).Width = 1500
            dgRujukanAsal.Columns(1).Width = 4000
            dgRujukanAsal.Columns(2).Width = 1000
            dgRujukanAsal.ReBind
        Case 6 'Detail Rujukan Asal
            strSQL = "SELECT  dbo.DetailRujukanAsal.KdDetailRujukanAsal,dbo.RujukanAsal.RujukanAsal, dbo.DetailRujukanAsal.DetailRujukanAsal" & _
            " FROM  dbo.RujukanAsal INNER JOIN" & _
            " dbo.DetailRujukanAsal ON dbo.RujukanAsal.KdRujukanAsal = dbo.DetailRujukanAsal.KdRujukanAsal"
            Call msubRecFO(rs, strSQL)
            Set dgDetailRujukanAsal.DataSource = rs
            With dgDetailRujukanAsal
                .Columns(0).Width = 1000
                .Columns(1).Width = 2200
                .Columns(2).Width = 3300
                .Columns(0).Caption = "Kode"
                .Columns(1).Caption = "Detail Rujukan Asal"
                .Columns(2).Caption = "Singkatan"
            End With
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad
    strSQL = "SELECT     KdRujukanAsal, RujukanAsal FROM         dbo.RujukanAsal"
    Call msubDcSource(dcRujukanAsal, rs, strSQL)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub clear()
    On Error Resume Next
    Select Case sstDataPendukung.Tab
        Case 0 'Status Pulang
            txtKdStatusPulang.Text = ""
            txtStatusPulang.Text = ""
            txtStatusPulang.SetFocus

        Case 1 'Kondisi Pulang
            txtKdKondisiPulang.Text = ""
            txtKondisiPulang.Text = ""
            txtSingkatan.Text = ""
            txtKondisiPulang.SetFocus

        Case 2 'Kesadaran
            txtKdKesadaran.Text = ""
            txtKesadaran.Text = ""
            txtKesadaran.SetFocus

        Case 3 'Jenis Diagnosa
            txtKdJenisDiagnosa.Text = ""
            txtJenisDiagnosa.Text = ""
            txtJenisDiagnosa.SetFocus

        Case 4 'Status Keluar Kamar
            txtKdStatusKeluar.Text = ""
            txtStatusKeluar.Text = ""
            txtStatusKeluar.SetFocus

        Case 5 'Rujukan Asal
            txtKdRujukanAsal.Text = ""
            txtRujukanAsal.Text = ""
            txtSingkatanRujukan.Text = ""
            txtRujukanAsal.SetFocus

        Case 6 'Detail Rujukan Asal
            txtKDDetailRujukanAsal.Text = ""
            dcRujukanAsal.Text = ""
            txtDetailRujukanAsal.Text = ""
            txtDetailRujukanAsal.SetFocus
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
            Case 4 'Rujukan Asal
                txtRujukanAsal.SetFocus
            Case 5 'Status Keluar Kamar
                txtStatusKeluar.SetFocus
        End Select
    End If
errLoad:
End Sub

Private Sub txtDetailRujukanAsal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtJenisDiagnosa_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgJenisDiagnosa.SetFocus
    End Select
End Sub

Private Sub txtJenisDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtJenisDiagnosa_LostFocus()
    txtJenisDiagnosa.Text = Trim(StrConv(txtJenisDiagnosa.Text, vbProperCase))
End Sub

Private Sub txtKesadaran_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgKesadaran.SetFocus
    End Select
End Sub

Private Sub txtKesadaran_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKesadaran_LostFocus()
    txtKesadaran.Text = Trim(StrConv(txtKesadaran.Text, vbProperCase))
End Sub

Private Sub txtKondisiPulang_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgKondisiPulang.SetFocus
    End Select
End Sub

Private Sub txtKondisiPulang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKondisiPulang_LostFocus()
    txtKondisiPulang.Text = Trim(StrConv(txtKondisiPulang.Text, vbProperCase))
End Sub

Private Sub txtRujukanAsal_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgRujukanAsal.SetFocus
    End Select
End Sub

Private Sub txtRujukanAsal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtRujukanAsal_LostFocus()
    txtRujukanAsal.Text = Trim(StrConv(txtRujukanAsal.Text, vbProperCase))
End Sub

Private Sub txtStatusKeluar_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgStatusKeluar.SetFocus
    End Select
End Sub

Private Sub txtStatusKeluar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtStatusKeluar_LostFocus()
    txtStatusKeluar.Text = Trim(StrConv(txtStatusKeluar.Text, vbProperCase))
End Sub

Private Sub txtStatusPulang_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgStatusPulang.SetFocus
    End Select
End Sub

Private Sub txtStatusPulang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtStatusPulang_LostFocus()
    txtStatusPulang.Text = Trim(StrConv(txtStatusPulang.Text, vbProperCase))
End Sub

