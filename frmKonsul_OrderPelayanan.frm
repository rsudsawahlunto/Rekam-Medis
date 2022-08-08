VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKonsul_OrderPelayanan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Konsul dan Pesan Pelayanan Tindakan Pasien"
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKonsul_OrderPelayanan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   12615
   Begin VB.TextBox txtNamaForm 
      Height          =   435
      Left            =   0
      TabIndex        =   107
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
   End
   Begin MSComctlLib.ListView LvSirkuler 
      Height          =   375
      Left            =   9360
      TabIndex        =   100
      Top             =   4920
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nama Pemeriksa"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView LvInstrumen 
      Height          =   375
      Left            =   6480
      TabIndex        =   97
      Top             =   4920
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nama Pemeriksa"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvAsisten 
      Height          =   375
      Left            =   3480
      TabIndex        =   94
      Top             =   4920
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nama Pemeriksa"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid fgPerawatPerPelayanan 
      Height          =   1575
      Left            =   720
      TabIndex        =   77
      Top             =   -480
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2778
      _Version        =   393216
      FixedCols       =   0
   End
   Begin MSComctlLib.ListView lvPemeriksa 
      Height          =   1815
      Left            =   13800
      TabIndex        =   72
      Top             =   3360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nama Pemeriksa"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fraDokterPembantu 
      Caption         =   "Dokter Pendamping"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   13440
      TabIndex        =   75
      Top             =   5760
      Visible         =   0   'False
      Width           =   7455
      Begin MSDataGridLib.DataGrid dgDokterPembantu 
         Height          =   2295
         Left            =   360
         TabIndex        =   76
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
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
   Begin VB.Frame fraDokterAnestesi 
      Caption         =   "Dokter Anestesi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   13680
      TabIndex        =   73
      Top             =   5280
      Visible         =   0   'False
      Width           =   6735
      Begin MSDataGridLib.DataGrid dgDokterAnestesi 
         Height          =   2295
         Left            =   240
         TabIndex        =   74
         Top             =   360
         Width           =   6255
         _ExtentX        =   11033
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
   Begin VB.Frame fraDokter 
      Caption         =   "Data Dokter Pemeriksa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   14760
      TabIndex        =   70
      Top             =   3720
      Visible         =   0   'False
      Width           =   8895
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   2295
         Left            =   240
         TabIndex        =   71
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
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
   Begin VB.TextBox txtNoAntrian 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   10200
      MaxLength       =   15
      TabIndex        =   38
      Top             =   7680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtKdRuangan 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   7560
      MaxLength       =   15
      TabIndex        =   37
      Top             =   8040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid dgDokterTM 
      Height          =   2415
      Left            =   13200
      TabIndex        =   0
      Top             =   960
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4260
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   1
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   706
      TabCaption(0)   =   "Konsul dan Pesan Pelayanan"
      TabPicture(0)   =   "frmKonsul_OrderPelayanan.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtKdDokterTM"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtKdIsiTM"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.TextBox txtKdIsiTM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   121
         Top             =   4320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Pesanan Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         TabIndex        =   17
         Top             =   4800
         Width           =   12375
         Begin MSDataGridLib.DataGrid dgPelayananRS 
            Height          =   2055
            Left            =   2040
            TabIndex        =   105
            Top             =   555
            Visible         =   0   'False
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   3625
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
         Begin VB.Frame Frame2 
            Caption         =   "Status CITO"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6960
            TabIndex        =   84
            Top             =   0
            Width           =   1695
            Begin VB.OptionButton optCito 
               Caption         =   "Ya"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   86
               Top             =   240
               Width           =   615
            End
            Begin VB.OptionButton optCito 
               Caption         =   "Tidak"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   85
               Top             =   240
               Value           =   -1  'True
               Width           =   735
            End
         End
         Begin VB.CommandButton cmdHapus 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10560
            TabIndex        =   45
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtKuantitas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   6120
            MaxLength       =   3
            TabIndex        =   43
            Text            =   "1"
            Top             =   240
            Width           =   615
         End
         Begin VB.ComboBox dcJenisPemeriksaan 
            Enabled         =   0   'False
            Height          =   330
            ItemData        =   "frmKonsul_OrderPelayanan.frx":0CE6
            Left            =   17400
            List            =   "frmKonsul_OrderPelayanan.frx":0CE8
            TabIndex        =   34
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox txtRP 
            Height          =   495
            Left            =   9000
            TabIndex        =   24
            Top             =   2760
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtTotalDiscount 
            Height          =   495
            Left            =   6120
            TabIndex        =   23
            Top             =   2760
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtHarusDibayar 
            Height          =   495
            Left            =   4800
            TabIndex        =   22
            Top             =   2760
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtTanggunganRS 
            Height          =   495
            Left            =   3480
            TabIndex        =   21
            Top             =   2760
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtHutangPenjamin 
            Height          =   495
            Left            =   2040
            TabIndex        =   20
            Top             =   2760
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtTotalBiaya 
            Height          =   495
            Left            =   720
            TabIndex        =   19
            Top             =   2760
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtIsiTM 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2040
            TabIndex        =   18
            Top             =   240
            Width           =   3255
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgPelayanan 
            Height          =   1815
            Left            =   120
            TabIndex        =   42
            Top             =   720
            Width           =   12135
            _ExtentX        =   21405
            _ExtentY        =   3201
            _Version        =   393216
            Rows            =   50
            Cols            =   5
            FixedCols       =   0
            BackColorFixed  =   8577768
            BackColorBkg    =   16777215
            FocusRect       =   2
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.CommandButton cmdTambah 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8880
            TabIndex        =   120
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Nama Pelayanan"
            Height          =   255
            Left            =   480
            TabIndex        =   48
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Jumlah"
            Height          =   255
            Left            =   5520
            TabIndex        =   44
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Jenis Pemeriksaan"
            Height          =   255
            Index           =   0
            Left            =   17160
            TabIndex        =   25
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox txtKdDokterTM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         MaxLength       =   15
         TabIndex        =   16
         Top             =   4395
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame Frame4 
         Caption         =   "Data Konsul"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   12375
         Begin VB.CheckBox chkOperasiBersama 
            Caption         =   "Operasi Bersama"
            Height          =   255
            Left            =   9960
            TabIndex        =   119
            ToolTipText     =   "Operasi bersama dokter Operator 1 dan dokter Operator 2"
            Top             =   4200
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txtDokter 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   6840
            TabIndex        =   109
            Top             =   1080
            Width           =   5415
         End
         Begin VB.CheckBox chkDilayaniDokter 
            Caption         =   "Dokter Pemeriksa / Dokter 1"
            Height          =   255
            Left            =   9720
            TabIndex        =   108
            Top             =   840
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.TextBox txtNoOrder1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   0
            MaxLength       =   15
            TabIndex        =   78
            Top             =   0
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox chkDibayardimuka 
            Caption         =   "Langsung"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   6960
            TabIndex        =   69
            Top             =   1080
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Frame fraButton 
            Height          =   1575
            Left            =   120
            TabIndex        =   64
            Top             =   2520
            Width           =   12135
            Begin VB.CheckBox ChkSirkuler 
               Caption         =   "Sirkuler"
               Height          =   255
               Left            =   9120
               TabIndex        =   99
               Top             =   840
               Width           =   1935
            End
            Begin VB.TextBox txtNamaSirkuler 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   9120
               TabIndex        =   98
               Top             =   1125
               Width           =   2535
            End
            Begin VB.CheckBox chkInstrumen 
               Caption         =   "Instrumen"
               Height          =   255
               Left            =   6240
               TabIndex        =   96
               Top             =   840
               Width           =   2775
            End
            Begin VB.TextBox txtNamaInstrumen 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   6240
               TabIndex        =   95
               Top             =   1125
               Width           =   2655
            End
            Begin VB.CheckBox chkAsisten 
               Caption         =   "Asisten"
               Height          =   255
               Left            =   3240
               TabIndex        =   93
               Top             =   840
               Width           =   1935
            End
            Begin VB.TextBox txtNamaAsisten 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   3240
               TabIndex        =   92
               Top             =   1125
               Width           =   2775
            End
            Begin VB.CheckBox chkDokterResusitasi 
               Caption         =   "Dokter Resusitasi Neonatus"
               Height          =   255
               Left            =   120
               TabIndex        =   91
               Top             =   840
               Width           =   3135
            End
            Begin VB.TextBox txtDokterResusitasi 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   120
               TabIndex        =   90
               Top             =   1125
               Width           =   2895
            End
            Begin VB.CheckBox chkDokter2 
               Caption         =   "Dokter 2"
               Height          =   255
               Left            =   3240
               TabIndex        =   87
               Top             =   120
               Width           =   2055
            End
            Begin VB.TextBox txtDokter2 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   3240
               TabIndex        =   79
               Top             =   405
               Width           =   2775
            End
            Begin VB.TextBox txtDokterAnestesi 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   6240
               TabIndex        =   66
               Top             =   405
               Width           =   2655
            End
            Begin VB.TextBox txtDokterPembantu 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   9120
               TabIndex        =   65
               Top             =   405
               Width           =   2535
            End
            Begin MSDataListLib.DataCombo dcJenisOperasi 
               Height          =   330
               Left            =   120
               TabIndex        =   82
               Top             =   405
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               Text            =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Jenis Operasi"
               Height          =   210
               Left            =   120
               TabIndex        =   83
               Top             =   120
               Width           =   1050
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Dokter Anestesi"
               Height          =   240
               Index           =   1
               Left            =   6240
               TabIndex        =   68
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Dokter Pendamping"
               Height          =   240
               Index           =   3
               Left            =   9120
               TabIndex        =   67
               Top             =   120
               Width           =   1665
            End
         End
         Begin VB.Frame fraPDokter 
            Height          =   855
            Left            =   120
            TabIndex        =   63
            Top             =   2880
            Width           =   12135
            Begin VB.CheckBox chkDelegasi 
               Caption         =   "Di Delegasikan"
               Height          =   255
               Left            =   3240
               TabIndex        =   89
               Top             =   120
               Width           =   1935
            End
            Begin VB.TextBox txtDokterDelegasi 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   3240
               TabIndex        =   88
               Top             =   405
               Width           =   2775
            End
            Begin VB.TextBox txtNamaPerawat 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   6240
               TabIndex        =   81
               Top             =   405
               Width           =   2655
            End
            Begin VB.CheckBox chkPerawat 
               Caption         =   "Paramedis/Penata"
               Height          =   255
               Left            =   6240
               TabIndex        =   80
               Top             =   120
               Width           =   3495
            End
         End
         Begin VB.TextBox txtSex 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   8640
            TabIndex        =   61
            Top             =   480
            Width           =   1095
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
            Left            =   9840
            TabIndex        =   54
            Top             =   240
            Width           =   2415
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
               TabIndex        =   57
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
               TabIndex        =   56
               Top             =   240
               Width           =   375
            End
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
               TabIndex        =   55
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "thn"
               Height          =   210
               Left            =   550
               TabIndex        =   60
               Top             =   277
               Width           =   285
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "bln"
               Height          =   210
               Left            =   1350
               TabIndex        =   59
               Top             =   277
               Width           =   240
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "hr"
               Height          =   210
               Left            =   2130
               TabIndex        =   58
               Top             =   270
               Width           =   165
            End
         End
         Begin VB.TextBox txtNamaPasien 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   52
            Top             =   480
            Width           =   2775
         End
         Begin VB.CommandButton cmdHapusKonsul 
            Caption         =   "Hapus Konsul"
            Height          =   375
            Left            =   15120
            TabIndex        =   51
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtTglKonsul 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   3000
            MaxLength       =   250
            TabIndex        =   47
            Top             =   -120
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox txtNoUrutTMP 
            Height          =   315
            Left            =   8640
            TabIndex        =   40
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtKdRuanganTujuanTMP 
            Height          =   315
            Left            =   6480
            TabIndex        =   39
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdSimpanKonsul 
            Caption         =   "Simpan Konsul"
            Height          =   375
            Left            =   15120
            TabIndex        =   36
            Top             =   360
            Width           =   1815
         End
         Begin MSDataGridLib.DataGrid dgHistoryPelayanan 
            Height          =   1455
            Left            =   8160
            TabIndex        =   35
            Top             =   4920
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   2566
            _Version        =   393216
            HeadLines       =   1
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtNoCMTM 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2400
            MaxLength       =   15
            TabIndex        =   5
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtNoPendaftaranTM 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   4080
            MaxLength       =   15
            TabIndex        =   4
            Top             =   480
            Width           =   1575
         End
         Begin MSDataListLib.DataCombo dcInstalasi 
            Height          =   330
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcRuangan 
            Height          =   330
            Left            =   3360
            TabIndex        =   7
            Top             =   1080
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker dtpTglOrderTM 
            Height          =   330
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MM/dd/yyyy HH:mm:ss"
            Format          =   108462083
            UpDown          =   -1  'True
            CurrentDate     =   37760
         End
         Begin MSDataListLib.DataCombo dcDokterPerujuk 
            Height          =   330
            Left            =   8400
            TabIndex        =   50
            Top             =   1080
            Visible         =   0   'False
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            Text            =   ""
         End
         Begin MSDataGridLib.DataGrid dgKonsul 
            Height          =   1455
            Left            =   120
            TabIndex        =   9
            Top             =   4320
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   2566
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcTempatPerujuk 
            Height          =   330
            Left            =   2160
            TabIndex        =   111
            Top             =   1800
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcNamaPerujuk 
            Height          =   330
            Left            =   5520
            TabIndex        =   112
            Top             =   1800
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcDiagnosa 
            Height          =   330
            Left            =   8640
            TabIndex        =   113
            Top             =   1800
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComCtl2.DTPicker dtpTglRujuk 
            Height          =   330
            Left            =   120
            TabIndex        =   117
            Top             =   1800
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy HH:mm"
            Format          =   108462083
            UpDown          =   -1  'True
            CurrentDate     =   37813
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Tgl. Dirujuk"
            Height          =   210
            Index           =   2
            Left            =   120
            TabIndex        =   118
            Top             =   1560
            Width           =   930
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Nama Tempat Perujuk"
            Height          =   210
            Index           =   4
            Left            =   2160
            TabIndex        =   116
            Top             =   1560
            Width           =   1830
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Nama Perujuk"
            Height          =   210
            Index           =   7
            Left            =   5520
            TabIndex        =   115
            Top             =   1560
            Width           =   1125
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Diagnosa (Penyakit) Rujukan"
            Height          =   210
            Index           =   8
            Left            =   8640
            TabIndex        =   114
            Top             =   1560
            Width           =   2325
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Dokter Pemeriksa"
            Height          =   210
            Left            =   6840
            TabIndex        =   110
            Top             =   840
            Width           =   2145
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Kelamin"
            Height          =   210
            Left            =   8640
            TabIndex        =   62
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Pasien"
            Height          =   210
            Index           =   0
            Left            =   5760
            TabIndex        =   53
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            Height          =   210
            Index           =   5
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Ruangan Tujuan"
            Height          =   210
            Left            =   3360
            TabIndex        =   14
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Instalasi Tujuan"
            Height          =   210
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   1260
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Dokter Perujuk"
            Height          =   210
            Left            =   12000
            TabIndex        =   12
            Top             =   840
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Rekam Medis"
            Height          =   210
            Index           =   7
            Left            =   2400
            TabIndex        =   11
            Top             =   240
            Width           =   1545
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Pendaftaran"
            Height          =   210
            Index           =   6
            Left            =   4080
            TabIndex        =   10
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   120
         TabIndex        =   26
         Top             =   7440
         Width           =   12375
         Begin VB.CommandButton cmdCetak 
            Caption         =   "&Cetak"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6960
            TabIndex        =   106
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdBatal 
            Caption         =   "&Baru"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5160
            TabIndex        =   104
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtNoIEDTA 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   49
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtNoRad 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   46
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtNoLab 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   41
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdTutup 
            Caption         =   "Tutu&p"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10575
            TabIndex        =   32
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdSimpan 
            Caption         =   "&Simpan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8760
            TabIndex        =   31
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtNoOrder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   5640
            MaxLength       =   15
            TabIndex        =   30
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtNoPendaftaran 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   4440
            MaxLength       =   15
            TabIndex        =   29
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtNoCM 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   28
            Top             =   0
            Width           =   1575
         End
         Begin VB.TextBox txtTotalBayar 
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
            Height          =   450
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   27
            Top             =   200
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Total Biaya"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   310
            Width           =   1455
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgAsistenPerPelayanan 
      Height          =   1455
      Left            =   5160
      TabIndex        =   101
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2566
      _Version        =   393216
      FixedCols       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid fgInstrumenPerPelayanan 
      Height          =   1455
      Left            =   7920
      TabIndex        =   102
      Top             =   480
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2566
      _Version        =   393216
      FixedCols       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid fgSirkulerPerPelayanan 
      Height          =   1455
      Left            =   11160
      TabIndex        =   103
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2566
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "frmKonsul_OrderPelayanan.frx":0CEA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9360
      Picture         =   "frmKonsul_OrderPelayanan.frx":36AB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3435
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   1800
      Picture         =   "frmKonsul_OrderPelayanan.frx":4433
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13935
   End
End
Attribute VB_Name = "frmKonsul_OrderPelayanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''splakuk 2013-05-18 Konsul dan pesan pelayanan

Option Explicit
Dim subintJmlArray As Integer
Dim subcurHargaSatuan As Currency
Dim subcurTarifService As Currency
Dim subcurHarusDibayar As Currency
Dim curTanggunganRS As Currency
Dim curHutangPenjamin As Currency
Dim subintJmlService As Integer
Dim tempStatusTampil As Boolean
Dim subJenisHargaNetto As Integer
Dim Cancel As Boolean
Dim subJmlTotalPely As Integer
Dim strKdPelayanan() As String
Dim subJmlTotal As Integer
Dim riilnya As Double
Dim blt As Integer
Dim subcurBiayaAdministrasi As Currency
Dim kolom As Integer
Dim ceking As Boolean
Dim varKdJenisPeriksa As String

Dim strFilterPelayanan As String
Dim strCito As String
Dim strKodePelayananRS As String
Dim curBiaya As Currency
Dim curJP As Currency
Dim intJmlPelayanan As Integer
Dim strKdKelas As String
Dim strKelas As String
Dim strKdJenisTarif As String
Dim strJenisTarif As String
Dim intBarang As Integer
Dim intJmlBarang As Integer
Dim intMaxJmlBarang As Integer
Dim strStatusAPBD As String

Dim subKdPemeriksa() As String
Dim curTarifCito As Currency
Dim subcurTarifCito As Currency
Dim subcurTarifBiayaSatuan As Currency
Dim subcurTarifHargaSatuan As Currency
Dim mstrKdDokter2 As String
Dim mstrKdDokterD As String
Dim strPilihGrid As String
Dim i As Integer
Dim j As Integer
Dim intJmlDokterPembantu As Integer
Dim subKdDokterPembantu As String
Dim strFilterDokter As String
Dim subKdDokterAnestesi As String
Dim intJmlDokterAnestesi As Integer
Dim rsNoLabRad As New ADODB.recordset
Dim strSQLNoLabRad As String
Dim dTglPlyn As Date
Dim subKdDokterResusitasi As String
Dim subKdAsisten() As String
Dim subKdPenata() As String
Dim subKdInstrumen() As String
Dim subKdSirkuler() As String

Dim subJmlTotalAsisten As Integer
Dim subJmlTotalPenata As Integer
Dim subJmlTotalInstrumen As Integer
Dim subJmlTotalSirkuler As Integer
Dim intJmlDokterResusitasi As Integer
Dim mstrKdDokterResusitasi As String


Private Sub chkDelegasi_Click()
If chkDelegasi.value = vbChecked Then
If MsgBox("Akan Didelegasikan ke Dokter Atau Paramedis ?? " & vbCrLf & "Pilih YES Untuk DOKTER atau Pilih NO Untuk PARAMEDIS ", vbYesNo, "Validasi") = vbYes Then
    chkPerawat.value = vbUnchecked
    chkPerawat.Enabled = False
    txtDokterDelegasi.Enabled = True
    lvPemeriksa.Enabled = False
Else
    chkPerawat.value = vbChecked
    chkPerawat.Enabled = True
    txtDokterDelegasi.Enabled = False
    lvPemeriksa.Enabled = True
End If
Else
    chkPerawat.value = vbChecked
    chkPerawat.Enabled = True
    txtDokterDelegasi.Enabled = False
    lvPemeriksa.Enabled = True
End If
End Sub

Private Sub chkDelegasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkDelegasi.value = 0 Then
            chkPerawat.SetFocus
        Else
            txtDokterDelegasi.SetFocus
        End If
    End If
End Sub

Public Sub chkDibayardimuka_Click()
If chkDibayardimuka.value = vbChecked Then
    fraPDokter.Visible = True
'    If dcInstalasi.BoundText = "04" Then
'        fraButton.Visible = True
'        chkOperasiBersama.Visible = True
'        Frame4.Height = 3975
'        Frame1.Top = 4155
'        Frame3.Top = 6840
'        Me.Height = 9330
'        chkDelegasi.Enabled = False
'    Else
        fraButton.Visible = False
        chkOperasiBersama.Visible = False
        Frame4.Height = 2415
        Frame1.Top = 2595
        Frame3.Top = 5280
        Me.Height = 7665
        chkDelegasi.Enabled = True
'    End If
Else
    fraPDokter.Visible = False
    fraButton.Visible = False
    chkOperasiBersama.Visible = False
'    Frame4.Height = 1575
'    Frame1.Top = 1755
'    Frame3.Top = 4320
'    Me.Height = 6705
        Frame4.Height = 2415
        Frame1.Top = 2595
        Frame3.Top = 5280
        Me.Height = 7665
    chkDelegasi.Enabled = True
End If
Call centerForm(Me, MDIUtama)
End Sub

Private Sub chkDibayardimuka_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtDokter.SetFocus
'    If chkDibayardimuka.value = vbChecked Then txtDokter.SetFocus Else txtIsiTM.SetFocus
'End If
'
End Sub

Private Sub chkDokter2_Click()
If chkDokter2.value = vbChecked Then
    txtDokter2.Enabled = True
Else
    txtDokter2.Enabled = False
End If
End Sub

Private Sub chkDokter2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If chkDokter2.value = 0 Then
        txtDokterAnestesi.SetFocus
    Else
        txtDokter2.SetFocus
    End If
End If
End Sub

Private Sub chkDokterResusitasi_Click()
On Error Resume Next
    
    If chkDokterResusitasi.value = 0 Then
        txtDokterResusitasi.Enabled = False
        txtDokterResusitasi.Text = ""
        
        If fraDokter.Visible = True Then fraDokter.Visible = False
        If fraButton.Visible = True Then
        chkAsisten.SetFocus
        End If
    Else
        txtDokterResusitasi.Enabled = True
        txtDokterResusitasi.SetFocus
    End If

End Sub

Private Sub cmdCetak_Click()
On Error GoTo errLoad
    If cmdSimpan.Enabled = True Then Exit Sub
    mdTglAwal = dtpTglOrderTM.value
    mdTglAkhir = dtpTglOrderTM.value
    mstrKdRuanganORS = dcRuangan.BoundText
    strSQL = "select StatusCito from DetailOrderPelayananTM where noPendaftaran='" & mstrNoPen & "'"
    Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            strCito = rs("StatusCito").value ' IIf(rs("StatusCito").value = 0, "Y", "T")
        End If
    If strCito = "Y" Then
        strNStsCITO = "Ya"
    ElseIf strCito = "T" Then
        strNStsCITO = "Tidak"
    End If
    strNamaRuangan = dcRuangan.Text
'    mstrNama = dcDokterPerujuk.Text
'----- DiKomen karena tidak bayar dimuka jadi hanya Order aja
'    If chkDibayardimuka.value = 0 Then
'        strCetak = "0"
'    ElseIf chkDibayardimuka.value = 1 Then
'        strCetak = "1"
'    End If
'--------
    strCetak = "0"
    strCetak2 = "TM"
    frm_cetak_RincianBiayaKonsul.Show
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTambah_Click()
On Error GoTo errLoad
Dim i As Integer

Dim subcurTarifBiayaSatuan As Currency
    dTglPlyn = Now
    'chandra 27 02 2014
    ' untuk menghandle jumlah yang kosong
    
    If (txtKuantitas.Text = "") Then
        MsgBox "Jumlah harus di isi"
        txtKuantitas.SetFocus
        Exit Sub
    End If
    subcurTarifBiayaSatuan = Format(dgPelayananRS.Columns("tarif"), "###,###")
    mdTglBerlaku = dtpTglOrderTM.value
    If chkDilayaniDokter.value = vbChecked Then
        If mstrKdDokter = "" Then
            MsgBox "Silahkan pilih Dokter pemeriksa ", vbCritical, "Validasi"
            If fraPDokter.Visible = vbChecked Then
                If chkDilayaniDokter.value = vbUnchecked Then chkDilayaniDokter.SetFocus Else txtDokter.SetFocus

            End If
            Exit Sub
        End If
    End If
    
    If chkPerawat.value = vbChecked And subJmlTotal = 0 Then
        MsgBox "Nama perawat kosong", vbCritical, "Validasi"
        lvPemeriksa.Visible = True
        txtNamaPerawat.SetFocus
        Exit Sub
    End If
    
    
    If dcRuangan.BoundText = "" Then
        MsgBox "Silahkan isi Ruangan tujuan konsul ", vbCritical, "Validasi"
        dcRuangan.SetFocus
        Exit Sub
    End If
    
'    If dcDokterPerujuk.BoundText = "" Then
'        MsgBox "Silahkan isi Dokter perujuk ", vbCritical, "Validasi"
'        dcDokterPerujuk.SetFocus
'        Exit Sub
'    End If
    
    If chkAsisten.value = vbChecked And subJmlTotalAsisten = 0 Then
        MsgBox "Nama perawat kosong", vbCritical, "Validasi"
        lvAsisten.Visible = True
        txtNamaAsisten.SetFocus
        Exit Sub
    End If
    
    If chkInstrumen.value = vbChecked And subJmlTotalInstrumen = 0 Then
        MsgBox "Nama Instrumen kosong", vbCritical, "Validasi"
        LvInstrumen.Visible = True
        txtNamaInstrumen.SetFocus
        Exit Sub
    End If
    
    If ChkSirkuler.value = vbChecked And subJmlTotalSirkuler = 0 Then
        MsgBox "Nama Sirkuler kosong", vbCritical, "Validasi"
        LvSirkuler.Visible = True
        txtNamaSirkuler.SetFocus
        Exit Sub
    End If
    
    For i = 1 To lvPemeriksa.ListItems.Count
            lvPemeriksa.ListItems(i).Checked = False
    Next i
    For i = 1 To lvAsisten.ListItems.Count
            lvAsisten.ListItems(i).Checked = False
    Next i
    For i = 1 To LvInstrumen.ListItems.Count
            LvInstrumen.ListItems(i).Checked = False
    Next i
    For i = 1 To LvSirkuler.ListItems.Count
            LvSirkuler.ListItems(i).Checked = False
    Next i
    Dim a As Integer
    Dim max_i As Integer
    With fgPelayanan
      max_i = .Rows
        i = max_i - 1
       For a = 1 To i
           If txtIsiTM.Text = "" Then Exit Sub
           If (.TextMatrix(a, 0) = txtKdIsiTM.Text) And _
                (.TextMatrix(a, 20) = dtpTglOrderTM.value) Then txtIsiTM.SetFocus: txtIsiTM.SelStart = 0: txtIsiTM.SelLength = Len(txtIsiTM.Text): Exit Sub
       Next a

        .TextMatrix(i, 0) = dgPelayananRS.Columns("KdPelayananRS")
        .TextMatrix(i, 1) = dgPelayananRS.Columns("NamaPelayanan")
        .TextMatrix(i, 2) = CInt(txtKuantitas.Text)
        subcurTarifCito = sp_Take_TarifBPT(dgPelayananRS.Columns("KdPelayananRS"))
        .TextMatrix(i, 3) = IIf(subcurTarifBiayaSatuan = 0, 0, Format(subcurTarifBiayaSatuan, "#,###"))
        .TextMatrix(i, 4) = IIf(funcRoundUp(CStr(subcurTarifBiayaSatuan + subcurTarifCito)) * CInt(txtKuantitas.Text) = 0, 0, Format(funcRoundUp(CStr(subcurTarifBiayaSatuan + subcurTarifCito)) * CInt(txtKuantitas.Text), "#,###"))
        .TextMatrix(i, 5) = mdTglBerlaku
        .TextMatrix(i, 6) = mstrKdDokter
        If optCito(0).value = True Then strCito = "1" Else strCito = "0"

        .TextMatrix(i, 7) = strCito
        .TextMatrix(i, 8) = IIf(Val(subcurTarifCito) = 0, 0, Format(subcurTarifCito, "#,###"))
        .TextMatrix(i, 9) = dTglPlyn
        .TextMatrix(i, 10) = ""
        .TextMatrix(i, 11) = IIf(chkDelegasi.value = vbChecked, mstrKdDokterD, "")
        .TextMatrix(i, 12) = "0" 'for pesan pelayanan
        
        .TextMatrix(i, 13) = IIf(chkDilayaniDokter.value = vbChecked, mstrKdDokter, "")
        .TextMatrix(i, 14) = subKdDokterPembantu
        .TextMatrix(i, 15) = IIf(subKdDokterAnestesi = "", "", subKdDokterAnestesi)
        .TextMatrix(i, 16) = IIf(chkDokter2.value = vbChecked, mstrKdDokter2, "")
        .TextMatrix(i, 17) = IIf(chkDokterResusitasi.value = vbChecked, mstrKdDokterResusitasi, "")
        
        .TextMatrix(i, 18) = dcRuangan.BoundText
        .TextMatrix(i, 19) = ""
        .TextMatrix(i, 20) = dtpTglOrderTM.value
        .TextMatrix(i, 21) = dcJenisOperasi.BoundText
        ' chandra 28 02 2014
        ' karena rekamedis tidak bayar di muka langsung
        .TextMatrix(i, 22) = "T" 'IIf(chkDibayardimuka.value = vbChecked, "Y", "T")
        .TextMatrix(i, 23) = IIf(chkDelegasi.value = vbChecked, "Y", "T")
        .TextMatrix(i, 24) = IIf(chkOperasiBersama.value = vbChecked, "Y", "T")
        .Rows = .Rows + 1
        
    End With
    Call HitungTotal
'    dgPelayananRS.Visible = True
    Dim s As Integer
    
    
    If chkPerawat.value = vbChecked Then Call subLoadPelayananPerPerawat
    If chkAsisten.value = vbChecked Then Call subLoadPelayananPerAsisten
    If chkInstrumen.value = vbChecked Then Call subLoadPelayananPerInstrumen
    If ChkSirkuler.value = vbChecked Then Call subLoadPelayananPerSirkuler
 
    txtKuantitas.Text = 1
    txtIsiTM.SetFocus
    ceking = True
    txtIsiTM.Text = ""
    txtKdIsiTM.Text = ""
    dgPelayananRS.Visible = False
Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub dcRuangan_Change()
''here
    dtpTglOrderTM.Minute = Format(Now, "nn")
    dtpTglOrderTM.Second = Format(Now, "ss")
End Sub

Private Sub LvInstrumen_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Or KeyAscii = 27 Then LvInstrumen.Visible = False: txtNamaInstrumen.SetFocus
End Sub

Private Sub LvSirkuler_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 27 Then LvSirkuler.Visible = False: txtNamaSirkuler.SetFocus
End Sub

Private Sub optCito_KeyPress(index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtIsiTM.Text = "" Then
        If (cmdSimpan.Enabled = True) Then
            cmdSimpan.SetFocus
        End If
    Else
        If (cmdTambah.Enabled = True) Then
            cmdTambah.SetFocus
        End If
    End If
End If
End Sub

Private Sub txtDokterResusitasi_Change()
    On Error Resume Next
    strPilihGrid = "DokterResus"
    strFilterDokter = "WHERE NamaDokter like '%" & txtDokterResusitasi.Text & "%'"
    subKdDokterResusitasi = ""
    fraDokter.Visible = True
    Call subLoadDokterResusitasi
End Sub

Private Sub subLoadDokterResusitasi()
    On Error Resume Next
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & strFilterDokter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDokterResusitasi = rs.RecordCount
    With dgDokter
        Set .DataSource = rs
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
    fraDokter.Left = 360
    fraDokter.Top = 4920
End Sub

Private Sub chkPerawat_Click()
If chkPerawat.value = vbChecked Then
        txtNamaPerawat.Enabled = True
        strSQL = "SELECT IdPegawai FROM V_DaftarPemeriksaPasien WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
        
            txtNamaPerawat.Text = strNmPegawai
            If lvPemeriksa.ListItems.Count > 0 Then
                lvPemeriksa.ListItems.Item("key" & strIDPegawaiAktif).Checked = True
                Call lvPemeriksa_ItemCheck(lvPemeriksa.ListItems.Item("key" & strIDPegawaiAktif))
            End If
        Else
            txtNamaPerawat.Text = ""
        End If
    Else
        lvPemeriksa.ListItems.Clear
        subJmlTotal = 0
        chkPerawat.Caption = "Paramedis/Penata"
        txtNamaPerawat.Text = ""
        txtNamaPerawat.Enabled = False
    End If
    lvPemeriksa.Visible = False
End Sub

Private Sub chkPerawat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If dcInstalasi.BoundText = "04" Then dcJenisOperasi.SetFocus Else txtIsiTM.SetFocus
End If
End Sub

Private Sub chkAsisten_Click()
    If chkAsisten.value = vbChecked Then
        strSQL = "SELECT IdPegawai FROM V_DaftarPemeriksaPasien WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            txtNamaAsisten.Text = strNmPegawai
            If lvAsisten.ListItems.Count > 0 Then
                lvAsisten.ListItems.Item("key" & strIDPegawaiAktif).Checked = False
                Call lvAsisten_ItemCheck(lvAsisten.ListItems.Item("key" & strIDPegawaiAktif))
            End If
        Else
            txtNamaAsisten.Text = ""
        End If
        txtNamaAsisten.Enabled = True
        lvAsisten.Visible = True
        lvAsisten.Enabled = True
    Else
        chkAsisten.Caption = "Asisten"
        subJmlTotalAsisten = 0
        txtNamaAsisten.Text = ""
        txtNamaAsisten.BackColor = &HFFFFFF
        txtNamaAsisten.Enabled = False
        lvAsisten.Enabled = False ' yang Lama k gini
        lvAsisten.Visible = False
        For i = 1 To lvAsisten.ListItems.Count
            lvAsisten.ListItems(i).Checked = False
        Next i
    End If
    'lvAsisten.Visible = True 'yang lama k gini
    lvAsisten.Visible = False
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo hell

    If dcRuangan.Text = "" Then
        MsgBox "Ruangan tujuan konsul kosong ", vbCritical, "Validasi"
        dcRuangan.SetFocus
        Exit Sub
    End If
    If chkDilayaniDokter.value = vbChecked Then
        If mstrKdDokter = "" Then
            MsgBox "Dokter pemeriksa kosong ", vbCritical, "Validasi"
            If chkDilayaniDokter.value = vbUnchecked Then chkDilayaniDokter.SetFocus Else txtDokter.SetFocus
            Exit Sub
        End If
    End If
    
'    If fgPelayanan.TextMatrix(1, 0) = "" Then
'        MsgBox "Pelayanan kosong ", vbCritical, "Validasi"
'        txtIsiTM.SetFocus
'        Exit Sub
'    End If

    
    Call sp_Rujukan(dbcmd)
    ''proses simpan pasien konsul/pasien rujukan
    
    If fgPelayanan.TextMatrix(1, 0) <> "" Then
        If MsgBox("Apakah Anda Akan menggunakan CITO?" & vbNewLine & "Pilih No Jika Tidak Menggunakan CITO", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then
            strCito = "0"
        Else
            strCito = 1
        End If
      End If
 
    If sp_PasienRujukanNStrukOrder = False Then Exit Sub
    
    For i = 1 To fgPelayanan.Rows - 2
'        strSQLX = "select * from PasienRujukan where NoPendaftaran='" & txtNoPendaftaranTM.Text & "' and TglDirujuk='" & Format(fgPelayanan.TextMatrix(i, 20), "yyyy/MM/dd HH:mm:ss") & "' and Kdruangantujuan='" & fgPelayanan.TextMatrix(i, 18) & "'"
'        Call msubRecFO(rsx, strSQLX)
'        If rsx.EOF = False Then GoTo lanjut_
'        ''proses simpan pasien konsul/pasien rujukan
'        If sp_PasienRujukanNStrukOrder(fgPelayanan.TextMatrix(i, 18), fgPelayanan.TextMatrix(i, 20), fgPelayanan.TextMatrix(i, 21), fgPelayanan.TextMatrix(i, 22)) = False Then Exit Sub
'lanjut_:
        ''proses simpan pemesanan pelayanan
'        If sp_OrderPelayananTMBayarDimuka(fgPelayanan.TextMatrix(i, 0), fgPelayanan.TextMatrix(i, 2), fgPelayanan.TextMatrix(i, 7), fgPelayanan.TextMatrix(i, 6), _
'            fgPelayanan.TextMatrix(i, 16), fgPelayanan.TextMatrix(i, 15), fgPelayanan.TextMatrix(i, 11), fgPelayanan.TextMatrix(i, 19), fgPelayanan.TextMatrix(i, 21), fgPelayanan.TextMatrix(i, 22), fgPelayanan.TextMatrix(i, 23), fgPelayanan.TextMatrix(i, 24)) = False Then Exit Sub
'
     If sp_OrderPelayananTMBayarDimuka(fgPelayanan.TextMatrix(i, 0), fgPelayanan.TextMatrix(i, 2), strCito, fgPelayanan.TextMatrix(i, 6), _
            mstrKdDokter2, subKdDokterAnestesi, mstrKdDokterD, fgPelayanan.TextMatrix(i, 19), fgPelayanan.TextMatrix(i, 21), fgPelayanan.TextMatrix(i, 22), fgPelayanan.TextMatrix(i, 23), fgPelayanan.TextMatrix(i, 24)) = False Then Exit Sub
    
    Next i
    
    If chkDibayardimuka.value = vbChecked Then
        
        If dcInstalasi.BoundText = "04" Then
        
            If chkDokter2.value = 1 Then
                If txtDokter2.Text = "" Then
                    MsgBox "Dokter operator 2 kosong ", vbCritical, "Validasi"
                    txtDokter2.SetFocus
                    Exit Sub
                End If
            End If
            
            If chkDokterResusitasi.value = 1 Then
                If txtDokterResusitasi.Text = "" Then
                    MsgBox "Dokter Resusitasi Neonatus kosong ", vbCritical, "Validasi"
                    txtDokterResusitasi.SetFocus
                    Exit Sub
                End If
            End If
        
            For i = 1 To fgPelayanan.Rows - 2
                If fgPelayanan.TextMatrix(i, 13) <> "" Then
                    With fgPelayanan
                        If sp_DokterPelaksanaOperasi(.TextMatrix(i, 9), .TextMatrix(i, 0), .TextMatrix(i, 13), .TextMatrix(i, 16), .TextMatrix(i, 15), .TextMatrix(i, 14), .TextMatrix(i, 17)) = False Then Exit Sub
                    End With
                End If
            Next i
            
            
            For i = 1 To fgAsistenPerPelayanan.Rows - 1
                With fgAsistenPerPelayanan
                    If sp_PetugasPemeriksaBP(.TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 1)) = False Then Exit Sub
                    
                End With
            Next i
    
            'If chkInstrumen.Value = Checked Then
            For i = 1 To fgInstrumenPerPelayanan.Rows - 1
                With fgInstrumenPerPelayanan
                    If sp_PetugasInstrumenBP(.TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 1)) = False Then Exit Sub
                End With
            Next i
        
            'If ChkSirkuler.Value = Checked Then
            For i = 1 To fgSirkulerPerPelayanan.Rows - 1
                With fgSirkulerPerPelayanan
                    If sp_PetugasSirkulerBP(.TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 1)) = False Then Exit Sub
                End With
            Next i
            
            If chkPerawat.value = 1 Then
                If fgPerawatPerPelayanan.TextMatrix(1, 0) = "" Then
                    MsgBox "Perawat/Penata Anestesi kosong ", vbCritical, "Validasi"
                    chkPerawat.SetFocus
                    Exit Sub
                End If
        
                For i = 1 To fgPerawatPerPelayanan.Rows - 1
                    With fgPerawatPerPelayanan
                        If sp_PetugasPenataAnes(.TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 1)) = False Then Exit Sub
                    End With
                Next i
            End If
        
        Else
            If chkPerawat.value = 1 Then
                If fgPerawatPerPelayanan.TextMatrix(1, 0) = "" Then
                    MsgBox "Perawat kosong ", vbCritical, "Validasi"
                    chkPerawat.SetFocus
                    Exit Sub
                End If
        
                For i = 1 To fgPerawatPerPelayanan.Rows - 1
                    With fgPerawatPerPelayanan
                        If sp_PetugasPemeriksaBP(.TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 1)) = False Then Exit Sub
                    End With
                Next i
            End If
        End If
    End If
    MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
    cmdCetak.Enabled = True
    cmdSimpan.Enabled = False
    cmdCetak.SetFocus
    flagJikaPasienSudahRujukan = True
Exit Sub
hell:
    Call msubPesanError
End Sub
Private Sub sp_Rujukan(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaranTM.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCMTM.Text)
        .Parameters.Append .CreateParameter("NoRujukan", adVarChar, adParamInput, 30, Null)
        .Parameters.Append .CreateParameter("KdRujukanAsal", adChar, adParamInput, 2, mstrKdAsalRujukan)
        .Parameters.Append .CreateParameter("SubRujukanAsal", adVarChar, adParamInput, 100, dcTempatPerujuk.Text)
        .Parameters.Append .CreateParameter("NamaPerujuk", adVarChar, adParamInput, 50, dcNamaPerujuk.Text)
        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , Format(dtpTglRujuk.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("DiagnosaRujukan", adVarChar, adParamInput, 100, dcDiagnosa.Text)

        .ActiveConnection = dbConn
        .CommandText = "AU_Rujukan"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam Pemasukan Data Rujukan", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("AU_Rujukan")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
        mstrKdInstalasiPerujuk = ""
    End With
    Exit Sub
End Sub

'simpan data Penata Anestesi
Private Function sp_PetugasPenataAnes(F_dtTanggalPelayanan As Date, F_strKodePelayanan As String, F_StrIdPerawat As String, f_KdRuangan As String) As Boolean
On Error GoTo errLoad

    sp_PetugasPenataAnes = False
    
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(F_dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, F_strKodePelayanan)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, F_StrIdPerawat)  'kode perawat
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PetugasPelaksanaAnastesi"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data petugas Penata Anestesi BP", vbExclamation, "Validasi"
            sp_PetugasPenataAnes = False
        
        End If
    
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
        sp_PetugasPenataAnes = True
    End With

Exit Function
errLoad:
    Call msubPesanError
    sp_PetugasPenataAnes = False
End Function

'simpan data Instrumen
Private Function sp_PetugasInstrumenBP(F_dtTanggalPelayanan As Date, F_strKodePelayanan As String, F_StrIdPerawat As String, f_KdRuangan As String) As Boolean
On Error GoTo errLoad

    sp_PetugasInstrumenBP = False
    
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(F_dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, F_strKodePelayanan)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, F_StrIdPerawat)  'kode perawat
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PetugasPelaksanaInstrumen"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data petugas Instrumen BP", vbExclamation, "Validasi"
            sp_PetugasInstrumenBP = False
        
        End If
    
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
        sp_PetugasInstrumenBP = True
    End With

Exit Function
errLoad:
    Call msubPesanError
    sp_PetugasInstrumenBP = False
End Function

'simpan data Sirkuler
Private Function sp_PetugasSirkulerBP(F_dtTanggalPelayanan As Date, F_strKodePelayanan As String, F_StrIdPerawat As String, f_KdRuangan As String) As Boolean
On Error GoTo errLoad

    sp_PetugasSirkulerBP = False
    
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(F_dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, F_strKodePelayanan)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, F_StrIdPerawat)  'kode perawat
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PetugasPelaksanaSirkuler"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data petugas Sirkuler BP", vbExclamation, "Validasi"
            sp_PetugasSirkulerBP = False
        
        End If
    
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
        sp_PetugasSirkulerBP = True
    End With

Exit Function
errLoad:
    Call msubPesanError
    sp_PetugasSirkulerBP = False
End Function

Private Function sp_PasienRujukanNStrukOrder()
On Error Resume Next
sp_PasienRujukanNStrukOrder = True
Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaranTM.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCMTM.Text)
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, dcRuangan.BoundText)
        .Parameters.Append .CreateParameter("IdDokterPerujuk", adChar, adParamInput, 10, mstrKdDokterPerujuk)
        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , Format(dtpTglOrderTM, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("OutKode", adChar, adParamOutput, 10, Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrKdDokter)
        .Parameters.Append .CreateParameter("KdJenisOperasi", adChar, adParamInput, 2, IIf(dcJenisOperasi.Text = "", Null, dcJenisOperasi.BoundText))
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, TempKodeKelas)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("StatusBayarDimuka", adChar, adParamInput, 1, "Y")
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PasienRujukanDibayarDimukaRM"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam proses penyimpanan data", vbCritical, "Validasi"
            sp_PasienRujukanNStrukOrder = False
        Else
            If (Not IsNull(.Parameters("OutKode").value)) Then txtNoOrder1.Text = .Parameters("OutKode").value
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
'hell:
'    Call msubPesanError
'    sp_PasienRujukanNStrukOrder = False
End Function

Private Function sp_OrderPelayananTMBayarDimuka(f_KdPelayanan As String, f_JmlPelayanan As Integer, f_StatusCito As String, _
    f_IdPegawai1 As String, _
    f_IdPegawai2 As String, f_IdPegawai3 As String, f_IdPegawaiD As String, f_idDokterPerujuk As String, f_kdjenisoperasi As String, f_StatusBayar As String, fStatusDelegasi As String, f_statusOperasibersama As String) As Boolean
On Error GoTo errLoad

    sp_OrderPelayananTMBayarDimuka = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder1.Text)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, f_KdPelayanan)
        .Parameters.Append .CreateParameter("JmlPelayanan", adInteger, adParamInput, , f_JmlPelayanan)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("IdDokterOrder", adChar, adParamInput, 10, mstrKdDokterPerujuk)
        .Parameters.Append .CreateParameter("StatusCITO", adChar, adParamInput, 1, IIf(f_StatusCito = 1, "Y", "T")) ' f_StatusCito
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCMTM.Text)
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, IIf(mstrValid = "", Null, mstrValid))
        .Parameters.Append .CreateParameter("NoRetur", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("JmlRetur", adInteger, adParamInput, , Null)
        .Parameters.Append .CreateParameter("KdPelayananRSUsed", adChar, adParamInput, 6, Null)
        .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 200, Null)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, TempKodeKelas)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, f_IdPegawai1) 'mstrKdDokterPerujuk
        .Parameters.Append .CreateParameter("IdPegawai2", adChar, adParamInput, 10, IIf(f_IdPegawai2 = "", Null, f_IdPegawai2))
        .Parameters.Append .CreateParameter("IdPegawai3", adChar, adParamInput, 10, IIf(f_IdPegawai3 = "", Null, f_IdPegawai3))
        .Parameters.Append .CreateParameter("KdJenisOperasi", adChar, adParamInput, 2, IIf(f_kdjenisoperasi = "", Null, f_kdjenisoperasi))
        If dcInstalasi.BoundText = "04" Then
        .Parameters.Append .CreateParameter("StatusDokterBersama", adChar, adParamInput, 1, f_statusOperasibersama)
        Else
        .Parameters.Append .CreateParameter("StatusDelegasi", adChar, adParamInput, 1, fStatusDelegasi)
        End If
        .Parameters.Append .CreateParameter("StatusBayarDimuka", adChar, adParamInput, 1, f_StatusBayar)
        .Parameters.Append .CreateParameter("IdPegawaiD", adChar, adParamInput, 10, IIf(f_IdPegawaiD = "", Null, f_IdPegawaiD)) '
           
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_OrderPelayananTMBayarDimuka"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam proses penyimpanan data", vbExclamation, "Validasi"
            sp_OrderPelayananTMBayarDimuka = False
            GoTo errLoad
        End If
    
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    
Exit Function
errLoad:
    Call msubPesanError
    sp_OrderPelayananTMBayarDimuka = False
End Function
'simpan data perawat
Private Function sp_PetugasPemeriksaBP(F_dtTanggalPelayanan As Date, F_strKodePelayanan As String, F_StrIdPerawat As String, f_KdRuangan As String) As Boolean
On Error GoTo errLoad

    sp_PetugasPemeriksaBP = False
    
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(F_dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, F_strKodePelayanan)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, F_StrIdPerawat)  'kode perawat
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PetugasPemeriksaBP"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data petugas pemeriksa BP", vbExclamation, "Validasi"
            sp_PetugasPemeriksaBP = False
        
        End If
    
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
        sp_PetugasPemeriksaBP = True
    End With

Exit Function
errLoad:
    Call msubPesanError
    sp_PetugasPemeriksaBP = False
End Function

'simpan data dokter pelaksana operasi
Private Function sp_DokterPelaksanaOperasi(F_dtTanggalPelayanan As Date, F_strKodePelayanan As String, _
                        f_IdDokterOperator1 As String, f_IdDokterOperator2 As String, f_IdDokterAnastesi As String, f_IdDokterPendamping As String, f_IdDokterResus As String) As Boolean
On Error GoTo errLoad

    sp_DokterPelaksanaOperasi = False
    
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, dcRuangan.BoundText)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, F_strKodePelayanan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(F_dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdDokterOperator1", adChar, adParamInput, 10, f_IdDokterOperator1)
        .Parameters.Append .CreateParameter("IdDokterOperator2", adChar, adParamInput, 10, IIf(f_IdDokterOperator2 = "", Null, f_IdDokterOperator2))
        .Parameters.Append .CreateParameter("IdDokterAnastesi", adChar, adParamInput, 10, IIf(f_IdDokterAnastesi = "", Null, f_IdDokterAnastesi))
        .Parameters.Append .CreateParameter("IdDokterPendamping", adChar, adParamInput, 10, IIf(f_IdDokterPendamping = "", Null, f_IdDokterPendamping))
        .Parameters.Append .CreateParameter("IdDokterResusitasi", adChar, adParamInput, 10, IIf(f_IdDokterResus = "", Null, f_IdDokterResus))
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DokterPelaksanaOperasi"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data ", vbExclamation, "Validasi"
            sp_DokterPelaksanaOperasi = False
        
        End If
    
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
        sp_DokterPelaksanaOperasi = True
    End With

Exit Function
errLoad:
    Call msubPesanError
    sp_DokterPelaksanaOperasi = False
End Function

Private Sub dcInstalasi_Change()
On Error Resume Next
dcRuangan.BoundText = ""
'If chkDibayardimuka.value = vbChecked Then
'If dcInstalasi.BoundText = "04" Then
'    fraButton.Visible = True
'    chkOperasiBersama.Visible = True
'Else
'    fraButton.Visible = False
'    chkOperasiBersama.Visible = False
'End If
'End If
chkDibayardimuka_Click
End Sub

Private Sub dcJenisOperasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkDokter2.SetFocus
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
On Error Resume Next
If strPilihGrid = "Dokter" Then
    If KeyAscii = 13 Then
        If mintJmlDokter = 0 Then Exit Sub
        txtDokter.Text = dgDokter.Columns(1).value
        mstrKdDokter = dgDokter.Columns(0).value
        If mstrKdDokter = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        chkDilayaniDokter.value = 1
        fraDokter.Visible = False
        txtDokter.SetFocus
'        If chkDelegasi.Enabled = True Then
'        chkDelegasi.SetFocus
'        Else
'        chkPerawat.SetFocus
'        End If
    End If
ElseIf strPilihGrid = "Dokter2" Then
    If KeyAscii = 13 Then
        If mintJmlDokter = 0 Then Exit Sub
        txtDokter2.Text = dgDokter.Columns(1).value
        mstrKdDokter2 = dgDokter.Columns(0).value
        If mstrKdDokter = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter2.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        fraDokter.Visible = False
        txtDokterAnestesi.SetFocus
    End If
ElseIf strPilihGrid = "DokterD" Then
    If KeyAscii = 13 Then
        If mintJmlDokter = 0 Then Exit Sub
        txtDokterDelegasi.Text = dgDokter.Columns(1).value
        mstrKdDokterD = dgDokter.Columns(0).value
        If mstrKdDokterD = "" Then
            MsgBox "Pilih dulu Dokter Delegasi yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokterDelegasi.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        fraDokter.Visible = False
        chkPerawat.SetFocus
    End If
ElseIf strPilihGrid = "DokterResus" Then
    If KeyAscii = 13 Then
        If mintJmlDokter = 0 Then Exit Sub
        txtDokterResusitasi.Text = dgDokter.Columns(1).value
        mstrKdDokterResusitasi = dgDokter.Columns(0).value
        If mstrKdDokterResusitasi = "" Then
            MsgBox "Pilih dulu Dokter Resusitasi Pasien", vbCritical, "Validasi"
            txtDokterResusitasi.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        fraDokter.Visible = False
        chkAsisten.SetFocus
    End If
    If KeyAscii = 27 Then
        fraDokter.Visible = False
    End If
If KeyAscii = 27 Then
    fraDokter.Visible = False
End If
End If
End Sub

Private Sub chkAsisten_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkAsisten.value = vbChecked Then
            txtNamaAsisten.SetFocus
        Else
'            optCito(1).SetFocus
        End If
    End If
End Sub

Private Sub txtIsiTM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If dgPelayananRS.Visible = False Then Exit Sub
        dgPelayananRS.SetFocus
        
    End If
End Sub

Private Sub txtNamaAsisten_Change()
On Error GoTo errLoad

    Call subLoadListAsisten("where [Nama Pemeriksa] LIKE '%" & txtNamaAsisten.Text & "%'")
    lvAsisten.Visible = True

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadListAsisten(Optional strKriteria As String)
Dim strKey As String
    
    strSQL = "select * from v_daftarpemeriksapasien " & strKriteria & " order by [Nama Pemeriksa]"
    Call msubRecFO(rs, strSQL)
    
    With lvAsisten
        .ListItems.Clear
        For i = 0 To rs.RecordCount - 1
            strKey = "key" & rs(0).value
            .ListItems.Add , strKey, rs(1).value
            rs.MoveNext
        Next
    
      '  .Top = fraButton.Top + txtNamaAsisten.Top + txtNamaAsisten.Height
        .Left = fraButton.Left + txtNamaAsisten.Left
        .Height = 2000
        .ColumnHeaders.Item(1).Width = lvAsisten.Width - 500
        .Width = txtNamaAsisten.Width
        If subJmlTotalAsisten = 0 Then Exit Sub
        For i = 1 To .ListItems.Count
            For j = 1 To subJmlTotalAsisten
                If .ListItems(i).Key = subKdAsisten(j) Then .ListItems(i).Checked = True
            Next j
        Next i
    End With
End Sub

Private Sub txtNamaAsisten_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If lvAsisten.Visible = True Then If lvAsisten.ListItems.Count > 0 Then lvAsisten.SetFocus
        Case vbKeyEscape
            lvAsisten.Visible = False
    End Select
End Sub

Private Sub txtNamaAsisten_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 13 Then
        If lvAsisten.Visible = True Then
            lvAsisten.SetFocus
        Else
            lvAsisten.Visible = False
            chkInstrumen.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        lvAsisten.Visible = False
    End If
    If KeyAscii = 39 Then KeyAscii = 0
Exit Sub
hell:
End Sub

Private Sub lvAsisten_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim blnSelected As Boolean
    If Item.Checked = True Then
        lvAsisten.ListItems(Item.Key).ForeColor = vbBlue
        subJmlTotalAsisten = subJmlTotalAsisten + 1
        ReDim Preserve subKdAsisten(subJmlTotalAsisten)
        subKdAsisten(subJmlTotalAsisten) = Item.Key
    Else
        lvAsisten.ListItems(Item.Key).ForeColor = vbBlack
        blnSelected = False
        For i = 1 To subJmlTotalAsisten
            If subKdAsisten(i) = Item.Key Then blnSelected = True
            If blnSelected = True Then
                If i = subJmlTotalAsisten Then
                    subKdAsisten(i) = ""
                Else
                    subKdAsisten(i) = subKdAsisten(i + 1)
                End If
            End If
        Next i
        If subJmlTotalAsisten < 1 Then
            subJmlTotalAsisten = 0
        Else
            subJmlTotalAsisten = subJmlTotalAsisten - 1
        End If
    
    End If

    If subJmlTotalAsisten = 0 Then
        txtNamaAsisten.BackColor = &HFFFFFF
        chkAsisten.Caption = "Asisten"
    Else
        txtNamaAsisten.BackColor = &HC0FFFF
        chkAsisten.Caption = "Asisten (" & subJmlTotalAsisten & " org)"
    End If
End Sub

Private Sub lvAsisten_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 27 Then lvAsisten.Visible = False: txtNamaAsisten.SetFocus
End Sub

Private Sub chkInstrumen_Click()
    If chkInstrumen.value = vbChecked Then
        strSQL = "SELECT IdPegawai FROM V_DaftarPemeriksaPasien WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            txtNamaInstrumen.Text = strNmPegawai
            If LvInstrumen.ListItems.Count > 0 Then
                LvInstrumen.ListItems.Item("key" & strIDPegawaiAktif).Checked = False
                Call lvInstrumen_ItemCheck(LvInstrumen.ListItems.Item("key" & strIDPegawaiAktif))
            End If
        Else
            txtNamaInstrumen.Text = ""
        End If
        txtNamaInstrumen.Enabled = True
        LvInstrumen.Visible = True
        LvInstrumen.Enabled = True
    Else
        chkInstrumen.Caption = "Instrumen"
        subJmlTotalInstrumen = 0
        txtNamaInstrumen.Text = ""
        txtNamaInstrumen.BackColor = &HFFFFFF
        txtNamaInstrumen.Enabled = False
        LvInstrumen.Enabled = False ' yang Lama k gini
        LvInstrumen.Visible = False
        For i = 1 To LvInstrumen.ListItems.Count
            LvInstrumen.ListItems(i).Checked = False
        Next i
    End If
    LvInstrumen.Visible = False
End Sub

Private Sub txtNamaInstrumen_Change()
On Error GoTo errLoad

    Call subLoadListInstrumen("where [Nama Pemeriksa] LIKE '%" & txtNamaInstrumen.Text & "%'")
    LvInstrumen.Visible = True

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadListInstrumen(Optional strKriteria As String) ' Add By Indra
Dim strKey As String
    
    strSQL = "select * from v_daftarpemeriksapasien " & strKriteria & " order by [Nama Pemeriksa]"
    Call msubRecFO(rs, strSQL)
    
    With LvInstrumen
        .ListItems.Clear
        For i = 0 To rs.RecordCount - 1
            strKey = "key" & rs(0).value
            .ListItems.Add , strKey, rs(1).value
            rs.MoveNext
        Next
    
        '.Top = fraButton.Top + txtNamaInstrumen.Top + txtNamaInstrumen.Height
        .Left = fraButton.Left + txtNamaInstrumen.Left
        .Height = 2000
        .ColumnHeaders.Item(1).Width = LvInstrumen.Width - 500
        .Width = txtNamaInstrumen.Width
        If subJmlTotalInstrumen = 0 Then Exit Sub
        For i = 1 To .ListItems.Count
            For j = 1 To subJmlTotalInstrumen
                If .ListItems(i).Key = subKdInstrumen(j) Then .ListItems(i).Checked = True
            Next j
        Next i
    End With
End Sub

Private Sub txtNamaInstrumen_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If LvInstrumen.Visible = True Then If LvInstrumen.ListItems.Count > 0 Then LvInstrumen.SetFocus
        Case vbKeyEscape
            LvInstrumen.Visible = False
    End Select
End Sub

Private Sub txtNamaInstrumen_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 13 Then
        If LvInstrumen.Visible = True Then
            LvInstrumen.SetFocus
        Else
            LvInstrumen.Visible = False
            ChkSirkuler.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        LvInstrumen.Visible = False
    End If
    If KeyAscii = 39 Then KeyAscii = 0
Exit Sub
hell:
End Sub

Private Sub lvInstrumen_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim blnSelected As Boolean
    If Item.Checked = True Then
        LvInstrumen.ListItems(Item.Key).ForeColor = vbBlue
        subJmlTotalInstrumen = subJmlTotalInstrumen + 1
        ReDim Preserve subKdInstrumen(subJmlTotalInstrumen)
        subKdInstrumen(subJmlTotalInstrumen) = Item.Key
    Else
        LvInstrumen.ListItems(Item.Key).ForeColor = vbBlack
        blnSelected = False
        For i = 1 To subJmlTotalInstrumen
            If subKdInstrumen(i) = Item.Key Then blnSelected = True
            If blnSelected = True Then
                If i = subJmlTotalInstrumen Then
                    subKdInstrumen(i) = ""
                Else
                    subKdInstrumen(i) = subKdInstrumen(i + 1)
                End If
            End If
        Next i
        If subJmlTotalInstrumen < 1 Then
            subJmlTotalInstrumen = 0
        Else
            subJmlTotalInstrumen = subJmlTotalInstrumen - 1
        End If
    
    End If

    If subJmlTotalInstrumen = 0 Then
        txtNamaInstrumen.BackColor = &HFFFFFF
        chkInstrumen.Caption = "Instrumen"
    Else
        txtNamaInstrumen.BackColor = &HC0FFFF
        chkInstrumen.Caption = "Instrumen (" & subJmlTotalInstrumen & " org)"
    End If
End Sub

Private Sub ChkSirkuler_Click()
    If ChkSirkuler.value = vbChecked Then
        strSQL = "SELECT IdPegawai FROM V_DaftarPemeriksaPasien WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            txtNamaSirkuler.Text = strNmPegawai
            If LvSirkuler.ListItems.Count > 0 Then
                LvSirkuler.ListItems.Item("key" & strIDPegawaiAktif).Checked = False
                Call lvSirkuler_ItemCheck(LvSirkuler.ListItems.Item("key" & strIDPegawaiAktif))
            End If
        Else
            txtNamaSirkuler.Text = ""
        End If
        txtNamaSirkuler.Enabled = True
        LvSirkuler.Visible = True
        LvSirkuler.Enabled = True
    Else
        ChkSirkuler.Caption = "Sirkuler"
        subJmlTotalSirkuler = 0
        txtNamaSirkuler.Text = ""
        txtNamaSirkuler.BackColor = &HFFFFFF
        txtNamaSirkuler.Enabled = False
        LvSirkuler.Enabled = False ' yang Lama k gini
        LvSirkuler.Visible = False
        For i = 1 To LvSirkuler.ListItems.Count
            LvSirkuler.ListItems(i).Checked = False
        Next i
    End If
    LvSirkuler.Visible = False
End Sub

Private Sub txtNamaSirkuler_Change()
On Error GoTo errLoad

    Call subLoadListSirkuler("where [Nama Pemeriksa] LIKE '%" & txtNamaSirkuler.Text & "%'")
    LvSirkuler.Visible = True

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadListSirkuler(Optional strKriteria As String)
Dim strKey As String
    
    strSQL = "select * from v_daftarpemeriksapasien " & strKriteria & " order by [Nama Pemeriksa]"
    Call msubRecFO(rs, strSQL)
    
    With LvSirkuler
        .ListItems.Clear
        For i = 0 To rs.RecordCount - 1
            strKey = "key" & rs(0).value
            .ListItems.Add , strKey, rs(1).value
            rs.MoveNext
        Next
    
        '.Top = fraButton.Top + txtNamaInstrumen.Top + txtNamaInstrumen.Height
        .Left = fraButton.Left + txtNamaSirkuler.Left
        .Height = 2000
        .ColumnHeaders.Item(1).Width = LvInstrumen.Width - 500
        .Width = txtNamaSirkuler.Width
        If subJmlTotalSirkuler = 0 Then Exit Sub
        For i = 1 To .ListItems.Count
            For j = 1 To subJmlTotalSirkuler
                If .ListItems(i).Key = subKdSirkuler(j) Then .ListItems(i).Checked = True
            Next j
        Next i
    End With
End Sub

Private Sub txtNamaSirkuler_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If LvSirkuler.Visible = True Then If LvSirkuler.ListItems.Count > 0 Then LvSirkuler.SetFocus
        Case vbKeyEscape
            LvSirkuler.Visible = False
    End Select
End Sub

Private Sub txtNamaSirkuler_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 13 Then
        If LvSirkuler.Visible = True Then
            LvSirkuler.SetFocus
        Else
            LvSirkuler.Visible = False
            txtIsiTM.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        LvSirkuler.Visible = False
    End If
    If KeyAscii = 39 Then KeyAscii = 0
Exit Sub
hell:
End Sub

Private Sub lvSirkuler_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim blnSelected As Boolean
    If Item.Checked = True Then
        LvSirkuler.ListItems(Item.Key).ForeColor = vbBlue
        subJmlTotalSirkuler = subJmlTotalSirkuler + 1
        ReDim Preserve subKdSirkuler(subJmlTotalSirkuler)
        subKdSirkuler(subJmlTotalSirkuler) = Item.Key
    Else
        LvSirkuler.ListItems(Item.Key).ForeColor = vbBlack
        blnSelected = False
        For i = 1 To subJmlTotalSirkuler
            If subKdSirkuler(i) = Item.Key Then blnSelected = True
            If blnSelected = True Then
                If i = subJmlTotalSirkuler Then
                    subKdSirkuler(i) = ""
                Else
                    subKdSirkuler(i) = subKdSirkuler(i + 1)
                End If
            End If
        Next i
        If subJmlTotalSirkuler < 1 Then
            subJmlTotalSirkuler = 0
        Else
            subJmlTotalSirkuler = subJmlTotalSirkuler - 1
        End If
    
    End If

    If subJmlTotalSirkuler = 0 Then
        txtNamaSirkuler.BackColor = &HFFFFFF
        ChkSirkuler.Caption = "Sirkuler"
    Else
        txtNamaSirkuler.BackColor = &HC0FFFF
        ChkSirkuler.Caption = "Sirkuler (" & subJmlTotalSirkuler & " org)"
    End If
End Sub


Private Sub dgDokterPembantu_DblClick()
    Call dgDokterPembantu_KeyPress(13)
End Sub

Private Sub dgDokterPembantu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokterPembantu = 0 Then Exit Sub
        txtDokterPembantu.Text = dgDokterPembantu.Columns(1).value
        subKdDokterPembantu = dgDokterPembantu.Columns(0).value

        fraDokterPembantu.Visible = False
        chkPerawat.SetFocus
    End If
    If KeyAscii = 27 Then
        fraDokterPembantu.Visible = False
    End If
End Sub

Private Sub fgPelayanan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then Call cmdHapus_Click
End Sub

Private Sub txtDokterDelegasi_Change()
    strPilihGrid = "DokterD"
    fraDokter.Visible = True
    Call subLoadDokterDelegasi
End Sub

Private Sub txtDokterDelegasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkPerawat.SetFocus
    If KeyAscii = 27 Then
        fraDokter.Visible = False
    End If
End Sub

Private Sub txtDokterPembantu_Change()
On Error Resume Next
    strFilterDokter = "WHERE NamaDokter like '%" & txtDokterPembantu.Text & "%'"
    subKdDokterPembantu = ""
    fraDokterPembantu.Visible = True
    Call subLoadDokterPembantu
End Sub

'untuk meload data dokter pembantu di grid
Private Sub subLoadDokterPembantu()
    On Error Resume Next
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & strFilterDokter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDokterPembantu = rs.RecordCount
    With dgDokterPembantu
        Set .DataSource = rs
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
    fraDokterPembantu.Left = 5040
    fraDokterPembantu.Top = 4200
End Sub

Private Sub dgDokterAnestesi_DblClick()
    Call dgDokterAnestesi_KeyPress(13)
End Sub

'untuk meload data dokter delegasi di grid
Private Sub subLoadDokterDelegasi()
'    On Error Resume Next
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter WHERE NamaDokter like '%" & txtDokterDelegasi.Text & "%'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    mintJmlDokter = rs.RecordCount
    With dgDokter
        Set .DataSource = rs
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
    fraDokter.Left = 3480
    fraDokter.Top = 3360
End Sub

Private Sub lvPemeriksa_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim blnSelected As Boolean
    If Item.Checked = True Then
        subJmlTotal = subJmlTotal + 1
        ReDim Preserve subKdPemeriksa(subJmlTotal)
        subKdPemeriksa(subJmlTotal) = Item.Key
    Else
        blnSelected = False
        For i = 1 To subJmlTotal
            If subKdPemeriksa(i) = Item.Key Then blnSelected = True
            If blnSelected = True Then
                If i = subJmlTotal Then
                    subKdPemeriksa(i) = ""
                Else
                    subKdPemeriksa(i) = subKdPemeriksa(i + 1)
                End If
            End If
        Next i
        subJmlTotal = subJmlTotal - 1
    End If
    
    If subJmlTotal = 0 Then
        txtNamaPerawat.BackColor = &HFFFFFF
        chkPerawat.Caption = "Paramedis/Penata"
    Else
        txtNamaPerawat.BackColor = &HC0FFFF
        chkPerawat.Caption = "Paramedis/Penata (" & subJmlTotal & " org)"
    End If
End Sub
Private Sub lvPemeriksa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 27 Then lvPemeriksa.Visible = False: txtNamaPerawat.SetFocus
End Sub

Private Sub txtDokterAnestesi_Change()
    strFilterDokter = "WHERE NamaDokter like '%" & txtDokterAnestesi.Text & "%'"
    subKdDokterAnestesi = ""
    fraDokterAnestesi.Visible = True
    Call subLoadDokterAnestesi
End Sub

Private Sub txtDokterAnestesi_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 13 Then
        If intJmlDokterAnestesi = 0 Then txtDokterPembantu.SetFocus
        If fraDokterAnestesi.Visible = True Then
            dgDokterAnestesi.SetFocus
        Else
            txtDokterPembantu.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        fraDokterAnestesi.Visible = False
    End If
    If KeyAscii = 39 Then KeyAscii = 0
hell:
End Sub

Private Sub txtDokterPembantu_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 13 Then
        If intJmlDokterPembantu = 0 Then chkPerawat.SetFocus
        If fraDokterPembantu.Visible = True Then
            dgDokterPembantu.SetFocus
        Else
            chkDokterResusitasi.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        fraDokterPembantu.Visible = False
    End If
    If KeyAscii = 39 Then KeyAscii = 0
hell:
End Sub

Private Sub txtIsiTM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dgPelayananRS.Visible = True Then
            dgPelayananRS.SetFocus
        Else
            txtKuantitas.SetFocus
        End If
    End If
End Sub

Private Sub txtKuantitas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If (Frame2.Enabled = True) Then
            optCito(1).SetFocus
        Else
            cmdTambah.SetFocus
        End If
    End If
End Sub

Private Sub txtNamaPerawat_Change()
On Error GoTo errLoad

    Call subLoadListPemeriksa("where [Nama Pemeriksa] LIKE '%" & txtNamaPerawat.Text & "%'")
    lvPemeriksa.Visible = True

Exit Sub
errLoad:
    Call msubPesanError
End Sub
Private Sub subLoadListPemeriksa(Optional strKriteria As String)
Dim strKey As String
    
    strSQL = "select * from v_daftarpemeriksapasien " & strKriteria & " order by [Nama Pemeriksa]"
    Call msubRecFO(rs, strSQL)
    
    With lvPemeriksa
        .ListItems.Clear
        For i = 0 To rs.RecordCount - 1
            strKey = "key" & rs(0).value
            .ListItems.Add , strKey, rs(1).value
            rs.MoveNext
        Next
        .Top = 3360
        .Left = 6480
        .Height = 1815
        .ColumnHeaders.Item(1).Width = lvPemeriksa.Width - 500
        
        If subJmlTotal = 0 Then Exit Sub
        For i = 1 To .ListItems.Count
            For j = 1 To subJmlTotal
                If .ListItems(i).Key = subKdPemeriksa(j) Then .ListItems(i).Checked = True
            Next j
        Next i
    End With
End Sub

Private Sub txtNamaPerawat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If lvPemeriksa.Visible = True Then
            lvPemeriksa.SetFocus
        Else
            optCito(1).SetFocus
        End If
    End If
End Sub
Private Sub txtNamaPerawat_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If lvPemeriksa.Visible = True Then If lvPemeriksa.ListItems.Count > 0 Then lvPemeriksa.SetFocus
        Case vbKeyEscape
            lvPemeriksa.Visible = False
    End Select
End Sub
Private Sub optCito_Click(index As Integer)
    If index = 0 Then
        strCito = "1"
    Else
        strCito = "0"
    End If
End Sub

'untuk meload data dokter anestesi di grid
Private Sub subLoadDokterAnestesi()
    On Error Resume Next
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & strFilterDokter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDokterAnestesi = rs.RecordCount
    With dgDokterAnestesi
        Set .DataSource = rs
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
    fraDokterAnestesi.Left = 5760
    fraDokterAnestesi.Top = 4200
End Sub

Private Sub dgDokterAnestesi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokterAnestesi = 0 Then Exit Sub
        txtDokterAnestesi.Text = dgDokterAnestesi.Columns(1).value
        subKdDokterAnestesi = dgDokterAnestesi.Columns(0).value

        fraDokterAnestesi.Visible = False
        txtDokterPembantu.SetFocus
    End If
    If KeyAscii = 27 Then
        fraDokterAnestesi.Visible = False
    End If
End Sub

Private Sub txtDokter2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If fraDokter.Visible = True Then dgDokter.SetFocus
    If KeyAscii = 27 Then
        fraDokter.Visible = False
    End If
End Sub

Private Sub txtDokter2_Change()
    strPilihGrid = "Dokter2"
    fraDokter.Visible = True
    Call subLoadDokter2
End Sub

'untuk meload data dokter delegasi di grid
Private Sub subLoadDokter2()
On Error GoTo errLoad
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter WHERE NamaDokter like '%" & txtDokter2.Text & "%'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    mintJmlDokter = rs.RecordCount
    With dgDokter
        Set .DataSource = rs
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
    fraDokter.Left = 3480
    fraDokter.Top = 4200
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub chkDilayaniDokter_Click()
On Error GoTo errLoad
    
    If chkDilayaniDokter.value = 0 Then
        txtDokter.Enabled = False
        txtDokter.Text = ""
        
        If fraDokter.Visible = True Then fraDokter.Visible = False
    Else
        lvPemeriksa.Visible = False
        
        txtDokter.Enabled = True
'        strSQL = "SELECT dbo.RegistrasiRJ.IdDokter, dbo.DataPegawai.NamaLengkap " & _
'            " FROM dbo.RegistrasiRJ INNER JOIN dbo.DataPegawai ON dbo.RegistrasiRJ.IdDokter = dbo.DataPegawai.IdPegawai " & _
'            " WHERE (dbo.RegistrasiRJ.NoPendaftaran = '" & mstrNoPen & "')"
'        Call msubRecFO(rs, strSQL)
'
'        strSQLCari = "Select KdJenisPegawai From DataPegawai Where IdPegawai='" & rs(0).Value & "'"
'        Call msubRecFO(rsB, strSQLCari)
'
'        If rsB(0).Value = "001" Then
'        If Not rs.EOF Then
'            txtDokter.Text = rs(1).Value
'            mstrKdDokter = rs(0).Value
'            intJmlDokter = rs.RecordCount
'            fraDokter.Visible = False
'        End If
'        End If
    End If
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub chkDilayaniDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkDilayaniDokter.value = 0 Then
            If chkDelegasi.Enabled = True Then
            chkDelegasi.SetFocus
            Else
            chkPerawat.SetFocus
            End If
        Else
            txtDokter.SetFocus
        End If
    End If
End Sub

Public Sub cmdHapus_Click()
With fgPelayanan
If .Row = .Rows Then Exit Sub
If .Row = 0 Then Exit Sub
If .TextMatrix(.Row, 0) = "" Then Exit Sub
Call msubRemoveItem(fgPelayanan, .Row)
Call HitungTotal
End With
End Sub

Private Sub dcDokterPerujuk_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    With chkDibayardimuka
    .SetFocus
    mstrKdDokterPerujuk = dcDokterPerujuk.BoundText
    End With
    
End If
End Sub

Private Sub dcInstalasi_GotFocus()
On Error GoTo errLoad
Dim tempKode As String

    tempKode = dcInstalasi.BoundText
    strSQL = "select distinct KdInstalasi,NamaInstalasi from V_RuanganTujuanRujukan WHERE kdinstalasi <> '04' and StatusEnabled='1' "
    Call msubDcSource(dcInstalasi, rs, strSQL)
    dcInstalasi.BoundText = tempKode
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    With dcRuangan
    .SetFocus
    End With
End Sub

Private Sub dcRuangan_GotFocus()
On Error GoTo errLoad
Dim tempKode As String
    
    If mstrFormPengirim <> "frmDaftarPasienRJRIIGD" Then
        tempKode = dcRuangan.BoundText
        strSQL = "select distinct KdRuangan,NamaRuangan from V_RuanganTujuanRujukan where KdInstalasi='" & dcInstalasi.BoundText & "' and KdRuangan <> '" & mstrKdRuangan & "' and StatusEnabled='1'  order by NamaRuangan "
        Call msubDcSource(dcRuangan, rs, strSQL)
        dcRuangan.BoundText = tempKode
    Else
        tempKode = dcRuangan.BoundText
        strSQL = "select distinct KdRuangan,NamaRuangan from V_RuanganTujuanRujukan where KdInstalasi='" & dcInstalasi.BoundText & "' and StatusEnabled='1' order by NamaRuangan "
        Call msubDcSource(dcRuangan, rs, strSQL)
        dcRuangan.BoundText = tempKode
    End If
Exit Sub
errLoad:
    Call msubPesanError
End Sub
Private Sub dcRuangan_KeyPress(KeyAscii As Integer)
With dcDokterPerujuk
    If (chkDibayardimuka.Enabled = True) Then
        'chandra 27 02 2014
        ' kebutuhan bayar di muka buat apa
        txtDokter.SetFocus
    End If
End With
End Sub

Public Sub Add_HistoryLoginActivity(strNamaObjekDB)
On Error GoTo hell_
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdAplikasi", adChar, adParamInput, 3, strKdAplikasi)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("TglActivity", adDate, adParamInput, , Format(Now, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("HostName", adVarChar, adParamInput, 50, strNamaHostLocal)
        .Parameters.Append .CreateParameter("NamaObjekDB", adVarChar, adParamInput, 200, strNamaObjekDB)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_HistoryLoginActivity"
        .CommandType = adCmdStoredProc
        .Execute
    
        If .Parameters("RETURN_VALUE").value <> 0 Then
            MsgBox "Ada Kesalahan dalam Hapus Rekap Komponen Biaya Pelayanan", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    
Exit Sub
hell_:
     Call msubPesanError("-Add_HistoryLoginActivity")
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcRuanganTujuan_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 13 Then
        If dcRuangan.MatchedWithList = True Then fgPelayanan.SetFocus
        strSQL = "select KdRuangan, NamaRuangan from Ruangan WHERE (NamaRuangan LIKE '%" & dcRuangan.Text & "%') and StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcRuangan.BoundText = rs(0).value
        dcRuangan.Text = rs(1).value
    End If
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub HitungTotal()
Dim total As Currency
Dim i As Integer
total = 0
txtTotalBayar.Text = ""
    If fgPelayanan.Rows <> 0 Then
    For i = 1 To fgPelayanan.Rows - 1
        total = total + IIf(fgPelayanan.TextMatrix(i, 4) = "", 0, fgPelayanan.TextMatrix(i, 4))
        If i = fgPelayanan.Rows - 1 Then
            txtTotalBayar.Text = Format(total, "###,###")
        End If
    Next i
    Else
        txtTotalBayar.Text = 0
    End If
End Sub


Private Sub dcRuanganTujuanTM_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 13 Then
        If dcRuangan.MatchedWithList = True Then fgPelayanan.SetFocus
        strSQL = "select KdRuangan, NamaRuangan from Ruangan WHERE (NamaRuangan LIKE '%" & dcRuangan.Text & "%') and StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcRuangan.BoundText = rs(0).value
        dcRuangan.Text = rs(1).value
    End If
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dgPelayananRS_DblClick()
On Error GoTo gabril
    txtKdIsiTM.Text = dgPelayananRS.Columns(1).value
    txtIsiTM.Text = dgPelayananRS.Columns(2).value
    dgPelayananRS.Visible = False
    txtIsiTM.SetFocus
Exit Sub
gabril:
    MsgBox "Data Pelayanan Tidak di temukan"
    txtIsiTM.Text = ""
    txtIsiTM.SetFocus
End Sub

Private Sub subLoadPelayananPerInstrumen()

    With fgInstrumenPerPelayanan
        For i = 1 To subJmlTotalInstrumen
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = mstrNoPen
            .TextMatrix(.Rows - 1, 1) = dcRuangan.BoundText
            .TextMatrix(.Rows - 1, 2) = dTglPlyn
            .TextMatrix(.Rows - 1, 3) = dgPelayananRS.Columns("KdPelayananRS")
            .TextMatrix(.Rows - 1, 4) = Mid(subKdInstrumen(i), 4, Len(subKdInstrumen(i)) - 3)
            .TextMatrix(.Rows - 1, 5) = strIDPegawaiAktif
        Next
    End With
'    subJmlTotalInstrumen = 0
'    txtNamaAsisten.BackColor = &HFFFFFF
'    ReDim Preserve subKdInstrumen(subJmlTotalInstrumen)
'
'    chkInstrumen.Value = vbUnchecked

End Sub

Private Sub subLoadPelayananPerSirkuler()

    With fgSirkulerPerPelayanan
        For i = 1 To subJmlTotalSirkuler
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = mstrNoPen
            .TextMatrix(.Rows - 1, 1) = dcRuangan.BoundText
            .TextMatrix(.Rows - 1, 2) = dTglPlyn
            .TextMatrix(.Rows - 1, 3) = dgPelayananRS.Columns("KdPelayananRS")
            .TextMatrix(.Rows - 1, 4) = Mid(subKdSirkuler(i), 4, Len(subKdSirkuler(i)) - 3)
            .TextMatrix(.Rows - 1, 5) = strIDPegawaiAktif
        Next
    End With
'    subJmlTotalSirkuler = 0
'    txtNamaAsisten.BackColor = &HFFFFFF
'    ReDim Preserve subKdSirkuler(subJmlTotalSirkuler)
'
'    ChkSirkuler.Value = vbUnchecked

End Sub

Private Sub subLoadPelayananPerAsisten()
    With fgAsistenPerPelayanan
        For i = 1 To subJmlTotalAsisten
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = mstrNoPen
            .TextMatrix(.Rows - 1, 1) = dcRuangan.BoundText
            .TextMatrix(.Rows - 1, 2) = dTglPlyn
            .TextMatrix(.Rows - 1, 3) = dgPelayananRS.Columns("KdPelayananRS")
            .TextMatrix(.Rows - 1, 4) = Mid(subKdAsisten(i), 4, Len(subKdAsisten(i)) - 3)
            .TextMatrix(.Rows - 1, 5) = strIDPegawaiAktif
        Next
    End With
'    subJmlTotalAsisten = 0
'    txtNamaAsisten.BackColor = &HFFFFFF
'    ReDim Preserve subKdAsisten(subJmlTotalAsisten)
'
'    chkAsisten.Value = vbUnchecked

End Sub

Private Function sp_Take_TarifBPT(f_KdPelayanan As String) As Currency
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, f_KdPelayanan)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, TempKodeKelas)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, "01")
        .Parameters.Append .CreateParameter("TarifCito", adCurrency, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("TarifTotal", adCurrency, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("StatusCito", adChar, adParamInput, 1, IIf(optCito(0).value = True, "Y", "T"))
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, IIf(chkDilayaniDokter.value = vbChecked, mstrKdDokter, Null))
        .Parameters.Append .CreateParameter("IdDokter2", adChar, adParamInput, 10, IIf(subKdDokterAnestesi = "", IIf(chkDilayaniDokter.value = vbChecked, mstrKdDokter, Null), subKdDokterAnestesi))
'        .Parameters.Append .CreateParameter("IdDokter2", adChar, adParamInput, 10, IIf(subKdDokterAnestesi = "", Null, subKdDokterAnestesi))
        .Parameters.Append .CreateParameter("IdDokter3", adChar, adParamInput, 10, IIf(subKdDokterPembantu = "", Null, subKdDokterPembantu))
        .Parameters.Append .CreateParameter("StatusDokterBersama", adChar, adParamInput, 1, "T")
        'txtDokterAnestesi
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Take_TarifBPT_DB"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam Pengambilan biaya tarif", vbExclamation, "Validasi"
            sp_Take_TarifBPT = 0
            subcurTarifBiayaSatuan = 0
        Else
            sp_Take_TarifBPT = .Parameters("TarifCito").value
            subcurTarifBiayaSatuan = .Parameters("TarifTotal").value
            Call Add_HistoryLoginActivity("Take_TarifBPT_DB")
        End If
    
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub subLoadPelayananPerPerawat()

    With fgPerawatPerPelayanan
        For i = 1 To subJmlTotal
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = mstrNoPen
            .TextMatrix(.Rows - 1, 1) = dcRuangan.BoundText
            .TextMatrix(.Rows - 1, 2) = dTglPlyn
            .TextMatrix(.Rows - 1, 3) = dgPelayananRS.Columns("KdPelayananRS")
            .TextMatrix(.Rows - 1, 4) = Mid(subKdPemeriksa(i), 4, Len(subKdPemeriksa(i)) - 3)
            .TextMatrix(.Rows - 1, 5) = strIDPegawaiAktif
        Next
    End With

'    subJmlTotal = 0
'    txtNamaPerawat.BackColor = &HFFFFFF
'    ReDim Preserve subKdPemeriksa(subJmlTotal)
'    chkPerawat.Value = vbUnchecked

End Sub

Private Sub dgPelayananRS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgPelayananRS_DblClick
    If KeyAscii = 27 Then dgPelayananRS.Visible = False: fgPelayanan.SetFocus
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 27 Then dgPelayananRS.Visible = False
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
Dim strSQLDokter As String
Dim rsDokter As New ADODB.recordset
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    
    chkDibayardimuka.value = 1
'    txtDokter.SetFocus
    txtNoCM.Visible = False
    txtNoPendaftaran.Visible = False
    txtNoOrder.Visible = False
    
    dtpTglRujuk.value = Now
    dtpTglOrderTM.value = Now
    txtTglKonsul.Text = Now
    Call subSetGrid
    
    
    

    Call subLoadDcSource
    dgDokterTM.Visible = False
    mstrFilterDokter = ""

   
'   txtDokter.Enabled = False
   txtDokter2.Enabled = False
   txtNamaPerawat.Enabled = False
   lvPemeriksa.Visible = False
   strCito = "0"
   fraPDokter.Visible = False
   fraButton.Visible = False
   Call subSetGridPerawatPerPelayanan
   Call subSetGridAsistenPerPelayanan
   Call subSetGridInstrumenPerPelayanan
   Call subSetGridSirkulerPerPelayanan
   Call subLoadTempatRujukan
   
    Call chkAsisten_Click
    Call chkInstrumen_Click
    Call ChkSirkuler_Click
    
    chkDibayardimuka_Click
Exit Sub
errLoad:
    Call msubPesanError

End Sub
Private Sub subLoadTempatRujukan()
    On Error GoTo errLoad
    Dim tempKode As String
    mstrKdInstalasiPerujuk = mstrKdInstalasi

    tempKode = mstrKdAsalRujukan
    If tempKode = "08" Or tempKode = "09" Or tempKode = "10" Or tempKode = "01" Or tempKode = "06" Or tempKode = "07" Or tempKode = "11" Or tempKode = "12" Then
        strSQL = "Select KdRuangan,NamaRuangan From Ruangan where StatusEnabled='1'"
        Call msubDcSource(dcTempatPerujuk, rs, strSQL)
    ElseIf tempKode = "02" Or tempKode = "03" Or tempKode = "04" Or tempKode = "05" Then
        strSQL = "select KdDetailRujukanAsal,DetailRujukanAsal from dbo.DetailRujukanAsal where KdRujukanAsal = '" & tempKode & "' and StatusEnabled='1'"
        Call msubDcSource(dcTempatPerujuk, rs, strSQL)
    End If
    If rs.EOF = False Then dcTempatPerujuk.Text = rs(1)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcDiagnosa_GotFocus()
    strSQL = "SELECT NamaDiagnosa FROM Diagnosa where StatusEnabled='1' ORDER BY NamaDiagnosa"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcDiagnosa.RowSource = rs
    dcDiagnosa.ListField = rs.Fields(0).Name
    Set rs = Nothing
End Sub

Private Sub dcDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtIsiTM.SetFocus
End Sub

Private Sub dcDiagnosa_LostFocus()
    dcDiagnosa = StrConv(dcDiagnosa, vbProperCase)
End Sub

Private Sub dcNamaPerujuk_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcNamaPerujuk.BoundText
    strSQL = "SELECT NamaDokter FROM V_DaftarDokter ORDER BY NamaDokter"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcNamaPerujuk.RowSource = rs
    dcNamaPerujuk.ListField = rs.Fields(0).Name
    Set rs = Nothing
    dcNamaPerujuk.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcNamaPerujuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dcDiagnosa.SetFocus
    Set rs = Nothing
    strSQL = "SELECT KodeDokter,NamaDokter FROM V_DaftarDokter where NamaDokter like '%" & dcNamaPerujuk.Text & "%'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdDokterPerujuk = rs.Fields(0)
        mstrNama = rs.Fields(1)
    End If
End Sub

Private Sub dcNamaPerujuk_LostFocus()
    dcNamaPerujuk = StrConv(dcNamaPerujuk, vbProperCase)
End Sub

Private Sub dcRujukanAsal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcTempatPerujuk.SetFocus
End Sub

Private Sub dcTempatPerujuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dcNamaPerujuk.SetFocus
End Sub

Private Sub dcTempatPerujuk_LostFocus()
    dcTempatPerujuk = StrConv(dcTempatPerujuk, vbProperCase)
End Sub


Private Sub subSetGridAsistenPerPelayanan()
    With fgAsistenPerPelayanan
        .Cols = 6
        .Rows = 1
        
        .MergeCells = flexMergeFree
        
        .TextMatrix(0, 0) = "NoPendaftaran"
        .TextMatrix(0, 1) = "Kode Ruangan"
        .TextMatrix(0, 2) = "Tgl Pelayanan"
        .TextMatrix(0, 3) = "Kode Pelayanan"
        .TextMatrix(0, 4) = "IdPegawai"
        .TextMatrix(0, 5) = "IdUser"
    
    End With
End Sub

Private Sub subSetGridInstrumenPerPelayanan()
    With fgInstrumenPerPelayanan
        .Cols = 6
        .Rows = 1
        
        .MergeCells = flexMergeFree
        
        .TextMatrix(0, 0) = "NoPendaftaran"
        .TextMatrix(0, 1) = "Kode Ruangan"
        .TextMatrix(0, 2) = "Tgl Pelayanan"
        .TextMatrix(0, 3) = "Kode Pelayanan"
        .TextMatrix(0, 4) = "IdPegawai"
        .TextMatrix(0, 5) = "IdUser"
    
    End With
End Sub

Private Sub subSetGridSirkulerPerPelayanan()
    With fgSirkulerPerPelayanan
        .Cols = 6
        .Rows = 1
        
        .MergeCells = flexMergeFree
        
        .TextMatrix(0, 0) = "NoPendaftaran"
        .TextMatrix(0, 1) = "Kode Ruangan"
        .TextMatrix(0, 2) = "Tgl Pelayanan"
        .TextMatrix(0, 3) = "Kode Pelayanan"
        .TextMatrix(0, 4) = "IdPegawai"
        .TextMatrix(0, 5) = "IdUser"
    
    End With
End Sub

Private Sub subSetGridPerawatPerPelayanan()
    With fgPerawatPerPelayanan
        .Cols = 6
        .Rows = 1
        
        .MergeCells = flexMergeFree
        
        .TextMatrix(0, 0) = "NoPendaftaran"
        .TextMatrix(0, 1) = "Kode Ruangan"
        .TextMatrix(0, 2) = "Tgl Pelayanan"
        .TextMatrix(0, 3) = "Kode Pelayanan"
        .TextMatrix(0, 4) = "IdPegawai"
        .TextMatrix(0, 5) = "IdUser"
    
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

   
    If txtNamaForm.Text = "frmRegistrasiRJPenunjang" Then
        frmRegistrasiRJPenunjang.Enabled = True
    End If
End Sub

Private Sub cmdBatal_Click()
    Call subKosong
End Sub
Private Sub subKosong()
        
        chkDilayaniDokter.value = 0
        chkDelegasi.value = 0
        chkPerawat.value = 0
        chkDokter2.value = 0
        chkDokterResusitasi.value = 0
        chkAsisten.value = 0
        chkInstrumen.value = 0
        ChkSirkuler.value = 0
        chkDibayardimuka.value = 0
        txtDokterAnestesi.Text = ""
        txtDokterPembantu.Text = ""
        dcJenisOperasi.BoundText = ""
        dcInstalasi.BoundText = ""
        dcRuangan.BoundText = ""
        fraDokter.Visible = False
        fraDokterAnestesi.Visible = False
        fraDokterPembantu.Visible = False
        Call subSetGrid
'        For i = 1 To lvPemeriksa.ListItems.Count
'            lvPemeriksa.ListItems(i).Checked = False
'        Next i
'        For i = 1 To lvAsisten.ListItems.Count
'                lvAsisten.ListItems(i).Checked = False
'        Next i
'        For i = 1 To LvInstrumen.ListItems.Count
'                LvInstrumen.ListItems(i).Checked = False
'        Next i
'        For i = 1 To LvSirkuler.ListItems.Count
'                LvSirkuler.ListItems(i).Checked = False
'        Next i

        Call subSetGridPerawatPerPelayanan
        Call subSetGridAsistenPerPelayanan
        Call subSetGridInstrumenPerPelayanan
        Call subSetGridSirkulerPerPelayanan
        txtIsiTM.Text = ""
        dgPelayananRS.Visible = False
        Call HitungTotal
        cmdSimpan.Enabled = True
        dtpTglOrderTM.value = Now
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If chkDibayardimuka.value = 1 Then
        If mintJmlDokter = 0 Then dtpTglRujuk.SetFocus 'Exit Sub
        If fraDokter.Visible = True Then
            dgDokter.SetFocus
        Else
            If chkDelegasi.Enabled = True Then
            chkDelegasi.SetFocus
            Else
            chkPerawat.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        fraDokter.Visible = False
    End If
        If fraDokter.Visible = True Then
            dgDokter.SetFocus
        Else
            dcTempatPerujuk.SetFocus
        End If

Else
        fraDokter.Visible = True
        If fraDokter.Visible = True Then
            dgDokter.SetFocus
        Else
            dcTempatPerujuk.SetFocus
        End If
'    txtIsiTM.SetFocus
End If
End Sub

Private Sub txtDokter_Change()
    strPilihGrid = "Dokter"
    mstrFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
'    mstrKdDokter = ""
    fraDokter.Visible = True
    Call subLoadDokter
End Sub


Private Sub subSetGrid()
On Error GoTo errLoad
  
    With fgPelayanan
        With fgPelayanan
        .Clear
        .Rows = 2
        .Cols = 25
        .TextMatrix(0, 0) = "Kode Pelayanan"
        .TextMatrix(0, 1) = "Nama Pelayanan"
        .TextMatrix(0, 2) = "Jumlah"
        .TextMatrix(0, 3) = "Biaya Satuan"
        .TextMatrix(0, 4) = "Biaya Total"
        .TextMatrix(0, 5) = "Tgl Berlaku"
        .TextMatrix(0, 6) = "Kode Dokter"
        .TextMatrix(0, 7) = "Status CITO"
        .TextMatrix(0, 8) = "Biaya CITO"
        .TextMatrix(0, 9) = "Tanggal Pelayanan"
        .TextMatrix(0, 10) = "KodeLabLuar"
        .TextMatrix(0, 11) = "DokterDelegasi"
        .TextMatrix(0, 12) = "StatusOrder" 'for pesan pelayanan
        
        .TextMatrix(0, 13) = "KodeDokterOperator"
        .TextMatrix(0, 14) = "KodeDokterPembantu"
        .TextMatrix(0, 15) = "KodeDokterAnestesi"
        .TextMatrix(0, 16) = "KodeDokterOperator2"
        .TextMatrix(0, 17) = "KodeDokterResus"
        
        .TextMatrix(0, 18) = "KdRuanganTujuan"
        .TextMatrix(0, 19) = "IdDokterPerujuk"
        .TextMatrix(0, 20) = "TglOrder"
        .TextMatrix(0, 21) = "KdJenisOperasi"
        .TextMatrix(0, 22) = "StatusDibayarDimuka"
        .TextMatrix(0, 23) = "StatusDelegasi"
        .TextMatrix(0, 24) = "StatusOperasiBersama"
        
        
        .ColWidth(0) = 0
        .ColWidth(1) = 4700
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 1500
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        .ColWidth(16) = 0
        .ColWidth(17) = 0
        
        .ColWidth(18) = 0
        .ColWidth(19) = 0
        .ColWidth(20) = 0
        .ColWidth(21) = 0
        .ColWidth(22) = 0
        .ColWidth(23) = 0
        .ColWidth(24) = 0
    End With
  End With
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDcSource()
On Error GoTo errLoad
' cek dokter penanggung jawab
'Dim strSQLa As String
'Dim rsa1 As New ADODB.recordset
'    strSQLa = "SELECT * FROM RegistrasiRJ WHERE NoPendaftaran='" & mstrNoPen & "'"
'    Call msubRecFO(rsa1, strSQLa)
'    If rsa1.EOF = False Then
'    Call msubDcSource(dcDokterPerujuk, rs, "Select KodeDokter, NamaDokter FROM V_DaftarDokter order by NamaDokter ")
'    If rs.EOF = False Then
'        If rsa1.Fields("IdDokter") = "-" Then
'           MsgBox "Dokter Penanggung Jawab Ruangan belum di isi", vbExclamation, "Warning"
'        Else
'          dcDokterPerujuk.BoundText = rsa1("IdDokter").Value
'        End If
'    Else
'           MsgBox "Dokter ruangan asal kosong ", vbExclamation, "Validasi"
'    End If
'    Else
'
'    End If
    
    strSQL = "select distinct KdInstalasi,NamaInstalasi from V_RuanganTujuanRujukan WHERE StatusEnabled='1' "
    Call msubDcSource(dcInstalasi, rs, strSQL)
    strSQL = "select distinct KdRuangan,NamaRuangan from V_RuanganTujuanRujukan WHERE StatusEnabled='1' "
    Call msubDcSource(dcRuangan, rs, strSQL)
    
    strSQL = "SELECT KdJenisOperasi,JenisOperasi FROM JenisOperasi where statusenabled=1"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcJenisOperasi.RowSource = rs
    dcJenisOperasi.BoundColumn = rs(0).Name
    dcJenisOperasi.ListField = rs(1).Name
    Set rs = Nothing
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter()
On Error GoTo errLoad

    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & mstrFilterDokter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    mintJmlDokter = rs.RecordCount
    With dgDokter
        Set .DataSource = rs
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
    fraDokter.Left = 360
    fraDokter.Top = 3360

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subHitungTotal()
On Error GoTo errLoad
Dim i As Integer

    If fgPelayanan.TextMatrix(fgPelayanan.Row - 1, 11) = "" Then Exit Sub
    txtTotalBiaya.Text = 0
    txtHutangPenjamin.Text = 0
    txtTanggunganRS.Text = 0
    txtHarusDibayar.Text = 0
    txtTotalDiscount.Text = 0
            
    With fgPelayanan
        For i = 1 To IIf(fgPelayanan.TextMatrix(fgPelayanan.Rows - 1, 2) = "", fgPelayanan.Rows - 2, fgPelayanan.Rows - 1)
            txtTotalBiaya.Text = txtTotalBiaya.Text + CDbl(.TextMatrix(i, 11))
            txtHutangPenjamin.Text = txtHutangPenjamin.Text + CDbl(.TextMatrix(i, 19))
            txtTanggunganRS.Text = txtTanggunganRS.Text + CDbl(.TextMatrix(i, 20))
            txtTotalDiscount.Text = txtTotalDiscount.Text + CDbl(.TextMatrix(i, 21))
            txtHarusDibayar.Text = txtHarusDibayar.Text + CDbl(.TextMatrix(i, 22))
        Next i
    End With
    
    txtTotalBiaya.Text = IIf(Val(txtTotalBiaya.Text) = 0, 0, Format(txtTotalBiaya.Text, "#,###"))
    txtHutangPenjamin.Text = IIf(Val(txtHutangPenjamin.Text) = 0, 0, Format(txtHutangPenjamin.Text, "#,###"))
    txtTanggunganRS.Text = IIf(Val(txtTanggunganRS.Text) = 0, 0, Format(txtTanggunganRS.Text, "#,###"))
    txtHarusDibayar.Text = IIf(Val(txtHarusDibayar.Text) = 0, 0, Format(txtHarusDibayar.Text, "#,###"))
    txtTotalDiscount.Text = IIf(Val(txtTotalDiscount.Text) = 0, 0, Format(txtTotalDiscount.Text, "#,###"))
    
    subcurHarusDibayar = txtHarusDibayar.Text
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtIsiTM_Change()
On Error GoTo hell

    strSQL = "SELECT DISTINCT TOP (200) [Jenis Pelayanan], KdPelayananRS, [Nama Pelayanan] as NamaPelayanan, Tarif" & _
            " FROM V_TarifPelayananTindakan where Kdruangan='" & dcRuangan.BoundText & "' AND [Nama Pelayanan] like'%" & txtIsiTM.Text & "%' AND KdPelayananRS IS NOT NULL and KdKelas='" & TempKodeKelas & "' order by NamaPelayanan"

    Call msubRecFO(dbRst, strSQL)
            
    Set dgPelayananRS.DataSource = dbRst
        With dgPelayananRS
            .Columns("Jenis Pelayanan").Width = 0
            .Columns("KdPelayananRS").Width = 0
            .Columns("NamaPelayanan").Width = 3500
'            .Top = 4950
'            .Left = 260
            .Visible = True

        End With

Exit Sub
hell:
    Call msubPesanError
End Sub

