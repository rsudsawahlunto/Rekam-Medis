VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAsuhanKeperawatan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Asuhan Keperawatan"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAsuhanKeperawatan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   9270
   Begin VB.TextBox txtKdParamedis2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   4440
      TabIndex        =   52
      Text            =   "txtKdParamedis"
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtKdParamedis3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   5760
      TabIndex        =   51
      Text            =   "txtKdParamedis"
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtKdDiagnosaKeperawatan 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   5760
      TabIndex        =   50
      Text            =   "txtKdDiagnosaKeperawatan"
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtKdParamedis1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   3120
      TabIndex        =   49
      Text            =   "txtKdParamedis"
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtNoPakai 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   3120
      TabIndex        =   48
      Text            =   "txtNoPakai"
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame fraParamedis 
      Caption         =   "Data Paramedis"
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
      Left            =   -4920
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   6015
      Begin MSDataGridLib.DataGrid dgParamedis 
         Height          =   2055
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   5775
         _ExtentX        =   10186
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
   End
   Begin VB.Frame fraDiagnosa 
      Caption         =   "Data Diagnosa Keperawatan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   -5760
      TabIndex        =   35
      Top             =   3360
      Visible         =   0   'False
      Width           =   6855
      Begin MSDataGridLib.DataGrid dgDiagnosaKeperawatan 
         Height          =   2775
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   6375
         _ExtentX        =   11245
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
   End
   Begin VB.Frame fraEvaluasi 
      Caption         =   "Evaluasi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   41
      Top             =   4440
      Width           =   9255
      Begin VB.TextBox txtParamedisEva 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2400
         TabIndex        =   16
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtEvaluasi 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   1080
         Width           =   9015
      End
      Begin MSComCtl2.DTPicker dtpTglEvaluasi 
         Height          =   330
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   388562947
         CurrentDate     =   38077
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Paramedis"
         Height          =   210
         Index           =   5
         Left            =   2400
         TabIndex        =   44
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Evaluasi"
         Height          =   210
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Evaluasi"
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Frame fraImplementasi 
      Caption         =   "Implementasi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   37
      Top             =   2880
      Width           =   9255
      Begin VB.TextBox txtImplementasi 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1080
         Width           =   9015
      End
      Begin VB.TextBox txtParamedisImp 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2400
         TabIndex        =   13
         Top             =   480
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpTglImplementasi 
         Height          =   330
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   118095875
         CurrentDate     =   38077
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Implementasi"
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   40
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Implementasi"
         Height          =   210
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1785
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Paramedis"
         Height          =   210
         Index           =   2
         Left            =   2400
         TabIndex        =   38
         Top             =   240
         Width           =   810
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
      Height          =   975
      Left            =   0
      TabIndex        =   26
      Top             =   960
      Width           =   9255
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Top             =   480
         Width           =   1335
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
         Left            =   6600
         TabIndex        =   27
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
            TabIndex        =   4
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
         Begin VB.TextBox txtHari 
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
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   555
            TabIndex        =   30
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1350
            TabIndex        =   29
            Top             =   270
            Width           =   210
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2130
            TabIndex        =   28
            Top             =   270
            Width           =   150
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   5280
         TabIndex        =   34
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   2880
         TabIndex        =   33
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1800
         TabIndex        =   32
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraRencana 
      Caption         =   "Rencana"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   23
      Top             =   1920
      Width           =   9255
      Begin VB.TextBox txtParamedisRen 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7320
         TabIndex        =   10
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtDiagnosaKeperawatan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2400
         TabIndex        =   8
         Top             =   480
         Width           =   4695
      End
      Begin MSComCtl2.DTPicker dtpTglAskep 
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   388562947
         CurrentDate     =   38077
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Paramedis"
         Height          =   210
         Index           =   1
         Left            =   7320
         TabIndex        =   36
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal AsKep"
         Height          =   210
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Diagnosa Keperawatan"
         Height          =   210
         Index           =   0
         Left            =   2400
         TabIndex        =   24
         Top             =   240
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   21
      Top             =   7800
      Width           =   9255
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   465
         Left            =   5280
         TabIndex        =   19
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   465
         Left            =   7320
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "F1 - Cetak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   46
         Top             =   360
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   0
      TabIndex        =   45
      Top             =   6000
      Width           =   9255
      Begin MSDataGridLib.DataGrid dgAsKep 
         Height          =   1455
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
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
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   47
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
      Left            =   7440
      Picture         =   "frmAsuhanKeperawatan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmAsuhanKeperawatan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmAsuhanKeperawatan.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmAsuhanKeperawatan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilterDiagnosa As String
Dim subKdDiagnosaKeperawatan As String
Dim subKdParamedisRen As String
Dim subKdParamedisImp As String
Dim subKdParamedisEva As String
Dim intJmlDiagnosa As Integer
Dim intJmlParamedis As Integer
Dim intJmlAsKep As Integer
Dim subParamedis As String
Dim subNoPakai As String

Private Sub clear()
    txtKdDiagnosaKeperawatan.Text = ""
    txtKdParamedis1.Text = ""
    txtKdParamedis2.Text = ""
    txtKdParamedis3.Text = ""
    txtDiagnosaKeperawatan.Text = ""
    txtParamedisEva.Text = ""
    txtParamedisRen.Text = ""
    txtParamedisImp.Text = ""
    txtEvaluasi.Text = ""
    txtImplementasi.Text = ""
    fraParamedis.Visible = False
    fraDiagnosa.Visible = False
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If Periksa("text", txtDiagnosaKeperawatan, "Diagnosa Keperawatan kosong") = False Then Exit Sub
    If Periksa("text", txtParamedisRen, "Paramedis Rencana AsKep kosong") = False Then Exit Sub
    If Periksa("text", txtParamedisImp, "Paramedis Implementasi AsKep kosong") = False Then Exit Sub
    If Periksa("text", txtParamedisEva, "Paramedis Evaluasi AsKep kosong") = False Then Exit Sub
    If Periksa("text", txtEvaluasi, "Evaluasi AsKep kosong") = False Then Exit Sub
    If Periksa("text", txtImplementasi, "Implementasi Rencana AsKep kosong") = False Then Exit Sub

    If sp_RencanaAskep("A") = False Then Exit Sub
    If sp_ImplementasiAskep("A") = False Then Exit Sub
    If sp_EvaluasiAskep("A") = False Then Exit Sub

    Call Add_HistoryLoginActivity("AUD_RencanaAsKep+AUD_ImplementasiAsKep+AUD_EvaluasiAsKep")

    Call clear
    Call subLoadGridSource
    txtDiagnosaKeperawatan.SetFocus
    Exit Sub
hell:
    msubPesanError
End Sub

Private Sub cmdTutup_Click()
    If txtDiagnosaKeperawatan.Text <> "" Then
        If MsgBox("Simpan data Asuhan Keperawatan pasien?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub dgAsKep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDiagnosaKeperawatan.SetFocus
End Sub

Private Sub dgAsKep_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next

    If intJmlAsKep = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        txtDiagnosaKeperawatan.SetFocus
        Exit Sub
    End If
    dtpTglAskep.value = dgAsKep.Columns(3).value
    dtpTglImplementasi.value = dgAsKep.Columns(8).value
    dtpTglEvaluasi.value = dgAsKep.Columns(12).value

    txtParamedisRen.Text = dgAsKep.Columns(6).value
    txtKdParamedis1.Text = dgAsKep.Columns(5).value
    fraParamedis.Visible = False

    txtParamedisEva.Text = dgAsKep.Columns(15).value
    txtKdParamedis3.Text = dgAsKep.Columns(14).value
    fraParamedis.Visible = False

    txtParamedisImp.Text = dgAsKep.Columns(11).value
    txtKdParamedis2.Text = dgAsKep.Columns(10).value
    fraParamedis.Visible = False

    txtEvaluasi.Text = dgAsKep.Columns(13).value
    txtImplementasi.Text = dgAsKep.Columns(9).value
    txtDiagnosaKeperawatan.Text = dgAsKep.Columns(4).value
    fraDiagnosa.Visible = False

    txtKdDiagnosaKeperawatan.Text = dgAsKep.Columns("KdDiagnosaKeperawatan").value
End Sub

Private Sub dgDiagnosaKeperawatan_DblClick()
    Call dgDiagnosaKeperawatan_KeyPress(13)
End Sub

Private Sub dgDiagnosaKeperawatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dgDiagnosaKeperawatan.ApproxCount = 0 Then Exit Sub
        txtKdDiagnosaKeperawatan.Text = dgDiagnosaKeperawatan.Columns("Kode Diagnosa Keperawatan")
        txtDiagnosaKeperawatan.Text = dgDiagnosaKeperawatan.Columns("Diagnosa Keperawatan")
        txtParamedisRen.SetFocus
        fraDiagnosa.Visible = False
    End If
End Sub

Private Sub dgParamedis_DblClick()
    Call dgParamedis_KeyPress(13)
End Sub

Private Sub dgParamedis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case subParamedis
            Case "ren"
                txtParamedisRen.Text = dgParamedis.Columns(1).value
                txtKdParamedis1 = dgParamedis.Columns(0).value
                dtpTglImplementasi.SetFocus
            Case "eva"
                txtParamedisEva.Text = dgParamedis.Columns(1).value
                txtKdParamedis2 = dgParamedis.Columns(0).value
                txtEvaluasi.SetFocus
            Case "imp"
                txtParamedisImp.Text = dgParamedis.Columns(1).value
                txtKdParamedis3 = dgParamedis.Columns(0).value
                txtImplementasi.SetFocus
        End Select
        subKdParamedis = dgParamedis.Columns(0).value
        fraParamedis.Visible = False
    End If
    If KeyAscii = 27 Then
        fraParamedis.Visible = False
    End If
End Sub

Private Sub dtpTglAskep_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtDiagnosaKeperawatan.SetFocus
End Sub

Private Sub dtpTglEvaluasi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtParamedisEva.SetFocus
End Sub

Private Sub dtpTglImplementasi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtParamedisImp.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            If dgAsKep.ApproxCount = 0 Then Exit Sub
            mstrNoPen = txtNoPendaftaran.Text
            frmCetakAsuhanKeperawatan.Show
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglAskep.value = Now
    dtpTglAskep.MaxDate = Now
    dtpTglEvaluasi.value = Now
    dtpTglEvaluasi.MaxDate = Now
    dtpTglImplementasi.value = Now
    dtpTglImplementasi.MaxDate = Now

    strSQL = "select NoPakai from PemakaianKamar where NoPendaftaran='" & mstrNoPen & "' and NoCM='" & mstrNoCM & "'"
    Call msubRecFO(rs, strSQL)
    txtNoPakai.Text = rs.Fields(0)

    Call clear
    Call subLoadGridSource
    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub txtDiagnosaKeperawatan_Change()
    fraDiagnosa.Visible = True
    Call subLoadDiagnosa
End Sub

Private Sub txtDiagnosaKeperawatan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If fraDiagnosa.Visible = False Then Exit Sub
        dgDiagnosaKeperawatan.SetFocus
    End If
End Sub

Private Sub txtDiagnosaKeperawatan_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    If KeyAscii = 13 Then
        If intJmlDiagnosa = 0 Then Exit Sub
        If fraDiagnosa.Visible = True Then
            dgDiagnosaKeperawatan.SetFocus
        Else
            txtParamedisRen.SetFocus
            Set rs = Nothing
            strSQL = "select KdDiagnosaKeperawatan from DiagnosaKeperawatan where DiagnosaKeperawatan='" & txtDiagnosaKeperawatan.Text & "' and StatusEnabled='1'" & strFilterDiagnosa
            Call msubRecFO(rs, strSQL)
            subKdDiagnosaKeperawatan = rs.Fields(0).value
        End If
    End If
    If KeyAscii = 27 Then
        fraDiagnosa.Visible = False
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDiagnosa()
    On Error GoTo hell
    Set rs = Nothing
    strSQL = "SELECT TOP 100 dbo.DiagnosaKeperawatan.DiagnosaKeperawatan, dbo.DiagnosaKeperawatan.KdDiagnosaKeperawatan" & _
    " FROM dbo.DiagnosaKeperawatan LEFT OUTER JOIN dbo.TujuanNRencanaTindakan ON dbo.DiagnosaKeperawatan.KdDiagnosaKeperawatan = dbo.TujuanNRencanaTindakan.KdDiagnosaKeperawatan" & _
    " GROUP BY dbo.DiagnosaKeperawatan.DiagnosaKeperawatan, dbo.DiagnosaKeperawatan.KdDiagnosaKeperawatan, dbo.DiagnosaKeperawatan.StatusEnabled " & _
    " HAVING (dbo.DiagnosaKeperawatan.DiagnosaKeperawatan LIKE '%" & txtDiagnosaKeperawatan.Text & "%') OR (dbo.DiagnosaKeperawatan.KdDiagnosaKeperawatan LIKE '%" & txtDiagnosaKeperawatan.Text & "%') and dbo.DiagnosaKeperawatan.StatusEnabled=1" & _
    " ORDER BY dbo.DiagnosaKeperawatan.DiagnosaKeperawatan"
    Call msubRecFO(rs, strSQL)
    intJmlDiagnosa = rs.RecordCount
    Set dgDiagnosaKeperawatan.DataSource = rs
    With dgDiagnosaKeperawatan
        .Columns(1).Caption = "Kode Diagnosa Keperawatan"
        .Columns(1).Width = 1200
        .Columns(0).Caption = "Diagnosa Keperawatan"
        .Columns(0).Width = 4550
    End With
    fraDiagnosa.Left = 2100
    fraDiagnosa.Top = 3000
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()
    On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "select * from V_AsuhanKeperawatan where NoPakai='" & txtNoPakai.Text & "'"
    rs.Open strSQL, dbConn, adOpenDynamic, adLockOptimistic
    intJmlAsKep = rs.RecordCount
    Set dgAsKep.DataSource = rs
    With dgAsKep
        .Columns(0).Width = 0       'NoPakai
        .Columns(1).Width = 0       'NoPendaftaran
        .Columns(2).Width = 0       'No CM

        .Columns(3).Width = 1900    'Tgl AsKep
        .Columns(4).Width = 3000    'Diagnosa Keperawatan

        .Columns(5).Width = 0       'Id Pegawai Rencana
        .Columns(6).Width = 0       'Id Pegawai Rencana
        .Columns(7).Width = 0       'Id Pegawai Rencana

        .Columns(8).Width = 1900    'Tgl Implementasi
        .Columns(9).Width = 3000    'Implementasi

        .Columns(10).Width = 0       'Id Pegawai Implementasi

        .Columns(11).Width = 0      'Id Pegawai Evaluasi

        .Columns(12).Width = 1900    'Tgl Evaluasi
        .Columns(13).Width = 3000   'Evaluasi
    End With
    Set rs = Nothing

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtEvaluasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtEvaluasi_LostFocus()
    Dim i As Integer
    Dim tempText As String
    tempText = Trim(txtEvaluasi.Text)
    txtEvaluasi.Text = ""
    For i = 1 To Len(tempText)
        If Asc(Mid(tempText, i, 1)) <> 10 And Asc(Mid(tempText, i, 1)) <> 13 Then
            txtEvaluasi.Text = txtEvaluasi.Text & Mid(tempText, i, 1)
        End If
    Next i
End Sub

Private Sub txtImplementasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglEvaluasi.SetFocus
End Sub

Private Sub txtImplementasi_LostFocus()
    Dim i As Integer
    Dim tempText As String
    tempText = Trim(txtImplementasi.Text)
    txtImplementasi.Text = ""
    For i = 1 To Len(tempText)
        If Asc(Mid(tempText, i, 1)) <> 10 And Asc(Mid(tempText, i, 1)) <> 13 Then
            txtImplementasi.Text = txtImplementasi.Text & Mid(tempText, i, 1)
        End If
    Next i
End Sub

Private Sub txtParamedisEva_Change()
    subParamedis = "eva"
    Call subLoadParamedis("where [Nama Pemeriksa] LIKE '%" & txtParamedisEva.Text & "%'")
    fraParamedis.Visible = True
    fraParamedis.Left = 2400
    fraParamedis.Top = 5320

End Sub

Private Sub txtParamedisEva_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If fraParamedis.Visible = False Then Exit Sub
        dgParamedis.SetFocus
    End If
End Sub

Private Sub txtParamedisEva_KeyPress(KeyAscii As Integer)
    On Error GoTo errorLahYaw
    If KeyAscii = 13 Then
        If intJmlParamedis = 0 Then Exit Sub
        If fraParamedis.Visible = True Then
            dgParamedis.SetFocus
        Else
            txtEvaluasi.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        fraParamedis.Visible = False
    End If
    Call SetKeyPressToChar(KeyAscii)
    Exit Sub
errorLahYaw:
End Sub

Private Sub txtParamedisImp_Change()
    subParamedis = "imp"
    Call subLoadParamedis("where [Nama Pemeriksa] LIKE '%" & txtParamedisImp.Text & "%'")
    fraParamedis.Visible = True
    fraParamedis.Left = 2400
    fraParamedis.Top = 3740

End Sub

Private Sub txtParamedisImp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If fraParamedis.Visible = False Then Exit Sub
        dgParamedis.SetFocus
    End If
End Sub

Private Sub txtParamedisImp_KeyPress(KeyAscii As Integer)
    On Error GoTo errorLahYaw
    If KeyAscii = 13 Then
        If intJmlParamedis = 0 Then Exit Sub
        If fraParamedis.Visible = True Then
            dgParamedis.SetFocus
        Else
            txtImplementasi.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        fraParamedis.Visible = False
    End If
    Call SetKeyPressToChar(KeyAscii)
    Exit Sub
errorLahYaw:
End Sub

Private Sub txtParamedisRen_Change()
    subParamedis = "ren"
    Call subLoadParamedis("where [Nama Pemeriksa] LIKE '%" & txtParamedisRen.Text & "%'")
    fraParamedis.Visible = True
    fraParamedis.Left = 3240
    fraParamedis.Top = 2790

End Sub

Private Sub txtParamedisRen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If fraParamedis.Visible = False Then Exit Sub
        dgParamedis.SetFocus
    End If
End Sub

Private Sub txtParamedisRen_KeyPress(KeyAscii As Integer)
    On Error GoTo errorLahYaw
    If KeyAscii = 13 Then
        If intJmlParamedis = 0 Then Exit Sub
        If fraParamedis.Visible = True Then
            dgParamedis.SetFocus
        Else
            dtpTglImplementasi.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        fraParamedis.Visible = False
    End If
    Call SetKeyPressToChar(KeyAscii)
    Exit Sub
errorLahYaw:
End Sub

Private Sub subLoadParamedis(Optional strKriteria As String)
    On Error Resume Next
    Set rs = Nothing
    strSQL = "select * from V_DaftarPemeriksaPasien " & strKriteria & " order by [Nama Pemeriksa]"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlParamedis = rs.RecordCount
    Set dgParamedis.DataSource = rs
    With dgParamedis
        .Columns(0).Width = 1300
        .Columns(1).Width = 2600
        .Columns(2).Width = 300
        .Columns(3).Width = 1000
    End With
    Set rs = Nothing
End Sub

Private Function sp_RencanaAskep(f_Status As String) As Boolean
    sp_RencanaAskep = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, txtNoPakai.Text)
        .Parameters.Append .CreateParameter("TglAskep", adDate, adParamInput, , Format(dtpTglAskep.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, mstrNoCM)
        .Parameters.Append .CreateParameter("KdDiagnosaKeperawatan", adVarChar, adParamInput, 10, txtKdDiagnosaKeperawatan.Text)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtKdParamedis1.Text)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_RencanaAsKep"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            If f_Status = "A" Then
                MsgBox "Gagal menyimpan data", vbCritical, "Validasi"
            Else
                MsgBox "Gagal menghapus data", vbCritical, "Validasi"
            End If
            sp_RencanaAskep = False
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_ImplementasiAskep(f_Status As String) As Boolean
    sp_ImplementasiAskep = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, txtNoPakai.Text)
        .Parameters.Append .CreateParameter("TglAskep", adDate, adParamInput, , Format(dtpTglAskep.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglImplementasi", adDate, adParamInput, , Format(dtpTglImplementasi.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Implementasi", adVarChar, adParamInput, 500, Trim(txtImplementasi.Text))
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtKdParamedis2.Text)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_ImplementasiAsKep"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            If f_Status = "A" Then
                MsgBox "Gagal menyimpan data", vbCritical, "Validasi"
            Else
                MsgBox "Gagal menghapus data", vbCritical, "Validasi"
            End If
            sp_ImplementasiAskep = False
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_EvaluasiAskep(f_Status As String) As Boolean
    sp_EvaluasiAskep = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, txtNoPakai.Text)
        .Parameters.Append .CreateParameter("TglAskep", adDate, adParamInput, , Format(dtpTglAskep.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglEvaluasi", adDate, adParamInput, , Format(dtpTglEvaluasi.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Evaluasi", adVarChar, adParamInput, 500, Trim(txtEvaluasi.Text))
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtKdParamedis3.Text)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_EvaluasiAsKep"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            If f_Status = "A" Then
                MsgBox "Gagal menyimpan data", vbCritical, "Validasi"
            Else
                MsgBox "Gagal menghapus data", vbCritical, "Validasi"
            End If
            sp_EvaluasiAskep = False
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

