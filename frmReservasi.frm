VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmReservasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Reservasi Pendaftaran"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReservasi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   13500
   Begin VB.Frame fraDokter 
      Caption         =   "Data Dokter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   9240
      TabIndex        =   30
      Top             =   6120
      Visible         =   0   'False
      Width           =   7815
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   1455
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   16
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
   End
   Begin VB.TextBox txtKet 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   120
      TabIndex        =   49
      Top             =   4440
      Width           =   8535
   End
   Begin VB.Frame fraRI 
      Height          =   975
      Left            =   0
      TabIndex        =   42
      Top             =   3240
      Width           =   13455
      Begin MSDataListLib.DataCombo dcKelasKamarRI 
         Height          =   360
         Left            =   120
         TabIndex        =   43
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcNoKamarRI 
         Height          =   360
         Left            =   3240
         TabIndex        =   44
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcNoBedRI 
         Height          =   360
         Left            =   5280
         TabIndex        =   45
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Kamar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "No. Kamar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   3240
         TabIndex        =   47
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "No. Bed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   5280
         TabIndex        =   46
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.TextBox TxtNoAntrian 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   12240
      MaxLength       =   15
      TabIndex        =   40
      Top             =   660
      Width           =   1095
   End
   Begin VB.Frame fraAntrian 
      Height          =   735
      Left            =   13440
      TabIndex        =   37
      Top             =   240
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox txtKdAntrian 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4080
         MaxLength       =   15
         TabIndex        =   38
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Kode Antrian"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   840
         TabIndex        =   39
         Top             =   120
         Width           =   2715
      End
   End
   Begin VB.TextBox txtNoBKM 
      Height          =   375
      Left            =   2640
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtNoPakai 
      Height          =   495
      Left            =   480
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   18
      Top             =   4920
      Width           =   13455
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
         Left            =   9720
         TabIndex        =   11
         Top             =   240
         Width           =   1695
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
         Left            =   11520
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Data Pemesanan"
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
      TabIndex        =   13
      Top             =   2160
      Width           =   13455
      Begin VB.TextBox txtKdDokter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   13680
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtTlp 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10920
         TabIndex        =   10
         Top             =   600
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   360
         Left            =   2400
         TabIndex        =   8
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   360
         Left            =   4680
         TabIndex        =   9
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpTglReservasi 
         Height          =   360
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
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
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   128581635
         CurrentDate     =   38061
      End
      Begin MSDataListLib.DataCombo dcDokter 
         Height          =   360
         Left            =   7200
         TabIndex        =   51
         Top             =   600
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Pemeriksa"
         Height          =   210
         Left            =   7200
         TabIndex        =   34
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "No. Telp"
         Height          =   210
         Left            =   10920
         TabIndex        =   32
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pesan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan Pemeriksaan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4680
         TabIndex        =   15
         Top             =   360
         Width           =   1905
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Instalasi Pemeriksaan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         TabIndex        =   14
         Top             =   360
         Width           =   1860
      End
   End
   Begin VB.Frame Frame1 
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
      TabIndex        =   16
      Top             =   1080
      Width           =   13455
      Begin VB.CheckBox chkNoCM 
         Caption         =   "No.CM"
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
         Height          =   240
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cbJenisKelamin 
         Height          =   330
         ItemData        =   "frmReservasi.frx":0CCA
         Left            =   7440
         List            =   "frmReservasi.frx":0CD4
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.Frame Frame4 
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
         Height          =   850
         Left            =   10680
         TabIndex        =   20
         Top             =   120
         Width           =   2655
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
            Height          =   360
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   6
            Top             =   330
            Width           =   375
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
            Height          =   360
            Left            =   960
            MaxLength       =   6
            TabIndex        =   5
            Top             =   330
            Width           =   375
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
            Left            =   120
            MaxLength       =   6
            TabIndex        =   4
            Top             =   330
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2280
            TabIndex        =   23
            Top             =   360
            Width           =   195
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1440
            TabIndex        =   22
            Top             =   360
            Width           =   270
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   600
            TabIndex        =   21
            Top             =   360
            Width           =   315
         End
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
         Height          =   360
         Left            =   240
         MaxLength       =   12
         TabIndex        =   0
         Top             =   600
         Width           =   1815
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
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   5175
      End
      Begin MSMask.MaskEdBox meTglLahir 
         Height          =   390
         Left            =   8880
         TabIndex        =   3
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Lahir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8880
         TabIndex        =   35
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7440
         TabIndex        =   24
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2160
         TabIndex        =   17
         Top             =   360
         Width           =   1350
      End
   End
   Begin VB.TextBox txtNoReservasi 
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
      Height          =   360
      Left            =   0
      MaxLength       =   10
      TabIndex        =   25
      Top             =   1680
      Width           =   1695
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   29
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Keterangan"
      Height          =   210
      Left            =   120
      TabIndex        =   50
      Top             =   4200
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "No Antrian"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   10320
      TabIndex        =   41
      Top             =   720
      Width           =   1830
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   9720
      Picture         =   "frmReservasi.frx":0CEE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmReservasi.frx":1A76
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10335
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmReservasi.frx":30D4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "No. Pendaftaran"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   0
      TabIndex        =   26
      Top             =   1440
      Width           =   1605
   End
End
Attribute VB_Name = "frmReservasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFilter As String
Dim strFilterVerifikasi As String
Dim intRowNow As Integer
Dim strSubInstalasi As String
Dim strNoAntrian As String
Dim dTglberlaku As Date
Dim curTarif As Currency
Dim curTP As Currency
Dim curTRS As Currency
Dim curPemb As Currency
Dim Qstrsql As String
Dim strKdDokter As String

Private Sub subLoadData()
    sRuangPeriksa = dcRuangan.Text
    sNamaPasien = txtNamaPasien.Text
    sJK = cbJenisKelamin.Text
    sUmur = txtTahun.Text & " th " & txtBulan.Text & " bl " & txtHari.Text & " hr"

End Sub


Private Sub chkNoCM_Click()
    If chkNoCM.value = Checked Then
        txtNoCM.Enabled = True
        txtNoCM.SetFocus
    Else
        txtNoCM.Enabled = False
        txtNoCM.Text = ""
        txtNamaPasien.Text = ""
        txtNamaPasien.Enabled = True
        txtNamaPasien.SetFocus
        cbJenisKelamin.Enabled = True
       
       
        meTglLahir.Enabled = True
         'meTglLahir.Text = ""
        meTglLahir.Text = "__/__/____"
        
        txtTahun.Enabled = True
        txtBulan.Enabled = True
        txtHari.Enabled = True
        txtTahun.Text = ""
        txtBulan.Text = ""
        txtHari.Text = ""
        
        cmdSimpan.Enabled = True
        
    End If
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo errLoad
blnSibuk = True
'cmdRujukan.Enabled = False
    If funcCekValidasi = False Then Exit Sub

    'cek pasien reservasi by 3whall 17/05/2012
     strSQL = "SELECT NamaLengkap " & _
              " FROM v_DaftarReservasiPasien " & _
              " WHERE (NamaLengkap = '" & txtNamaPasien.Text & "%') AND (DAY([Tgl Pesan]) = '" & Day(dtpTglReservasi.value) & "') AND (MONTH([Tgl Pesan]) = '" & Month(dtpTglReservasi.value) & "') AND (YEAR([Tgl Pesan]) = '" & Year(dtpTglReservasi.value) & "') " & _
              " And KdRuangan = '" & dcRuangan.BoundText & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        If MsgBox("Pasien sudah pesan pendaftaran.., " & vbNewLine & "Lanjutkan", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    End If

    cmdSimpan.Enabled = False

    'simpan data registrasi
    Call sp_ReservasiPasien(dbcmd)

    cmdSimpan.Enabled = True

    Call subEnableButtonReg(True)
    
'========================================== Untuk Menampilkan No Antrian (Cyber 28 Juni 2012) ===================================
    
    strSQL = "SELECT NoAntrian " & _
              " FROM ReservasiPasien Where NoReservasi = '" & txtNoReservasi.Text & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
    txtNoAntrian.Text = ""
    Exit Sub
    End If
    txtNoAntrian.Text = rs(0).value
'========================================== Untuk Menampilkan No Antrian (Cyber 28 Juni 2012) ===================================


'********************************************************* Cek Validas Pasien Reservasi (Cyber 25 Okt 12) ********************************************************
'    strSQL = "SELECT NoCM, Ruangan, [Dokter Pemeriksa], TglMasuk FROM ReservasiPasien WHERE (NoCM = '" & TxtNoCM.Text & "') AND kdruangan = '" & DcRuangan.BoundText & "' AND [Dokter Pemeriksa] = '" & txtDokter.Text & "' AND StatusPeriksa = 'T' and TglMasuk BETWEEN '" & Format(dtpTglReservasi, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpTglReservasi, "yyyy/MM/dd 23:59:59") & "' "
'    Call msubRecFO(rs, strSQL)
'    If Not rs.EOF Then
'        If MsgBox("Pasien tersebut sudah terdaftar di Reservasi," & vbNewLine & "Ruangan " & rs("Ruangan") & " , Dengan Dokter : " & rs("Dokter Pemeriksa") & " ", vbCritical, "Perhatian") Then
'            mstrNoCM = ""
'            TxtNoCM = ""
'            subClearData
'            fraDokter.Visible = False
'            chkDetailPasien.Enabled = False
'            TxtNoCM.SetFocus
'        Exit Sub
'        Else
'            TxtNoCM = ""
'        Exit Sub
'        End If
'    End If
'*********************************************************(Cyber 25 Okt 12) ********************************************************

'------------------------------------------ Dayz --------------------------------------------
'    dcRuangan.Enabled = True
'    dcInstalasi.Enabled = True
'    dcNamaPoli.Enabled = True
'    txtKet.Enabled = True
'    dtpTglReservasi.Enabled = True
'    txtDokter.Enabled = True
'    txtTlp.Enabled = True
'    dcNamaPoli.Enabled = True
'
'    chkNoCM.value = vbUnchecked
'    mstrNoCM = ""
'    TxtNoCM.Text = ""
'    txtNamaPasien = ""
'    cbJenisKelamin.Text = ""
'    meTglLahir.Text = "__/__/____"
'    txtTahun.Text = ""
'    txtBulan.Text = ""
'    txtHari.Text = ""
'    txtDokter.Text = ""
'    dcRuangan.Text = ""
'    dcInstalasi.Text = ""
'    dcNamaPoli.Text = ""
'    txtKet.Text = ""
'    txtTlp.Text = ""
''    TxtNoAntrian.Text = ""
'    Call Form_Load
'    txtNamaPasien.SetFocus
'    fraDokter.Visible = False
'    CmdSimpan.Enabled = True
'---------------------------------------- 16/12/2012 ------------------------------------------------
Exit Sub
errLoad:
    Call msubPesanError
    cmdSimpan.Enabled = True
    blnSibuk = False
End Sub

Private Sub cmdTutup_Click()
    If cmdSimpan.Enabled = True And txtNamaPasien.Text <> "" Then
        If MsgBox("Simpan Data ", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    
    Call subClearData

    Unload Me
End Sub

'Private Sub Command1_Click()
'If cmdSimpan.Enabled = True Then Exit Sub
'    mstrNoCM = Trim(txtNoCM)
'    frmCetakCatatanMedis.Show
'End Sub

'*************************************** Cyber 13 April 2012 **************************
Private Sub dcInstalasi_Change()

    Call SubTampilReservasiRI
        
End Sub
'================================ Add By Dayz =====================================
Private Sub SubTampilReservasiRI()
    If dcInstalasi.BoundText = "03" Then
        fraRI.Visible = True
        Frame2.Top = 4920
        frmReservasi.Height = 6210
        Label4.Top = 4320
        txtKet.Top = 4560
        Call centerForm(Me, MDIUtama)
        
'        Call Animate(frmRegistrasiAll, 8280, True)
    Else
        fraRI.Visible = False
        Frame2.Top = 4080
        frmReservasi.Height = 5280
        Label4.Top = 3360
        txtKet.Top = 3600
        Call centerForm(Me, MDIUtama)
        
'        Call Animate(frmRegistrasiAll, 5385, False)
    End If
End Sub
'*************************************** EDITED By DAYZ ***********************************************
Private Sub dcInstalasi_GotFocus()
On Error GoTo errLoad
Dim tempKode As String

    tempKode = dcInstalasi.BoundText
'    strSQL = "SELECT DISTINCT KdInstalasi,NamaInstalasi FROM V_KelasPelayanan where KdInstalasi in ('02','03','32','04') and StatusEnabled='1'"
    strSQL = "SELECT DISTINCT KdInstalasi,NamaInstalasi FROM V_KelasPelayananinstalasireservasi"
    Call msubDcSource(dcInstalasi, rs, strSQL)
'    strSQL = "SELECT DISTINCT KdInstalasi,NamaInstalasi FROM V_KelasPelayanan where KdInstalasi in ('02','03','32') and StatusEnabled='1'"
'    Call msubDcSource(dcInstalasi, rs, strSQL)
    
    dcInstalasi.BoundText = tempKode
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcInstalasi.MatchedWithList = True Then dcRuangan.SetFocus
'        strSQL = "SELECT DISTINCT KdInstalasi,NamaInstalasi FROM V_KelasPelayanan where KdInstalasi in ('02','03','32','04') and StatusEnabled='1' and (NamaInstalasi LIKE '%" & dcInstalasi.Text & "%')"
        strSQL = "SELECT DISTINCT KdInstalasi,NamaInstalasi FROM V_KelasPelayananinstalasireservasi where(NamaInstalasi LIKE '%" & dcInstalasi.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcInstalasi.Text = ""
        Exit Sub
        End If
        dcInstalasi.BoundText = rs(0).value
        dcInstalasi.Text = rs(1).value
        dcRuangan.SetFocus
    End If
End Sub
'





Private Sub dcKelasKamarRI_GotFocus()
On Error GoTo errLoad
Dim tempKdKelas As String
Dim tempKdRuangan As String

    tempKdKelas = dcKelasKamarRI.BoundText
    
    'cek kelas intensif
    strSQL = "SELECT Distinct KdKelas, KdRuangan From V_KamarRegRawatInap WHERE KdRuangan = '" & dcRuangan.BoundText & "' and StatusEnabled='1' and Expr1='1'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF Then Exit Sub
    tempKdRuangan = rs("KdRuangan").value
    
    If rs(0).value = "04" Then
        strSQL = "SELECT DISTINCT KdKelas, Kelas " & _
            " FROM V_KamarRegRawatInap " & _
            " WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND  StatusEnabled='1'"
    Else
        strSQL = "SELECT DISTINCT KdKelas, Kelas " & _
            " FROM V_KamarRegRawatInap " & _
            " WHERE KdRuangan = '" & dcRuangan.BoundText & "'  and StatusEnabled='1'"
    End If
    
    Call msubDcSource(dcKelasKamarRI, rs, strSQL)
    dcKelasKamarRI.BoundText = tempKdKelas

Exit Sub
errLoad:
    Call msubPesanError
End Sub
Private Sub dcKelasKamarRI_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If dcKelasKamarRI.MatchedWithList = True Then dcNoKamarRI.SetFocus
'        strSQL = "SELECT DISTINCT KdKelas, Kelas from V_KamarRegRawatInap WHERE kelas LIKE '%" & dcKelasKamarRI.Text & "%'"
        strSQL = "SELECT Kdkelas, kelas FROM V_KelasPelayanan WHERE (KdInstalasi IN ('" & dcInstalasi.BoundText & "','08')) AND (kelas LIKE '%" & dcKelasKamarRI.Text & "%') and Expr3='1'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcKelasKamarRI.Text = ""
        dcNoKamarRI.SetFocus
        Exit Sub
        End If
        dcKelasKamarRI.BoundText = rs(0).value
        dcKelasKamarRI.Text = rs(1).value
    End If

End Sub

Private Sub dcNoBedRI_GotFocus()
On Error GoTo errLoad
Dim tempKode As String

    tempKode = dcNoBedRI.BoundText
    strSQL = "SELECT distinct dbo.StatusBed.NoBed, dbo.StatusBed.NoBed AS Alias, dbo.StatusBed.StatusEnabled, dbo.NoKamar.StatusEnabled as Expr" & _
        " FROM dbo.NoKamar INNER JOIN dbo.StatusBed ON dbo.NoKamar.KdKamar = dbo.StatusBed.KdKamar" & _
        " WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND (KdKelas = '" & dcKelasKamarRI.BoundText & "') AND (dbo.StatusBed.StatusBed = 'K') AND (dbo.NoKamar.KdKamar = '" & dcNoKamarRI.BoundText & "')"
    Call msubDcSource(dcNoBedRI, rs, strSQL)
    dcNoBedRI.BoundText = tempKode
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcNoBedRI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKet.SetFocus
End Sub

Private Sub dcNoKamarRI_GotFocus()
On Error GoTo errLoad
Dim tempKode As String
    
    tempKode = dcNoKamarRI.BoundText
'    strSQL = "SELECT dbo.NoKamar.NoKamar, dbo.NoKamar.NamaKamar AS Alias" & _
'        " FROM dbo.NoKamar INNER JOIN dbo.StatusBed ON dbo.NoKamar.NoKamar = dbo.StatusBed.NoKamar" & _
'        " WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND (KdKelas = '" & dcKelasKamarRI.BoundText & "') AND (dbo.StatusBed.StatusBed = 'K')"
        strSQL = "SELECT distinct dbo.NoKamar.KdKamar,dbo.NoKamar.NamaKamar AS Alias, dbo.NoKamar.StatusEnabled, dbo.StatusBed.StatusEnabled " & _
            " FROM dbo.NoKamar INNER JOIN dbo.StatusBed ON dbo.NoKamar.KdKamar = dbo.StatusBed.KdKamar " & _
            " WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND (KdKelas = '" & dcKelasKamarRI.BoundText & "') AND (dbo.StatusBed.StatusBed = 'K') and dbo.NoKamar.StatusEnabled='1' and dbo.StatusBed.StatusEnabled='1' "

    
    Call msubDcSource(dcNoKamarRI, rs, strSQL)
    dcNoKamarRI.BoundText = tempKode
    
Exit Sub
errLoad:
    Call msubPesanError

End Sub

Private Sub dcNoKamarRI_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 13 Then
        If dcNoKamarRI.MatchedWithList = True Then dcNoBedRI.SetFocus
        strSQL = "SELECT KdKamar,NamaKamar " & _
            " FROM dbo.NoKamar " & _
            " WHERE (NamaKamar = '" & dcNoKamarRI.Text & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcNoKamarRI.Text = ""
        dcNoBedRI.SetFocus
        Exit Sub
        End If
        dcNoKamarRI.BoundText = rs(0).value
        dcNoKamarRI.Text = rs(1).value
    End If
Exit Sub
hell:
    msubPesanError

End Sub

Private Sub dcRuangan_GotFocus()
On Error GoTo errLoad
Dim tempKode As String

    tempKode = dcRuangan.BoundText
    If dcInstalasi.BoundText = "03" Then
        strSQL = "SELECT distinct KdRuangan, NamaRuangan FROM V_KelasPelayanan WHERE (KdInstalasi = '" & dcInstalasi.BoundText & "') and Expr3='1' ORDER BY NamaRuangan"
    Else
        strSQL = "SELECT distinct KdRuangan, NamaRuangan FROM V_KelasPelayanan WHERE (KdInstalasi = '" & dcInstalasi.BoundText & "') and Expr3='1' ORDER BY NamaRuangan"
    End If
    Call msubDcSource(dcRuangan, rs, strSQL)
    dcRuangan.BoundText = tempKode

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcRuangan_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 13 Then
        strSQL = "SELECT KdRuangan, NamaRuangan FROM V_KelasPelayanan WHERE (KdInstalasi IN ('" & dcInstalasi.BoundText & "','08')) AND (NamaRuangan LIKE '%" & dcRuangan.Text & "') and Expr3='1'"
'        strSQL = "SELECT KdRuangan, NamaRuangan FROM V_KelasPelayanan WHERE (KdInstalasi IN ('" & dcInstalasi.BoundText & "','08')) and Expr3='1'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcRuangan.Text = ""
        Exit Sub
        End If
        dcRuangan.BoundText = rs(0).value
        dcRuangan.Text = rs(1).value
      
        dcDokter.SetFocus

    End If
Exit Sub
errLoad:
End Sub

Private Sub dtpTglReservasi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcInstalasi.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strCtrlKey As String
On Error Resume Next

    'deklarasi tombol control ditekan
    strCtrlKey = (Shift + vbCtrlMask)

Select Case KeyCode
        Case vbKeyF1
        'Jarakal
           If cmdSimpan.Enabled = True Then Exit Sub
            mstrNoPen = txtNoReservasi.Text
            mstrKdInstalasi = dcInstalasi.BoundText
            frm_cetak_label_viewer.Show
'            frm_cetak_label_viewer.Cetaklangsung
        Case vbKeyR
            If strCtrlKey = 4 Then
                Unload Me
                frmReservasi.Show
            End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglReservasi.value = Now
    
    
    ' untuk mendapatkan jumlah panjang NoCM pada setting global
    strSQL = "Select value from SettingGlobal where Prefix = 'LenNoCM'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        strBanyakNoCM = rs(0).value
'    Else
'        strBanyakNoCM = "6"
    End If
    
    txtNoCM.MaxLength = strBanyakNoCM
    
    strRegistrasi = "RJ"
    If mblnCariPasien = True Then frmCariPasien.Enabled = False
    If bolAntrian = True Then
        txtKdAntrian.Enabled = True
    Else
        txtKdAntrian.Enabled = False
    End If
    
    strSQL = "SELECT DISTINCT KdInstalasi,NamaInstalasi FROM V_KelasPelayanan where KdInstalasi in ('02','03','32','04') and StatusEnabled='1'"
    Call msubDcSource(dcInstalasi, rs, strSQL)
    Set rs = Nothing
    
    dcInstalasi.BoundText = rs.Fields(0).value
    dcInstalasi.BoundText = rs.Fields(1).value
    
    dcInstalasi.BoundText = "02"
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnCariPasien = True Then frmCariPasien.Enabled = True
        Call subClearData
End Sub

Private Sub txtKdAntrian_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNoCM.SetFocus
End Sub

Private Sub hgPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub
Private Sub txtKdAntrian_LostFocus()
On Error Resume Next
    If Update_AntrianPasienRegistrasi(txtKdAntrian.Text, 0, 0, 0, 0, 0, "PROSES") = False Then Exit Sub
Exit Sub
'hell_:
'    msubPesanError
End Sub

Private Sub txtKet_GotFocus()
'Me.Height = 5280
'fraDokter.Visible = False
End Sub

Private Sub txtKet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdSimpan.Enabled = False Then
            cmdTutup.SetFocus
        Else
            cmdSimpan.SetFocus
        End If
    End If
End Sub

Private Sub txtNoCM_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If txtNoCM <> "" Then
            blnSibuk = True
            Call CariData
            dcInstalasi.BoundText = "02"
           'meTglLahir.SetFocus
'           dtpTglReservasi.SetFocus
        Else
'            txtNamaPasien.Enabled = True
            cbJenisKelamin.Enabled = True
            txtTahun.Enabled = True
            txtBulan.Enabled = True
            txtHari.Enabled = True
            txtNamaPasien.SetFocus
        End If
    
    End If
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = 13 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    If KeyAscii = Asc(",") Then Exit Sub
    If KeyAscii = Asc(".") Then Exit Sub
    
    
End Sub

Private Sub txtNoCM_LostFocus()
    Call CariData
End Sub

'untuk enable/disable button reg
Private Sub subEnableButtonReg(blnStatus As Boolean)

    cmdSimpan.Enabled = Not blnStatus
'    dtpTglReservasi.Enabled = Not blnStatus
    dcInstalasi.Enabled = Not blnStatus
    dcRuangan.Enabled = Not blnStatus

End Sub

'Store procedure untuk mengisi reservasi pasien
Private Sub sp_ReservasiPasien(ByVal adoCommand As ADODB.Command)
 Set adoCommand = New ADODB.Command
    
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoReservasi", adInteger, adParamInput, , Null)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("NamaLengkap", adVarChar, adParamInput, 50, txtNamaPasien.Text)
        If cbJenisKelamin.Text = "Laki-laki" Then
            .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 2, "01")
        Else
            .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 2, "02")
        End If
        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(meTglLahir, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglPemesanan", adDate, adParamInput, , Format(Now, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, dcRuangan.BoundText)
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, strKdDokterReservasi)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dtpTglReservasi.value, "yyyy/MM/dd HH:mm:ss"))
'        .Parameters.Append .CreateParameter("RuanganPoli", adVarChar, adParamInput, 20, IIf((dcNamaPoli.Text = ""), dcNamaPoli.Text, dcNamaPoli.Text))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 150, IIf(txtKet.Text = "", Null, txtKet.Text))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("NoTlp", adVarChar, adParamInput, 15, IIf(txtTlp.Text = "", Null, txtTlp.Text))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")
        .Parameters.Append .CreateParameter("OutputNoReservasi", adInteger, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("OutputNoAntrian", adVarChar, adParamOutput, 3, Null)
        .Parameters.Append .CreateParameter("StatusDaftar", adVarChar, adParamInput, 1, "T")
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, dcKelasKamarRI.BoundText)
        .Parameters.Append .CreateParameter("KdKamar", adChar, adParamInput, 4, dcNoKamarRI.BoundText)
        .Parameters.Append .CreateParameter("NoBed", adChar, adParamInput, 2, dcNoBedRI.BoundText)
        .Parameters.Append .CreateParameter("StatusReservasi", adChar, adParamInput, 1, "Y")

        
'        .Parameters.Append .CreateParameter("KdJenisPoliklinik", adChar, adParamInput, 2, IIf((dcNamaPoli.BoundText = ""), dcNamaPoli.BoundText, dcNamaPoli.BoundText))

        .ActiveConnection = dbConn
        .CommandText = "AUD_ReservasiPasien"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data", vbCritical, "Validasi"
        Else
            txtNoReservasi.Text = .Parameters("OutputNoReservasi").value
'            TxtNoAntrian.Text = .Parameters("OutputNoAntrian").value
            MsgBox "Pasien berhasil direservasi", vbInformation
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
'    MousePointer = vbDefault
    Exit Sub
End Sub

'untuk cek validasi
Private Function funcCekValidasi() As Boolean
'Validasi Untuk digit No CM & Telepon (Cyber 27 Juni 2012)
    If chkNoCM.value = 1 And Len(txtNoCM.Text) < 6 Then
        MsgBox "No CM harus 6 Digit", vbExclamation, "Validasi"
        funcCekValidasi = False
        txtNoCM.SetFocus
        Exit Function
    End If

'   splakuk 22-02-2012

'        If txtKdDokter.Text = "" Then
'            MsgBox "Pilihan Dokter harus diisi sesuai data daftar dokter", vbExclamation, "Validasi"
'            funcCekValidasi = False
'            txtDokter.SetFocus
'            Exit Function
'        End If
'
    If Periksa("datacombo", dcInstalasi, "Nama instalasi kosong") = False Then funcCekValidasi = False: Exit Function

    If Periksa("datacombo", dcRuangan, "Nama ruangan kosong") = False Then funcCekValidasi = False: Exit Function
    If Periksa("text", txtNamaPasien, "Nama Pasien kosong") = False Then funcCekValidasi = False: Exit Function
    
    If meTglLahir.Text = "__/__/____" Then
        MsgBox "Umur Tidak Boleh Kosong", vbInformation
        meTglLahir.SetFocus
        Exit Function
    End If
    

'    triwall 17052012
    If txtTlp.Text = "" Then
        MsgBox "No. Telepon wajib diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        txtTlp.SetFocus
        Exit Function
    End If
    
    
    funcCekValidasi = True
End Function

'untuk membersihkan data pasien registrasi
Private Sub subClearData()
    txtNoPakai.Text = ""
    txtNoReservasi.Text = ""
    txtNamaPasien.Text = ""
    cbJenisKelamin.Text = ""
    txtTahun.Text = ""
    txtBulan.Text = ""
    txtHari.Text = ""
'    dcHubungan.BoundText = ""
    dtpTglReservasi.MaxDate = #9/9/2999#
    dtpTglReservasi.value = Now
    dcInstalasi.Text = ""
    dcRuangan.Text = ""
'    dcJenisKelas.Text = ""
'    dcKelompokPasien.Text = ""
'    dcKelas.Text = ""
'    txtDokter.Text = ""
    txtKet.Text = ""
    txtTlp.Text = ""
    strReservasi = ""
    fraDokter.Visible = False
    dcDokter.Text = ""
End Sub
Private Sub txtNamaPasien_KeyDown(KeyAscii As Integer, Shift As Integer)

    If KeyAscii = 13 Then
        cbJenisKelamin.SetFocus
'        TxtNoAntrian.Text = ""
    End If
End Sub

Private Sub cbJenisKelamin_keydown(KeyAscii As Integer, Shift As Integer)
    If KeyAscii = 13 Then
        meTglLahir.SetFocus
    End If
End Sub

Private Sub txtTlp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If dcInstalasi.BoundText <> "03" Then
            txtKet.SetFocus
        Else
            dcKelasKamarRI.SetFocus
        End If
    End If
'    Call SetKeyPressToNumber(KeyCode)
End Sub

'untuk mengganti nocm on change
Public Sub CariData()
On Error GoTo errLoad
    Call subClearData
    Call subEnableButtonReg(False)
    
    dcInstalasi.BoundText = "02"
    'cek pasien Meninggal
    strSQL = "SELECT NoCM FROM PasienMeninggal WHERE (NoCM = '" & txtNoCM.Text & "')"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        MsgBox "Pasien tersebut Sudah meninggal", vbInformation, "Informasi"
        mstrNoCM = ""
'        chkDetailPasien.Enabled = False
        cmdSimpan.Enabled = False
        txtNoCM.Text = ""
        txtNoCM.SetFocus
        Exit Sub
    End If
    
    'cek pasien igd
    strSQL = "SELECT NoCM FROM V_DaftarPasienIGDAktif WHERE (NoCM = '" & txtNoCM.Text & "')"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        MsgBox "Pasien tersebut belum keluar dari IGD", vbInformation, "Informasi"
        mstrNoCM = ""
'        chkDetailPasien.Enabled = False
        cmdSimpan.Enabled = False
        Exit Sub
    End If
    
    'cek pasien ri
    strSQL = "SELECT dbo.RegistrasiRI.NoCM, dbo.Ruangan.NamaRuangan FROM dbo.RegistrasiRI INNER JOIN dbo.Ruangan ON dbo.RegistrasiRI.KdRuangan = dbo.Ruangan.KdRuangan WHERE (NoCM = '" & txtNoCM.Text & "') AND StatusPulang = 'T'"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        MsgBox "Pasien tersebut belum keluar dari Rawat Inap," & vbNewLine & "Ruangan " & rs("NamaRuangan") & " ", vbInformation, "Informasi"
        mstrNoCM = ""
'        chkDetailPasien.Enabled = False
        cmdSimpan.Enabled = False
        Exit Sub
    End If
    
    strSQL = "Select * from v_CariPasien WHERE [No. CM]='" & txtNoCM.Text & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        mstrNoCM = ""
'        chkDetailPasien.Enabled = False
        cmdSimpan.Enabled = False
        Exit Sub
    End If
    
    mstrNoCM = txtNoCM.Text
    txtNamaPasien.Text = rs.Fields("Nama Lengkap").value
    txtNamaPasien.Enabled = False
    cbJenisKelamin.Enabled = False
    If rs.Fields("JK").value = "P" Then
        cbJenisKelamin.Text = "Perempuan"
    ElseIf rs.Fields("JK").value = "L" Then
        cbJenisKelamin.Text = "Laki-laki"
    End If
    meTglLahir.Text = rs.Fields("tgllahir").value
    meTglLahir.Enabled = False
    txtTahun.Text = rs.Fields("UmurTahun").value
    txtBulan.Text = rs.Fields("UmurBulan").value
    txtHari.Text = rs.Fields("UmurHari").value
    txtTahun.Enabled = False
    txtBulan.Enabled = False
    txtHari.Enabled = False
    Set rs = Nothing
'    chkDetailPasien.Enabled = True
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subCetakLabelRegistrasi()
On Error GoTo errLoad
    Printer.Print strNNamaRS
    Printer.Print strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
    Printer.Print strWebsite & ", " & strEmail
    
    If (mstrKdInstalasi = "02") Or (mstrKdInstalasi = "11") Or (mstrKdInstalasi = "06") Then
        strSQL = "SELECT * from V_CetakLabelRegistrasiPasienMRS WHERE (NoPendaftaran) =('" & mstrNoPen & "')"
    Else
        strSQL = "SELECT * from V_CetakLabelRegistrasiPasienMRS WHERE (NoPendaftaran) =('" & mstrNoPen & "')"
    End If
    Call msubRecFO(rs, strSQL)
    
    Printer.Print "No. Pendaftaran"
    Printer.Print "No. CM"
    Printer.Print "Nama Pasien"
    Printer.Print "Jenis Kelamin"
    Printer.Print "Kelompok Pasien"
    Printer.Print "Jenis Kelas"
    Printer.Print "Ruangan Tujuan"
    Printer.Print "Lokasi Ruangan"
    Printer.Print "No. Ruangan"

    Printer.Print "No. Antrian"
    Printer.Print "------------------------------"

    strSQL = "SELECT MessageToDay FROM MasterDataPendukung"
    Call msubRecFO(rs, strSQL)
    Printer.Print IIf(IsNull(rs(0)), "", rs(0))
    Printer.Print "------------------------------"
    Printer.Print "User :"
    
    Printer.EndDoc
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcDokter_GotFocus()
On Error GoTo gabril

    strhari = Format(dtpTglReservasi, "DDDD")
    'strhari = txtHari.Text
        strSQL = "SELECT KdDokter,NamaLengkap FROM V_JadwalPraktekDokter  where NamaRuangan='" & dcRuangan.Text & "' " & _
        "and hari ='" & strhari & "'"
        Call msubRecFO(rsCek, strSQL)
        
    If Not rsCek.EOF Then
        
        strSQL = "SELECT KdDokter,NamaLengkap FROM V_JadwalPraktekDokter  where NamaRuangan='" & dcRuangan.Text & "' " & _
                 "and hari ='" & strhari & "'"
    Else
        strSQL = "SELECT KodeDokter,NamaDokter FROM V_DaftarDokter "
    End If
    
        Call msubDcSource(dcDokter, rs, strSQL)
        dcDokter.BoundText = rs.Fields(0).value
        dcDokter.Text = rs.Fields(1).value
        
        strKdDokterReservasi = rs.Fields(0).value
    
Exit Sub
gabril:
    Call msubPesanError

End Sub
Private Sub dcDokter_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then
'        If intJmlDokter = 0 Then Exit Sub
        txtTlp.SetFocus
    End If
Exit Sub
errLoad:
End Sub
Private Sub meTglLahir_KeyPress(KeyAscii As Integer)
    On Error GoTo errTglLahir
    If KeyAscii = 13 Then
        If meTglLahir.Text = "__/__/____" Then
            txtTahun.SetFocus
            Exit Sub
        End If
        If funcCekValidasiTgl("TglLahir", meTglLahir) = "NoErr" Then
            Call subYearOldCount(Format(meTglLahir.Text, "yyyy/mm/dd"))
            txtTahun.Text = YOC_intYear
            txtBulan.Text = YOC_intMonth
            txtHari.Text = YOC_intDay
        End If
        dtpTglReservasi.SetFocus
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
        txtHari.Text = YOC_intDay
    Else
        txtTahun.Text = ""
        txtBulan.Text = ""
        txtHari.Text = ""
    End If
    
    
Exit Sub
errTglLahir:
    MsgBox "Sudahkah anda mengganti Regional Setting komputer anda menjadi 'Indonesia'?" _
        & vbNewLine & "Bila sudah hubungi Administrator dan laporkan pesan kesalahan berikut:" _
        & vbNewLine & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
End Sub


Private Sub txtBulan_Change()
    Dim dTglLahir As Date
    
  If chkNoCM.value = 0 Then
    
    If txtBulan.Text = "" And txtTahun.Text = "" Then txtHari.SetFocus: Exit Sub
    If txtBulan.Text = "" Then txtBulan.Text = 0
    If txtTahun.Text = "" And txtHari.Text = "" Then
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
    ElseIf txtTahun.Text <> "" And txtHari.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    ElseIf txtTahun.Text = "" And txtHari.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
    ElseIf txtTahun.Text <> "" And txtHari.Text = "" Then
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    End If
 
 End If
    
    
'    meTglLahir.Text = dTglLahir
End Sub

Private Sub txtBulan_KeyPress(KeyAscii As Integer)
    Dim dTglLahir As Date
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If txtBulan.Text = "" And txtTahun.Text = "" Then txtHari.SetFocus: Exit Sub
        If txtBulan.Text = "" Then txtBulan.Text = 0
        If txtTahun.Text = "" And txtHari.Text = "" Then
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
        ElseIf txtTahun.Text <> "" And txtHari.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        ElseIf txtTahun.Text = "" And txtHari.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        ElseIf txtTahun.Text <> "" And txtHari.Text = "" Then
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        End If
        meTglLahir.Text = dTglLahir
        txtHari.SetFocus
    End If
End Sub

Private Sub txtHari_Change()
    Dim dTglLahir As Date
    If txtHari.Text = "" And txtBulan.Text = "" And txtTahun.Text = "" Then dtpTglReservasi.SetFocus: Exit Sub
    If txtHari.Text = "" Then txtHari.Text = 0
    If txtTahun.Text = "" And txtBulan.Text = "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
    ElseIf txtTahun.Text <> "" And txtBulan.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    ElseIf txtTahun.Text = "" And txtBulan.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
    ElseIf txtTahun.Text <> "" And txtBulan.Text = "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    End If
'    meTglLahir.Text = dTglLahir
End Sub

Private Sub txtHari_KeyPress(KeyAscii As Integer)
    Dim dTglLahir As Date
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If txtHari.Text = "" And txtBulan.Text = "" And txtTahun.Text = "" Then dtpTglReservasi.SetFocus: Exit Sub
        If txtHari.Text = "" Then txtHari.Text = 0
        If txtTahun.Text = "" And txtBulan.Text = "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        ElseIf txtTahun.Text <> "" And txtBulan.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        ElseIf txtTahun.Text = "" And txtBulan.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        ElseIf txtTahun.Text <> "" And txtBulan.Text = "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        End If

        meTglLahir.Text = dTglLahir
        dtpTglReservasi.SetFocus
    End If
End Sub

Private Sub txtTahun_Change()

 Dim dTglLahir As Date
    
 If chkNoCM.value = 0 Then
 
    If txtTahun = "" Then txtBulan.SetFocus: Exit Sub
    If txtBulan.Text = "" And txtHari.Text = "" Then
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), Date)
    ElseIf txtBulan.Text <> "" And txtHari.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    ElseIf txtBulan.Text = "" And txtHari.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    ElseIf txtBulan.Text <> "" And txtHari.Text = "" Then
        dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
    End If
 End If

'    dtpTglReservasi.SetFocus
'    meTglLahir.Text = dTglLahir
End Sub

Private Sub txtTahun_KeyPress(KeyAscii As Integer)
    Dim dTglLahir As Date
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If txtTahun = "" Then txtBulan.SetFocus: Exit Sub
        If txtBulan.Text = "" And txtHari.Text = "" Then
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), Date)
        ElseIf txtBulan.Text <> "" And txtHari.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), dTglLahir)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        ElseIf txtBulan.Text = "" And txtHari.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHari.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        ElseIf txtBulan.Text <> "" And txtHari.Text = "" Then
            dTglLahir = DateAdd("m", -1 * CInt(txtBulan.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahun.Text), dTglLahir)
        End If
        
        meTglLahir.Text = dTglLahir
        txtBulan.SetFocus
    End If
End Sub

Private Sub txtNamaPasien_LostFocus()
'Ganti Uppercase krn ketentuan RSM Cyber 14042012
    txtNamaPasien = StrConv(txtNamaPasien, vbUpperCase)
'    txtNamaPasien = StrConv(txtNamaPasien, vbProperCase)
End Sub

Private Sub txtTlp_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If dcInstalasi.BoundText <> "03" Then
'        txtKet.SetFocus
'    Else
'        dcKelasKamarRI.SetFocus
'    End If
    
    If KeyAscii = 13 Then txtTlp.SetFocus
    Call SetKeyPressToNumber(KeyAscii)

'End If
End Sub
