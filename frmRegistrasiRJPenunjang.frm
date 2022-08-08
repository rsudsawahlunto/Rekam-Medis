VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRegistrasiRJPenunjang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Registrasi Pasien RJ dan Penunjang"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegistrasiRJPenunjang.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   12165
   Begin VB.Frame fraKonsul 
      Caption         =   "Konsul Penunjang dan Antar Ruangan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   100
      Top             =   6480
      Visible         =   0   'False
      Width           =   12135
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   0
         TabIndex        =   103
         Top             =   240
         Width           =   12135
         Begin MSDataListLib.DataCombo dcDokterPerujuk 
            Height          =   330
            Left            =   6600
            TabIndex        =   131
            Top             =   1200
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcRuanganTM 
            Height          =   330
            Left            =   3240
            TabIndex        =   130
            Top             =   1200
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcInstalasiTM 
            Height          =   330
            Left            =   120
            TabIndex        =   129
            Top             =   1200
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker dtpTglOrderTM 
            Height          =   330
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy HH:mm"
            Format          =   157745155
            UpDown          =   -1  'True
            CurrentDate     =   38061
         End
         Begin VB.PictureBox dgKonsul 
            Appearance      =   0  'Flat
            Height          =   1455
            Left            =   120
            ScaleHeight     =   1425
            ScaleWidth      =   7905
            TabIndex        =   119
            Top             =   4320
            Width           =   7935
         End
         Begin VB.PictureBox dgHistoryPelayanan 
            Height          =   1455
            Left            =   8160
            ScaleHeight     =   1395
            ScaleWidth      =   8595
            TabIndex        =   118
            Top             =   4920
            Width           =   8655
         End
         Begin VB.CheckBox chkOperasiBersama 
            Caption         =   "Operasi Bersama"
            Height          =   255
            Left            =   9960
            TabIndex        =   117
            ToolTipText     =   "Operasi bersama dokter Operator 1 dan dokter Operator 2"
            Top             =   4200
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txtSex 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   8400
            TabIndex        =   116
            Top             =   480
            Width           =   1095
         End
         Begin VB.Frame Frame7 
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
            Left            =   9600
            TabIndex        =   109
            Top             =   240
            Width           =   2415
            Begin VB.TextBox txtthntm 
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
               TabIndex        =   112
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtblntm 
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
               TabIndex        =   111
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txthrtm 
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
               TabIndex        =   110
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "thn"
               Height          =   210
               Left            =   550
               TabIndex        =   115
               Top             =   277
               Width           =   285
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "bln"
               Height          =   210
               Left            =   1350
               TabIndex        =   114
               Top             =   277
               Width           =   240
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "hr"
               Height          =   210
               Left            =   2130
               TabIndex        =   113
               Top             =   270
               Width           =   165
            End
         End
         Begin VB.TextBox txtNamaPasienTM 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   108
            Top             =   480
            Width           =   2535
         End
         Begin VB.CommandButton cmdHapusKonsul 
            Caption         =   "Hapus Konsul"
            Height          =   375
            Left            =   15120
            TabIndex        =   107
            Top             =   720
            Width           =   1815
         End
         Begin VB.CommandButton cmdSimpanKonsul 
            Caption         =   "Simpan Konsul"
            Height          =   375
            Left            =   15120
            TabIndex        =   106
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtNoCMTM 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2400
            MaxLength       =   15
            TabIndex        =   105
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
            TabIndex        =   104
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Dokter Pemeriksa"
            Height          =   210
            Left            =   6600
            TabIndex        =   127
            Top             =   960
            Width           =   2145
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Kelamin"
            Height          =   210
            Left            =   8400
            TabIndex        =   126
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
            TabIndex        =   125
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
            TabIndex        =   124
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Ruangan Tujuan"
            Height          =   210
            Left            =   3240
            TabIndex        =   123
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Instalasi Tujuan"
            Height          =   210
            Left            =   120
            TabIndex        =   122
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Rekam Medis"
            Height          =   210
            Index           =   7
            Left            =   2400
            TabIndex        =   121
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
            TabIndex        =   120
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdSimpanTM 
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
         Left            =   8520
         TabIndex        =   102
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutupTM 
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
         Left            =   10320
         TabIndex        =   101
         Top             =   2280
         Width           =   1695
      End
   End
   Begin VB.TextBox txtFormPengirim 
      Height          =   375
      Left            =   720
      TabIndex        =   98
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
   End
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
      Left            =   1320
      TabIndex        =   85
      Top             =   6360
      Visible         =   0   'False
      Width           =   7815
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   1455
         Left            =   240
         TabIndex        =   86
         Top             =   360
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
   Begin VB.Frame Frame3 
      Caption         =   "Data Registrasi Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   38
      Top             =   2520
      Width           =   12135
      Begin MSDataListLib.DataCombo dcDokter 
         Height          =   360
         Left            =   3480
         TabIndex        =   97
         Top             =   1320
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
      Begin VB.Frame FraMaks 
         Caption         =   "Kunj. Maks"
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
         Left            =   9240
         TabIndex        =   93
         Top             =   1080
         Width           =   2775
         Begin VB.TextBox txtMaks 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   94
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblKonsul 
            Caption         =   "Kunj. ke: 1"
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
            Left            =   1200
            TabIndex        =   95
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CheckBox chkPaket 
         Caption         =   "Paket MCU / Rehab Medik"
         Enabled         =   0   'False
         Height          =   245
         Left            =   6600
         TabIndex        =   91
         Top             =   1080
         Width           =   2535
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   360
         Left            =   480
         TabIndex        =   9
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
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
         TabIndex        =   12
         Top             =   600
         Width           =   3855
         _ExtentX        =   6800
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
      Begin MSDataListLib.DataCombo dcKelas 
         Height          =   360
         Left            =   8520
         TabIndex        =   11
         Top             =   2280
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSComCtl2.DTPicker dtpTglPendaftaran 
         Height          =   360
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
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
         Format          =   157614083
         UpDown          =   -1  'True
         CurrentDate     =   38061
      End
      Begin MSDataListLib.DataCombo dcKelompokPasien 
         Height          =   360
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
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
      Begin MSDataListLib.DataCombo dcJenisKelas 
         Height          =   360
         Left            =   2280
         TabIndex        =   10
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
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
      Begin MSDataListLib.DataCombo dcSubInstalasi 
         Height          =   360
         Left            =   3120
         TabIndex        =   13
         Top             =   2160
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo dcRujukanRI 
         Height          =   360
         Left            =   8640
         TabIndex        =   14
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSDataListLib.DataCombo dcPaket 
         Height          =   360
         Left            =   6600
         TabIndex        =   92
         Top             =   1320
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Pemeriksa"
         Height          =   210
         Left            =   3480
         TabIndex        =   90
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Asal Rujukan"
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
         Index           =   0
         Left            =   8700
         TabIndex        =   64
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "SMF (Kasus Penyakit)"
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
         Index           =   2
         Left            =   3090
         TabIndex        =   58
         Top             =   1920
         Width           =   1845
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pendaftaran"
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
         Left            =   2280
         TabIndex        =   53
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Penjamin"
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
         Left            =   240
         TabIndex        =   47
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pendaftaran"
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
         Left            =   240
         TabIndex        =   46
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Pelayanan"
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
         Left            =   8520
         TabIndex        =   41
         Top             =   2040
         Width           =   1380
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan Pemeriksaan / Perawatan"
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
         Left            =   4680
         TabIndex        =   40
         Top             =   360
         Width           =   3135
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
         Left            =   525
         TabIndex        =   39
         Top             =   1920
         Width           =   1860
      End
   End
   Begin VB.TextBox txtHari 
      Height          =   375
      Left            =   6960
      TabIndex        =   88
      Text            =   "Text1"
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtKdDokter 
      Height          =   375
      Left            =   2760
      TabIndex        =   87
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtNoReservasi 
      Height          =   375
      Left            =   4920
      TabIndex        =   84
      Text            =   "Text1"
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame fraAntrian 
      Height          =   855
      Left            =   6720
      TabIndex        =   80
      Top             =   720
      Width           =   5415
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
         Height          =   690
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   81
         Top             =   120
         Width           =   1095
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
         Index           =   2
         Left            =   1080
         TabIndex        =   82
         Top             =   120
         Width           =   3105
      End
   End
   Begin VB.TextBox txtNoBKM 
      Height          =   375
      Left            =   2640
      TabIndex        =   78
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame5 
      Caption         =   "Data Penanggungjawab Pasien"
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
      Left            =   0
      TabIndex        =   66
      Top             =   6360
      Width           =   12615
      Begin VB.CheckBox chkDiriSendiri 
         Caption         =   "&Diri Sendiri"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   11040
         TabIndex        =   22
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtTlpRI 
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
         Height          =   390
         Left            =   10560
         MaxLength       =   50
         TabIndex        =   32
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox txtAlamatRI 
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
         Height          =   1110
         Left            =   8040
         MaxLength       =   50
         TabIndex        =   33
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox txtNamaRI 
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
         Height          =   390
         Left            =   120
         MaxLength       =   20
         TabIndex        =   23
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtKodePos 
         Alignment       =   2  'Center
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
         Height          =   390
         Left            =   9000
         MaxLength       =   5
         TabIndex        =   31
         Top             =   2040
         Width           =   1455
      End
      Begin MSMask.MaskEdBox meRTRWPJ 
         Height          =   390
         Left            =   8040
         TabIndex        =   30
         Top             =   2040
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
      Begin MSDataListLib.DataCombo dcKotaPJ 
         Height          =   360
         Left            =   3840
         TabIndex        =   27
         Top             =   1320
         Width           =   4095
         _ExtentX        =   7223
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
      Begin MSDataListLib.DataCombo dcKecamatanPJ 
         Height          =   360
         Left            =   120
         TabIndex        =   28
         Top             =   2040
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
      Begin MSDataListLib.DataCombo dcKelurahanPJ 
         Height          =   360
         Left            =   3840
         TabIndex        =   29
         Top             =   2040
         Width           =   4095
         _ExtentX        =   7223
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
      Begin MSDataListLib.DataCombo dcPropinsiPJ 
         Height          =   360
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         OLEDropMode     =   1
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
      Begin MSDataListLib.DataCombo dcHubungan 
         Height          =   360
         Left            =   2880
         TabIndex        =   24
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
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
      Begin MSDataListLib.DataCombo dcPekerjaanPJ 
         Height          =   360
         Left            =   5280
         TabIndex        =   25
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Pekerjaan"
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
         Left            =   5280
         TabIndex        =   77
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hubungan"
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
         Index           =   1
         Left            =   2880
         TabIndex        =   76
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Telepon"
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
         Index           =   9
         Left            =   10560
         TabIndex        =   75
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Alamat Lengkap"
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
         Index           =   8
         Left            =   8040
         TabIndex        =   74
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Nama Lengkap"
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
         Index           =   7
         Left            =   120
         TabIndex        =   73
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Kode Pos"
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
         Left            =   9000
         TabIndex        =   72
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "RT/RW"
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
         Left            =   8040
         TabIndex        =   71
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Kelurahan"
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
         Left            =   3840
         TabIndex        =   70
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Kecamatan"
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
         TabIndex        =   69
         Top             =   1800
         Width           =   945
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Propinsi"
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
         TabIndex        =   68
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Kota/Kabupaten"
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
         Left            =   3840
         TabIndex        =   67
         Top             =   1080
         Width           =   1350
      End
   End
   Begin VB.TextBox txtNoPakai 
      Height          =   495
      Left            =   480
      TabIndex        =   65
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraRawatGabung 
      Caption         =   "Rawat Gabung ?"
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
      Left            =   10800
      TabIndex        =   63
      Top             =   6240
      Width           =   1695
      Begin VB.OptionButton optYa 
         Caption         =   "Ya"
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
         Left            =   960
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optTidak 
         Caption         =   "Tidak"
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
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar stbInformasi 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   62
      Top             =   5730
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4321
            MinWidth        =   1411
            Text            =   "Cetak Label (F1)"
            TextSave        =   "Cetak Label (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4674
            MinWidth        =   1764
            Text            =   "Pasien Baru Ctrl+B"
            TextSave        =   "Pasien Baru Ctrl+B"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4321
            MinWidth        =   1411
            Text            =   "Cari Pasien (F3)"
            TextSave        =   "Cari Pasien (F3)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   3731
            Text            =   "Lembar Masuk Ctrl+R"
            TextSave        =   "Lembar Masuk Ctrl+R"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5450
            Text            =   "Surat Keterangan Ctrl+Z"
            TextSave        =   "Surat Keterangan Ctrl+Z"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "Cetak SJP (F9)"
            TextSave        =   "Cetak SJP (F9)"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Text            =   "C. Medis Ctrl+M"
            TextSave        =   "C. Medis Ctrl+M"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraRegistrasiRI 
      Caption         =   "Data Masuk Rawat Inap"
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
      TabIndex        =   56
      Top             =   5520
      Width           =   12615
      Begin MSDataListLib.DataCombo dcCaraMasukRI 
         Height          =   360
         Left            =   1920
         TabIndex        =   16
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
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
      Begin MSDataListLib.DataCombo dcKelasKamarRI 
         Height          =   360
         Left            =   4680
         TabIndex        =   17
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
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
         Left            =   7440
         TabIndex        =   18
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
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
         Left            =   9360
         TabIndex        =   19
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   9360
         TabIndex        =   61
         Top             =   240
         Width           =   660
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
         Left            =   7440
         TabIndex        =   60
         Top             =   240
         Width           =   900
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
         Left            =   4680
         TabIndex        =   59
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Cara Masuk"
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
         Index           =   1
         Left            =   1920
         TabIndex        =   57
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   0
      TabIndex        =   45
      Top             =   4320
      Width           =   12135
      Begin VB.TextBox txtNoRujukan 
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
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   133
         Top             =   480
         Width           =   3615
      End
      Begin VB.CheckBox chkKonsul 
         Caption         =   "Konsul"
         Height          =   495
         Left            =   120
         TabIndex        =   128
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdTindakan 
         Caption         =   "Rujukan && Pemeriksaan"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6720
         TabIndex        =   96
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdRujukan 
         Caption         =   "&Data Rujukan"
         Enabled         =   0   'False
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
         Left            =   4920
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
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
         Left            =   8520
         TabIndex        =   36
         Top             =   360
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
         Left            =   10320
         TabIndex        =   37
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdAsuransiP 
         Caption         =   "&Asuransi Pasien"
         Enabled         =   0   'False
         Height          =   495
         Left            =   8520
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "No. Rujukan"
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
         Left            =   1200
         TabIndex        =   134
         Top             =   240
         Width           =   1035
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
      TabIndex        =   42
      Top             =   1440
      Width           =   12135
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
         Left            =   120
         MaxLength       =   12
         TabIndex        =   132
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cboJK 
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
         ItemData        =   "frmRegistrasiRJPenunjang.frx":0CCA
         Left            =   7440
         List            =   "frmRegistrasiRJPenunjang.frx":0CD4
         TabIndex        =   99
         Top             =   600
         Width           =   1575
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
         Left            =   9120
         TabIndex        =   48
         Top             =   150
         Width           =   2895
         Begin VB.TextBox txtHr 
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
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   5
            Top             =   330
            Width           =   375
         End
         Begin VB.TextBox txtBln 
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
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   4
            Top             =   330
            Width           =   375
         End
         Begin VB.TextBox txtThn 
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
            MaxLength       =   6
            TabIndex        =   3
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
            Left            =   2400
            TabIndex        =   51
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
            Left            =   1560
            TabIndex        =   50
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
            Left            =   720
            TabIndex        =   49
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.TextBox txtJK 
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
         Left            =   7440
         MaxLength       =   9
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtNoCM 
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
         MaxLength       =   12
         TabIndex        =   0
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtNamaPasien 
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
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   4935
      End
      Begin VB.CheckBox chkDetailPasien 
         Caption         =   "Detail Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10440
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkPrintKartu 
         Caption         =   "Print Kartu /u Pasien Lama"
         Height          =   375
         Left            =   10440
         TabIndex        =   83
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
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
         TabIndex        =   52
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
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
         TabIndex        =   44
         Top             =   360
         Width           =   615
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
         Left            =   2400
         TabIndex        =   43
         Top             =   360
         Width           =   1710
      End
   End
   Begin VB.TextBox txtNoPendaftaran 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   54
      Top             =   960
      Visible         =   0   'False
      Width           =   3135
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   79
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
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "No Pendaftaran"
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
      Left            =   3480
      TabIndex        =   89
      Top             =   720
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   10200
      Picture         =   "frmRegistrasiRJPenunjang.frx":0CEE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRegistrasiRJPenunjang.frx":1A76
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRegistrasiRJPenunjang.frx":30D4
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
      TabIndex        =   55
      Top             =   1440
      Width           =   1605
   End
End
Attribute VB_Name = "frmRegistrasiRJPenunjang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFilter As String
Dim intRowNow As Integer
Dim strSubInstalasi As String
Dim strNoAntrian As String
Dim dTglberlaku As Date
Dim curTarif As Currency
Dim curTP As Currency
Dim curTRS As Currency
Dim curPemb As Currency
Dim Qstrsql As String
Public tempKelompokPasien As String
Dim strkdRuanganPerujuk As String
Dim strJK As String

'update aan fauzan 05/06/2013
Dim subPrinterZebra As Printer
Dim X As String
Dim temkdInstalasi As String
Dim strkdruangankasir As String
Dim dbRshy As New ADODB.recordset
Dim strcetakID As String
'Dim TempKodeKelas As String
Dim TempKodeSubinstalasi As String
Dim tempDaftarReservasi As Boolean
Dim TempKdRuanganPoliMCU As String
Dim TempKdRuanganPoliRehapMedik As String
Dim blnStatusRuangan As Boolean

Private Sub subCekJenisKelas()
    strSQL = "SELECT distinct KdKelas,Kelas FROM V_KelasPelayananRJPenunjang where Expr1='1'and DetailJenisJasaPelayanan LIKE '%" & dcJenisKelas.Text & "%' and KdInstalasi <> '03'  and kdruangan = '" & dcRuangan.BoundText & "'"
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = True Then
            dcKelas.Text = ""
            Exit Sub
        End If
        
        dcKelas.BoundText = rs(0).value
        dcKelas.Text = rs(1).value
        
        TempKodeKelas = rs(0).value
'        dcJenisKelas.BoundText = rs(0).value
'        dcJenisKelas.Text = rs(1).value
End Sub

Private Function sp_PasienRujukan() As Boolean
On Error GoTo hell
sp_PasienRujukan = True
Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaranTM.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCMTM.Text)
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, strkdRuanganPerujuk)
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, dcRuanganTM.BoundText)
        .Parameters.Append .CreateParameter("IdDokterPerujuk", adChar, adParamInput, 10, dcDokterPerujuk.BoundText)
        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , Format(dtpTglOrderTM.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dcSubInstalasi.BoundText)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PasienRujukan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam proses penyimpanan data", vbCritical, "Validasi"
            sp_PasienRujukan = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
hell:
    Call msubPesanError
    sp_PasienRujukan = False
    'Resume 0
End Function


Private Sub subclearkonsul()
    dcInstalasiTM.BoundText = ""
    dcInstalasiTM.Text = ""
    dcRuanganTM.BoundText = ""
    dcRuanganTM.Text = ""
    dcDokter.Text = ""
    dcDokter.BoundText = ""
    Call msubDcSource(dcDokterPerujuk, rs, "Select KodeDokter, NamaDokter FROM V_DaftarDokter order by NamaDokter")
End Sub

Private Sub cmdSimpanTM_Click()
    Dim adoCommand As New ADODB.Command
    If Periksa("datacombo", dcInstalasiTM, "Instalasi tujuan konsul masih kosong") = False Then Exit Sub
    If Periksa("datacombo", dcRuanganTM, "Ruangan tujuan konsul masih kosong") = False Then Exit Sub
    If Periksa("datacombo", dcDokterPerujuk, "Dokter perujuk masih kosong") = False Then Exit Sub
    
    If sp_PasienRujukan() = False Then Exit Sub
'    With adoCommand
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
'        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
'        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
'        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, mstrKdRuangan)
'        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, dcRuangan.BoundText)
'        .Parameters.Append .CreateParameter("IdDokterPerujuk", adChar, adParamInput, 10, dcDokter.BoundText)
'        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , Format(dtpTglOrderTM.value, "yyyy/MM/dd HH:mm:ss"))
'        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dcSubInstalasi.BoundText)
'        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
'
'        .ActiveConnection = dbConn
'        .CommandText = "dbo.Add_PasienRujukan"
'        .CommandType = adCmdStoredProc
'        .Execute
'        If Not (.Parameters("RETURN_VALUE").value = 0) Then
'            MsgBox "Ada Kesalahan dalam Penyimpanan Data", vbCritical, "Validasi"
'        Else
'            Call Add_HistoryLoginActivity("Add_PasienRujukan")
'        End If
'        Call deleteADOCommandParameters(adoCommand)
'        Set adoCommand = Nothing
'    End With
    
    MsgBox "Proses konsul berhasil", vbInformation, "Informasi"
    fraKonsul.Visible = False: Me.Enabled = True
    Frame2.Enabled = True: cmdTutup.Enabled = True
    Call subclearkonsul

End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub cmdTutupTM_Click()
    fraKonsul.Visible = False
    Frame1.Enabled = True
    Frame2.Enabled = True
End Sub

Private Sub dcDokter_Change()
 Call dcDokter_KeyPress(13)
End Sub

Private Sub dcDokter_LostFocus()
    mstrKdDokter = dcRuangan.BoundText
   
End Sub

Private Sub dcDokterPerujuk_GotFocus()
On Error GoTo errLoad
Dim tempKode As String
        tempKode = dcDokterPerujuk.BoundText
        strSQL = "Select KodeDokter, NamaDokter FROM V_DaftarDokter order by NamaDokter "
        Call msubDcSource(dcDokterPerujuk, rs, strSQL)
        dcDokterPerujuk.BoundText = tempKode
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcDokterPerujuk_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSimpanTM.SetFocus
End If
End Sub

Private Sub dcDokterPerujuk_LostFocus()
 Call dcDokterPerujuk_KeyPress(13)
    cmdSimpanTM.SetFocus
End Sub

Private Sub dcInstalasiTM_Change()
On Error Resume Next
    dcRuangan.BoundText = ""
End Sub

Private Sub dcInstalasiTM_GotFocus()
On Error GoTo errLoad
Dim tempKode As String
    
        tempKode = dcInstalasiTM.BoundText
        strSQL = "select distinct KdInstalasi, NamaInstalasi from V_RuanganTujuanRujukan WHERE StatusEnabled='1' "
        Call msubDcSource(dcInstalasiTM, rs, strSQL)
        dcInstalasiTM.BoundText = tempKode
        
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcInstalasiTM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    dcRuanganTM.SetFocus
End If
End Sub

Private Sub dcInstalasiTM_LostFocus()
Call dcInstalasiTM_KeyPress(13)
    dcRuanganTM.SetFocus
End Sub

Private Sub dcRuangan_Change()
    If dcRuangan.BoundText = TempKdRuanganPoliMCU Then
        strSQL = "Select QtyMax,KunjunganKe From PasienDaftarPaketPelayananTM where NoCM='" & txtNoCM.Text & "' and KdRuangan='" & dcRuangan.BoundText & "'"
        Call msubRecFO(rs, strSQL)
        
        strSQL2 = "Select Max(KunjunganKe) as KunjunganKe From PasienDaftarPaketPelayananTM where NoCM='" & txtNoCM.Text & "' and KdRuangan='" & dcRuangan.BoundText & "'"
        Call msubRecFO(rs2, strSQL2)

        If rs.EOF = False Then
            If rs2.Fields("KunjunganKe").value > rs.Fields("QtyMax").value Then
                lblKonsul.Caption = "Kontrol Ke: 1"
                txtMaks.Text = ""
                MsgBox "Pasien sudah melewati batas kunjungan maksimal", vbCritical, "Validasi": Exit Sub
                strSQL3 = "delete PasienDaftarPaketPelayananTM where NoCM='" & txtNoCM.Text & "' and KdRuangan='" & dcRuangan.BoundText & "'"
                Call msubRecFO(rs3, strSQL3)
                txtMaks.Text = ""
                 lblKonsul.Caption = "Kunj. ke: 1"
            End If
            lblKonsul.Caption = "Kontrol Ke: " & rs2.Fields("KunjunganKe").value + 1
            txtMaks.Enabled = False
            txtMaks.Text = rs.Fields("QtyMax").value
        End If
    Else
   End If
   
End Sub

Private Sub dcRuanganTM_Change()
    dtpTglOrderTM.Minute = Format(Now, "nn")
    dtpTglOrderTM.Second = Format(Now, "ss")
End Sub

Private Sub dcRuanganTM_GotFocus()
On Error GoTo errLoad
Dim tempKode As String
    tempKode = dcRuanganTM.BoundText
    strSQL = "select distinct KdRuangan, NamaRuangan from V_RuanganTujuanRujukan where KdInstalasi= '" & dcInstalasiTM.BoundText & "' and KdRuangan <> '" & mstrKdRuangan & "' and StatusEnabled ='1' order by NamaRuangan "
    Call msubDcSource(dcRuanganTM, rs, strSQL)
    dcRuanganTM.BoundText = tempKode
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcRuanganTM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    dcDokterPerujuk.SetFocus
End If
End Sub

Private Sub dcRuanganTM_LostFocus()
Call dcRuanganTM_KeyPress(13)
    dcDokterPerujuk.SetFocus
    
End Sub

Private Sub dcRujukanRI_Click(Area As Integer)
'Call dcRujukanRI_KeyPress(13)
End Sub

Private Sub dcSubInstalasi_Click(Area As Integer)
'    Call dcSubInstalasi_KeyPress(13)
End Sub

Private Sub txtNamaPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboJK.SetFocus
End Sub
Private Sub cboJK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtThn.SetFocus
End Sub

Private Sub subLoadData()
    sRuangPeriksa = dcRuangan.Text
    sNamaPasien = txtNamaPasien.Text
    sJK = cboJK.Text
    sUmur = txtThn.Text & " th " & txtBln.Text & " bl " & txtHr.Text & " hr"
    sAlamat = ""
    sPenjamin = dcKelompokPasien.Text
    sKelas = dcJenisKelas.Text
    sNoBed = dcNoBedRI.Text
    iNoAntrian = strNoAntrian
End Sub

'Store procedure untuk mengisi struk billing pasien
Private Function sp_AddStrukBuktiKasMasuk() As Boolean
On Error GoTo errLoad
    Dim strLokal As String
    sp_AddStrukBuktiKasMasuk = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglBKM", adDate, adParamInput, , Format(dtpTglPendaftaran.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdCaraBayar", adChar, adParamInput, 2, "01")
        .Parameters.Append .CreateParameter("KdJenisKartu", adChar, adParamInput, 2, Null)
        .Parameters.Append .CreateParameter("NamaBank", adVarChar, adParamInput, 100, Null)
        .Parameters.Append .CreateParameter("NoKartu", adVarChar, adParamInput, 50, Null)
        .Parameters.Append .CreateParameter("AtasNama", adVarChar, adParamInput, 50, Null)
       ' If CCur(txtJmlUang.Text) > CCur(lblTotalTagihan.Caption) Then
        '    .Parameters.Append .CreateParameter("JmlBayar", adCurrency, adParamInput, , CCur(lblTotalTagihan.Caption))
        'Else
        .Parameters.Append .CreateParameter("JmlBayar", adCurrency, adParamInput, , mcurAll_HrsDibyr)
        'End If
'        .Parameters.Append .CreateParameter("JmlBayar", adCurrency, adParamInput, , CCur(txtJmlUang.Text))
        .Parameters.Append .CreateParameter("Administrasi", adCurrency, adParamInput, , 0)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, Null)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, "176")
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, noidpegawai)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("OutputNoBKM", adChar, adParamOutput, 10, Null)
        
        .ActiveConnection = dbConn
        .CommandText = "Add_StrukBuktiKasMasukPelayananPasien"
        .CommandType = adCmdStoredProc
        .Execute
    
        If .Parameters("RETURN_VALUE").value <> 0 Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Struk Billing Pasien", vbCritical, "Validasi"
            sp_AddStrukBuktiKasMasuk = False
        Else
            If Not IsNull(.Parameters("OutputNoBKM").value) Then txtNoBKM.Text = .Parameters("OutputNoBKM").value
            If Len(txtNoBKM.Text) = 0 Then
                strLokal = "SELECT NoBKM from StrukBuktiKasMasuk where tglBKM = '" & Format(dtpTglPendaftaran.value, "yyyy/MM/dd HH:mm:ss") & "' and kdRuangan = '176' and idUser = '" & noidpegawai & "'"
                Call msubRecFO(rs, strLokal)
                txtNoBKM.Text = rs("NoBKM").value
            End If
            Call Add_HistoryLoginActivity("Add_StrukBuktiKasMasukPelayananPasien")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    
Exit Function
errLoad:
    sp_AddStrukBuktiKasMasuk = False
    Call msubPesanError("-Add_StrukBuktiKasMasukPelayananPasien")
End Function

'Store procedure untuk mengisi struk billing pasien
Private Function sp_AddStruk(ByVal adoCommand As ADODB.Command, strStsByr As String) As Boolean
    On Error GoTo errLoad
    Dim strLokal As String
    sp_AddStruk = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoBKM", adChar, adParamInput, 10, mstrNoBKM)
        .Parameters.Append .CreateParameter("OutputNoStruk", adChar, adParamOutput, 10, Null)
        .Parameters.Append .CreateParameter("TglStruk", adDate, adParamInput, , Format(dtpTglPendaftaran.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, mstrNoCM)
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, dcKelompokPasien.BoundText)
        If dcKelompokPasien.BoundText = "01" Then
            .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, "2222222222")
        Else
            .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, typAsuransi.strIdPenjamin)
        End If
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, "176")
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, noidpegawai)
        .Parameters.Append .CreateParameter("TotalBiaya", adCurrency, adParamInput, , CCur(mcurBayar))
        .Parameters.Append .CreateParameter("JmlHutangPenjamin", adCurrency, adParamInput, , CCur(mcurAll_TP))
        .Parameters.Append .CreateParameter("JmlTanggunganRS", adCurrency, adParamInput, , CCur(mcurAll_TRS))
        .Parameters.Append .CreateParameter("JmlPembebasan", adCurrency, adParamInput, , CCur(mcurAll_Pemb))
        .Parameters.Append .CreateParameter("JmlHrsDibayar", adCurrency, adParamInput, , CCur(mcurAll_HrsDibyr))
        .Parameters.Append .CreateParameter("JmlDiscount", adCurrency, adParamInput, , "0")
      
        .ActiveConnection = dbConn
        .CommandText = "Add_NoStrukPelayananPasien"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Struk Billing Pasien", vbCritical, "Validasi"
            sp_AddStruk = False
        Else
            If Not IsNull(.Parameters("OutputNoStruk").value) Then mstrNoStruk = .Parameters("OutputNoStruk").value
            If Len(mstrNoStruk) = 0 Then
                strLokal = "SELECT NoStruk from StrukPelayananPasien where tglStruk = '" & Format(dtpTglPendaftaran.value, "yyyy/MM/dd HH:mm:ss") & "' and NoPendaftaran = '" & mstrNoPen & "' and NoCM = '" & mstrNoCM & "' and idUser = '" & noidpegawai & "'"
                Call msubRecFO(rs, strLokal)
                mstrNoStruk = rs("NoStruk").value
            End If
            Call Add_HistoryLoginActivity("Add_NoStrukPelayananPasien")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Function
errLoad:
    msubPesanError ("-Add_NoStrukPelayananPasien")
End Function

Private Sub CboNamaPoli_Keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    dcSubInstalasi.SetFocus
End If
End Sub

Private Sub chkDetailPasien_Click()
    If chkDetailPasien.value = 1 Then
        strPasien = "View"
        strRegistrasi = "PasienLama"
        Load frmPasienBaru
        frmPasienBaru.Show
    Else
        Unload frmPasienBaru
        Unload frmDetailPasien
    End If
End Sub

Private Sub chkDetailPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglPendaftaran.SetFocus
End Sub

Private Sub chkDiriSendiri_Click()
    On Error GoTo errLoad
    If chkDiriSendiri.value = vbChecked Then
        strSQL = "SELECT NamaLengkap, Alamat, Telepon,Propinsi,Kota,Kecamatan,Kelurahan,RTRW,Kodepos FROM Pasien WHERE NocM='" & txtNoCM.Text & "'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount <> 0 Then
            txtNamaRI.Text = rs("NamaLengkap").value
            txtAlamatRI.Text = IIf(IsNull(rs("Alamat").value), "-", rs("Alamat").value)
            txtTlpRI.Text = IIf(IsNull(rs("Telepon")), "-", rs("Telepon").value)
            dcPropinsiPJ.Text = IIf(IsNull(rs("Propinsi")), "-", rs("Propinsi"))
            dcKotaPJ.Text = IIf(IsNull(rs("Kota")), "-", rs("Kota"))
            dcKecamatanPJ.Text = IIf(IsNull(rs("Kecamatan")), "-", rs("Kecamatan"))
            dcKelurahanPJ.Text = IIf(IsNull(rs("Kelurahan")), "-", rs("Kelurahan"))
            
            'load Pekerjaan Pasien
            strSQL = "SELECT Pekerjaan FROM detailPasien WHERE NocM='" & txtNoCM.Text & "'"
            Call msubRecFO(rs, strSQL)
            dcPekerjaanPJ.Text = IIf(rs.RecordCount = 0, "-", rs("Pekerjaan"))
            
        Else
        txtNamaRI.Text = ""
        txtAlamatRI.Text = ""
        txtTlpRI.Text = ""
        dcPropinsiPJ.Text = ""
        dcKotaPJ.Text = ""
        dcKecamatanPJ.Text = ""
        dcKelurahanPJ.Text = ""
        End If
    Else
        txtNamaRI.Text = ""
        txtAlamatRI.Text = ""
        txtTlpRI.Text = ""
        dcPropinsiPJ.Text = ""
        dcKotaPJ.Text = ""
        dcKecamatanPJ.Text = ""
        dcKelurahanPJ.Text = ""
    End If
    dcHubungan.BoundText = ""
    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub chkDiriSendiri_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If chkDiriSendiri.value = vbChecked Then
            cmdSimpan.SetFocus
        Else
            txtNamaRI.SetFocus
        End If
    End If
End Sub

Private Sub chkPaket_Click()
    If chkPaket.value = 1 Then
        FraMaks.Enabled = True
'        txtMaks.SetFocus
        dcPaket.Enabled = True
        dcPaket.SetFocus
    Else
        FraMaks.Enabled = False
        txtMaks.Text = ""
        dcPaket.Enabled = False
        dcPaket.Text = ""
    End If
End Sub

Private Sub cmdAsuransiP_Click()
On Error GoTo hell_
    'Me.Enabled = False
    mblnTemp = True
    mstrNoPen = ""
    mstrNoCM = txtNoCM.Text
    mstrKdJenisPasien = dcKelompokPasien.BoundText
    mstrKdInstalasi = dcInstalasi.BoundText
'    mstrKdPenjaminPasien = "2222222222"
    With frmUbahJenisPasien
        .Show
        .txtNamaFormPengirim.Text = "tampung"
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = txtNamaPasien.Text
        .txtJK.Text = cboJK.Text
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHr.Text
        .txttglpendaftaran.Text = dtpTglPendaftaran.value
        .lblNoPendaftaran.Visible = False
        .txtNoPendaftaran.Visible = False
        .dcJenisPasien.BoundText = mstrKdJenisPasien
'        If mblnAdmin = True Then
'            .dcJenisPasien.Enabled = True
'        Else
'            .dcJenisPasien.Enabled = False
'        End If
        .dcPenjamin.BoundText = mstrKdPenjaminPasien
        .dcAsalRujukan.BoundText = dcRujukanRI.BoundText
        .dcJenisPasien.Enabled = False
        
'strSQLX = "Select IdAsuransi From AsuransiPasien Where NoCM = '" & mstrNoCM & "' "
strSQLX = "Select TOP 1 IdAsuransi From PemakaianAsuransi Where NoCM = '" & mstrNoCM & "' order by TglSJP DESC"
            Call msubRecFO(rsx, strSQLX)
            
            If rsx.EOF = False Then
                .txtNoKartuPA.Text = rsx(0).value
            End If
        .txtNoRujukan.Text = txtNoRujukan.Text
        .txtNoKartuPA.SetFocus
    End With
Exit Sub
hell_:
    msubPesanError
End Sub

'Private Sub cmdOrder_Click()
'On Error GoTo Gabril
'    fraOrder.Visible = True
'    fraOrder.Top = -120
'    fraOrder.Left = 2760
'    optTindakan.value = False
'    optObat.value = False
'Exit Sub
'Gabril:
'    Call msubPesanError
'End Sub

Private Sub cmdRujukan_Click()
On Error GoTo hell_
If dcRujukanRI.BoundText = "01" Then cmdTutup.SetFocus: Exit Sub   ' datang sendiri"
    With frmRujukan
        .Show
        .txtNoCM.Text = txtNoCM
        .txtNamaPasien.Text = txtNamaPasien.Text
        .txtJK.Text = cboJK.Text
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHr.Text
        .txtNoPendaftaran.Text = txtNoPendaftaran.Text
        .dcRujukanAsal.Text = dcRujukanRI.Text
        mstrKdInstalasiPerujuk = dcRujukanRI.BoundText
    End With
Exit Sub
hell_:
    msubPesanError
End Sub

Private Sub cmdSimpan_Click()
On Error Resume Next
Dim strAntrian As String
blnSibuk = True

    
    If dcInstalasi.BoundText = "03" Then
       MsgBox "Registrasi Rawat Inap tidak bisa dilakukan di Form ini", vbCritical, "Validasi"
       Exit Sub
    End If
 
    cmdRujukan.Enabled = False
    If funcCekValidasi = False Then Exit Sub
    If dcInstalasi.BoundText = "04" Then
        MsgBox "Maaf pasien tidak dapat langsung dioperasi Harus ada Rujukan Dari Dokter...", vbCritical, "Validasi"
        Exit Sub
    End If
    Call msubRecFO(dbRst, "SELECT IdPenjamin FROM PenjaminKelompokPasien WHERE KdKelompokPasien = '" & strKdKelompokPasien & "'")
    If dbRst.EOF = True Then
        MsgBox "Lengkapi dulu data panjamin pasien" & vbNewLine & "" & dcKelompokPasien.Text & "", vbExclamation, "Validasi"
        dcKelompokPasien.SetFocus
        Exit Sub
    End If
    If dbRst(0).value <> "2222222222" And typAsuransi.blnSuksesAsuransi = False Then
        cmdSimpan.SetFocus
        'simpan data penjamin
        mstrKdInstalasi = dcInstalasi.BoundText
        mstrKdRuanganPasien = dcRuangan.BoundText
        Call cmdAsuransiP_Click 'mstrKdPenjaminPasien selalu 2222222222
        Call SubLoadAsuransi
        Exit Sub
    End If

    If dcInstalasi.BoundText = "03" Then
        mstrKdInstalasi = dcInstalasi.BoundText
        'validasi data registrasi ri
        If Periksa("datacombo", dcCaraMasukRI, "Asal Rujukan belum di isi") = False Then Exit Sub
      '  If Periksa("datacombo", dcKelasKamarRI, "Kelas kamar belum di isi") = False Then Exit Sub
      '  If Periksa("datacombo", dcNoKamarRI, "Nomor kamar belum di isi") = False Then Exit Sub
      '  If Periksa("datacombo", dcNoBedRI, "Nomor bed belum di isi") = False Then Exit Sub
      '  If Periksa("text", txtNamaRI, "Nama penanggung jawab belum di isi") = False Then Exit Sub
      '  If Periksa("text", txtAlamatRI, "Alamat penanggung jawab?") = False Then Exit Sub
      '  If Len(Trim(dcHubungan.Text)) > 0 Then
      '      If Periksa("datacombo", dcHubungan, "Data hubungan peserta pasien belum di isi") = False Then Exit Sub
      '  End If

    '        strSQL = "SELECT StatusBed FROM StatusBed WHERE (KdKamar = '" & dcNoKamarRI.BoundText & "') AND (NoBed = '" & dcNoBedRI.BoundText & "') and StatusEnabled='1'"
    '        Call msubRecFO(rs, strSQL)
    '        If UCase(rs(0).value) = "I" Then
    '            MsgBox "No. Bed sudah terpakai", vbInformation, "Informasi"
    '            strSQL = "SELECT distinct dbo.StatusBed.NoBed, dbo.StatusBed.NoBed AS Alias, dbo.StatusBed.StatusEnabled" & _
    '                " FROM dbo.NoKamar INNER JOIN dbo.StatusBed ON dbo.NoKamar.KdKamar = dbo.StatusBed.KdKamar" & _
    '                " WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND (KdKelas = '" & dcKelasKamarRI.BoundText & "') AND (dbo.NoKamar.KdKamar = '" & dcNoKamarRI.BoundText & "') AND (dbo.StatusBed.StatusBed = 'K') and StatusEnabled='1'"
    '            Call msubDcSource(dcNoBedRI, rs, strSQL)
    '            Exit Sub
    '        End If
    End If

    
'================================== add By Dayz =======================================
'If dcRuangan.BoundText = "220" Then 'Or dcInstalasi.BoundText = "06" Then
'    Call sp_KonsulTemp(dbcmd)
'End If

'If dcRuangan.BoundText = "220" Then
'    strSQL = "Select QtyMax,KunjunganKe From PasienDaftarPaketPelayananTM where NoCM='" & txtNoCM.Text & "' and KdRuangan='" & dcRuangan.BoundText & "'"
'    Call msubRecFO(rs, strSQL)
'    If rs.EOF = False Then
'        lblKonsul.Visible = True
'        lblKonsul.Caption = "Kontrol Ke - " & rs.Fields("KunjunganKe").value + 1
'
'        If rs.Fields("KunjunganKe").value >= rs.Fields("QtyMax").value Then
'            MsgBox "Pasien sudah melewati batas kunjungan maksimal", vbCritical, "Validasi": Exit Sub
'
'        End If
'    End If
    
    If dcRuangan.BoundText = TempKdRuanganPoliMCU Then
        If txtMaks.Text = "" Then
            MsgBox "Kunjungan Maksimal kosong", vbCritical, "Validasi": Exit Sub
        End If
        If dcPaket.Text = "" Then
            MsgBox "Paket Pelayanan kosong", vbCritical, "Validasi": Exit Sub
        End If
        
    End If
'End If
If dcRuangan.BoundText = TempKdRuanganPoliMCU Then
     strSQL = "Select QtyMax,KunjunganKe From PasienDaftarPaketPelayananTM where NoCM='" & txtNoCM.Text & "' and KdRuangan='" & dcRuangan.BoundText & "'"
     Call msubRecFO(rs, strSQL)
    
    strSQL2 = "Select Max(KunjunganKe) AS KunjunganKe From PasienDaftarPaketPelayananTM where NoCM='" & txtNoCM.Text & "' and KdRuangan='" & dcRuangan.BoundText & "'"
    Call msubRecFO(rs2, strSQL2)
    If rs.EOF = False Then
    
        If txtMaks.Text = "" Then
            MsgBox "Kunjungan Maksimal kosong", vbCritical, "Validasi": Exit Sub
        End If
        If dcPaket.Text = "" Then
            MsgBox "Paket Pelayanan kosong", vbCritical, "Validasi": Exit Sub
        End If
    
        lblKonsul.Visible = True
        lblKonsul.Caption = "Kontrol Ke - " & rs2.Fields("KunjunganKe").value + 1
        
        If rs2.Fields("KunjunganKe").value >= rs.Fields("QtyMax").value Then
            MsgBox "Pasien sudah melewati batas kunjungan maksimal", vbCritical, "Validasi": Exit Sub
            
        End If
    End If
        
   
End If
'
'Else
    'simpan data registrasi
'================================================= Untuk Validasi Pendaftaran =========================================================
    strSQL = "SELECT NoCM, Ruangan, NamaDokter, TglMasuk FROM V_DaftarAntrianPasienMRS WHERE (NoCM = '" & txtNoCM.Text & "') AND kdruangan = '" & dcRuangan.BoundText & "' AND TglMasuk BETWEEN '" & Format(dtpTglPendaftaran, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpTglPendaftaran, "yyyy/MM/dd 23:59:59") & "' and [Status Periksa] <> 'Sudah' and NoBKM is not null"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
    MsgBox "Pasien tersebut sudah terdaftar di Rawat Jalan," & vbNewLine & "Ruangan " & rs("Ruangan") & " ", vbCritical, "Perhatian"
            mstrNoCM = ""
            txtNoCM = ""
            subClearData
            fraDokter.Visible = False
            chkDetailPasien.Enabled = False
            txtNoCM.SetFocus
        Exit Sub
    End If
    
        strSQL = "SELECT NoCM, Title + ' ' + [Nama Lengkap] AS NamaPasien FROM V_CariPasien WHERE ([No. CM] = '" & txtNoCM.Text & "' )"
        Call msubRecFO(rsCek, strSQL)
        
        If rsCek.EOF = False Then
            Set rs = Nothing
            strSQL = "Select NoAntrian From ReservasiPasien Where KdRuangan = '" & dcRuangan.BoundText & "' and TglMasuk between '" & Format(dtpTglPendaftaran, "yyyy/mm/dd 00:00:00") & "' And '" & Format(dtpTglPendaftaran, "yyyy/mm/dd hh:mm:ss") & "' "
            Call msubRecFO(rs, strSQL)

            If rs.EOF = False Then
                strAntrian = rs.Fields(0)
            End If


            Call sp_RegistrasiAll(dbcmd)
        Else
            With frmPasienBaru
                .Show
                .txtFormPengirim.Text = Me.Name
                .txtKdAntrian.Text = txtKdAntrian.Text
                .chkNoCM.value = vbChecked
                .txtNoCM.Text = txtNoCM.Text
                .txtNamaPasien.Text = txtNamaPasien.Text
                .cboJnsKelaminPasien.Text = cboJK.Text
                .txtTahun.Text = txtThn.Text
                .txtBulan.Text = txtBln.Text
                .txthari.Text = txtHr.Text
            End With
        Exit Sub
        End If
        cmdSimpan.Enabled = False
        
'End If
'    Call sp_RegistrasiAll(dbcmd)

'=================================== End Dayz ============================================
    If txtNoPendaftaran = "" Then
        MsgBox "No Pendaftaran kosong", vbExclamation, "Validasi"
        Exit Sub
    End If
'====== Add By Dayz ==========
    If dcRuangan.BoundText = TempKdRuanganPoliMCU Then ' Or dcInstalasi.BoundText = "06" Then
'        Call sp_KonsulTemp(dbcmd)
        Call sp_PasienDaftarPaketPelayananTM(dbcmd)
        
        
        strSQL = "Select QtyMax,KunjunganKe From PasienDaftarPaketPelayananTM where NoCM='" & txtNoCM.Text & "' and KdRuangan='" & dcRuangan.BoundText & "'"
        Call msubRecFO(rs, strSQL)
        
        strSQL2 = "Select Max(KunjunganKe) as KunjunganKe From PasienDaftarPaketPelayananTM where NoCM='" & txtNoCM.Text & "' and KdRuangan='" & dcRuangan.BoundText & "'"
        Call msubRecFO(rs2, strSQL2)
        
        If rs.EOF = False Then
'            If rs2.Fields("KunjunganKe").value = rs.Fields("QtyMax").value Then
'                strSQL3 = "delete PasienDaftarPaketPelayananTM where NoCM='" & txtNoCM.Text & "' and KdRuangan='" & dcRuangan.BoundText & "'"
'                Call msubRecFO(rs3, strSQL3)
'
'            End If
        End If
        
    End If
'===== end Dayz =====
    If dcInstalasi.BoundText = "03" Then
        'simpan registrasi pasien RI
        Call sp_RegistrasiPasienRI(dbcmd)
        
        'simpan pasien masuk kamar
        Call sp_PasienMasukKamar(dbcmd)
    End If

    Call msubRecFO(dbRst, "SELECT IdPenjamin FROM PenjaminKelompokPasien WHERE KdKelompokPasien = '" & strKdKelompokPasien & "'")
    If dbRst(0).value <> "2222222222" Then
         'If sp_JenisPasienJoinProgramAskes = False Then Exit Sub
        Call sp_AsuransiPasien(dbcmd)
    End If

    
    If chkPrintKartu.value = 1 Then
        Call sp_TempPrintKartu("A", txtNoPendaftaran.Text)
    End If
    
    If dcInstalasi.BoundText <> "16" Then
    
        If sp_PelayananOtomatis() = False Then Exit Sub
        'jika triger insert to pasien belum bayar gagal
        'add 2008-06-04 onede
        Set rs = Nothing
        strSQL = "select NoPendaftaran from PasienBelumBayar where NoPendaftaran='" & txtNoPendaftaran.Text & "'AND NoCM ='" & txtNoCM.Text & "'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dbConn.Execute "Insert into PasienBelumbayar values('" & txtNoPendaftaran.Text & "','" & txtNoCM.Text & "') "
        End If
        
        strCetak = ""
        If dcInstalasi.BoundText = "01" Then
            strCetak = "IGD"
        ElseIf dcInstalasi.BoundText = "04" Or dcInstalasi.BoundText = "09" Or dcInstalasi.BoundText = "10" Then
            strCetak = "Penunjang"
            Set rs = Nothing
            strSQL = "select KdRuangan from RuanganLangsungBayar where KdRuangan='" & mstrKdRuangan & "' and KdKelompokPasien='" & strKdKelompokPasien & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
            Call msubRecFO(dbRshy, "select KdRuanganKasir from RuanganLangsungBayar where KdRuangan='" & mstrKdRuangan & "' and KdKelompokPasien='" & strKdKelompokPasien & "'")
            strkdruangankasir = dbRshy(0).value
                dbConn.Execute "execute PostingHutangPenjaminPasien_AU '" & txtNoPendaftaran.Text & "','A' "
                dbConn.Execute "execute Add_StrukPelayananPasienBayarByBackOfficeNew '" & Format(dtpTglPendaftaran.value, "yyyy/MM/dd HH:mm:ss") & "','05','" & txtNoPendaftaran.Text & "','" & txtNoCM.Text & "','Daftar Bayar','" & strkdruangankasir & "','" & strIDPegawai & "' "
                
    
            End If
        Else
            strCetak = "RJ"
            Set rs = Nothing
            strSQL = "select KdRuangan from RuanganLangsungBayar where KdRuangan='" & mstrKdRuangan & "' and KdKelompokPasien='" & strKdKelompokPasien & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
            Call msubRecFO(dbRshy, "select KdRuanganKasir from RuanganLangsungBayar where KdRuangan='" & mstrKdRuangan & "' and KdKelompokPasien='" & strKdKelompokPasien & "'")
            strkdruangankasir = dbRshy(0).value
                dbConn.Execute "execute PostingHutangPenjaminPasien_AU '" & txtNoPendaftaran.Text & "','A' "
                dbConn.Execute "execute Add_StrukPelayananPasienBayarByBackOfficeNew '" & Format(dtpTglPendaftaran.value, "yyyy/MM/dd HH:mm:ss") & "','05','" & txtNoPendaftaran.Text & "','" & txtNoCM.Text & "','Daftar Bayar','" & strkdruangankasir & "','" & strIDPegawai & "' "
    
            End If
        End If
        
    End If

    cmdSimpan.Enabled = False

    Call subEnableButtonReg(True)
    If dcInstalasi.BoundText = "09" Or dcInstalasi.BoundText = "04" Or dcInstalasi.BoundText = "10" Or dcInstalasi.BoundText = "16" Then
        cmdTindakan.Enabled = True
        cmdTindakan.SetFocus
    Else
        cmdTindakan.Enabled = False
    End If
cmdRujukan.Enabled = True
''splakuk
If dcInstalasi.BoundText = "03" Then
    Call Add_HistoryLoginActivity("Add_RegistrasiPasienMRS+Add_RegistrasiPasienRI+Add_PasienMasukKamar")
Else
    Call Add_HistoryLoginActivity("Add_RegistrasiPasienMRS")
End If

'untuk update no antrian
'If Periksa("text", txtKdAntrian, "Nomor Antrian belum di isi") = False Then Exit Sub
If bolAntrian = True Then
    If TempKodeKelas = "" Then
        If Update_AntrianPasienRegistrasi(txtKdAntrian.Text, txtNoCM, dcRuangan.BoundText, dcKelas.BoundText, dcKelompokPasien.BoundText, txtNoPendaftaran.Text, "SELESAI") = False Then Exit Sub
    Else
        If Update_AntrianPasienRegistrasi(txtKdAntrian.Text, txtNoCM, dcRuangan.BoundText, TempKodeKelas, dcKelompokPasien.BoundText, txtNoPendaftaran.Text, "SELESAI") = False Then Exit Sub
    End If
End If

strNamaDokter = dcDokter.Text

Dim path As String
     Dim pathtemp As String
    
    strSQL = "select Value from SettingGlobal where Prefix='PathSdkAntrian'"
    Call msubRecFO(rs, strSQL)
      
    If Not rs.EOF Then
        If rs(0).value <> "" Then
            path = rs(0).value
        End If
    End If
    
'    strSQL = "select StatusAntrian from SettingDataUmum"
'    Call msubRecFO(rs, strSQL)
'    Dim coba As Long
'    If Not rs.EOF Then
'        If rs(0).value = "1" Then
'            If Dir(path) <> "" Then
'                path = path + " Type:" & Chr(34) & "Update Patient" & Chr(34)
'                coba = Shell(path, vbNormalFocus)
'            End If
'        Else
'            txtKdAntrian.Text = strNoAntrian
'        End If
'    End If
'---------------------------------------------------------------
    strSQL = "select StatusAntrian from SettingDataUmum"
    Call msubRecFO(rs, strSQL)
    Dim coba1 As Long
    If Not rs.EOF Then
        If rs(0).value = "1" Then
            If Dir(path) <> "" Then
                path = path + " Type:" & Chr(34) & "Update Patient" & Chr(34)
'                path = pathtemp + " Type:" & Chr(34) & "Update Patient" & Chr(34)
                coba1 = Shell(path, vbNormalFocus)
                strSQL = "select * from AntrianEndpoint  where kdRuangan='" & dcRuangan.BoundText & "'"
                Call msubRecFO(rs, strSQL)
                If (rs.EOF = False) Then
'                    path = pathtemp & " endpoint:" & rs("EndPointAntrian") & " loket:" & dcRuangan.BoundText & " Type:" & Chr(34) & "Update Patient" & Chr(34)
                    path = path & " endpoint:" & rs("EndPointAntrian") & " loket:" & dcRuangan.BoundText & " Type:" & Chr(34) & "Update Patient" & Chr(34)
                    coba1 = Shell(path, vbNormalFocus)
                End If
                
            End If
        Else
            txtKdAntrian.Text = strNoAntrian
        End If
    End If
    
' chandra 27 02 2014
    ' Tambahan untuk mengupdate no pendaftaran jika pasien reservasi
    If strNoReservasi <> "" Then
        strSQL = "update ReservasiPasien set NoPendaftaran='" & txtNoPendaftaran.Text & "', StatusDaftar='Y'  where NoReservasi='" & strNoReservasi & "'"
        Call msubRecFO(rs, strSQL)
    End If

If dcInstalasi.BoundText = "03" Then Call Form_KeyDown(121, 2)
mstrKdAsalRujukan = dcRujukanRI.BoundText
mstrNoPen = txtNoPendaftaran.Text


frm_cetak_label_viewer.Show
'frm_cetak_label_viewer.Cetaklangsung

If chkKonsul.value = Checked Then
    Frame1.Enabled = False
    Frame2.Enabled = False
    fraKonsul.Visible = True
    dtpTglOrderTM.value = dtpTglPendaftaran.value
    txtNoCMTM.Text = txtNoCM.Text
    txtNoPendaftaranTM.Text = txtNoPendaftaran.Text
    txtNamaPasienTM.Text = txtNamaPasien.Text
    txtSex.Text = cboJK.Text
    txtthntm.Text = txtThn.Text
    txtblntm.Text = txtblntm.Text
    txthrtm.Text = txtHr.Text
    strkdRuanganPerujuk = dcRuangan.BoundText
'    Call msubDcSource(dcDokterPerujuk, rs, "Select KodeDokter, NamaDokter FROM V_DaftarDokter order by NamaDokter ")
End If

Exit Sub
errLoad:
'    Resume 0
'    Call msubPesanError
'    cmdSimpan.Enabled = True
'    blnSibuk = False

End Sub

Private Sub cmdTindakan_Click()
On Error GoTo errLoad
If txtNoCM.Text = "" Then Exit Sub
    
    mstrNoPen = txtNoPendaftaran.Text
    With frmKonsul_OrderPelayanan
        .Show
        .txtNoCMTM.Text = txtNoCM.Text
        .txtNoPendaftaranTM.Text = mstrNoPen

        .txtNamaPasien.Text = txtNamaPasien.Text
        .txtSex.Text = cboJK.Text
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHr.Text
        .dcDokterPerujuk.Enabled = False
        .dcRuangan.Enabled = False
        .txtDokter.Enabled = False
        .dcInstalasi.Enabled = False
        .dcInstalasi.BoundText = dcInstalasi.BoundText
        .dcRuangan.BoundText = dcRuangan.BoundText
        .dcTempatPerujuk.BoundText = dcRujukanRI.BoundText
'        .dcTempatPerujuk.Text = dcRujukanRI.Text
        .dcTempatPerujuk.SetFocus

        mstrKdKelas = dcKelas.BoundText
        mstrKdSubInstalasi = dcSubInstalasi.BoundText
        mstrValid = ""
'        .dcInstalasi.BoundText = dcInstalasi.BoundText
'        .dcRuangan.BoundText = dcRuangan.BoundText
'        .dcDokterPerujuk.BoundText = dcDokter.BoundText
        
        .txtKdDokterTM.Text = mstrKdDokter
        .txtDokter.Text = dcDokter.Text
        
        .txtNamaForm.Text = Me.Name
        .fraDokter.Visible = False
        
    End With
    cmdTindakan.Enabled = False
    Me.Enabled = False
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
On Error GoTo hell
    
'    If cmdSimpan.Enabled = True And txtNamaPasien.Text <> "" Then
'        If MsgBox("Simpan Data Registrasi Pasien ", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
'            Call CmdSimpan_Click
'            Exit Sub
'        End If
'    End If
If txtFormPengirim.Text = "frmDaftarReservasiPasien" Then frmDaftarReservasiPasien.cmdCari_Click

If cmdSimpan.Enabled = True Then
    Unload Me
Else
    If dcInstalasi.BoundText = "09" Or dcInstalasi.BoundText = "16" Then
        strSQL = "Select * from BiayaPelayanan where NoPendaftaran='" & txtNoPendaftaran.Text & "' and (year(TglPelayanan)='" & Format(dtpTglPendaftaran, "yyyy") & "' and month(TglPelayanan)='" & Format(dtpTglPendaftaran, "MM") & "' and day(TglPelayanan)='" & Format(dtpTglPendaftaran, "dd") & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True And flagJikaPasienSudahRujukan = False Then
           If MsgBox("Apakah anda akan keluar tanpa melengkapi Tindakan Pelayanan ?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
                Unload Me
           Else
                cmdTindakan.SetFocus
           End If
            Exit Sub
        End If
    End If
'    If dcInstalasi.BoundText = "04" Then
'        strSQL = "Select * from RegistrasiIBS where NoPendaftaran='" & txtnopendaftaran.Text & "' and (year(TglPendaftaran)='" & Format(dtpTglPendaftaran, "yyyy") & "' and month(TglPendaftaran)='" & Format(dtpTglPendaftaran, "MM") & "' and day(TglPendaftaran)='" & Format(dtpTglPendaftaran, "dd") & "')"
'        Call msubRecFO(rs, strSQL)
'        If rs.EOF = True Then
'            MsgBox "Pasien belum terdaftar di ruang Bedah Sentral, silahkan lengkapi Tindakan Pelayanan", vbCritical, "Validasi"
'            cmdTindakan.SetFocus
'            Exit Sub
'        End If
'    End If
    If dcInstalasi.BoundText = "10" Then
        strSQL = "Select * from BiayaPelayanan where NoPendaftaran='" & txtNoPendaftaran.Text & "' and (year(TglPelayanan)='" & Format(dtpTglPendaftaran, "yyyy") & "' and month(TglPelayanan)='" & Format(dtpTglPendaftaran, "MM") & "' and day(TglPelayanan)='" & Format(dtpTglPendaftaran, "dd") & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            MsgBox "Pasien belum terdaftar di ruang Radiologi, silahkan lengkapi Tindakan Pelayanan", vbCritical, "Validasi"
            cmdTindakan.SetFocus
            Exit Sub
        End If
    End If
    Unload Me
End If
Exit Sub
hell:
End Sub

Private Sub Command1_Click()
If cmdSimpan.Enabled = True Then Exit Sub
    mstrNoCM = Trim(txtNoCM)
    frmCetakCatatanMedis.Show
End Sub

Private Sub dcCaraMasukRI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcCaraMasukRI.MatchedWithList = True Then dcKelasKamarRI.SetFocus
        strSQL = "SELECT KdCaraMasuk, CaraMasuk FROM CaraMasuk where StatusEnabled='1' and (CaraMasuk LIKE '%" & dcCaraMasukRI.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcCaraMasukRI.Text = ""
        Exit Sub
        End If
        dcCaraMasukRI.BoundText = rs(0).value
        dcCaraMasukRI.Text = rs(1).value
    End If
End Sub

Private Sub dcHubungan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then dcPekerjaanPJ.SetFocus
End Sub

Private Sub dcHubungan_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad

    If KeyAscii = 13 Then
        If Len(Trim(dcHubungan.Text)) = 0 Then cmdSimpan.SetFocus
        If dcHubungan.MatchedWithList = True Then dcPekerjaanPJ.SetFocus
        strSQL = "SELECT Hubungan, NamaHubungan FROM HubunganKeluarga WHERE (NamaHubungan LIKE '%" & dcHubungan.Text & "%') and StatusEnabled='1'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcHubungan.Text = ""
        Exit Sub
        End If
        dcHubungan.BoundText = rs(0).value
        dcHubungan.Text = rs(1).value
       
    End If

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcInstalasix()
On Error GoTo errLoad

    dcJenisKelas.BoundText = ""
    If dcInstalasi.BoundText = "02" Then 'RJ
        dcJenisKelas.BoundText = "01" 'UMUM
    Else
        dcJenisKelas.BoundText = ""
    End If
    dcSubInstalasi.BoundText = ""
'    Call subTampilRegistrasiRI
    Call subDcSource
Exit Sub
errLoad:
    Call msubPesanError
End Sub


Private Sub dcInstalasi_Change()
On Error GoTo errLoad

    If dcRuangan.BoundText = TempKdRuanganPoliMCU Then
        
        strSQL = "Select QtyMax,KunjunganKe From PasienDaftarPaketPelayananTM where NoCM='" & txtNoCM.Text & "' and KdRuangan='" & dcRuangan.BoundText & "'"
        Call msubRecFO(rs, strSQL)
        
          strSQL2 = "Select Max(KunjunganKe) as KunjunganKe From PasienDaftarPaketPelayananTM where NoCM='" & txtNoCM.Text & "' and KdRuangan='" & dcRuangan.BoundText & "'"
        Call msubRecFO(rs2, strSQL2)
        If rs.EOF = False Then
            
            lblKonsul.Caption = "Kontrol Ke: " & rs2.Fields("KunjunganKe").value + 1
            txtMaks.Enabled = False
            txtMaks.Text = rs.Fields("QtyMax").value
            If rs2.Fields("KunjunganKe").value >= rs.Fields("QtyMax").value Then
                MsgBox "Pasien sudah melewati batas kunjungan maksimal", vbCritical, "Validasi": 'Exit Sub
'                 strSQL3 = "delete PasienDaftarPaketPelayananTM where NoCM='" & txtNoCM.Text & "' and KdRuangan='" & dcRuangan.BoundText & "'"
'                Call msubRecFO(rs3, strSQL3)
'                txtMaks.Text = ""
'                 lblKonsul.Caption = "Kunj. ke: 1"

            'End If
                
            End If
        End If
    End If
    

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcInstalasi_GotFocus()
On Error GoTo errLoad
Dim tempKode As String

    tempKode = dcInstalasi.BoundText

    dcInstalasi.BoundText = tempKode
    mstrKdInstalasi = tempKode
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcInstalasi.BoundText = "41" Or dcInstalasi.BoundText = "06" Then
            chkPaket.Enabled = True
        Else
            chkPaket.Enabled = False
        End If
        
        dcJenisKelas.SetFocus

    End If
End Sub

Private Sub dcKecamatanPJx()
    strSQL = "SELECT DISTINCT KdKelurahan, NamaKelurahan FROM Kelurahan where KdKecamatan = '" & dcKecamatanPJ.BoundText & "' and StatusEnabled='1' order by NamaKelurahan"
    Call msubDcSource(dcKelurahanPJ, rs, strSQL)
    If rs.RecordCount <> 0 Then
        dcKelurahanPJ.Text = rs("NamaKelurahan")
    Else
        dcKelurahanPJ.Text = ""
    End If
End Sub


Private Sub dcJenisKelas_LostFocus()
If dcJenisKelas.MatchedWithList = True Then dcRuangan.SetFocus
Call subCekJenisKelas
End Sub

Private Sub dcKecamatanPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKecamatanPJ.MatchedWithList = True Then dcKelurahanPJ.SetFocus
        strSQL = "SELECT DISTINCT KdKecamatan, NamaKecamatan FROM Kecamatan where KdKotaKabupaten = '" & dcKotaPJ.BoundText & "' and StatusEnabled='1' and (NamaKecamatan LIKE '%" & dcKecamatanPJ.Text & "%')order by NamaKecamatan"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcKecamatanPJ.Text = ""
        Exit Sub
        End If
        dcKecamatanPJ.BoundText = rs(0).value
        dcKecamatanPJ.Text = rs(1).value
        Call dcKecamatanPJx
    End If
End Sub

Private Sub dcKecamatanPJ_LostFocus()
    dcKecamatanPJ = Trim(StrConv(dcKecamatanPJ, vbProperCase))
End Sub

Private Sub dcKelas_LostFocus()
        If dcJenisKelas.MatchedWithList = True Then dcRuangan.SetFocus
        strSQL = "SELECT distinct KdDetailJenisJasaPelayanan,DetailJenisJasaPelayanan FROM V_InstalasiJenisKelas Where KdInstalasi='" & dcInstalasi.BoundText & "'"

        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcKelas.Text = ""
        Exit Sub
        End If
        dcKelas.BoundText = rs(0).value
        dcKelas.Text = rs(1).value

End Sub

Private Sub dcKelasKamarRI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKelasKamarRI.MatchedWithList = True Then dcNoKamarRI.SetFocus
        strSQL = "SELECT Kdkelas, kelas FROM V_KelasPelayanan WHERE (KdInstalasi IN ('" & dcInstalasi.BoundText & "','08')) AND (KdDetailJenisJasaPelayanan  = '" & dcJenisKelas.BoundText & "') AND (KdKelas IN ('" & dcKelas.BoundText & "','04')) AND (kelas LIKE '%" & dcKelasKamarRI.Text & "%') and Expr3='1'"
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

Private Sub dcKelurahanPJx()
    strSQL = "SELECT DISTINCT KodePos FROM Kelurahan where KdKelurahan = '" & dcKelurahanPJ.BoundText & "' and StatusEnabled='1'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then
        If IsNull(rs("KodePos")) = True Then
            txtKodePos.Text = ""
        Else
            txtKodePos.Text = rs("KodePos")
        End If
    Else
        txtKodePos.Text = ""
    End If
End Sub

Private Sub dcKelurahanPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then meRTRWPJ.SetFocus
End Sub

Private Sub dcKelurahanPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKelurahanPJ.MatchedWithList = True Then meRTRWPJ.SetFocus
        strSQL = "SELECT DISTINCT KdKelurahan, NamaKelurahan FROM Kelurahan where KdKecamatan = '" & dcKecamatanPJ.BoundText & "' and StatusEnabled='1' and (NamaKelurahan LIKE '%" & dcKelurahanPJ.Text & "%')order by NamaKelurahan"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcKelurahanPJ.Text = ""
        Exit Sub
        End If
        dcKelurahanPJ.BoundText = rs(0).value
        dcKelurahanPJ.Text = rs(1).value
        Call dcKelurahanPJx
    End If
End Sub

Private Sub dcKelurahanPJ_LostFocus()
    dcKelurahanPJ = Trim(StrConv(dcKelurahanPJ, vbProperCase))
End Sub

Private Sub dcKotaPJx()
    strSQL = "SELECT DISTINCT KdKecamatan, NamaKecamatan FROM Kecamatan where KdKotaKabupaten = '" & dcKotaPJ.BoundText & "' and StatusEnabled='1' order by NamaKecamatan"
    Call msubDcSource(dcKecamatanPJ, rs, strSQL)
    If rs.RecordCount <> 0 Then
        dcKecamatanPJ.Text = rs("NamaKecamatan")
    Else
        dcKecamatanPJ.Text = ""
    End If
End Sub



Private Sub dcKotaPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKotaPJ.MatchedWithList = True Then dcKecamatanPJ.SetFocus
        strSQL = "SELECT DISTINCT KdKotaKabupaten, NamaKotaKabupaten FROM KotaKabupaten where KdPropinsi = '" & dcPropinsiPJ.BoundText & "' and StatusEnabled='1' and (NamaKotaKabupaten LIKE '%" & dcKotaPJ.Text & "%')order by NamaKotaKabupaten"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcKotaPJ.Text = ""
        Exit Sub
        End If
        dcKotaPJ.BoundText = rs(0).value
        dcKotaPJ.Text = rs(1).value
        Call dcKotaPJx
    End If
End Sub

Private Sub dcKotaPJ_LostFocus()
    dcKotaPJ = Trim(StrConv(dcKotaPJ, vbProperCase))
End Sub

Private Sub dcNoBedRI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcNoBedRI.MatchedWithList = True Then optTidak.SetFocus
    strSQL = "SELECT distinct dbo.StatusBed.NoBed, dbo.StatusBed.NoBed AS Alias, dbo.StatusBed.StatusEnabled, dbo.NoKamar.StatusEnabled as Expr" & _
        " FROM dbo.NoKamar INNER JOIN dbo.StatusBed ON dbo.NoKamar.KdKamar = dbo.StatusBed.KdKamar" & _
        " WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND (KdKelas = '" & dcKelasKamarRI.BoundText & "') AND (dbo.NoKamar.KdKamar = '" & dcNoKamarRI.BoundText & "') and (NoBed LIKE '%" & dcNoBedRI.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcNoBedRI.Text = ""
        Exit Sub
        End If
        dcNoBedRI.BoundText = rs(0).value
        dcNoBedRI.Text = rs(1).value
    End If
End Sub

Private Sub dcNoBedRI_LostFocus()
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
    
 '============================================== Added By Dayz ==============================================
    Set rs1 = Nothing
    strSQL = "Select Distinct JenisKelamin from V_ValidasiJenisKelamin Where KdKamar='" & dcNoKamarRI.BoundText & "'"
    Call msubRecFO(rs1, strSQL)
    If rs1.EOF = False Then
        strJK = rs1.Fields(0).value
    Else
        strJK = "-"
    End If

    strSQL = "Select Distinct JenisKelamin from V_ValidasiJenisKelamin Where JenisKelamin='" & strJK & "' and KdKamar='" & dcNoKamarRI.BoundText & "'"
    Call msubRecFO(rsCek, strSQL)
    If rsCek.EOF = False Then
        If cboJK.Text = strJK Then
            dcNoBedRI.SetFocus
        Else
            MsgBox "Jenis Kelamin Berbeda", vbInformation
            dcNoKamarRI.SetFocus
            dcNoBedRI.Text = ""
        End If
    Else
        dcNoBedRI.SetFocus
    End If
'============================================== End Dayz =================================================
    End If
Exit Sub
hell:
    msubPesanError
End Sub



Private Sub dcPaket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtMaks.Text <> "" Then
            If (cmdSimpan.Enabled = True) Then
                cmdSimpan.SetFocus
            End If
        Else
            FraMaks.Enabled = True
            txtMaks.Enabled = True
            txtMaks.SetFocus
        End If
    End If
End Sub

Private Sub dcPaket_LostFocus()
On Error GoTo hell
    If dcRuangan.BoundText = "223" Then
    strSQL = "Select * from SettingGlobal where prefix='kdRuanganMCU'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
      txtMaks.Text = "0"
      txtMaks.Enabled = False
    
    End If
End If
Exit Sub
hell:
    msubPesanError
End Sub

Private Sub dcPekerjaanPJ_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtAlamatRI.SetFocus
End Sub

Private Sub dcPekerjaanPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcPekerjaanPJ.MatchedWithList = True Then txtAlamatRI.SetFocus
        strSQL = "SELECT DISTINCT KdPekerjaan,Pekerjaan FROM Pekerjaan where StatusEnabled='1' and (Pekerjaan LIKE '%" & dcPekerjaanPJ.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcPekerjaanPJ.BoundText = rs(0).value
        dcPekerjaanPJ.Text = rs(1).value
    End If
End Sub

Private Sub dcPropinsiPJx()
    strSQL = "SELECT DISTINCT KdKotaKabupaten, NamaKotaKabupaten FROM KotaKabupaten where KdPropinsi = '" & dcPropinsiPJ.BoundText & "' and StatusEnabled='1' order by NamaKotaKabupaten"
    Call msubDcSource(dcKotaPJ, rs, strSQL)
    If rs.RecordCount <> 0 Then
        dcKotaPJ.Text = rs("NamaKotaKabupaten")
    Else
        dcKotaPJ.Text = ""
    End If
End Sub

Private Sub dcPropinsiPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcPropinsiPJ.MatchedWithList = True Then dcKotaPJ.SetFocus
        strSQL = "SELECT DISTINCT KdPropinsi, NamaPropinsi FROM Propinsi where StatusEnabled='1' and  (NamaPropinsi LIKE '%" & dcPropinsiPJ.Text & "%')order by NamaPropinsi"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcPropinsiPJ.Text = ""
        Exit Sub
        End If
        dcPropinsiPJ.BoundText = rs(0).value
        dcPropinsiPJ.Text = rs(1).value
        Call dcPropinsiPJx
    End If
End Sub

Private Sub dcPropinsiPJ_LostFocus()
    dcPropinsiPJ = Trim(StrConv(dcPropinsiPJ, vbProperCase))
End Sub

Private Sub dcRuangan_LostFocus()
On Error GoTo errLoad
If dcRuangan.MatchedWithList = True Then dcRujukanRI.SetFocus
    Call msubDcSource(dcSubInstalasi, rs, "SELECT KdSubInstalasi, NamaSubInstalasi FROM V_SubInstalasiRuangan WHERE KdRuangan = '" & dcRuangan.BoundText & "' and StatusEnabled='1' ORDER BY NamaSubInstalasi")
    If rs.EOF = True Then
    dcSubInstalasi.Text = ""
    dcInstalasi.Text = ""
    Exit Sub

    End If
    
        dcSubInstalasi.BoundText = rs(0).value
        dcSubInstalasi.Text = rs(1).value
        
        TempKodeSubinstalasi = dcSubInstalasi.BoundText
        
    strSQLX = "select KdInstalasi from Ruangan where KdRuangan like '%" & dcRuangan.BoundText & "%'"
    Call msubRecFO(rsx, strSQLX)
    If rs.EOF = False Then
        dcInstalasi.BoundText = rsx(0).value
    End If
    
 If dcRuangan.BoundText = TempKdRuanganPoliMCU Then
    strSQL = "Select * from SettingGlobal where prefix='kdRuanganMCU'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
      txtMaks.Text = "0"
      txtMaks.Enabled = False
    
    End If
  End If
  
   If dcRuangan.BoundText = TempKdRuanganPoliMCU Then
    Call dcRuangan_Change
    'Call dcJenisKelas_LostFocus
   End If
   Call subCekJenisKelas
 Exit Sub
errLoad:
 Call msubPesanError
End Sub

Private Sub dcRujukanRI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcRujukanRI.MatchedWithList = True Then dcKelompokPasien.SetFocus
        strSQL = "SELECT KdRujukanAsal, RujukanAsal FROM RujukanAsal where StatusEnabled='1' and (RujukanAsal LIKE '%" & dcRujukanRI.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcRujukanRI.Text = ""
        Exit Sub
        End If
        dcRujukanRI.BoundText = rs(0).value
        dcRujukanRI.Text = rs(1).value
    End If
End Sub

Private Sub dcRujukanRI_LostFocus()
'    If dcInstalasi.BoundText = "09" Or dcInstalasi.BoundText = "04" Or dcInstalasi.BoundText = "10" Or dcInstalasi.BoundText = "16" Then
'        Call msubDcSource(dcRujukanRI, rs, "SELECT KdRujukanAsal, RujukanAsal FROM RujukanAsal where StatusEnabled='1'")
'    End If
'        If dcRujukanRI.MatchedWithList = True Then dcKelompokPasien.SetFocus
        strSQL = "SELECT KdRujukanAsal, RujukanAsal FROM RujukanAsal where StatusEnabled='1' and (RujukanAsal LIKE '%" & dcRujukanRI.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcRujukanRI.Text = ""
        Exit Sub
        End If
        dcRujukanRI.BoundText = rs(0).value
        dcRujukanRI.Text = rs(1).value
End Sub

Private Sub dcSubInstalasi_GotFocus()
On Error GoTo errLoad
Dim tempKode As String
    
    tempKode = dcSubInstalasi.BoundText
    Call msubDcSource(dcSubInstalasi, rs, "SELECT KdSubInstalasi, NamaSubInstalasi FROM V_SubInstalasiRuangan WHERE KdRuangan = '" & dcRuangan.BoundText & "' and StatusEnabled='1' ORDER BY NamaSubInstalasi")
    If rs.EOF = False Then
        dcSubInstalasi.BoundText = rs.Fields(0).value
        dcSubInstalasi.Text = rs.Fields(1).value
        dcSubInstalasi.BoundText = tempKode
        dcRujukanRI.SetFocus
    Else
        dcSubInstalasi.Enabled = True
        dcSubInstalasi.SetFocus
    End If
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcJenisKelas_Change()
'    dcKelas.Text = ""
End Sub

Private Sub dcJenisKelas_GotFocus()
On Error GoTo errLoad
If dcRuangan.Text = "" Then
    tempKode = dcJenisKelas.BoundText
    strSQL = "SELECT distinct KdDetailJenisJasaPelayanan,DetailJenisJasaPelayanan FROM V_KelasPelayanan " ' where KdDetailJenisJasaPelayanan NOT IN('01','02','03','04') "

    Call msubDcSource(dcJenisKelas, rs, strSQL)
    dcJenisKelas.BoundText = tempKode
End If
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcJenisKelas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcJenisKelas.MatchedWithList = True Then dcRuangan.SetFocus
        strSQL = "SELECT distinct KdDetailJenisJasaPelayanan,DetailJenisJasaPelayanan FROM V_KelasPelayanan where KdDetailJenisJasaPelayanan is not null "

        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcKelas.Text = ""
        Exit Sub
        End If
        dcKelas.BoundText = rs(0).value
        dcKelas.Text = rs(1).value
        
        dcRuangan.SetFocus
       
    End If
End Sub

Private Sub dcKelas_Change()
    'dcRuangan.Text = ""
End Sub

'Private Sub dcKelas_GotFocus()
'On Error GoTo errLoad
'Dim TempKode As String
'
'    TempKode = dcKelas.BoundText
'
'    strSQL = "SELECT distinct KdKelas, Kelas FROM V_KelasPelayanan WHERE KdInstalasi = '" & dcInstalasi.BoundText & "' and KdDetailJenisJasaPelayanan = '" & dcJenisKelas.BoundText & "' AND KdKelas<>04 and Expr2='1'"
'    Call msubDcSource(dcKelas, rs, strSQL)
'
'    dcKelas.BoundText = TempKode
'    dcKelas.BoundText = rs(0).value
'    dcKelas.Text = rs(1).value
'    dcRuangan.SetFocus
'Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub
'
'Private Sub dcKelas_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
''        If dcKelas.MatchedWithList = True Then dcRuangan.SetFocus
'        strSQL = "SELECT distinct KdKelas, Kelas FROM V_KelasPelayanan WHERE KdInstalasi = '" & dcInstalasi.BoundText & "' and KdDetailJenisJasaPelayanan ='" & dcJenisKelas.BoundText & "' AND KdKelas<>04 and Expr2='1' and Kelas LIKE '" & dcKelas.Text & "'"
'        Call msubRecFO(rs, strSQL)
'        If rs.EOF = True Then
'        dcKelas.Text = ""
'        Exit Sub
'        End If
'        dcKelas.BoundText = rs(0).value
'        dcKelas.Text = rs(1).value
'        dcRuangan.SetFocus
'    End If
'End Sub

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
            " WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND (KdKelas IN ('" & dcKelas.BoundText & "','04')) and StatusEnabled='1'"
    Else
        strSQL = "SELECT DISTINCT KdKelas, Kelas " & _
            " FROM V_KamarRegRawatInap " & _
            " WHERE KdRuangan = '" & dcRuangan.BoundText & "' AND (KdKelas in ('" & dcKelas.BoundText & "','04')) and StatusEnabled='1'"
    End If
    
    Call msubDcSource(dcKelasKamarRI, rs, strSQL)
    dcKelasKamarRI.BoundText = tempKdKelas

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKelasKamarRI_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then dcNoKamarRI.SetFocus
End Sub

Private Sub dcKelompokPasien_Change()
   strKdKelompokPasien = dcKelompokPasien.BoundText
   typAsuransi.blnSuksesAsuransi = False
End Sub

Private Sub dcKelompokPasien_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 13 Then
        strSQL = "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where StatusEnabled='1' and JenisPasien LIKE '%" & dcKelompokPasien.Text & "%' order by JenisPasien "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKelompokPasien.BoundText = rs(0).value
        dcKelompokPasien.Text = rs(1).value
        
        If dcKelompokPasien.Text = "" Then
            dcKelompokPasien.SetFocus
            Exit Sub
        End If
        strKdKelompokPasien = dcKelompokPasien.BoundText
        Call msubRecFO(dbRst, "SELECT IdPenjamin FROM PenjaminKelompokPasien WHERE KdKelompokPasien = '" & strKdKelompokPasien & "'")
        If dbRst.EOF = True Then
            MsgBox "Lengkapi dulu data Penjamin Kelompok Pasien " & vbNewLine & "" & dcKelompokPasien.Text & "", vbExclamation, "Validasi"
            dcKelompokPasien.SetFocus
            Exit Sub
        End If
        If dbRst(0).value <> "2222222222" And typAsuransi.blnSuksesAsuransi = False Then
            Call cmdAsuransiP_Click
            Call SubLoadAsuransi
        Else
'            Call subTampilRegistrasiRI
            If dcInstalasi.BoundText = "03" Then
                dcCaraMasukRI.SetFocus
                dcDokter.SetFocus
            Else
'                cmdSimpan.SetFocus
                dcDokter.SetFocus
            End If
        End If
    End If
Exit Sub
errLoad:
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

Private Sub dcNoBedRI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then optTidak.SetFocus
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
    
'    strSQL = "Select Distinct JenisKelamin from V_ValidasiJenisKelamin Where JenisKelamin='" & cboJK.Text & "' and KdKamar='" & dcNoKamarRI.BoundText & "'"
'    Call msubRecFO(rsCek, strSQL)
'    If rsCek.EOF = False Then
'        dcNoBedRI.SetFocus
'    Else
'        MsgBox "Jenis Kelamin Berbeda", vbInformation
'        dcNoKamarRI.SetFocus
'        dcNoBedRI.Text = ""
'    End If
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcNoKamarRI_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then dcNoBedRI.SetFocus
End Sub

Private Sub dcRuanganx()
On Error GoTo errLoad
    
    If dcInstalasi.BoundText = "03" Then
        Call msubDcSource(dcKelasKamarRI, rsb, "SELECT DISTINCT KdKelas, Kelas FROM V_KamarRegRawatInap WHERE (KdRuangan = '" & dcRuangan.BoundText & "')")
        dcKelasKamarRI.BoundText = ""
    End If

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcRuangan_GotFocus()
On Error GoTo errLoad
'Dim TempKode As String
If dcRuangan.Text = "" Then
    tempKode = dcRuangan.BoundText
'    If dcInstalasi.BoundText = "03" Then
'        strSQL = "SELECT distinct KdRuangan, NamaRuangan FROM V_KelasPelayanan WHERE (KdInstalasi = '" & dcInstalasi.BoundText & "') AND (KdDetailJenisJasaPelayanan  = '" & dcJenisKelas.BoundText & "') AND ((KdKelas = '" & dcKelas.BoundText & "') OR KdKelas='04') and Expr3='1' ORDER BY NamaRuangan"
'    Else
'        strSQL = "SELECT distinct KdRuangan, NamaRuangan FROM V_KelasPelayanan WHERE (KdInstalasi = '" & dcInstalasi.BoundText & "') AND (KdDetailJenisJasaPelayanan  = '" & dcJenisKelas.BoundText & "') AND (KdKelas = '" & dcKelas.BoundText & "') and Expr3='1' ORDER BY NamaRuangan"
'    End If
'    If dcInstalasi.BoundText = "03" Then
'        strSQL = "SELECT distinct KdRuangan, NamaRuangan FROM V_KelasPelayanan WHERE (KdInstalasi = '" & dcInstalasi.BoundText & "') AND (KdDetailJenisJasaPelayanan  = '" & dcJenisKelas.BoundText & "') OR (KdKelas='04') and Expr3='1' ORDER BY NamaRuangan"
'    Else
       ' strSQL = "SELECT distinct KdRuangan, NamaRuangan FROM V_KelasPelayanan WHERE (KdDetailJenisJasaPelayanan  = '" & dcJenisKelas.BoundText & "') and Expr3='1' ORDER BY NamaRuangan"
        strSQL = "SELECT KdRuangan,NamaRuangan FROM V_RuanganPemeriksaanRegistrasiPenunjang  WHERE KdDetailJenisJasaPelayanan = '" & dcJenisKelas.BoundText & "'"
'    End If

    Call msubDcSource(dcRuangan, rs, strSQL)
    dcRuangan.BoundText = tempKode
End If
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcRuangan_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 13 Then
'        strSQL = "SELECT KdRuangan, NamaRuangan FROM V_KelasPelayanan WHERE (KdInstalasi IN ('" & dcInstalasi.BoundText & "','08')) AND (KdDetailJenisJasaPelayanan  = '" & dcJenisKelas.BoundText & "') AND (KdKelas IN ('" & dcKelas.BoundText & "','04')) AND (NamaRuangan LIKE '%" & dcRuangan.Text & "%') and Expr3='1'"
       If dcRuangan.MatchedWithList = True Then dcRujukanRI.SetFocus
        strSQL = "SELECT KdRuangan,NamaRuangan FROM V_RuanganPemeriksaanRegistrasiPenunjang  WHERE (KdDetailJenisJasaPelayanan = '" & dcJenisKelas.BoundText & "') AND (NamaRuangan LIKE '%" & dcRuangan.Text & "%') "

        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcRuangan.Text = ""
        Exit Sub
        End If
        dcRuangan.BoundText = rs(0).value

    Call dcRuanganx
        strSQLX = "select KdInstalasi from Ruangan where KdRuangan like '%" & dcRuangan.BoundText & "%'"
    Call msubRecFO(rsx, strSQLX)
    If rs.EOF = False Then
        dcInstalasi.BoundText = rsx(0).value
        tempKode = rsx(0).value
    End If
    End If
Exit Sub
errLoad:
End Sub

Private Sub dcRujukanRI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcKelompokPasien.SetFocus
End Sub

Private Sub dcSubInstalasi_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 13 Then
        strSQL = "SELECT KdSubInstalasi, NamaSubInstalasi FROM V_SubInstalasiRuangan WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND (NamaSubInstalasi LIKE '%" & dcSubInstalasi.Text & "%') and StatusEnabled='1'"
        Call msubRecFO(dbRst, strSQL)
        If dbRst.EOF = True Then
        dcSubInstalasi.Text = ""
        Exit Sub
        End If
        dcSubInstalasi.BoundText = dbRst(0).value
        dcSubInstalasi.Text = dbRst(1).value
        dcRujukanRI.SetFocus
    End If
Exit Sub
errLoad:
    Call msubPesanError
End Sub
Private Sub SubLoadAsuransi()
On Error GoTo errLoad
            
strSQL = "Select * from V_DaftarAsuransi where NoCM = '" & txtNoCM.Text & "'"
Call msubRecFO(rs, strSQL)

If rs.EOF = False Then
    With frmUbahJenisPasien
        .Show
        If (tempKelompokPasien = dcKelompokPasien.BoundText) Then
            .dcPenjamin.Text = rs.Fields("NamaPenjamin")
            .dcPerusahaan.Text = rs.Fields("InstitusiAsal")
            .dcGolonganAsuransi.Text = rs.Fields("NamaGolongan")
            .txtNoKartuPA.Text = rs.Fields("IdAsuransi")
            .txtNipPA.Text = rs.Fields("IdPeserta")
        End If
        .txtNoRujukan.Text = txtNoRujukan.Text
        .txtNamaPA.Text = rs.Fields("NamaPeserta")
        .dtpTglLahirPA.value = rs.Fields("TglLahir")
        .txtAlamatPA.Text = rs.Fields("Alamat")
    End With
End If

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcSubInstalasi_LostFocus()
On Error GoTo errLoad
Dim tempKode As String
    
    tempKode = dcSubInstalasi.BoundText
    Call msubDcSource(dcSubInstalasi, rs, "SELECT KdSubInstalasi, NamaSubInstalasi FROM V_SubInstalasiRuangan WHERE KdRuangan = '" & dcRuangan.BoundText & "' and StatusEnabled='1' ORDER BY NamaSubInstalasi")
    If rs.EOF = False Then
        tempKode = rs.Fields(0).value
        dcSubInstalasi.BoundText = rs.Fields(0).value
        dcSubInstalasi.Text = rs.Fields(1).value
        dcSubInstalasi.BoundText = tempKode
        dcRujukanRI.SetFocus
    Else
        dcSubInstalasi.Enabled = True
        dcSubInstalasi.SetFocus
    End If
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Private Sub dgDokter_DblClick()
'On Error GoTo gabril
'    txtDokter.Text = dgDokter.Columns("NamaLengkap").value
'    txtKdDokter.Text = dgDokter.Columns("KdDokter").value
'    fraDokter.Visible = False
'Exit Sub
'gabril:
'    Call msubPesanError
'End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
'
'    txtDokter.Text = dgDokter.Columns("NamaLengkap").value
'    txtKdDokter.Text = dgDokter.Columns("KdDokter").value
'    fraDokter.Visible = False
If dcInstalasi.BoundText <> "06" Or dcInstalasi.BoundText <> "41" Then
    cmdSimpan.SetFocus
Else
    chkPaket.SetFocus
End If

End Sub
Private Sub dtpTglPendaftaran_Change()
    dtpTglPendaftaran.MaxDate = Now
End Sub

Private Sub dtpTglPendaftaran_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcJenisKelas.SetFocus
'    If KeyCode = 13 Then dcInstalasi.SetFocus
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo hell
Dim strCtrlKey As String
Dim intTracer As String
    'deklarasi tombol control ditekan
    strCtrlKey = (Shift + vbCtrlMask)

Select Case KeyCode
'        Case vbKeyF1
        'Jarakal
        Case vbKeyF1
            'Jarakal
            If cmdSimpan.Enabled = True Then Exit Sub
            mstrNoPen = frmRegistrasiRJPenunjang.txtNoPendaftaran.Text
            mstrKdInstalasi = frmRegistrasiRJPenunjang.dcInstalasi.BoundText
            frm_cetak_label_viewer.Show
'            frm_cetak_label_viewer.Cetaklangsung
        Case vbKeyB
            If strCtrlKey = 4 Then
                Unload Me
               'Call cmdTutup_Click

                strPasien = "Baru"
                frmPasienBaru.Show
'                Unload Me
            End If
        Case vbKeyF2
            Unload Me
            frmCariPasien.Show
        Case vbKeyJ
            If strCtrlKey = 4 Then
                Unload Me
                strPasien = "Lama"
                frmRegistrasiRJPenunjang.Show
            End If
'        Case vbKeyR
'            If txtnopendaftaran.Text = "" Then Exit Sub
'            If dcInstalasi.BoundText <> "03" Then Exit Sub
'               'If dcInstalasi.BoundText <> "02" Then Exit Sub
'            'End If
'            frmCetakLembarMasukDanKeluarV2.Show
        Case vbKeyZ
            If txtNoPendaftaran.Text = "" Then
                MsgBox "Harus daftar terlebih dahulu", vbExclamation, "Validasi"
                Exit Sub
            End If
            mstrNoPen = txtNoPendaftaran.Text
            'If dcInstalasi.BoundText <> "03" Then Exit Sub
            mstrNamaDokter = dcDokter.Text
            frmCetakSuratKeterangan.Show
        Case vbKeyF9
            If cmdSimpan.Enabled = True Then Exit Sub

            strSQL = "select *  from SettingGlobal where Prefix = 'KdKelompokPasienUmum'"

            Call msubRecFO(rsCek, strSQL)
            If rsCek.EOF = False Then
                If rsCek.Fields("Value").value <> dcKelompokPasien.BoundText Then
                    If mstrNoSJP = "" Then
                        MsgBox "No SJP kosong", vbExclamation, "Validasi"
                        Exit Sub
                    End If
                    frmViewerSJP.Show
                End If
            End If


    End Select
Exit Sub
hell:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglPendaftaran.value = Now
    
    ' untuk mendapatkan jumlah panjang NoCM pada setting global
    strSQL = "Select value from SettingGlobal where Prefix = 'LenNoCM'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        strBanyakNoCM = rs(0).value
'    Else
'        strBanyakNoCM = "6"
    End If
    
    txtNoCM.MaxLength = strBanyakNoCM
    
    
    ' untuk mendapatkan kode ruangan poli MCU dan setting global di sesuaikan dengan kode ruangan MCU
    strSQL = "Select Value from SettingGlobal where prefix='KdRuanganPoliMCU'"
    Call msubRecFO(rs1, strSQL)
    If rs1.EOF = False Then
        TempKdRuanganPoliMCU = rs1.Fields("Value").value
    End If

    ' untuk mendapatkan kode ruangan poli Rehap medik dan setting global di sesuaikan dengan kode ruangan Rehap medik
    strSQL = "Select Value from SettingGlobal where prefix='KdRuanganPoliRehapMedik'"
    Call msubRecFO(rs2, strSQL)
    If rs1.EOF = False Then
        TempKdRuanganPoliRehapMedik = rs2.Fields("Value").value
    End If

    
    flagJikaPasienSudahRujukan = False
    strSQLinst = ""
    If mstrKdRuangan = strKdRuanganRekamMedis Then ' untuk ruangan rekam Medis
        strSQLinst = "SELECT DISTINCT KdInstalasi,NamaInstalasi FROM V_KelasPelayanan"
    ElseIf mstrKdRuangan = strKdRuanganRegistrasiRJ Then ' untuk ruangan registrasi RJ
        strSQLinst = "SELECT DISTINCT KdInstalasi,NamaInstalasi FROM Instalasi where KdInstalasi in('06','01','22','02','09')"
    ElseIf mstrKdRuangan = strKdRuanganRegistrasiRI Then ' untuk ruangan registrasi RI
        strSQLinst = "SELECT DISTINCT KdInstalasi,NamaInstalasi FROM Instalasi where KdInstalasi in('08','09') "
    Else
        strSQLinst = "SELECT DISTINCT KdInstalasi,NamaInstalasi FROM V_KelasPelayanan"
    End If
    
    'strSQLinst = "SELECT DISTINCT KdInstalasi,NamaInstalasi FROM V_KelasPelayanan"
    Call msubDcSource(dcInstalasi, rsint, strSQLinst)
    
    strRegistrasi = "RJ"
    If mblnCariPasien = True Then frmCariPasien.Enabled = False
    Call subDcSource
'    Call subTampilRegistrasiRI
    If bolAntrian = True Then
        txtKdAntrian.Enabled = True
    Else
        txtKdAntrian.Enabled = False
    End If
    
'    dcInstalasi.BoundText = "02"
    dcKelompokPasien.BoundText = "01"
    
     txtKdAntrian.Text = strnoAntrianPasien
'    Call subClearData
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
''txtDokter.Visible = False
'Label20.Visible = False
'
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnCariPasien = True Then frmCariPasien.Enabled = True
    If cmdSimpan.Enabled = True And txtNamaPasien.Text <> "" Then
        If MsgBox("Simpan Data Registrasi Pasien ", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    
        
End Sub



Private Sub subLoadDokter()
On Error GoTo gabril

txthari.Text = Format(dtpTglPendaftaran, "DDDD")
strhari = txthari.Text
    strSQL = "SELECT KdDokter,NamaLengkap FROM V_JadwalPraktekDokter  where NamaRuangan='" & dcRuangan.Text & "' " & _
    "and hari ='" & strhari & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDokter = rs.RecordCount
If Not rs.EOF Then
    dcDokter.BoundText = rs.Fields(0).value
    dcDokter.Text = rs.Fields(1).value
    mstrKdDokter = rs.Fields(0).value
'        Set dgDokter.DataSource = rs
'        With dgDokter
'            .Columns(0).Width = 3000 'nama dokter
'            .Columns(1).Width = 0 'kode dokter
'            .Columns(2).Width = 2500
''            .Columns(3).Width = 3000
'        End With
'
'        fraDokter.Left = 6600
'        fraDokter.Top = Frame3.Top + txtDokter.Top + txtDokter.Height
'        fraDokter.Visible = True
'        fraDokter = 1
Else
    If dcInstalasi.BoundText <> "02" Then
        dcDokter.SetFocus
    Else
        MsgBox "Tidak ada dokter yang prakter di hari, ruangan, dan poli yang dipilih", vbInformation, "Jadwal Praktek Dokter"
        dcDokter.Text = ""
        dcDokter.SetFocus
    End If
End If
Exit Sub
gabril:
    Call msubPesanError
End Sub
Private Sub dcDokter_GotFocus()
On Error GoTo gabril
 'chandra 27 02 2014
    ' untuk handle dokter lab penunjang
        Dim strSQL2 As String
        Dim hari As String
        Dim dayOfWeek  As Integer
        dayOfWeek = Weekday(Now)
        
        If (dayOfWeek = DayConstants.mvwFriday) Then
            hari = "JUMAT"
        ElseIf (dayOfWeek = DayConstants.mvwMonday) Then
            hari = "SENIN"
        ElseIf (dayOfWeek = DayConstants.mvwSaturday) Then
            hari = "SABTU"
        ElseIf (dayOfWeek = DayConstants.mvwSunday) Then
            hari = "MINGGU"
        ElseIf (dayOfWeek = DayConstants.mvwThursday) Then
            hari = "KAMIS"
        ElseIf (dayOfWeek = DayConstants.mvwTuesday) Then
            hari = "SELASA"
        ElseIf (dayOfWeek = DayConstants.mvwWednesday) Then
            hari = "RABU"
        End If
        
        
        strSQL = "SELECT KdDokter,NamaLengkap FROM V_JadwalPraktekDokter  where NamaRuangan='" & dcRuangan.Text & "' AND hari='" & hari & "' "
        Call msubRecFO(rsCek, strSQL)
        If Not rsCek.EOF Then
            
            strSQL = "SELECT KdDokter,NamaLengkap FROM V_JadwalPraktekDokter  where NamaRuangan='" & dcRuangan.Text & "'  AND hari='" & hari & "'  "
        Else
            strSQL = "SELECT KodeDokter,NamaDokter FROM V_DaftarDokter "
        End If
        
        strSQL2 = "SELECT     Ruangan.KdInstalasi FROM  Ruangan INNER JOIN  SettingGlobal ON Ruangan.KdInstalasi = SettingGlobal.Value where SettingGlobal.Prefix= 'KdInstalasiLaboratorium' and Ruangan.KdRuangan='" & dcRuangan.BoundText & "'"
        Call msubRecFO(rs, strSQL2)
        ' cek jika laboratory
        If (Not rs.EOF) Then
            strSQL = "select idPegawai as KodeDokter, NamaLengkap as NamaDokter from datapegawai where idPegawai='1111111111'"
        Else
            strSQL2 = "SELECT     Ruangan.KdInstalasi FROM  Ruangan INNER JOIN  SettingGlobal ON Ruangan.KdInstalasi = SettingGlobal.Value where SettingGlobal.Prefix= 'KdInstalasiRadiology' and Ruangan.KdRuangan='" & dcRuangan.BoundText & "'"
            Set rs = Nothing
            Call msubRecFO(rs, strSQL2)
            ' jika radiology
            If (Not rs.EOF) Then
                strSQL = "select idPegawai as KodeDokter, NamaLengkap as NamaDokter from datapegawai where idPegawai='2222222222'"
            End If
        End If
        
        Set rs = Nothing
        Call msubDcSource(dcDokter, rs, strSQL)
        dcDokter.BoundText = rs.Fields(0).value
        dcDokter.Text = rs.Fields(1).value
        
        mstrKdDokter = rs.Fields(0).value
    
Exit Sub
gabril:
    Call msubPesanError
End Sub
Private Sub dcDokter_KeyPress(KeyAscii As Integer)
On Error GoTo gabril
If KeyAscii = 13 Then
    If dcRuangan.BoundText = TempKdRuanganPoliMCU Or dcRuangan.BoundText = TempKdRuanganPoliRehapMedik Then
        chkPaket.Enabled = True
        chkPaket.SetFocus
    Else
        chkPaket.Enabled = False
        chkPaket.value = 0
        dcPaket.Text = ""
        cmdSimpan.SetFocus
    End If
        If KeyAscii = 39 Then KeyAscii = 0
        If KeyAscii = 27 Then
        End If
End If
Exit Sub
gabril:
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
End Sub

Private Sub meRTRWPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKodePos.SetFocus
    If KeyCode = 39 Then KeyCode = 0
    Call SetKeyPressToNumber(KeyCode)
End Sub
    
Private Sub meRTRWPJ_KeyPress(KeyCode As Integer)
    If KeyCode = 13 Then txtKodePos.SetFocus
End Sub

Private Sub optTidak_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkDiriSendiri.SetFocus
End Sub

Private Sub optYa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkDiriSendiri.SetFocus
End Sub

Private Sub txtAlamatRI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcPropinsiPJ.SetFocus
End Sub

Private Sub txtKodePos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtTlpRI.SetFocus
    Call SetKeyPressToNumber(KeyCode)
End Sub

Private Sub txtKodePos_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then txtTlpRI.SetFocus
End Sub


Private Sub txtMaks_KeyPress(KeyAscii As Integer)
Call SetKeyPressToNumber(KeyAscii)
If KeyAscii = 13 Then
    cmdSimpan.SetFocus
End If
End Sub

Private Sub txtNamaRI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcHubungan.SetFocus
    Call SetKeyPressToChar(KeyCode)
End Sub

Private Sub txtNoCM_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
''        txtNoBKM = ""
''=============================== Untuk default 6 digit ==================================
'
'        txtNoCM.Text = Format(txtNoCM, "000000")
'        blnSibuk = True
'        Call CariData
'        dtpTglPendaftaran.SetFocus
''        If dcInstalasi.BoundText = "02" Then
''            dcInstalasi.BoundText = "02"
''        ElseIf dcInstalasi.BoundText = "03" Then
''            dcInstalasi.BoundText = "03"
''        ElseIf dcInstalasi.BoundText = "01" Then
''            dcInstalasi.BoundText = "01"
''        End If
'
'        dcKelompokPasien.BoundText = "01"
''        If chkDetailPasien.Enabled = True Then chkDetailPasien.SetFocus
'    End If
'    If KeyAscii = vbKeyBack Then Exit Sub
'    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
'    If KeyAscii = Asc(",") Then Exit Sub
'    If KeyAscii = Asc(".") Then Exit Sub
    
    If KeyAscii = 13 Then
        strSQL = "SELECT NoCM, Title + ' ' + [Nama Lengkap] AS NamaPasien FROM V_CariPasien WHERE ([No. CM] = '" & txtNoCM.Text & "' )"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            Call CariData
            If txtNoCM.Text <> "" Then dcJenisKelas.SetFocus
        Else
            Call subClearData
            txtNamaPasien.Enabled = True
            cboJK.Enabled = True
            txtThn.Enabled = True
            txtBln.Enabled = True
            txtHr.Enabled = True
            chkDetailPasien.Enabled = True
            txtNamaPasien.SetFocus
        End If
    End If
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = 13 Then Exit Sub
    Call SetKeyPressToNumber(KeyAscii)
'    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    If KeyAscii = Asc(",") Then Exit Sub
    If KeyAscii = Asc(".") Then Exit Sub

End Sub
'Ditutup, Dikarenakan Jika Tombol Tutup ditekan tanpa mengisi nocm maka error, jadi ditutup
Private Sub txtNoCM_LostFocus()
    'Call txtNoCM_KeyPress(13)
End Sub

Private Sub txtNoPendaftaran_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'untuk enable/disable button reg
Private Sub subEnableButtonReg(blnStatus As Boolean)
    cmdRujukan.Enabled = blnStatus
    cmdAsuransiP.Enabled = blnStatus
    cmdSimpan.Enabled = Not blnStatus
    dtpTglPendaftaran.Enabled = Not blnStatus
    dcInstalasi.Enabled = Not blnStatus
    dcRuangan.Enabled = Not blnStatus
    dcSubInstalasi.Enabled = Not blnStatus
    dcRujukanRI.Enabled = Not blnStatus
    dcKelompokPasien.Enabled = Not blnStatus
    dcKelas.Enabled = Not blnStatus
    dcJenisKelas.Enabled = Not blnStatus
End Sub
Private Sub sp_KonsulTemp(ByVal adoCommand As ADODB.Command)
 Set adoCommand = New ADODB.Command
 
    MousePointer = vbHourglass
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("NamaPasien", adVarChar, adParamInput, 100, txtNamaPasien.Text)
        If cboJK.Text = "Laki-laki" Then
            strJK = "L"
        Else
            strJK = "P"
        End If
        .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 1, strJK)
        .Parameters.Append .CreateParameter("KdInstalasi", adChar, adParamInput, 2, dcInstalasi.BoundText)
        
        .ActiveConnection = dbConn
        .CommandText = "Add_KonsulTemp"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam Loading Data", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    MousePointer = vbDefault
Exit Sub
End Sub

Private Sub sp_PasienDaftarPaketPelayananTM(ByVal adoCommand As ADODB.Command)
On Error GoTo errLoad
 Set adoCommand = New ADODB.Command
 
    MousePointer = vbHourglass
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("KdPaket", adVarChar, adParamInput, 3, dcPaket.BoundText)
        .Parameters.Append .CreateParameter("QtyMax", adTinyInt, adParamInput, 1, txtMaks.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, dcRuangan.BoundText)
        
        .ActiveConnection = dbConn
        .CommandText = "Add_PasienDaftarPaketPelayananTM"
        .CommandType = adCmdStoredProc
'        .CommandTimeout = 120
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam proses Data", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    MousePointer = vbDefault
Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Store procedure untuk mengisi registrasi pasien RI
Private Sub sp_RegistrasiPasienRI(ByVal adoCommand As ADODB.Command)
 Set adoCommand = New ADODB.Command
 
    MousePointer = vbHourglass
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dcSubInstalasi.BoundText)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, dcKelas.BoundText) ' dcKelasKamarRI.BoundText)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dtpTglPendaftaran.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 2, Null)
        .Parameters.Append .CreateParameter("KdCaraMasuk", adChar, adParamInput, 2, dcCaraMasukRI.BoundText)
        .Parameters.Append .CreateParameter("KdRujukanAsal", adChar, adParamInput, 2, dcRujukanRI.BoundText)
        .Parameters.Append .CreateParameter("NamaPJ", adVarChar, adParamInput, 20, txtNamaRI.Text)
        .Parameters.Append .CreateParameter("PekerjaanPJ", adVarChar, adParamInput, 30, dcPekerjaanPJ.Text)
        .Parameters.Append .CreateParameter("Hubungan", adChar, adParamInput, 2, IIf(dcHubungan.BoundText = "", Null, dcHubungan.BoundText))
        .Parameters.Append .CreateParameter("AlamatPJ", adVarChar, adParamInput, 50, IIf(txtAlamatRI.Text = "", Null, txtAlamatRI.Text))
        .Parameters.Append .CreateParameter("PropinsiPJ", adVarChar, adParamInput, 25, IIf(dcPropinsiPJ.Text = "", Null, dcPropinsiPJ.Text))
        .Parameters.Append .CreateParameter("KotaPJ", adVarChar, adParamInput, 25, IIf(dcKotaPJ.Text = "", Null, dcKotaPJ.Text))
        .Parameters.Append .CreateParameter("KecamatanPJ", adVarChar, adParamInput, 25, IIf(dcKecamatanPJ.Text = "", Null, dcKecamatanPJ.Text))
        .Parameters.Append .CreateParameter("KelurahanPJ", adVarChar, adParamInput, 25, IIf(dcKelurahanPJ.Text = "", Null, dcKelurahanPJ.Text))
        .Parameters.Append .CreateParameter("RTRWPJ", adVarChar, adParamInput, 25, IIf(meRTRWPJ.Text = "", Null, meRTRWPJ.Text))
        .Parameters.Append .CreateParameter("KodePosPJ", adVarChar, adParamInput, 25, IIf(meRTRWPJ.Text = "", Null, txtKodePos.Text))
        .Parameters.Append .CreateParameter("TeleponPJ", adVarChar, adParamInput, 20, IIf(Len(Trim(txtTlpRI.Text)) = 0, Null, Trim(txtTlpRI.Text)))
        
        .ActiveConnection = dbConn
        .CommandText = "Add_RegistrasiPasienRI"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan registrasi RI", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    MousePointer = vbDefault
    Exit Sub
End Sub

'Store procedure untuk mengisi pasien masuk RI
Private Sub sp_PasienMasukKamar(ByVal adoCommand As ADODB.Command)
    Set adoCommand = New ADODB.Command
    
    MousePointer = vbHourglass
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dcSubInstalasi.BoundText)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, dcRuangan.BoundText)
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 2, Null)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, dcKelasKamarRI.BoundText)
        .Parameters.Append .CreateParameter("KdKamar", adChar, adParamInput, 4, dcNoKamarRI.BoundText)
        .Parameters.Append .CreateParameter("NoBed", adChar, adParamInput, 2, dcNoBedRI.BoundText)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dtpTglPendaftaran.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, dcKelas.BoundText)
        
        .Parameters.Append .CreateParameter("OutputNoPakai", adChar, adParamOutput, 10, Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("KdCaraMasuk", adChar, adParamInput, 2, dcCaraMasukRI.BoundText)
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, Null)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 2, IIf(optTidak.value = True, "MA", "RG"))
        
        .ActiveConnection = dbConn
        .CommandText = "Add_PasienMasukKamar"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam proses pasien masuk kamar", vbCritical, "Validasi"
        Else
            txtNoPakai.Text = .Parameters("OutputNoPakai").value
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    MousePointer = vbDefault
    Exit Sub
End Sub

'Store procedure untuk mengisi registrasi pasien
Private Sub sp_RegistrasiAll(ByVal adoCommand As ADODB.Command)
Dim strLokal As String
    Set adoCommand = New ADODB.Command
    
    MousePointer = vbHourglass
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dcSubInstalasi.BoundText)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, dcRuangan.BoundText)
        .Parameters.Append .CreateParameter("TglPendaftaran", adDate, adParamInput, , Format(dtpTglPendaftaran.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dtpTglPendaftaran.value, "yyyy/MM/dd HH:mm:ss"))
        If TempKodeKelas = "" Then
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, "01")
        Else
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, TempKodeKelas)
        End If
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, strKdKelompokPasien)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, noidpegawai)
        .Parameters.Append .CreateParameter("OutputNoPendaftaran", adChar, adParamOutput, 10, Null)
        If txtFormPengirim = "frmDaftarReservasiPasien" Then
            .Parameters.Append .CreateParameter("OutputNoAntrian", adChar, adParamOutput, 3, strnoAntrianPasien)
        Else
            .Parameters.Append .CreateParameter("OutputNoAntrian", adChar, adParamOutput, 3, Null)
        End If
        .Parameters.Append .CreateParameter("KdDetailJenisJasaPelayanan", adChar, adParamInput, 2, dcJenisKelas.BoundText)
'        add Dayz, utk PaketPelayananTM MCU & Rehab Medik
        .Parameters.Append .CreateParameter("KdPaket", adVarChar, adParamInput, 3, IIf(dcPaket.BoundText = "", Null, dcPaket.BoundText))
'        end Dayz
        .Parameters.Append .CreateParameter("KdRujukanAsal", adChar, adParamInput, 2, dcRujukanRI.BoundText)
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, IIf(dcDokter.BoundText = "", Null, dcDokter.BoundText))
'         .Parameters.Append .CreateParameter("KdJenisPoliklinik", adChar, adParamInput, 2, IIf((CboNamaPoli.BoundText = ""), CboNamaPoli.BoundText, CboNamaPoli.BoundText))

        .ActiveConnection = dbConn
        .CommandText = "Add_RegistrasiPasienMRS"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Pendaftaran Pasien..", vbCritical, "Validasi"
        Else
            If Not IsNull(.Parameters("OutputNoPendaftaran").value) Then mstrNoPen = .Parameters("OutputNoPendaftaran").value
            If Not IsNull(.Parameters("OutputNoAntrian").value) Then strNoAntrian = .Parameters("OutputNoAntrian").value
            txtKdAntrian.Text = strNoAntrian
            txtNoPendaftaran.Text = mstrNoPen
            If Len(mstrNoPen) = 0 Then
                strLokal = "SELECT NoPendaftaran, NoAntrian from PasienMasukRumahSakit where kdRuangan = '" & dcRuangan.BoundText & "' and tglMasuk = '" & Format(dtpTglPendaftaran.value, "yyyy/MM/dd HH:mm:ss") & "' and NoCM = '" & Trim(txtNoCM.Text) & "' and idUser = '" & noidpegawai & "'"
                Call msubRecFO(rs, strLokal)
                mstrNoPen = rs("NoPendaftaran").value
                strNoAntrian = rs("NoAntrian").value
            End If
'            If txtFormPengirim.Text = "frmDaftarReservasiPasien" Then
'                strLokal = "Select NoAntrian From ReservasiPasien where KdRuangan='" & dcRuangan.BoundText & "' and year(tglMasuk) = '" & Format(dtpTglPendaftaran.value, "yyyy") & "' and month(tglMasuk) = '" & Format(dtpTglPendaftaran.value, "MM") & "' and day(tglMasuk) = '" & Format(dtpTglPendaftaran.value, "dd") & "' and NoCM = '" & Trim(txtNoCM.Text) & "' and Noreservasi = '" & frmDaftarReservasiPasien.dgDaftarReservasiPasien.Columns("NoReservasi") & "' and StatusDaftar='T'"
'                Call msubRecFO(rs, strLokal)
'                dbConn.Execute "Update PasienMasukRumahSakit Set NoAntrian='" & rs(0).value & "' Where NoPendaftaran = '" & mstrNoPen & "' and NoCM='" & txtNoCM.Text & "'"
'                strNoAntrian = rs(0).value
'                txtKdAntrian.Text = strNoAntrian
'            End If
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    MousePointer = vbDefault
    Exit Sub
End Sub

'Store procedure untuk mengisi pelayanan otomatis
Private Function sp_PelayananOtomatis() As Boolean
On Error GoTo errLoad
    sp_PelayananOtomatis = True
    Set dbcmd = New ADODB.Command
    
    MousePointer = vbHourglass
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dcSubInstalasi.BoundText)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, dcRuangan.BoundText)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dtpTglPendaftaran.value, "yyyy/MM/dd HH:mm:ss"))
        If dcInstalasi.BoundText <> "03" And dcInstalasi.BoundText <> "08" Then
            .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, Null)
        Else
            .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, dcKelasKamarRI.BoundText)
        End If
        
        If TempKodeKelas = "" Then
        .Parameters.Append .CreateParameter("KdKelasPel", adChar, adParamInput, 2, "01")
        Else
        .Parameters.Append .CreateParameter("KdKelasPel", adChar, adParamInput, 2, TempKodeKelas)
        End If
        
        If dcInstalasi.BoundText <> "03" And dcInstalasi.BoundText <> "08" Then
            .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, Null)
        Else
            .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, txtNoPakai.Text)
        End If
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strIDPegawaiAktif)
        If dcInstalasi.BoundText <> "03" And dcInstalasi.BoundText <> "08" Then
            .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 2, "AL")
        Else
            .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 2, IIf(optTidak.value = True, "MA", "RG"))
        End If
                
        .ActiveConnection = dbConn
        .CommandText = "Add_BiayaPelayananOtomatisNew"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            sp_PelayananOtomatis = False
            MsgBox "Ada kesalahan proses penyimpanan data biaya otomatis", vbCritical, "Validasi"
            GoTo errLoad
        Else
            Call Add_HistoryLoginActivity("Add_BiayaPelayananOtomatisNew")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    MousePointer = vbDefault
    Exit Function
Exit Function
errLoad:
    sp_DelBiayaPelayananCek Trim(txtNoPendaftaran.Text)
    Exit Function
    Call msubPesanError("sp_PelayananOtomatis")
End Function

'Store procedure untuk menghapus biaya pelayanan pasien yang gagal disimpan
Private Sub sp_DelBiayaPelayananCek(varNoPendaftaran As String)
Dim adoCek As ADODB.Command
    Set adoCek = New ADODB.Command
    With adoCek
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.CEK_BiayaPelayananOTO"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Penghapusan Biaya Pelayanan Pasien", vbCritical, "Validasi"
        Else
'            MsgBox "Pemasukan Biaya Pelayanan Pasien sukses", vbExclamation, "Validasi"
            Call Add_HistoryLoginActivity("Delete_BiayaPelayanan")
        End If
        Call deleteADOCommandParameters(adoCek)
        Set adoCek = Nothing
    End With
    Exit Sub
End Sub

'Store procedure untuk mengisi asuransi pasien
Private Sub sp_AsuransiPasien(ByVal adoCommand As ADODB.Command)
Dim xrtSQL As String
    Set dbcmd = New ADODB.Command
    
    MousePointer = vbHourglass
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, typAsuransi.strIdPenjamin)
        .Parameters.Append .CreateParameter("IdAsuransi", adVarChar, adParamInput, 25, typAsuransi.strIdAsuransi)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, mstrNoCM)
        .Parameters.Append .CreateParameter("KdHubKeluarga", adChar, adParamInput, 2, typAsuransi.strHubungan)
        .Parameters.Append .CreateParameter("NamaPeserta", adVarChar, adParamInput, 50, typAsuransi.strNamaPeserta)
        
        .Parameters.Append .CreateParameter("IDPeserta", adVarChar, adParamInput, 16, typAsuransi.strIdPeserta)
        .Parameters.Append .CreateParameter("KdGolongan", adChar, adParamInput, 2, IIf(Len(Trim(typAsuransi.strKdGolongan)) = 0, Null, Trim(typAsuransi.strKdGolongan)))
        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(typAsuransi.dTglLahir, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("Alamat", adVarChar, adParamInput, 100, typAsuransi.strAlamat)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        
'        .Parameters.Append .CreateParameter("KdHubungan", adChar, adParamInput, 2, typAsuransi.strHubungan)
        If typAsuransi.strNoSJP <> "" Then
            .Parameters.Append .CreateParameter("NoSJP", adVarChar, adParamInput, 30, typAsuransi.strNoSJP)
        Else
            .Parameters.Append .CreateParameter("NoSJP", adVarChar, adParamInput, 30, Null)
        End If
        .Parameters.Append .CreateParameter("TglSJP", adDate, adParamInput, , Format(typAsuransi.dTglSJP, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("NoBP", adVarChar, adParamInput, 3, IIf(Len(Trim(typAsuransi.strNoBp)) = 0, Null, Trim(typAsuransi.strNoBp)))
        
        .Parameters.Append .CreateParameter("KunjunganKe", adInteger, adParamInput, , typAsuransi.intNoKunjungan)
        .Parameters.Append .CreateParameter("OutputNoSJP", adVarChar, adParamOutput, 30, Null)
        .Parameters.Append .CreateParameter("StatusNoSJP", adChar, adParamInput, 1, typAsuransi.strStatusNoSJP)
        .Parameters.Append .CreateParameter("AnakKe", adInteger, adParamInput, , typAsuransi.intAnakKe)
        .Parameters.Append .CreateParameter("UnitBagian", adVarChar, adParamInput, 50, IIf(Len(Trim(typAsuransi.strUnitBagian)) = 0, Null, Trim(typAsuransi.strUnitBagian)))
        
        .Parameters.Append .CreateParameter("KdPaket", adVarChar, adParamInput, 3, Null)
        .Parameters.Append .CreateParameter("NoRujukan", adVarChar, adParamInput, 30, typAsuransi.strNoRujukan)
        .Parameters.Append .CreateParameter("KdRujukanAsal", adChar, adParamInput, 2, typAsuransi.strKdRujukanAsal)
        .Parameters.Append .CreateParameter("DetailRujukanAsal", adVarChar, adParamInput, 100, typAsuransi.strDetailRujukanAsal)
        .Parameters.Append .CreateParameter("KdDetailRujukanAsal", adChar, adParamInput, 8, typAsuransi.strKdDetailRujukanAsal)
        
        .Parameters.Append .CreateParameter("NamaPerujuk", adVarChar, adParamInput, 50, typAsuransi.strNamaPerujuk)
        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , typAsuransi.dTglDirujuk)
        .Parameters.Append .CreateParameter("DiagnosaRujukan", adVarChar, adParamInput, 100, typAsuransi.strDiagnosaRujukan)
        .Parameters.Append .CreateParameter("KdDiagnosa", adVarChar, adParamInput, 7, typAsuransi.strKdDiagnosa)
        
        '###24-4-2008 by john ----'edit splakuk
        xrtSQL = "SELECT  KdinstitusiAsal, InstitusiAsal FROM InstitusiAsalPasien WHERE InstitusiAsal LIKE '" & typAsuransi.strPerusahaanPenjamin & "%' or KdInstitusiAsal LIKE '" & typAsuransi.strPerusahaanPenjamin & "' and StatusEnabled='1'"
        Call msubRecFO(rsx, xrtSQL)
        .Parameters.Append .CreateParameter("KdInstitusiAsal", adVarChar, adParamInput, 4, IIf(Len(Trim(rsx(0).value)) = 0, Null, Trim(rsx(0).value)))
        .Parameters.Append .CreateParameter("KdKelasDiTanggung", adChar, adParamInput, 2, typAsuransi.strKdKelasDitanggung)
        
        .ActiveConnection = dbConn
        .CommandText = "AU_AsuransiPasienJoinProgramAskes"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan Asuransi Pasien", vbCritical, "Validasi"
            mstrNoSJP = typAsuransi.strNoSJP
        Else
            mstrNoSJP = typAsuransi.strNoSJP
            Call Add_HistoryLoginActivity("AU_AsuransiPasienJoinProgramAskes")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    MousePointer = vbDefault
    Exit Sub
End Sub

'untuk cek validasi
Private Function funcCekValidasi() As Boolean
    If txtNamaPasien.Text = "" Then
        MsgBox "Nama Pasien Harus Diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        txtNoCM.SetFocus
        Exit Function
    End If
    If Periksa("datacombo", dcJenisKelas, "Jenis kelas pelayanan kosong") = False Then funcCekValidasi = False: Exit Function
    If Periksa("datacombo", dcRuangan, "Nama ruangan kosong") = False Then funcCekValidasi = False: Exit Function
    If Periksa("datacombo", dcRujukanRI, "Data rujukan kosong!") = False Then funcCekValidasi = False: Exit Function
    If Periksa("datacombo", dcKelompokPasien, "Jenis pasien kosong!") = False Then funcCekValidasi = False: Exit Function
    funcCekValidasi = True
End Function

'untuk membersihkan data pasien registrasi
Private Sub subClearData()
    txtNoPakai.Text = ""
    txtNoPendaftaran.Text = ""
    txtNamaPasien.Text = ""
    cboJK.Text = ""
    txtThn.Text = ""
    txtBln.Text = ""
    txtHr.Text = ""
    dcHubungan.BoundText = ""
    dtpTglPendaftaran.MaxDate = #9/9/2999#
    dtpTglPendaftaran.value = Now
'    If dcInstalasi.Text = "" Then
    dcInstalasi.Text = ""
'    Else
'    End If
    dcRuangan.Text = ""
    dcJenisKelas.Text = ""
    dcKelompokPasien.Text = ""
    dcKelas.Text = ""
    dcDokter.Text = ""
    
'    TempKodeKelas = ""
End Sub

Private Sub subDcSource()
On Error GoTo errLoad
    'Call msubDcSource(dcKelompokPasien, rs, "SELECT KdKelompokPasien,JenisPasien FROM KelompokPasien WHERE NOT (KdKelompokPasien = '05') order by JenisPasien") 'askes gakin di tutup by splakuk
    Call msubDcSource(dcKelompokPasien, rs, "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where StatusEnabled='1' order by JenisPasien") 'askes gakin di buka
    'Call msubDcSource(dcKelompokPasien, rs, "SELECT IdPenjamin, NamaPenjamin FROM Penjamin where StatusEnabled='1'") 'order by NamaPenjamin")
    If dcInstalasi.BoundText = "09" Or dcInstalasi.BoundText = "04" Or dcInstalasi.BoundText = "10" Or dcInstalasi.BoundText = "16" Then
        Call msubDcSource(dcRujukanRI, rs, "SELECT KdRujukanAsal, RujukanAsal FROM RujukanAsal where StatusEnabled='1' and KdRujukanAsal<>'01'")
        If Not rs.EOF Then
         '   dcRujukanRI.BoundText = rs(0).value
        End If
    Else
        Call msubDcSource(dcRujukanRI, rs, "SELECT KdRujukanAsal, RujukanAsal FROM RujukanAsal where StatusEnabled='1'")
'        If Not rs.EOF Then
'        '    dcRujukanRI.BoundText = rs(0).value
'        End If
    End If
    Call msubDcSource(dcCaraMasukRI, rs, "SELECT KdCaraMasuk, CaraMasuk FROM CaraMasuk where StatusEnabled='1'")
    Call msubDcSource(dcHubungan, rs, "SELECT Hubungan, NamaHubungan FROM HubunganKeluarga where StatusEnabled='1'")
    
    strSQL = "SELECT DISTINCT KdPropinsi, NamaPropinsi FROM Propinsi where StatusEnabled='1' order by NamaPropinsi"
    Call msubDcSource(dcPropinsiPJ, rs, strSQL)

    strSQL = "SELECT DISTINCT KdKotaKabupaten, NamaKotaKabupaten FROM KotaKabupaten where KdPropinsi = '" & dcPropinsiPJ.BoundText & "' and StatusEnabled='1' order by NamaKotaKabupaten"
    Call msubDcSource(dcKotaPJ, rs, strSQL)

    strSQL = "SELECT DISTINCT KdKecamatan, NamaKecamatan FROM Kecamatan where KdKotaKabupaten = '" & dcKotaPJ.BoundText & "' and StatusEnabled='1' order by NamaKecamatan"
    Call msubDcSource(dcKecamatanPJ, rs, strSQL)

    strSQL = "SELECT DISTINCT KdKelurahan, NamaKelurahan FROM Kelurahan where KdKecamatan = '" & dcKecamatanPJ.BoundText & "' and StatusEnabled='1' order by NamaKelurahan"
    Call msubDcSource(dcKelurahanPJ, rs, strSQL)
    
    strSQL = "SELECT DISTINCT KdPekerjaan,Pekerjaan FROM Pekerjaan where StatusEnabled='1'"
    Call msubDcSource(dcPekerjaanPJ, rs, strSQL)

   Call msubDcSource(dcPaket, rs, "SELECT KdPaket,NamaPaket FROM PaketPelayananTM where StatusEnabled='1' order by NamaPaket")
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtNoRujukan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dcRujukanRI.SetFocus
    
    If KeyAscii = 13 Then
        Call CekNoRujukan
    End If
End Sub

Private Sub txtThn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtBln.SetFocus
End Sub
Private Sub txtBln_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtHr.SetFocus
End Sub
Private Sub txtHr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJenisKelas.SetFocus
End Sub
Private Sub txtTlpRI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSimpan.SetFocus
    Call SetKeyPressToNumber(KeyCode)
End Sub
'untuk mengganti nocm on change
Public Sub CariData()
On Error GoTo errLoad
    Call subClearData
    Call subEnableButtonReg(False)
    
    'cek pasien igd
    strSQL = "SELECT NoCM FROM PasienMeninggal WHERE (NoCM = '" & txtNoCM.Text & "')"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        MsgBox "Pasien tersebut Sudah meninggal", vbInformation, "Informasi"
        mstrNoCM = ""
        chkDetailPasien.Enabled = False
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
        chkDetailPasien.Enabled = False
        cmdSimpan.Enabled = False
        Exit Sub
    End If
    
    'cek pasien ri
    strSQL = "SELECT dbo.RegistrasiRI.NoCM, dbo.Ruangan.NamaRuangan FROM dbo.RegistrasiRI INNER JOIN dbo.Ruangan ON dbo.RegistrasiRI.KdRuangan = dbo.Ruangan.KdRuangan WHERE (NoCM = '" & txtNoCM.Text & "') AND StatusPulang = 'T'"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        MsgBox "Pasien tersebut belum keluar dari Rawat Inap," & vbNewLine & "Ruangan " & rs("NamaRuangan") & " ", vbInformation, "Informasi"
        mstrNoCM = ""
        chkDetailPasien.Enabled = False
        cmdSimpan.Enabled = False
        Exit Sub
    End If
    
    strSQL = "Select * from v_CariPasien WHERE [No. CM]='" & txtNoCM.Text & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        mstrNoCM = ""
        chkDetailPasien.Enabled = False
        cmdSimpan.Enabled = False
        Exit Sub
    End If
    
    ' Cek Data Asuransi Pasien
    strSQL = "Select Top(1) * from V_DataPesertaAsuransi where NoCM = '" & txtNoCM.Text & "'"
    Call msubRecFO(rs1, strSQL)
    
    If rs1.EOF = False Then
        
        dcKelompokPasien.BoundText = rs1.Fields("KdKelompokPasien")
        dcKelompokPasien.Text = rs1.Fields("JenisPasien")
    End If
    ''

    
    mstrNoCM = txtNoCM.Text
    txtNamaPasien.Text = rs.Fields("Nama Lengkap").value
    If rs.Fields("JK").value = "P" Then
        cboJK.Text = "Perempuan"
    ElseIf rs.Fields("JK").value = "L" Then
        cboJK.Text = "Laki-laki"
    End If
    txtThn.Text = rs.Fields("UmurTahun").value
    txtBln.Text = rs.Fields("UmurBulan").value
    txtHr.Text = rs.Fields("UmurHari").value
    Set rs = Nothing
    chkDetailPasien.Enabled = True
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


Private Function sp_TempPrintKartu(Status As String, f_NoPendaftaran As String) As Boolean
On Error GoTo onError
    sp_TempPrintKartu = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("TglPrintKartu", adDate, adParamInput, , Format(Now, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, Status)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_TempPrintKartu"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_TempPrintKartu = False
        Else
            Call Add_HistoryLoginActivity("AUD_TempPrintKartu")
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    
Exit Function
onError:
        Call msubPesanError
End Function

Private Sub CekNoRujukan()
    If (Dir("C:\SDK\Vclaim\result.tlb") <> "") Then
        Dim context As ContextVclaim
        Set context = New ContextVclaim
        Dim result() As String
        
        strSQL = "Select Value From SettingGlobal where Prefix In('ConsumerID','PasswordKey')"
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = False Then
            context.ConsumerID = rs(0).value
            rs.MoveNext
            context.PasswordKey = rs(0).value
        End If
        
        strSQL = "SELECT Value FROM SettingGlobal where Prefix='UrlGenerateSEP'"
        Call msubRecFO(rs, strSQL)
        Dim URL  As String
        
        If rs.EOF = False Then
            URL = rs.Fields(0)
            context.URL = URL
        End If
        
        result = context.RujukanPcareByNoRujukan(txtNoRujukan.Text)
        
        Dim i As Long
        For i = LBound(result) To UBound(result)
            Dim arr() As String
            Dim NoCMRujukan As String
            Dim KdPoli As String
            
            KdPoli = ""
            arr = Split(result(i), ":")
            
            Select Case arr(0)
                Case "DIAGNOSA-NOKUNJUNGAN"
                    If arr(1) = txtNoRujukan.Text Then
                        Call dcJenisKelas_GotFocus
                        
                        dcJenisKelas.BoundText = "01"
                        dcJenisKelas.Text = "STANDAR"
                        
                        dcRuangan.Text = ""
                        dcRuangan.BoundText = ""
                        Call dcRuangan_GotFocus
                        
                        dcKelompokPasien.BoundText = "10"
                        dcKelompokPasien.Text = "BPJS"
                    End If
                Case "PESERTA-NOMR"
                    If arr(1) <> mstrNoCM Then
                        MsgBox "NoCM Rujukan (" & arr(1) & ") tidak sesuai.", vbCritical, "Peringatan"
                        Exit Sub
                    End If
                Case "POLIRUJUKAN-KODE"
                    If arr(1) = "ANA" Then
                        dcRuangan.BoundText = "204"
                        dcRuangan.Text = "Poliklinik Anak"
                    ElseIf arr(1) = "BED" Then
                        dcRuangan.BoundText = "205"
                        dcRuangan.Text = "Poliklinik Bedah"
                    ElseIf arr(1) = "INT" Then
                        dcRuangan.BoundText = "202"
                        dcRuangan.Text = "Poliklinik Dalam"
                    ElseIf arr(1) = "GIG" Then
                        dcRuangan.BoundText = "211"
                        dcRuangan.Text = "Poliklinik Gigi"
                    ElseIf arr(1) = "JAN" Then
                        dcRuangan.BoundText = "224"
                        dcRuangan.Text = "Poliklinik Jantung"
                    ElseIf arr(1) = "OBG" Then
                        dcRuangan.BoundText = "206"
                        dcRuangan.Text = "Poliklinik Kebidanan dan Kandungan"
                    ElseIf arr(1) = "KLT" Then
                        dcRuangan.BoundText = "209"
                        dcRuangan.Text = "Poliklinik Kulit dan Kelamin"
                    ElseIf arr(1) = "MAT" Then
                        dcRuangan.BoundText = "201"
                        dcRuangan.Text = "Poliklinik Mata"
                    ElseIf arr(1) = "SAR" Then
                        dcRuangan.BoundText = "221"
                        dcRuangan.Text = "Poliklinik Neurologi"
                    ElseIf arr(1) = "PAR" Then
                        dcRuangan.BoundText = "212"
                        dcRuangan.Text = "Poliklinik Paru"
                    ElseIf arr(1) = "THT" Then
                        dcRuangan.BoundText = "208"
                        dcRuangan.Text = "Poliklinik THT"
                    ElseIf arr(1) = "FIS" Then
                        dcRuangan.BoundText = "601"
                        dcRuangan.Text = "Fisioterapi"
                    End If
            End Select
        Next i
    End If
End Sub
