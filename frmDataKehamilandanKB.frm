VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDataKehamilandanKB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Kehamilan dan Keluarga Berencana Pasien"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDataKehamilandanKB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   9480
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   2160
      TabIndex        =   30
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   3960
      TabIndex        =   32
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   5760
      TabIndex        =   31
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   7605
      TabIndex        =   33
      Top             =   8400
      Width           =   1815
   End
   Begin TabDlg.SSTab ssData 
      Height          =   6135
      Left            =   0
      TabIndex        =   43
      Top             =   2160
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Kehamilan"
      TabPicture(0)   =   "frmDataKehamilandanKB.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraDokterKehamilan"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Keluarga Berencana"
      TabPicture(1)   =   "frmDataKehamilandanKB.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraDataKehamilan"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraDokterKB"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fraDokterKB 
         Caption         =   "Data Pemeriksa"
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
         Left            =   240
         TabIndex        =   67
         Top             =   4920
         Visible         =   0   'False
         Width           =   9015
         Begin MSDataGridLib.DataGrid dgDokterKB 
            Height          =   2055
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   8535
            _ExtentX        =   15055
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
      Begin VB.Frame fraDokterKehamilan 
         Caption         =   "Data Pemeriksa"
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
         Left            =   -74760
         TabIndex        =   66
         Top             =   4920
         Visible         =   0   'False
         Width           =   9015
         Begin MSDataGridLib.DataGrid dgDokterKehamilan 
            Height          =   2055
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   8535
            _ExtentX        =   15055
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
      Begin VB.Frame fraDataKehamilan 
         Caption         =   "Data Keluarga Berencana (KB)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   240
         TabIndex        =   54
         Top             =   480
         Width           =   9015
         Begin MSDataGridLib.DataGrid dgKB 
            Height          =   1935
            Left            =   240
            TabIndex        =   29
            Top             =   3480
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   3413
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtKdDokterKB 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2760
            MaxLength       =   150
            TabIndex        =   69
            Top             =   0
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txtPemeriksaKB 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2280
            MaxLength       =   150
            TabIndex        =   21
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox txtKeteranganKB 
            Appearance      =   0  'Flat
            Height          =   675
            Left            =   240
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   2760
            Width           =   8535
         End
         Begin VB.TextBox txtTindakan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   27
            Top             =   2160
            Width           =   6855
         End
         Begin VB.TextBox txtKegagalan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   240
            MaxLength       =   3
            TabIndex        =   26
            Top             =   2160
            Width           =   1575
         End
         Begin VB.TextBox txtEfekSamping 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   25
            Top             =   1560
            Width           =   6855
         End
         Begin MSDataListLib.DataCombo dcJenisKontrasepsi 
            Height          =   330
            Left            =   240
            TabIndex        =   24
            Top             =   1560
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker dtpTglPeriksaKB 
            Height          =   330
            Left            =   240
            TabIndex        =   20
            Top             =   600
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy HH:mm"
            Format          =   118947843
            UpDown          =   -1  'True
            CurrentDate     =   38076
         End
         Begin MSDataListLib.DataCombo dcPerawatKB 
            Height          =   330
            Left            =   6120
            TabIndex        =   23
            Top             =   600
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   240
            X2              =   8760
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Tanggal Periksa"
            Height          =   210
            Index           =   26
            Left            =   240
            TabIndex        =   65
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Dokter/Perawat Pemeriksa"
            Height          =   210
            Index           =   25
            Left            =   2280
            TabIndex        =   64
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Paramedis Pemeriksa"
            Height          =   210
            Index           =   9
            Left            =   6120
            TabIndex        =   63
            Top             =   360
            Width           =   1680
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Keterangan"
            Height          =   210
            Index           =   15
            Left            =   240
            TabIndex        =   59
            Top             =   2520
            Width           =   945
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Kontrasepsi"
            Height          =   210
            Index           =   14
            Left            =   240
            TabIndex        =   58
            Top             =   1320
            Width           =   1380
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Tindakan"
            Height          =   210
            Index           =   13
            Left            =   1920
            TabIndex        =   57
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Kegagalan ( kali )"
            Height          =   210
            Index           =   12
            Left            =   240
            TabIndex        =   56
            Top             =   1920
            Width           =   1395
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Efek Samping"
            Height          =   210
            Index           =   11
            Left            =   1920
            TabIndex        =   55
            Top             =   1320
            Width           =   1110
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Kehamilan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   -74760
         TabIndex        =   44
         Top             =   480
         Width           =   9015
         Begin MSDataGridLib.DataGrid dgKehamilan 
            Height          =   1935
            Left            =   240
            TabIndex        =   19
            Top             =   3480
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   3413
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtKdDokterKehamilan 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2280
            MaxLength       =   150
            TabIndex        =   68
            Top             =   0
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txtPemeriksaKehamilan 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2280
            MaxLength       =   150
            TabIndex        =   8
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox txtTahunKehamilan 
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
            Left            =   240
            MaxLength       =   1
            TabIndex        =   11
            Top             =   1560
            Width           =   375
         End
         Begin VB.TextBox txtBulanKehamilan 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   1020
            MaxLength       =   2
            TabIndex        =   12
            Top             =   1560
            Width           =   375
         End
         Begin VB.TextBox txtHariKehamilan 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   13
            Top             =   1560
            Width           =   375
         End
         Begin VB.TextBox txtGravite 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2520
            MaxLength       =   3
            TabIndex        =   14
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtPartus 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3840
            MaxLength       =   3
            TabIndex        =   15
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtAbortus 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5160
            MaxLength       =   3
            TabIndex        =   16
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtKeteranganKehamilan 
            Appearance      =   0  'Flat
            Height          =   1155
            Left            =   240
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   2280
            Width           =   8535
         End
         Begin MSDataListLib.DataCombo dcImunisasi 
            Height          =   330
            Left            =   6480
            TabIndex        =   17
            Top             =   1560
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker dtpTglPeriksaKehamilan 
            Height          =   330
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy HH:mm"
            Format          =   119013379
            UpDown          =   -1  'True
            CurrentDate     =   38076
         End
         Begin MSDataListLib.DataCombo dcPerawatKehamilan 
            Height          =   330
            Left            =   6120
            TabIndex        =   10
            Top             =   600
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin MSMask.MaskEdBox meTglLahir 
            Height          =   390
            Left            =   240
            TabIndex        =   70
            Top             =   960
            Visible         =   0   'False
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
         Begin VB.Line Line1 
            Index           =   1
            X1              =   240
            X2              =   8760
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Tanggal Periksa"
            Height          =   210
            Index           =   10
            Left            =   240
            TabIndex        =   62
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Dokter/Perawat Pemeriksa"
            Height          =   210
            Index           =   8
            Left            =   2280
            TabIndex        =   61
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Paramedis Pemeriksa"
            Height          =   210
            Index           =   7
            Left            =   6120
            TabIndex        =   60
            Top             =   360
            Width           =   1680
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Umur Kehamilan"
            Height          =   210
            Index           =   24
            Left            =   240
            TabIndex        =   53
            Top             =   1320
            Width           =   1305
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Index           =   23
            Left            =   675
            TabIndex        =   52
            Top             =   1590
            Width           =   285
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Index           =   22
            Left            =   1470
            TabIndex        =   51
            Top             =   1590
            Width           =   240
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Index           =   21
            Left            =   2250
            TabIndex        =   50
            Top             =   1590
            Width           =   165
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Gravite"
            Height          =   210
            Index           =   20
            Left            =   2520
            TabIndex        =   49
            Top             =   1320
            Width           =   570
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Partus"
            Height          =   210
            Index           =   19
            Left            =   3840
            TabIndex        =   48
            Top             =   1320
            Width           =   510
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Abortus"
            Height          =   210
            Index           =   18
            Left            =   5160
            TabIndex        =   47
            Top             =   1320
            Width           =   645
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Imunisasi"
            Height          =   210
            Index           =   17
            Left            =   6480
            TabIndex        =   46
            Top             =   1320
            Width           =   720
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Keterangan"
            Height          =   210
            Index           =   16
            Left            =   240
            TabIndex        =   45
            Top             =   2040
            Width           =   945
         End
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
      TabIndex        =   34
      Top             =   960
      Width           =   9495
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5520
         MaxLength       =   9
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
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
         Left            =   6840
         TabIndex        =   35
         Top             =   360
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
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Index           =   4
            Left            =   550
            TabIndex        =   38
            Top             =   277
            Width           =   285
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Index           =   5
            Left            =   1350
            TabIndex        =   37
            Top             =   277
            Width           =   240
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Index           =   6
            Left            =   2130
            TabIndex        =   36
            Top             =   270
            Width           =   165
         End
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Index           =   3
         Left            =   5520
         TabIndex        =   42
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Index           =   2
         Left            =   3000
         TabIndex        =   41
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Index           =   1
         Left            =   1800
         TabIndex        =   40
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   1335
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   71
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
      Left            =   7680
      Picture         =   "frmDataKehamilandanKB.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDataKehamilandanKB.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataKehamilandanKB.frx":444B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDataKehamilandanKB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim subBolTampil As Boolean

Private Function sp_DelCatatanKehamilanPasien() As Boolean
    On Error GoTo errLoad

    sp_DelCatatanKehamilanPasien = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, dgKehamilan.Columns("NoPendaftaran"))
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dgKehamilan.Columns("TglPeriksa"), "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuanganPasien)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Del_CatatanKehamilanPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_DelCatatanKehamilanPasien = False
        End If
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    Call msubPesanError("sp_DelCatatanKehamilanPasien")
End Function

Private Function sp_DelCatatanProgramKBPasien() As Boolean
    On Error GoTo errLoad

    sp_DelCatatanProgramKBPasien = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, dgKB.Columns("NoPendaftaran"))
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dgKB.Columns("TglPeriksa"), "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuanganPasien)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Del_CatatanProgramKBPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_DelCatatanProgramKBPasien = False
        End If
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    Call msubPesanError("sp_DelCatatanProgramKBPasien")
End Function

Private Sub subLoadDataGrid()
    On Error GoTo errLoad
    Dim i As Integer

    Select Case ssData.Tab
        Case 0
            strSQL = "SELECT * FROM V_DataKehamilan WHERE NoCM = '" & txtNoCM.Text & "'"
            Call msubRecFO(rs, strSQL)
            Set dgKehamilan.DataSource = rs
            With dgKehamilan
                For i = 0 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next i
                .Columns("TglPeriksa").Width = 1559
                .Columns("NamaRuangan").Width = 1559
                .Columns("TglKehamilan").Width = 1200
                .Columns("Gravite").Width = 800
                .Columns("Partus").Width = 800
                .Columns("Abortus").Width = 800
                .Columns("NamaImunisasi").Width = 1559
                .Columns("Keterangan").Width = 1559
            End With

        Case 1
            strSQL = "SELECT * FROM V_DataKB WHERE NoCM = '" & txtNoCM.Text & "'"
            Call msubRecFO(rs, strSQL)
            Set dgKB.DataSource = rs
            With dgKB
                For i = 0 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next i
                .Columns("TglPeriksa").Width = 1200
                .Columns("NamaRuangan").Width = 1559
                .Columns("JenisKontrasepsi").Width = 1200
                .Columns("EfekSamping").Width = 800
                .Columns("Kegagalan").Width = 800
                .Columns("Tindakan").Width = 800
                .Columns("Pemeriksa").Width = 1559
                .Columns("Keterangan").Width = 1559
            End With
    End Select

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    'data kehamilan
    strSQL = "SELECT KdImunisasi, NamaImunisasi From Imunisasi where StatusEnabled='1' ORDER BY NamaImunisasi"
    Call msubDcSource(dcImunisasi, rs, strSQL)
    If rs.EOF = False Then dcImunisasi.BoundText = rs(0).value

    strSQL = "SELECT  IdPegawai, [Nama Pemeriksa] From V_DaftarPemeriksaPasien ORDER BY  [Nama Pemeriksa]"
    Call msubDcSource(dcPerawatKehamilan, rs, strSQL)
    If rs.EOF = False Then dcPerawatKehamilan.BoundText = strIDPegawaiAktif

    'data keluarga berencana
    strSQL = "SELECT KdJenisKontrasepsi, JenisKontrasepsi From JenisKontrasepsi where StatusEnabled='1' ORDER BY JenisKontrasepsi"
    Call msubDcSource(dcJenisKontrasepsi, rs, strSQL)
    If rs.EOF = False Then dcJenisKontrasepsi.BoundText = rs(0).value

    strSQL = "SELECT  IdPegawai, [Nama Pemeriksa] From V_DaftarPemeriksaPasien ORDER BY  [Nama Pemeriksa]"
    Call msubDcSource(dcPerawatKB, rs, strSQL)
    If rs.EOF = False Then dcPerawatKB.BoundText = strIDPegawaiAktif

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subKosong()
    Select Case ssData.Tab
        Case 0
            dtpTglPeriksaKehamilan.value = Now
            txtKdDokterKehamilan.Text = ""
            txtPemeriksaKehamilan.Text = ""
            dcPerawatKehamilan.BoundText = ""

            txtTahunKehamilan.Text = 0
            txtBulanKehamilan.Text = 0
            txtHariKehamilan.Text = 0
            txtGravite.Text = ""
            txtPartus.Text = ""
            txtAbortus.Text = ""
            dcImunisasi.BoundText = ""
            txtKeteranganKehamilan.Text = ""
        Case 1
            dtpTglPeriksaKB.value = Now
            txtKdDokterKB.Text = ""
            txtPemeriksaKB.Text = ""
            dcPerawatKB.BoundText = ""

            dcJenisKontrasepsi.BoundText = ""
            txtEfekSamping.Text = ""
            txtKegagalan.Text = ""
            txtTindakan.Text = ""
            txtKeteranganKB.Text = ""
    End Select
End Sub

Private Function sp_CatatanKehamilanPasien() As Boolean
    On Error GoTo errLoad

    sp_CatatanKehamilanPasien = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dtpTglPeriksaKehamilan.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuanganPasien)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("IdDokter", adVarChar, adParamInput, 10, txtKdDokterKehamilan.Text)
        .Parameters.Append .CreateParameter("TglKehamilan", adDate, adParamInput, , Format(meTglLahir.Text, "yyyy/MM/dd"))

        .Parameters.Append .CreateParameter("Gravite", adTinyInt, adParamInput, , Val(txtGravite.Text))
        .Parameters.Append .CreateParameter("Partus", adTinyInt, adParamInput, , Val(txtPartus.Text))
        .Parameters.Append .CreateParameter("Abortus", adTinyInt, adParamInput, , Val(txtAbortus.Text))
        .Parameters.Append .CreateParameter("KdImunisasi", adChar, adParamInput, 3, dcImunisasi.BoundText)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(Len(Trim(txtKeteranganKehamilan.Text)) = 0, Null, Trim(txtKeteranganKehamilan.Text)))

        .Parameters.Append .CreateParameter("IdParamedis", adChar, adParamInput, 10, IIf(dcPerawatKehamilan.BoundText = "", Null, dcPerawatKehamilan.BoundText))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AU_CatatanKehamilanPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_CatatanKehamilanPasien = False
        Else
            Call Add_HistoryLoginActivity("AU_CatatanKehamilanPasien")
        End If
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    sp_CatatanKehamilanPasien = False
    Call msubPesanError("sp_CatatanKehamilanPasien")
End Function

Private Function sp_CatatanProgramKBPasien() As Boolean
    On Error GoTo errLoad

    sp_CatatanProgramKBPasien = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dtpTglPeriksaKB.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuanganPasien)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("IdDokter", adVarChar, adParamInput, 10, txtKdDokterKB.Text)

        .Parameters.Append .CreateParameter("KdJenisKontrasepsi", adChar, adParamInput, 2, dcJenisKontrasepsi.BoundText)
        .Parameters.Append .CreateParameter("EfekSamping", adVarChar, adParamInput, 200, Trim(txtEfekSamping.Text))
        .Parameters.Append .CreateParameter("Kegagalan", adTinyInt, adParamInput, , Val(txtKegagalan.Text))
        .Parameters.Append .CreateParameter("Tindakan", adVarChar, adParamInput, 200, IIf(Len(Trim(txtTindakan.Text)) = 0, Null, Trim(txtTindakan.Text)))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(Len(Trim(txtKeteranganKB.Text)) = 0, Null, Trim(txtKeteranganKB.Text)))

        .Parameters.Append .CreateParameter("IdParamedis", adChar, adParamInput, 10, IIf(dcPerawatKB.BoundText = "", Null, dcPerawatKB.BoundText))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AU_CatatanProgramKBPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_CatatanProgramKBPasien = False
        Else
            Call Add_HistoryLoginActivity("AU_CatatanProgramKBPasien")
        End If
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    sp_CatatanProgramKBPasien = False
    Call msubPesanError("sp_CatatanProgramKBPasien")
End Function

Private Sub subSimpanKB()
    On Error GoTo errLoad

    If Trim(txtKdDokterKB.Text) = "" Then
        MsgBox "Lengkap nama Dokter/Perawat Pemeriksa", vbExclamation, "Validasi"
        txtPemeriksaKB.SetFocus
        Exit Sub
    End If
    If Periksa("datacombo", dcJenisKontrasepsi, "Jenis kontrasepsi kosong") = False Then Exit Sub
    If sp_CatatanProgramKBPasien() = False Then Exit Sub
    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError("subSimpanKB")
End Sub

Private Sub subSimpanKehamilan()
    On Error GoTo errLoad

    If Trim(txtKdDokterKehamilan.Text) = "" Then
        MsgBox "Lengkap nama Dokter/Perawat Pemeriksa", vbExclamation, "Validasi"
        txtPemeriksaKehamilan.SetFocus
        Exit Sub
    End If
    If Periksa("text", txtGravite, "Nilai Gravite kosong") = False Then Exit Sub
    If Periksa("text", txtPartus, "Nilai Partus kosong") = False Then Exit Sub
    If Periksa("text", txtAbortus, "Nilai Abortus kosong") = False Then Exit Sub
    If Periksa("datacombo", dcImunisasi, "Imunisasi kosong") = False Then Exit Sub
    If Periksa("datacombo", dcPerawatKehamilan, "Paramedis pemeriksa kosong") = False Then Exit Sub
    If sp_CatatanKehamilanPasien() = False Then Exit Sub
    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError("subSimpanKehamilan")
End Sub

Private Sub cmdBatal_Click()
    On Error GoTo errLoad

    Call subKosong
    Call subLoadDcSource
    Call subLoadDataGrid
    fraDokterKehamilan.Visible = False: fraDokterKB.Visible = False
    Exit Sub
errLoad:
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errLoad

    Select Case ssData.Tab
        Case 0
            If dgKehamilan.ApproxCount = 0 Then Exit Sub
            If MsgBox("Apakah anda yakin akan menghapus data kehamilan '" _
            & txtNamaPasien.Text & "'" & vbNewLine _
            & "Dengan tanggal periksa '" & dgKehamilan.Columns("TglPeriksa").value _
            & "'", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
            If sp_DelCatatanKehamilanPasien = False Then Exit Sub
        Case 1
            If dgKB.ApproxCount = 0 Then Exit Sub
            If MsgBox("Apakah anda yakin akan menghapus data keluarga berencana (KB) '" _
            & txtNamaPasien.Text & "'" & vbNewLine _
            & "Dengan tanggal periksa '" & dgKB.Columns("TglPeriksa").value _
            & "'", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
            If sp_DelCatatanProgramKBPasien = False Then Exit Sub
    End Select
    MsgBox "Penghapusan data berhasil", vbInformation, "Informasi"
    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    Select Case ssData.Tab
        Case 0
            'data kehamilan
            Call subSimpanKehamilan
        Case 1
            'data kb
            Call subSimpanKB
    End Select
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcImunisasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcImunisasi.MatchedWithList = True Then txtKeteranganKehamilan.SetFocus
        strSQL = "SELECT KdImunisasi, NamaImunisasi From Imunisasi where StatusEnabled='1' and (NamaImunisasi LIKE '%" & dcImunisasi.Text & "%') ORDER BY NamaImunisasi "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcImunisasi.Text = ""
            Exit Sub
        End If
        dcImunisasi.BoundText = rs(0).value
        dcImunisasi.Text = rs(1).value
    End If
End Sub

Private Sub dcJenisKontrasepsi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcJenisKontrasepsi.MatchedWithList = True Then txtEfekSamping.SetFocus
        strSQL = "SELECT KdJenisKontrasepsi, JenisKontrasepsi From JenisKontrasepsi where StatusEnabled='1' and (JenisKontrasepsi LIKE '%" & dcJenisKontrasepsi.Text & "%')  ORDER BY JenisKontrasepsi "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcJenisKontrasepsi.Text = ""
            Exit Sub
        End If
        dcJenisKontrasepsi.BoundText = rs(0).value
        dcJenisKontrasepsi.Text = rs(1).value
    End If
End Sub

Private Sub dcJenisKontrasepsi_LostFocus()
    If dcJenisKontrasepsi.MatchedWithList = True Then txtEfekSamping.SetFocus
    strSQL = "SELECT KdJenisKontrasepsi, JenisKontrasepsi From JenisKontrasepsi where StatusEnabled='1' and (JenisKontrasepsi LIKE '%" & dcJenisKontrasepsi.Text & "%')  ORDER BY JenisKontrasepsi "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        dcJenisKontrasepsi.Text = ""
        Exit Sub
    End If
    dcJenisKontrasepsi.BoundText = rs(0).value
    dcJenisKontrasepsi.Text = rs(1).value

End Sub

Private Sub dcPerawatKB_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    If KeyAscii = 13 Then
        If Len(Trim(dcPerawatKB.Text)) > 0 Then
            strSQL = "SELECT  IdPegawai, [Nama Pemeriksa]" & _
            " From V_DaftarPemeriksaPasien" & _
            " WHERE ([Nama Pemeriksa] LIKE '%" & dcPerawatKB.Text & "%')"
            Call msubRecFO(rs, strSQL)
            dcPerawatKB.Text = ""
            If rs.EOF = False Then dcPerawatKB.BoundText = rs(0).value: dcJenisKontrasepsi.SetFocus
        Else
            dcJenisKontrasepsi.SetFocus
        End If
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPerawatKB_LostFocus()
        If Len(Trim(dcPerawatKB.Text)) > 0 Then
            strSQL = "SELECT  IdPegawai, [Nama Pemeriksa]" & _
            " From V_DaftarPemeriksaPasien" & _
            " WHERE ([Nama Pemeriksa] LIKE '%" & dcPerawatKB.Text & "%')"
            Call msubRecFO(rs, strSQL)
            dcPerawatKB.Text = ""
            If rs.EOF = False Then dcPerawatKB.BoundText = rs(0).value: dcJenisKontrasepsi.SetFocus
        Else
            dcJenisKontrasepsi.SetFocus
        End If

End Sub

Private Sub dcPerawatKehamilan_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    If KeyAscii = 13 Then
        If Len(Trim(dcPerawatKehamilan.Text)) > 0 Then
            strSQL = "SELECT  IdPegawai, [Nama Pemeriksa]" & _
            " From V_DaftarPemeriksaPasien" & _
            " WHERE ([Nama Pemeriksa] LIKE '%" & dcPerawatKehamilan.Text & "%')"
            Call msubRecFO(rs, strSQL)
            dcPerawatKehamilan.Text = ""
            If rs.EOF = False Then dcPerawatKehamilan.BoundText = rs(0).value: txtBulanKehamilan.SetFocus
        Else
            txtBulanKehamilan.SetFocus
        End If
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPerawatKehamilan_LostFocus()
        If Len(Trim(dcPerawatKehamilan.Text)) > 0 Then
            strSQL = "SELECT  IdPegawai, [Nama Pemeriksa]" & _
            " From V_DaftarPemeriksaPasien" & _
            " WHERE ([Nama Pemeriksa] LIKE '%" & dcPerawatKehamilan.Text & "%')"
            Call msubRecFO(rs, strSQL)
            dcPerawatKehamilan.Text = ""
            If rs.EOF = False Then dcPerawatKehamilan.BoundText = rs(0).value: txtBulanKehamilan.SetFocus
        Else
            txtBulanKehamilan.SetFocus
        End If

End Sub

Private Sub dgDokterKB_DblClick()
    txtKdDokterKB.Text = dgDokterKB.Columns(0)
    subBolTampil = True
    txtPemeriksaKB.Text = dgDokterKB.Columns(1)
    subBolTampil = False
    fraDokterKB.Visible = False
    dcPerawatKB.SetFocus
End Sub

Private Sub dgDokterKB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dgDokterKB.ApproxCount = 0 Then Exit Sub
        Call dgDokterKB_DblClick
    End If
End Sub

Private Sub dgDokterKehamilan_DblClick()
    txtKdDokterKehamilan.Text = dgDokterKehamilan.Columns(0)
    subBolTampil = True
    txtPemeriksaKehamilan.Text = dgDokterKehamilan.Columns(1)
    subBolTampil = False
    fraDokterKehamilan.Visible = False
    dcPerawatKehamilan.SetFocus
End Sub

Private Sub dgDokterKehamilan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dgDokterKehamilan.ApproxCount = 0 Then Exit Sub
        Call dgDokterKehamilan_DblClick
    End If
End Sub

Private Sub dgKehamilan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    With dgKehamilan
        dtpTglPeriksaKehamilan.value = .Columns("TglPeriksa").value
        txtKdDokterKehamilan.Text = .Columns("IdDokter").value
        txtGravite.Text = .Columns("Gravite").value
        txtPartus.Text = .Columns("Partus").value
        txtAbortus.Text = .Columns("Abortus").value
        txtPemeriksaKehamilan.Text = .Columns("Dokter").value
        dcPerawatKehamilan.BoundText = .Columns("IdParamedis").value
        dcImunisasi.BoundText = .Columns("KdImunisasi").value
        txtKeteranganKehamilan.Text = .Columns("Keterangan")
        fraDokterKehamilan.Visible = False
    End With
End Sub

Private Sub dtpTglPeriksaKB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtPemeriksaKB.SetFocus
End Sub

Private Sub dtpTglPeriksaKehamilan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtPemeriksaKehamilan.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    Select Case KeyCode
        Case vbKey1
            If strCtrlKey = 4 Then ssData.SetFocus: ssData.Tab = 0
        Case vbKey2
            If strCtrlKey = 4 Then ssData.SetFocus: ssData.Tab = 1
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)

    ssData.Tab = 1
    ssData.Tab = 0

    With frmTransaksiPasien
        txtNoPendaftaran = .txtNoPendaftaran.Text
        txtNoCM = .txtNoCM.Text
        txtNamaPasien = .txtNamaPasien.Text
        txtSex.Text = .txtSex.Text
        txtThn = .txtThn.Text
        txtBln = .txtBln.Text
        txtHari = .txtHr.Text
    End With
    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmTransaksiPasien.Enabled = True
End Sub

Private Sub ssData_Click(PreviousTab As Integer)
    On Error GoTo errLoad

    Call subKosong
    Call subLoadDataGrid

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub ssData_KeyPress(KeyAscii As Integer)
    Select Case ssData.Tab
        Case 0
            dtpTglPeriksaKehamilan.SetFocus
        Case 1
            dtpTglPeriksaKB.SetFocus
    End Select
End Sub

Private Sub txtAbortus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcImunisasi.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtBulanKehamilan_Change()
    Dim dTglLahir As Date
    If txtBulanKehamilan.Text = "" And txtTahunKehamilan.Text = "" Then txtHariKehamilan.SetFocus: Exit Sub
    If txtBulanKehamilan.Text = "" Then txtBulanKehamilan.Text = 0
    If txtTahunKehamilan.Text = "" And txtHariKehamilan.Text = "" Then
        dTglLahir = DateAdd("m", -1 * CInt(txtBulanKehamilan.Text), Date)
    ElseIf txtTahunKehamilan.Text <> "" And txtHariKehamilan.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHariKehamilan.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulanKehamilan.Text), dTglLahir)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahunKehamilan.Text), dTglLahir)
    ElseIf txtTahunKehamilan.Text = "" And txtHariKehamilan.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHariKehamilan.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulanKehamilan.Text), dTglLahir)
    ElseIf txtTahunKehamilan.Text <> "" And txtHariKehamilan.Text = "" Then
        dTglLahir = DateAdd("m", -1 * CInt(txtBulanKehamilan.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahunKehamilan.Text), dTglLahir)
    End If
    meTglLahir.Text = dTglLahir
End Sub

Private Sub txtBulanKehamilan_KeyPress(KeyAscii As Integer)
    Dim dTglLahir As Date
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If txtBulanKehamilan.Text = "" And txtTahunKehamilan.Text = "" Then txtHariKehamilan.SetFocus: Exit Sub
        If txtBulanKehamilan.Text = "" Then txtBulanKehamilan.Text = 0
        If txtTahunKehamilan.Text = "" And txtHariKehamilan.Text = "" Then
            dTglLahir = DateAdd("m", -1 * CInt(txtBulanKehamilan.Text), Date)
        ElseIf txtTahunKehamilan.Text <> "" And txtHariKehamilan.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHariKehamilan.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulanKehamilan.Text), dTglLahir)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahunKehamilan.Text), dTglLahir)
        ElseIf txtTahunKehamilan.Text = "" And txtHariKehamilan.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHariKehamilan.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulanKehamilan.Text), dTglLahir)
        ElseIf txtTahunKehamilan.Text <> "" And txtHariKehamilan.Text = "" Then
            dTglLahir = DateAdd("m", -1 * CInt(txtBulanKehamilan.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahunKehamilan.Text), dTglLahir)
        End If
        meTglLahir.Text = dTglLahir
        txtHariKehamilan.SetFocus
    End If
End Sub

Private Sub txtBulanKehamilan_LostFocus()
    txtBulanKehamilan.Text = Val(txtBulanKehamilan.Text)
    If Val(txtBulanKehamilan.Text) > 10 Then txtBulanKehamilan.Text = "10"

End Sub

Private Sub txtEfekSamping_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKegagalan.SetFocus
End Sub

Private Sub txtGravite_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPartus.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtHariKehamilan_Change()
    Dim dTglLahir As Date
    If txtHariKehamilan.Text = "" And txtBulanKehamilan.Text = "" And txtTahunKehamilan.Text = "" Then txtGravite.SetFocus: Exit Sub
    If txtHariKehamilan.Text = "" Then txtHariKehamilan.Text = 0
    If txtTahunKehamilan.Text = "" And txtBulanKehamilan.Text = "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHariKehamilan.Text), Date)
    ElseIf txtTahunKehamilan.Text <> "" And txtBulanKehamilan.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHariKehamilan.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulanKehamilan.Text), dTglLahir)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahunKehamilan.Text), dTglLahir)
    ElseIf txtTahunKehamilan.Text = "" And txtBulanKehamilan.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHariKehamilan.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulanKehamilan.Text), dTglLahir)
    ElseIf txtTahunKehamilan.Text <> "" And txtBulanKehamilan.Text = "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHariKehamilan.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahunKehamilan.Text), dTglLahir)
    End If
    meTglLahir.Text = dTglLahir
End Sub

Private Sub txtHariKehamilan_KeyPress(KeyAscii As Integer)
    Dim dTglLahir As Date
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If txtHariKehamilan.Text = "" And txtBulanKehamilan.Text = "" And txtTahunKehamilan.Text = "" Then txtGravite.SetFocus: Exit Sub
        If txtHariKehamilan.Text = "" Then txtHariKehamilan.Text = 0
        If txtTahunKehamilan.Text = "" And txtBulanKehamilan.Text = "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHariKehamilan.Text), Date)
        ElseIf txtTahunKehamilan.Text <> "" And txtBulanKehamilan.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHariKehamilan.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulanKehamilan.Text), dTglLahir)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahunKehamilan.Text), dTglLahir)
        ElseIf txtTahunKehamilan.Text = "" And txtBulanKehamilan.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHariKehamilan.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulanKehamilan.Text), dTglLahir)
        ElseIf txtTahunKehamilan.Text <> "" And txtBulanKehamilan.Text = "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHariKehamilan.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahunKehamilan.Text), dTglLahir)
        End If
        meTglLahir.Text = dTglLahir
        txtGravite.SetFocus
    End If
End Sub

Private Sub txtHariKehamilan_LostFocus()
    txtHariKehamilan.Text = Val(txtHariKehamilan.Text)
    If Val(txtHariKehamilan.Text) > 31 Then txtHariKehamilan.Text = "31"
End Sub

Private Sub txtKegagalan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTindakan.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtKeteranganKB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKeteranganKB_LostFocus()
    Dim i As Integer
    Dim tempText As String
    tempText = Trim(txtKeteranganKB.Text)
    txtKeteranganKB.Text = ""
    For i = 1 To Len(tempText)
        If Asc(Mid(tempText, i, 1)) <> 10 And Asc(Mid(tempText, i, 1)) <> 13 Then
            txtKeteranganKB.Text = txtKeteranganKB.Text & Mid(tempText, i, 1)
        End If
    Next i
End Sub

Private Sub txtKeteranganKehamilan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKeteranganKehamilan_LostFocus()
    Dim i As Integer
    Dim tempText As String
    tempText = Trim(txtKeteranganKehamilan.Text)
    txtKeteranganKehamilan.Text = ""
    For i = 1 To Len(tempText)
        If Asc(Mid(tempText, i, 1)) <> 10 And Asc(Mid(tempText, i, 1)) <> 13 Then
            txtKeteranganKehamilan.Text = txtKeteranganKehamilan.Text & Mid(tempText, i, 1)
        End If
    Next i
End Sub

Private Sub txtPartus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAbortus.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtPemeriksaKB_Change()
    On Error GoTo errLoad

    If subBolTampil = True Then Exit Sub
    strSQL = "SELECT IdPegawai AS [Kode Pemeriksa], [Nama Pemeriksa],JK,[Jenis Pemeriksa] " & _
    " FROM V_DaftarDokterdanPemeriksaPasien WHERE [Nama Pemeriksa] LIKE '%" & txtPemeriksaKB.Text & "%'"
    Call msubRecFO(rs, strSQL)
    Set dgDokterKB.DataSource = rs
    With dgDokterKB
        .Columns(0).Width = 1500
        .Columns(1).Width = 4000
        .Columns(2).Width = 400
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 2000
    End With
    fraDokterKB.Left = 240
    fraDokterKB.Top = 1440
    fraDokterKB.Visible = True
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtPemeriksaKB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If fraDokterKB.Visible = False Then Exit Sub
        dgDokterKB.SetFocus
    End If
End Sub

Private Sub txtPemeriksaKB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If fraDokterKB.Visible = True Then dgDokterKB.SetFocus Else dcPerawatKB.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtPemeriksaKehamilan_Change()
    On Error GoTo errLoad

    If subBolTampil = True Then Exit Sub
    strSQL = "SELECT IdPegawai AS [Kode Pemeriksa], [Nama Pemeriksa],JK,[Jenis Pemeriksa] " & _
    " FROM V_DaftarDokterdanPemeriksaPasien WHERE [Nama Pemeriksa] LIKE '%" & txtPemeriksaKehamilan.Text & "%'"
    Call msubRecFO(rs, strSQL)
    Set dgDokterKehamilan.DataSource = rs
    With dgDokterKehamilan
        .Columns(0).Width = 1500
        .Columns(1).Width = 4000
        .Columns(2).Width = 400
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 2000
    End With
    fraDokterKehamilan.Left = 240
    fraDokterKehamilan.Top = 1440
    fraDokterKehamilan.Visible = True
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtPemeriksaKehamilan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If fraDokterKehamilan.Visible = False Then Exit Sub
        dgDokterKehamilan.SetFocus
    End If
End Sub

Private Sub txtPemeriksaKehamilan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If fraDokterKehamilan.Visible = True Then dgDokterKehamilan.SetFocus Else dcPerawatKehamilan.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtTahunKehamilan_Change()
    Dim dTglLahir As Date
    If txtTahunKehamilan = "" Then txtBulanKehamilan.SetFocus: Exit Sub
    If txtBulanKehamilan.Text = "" And txtHariKehamilan.Text = "" Then
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahunKehamilan.Text), Date)
    ElseIf txtBulanKehamilan.Text <> "" And txtHariKehamilan.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHariKehamilan.Text), Date)
        dTglLahir = DateAdd("m", -1 * CInt(txtBulanKehamilan.Text), dTglLahir)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahunKehamilan.Text), dTglLahir)
    ElseIf txtBulanKehamilan.Text = "" And txtHariKehamilan.Text <> "" Then
        dTglLahir = DateAdd("d", -1 * CInt(txtHariKehamilan.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahunKehamilan.Text), dTglLahir)
    ElseIf txtBulanKehamilan.Text <> "" And txtHariKehamilan.Text = "" Then
        dTglLahir = DateAdd("m", -1 * CInt(txtBulanKehamilan.Text), Date)
        dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahunKehamilan.Text), dTglLahir)
    End If
    meTglLahir.Text = dTglLahir
End Sub

Private Sub txtTahunKehamilan_KeyPress(KeyAscii As Integer)
    Dim dTglLahir As Date
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If txtTahunKehamilan = "" Then txtBulanKehamilan.SetFocus: Exit Sub
        If txtBulanKehamilan.Text = "" And txtHariKehamilan.Text = "" Then
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahunKehamilan.Text), Date)
        ElseIf txtBulanKehamilan.Text <> "" And txtHariKehamilan.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHariKehamilan.Text), Date)
            dTglLahir = DateAdd("m", -1 * CInt(txtBulanKehamilan.Text), dTglLahir)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahunKehamilan.Text), dTglLahir)
        ElseIf txtBulanKehamilan.Text = "" And txtHariKehamilan.Text <> "" Then
            dTglLahir = DateAdd("d", -1 * CInt(txtHariKehamilan.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahunKehamilan.Text), dTglLahir)
        ElseIf txtBulanKehamilan.Text <> "" And txtHariKehamilan.Text = "" Then
            dTglLahir = DateAdd("m", -1 * CInt(txtBulanKehamilan.Text), Date)
            dTglLahir = DateAdd("yyyy", -1 * CInt(txtTahunKehamilan.Text), dTglLahir)
        End If
        meTglLahir.Text = dTglLahir
        txtBulanKehamilan.SetFocus
    End If
End Sub

Private Sub txtTahunKehamilan_LostFocus()
    txtTahunKehamilan.Text = Val(txtTahunKehamilan.Text)
    If Val(txtTahunKehamilan.Text) > 100 Then txtTahunKehamilan.Text = "100"
End Sub

Private Sub txtTindakan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeteranganKB.SetFocus
End Sub

