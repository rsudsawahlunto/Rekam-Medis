VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDataKegiatanKesehatanLingkunganRL5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Kegiatan Kesehatan Lingkungan"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDataKegiatanKesehatanLingkunganRL5.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   14325
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   100
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton cmdBatal 
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
      Height          =   495
      Left            =   10320
      TabIndex        =   99
      Top             =   7680
      Width           =   1935
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
      Height          =   495
      Left            =   12360
      TabIndex        =   98
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Frame Frame17 
      Caption         =   "D. Penyehatan Air"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   72
      Top             =   5880
      Width           =   14295
      Begin VB.Frame Frame21 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10560
         TabIndex        =   89
         Top             =   315
         Width           =   3135
         Begin VB.OptionButton optDKuartal1TMS 
            Caption         =   "TMS"
            Height          =   255
            Left            =   1680
            TabIndex        =   91
            Top             =   160
            Width           =   1335
         End
         Begin VB.OptionButton optDKuartal1MS 
            Caption         =   "MS"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame18 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   73
         Top             =   120
         Width           =   3735
         Begin VB.OptionButton optAirBersihSumur 
            Caption         =   "SUMUR BOR/ SUMUR GALI"
            Height          =   255
            Left            =   1320
            TabIndex        =   75
            Top             =   160
            Width           =   2295
         End
         Begin VB.OptionButton optAirBersihPDAM 
            Caption         =   "PDAM"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame19 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   77
         Top             =   480
         Width           =   3735
         Begin VB.OptionButton optKuantitasTdkCkp 
            Caption         =   "TIDAK CUKUP"
            Height          =   255
            Left            =   1320
            TabIndex        =   79
            Top             =   160
            Width           =   1335
         End
         Begin VB.OptionButton optKuantitasCkp 
            Caption         =   "CUKUP"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame20 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   81
         Top             =   840
         Width           =   3735
         Begin VB.OptionButton optKontinuitasYa 
            Caption         =   "YA"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optKontinuitasTdk 
            Caption         =   "TIDAK"
            Height          =   255
            Left            =   1320
            TabIndex        =   82
            Top             =   160
            Width           =   1335
         End
      End
      Begin VB.Frame Frame22 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10560
         TabIndex        =   92
         Top             =   675
         Width           =   3135
         Begin VB.OptionButton optDKuartal2MS 
            Caption         =   "MS"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optDKuartal2TMS 
            Caption         =   "TMS"
            Height          =   255
            Left            =   1680
            TabIndex        =   93
            Top             =   160
            Width           =   1335
         End
      End
      Begin VB.Frame Frame23 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10560
         TabIndex        =   95
         Top             =   1035
         Width           =   3135
         Begin VB.OptionButton optDKuartal3MS 
            Caption         =   "MS"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optDKuartal3TMS 
            Caption         =   "TMS"
            Height          =   255
            Left            =   1680
            TabIndex        =   96
            Top             =   160
            Width           =   1335
         End
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "-  Kuartal III"
         Height          =   195
         Left            =   8040
         TabIndex        =   88
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "-  Kuartal II"
         Height          =   195
         Left            =   8040
         TabIndex        =   87
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "-  Kuartal I"
         Height          =   195
         Left            =   8040
         TabIndex        =   86
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "4. Kualitas air Minum (mikrobiologi)"
         Height          =   195
         Left            =   7560
         TabIndex        =   85
         Top             =   240
         Width           =   2460
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "3. Kontinuitas (tersedia dalam 24 jam)"
         Height          =   195
         Left            =   240
         TabIndex        =   84
         Top             =   1035
         Width           =   2730
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "2. Kuantitas"
         Height          =   195
         Left            =   240
         TabIndex        =   80
         Top             =   690
         Width           =   870
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "1. Tersedia sarana air bersih"
         Height          =   195
         Left            =   240
         TabIndex        =   76
         Top             =   330
         Width           =   2055
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "C. Pengelolaan Limbah Padat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   7200
      TabIndex        =   45
      Top             =   2160
      Width           =   7095
      Begin VB.Frame Frame16 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   67
         Top             =   3120
         Width           =   3615
         Begin VB.OptionButton optCSumberDanaAPBD 
            Caption         =   "APBD"
            Height          =   255
            Left            =   1080
            TabIndex        =   71
            Top             =   160
            Width           =   855
         End
         Begin VB.OptionButton optCSumberDanaAPBN 
            Caption         =   "APBN"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   160
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optCSumberDanaBLN 
            Caption         =   "BLN"
            Height          =   255
            Left            =   2040
            TabIndex        =   69
            Top             =   160
            Width           =   855
         End
         Begin VB.OptionButton optCSumberDanaRS 
            Caption         =   "RS"
            Height          =   255
            Left            =   2880
            TabIndex        =   68
            Top             =   160
            Width           =   615
         End
      End
      Begin VB.TextBox txtC5 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4680
         TabIndex        =   61
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtC4 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3480
         TabIndex        =   59
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Frame Frame13 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   46
         Top             =   120
         Width           =   3135
         Begin VB.OptionButton optSaranaInsineratorTdkAda 
            Caption         =   "TIDAK ADA"
            Height          =   255
            Left            =   1680
            TabIndex        =   48
            Top             =   160
            Width           =   1335
         End
         Begin VB.OptionButton optSaranaInsineratorAda 
            Caption         =   "ADA"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame14 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   50
         Top             =   480
         Width           =   3135
         Begin VB.OptionButton optInsineratorYa 
            Caption         =   "YA"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optInsineratorTdk 
            Caption         =   "TIDAK"
            Height          =   255
            Left            =   1680
            TabIndex        =   51
            Top             =   160
            Width           =   1335
         End
      End
      Begin VB.Frame Frame15 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   54
         Top             =   840
         Width           =   3135
         Begin VB.OptionButton optPermenkesTdk 
            Caption         =   "TIDAK"
            Height          =   255
            Left            =   1680
            TabIndex        =   56
            Top             =   160
            Width           =   1335
         End
         Begin VB.OptionButton optPermenkesYa 
            Caption         =   "YA"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin MSComCtl2.DTPicker dtpInsinerator 
         Height          =   345
         Left            =   3960
         TabIndex        =   63
         Top             =   2760
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   126812163
         UpDown          =   -1  'True
         CurrentDate     =   38212
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "7. Sumber dana"
         Height          =   195
         Left            =   360
         TabIndex        =   66
         Top             =   3240
         Width           =   1140
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "tahun"
         Height          =   195
         Left            =   3360
         TabIndex        =   65
         Top             =   2820
         Width           =   420
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "6. Insinerator dibangun"
         Height          =   195
         Left            =   360
         TabIndex        =   62
         Top             =   2820
         Width           =   1695
      End
      Begin VB.Label Label17 
         Caption         =   "5. Bila tidak mempunyai insinerator kemana dimusnahkan"
         Height          =   435
         Left            =   360
         TabIndex        =   60
         Top             =   2400
         Width           =   4365
      End
      Begin VB.Label Label16 
         Caption         =   "4. Bila insinerator tidak berfungsi, kemana        Limbah padat medis dimusnahkan"
         Height          =   435
         Left            =   360
         TabIndex        =   58
         Top             =   1920
         Width           =   3165
      End
      Begin VB.Label Label15 
         Caption         =   $"frmDataKegiatanKesehatanLingkunganRL5.frx":0CCA
         Height          =   1020
         Left            =   360
         TabIndex        =   57
         Top             =   1035
         Width           =   2685
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "2. Insinirator berfungsi"
         Height          =   195
         Left            =   360
         TabIndex        =   53
         Top             =   675
         Width           =   1650
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "1. Sarana insinerator"
         Height          =   195
         Left            =   360
         TabIndex        =   49
         Top             =   330
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "B. Pengelolaan Limbah Cair"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   15
      Top             =   2160
      Width           =   7095
      Begin VB.Frame Frame8 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   28
         Top             =   1280
         Width           =   3135
         Begin VB.OptionButton optBKuartal1MS 
            Caption         =   "MS"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optBKuartal1TMS 
            Caption         =   "TMS"
            Height          =   255
            Left            =   1680
            TabIndex        =   29
            Top             =   160
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   16
         Top             =   120
         Width           =   3135
         Begin VB.OptionButton optInstLimbahCairAda 
            Caption         =   "ADA"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optInstLimbahCairTdkAda 
            Caption         =   "TIDAK ADA"
            Height          =   255
            Left            =   1680
            TabIndex        =   17
            Top             =   160
            Width           =   1335
         End
      End
      Begin VB.Frame Frame7 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   19
         Top             =   480
         Width           =   3135
         Begin VB.OptionButton optIPALTdk 
            Caption         =   "TIDAK"
            Height          =   255
            Left            =   1680
            TabIndex        =   21
            Top             =   160
            Width           =   1335
         End
         Begin VB.OptionButton optIPALYa 
            Caption         =   "YA"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame9 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   31
         Top             =   1640
         Width           =   3135
         Begin VB.OptionButton optBKuartal2TMS 
            Caption         =   "TMS"
            Height          =   255
            Left            =   1680
            TabIndex        =   33
            Top             =   160
            Width           =   1335
         End
         Begin VB.OptionButton optBKuartal2MS 
            Caption         =   "MS"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame10 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   34
         Top             =   2000
         Width           =   3135
         Begin VB.OptionButton optBKuartal3TMS 
            Caption         =   "TMS"
            Height          =   255
            Left            =   1680
            TabIndex        =   36
            Top             =   160
            Width           =   1335
         End
         Begin VB.OptionButton optBKuartal3MS 
            Caption         =   "MS"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin MSComCtl2.DTPicker dtpIPAL 
         Height          =   345
         Left            =   3960
         TabIndex        =   38
         Top             =   2580
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   126943235
         UpDown          =   -1  'True
         CurrentDate     =   38212
      End
      Begin VB.Frame Frame11 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   40
         Top             =   2880
         Width           =   3615
         Begin VB.OptionButton optBSumberDanaRS 
            Caption         =   "RS"
            Height          =   255
            Left            =   2880
            TabIndex        =   44
            Top             =   160
            Width           =   615
         End
         Begin VB.OptionButton optBSumberDanaBLN 
            Caption         =   "BLN"
            Height          =   255
            Left            =   2040
            TabIndex        =   43
            Top             =   160
            Width           =   855
         End
         Begin VB.OptionButton optBSumberDanaAPBN 
            Caption         =   "APBN"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   160
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optBSumberDanaAPBD 
            Caption         =   "APBD"
            Height          =   255
            Left            =   1080
            TabIndex        =   41
            Top             =   160
            Width           =   855
         End
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "tahun"
         Height          =   195
         Left            =   3360
         TabIndex        =   64
         Top             =   2640
         Width           =   420
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "5. Sumber dana"
         Height          =   195
         Left            =   360
         TabIndex        =   39
         Top             =   3000
         Width           =   1140
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "4. Sarana IPAL dibangun"
         Height          =   195
         Left            =   360
         TabIndex        =   37
         Top             =   2640
         Width           =   1785
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "-  Kuartal III"
         Height          =   195
         Left            =   840
         TabIndex        =   27
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "-  Kuartal II"
         Height          =   195
         Left            =   840
         TabIndex        =   26
         Top             =   1800
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "-  Kuartal I"
         Height          =   195
         Left            =   840
         TabIndex        =   25
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "3. Kualitas effluent (sesuai Kepmen LH 58/95 atau perda setempat)"
         Height          =   195
         Left            =   360
         TabIndex        =   24
         Top             =   1080
         Width           =   4845
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "1. Instalasi Pengelolaan Limbah Cair"
         Height          =   195
         Left            =   360
         TabIndex        =   23
         Top             =   330
         Width           =   2580
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "2. IPAL Berfungsi"
         Height          =   195
         Left            =   360
         TabIndex        =   22
         Top             =   675
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "A. Dokumen"
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
      Left            =   7200
      TabIndex        =   6
      Top             =   1080
      Width           =   7095
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   9
         Top             =   120
         Width           =   3135
         Begin VB.OptionButton optDokAmdalTdkAda 
            Caption         =   "TIDAK ADA"
            Height          =   255
            Left            =   1680
            TabIndex        =   11
            Top             =   160
            Width           =   1335
         End
         Begin VB.OptionButton optDokAmdalAda 
            Caption         =   "ADA"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   12
         Top             =   480
         Width           =   3135
         Begin VB.OptionButton optDokUKLAda 
            Caption         =   "ADA"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optDokUKLTdkAda 
            Caption         =   "TIDAK ADA"
            Height          =   255
            Left            =   1680
            TabIndex        =   13
            Top             =   160
            Width           =   1335
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "2. Dokumen UKL dan UPL"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   675
         Width           =   1800
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "1. Dokumen Amdal"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   330
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   7095
      Begin VB.TextBox txtKdRS 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   3360
         TabIndex        =   4
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtNamaRS 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   3360
         TabIndex        =   2
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Kode Rumah Sakit"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   630
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Rumah Sakit"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   1335
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
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
      Left            =   12480
      Picture         =   "frmDataKegiatanKesehatanLingkunganRL5.frx":0D62
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataKegiatanKesehatanLingkunganRL5.frx":1AEA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmDataKegiatanKesehatanLingkunganRL5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub cmdBatal_Click()
    Call subKosong
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    Dim sBSmbrDana, sCSmbrDana As String

    Set rs = Nothing
    Call msubRecFO(rs, "Delete From TempDataKegiatanKesehatanLingkungan_RL5")

    If optBSumberDanaAPBN.value = True Then
        sBSmbrDana = "APBN"
    ElseIf optBSumberDanaAPBD.value = True Then
        sBSmbrDana = "APBD"
    ElseIf optBSumberDanaBLN.value = True Then
        sBSmbrDana = "BLN"
    ElseIf optBSumberDanaRS.value = True Then
        sBSmbrDana = "RS"
    End If

    If optCSumberDanaAPBN.value = True Then
        sCSmbrDana = "APBN"
    ElseIf optCSumberDanaAPBD.value = True Then
        sCSmbrDana = "APBD"
    ElseIf optCSumberDanaBLN.value = True Then
        sCSmbrDana = "BLN"
    ElseIf optCSumberDanaRS.value = True Then
        sCSmbrDana = "RS"
    End If

    strSQL = "Insert Into TempDataKegiatanKesehatanLingkungan_RL5 Values ('" & strKdRS & "','" & strNNamaRS & "', " & IIf(optDokAmdalAda.value = True, 0, 1) & "," & IIf(optDokUKLAda.value = True, 0, 1) & "," & IIf(optInstLimbahCairAda.value = True, 0, 1) & "," & IIf(optIPALYa.value = True, 0, 1) & "," & IIf(optBKuartal1MS.value = True, 0, 1) & "," & IIf(optBKuartal2MS.value = True, 0, 1) & "," & IIf(optBKuartal3MS.value = True, 0, 1) & ",'" & Format(dtpIPAL.value, "yyyy") & "','" & sBSmbrDana & "'," & _
    "" & IIf(optSaranaInsineratorAda.value = True, 0, 1) & "," & IIf(optInsineratorYa.value = True, 0, 1) & "," & IIf(optPermenkesYa.value = True, 0, 1) & ",'" & IIf(Trim(txtC4.Text) = "", Null, Trim(txtC4.Text)) & "','" & IIf(Trim(txtC5.Text) = "", Null, Trim(txtC5.Text)) & "','" & Format(dtpInsinerator.value, "yyyy") & "','" & sCSmbrDana & "'," & IIf(optAirBersihPDAM.value = True, 0, 1) & "," & IIf(optKuantitasCkp.value = True, 0, 1) & "," & IIf(optKontinuitasYa = True, 0, 1) & "," & _
    "" & IIf(optDKuartal1MS.value = True, 0, 1) & "," & IIf(optDKuartal2MS.value = True, 0, 1) & "," & IIf(optDKuartal3MS.value = True, 0, 1) & ")"

    Set rs = Nothing
    Call msubRecFO(rs, strSQL)

    frmLapRL5.Show

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpInsinerator_Change()
    dtpInsinerator.MaxDate = Now
End Sub

Private Sub dtpIPAL_Change()
    dtpIPAL.MaxDate = Now
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call subKosong

    txtNamaRS.Text = strNNamaRS
    txtKdRS.Text = strKdRS
End Sub

Sub subKosong()
    optDokAmdalAda.value = True
    optDokUKLAda.value = True
    optInstLimbahCairAda.value = True
    optIPALYa.value = True
    optBKuartal1MS.value = True
    optBKuartal2MS.value = True
    optBKuartal3MS.value = True
    dtpIPAL.value = Now
    optBSumberDanaAPBN.value = True
    optSaranaInsineratorAda.value = True
    optInsineratorYa.value = True
    optPermenkesYa.value = True
    txtC4.Text = ""
    txtC5.Text = ""
    dtpInsinerator.value = Now
    optCSumberDanaAPBN.value = True
    optAirBersihPDAM.value = True
    optKuantitasCkp.value = True
    optKontinuitasYa.value = True
    optDKuartal1MS.value = True
    optDKuartal2MS.value = True
    optDKuartal3MS.value = True
End Sub

