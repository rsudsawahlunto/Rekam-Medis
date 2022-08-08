VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUbahJenisPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Asuransi Pasien"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   10635
   Begin MSFlexGridLib.MSFlexGrid fgDPJP 
      Height          =   2055
      Left            =   2505
      TabIndex        =   144
      Top             =   7200
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3625
      _Version        =   393216
   End
   Begin VB.Frame fraLakalantas 
      Height          =   5295
      Left            =   2520
      TabIndex        =   105
      Top             =   9240
      Visible         =   0   'False
      Width           =   6495
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2055
         Left            =   1200
         TabIndex        =   106
         Top             =   3120
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3625
         _Version        =   393216
      End
      Begin VB.Frame Frame3 
         Caption         =   "SUPLESI"
         Height          =   855
         Left            =   120
         TabIndex        =   136
         Top             =   3360
         Width           =   6255
         Begin VB.CheckBox chkSuplesi 
            Height          =   255
            Left            =   240
            TabIndex        =   139
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtSuplesi 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1680
            MaxLength       =   30
            TabIndex        =   137
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   360
            Width           =   4455
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SEP Suplesi"
            Height          =   210
            Index           =   31
            Left            =   600
            TabIndex        =   138
            Top             =   360
            Width           =   930
         End
      End
      Begin VB.TextBox txtKet 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   134
         TabStop         =   0   'False
         Text            =   "-"
         Top             =   4320
         Width           =   5175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Batal"
         Height          =   375
         Left            =   5280
         TabIndex        =   131
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdSimpanLaka 
         Caption         =   "Simpan"
         Height          =   375
         Left            =   4200
         TabIndex        =   130
         Top             =   4800
         Width           =   975
      End
      Begin VB.Frame Frame6 
         Caption         =   "Lokasi Kejadian"
         Height          =   2055
         Left            =   120
         TabIndex        =   114
         Top             =   1320
         Width           =   6255
         Begin VB.TextBox txtPropinsi 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1080
            TabIndex        =   123
            Top             =   420
            Width           =   3135
         End
         Begin VB.TextBox txtKdPropinsi 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4320
            TabIndex        =   122
            Top             =   420
            Width           =   1215
         End
         Begin VB.CommandButton cmdPropinsi 
            Caption         =   "..."
            Height          =   375
            Left            =   5520
            TabIndex        =   121
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton cmdKota 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   120
            Top             =   855
            Width           =   615
         End
         Begin VB.TextBox txtKdKota 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4320
            TabIndex        =   119
            Top             =   900
            Width           =   1215
         End
         Begin VB.TextBox txtKota 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1080
            TabIndex        =   118
            Top             =   900
            Width           =   3135
         End
         Begin VB.TextBox txtKec 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1080
            TabIndex        =   117
            Top             =   1425
            Width           =   3135
         End
         Begin VB.TextBox txtKdKec 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4320
            TabIndex        =   116
            Top             =   1425
            Width           =   1215
         End
         Begin VB.CommandButton cmdKec 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   115
            Top             =   1380
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Propinsi"
            Height          =   255
            Left            =   120
            TabIndex        =   126
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Kab/Kota"
            Height          =   255
            Left            =   120
            TabIndex        =   125
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Kecamatan"
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   1425
            Width           =   975
         End
      End
      Begin VB.Frame fraPenjaminLaka 
         Caption         =   "Penjamin Lakalantas"
         Height          =   735
         Left            =   120
         TabIndex        =   109
         Top             =   120
         Width           =   6255
         Begin VB.CheckBox chkJasaRaharja 
            Caption         =   "Jasa Raharja"
            Height          =   210
            Left            =   240
            TabIndex        =   113
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox chkBPJSKK 
            Caption         =   "BPJS Ketenagakerjaan"
            Height          =   210
            Left            =   1680
            TabIndex        =   112
            Top             =   360
            Width           =   2175
         End
         Begin VB.CheckBox chkTaspen 
            Caption         =   "TASPEN"
            Height          =   210
            Left            =   3960
            TabIndex        =   111
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox chkAsabri 
            Caption         =   "ASABRI"
            Height          =   210
            Left            =   5160
            TabIndex        =   110
            Top             =   360
            Width           =   975
         End
      End
      Begin MSComCtl2.DTPicker dtpTglKejadian 
         Height          =   315
         Left            =   1800
         TabIndex        =   108
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   163315713
         UpDown          =   -1  'True
         CurrentDate     =   37694
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ket Laka"
         Height          =   210
         Index           =   28
         Left            =   240
         TabIndex        =   135
         Top             =   4320
         Width           =   705
      End
      Begin VB.Label Label6 
         Caption         =   "Tanggal Kejadian"
         Height          =   255
         Left            =   240
         TabIndex        =   107
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame fraPengajuanSEP 
      Caption         =   "Pengajuan SEP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   99
      Top             =   5040
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdBatalPengajuanSEP 
         Caption         =   "&Tutup"
         Height          =   375
         Left            =   4800
         TabIndex        =   102
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdSimpanPengajuanSEP 
         Caption         =   "Simpan"
         Height          =   375
         Left            =   3360
         TabIndex        =   101
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtKetPengajuan 
         BackColor       =   &H0080FF80&
         Height          =   735
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   100
         Top             =   480
         Width           =   6015
      End
      Begin VB.Label Label1 
         Caption         =   "Silahkan masukkan keterangan pengajuan SEP"
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdUpdateTglPulangBPJS 
      Caption         =   "Update Tgl Pulang BPJS"
      Height          =   735
      Left            =   1200
      TabIndex        =   98
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdLPK 
      Caption         =   "Lembar Pengajun Klaim"
      Height          =   735
      Left            =   4680
      TabIndex        =   97
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdPengajuanSEP 
      Caption         =   "Pengajuan SEP"
      Height          =   735
      Left            =   3600
      TabIndex        =   96
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton cmdApprovalSEP 
      Caption         =   "Approval SEP"
      Height          =   735
      Left            =   2400
      TabIndex        =   95
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton cmdRujukan 
      Caption         =   "&Rujuk Pasien BPJS"
      Height          =   735
      Left            =   0
      TabIndex        =   94
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdateSEP 
      Caption         =   "Update SEP"
      Height          =   735
      Left            =   7080
      TabIndex        =   93
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdHapusSEP 
      Caption         =   "Hapus SEP"
      Height          =   735
      Left            =   5880
      TabIndex        =   88
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Frame fraHapusSEP 
      Caption         =   "Hapus SEP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   83
      Top             =   5040
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdBatalKetHapus 
         Caption         =   "Batal"
         Height          =   375
         Left            =   4800
         TabIndex        =   86
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdSimpanKetHapus 
         Caption         =   "Simpan"
         Height          =   375
         Left            =   3360
         TabIndex        =   85
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtKetHapus 
         Height          =   735
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   84
         Top             =   480
         Width           =   6015
      End
      Begin VB.Label Label2 
         Caption         =   "Silakan masukkan penyebab hapus SEP"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   7440
      TabIndex        =   78
      Top             =   4560
      Width           =   3135
      Begin VB.CheckBox chkLakalantas 
         Caption         =   "Lakalantas"
         Height          =   210
         Left            =   120
         TabIndex        =   129
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkKatarak 
         Caption         =   "Katarak"
         Height          =   210
         Left            =   120
         TabIndex        =   128
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox chkCob 
         Caption         =   "COB"
         Height          =   210
         Left            =   120
         TabIndex        =   127
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtPpkRujukan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PPK Rujukan"
         Height          =   210
         Index           =   29
         Left            =   120
         TabIndex        =   80
         Top             =   120
         Visible         =   0   'False
         Width           =   1020
      End
   End
   Begin VB.TextBox txtTempHakKelas 
      Height          =   315
      Left            =   8760
      TabIndex        =   77
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtFormRegistrasiPengirim 
      Height          =   495
      Left            =   3360
      TabIndex        =   73
      Text            =   "Text1"
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtNoBKM 
      Height          =   375
      Left            =   4800
      TabIndex        =   69
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtKdInstalasi 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      TabIndex        =   67
      Text            =   "txtKdInstalasi"
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraDataRujukan 
      Caption         =   "Data Rujukan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   57
      Top             =   6240
      Width           =   10575
      Begin VB.TextBox txtKdDPJP 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5400
         MaxLength       =   30
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtDPJP 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtNoSKDP 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         MaxLength       =   6
         TabIndex        =   141
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txtCatatan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6600
         MaxLength       =   30
         TabIndex        =   132
         TabStop         =   0   'False
         Text            =   "-"
         Top             =   1680
         Width           =   3855
      End
      Begin VB.TextBox txtNoRujukan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   480
         Width           =   3975
      End
      Begin MSDataListLib.DataCombo dcAsalRujukan 
         Height          =   330
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
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
      Begin MSComCtl2.DTPicker dtpTglDirujuk 
         Height          =   315
         Left            =   240
         TabIndex        =   27
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
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
         Format          =   163905539
         UpDown          =   -1  'True
         CurrentDate     =   37694
      End
      Begin MSDataListLib.DataCombo dcNamaPerujuk 
         Height          =   330
         Left            =   2520
         TabIndex        =   28
         Top             =   1080
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcNamaAsalRujukan 
         Height          =   330
         Left            =   6600
         TabIndex        =   26
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcDiagnosa 
         Height          =   330
         Left            =   6600
         TabIndex        =   29
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KD DPJP"
         Height          =   210
         Index           =   35
         Left            =   5400
         TabIndex        =   146
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DPJP"
         Height          =   210
         Index           =   34
         Left            =   2520
         TabIndex        =   142
         Top             =   1440
         Width           =   405
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. SKDP"
         Height          =   210
         Index           =   33
         Left            =   240
         TabIndex        =   140
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan"
         Height          =   210
         Index           =   30
         Left            =   6600
         TabIndex        =   133
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama tempat Perujuk = Nama Puskesmas/ Nama Klinik/ Tempat Dokter Praktek/ Nama Rumah Sakit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   240
         TabIndex        =   64
         Top             =   2040
         Width           =   8565
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Perujuk (Dokter, Bidan, Mantri, dll)"
         Height          =   210
         Index           =   21
         Left            =   2520
         TabIndex        =   63
         Top             =   840
         Width           =   3345
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Rujukan"
         Height          =   210
         Index           =   24
         Left            =   2520
         TabIndex        =   62
         Top             =   240
         Width           =   930
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asal Rujukan"
         Height          =   210
         Index           =   25
         Left            =   240
         TabIndex        =   61
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Asal Rujukan (Nama Tempat Rujukan)"
         Height          =   210
         Index           =   27
         Left            =   6360
         TabIndex        =   60
         Top             =   240
         Width           =   3840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Dirujuk"
         Height          =   210
         Index           =   26
         Left            =   240
         TabIndex        =   59
         Top             =   840
         Width           =   930
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnosa (Penyakit) Rujukan"
         Height          =   210
         Index           =   22
         Left            =   6600
         TabIndex        =   58
         Top             =   840
         Width           =   2565
      End
   End
   Begin VB.Frame fraPemakaianAsuransi 
      Caption         =   "Pemakaian Asuransi  (SJP = Surat Jaminan Pelayanan)"
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
      TabIndex        =   50
      Top             =   4560
      Width           =   7335
      Begin VB.TextBox txtNoSJP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   18
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtNoKunjungan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         MaxLength       =   1
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtNoBP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         MaxLength       =   3
         TabIndex        =   20
         Text            =   "a24"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtAnakKe 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chkNoSJP 
         Caption         =   "No. SJP Otomatis"
         Enabled         =   0   'False
         Height          =   210
         Left            =   2880
         TabIndex        =   17
         Top             =   360
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo dcHubungan 
         Height          =   330
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSComCtl2.DTPicker dtpTglSJP 
         Height          =   315
         Left            =   5280
         TabIndex        =   19
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
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
         Format          =   163905539
         UpDown          =   -1  'True
         CurrentDate     =   37694
      End
      Begin MSDataListLib.DataCombo dcUnitKerja 
         Height          =   330
         Left            =   1920
         TabIndex        =   22
         Top             =   1200
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcKelasDitanggung 
         Height          =   330
         Left            =   5280
         TabIndex        =   23
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kelas Ditanggung"
         Height          =   210
         Index           =   13
         Left            =   5400
         TabIndex        =   70
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. SJP"
         Height          =   210
         Index           =   17
         Left            =   5280
         TabIndex        =   56
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hubungan Pasien"
         Height          =   210
         Index           =   15
         Left            =   240
         TabIndex        =   55
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anak Ke -"
         Height          =   210
         Index           =   16
         Left            =   1920
         TabIndex        =   54
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. BP"
         Height          =   210
         Index           =   18
         Left            =   240
         TabIndex        =   53
         Top             =   960
         Width           =   555
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kunj. Ke -"
         Height          =   210
         Index           =   20
         Left            =   960
         TabIndex        =   52
         Top             =   960
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bertugas di Unit / Bagian"
         Height          =   210
         Index           =   19
         Left            =   1920
         TabIndex        =   51
         Top             =   960
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   735
      Left            =   9480
      TabIndex        =   31
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   735
      Left            =   8280
      TabIndex        =   30
      Top             =   8760
      Width           =   1095
   End
   Begin VB.TextBox txtNamaFormPengirim 
      Height          =   495
      Left            =   0
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTglPendaftaran 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Frame fraDataKartuPeserta 
      Caption         =   "Data Kartu Peserta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   40
      Top             =   2160
      Width           =   10575
      Begin VB.OptionButton optRujukanRS 
         Caption         =   "Rujukan RS"
         Height          =   210
         Left            =   2280
         TabIndex        =   147
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optNoKartu 
         Caption         =   "No. Kartu"
         Height          =   210
         Left            =   120
         TabIndex        =   92
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optNoRujukan 
         Caption         =   "Rujukan"
         Height          =   210
         Left            =   1320
         TabIndex        =   91
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtJenisPasien 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   315
         Left            =   7920
         MaxLength       =   16
         ScrollBars      =   1  'Horizontal
         TabIndex        =   89
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CheckBox chkCheckkartu 
         Caption         =   "Check Online No Kartu Peserta"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   240
         TabIndex        =   74
         Top             =   240
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkDiriSendiri 
         Caption         =   "Diri Sendiri"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   9240
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtAlamatPA 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   14
         Top             =   1920
         Width           =   8175
      End
      Begin VB.TextBox txtNipPA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5520
         MaxLength       =   16
         ScrollBars      =   1  'Horizontal
         TabIndex        =   12
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtNamaPA 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtNoKartuPA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   315
         Left            =   240
         MaxLength       =   25
         TabIndex        =   9
         Top             =   720
         Width           =   3255
      End
      Begin MSDataListLib.DataCombo dcPenjamin 
         Height          =   330
         Left            =   3600
         TabIndex        =   7
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8454016
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
      Begin MSComCtl2.DTPicker dtpTglLahirPA 
         Height          =   315
         Left            =   3960
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   163840001
         UpDown          =   -1  'True
         CurrentDate     =   37694
      End
      Begin MSDataListLib.DataCombo dcPerusahaan 
         Height          =   330
         Left            =   6720
         TabIndex        =   8
         Top             =   720
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8454016
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
      Begin MSDataListLib.DataCombo dcGolonganAsuransi 
         Height          =   330
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Peserta (BPJS)"
         Height          =   210
         Index           =   32
         Left            =   7920
         TabIndex        =   90
         Top             =   1080
         Width           =   1665
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Golongan Asuransi"
         Height          =   210
         Index           =   7
         Left            =   240
         TabIndex        =   71
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Institusi Asal Pasien"
         Height          =   210
         Index           =   2
         Left            =   6720
         TabIndex        =   68
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Peserta"
         Height          =   210
         Index           =   14
         Left            =   2160
         TabIndex        =   46
         Top             =   1680
         Width           =   1230
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Lahir"
         Height          =   210
         Index           =   11
         Left            =   3960
         TabIndex        =   45
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. KTP / SIM Peserta"
         Height          =   210
         Index           =   12
         Left            =   5520
         TabIndex        =   44
         Top             =   1080
         Width           =   1605
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Peserta"
         Height          =   210
         Index           =   10
         Left            =   240
         TabIndex        =   43
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kartu Peserta"
         Height          =   210
         Index           =   9
         Left            =   240
         TabIndex        =   42
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Penjamin"
         Height          =   210
         Index           =   8
         Left            =   3600
         TabIndex        =   41
         Top             =   480
         Width           =   1365
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
      TabIndex        =   32
      Top             =   960
      Width           =   10575
      Begin VB.TextBox txtNoTlpPasien 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   0
         MaxLength       =   10
         TabIndex        =   82
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtppkpelayanan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   0
         MaxLength       =   10
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   75
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   65
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo dcJenisPasien 
         Height          =   315
         Left            =   8520
         TabIndex        =   5
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   0
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4560
         MaxLength       =   9
         TabIndex        =   1
         Top             =   600
         Width           =   1095
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
         Height          =   615
         Left            =   5880
         TabIndex        =   33
         Top             =   360
         Width           =   2535
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            MaxLength       =   6
            TabIndex        =   2
            Top             =   250
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            MaxLength       =   6
            TabIndex        =   3
            Top             =   250
            Width           =   375
         End
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   4
            Top             =   250
            Width           =   375
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "thn"
            Height          =   210
            Index           =   4
            Left            =   600
            TabIndex        =   36
            Top             =   302
            Width           =   285
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bln"
            Height          =   210
            Index           =   5
            Left            =   1440
            TabIndex        =   35
            Top             =   302
            Width           =   240
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "hr"
            Height          =   210
            Index           =   6
            Left            =   2280
            TabIndex        =   34
            Top             =   302
            Width           =   165
         End
      End
      Begin VB.Label lblNoPendaftaran 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   8520
         TabIndex        =   66
         Top             =   -120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblJenisPasien 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Pasien"
         Height          =   210
         Left            =   8520
         TabIndex        =   47
         Top             =   360
         Width           =   960
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. CM"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pasien"
         Height          =   210
         Index           =   1
         Left            =   1680
         TabIndex        =   38
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Index           =   3
         Left            =   4560
         TabIndex        =   37
         Top             =   360
         Width           =   1065
      End
   End
   Begin MSComctlLib.ProgressBar pbData 
      Height          =   360
      Left            =   120
      TabIndex        =   72
      Top             =   10080
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   200
      Scrolling       =   1
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   76
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
   Begin MSComCtl2.DTPicker dtpTglPulang 
      Height          =   315
      Left            =   720
      TabIndex        =   104
      Top             =   9960
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
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
      Format          =   163905539
      UpDown          =   -1  'True
      CurrentDate     =   37694
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmUbahJenisPasien.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmUbahJenisPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
        Dim result() As String
        Dim mstrPpkRujukan As String
        Dim mstrkartuPeserta As String
        Dim mstrnorujukan As String
        Dim mstrJenisPeserta As String
        Dim mstrNamaAsalRujukanMon As String
        Dim fKdKelasDitanggung As String
        Dim mstrNamaPeserta As String
        Dim strNoSEPForSimpan As String
        Dim mstrCatatatnBPJS As String
        Dim bolGenerateSEPSukse As Boolean
        Dim pisa As String
        Dim sex As String
        Dim kdCabang As String
        Dim nmCabang As String
        Dim kdKelas As String
        Dim kdJenisPeserta As String
        Dim potensiprb As String
        
        Dim n As Long
        Dim arr() As String
        
Dim fTglLahir As Date
Dim fNoPendaftaran As String
Dim fNoSJP As String
Dim fNoBP As String
Dim fNoKunjungan As Integer
Dim fChkNoSJP As String
Dim fDcUnitKerja As String
Dim fNamaAsalRujukan As String
Dim fNamaPerujuk As String
Dim fDiagnosa As String
Dim fAlamatPA As String
Dim fIDPeserta As String
Dim fKdPerusahaan As String
Dim ppkRujukan As String
Dim mstrKelasDitanggung As String

'Private ppkRujukan As String
Public mstrTglVerifBPJS As String
Public mbolSEP As Boolean
Dim mdtptglsjp As String
Public mstrPilihanSEP As String
Dim strJenisID As String

'---------- detailKartuBPJS------------------
Dim noKartu As String
Dim nik As String
Dim nama As String
'Dim pisa As String
'Dim sex As String
Dim tgllahir As String
Dim tglCetakKartu As String
Dim kdProvider As String
Dim nmProvider As String
Dim statusPeserta As String
'Dim kdCabang As String
'Dim nmCabang As String
'Dim kdJenisPeserta As String
Dim nmJenisPeserta As String
'Dim kdKelas As String
Dim nmKelas As String

Dim diagAwal As String
Dim strNamaPegawai As String
Dim jnsPelayanan As String
Dim asalRujukan As String
Dim keluhan As String
Dim keterangan As String
'----------------------------
Dim blnKartuAktif As Boolean
Dim strKdDiagnosa As String
Dim StatusVclaim As String

Dim lakalantas As String
Dim lokasiLaka As String
Dim penjaminLakalantas As String
Dim cob As String
Dim katarak As String
Dim jnsBtn As String
Dim suplesi As String

'Store procedure untuk mengisi struk billing pasien
Public Function sp_AddStrukBuktiKasMasuk() As Boolean
    On Error GoTo errLoad
    Dim strLokal As String
    sp_AddStrukBuktiKasMasuk = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglBKM", adDate, adParamInput, , Format(dTglDaftar, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdCaraBayar", adChar, adParamInput, 2, "01")
        .Parameters.Append .CreateParameter("KdJenisKartu", adChar, adParamInput, 2, Null)
        .Parameters.Append .CreateParameter("NamaBank", adVarChar, adParamInput, 100, Null)
        .Parameters.Append .CreateParameter("NoKartu", adVarChar, adParamInput, 50, Null)
        .Parameters.Append .CreateParameter("AtasNama", adVarChar, adParamInput, 50, Null)
        .Parameters.Append .CreateParameter("JmlBayar", adCurrency, adParamInput, , mcurAll_HrsDibyr)
        .Parameters.Append .CreateParameter("Administrasi", adCurrency, adParamInput, , 0)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, Null)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, "176")
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, noidpegawai)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("OutputNoBKM", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_StrukBuktiKasMasukPelayananPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("RETURN_VALUE").value <> 0 Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Struk Billing Pasien", vbCritical, "Validasi"
            sp_AddStrukBuktiKasMasuk = False
        Else
            If Not IsNull(.Parameters("OutputNoBKM").value) Then txtNoBKM.Text = .Parameters("OutputNoBKM").value
            If Len(Trim(txtNoBKM.Text)) = 0 Then
                strLokal = "SELECT NoBKM from StrukBuktiKasMasuk where tglBKM = '" & Format(dTglDaftar, "yyyy/MM/dd HH:mm:ss") & "' and idUser = '" & noidpegawai & "' and kdRuangan = '176'"
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
Public Function sp_AddStruk(ByVal adoCommand As ADODB.Command, strStsByr As String) As Boolean
    On Error GoTo errLoad
    Dim strLokal As String
    sp_AddStruk = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoBKM", adChar, adParamInput, 10, mstrNoBKM)
        .Parameters.Append .CreateParameter("OutputNoStruk", adChar, adParamOutput, 10, Null)
        .Parameters.Append .CreateParameter("TglStruk", adDate, adParamInput, , Format(dTglDaftar, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, mstrNoCM)
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, mstrKdJenisPasien)

        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, IIf(mstrKdPenjamin = "", "2222222222", mstrKdPenjamin))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, "176")
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, noidpegawai)
        .Parameters.Append .CreateParameter("TotalBiaya", adCurrency, adParamInput, , CCur(mcurBayar))
        .Parameters.Append .CreateParameter("JmlHutangPenjamin", adCurrency, adParamInput, , CCur(mcurAll_TP))
        .Parameters.Append .CreateParameter("JmlTanggunganRS", adCurrency, adParamInput, , CCur(mcurAll_TRS))
        .Parameters.Append .CreateParameter("JmlPembebasan", adCurrency, adParamInput, , CCur(mcurAll_Pemb))
        .Parameters.Append .CreateParameter("JmlHrsDibayar", adCurrency, adParamInput, , CCur(mcurAll_HrsDibyr))
        .Parameters.Append .CreateParameter("JmlDiscount", adCurrency, adParamInput, , "0")

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_NoStrukPelayananPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Struk Billing Pasien", vbCritical, "Validasi"
            sp_AddStruk = False
        Else
            If Not IsNull(.Parameters("OutputNoStruk").value) Then mstrNoStruk = .Parameters("OutputNoStruk").value
            If Len(mstrNoStruk) = 0 Then
                strLokal = "SELECT NoStruk from StrukPelayananPasien where tglStruk = '" & Format(dTglDaftar, "yyyy/MM/dd HH:mm:ss") & "' and NoPendaftaran = '" & mstrNoPen & "' and NoCM = '" & mstrNoCM & "' and idUser = '" & noidpegawai & "'"
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

Private Sub subLoadPemakaianAsuransi(s_NoPendaftaran As String, s_IdPenjamin As String)
    On Error GoTo errLoad

    strSQL = "SELECT * FROM v_PemakaianAsuransi WHERE NoPendaftaran = '" & s_NoPendaftaran & "' AND IdPenjamin='" & s_IdPenjamin & "' and StatusEnabled='1'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF = False Then
        dcHubungan.BoundText = IIf(IsNull(rs("KdHubungan")), "", rs("KdHubungan"))
        txtAnakKe.Text = IIf(IsNull(rs("AnakKe")), "", rs("AnakKe"))
        txtNoSJP.Text = IIf(IsNull(rs("NoSJP")), "", rs("NoSJP"))
        dtpTglSJP.value = IIf(IsNull(rs("TglSJP")), Now, rs("TglSJP"))
        txtNoBP.Text = IIf(IsNull(rs("NoBP")), "", rs("NoBP"))
        txtNoKunjungan.Text = IIf(IsNull(rs("KunjunganKe")), "", rs("KunjunganKe"))
        dcUnitKerja.Text = IIf(IsNull(rs("UnitBagian")), "", rs("UnitBagian"))
    Else
        dcHubungan.BoundText = ""
        txtAnakKe.Text = ""
        txtNoSJP.Text = ""
        dtpTglSJP.value = Now
        txtNoBP.Text = ""
        txtNoKunjungan.Text = ""
        dcUnitKerja.Text = ""
        dcKelasDitanggung.BoundText = ""
    End If

    Exit Sub
errLoad:
    Call msubPesanError("subLoadPemakaianAsuransi")
End Sub

Private Sub subTampungDataPenjamin()
    typAsuransi.strIdPenjamin = dcPenjamin.BoundText
    typAsuransi.strIdAsuransi = txtNoKartuPA.Text
    typAsuransi.strNoCm = txtNoCM.Text
    typAsuransi.strNamaPeserta = txtNamaPA.Text

    typAsuransi.strIdPeserta = IIf(txtNipPA.Text = "", "-", txtNipPA.Text)   ''allow null

    typAsuransi.strKdGolongan = dcGolonganAsuransi.BoundText
    typAsuransi.dTglLahir = dtpTglLahirPA.value
    typAsuransi.strAlamat = IIf(txtAlamatPA.Text = "", "-", txtAlamatPA.Text)
    typAsuransi.strNoPendaftaran = IIf(txtNoPendaftaran.Text <> "", txtNoPendaftaran.Text, mstrNoPen)

    typAsuransi.strHubungan = dcHubungan.BoundText
    typAsuransi.strNoSJP = txtNoSJP.Text
    typAsuransi.dTglSJP = dtpTglSJP.value
    typAsuransi.strNoBp = IIf(txtNoBP.Text = "", "-", txtNoBP.Text)
    typAsuransi.intNoKunjungan = IIf(Val(txtNoKunjungan.Text) = 0, 1, Val(txtNoKunjungan.Text))

    typAsuransi.strStatusNoSJP = IIf(chkNoSJP.value = vbChecked, "O", "M")
    typAsuransi.intAnakKe = IIf(Val(txtAnakKe.Text) = 0, 0, Val(txtAnakKe.Text))
    typAsuransi.strUnitBagian = IIf(dcUnitKerja.Text = "", "-", Trim(dcUnitKerja.Text))

    typAsuransi.strNoRujukan = IIf(txtNoRujukan.Text = "", "-", txtNoRujukan.Text)
    typAsuransi.strKdRujukanAsal = dcAsalRujukan.BoundText
    typAsuransi.strDetailRujukanAsal = IIf(dcNamaAsalRujukan.Text = "", "-", dcNamaAsalRujukan.Text)
    typAsuransi.strKdDetailRujukanAsal = IIf(ppkRujukan <> "", ppkRujukan, dcNamaAsalRujukan.BoundText)
    typAsuransi.strNamaPerujuk = IIf(dcNamaPerujuk.Text = "", "-", dcNamaPerujuk.Text)

    typAsuransi.dTglDirujuk = dtpTglDirujuk.value
    typAsuransi.strDiagnosaRujukan = IIf(dcDiagnosa.Text = "", "-", dcDiagnosa.Text)
    typAsuransi.strKdDiagnosa = dcDiagnosa.BoundText

    typAsuransi.strKdKelompokPasien = dcJenisPasien.BoundText
    typAsuransi.strPerusahaanPenjamin = dcPerusahaan.BoundText
    typAsuransi.strKdKelasDitanggung = dcKelasDitanggung.BoundText
'    typAsuransi.strJenisPasien = txtJenisPasien.Text

    typAsuransi.blnSuksesAsuransi = True
End Sub

Private Function AUD_DetailKartuBPJS() As Boolean
    On Error GoTo StatusErr

    MousePointer = vbHourglass
    AUD_DetailKartuBPJS = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("noKartu", adChar, adParamInput, 13, txtNoKartuPA)
        .Parameters.Append .CreateParameter("nik", adChar, adParamInput, 16, txtNipPA)
        .Parameters.Append .CreateParameter("nama", adVarChar, adParamInput, 50, txtNamaPA)
        .Parameters.Append .CreateParameter("pisa", adVarChar, adParamInput, 3, pisa)
        .Parameters.Append .CreateParameter("sex", adVarChar, adParamInput, 3, sex)
        .Parameters.Append .CreateParameter("tglLahir", adVarChar, adParamInput, 20, Format(dtpTglLahirPA, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("tglCetakKartu", adVarChar, adParamInput, 20, Format(dtpTglSJP, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("kdProvider", adVarChar, adParamInput, 20, ppkRujukan)
        .Parameters.Append .CreateParameter("nmProvider", adVarChar, adParamInput, 50, dcNamaAsalRujukan.Text)
        .Parameters.Append .CreateParameter("kdCabang", adVarChar, adParamInput, 20, kdCabang)
        .Parameters.Append .CreateParameter("nmCabang", adVarChar, adParamInput, 50, nmCabang)
        .Parameters.Append .CreateParameter("kdJenisPeserta", adVarChar, adParamInput, 20, kdJenisPeserta)
        .Parameters.Append .CreateParameter("nmJenisPeserta", adVarChar, adParamInput, 50, txtJenisPasien.Text)
        .Parameters.Append .CreateParameter("kdKelas", adVarChar, adParamInput, 10, kdKelas)
        .Parameters.Append .CreateParameter("nmKelas", adVarChar, adParamInput, 20, dcKelasDitanggung.Text)
    End With
        Exit Function
StatusErr:
    cmdSimpan.Enabled = True
    MousePointer = vbDefault
'    sp_JenisPasienJoinProgramAskes = False
    Call msubPesanError("sp_JenisPasienJoinProgramAskes")
    MsgBox "Ulangi proses simpan ", vbCritical, "Validasi"
End Function

Private Function sp_JenisPasienJoinProgramAskes() As Boolean
    On Error GoTo StatusErr

    MousePointer = vbHourglass
    sp_JenisPasienJoinProgramAskes = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, dcPenjamin.BoundText)
        .Parameters.Append .CreateParameter("IdAsuransi", adVarChar, adParamInput, 25, txtNoKartuPA)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM)
        .Parameters.Append .CreateParameter("KdHubKeluarga", adChar, adParamInput, 2, dcHubungan.BoundText)
        .Parameters.Append .CreateParameter("NamaPeserta", adVarChar, adParamInput, 50, txtNamaPA.Text)
        '5
        .Parameters.Append .CreateParameter("IDPeserta", adVarChar, adParamInput, 16, txtNipPA)
        .Parameters.Append .CreateParameter("KdGolongan", adChar, adParamInput, 2, IIf(Len(Trim(dcGolonganAsuransi.Text)) = 0, Null, Trim(dcGolonganAsuransi.BoundText)))
        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(dtpTglLahirPA, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Alamat", adVarChar, adParamInput, 100, txtAlamatPA)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, IIf(Len(Trim(txtNoPendaftaran.Text)) = 0, Null, txtNoPendaftaran.Text))
        '10
        .Parameters.Append .CreateParameter("KdHubungan", adChar, adParamInput, 2, dcHubungan.BoundText)
        .Parameters.Append .CreateParameter("NoSJP", adVarChar, adParamInput, 30, IIf(Len(Trim(txtNoSJP.Text)) = 0, Null, Trim(txtNoSJP.Text)))
        .Parameters.Append .CreateParameter("TglSJP", adDate, adParamInput, , Format(dtpTglSJP, "yyyy/MM/dd hh:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("NoBP", adChar, adParamInput, 3, IIf(Len(Trim(txtNoBP.Text)) = 0, Null, Trim(txtNoBP.Text)))
        '15
        .Parameters.Append .CreateParameter("KunjunganKe", adInteger, adParamInput, , IIf(Val(txtNoKunjungan.Text) = 0, "1", txtNoKunjungan.Text))
        .Parameters.Append .CreateParameter("OutputNoSJP", adVarChar, adParamOutput, 30, Null)
        .Parameters.Append .CreateParameter("StatusNoSJP", adChar, adParamInput, 1, IIf(chkNoSJP.value = vbChecked, "O", "M"))
        .Parameters.Append .CreateParameter("AnakKe", adInteger, adParamInput, , Val(txtAnakKe.Text))
        .Parameters.Append .CreateParameter("UnitBagian", adVarChar, adParamInput, 50, IIf(Len(Trim(dcUnitKerja.Text)) = 0, Null, Trim(dcUnitKerja.Text)))
        '20
        .Parameters.Append .CreateParameter("KdPaket", adVarChar, adParamInput, 3, Null)
        .Parameters.Append .CreateParameter("NoRujukan", adVarChar, adParamInput, 30, txtNoRujukan.Text)
        .Parameters.Append .CreateParameter("KdRujukanAsal", adChar, adParamInput, 2, dcAsalRujukan.BoundText)
        .Parameters.Append .CreateParameter("DetailRujukanAsal", adVarChar, adParamInput, 100, IIf(Len(Trim(dcNamaAsalRujukan.Text)) = 0, Null, dcNamaAsalRujukan.Text))
'        .Parameters.Append .CreateParameter("KdDetailRujukanAsal", adChar, adParamInput, 8, IIf(chkNoSJP.value = vbChecked, "12345678", dcNamaAsalRujukan.BoundText))
        .Parameters.Append .CreateParameter("KdDetailRujukanAsal", adChar, adParamInput, 8, IIf(ppkRujukan <> "", ppkRujukan, dcNamaAsalRujukan.BoundText))
        '25
        .Parameters.Append .CreateParameter("NamaPerujuk", adVarChar, adParamInput, 50, IIf(Len(Trim(dcNamaPerujuk.Text)) = 0, Null, Trim(dcNamaPerujuk.Text)))
        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , Format(dtpTglDirujuk.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("DiagnosaRujukan", adVarChar, adParamInput, 100, IIf(Len(Trim(dcDiagnosa.Text)) = 0, Null, Trim(dcDiagnosa.Text)))
        .Parameters.Append .CreateParameter("KdDiagnosa", adVarChar, adParamInput, 7, dcDiagnosa.BoundText)
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, dcJenisPasien.BoundText)
        .Parameters.Append .CreateParameter("KdInstitusiAsal", adVarChar, adParamInput, 4, IIf(dcPerusahaan.Text = "", Null, dcPerusahaan.BoundText))
        .Parameters.Append .CreateParameter("KdKelasDiTanggung", adChar, adParamInput, 2, dcKelasDitanggung.BoundText)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_JenisPasienJoinProgramAskesNew"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_JenisPasienJoinProgramAskes = False
        Else
            txtNoSJP.Text = IIf(IsNull(.Parameters("OutputNoSJP")), "", .Parameters("OutputNoSJP"))
            cmdSimpan.Enabled = False
            Call Add_HistoryLoginActivity("Update_JenisPasienJoinProgramAskesNew")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    MousePointer = vbDefault

    Exit Function
StatusErr:
    cmdSimpan.Enabled = True
    MousePointer = vbDefault
    sp_JenisPasienJoinProgramAskes = False
    Call msubPesanError("sp_JenisPasienJoinProgramAskes")
    MsgBox "Ulangi proses simpan ", vbCritical, "Validasi"
End Function

Private Sub subKosong()
    txtNoCM.Text = ""
    txtNamaPasien.Text = ""
    txtJK.Text = ""
    txtThn.Text = ""
    txtBln.Text = ""
    txtHr.Text = ""
    txtNoPendaftaran.Text = ""
    txtFormRegistrasiPengirim.Text = ""

    chkDiriSendiri.value = vbUnchecked

    dcPenjamin.BoundText = ""
    txtNoKartuPA.Text = ""
    txtNamaPA.Text = ""
    dtpTglLahirPA.value = Now
    txtNipPA.Text = ""
    dcKelasDitanggung.BoundText = ""
    txtAlamatPA.Text = ""

    dcHubungan.BoundText = ""
    txtAnakKe.Text = ""
    chkNoSJP.value = vbUnchecked
    dtpTglSJP.value = Now
    txtNoBP.Text = ""
    txtNoKunjungan.Text = ""
    dcUnitKerja.BoundText = ""

    dcAsalRujukan.BoundText = ""
    txtNoRujukan.Text = ""
    dcNamaAsalRujukan.BoundText = ""
    dtpTglDirujuk.value = Now
    dcNamaPerujuk.BoundText = ""
    dcDiagnosa.BoundText = ""
    dcGolonganAsuransi.BoundText = ""
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    Call msubDcSource(dcJenisPasien, rs, "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where StatusEnabled='1' order by JenisPasien")
    Call msubDcSource(dcHubungan, rs, "SELECT KdHubungan, NamaHubungan FROM HubunganPesertaAsuransi where StatusEnabled='1'")
    Call msubDcSource(dcGolonganAsuransi, rs, "SELECT     KdGolongan, NamaGolongan FROM GolonganAsuransi where StatusEnabled='1'")
    Call msubDcSource(dcUnitKerja, rs, "SELECT KdRuangan, NamaRuangan FROM Ruangan where StatusEnabled='1' ORDER BY NamaRuangan")
    Call msubDcSource(dcAsalRujukan, rs, "SELECT KdRujukanAsal, RujukanAsal FROM RujukanAsal where StatusEnabled='1'")
    strSQL = "SELECT KdDetailRujukanAsal, DetailRujukanAsal" & _
    " FROM DetailRujukanAsal " & _
    " WHERE (KdRujukanAsal = '" & dcAsalRujukan.BoundText & "')"
    Call msubDcSource(dcNamaAsalRujukan, rs, strSQL)
    Call msubDcSource(dcNamaPerujuk, rs, "SELECT KodeDokter, NamaDokter FROM V_DaftarDokter")
    Call msubDcSource(dcDiagnosa, rs, "SELECT KdDiagnosa, NamaDiagnosa FROM Diagnosa where StatusEnabled='1' ORDER BY NamaDiagnosa")

    strSQL = "SELECT  KdInstitusiAsal, InstitusiAsal FROM InstitusiAsalPasien where StatusEnabled='1' order by InstitusiAsal"
    Call msubDcSource(dcPerusahaan, rs, strSQL)
'    Call msubDcSource(dcLakalantas, rs, "SELECT kdLakalantas, nmLakalantas FROM Lakalantas")
    Call msubDcSource(dcKelasDitanggung, rs, "SELECT DISTINCT KdKelas, DeskKelas FROM V_KelasDitanggungPenjamin ")
    
    Exit Sub
errLoad:
    Set rs = Nothing
    Call msubPesanError
End Sub

Private Sub chkDiriSendiri_Click()
    On Error GoTo errLoad
    If chkDiriSendiri.value = 1 Then
        strSQL = "SELECT NamaLengkap, NoIdentitas, Alamat,TglLahir FROM v_S_RegistrasiDataPasien WHERE NocM='" & txtNoCM.Text & "'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount <> 0 Then
            txtNamaPA.Text = rs("NamaLengkap")
            txtNipPA.Text = rs("NoIdentitas") & ""
            txtAlamatPA.Text = rs("Alamat") & ""
            dtpTglLahirPA.value = Format(rs("TglLahir"), "dd/mm/yyyy")
            dcHubungan.Text = "Peserta"
            
'            strSQL = "select * from asuransipasien where nocm='" & txtNoCM.Text & "'"
'            Call msubRecFO(rs, strSQL)
'            If (rs.RecordCount <> 0) Then
'                txtNoKartuPA.Text = rs("IdAsuransi").value
'            End If
'        Call msubRecFO(rs, strSQL)
        
        Else
            txtNamaPA.Text = ""
            txtNipPA.Text = ""
          '  txtNoKartuPA.Text = ""
            txtAlamatPA.Text = ""
            dtpTglLahirPA.value = Now
            dcHubungan.Text = ""
        End If
    Else
        txtNamaPA.Text = ""
        txtNipPA.Text = ""
       ' txtNoKartuPA.Text = ""
        txtAlamatPA.Text = ""
        dcHubungan.Text = ""
        dtpTglLahirPA.value = Now
    End If
    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub chkDiriSendiri_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then txtNoKartuPA.SetFocus
    If KeyAscii = 13 Then dcGolonganAsuransi.SetFocus
End Sub

Private Sub chkLakalantas_Click()
    With fraLakalantas
        If chkLakalantas.value = 1 Then
            .Visible = True
            .Left = 2040
            .Top = 2880
        Else
            .Visible = False
        End If
    End With
End Sub

Private Sub chkNoSJP_Click()
    If chkNoSJP.value = vbChecked Then txtNoSJP.Enabled = False Else txtNoSJP.Enabled = True
End Sub

Private Sub chkNoSJP_KeyPress(KeyAscii As Integer)
    If chkDiriSendiri.value = vbChecked Then dtpTglSJP.SetFocus Else txtNoSJP.SetFocus
End Sub

Private Sub chkSuplesi_Click()
    If chkSuplesi.value = vbChecked Then txtSuplesi.Enabled = True Else txtSuplesi.Enabled = False
End Sub

Private Sub cmdApprovalSEP_Click()
On Error GoTo pesan
    
'    Call msubRecFO(dbrs, "SELECT Value FROM SettingGlobal WHERE Prefix='KdJenisPasienBPJS'")
'    If Not dbrs.EOF Then
        If dcPenjamin.BoundText = "0000000019" Then
            
            If MsgBox("Yakin Anda Akan Setujui Pembuatan SEP " & vbCrLf & "atas nama " & txtNamaPasien.Text & " ?", vbQuestion + vbYesNo, "Approval Pengajuan SEP") = vbNo Then Exit Sub
            If ApprovalPengajuanSEP(Trim(txtNoKartuPA.Text)) = False Then Exit Sub
            If sp_PengajuanSEPVclaim("Disetujui") = False Then Exit Sub
        Else
            MsgBox "Fitur Ini hanya untuk pasien BPJS", vbExclamation, "Validasi"
        End If
'    End If
    
Exit Sub
pesan:
    Call msubPesanError
End Sub

Private Sub cmdHapusSEP_Click()
On Error Resume Next
    
    Call msubRecFO(dbRst, "SELECT IdPenjamin FROM PenjaminKelompokPasien WHERE KdKelompokPasien = '" & dcJenisPasien.BoundText & "'")
'    If dbRst.EOF = False Then
'        Call msubRecFO(rs3, "select * from SettingGlobal where prefix='KdJenisPasienBPJS' and value='" & dbRst(0).value & "'")
'        If rs3.EOF = True Then
'            MsgBox "Fitur ini hanya diperuntukan untuk Pasien BPJS", vbInformation + vbOKOnly, "Hapus SEP"
'            Exit Sub
'        End If
'    Else
'            MsgBox "Fitur ini hanya diperuntukan untuk Pasien BPJS", vbInformation + vbOKOnly, "Hapus SEP"
'            Exit Sub
'    End If
    If dcJenisPasien.BoundText <> "10" Then
        MsgBox "Fitur ini hanya diperuntukan untuk Pasien BPJS", vbInformation + vbOKOnly, "Hapus SEP"
        Exit Sub
    End If
    
    If Len(Trim(txtNoSJP.Text)) < 5 Then MsgBox "Silakan cek kembali pasien No SEP", vbInformation + vbOKOnly, "Hapus SEP": Exit Sub
    If MsgBox("Apakah anda yakin akan menghapus SEP?", vbQuestion + vbYesNo, "Hapus SEP") = vbNo Then Exit Sub
    fraHapusSEP.Visible = True
    txtKetHapus.SetFocus
End Sub

Private Sub cmdKec_Click()
jnsBtn = "Kecamatan"
MSFlexGrid1.Visible = True
MSFlexGrid1.Top = 3000
MSFlexGrid1.Left = 1200

On Error Resume Next

    If (Dir("C:\SDK\vclaim\result.tlb") <> "") Then
        Dim context As ContextVclaim
        Set context = New ContextVclaim
        
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
        
        result = context.RefKecamatan(Trim(txtKdKota.Text))
        Dim strResult As String
        strResult = ""
        For n = LBound(result) To UBound(result)
            arr = Split(result(n), ":")
            strResult = strResult & vbCrLf & result(n)
            If arr(0) = "error" Then
                MsgBox Replace(result(0), "error:", ""), vbExclamation, "error"
                Exit Sub
            End If
            
            
        Next n
        Call fillGridWithPropinsi(MSFlexGrid1, result)
    End If
End Sub

Private Sub cmdKota_Click()
jnsBtn = "Kota"
MSFlexGrid1.Visible = True
MSFlexGrid1.Top = 2520
MSFlexGrid1.Left = 1200

On Error Resume Next

    If (Dir("C:\SDK\vclaim\result.tlb") <> "") Then
        Dim context As ContextVclaim
        Set context = New ContextVclaim
        
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
        
        result = context.RefKotaKabupaten(Trim(txtKdPropinsi.Text))
        Dim strResult As String
        strResult = ""
        For n = LBound(result) To UBound(result)
            arr = Split(result(n), ":")
            strResult = strResult & vbCrLf & result(n)
            If arr(0) = "error" Then
                MsgBox Replace(result(0), "error:", ""), vbExclamation, "error"
                Exit Sub
            End If
            
            
        Next n
        Call fillGridWithPropinsi(MSFlexGrid1, result)
    End If
End Sub

Private Sub cmdLPK_Click()
On Error GoTo pesan
    
    Call msubRecFO(dbrs, "SELECT Value FROM SettingGlobal WHERE Prefix='KdJenisPasienBPJS'")
    If Not dbrs.EOF Then
        If dcPenjamin.BoundText = dbrs(0).value Then
            Me.Enabled = False
            
        Else
            MsgBox "Fitur Ini hanya untuk pasien BPJS", vbExclamation, "Validasi"
        End If
    End If
    
Exit Sub
pesan:
    Call msubPesanError
End Sub

Private Sub cmdPengajuanSEP_Click()
'    Call msubRecFO(dbrs, "SELECT Value FROM SettingGlobal WHERE Prefix='KdJenisPasienBPJS'")
'    If Not dbrs.EOF Then
        If dcPenjamin.BoundText = "0000000019" Then
            
            If Len(Trim(txtNoSJP.Text)) < 19 Then
                MsgBox "Silakan verifikasi NoSEP aktif", vbExclamation, "Pengajuan SEP BPJS"
            End If
            
            fraPengajuanSEP.Visible = True
            txtKetPengajuan.SetFocus
        Else
            MsgBox "Fitur Ini hanya untuk pasien BPJS", vbExclamation, "Validasi"
        End If
'    End If
    
Exit Sub
pesan:
    Call msubPesanError
End Sub

Private Sub cmdPropinsi_Click()
jnsBtn = "Propinsi"
MSFlexGrid1.Visible = True
MSFlexGrid1.Top = 2040
MSFlexGrid1.Left = 1200

On Error Resume Next

    If (Dir("C:\SDK\vclaim\result.tlb") <> "") Then
        Dim context As ContextVclaim
        Set context = New ContextVclaim
        
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
        
        result = context.RefPropinsi
        Dim strResult As String
        strResult = ""
        For n = LBound(result) To UBound(result)
            arr = Split(result(n), ":")
            strResult = strResult & vbCrLf & result(n)
            If arr(0) = "error" Then
                MsgBox Replace(result(0), "error:", ""), vbExclamation, "error"
                Exit Sub
            End If
            
            
        Next n
        Call fillGridWithPropinsi(MSFlexGrid1, result)
    End If
End Sub

Private Sub cmdRujukan_Click()
On Error GoTo pesan
    
'    Call msubRecFO(dbrs, "SELECT Value FROM SettingGlobal WHERE Prefix='KdJenisPasienBPJS'")
'    If Not dbrs.EOF Then
        If dcPenjamin.BoundText = "0000000019" Then
'            Me.Enabled = False
'            Call subLoadDataPasien(Trim(txtNoPendaftaran.Text), Trim(txtNoSJP.Text))

            Call subLoadDataRujukan(Trim(txtNoPendaftaran.Text), Trim(txtNoSJP.Text))
        Else
            MsgBox "Fitur Ini hanya untuk pasien BPJS", vbExclamation, "Validasi"
        End If
'    End If
    
Exit Sub
pesan:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    
    Dim context As ServiceAskes.context
    Set context = New context

    Call msubRecFO(dbRst, "SELECT IdPenjamin FROM PenjaminKelompokPasien WHERE KdKelompokPasien = '" & dcJenisPasien.BoundText & "'")
    If dbRst(0).value = "2222222222" Then
        'begin pasien jaminan
        Set rs = Nothing
        strSQL = "SELECT IdPenjamin FROM PasienDaftar WHERE NoPendaftaran = '" & txtNoPendaftaran.Text & "'"
        Call msubRecFO(rs, strSQL)
        If rs(0).value = "0000000025" Then
            If sp_UpdateJenisPasienJaminan(dcJenisPasien.BoundText, txtNoPendaftaran.Text) = False Then Exit Sub
        Else
            If sp_UpdateJenisPasienUmum(dcJenisPasien.BoundText, txtNoPendaftaran.Text) = False Then Exit Sub
        End If
        MousePointer = vbHourglass

        MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
        cmdSimpan.Enabled = False
        MousePointer = vbDefault
        Exit Sub
    End If
    
    If dcPenjamin.BoundText = "0000000019" Then
        If Periksa("text", txtKdDPJP, "DPJP belum di isi") = False Then Exit Sub
    End If
    
    If Periksa("datacombo", dcPenjamin, "Penjamin belum di isi") = False Then Exit Sub
    If Periksa("text", txtNoKartuPA, "Nomor kartu belum di isi") = False Then Exit Sub
    If Periksa("text", txtNamaPA, "Nama peserta asuransi ?") = False Then Exit Sub
    If Periksa("datacombo", dcKelasDitanggung, "Kelas ditanggung ?") = False Then Exit Sub
    If Periksa("datacombo", dcHubungan, "Hubungan peserta asuransi belum di isi") = False Then Exit Sub
    'If Periksa("text", txtNoSJP, "No SJP harus diisi") = False Then Exit Sub
    If Periksa("datacombo", dcAsalRujukan, "Asal rujukan belum di isi") = False Then Exit Sub
    If Periksa("text", txtNoRujukan, "No rujukan belum di isi") = False Then Exit Sub
    If Periksa("datacombo", dcGolonganAsuransi, "Golongan Asuransi belum di isi") = False Then Exit Sub
    If dcJenisPasien.BoundText = "10" Then
        If Periksa("datacombo", dcDiagnosa, "Diagnosa Belum di isi") = False Then Exit Sub
        If dtpTglDirujuk.value = dtpTglSJP.value Then MsgBox "Tgl dirujuk belum sesuai", vbInformation, "Informasi": Exit Sub
    End If
    mstrNoSJP = txtNoSJP.Text
    
    '---untuk generate no SEP BPJS----------------------------------------------
    Call msubRecFO(dbRst3, "Select * from PemakaianAsuransi where NoCM='" & Trim(txtNoCM.Text) & "' and NoPendaftaran ='" & mstrNoPen & "'")
  If dbRst3.EOF = True Then
TarikSEP:

    mdtptglsjp = Format(dtpTglDirujuk, "yyyy-MM-dd hh:mm:ss")
    If dcPenjamin.BoundText = "0000000019" Then
        If chkNoSJP.value = vbChecked Then
            'Untuk Mendapatkan SEP ke Database BPZS Bila = T Tidak connect, Bila <> T Connect
            If txtNoSJP.Text = "0" Then
            
                'Call sp_JenisPasienJoinProgramBPJS
            Else
              
            strSQL = "SELECT value FROM SettingGlobal where prefix='SepGenerateBPJS'"
            Call msubRecFO(rs, strSQL)

                If rs.EOF = False Then
                    mstrPilihanSEP = rs(0)
                End If
                mstrNoSJP = ""
                If mstrPilihanSEP = "Y" Then
                    strSQL = "SELECT Value FROM SettingGlobal WHERE Prefix='Versi VClaim'"
                    Call msubRecFO(rs, strSQL)
                    If rs(0).value = "1.0" Then
                        Call GenerateSEPBPJS
                    ElseIf rs(0).value = "1.1" Then
                        Call GenerateSEPBPJSNew
                    End If
                Else
'                  Call sp_JenisPasienJoinProgramBPJS
                End If
                
                If (Mid(mstrNoSJP, 1, 8) = context.KodeRumahSakit) Then
                    MsgBox "No SEP = " & mstrNoSJP & ""
                    'Exit Sub
                End If
                If mbolSEP = True Then
                    Exit Sub
                End If
                
            End If
        Else
                'Call sp_JenisPasienJoinProgramBPJS
        End If
    End If
  Else
    If txtNoSJP.Text = "Error" Or txtNoSJP.Text = "" Or txtNoSJP.Text = "-" Or Len(txtNoSJP.Text) = 19 Then GoTo TarikSEP:
    txtNoSJP.Text = dbRst3("NoSJP").value
        dtpTglSJP.value = dbRst3("TglSJP").value
'        Call msubDcSource(dcKelasDitanggung, rs, "SELECT KdKelas, DeskKelas FROM KelasPelayanan where KdKelas='" & dbRst3("KdKelasDitanggung").value & "'")
'        dcKelasDitanggung.Text = rs("DeskKelas").value
   bolGenerateSEPSukse = False

  End If
    mstrNoSJP = txtNoSJP.Text
    '-------------------------------------------------------------------
    If bolGenerateSEPSukse = False Then MsgBox "gagal simpan", vbInformation, "Informasi": Exit Sub
    ' chandra 27 02 2014
    ' Tambahan untuk txtNamaFormPengirim.Text = "Tampung" karena ada parameter yang T nya besar
    If txtNamaFormPengirim.Text = "tampung" Or txtNamaFormPengirim.Text = "Tampung" Then
        Call subTampungDataPenjamin
    Else
        If sp_JenisPasienJoinProgramAskes = False Then Exit Sub
    End If
    
    If typAsuransi.blnSuksesAsuransi = True Then cmdSimpan.Enabled = False Else cmdSimpan.Enabled = True

    MousePointer = vbDefault
    MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
    Exit Sub
errLoad:
    Call msubPesanError("cmdSimpan_Click")
    MousePointer = vbDefault
End Sub

Private Sub cmdSimpanKetHapus_Click()
On Error Resume Next
If Periksa("text", txtKetHapus, "") = False Then
    MsgBox "Silakan isi keterangan penyebab hapus SEP", vbCritical, "Hapus SEP"
    Exit Sub
End If

'If StatusVclaim = "Y" Then
    Call HapusSEPVclaim
'Else
'    Call HapusSEPV21
'End If

fraHapusSEP.Visible = False
End Sub

Private Sub cmdSimpanLaka_Click()
    If chkJasaRaharja.value = False And chkBPJSKK.value = False And chkTaspen.value = False And chkAsabri.value = False Then
        MsgBox "Silahkan pilih penjamin lakalantas dahulu...", vbCritical, "Peringatan"
        Exit Sub
    End If
    If txtKdPropinsi.Text = "" Or txtKota.Text = "" Or txtKec.Text = "" Then
        MsgBox "Silahkan isi lokasi lakalantas dahulu...", vbCritical, "Peringatan"
        Exit Sub
    End If
    fraLakalantas.Visible = False
End Sub

Private Sub cmdSimpanPengajuanSEP_Click()
On Error GoTo pesan
    
    If Periksa("text", txtKetPengajuan, "Keterangan Pengajuan SEP belum diisi") = False Then Exit Sub
    If PengajuanSEP(Trim(txtNoKartuPA.Text)) = False Then Exit Sub
    If sp_PengajuanSEPVclaim("Pengajuan") = False Then Exit Sub
    
Exit Sub
pesan:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    If mblnTemp = True Then
'            Unload Me
        If txtFormRegistrasiPengirim.Text = "frmRegistrasiAll" Then
            frmRegistrasiAll.Enabled = True
            Unload Me
        Else
            frmRegistrasiRJPenunjang.Enabled = True
            Unload Me
        End If
            mblnTemp = False
    Else
        Unload Me
    End If
End Sub

Private Sub cmdUpdateSEP_Click()
If chkNoSJP.value = vbChecked Then
             
          If (Dir("C:\SDK\Vclaim\result.tlb") <> "") Then
            Dim context As ContextVclaim
            Set context = New ContextVclaim
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
                 'Exit Sub
              End If
            
            mstrnorujukan = txtNoRujukan.Text
           
           Dim kodeKelas As Integer
           If (dcKelasDitanggung.Text = "Kelas I") Then
            kodeKelas = 1
           ElseIf (dcKelasDitanggung.Text = "Kelas II") Then
            kodeKelas = 2
           Else
            kodeKelas = 3
           End If
           
            strKdDiagnosa = Trim(dcDiagnosa.BoundText)
           
            Dim rsDataPoliBPJS As New ADODB.recordset
            Dim qryPoliBPJS As String
            Dim strKdDataPoliBPJS As String
            
            ''''''''' langsung ngambil kdruangan dri RS ''''''''''
            qryPoliBPJS = "SELECT KodeExternal FROM dbo.Ruangan WHERE KdRuangan='" & mstrKdRuanganPasien & "'"
            Call msubRecFO(rsDataPoliBPJS, qryPoliBPJS)
            If rsDataPoliBPJS.EOF = False Then
                strKdDataPoliBPJS = IIf(IsNull(rsDataPoliBPJS(0)) = True, "", rsDataPoliBPJS(0))
            Else
                strKdDataPoliBPJS = "" 'nilai default kalau belum dimapping
            End If
            
            '2+2=5, untuk testing generate ke DVLP IDPegawai panjangnnya tidak boleh lebih dari 9. Dipatok saja ke 888888888(9 Digit)
            Dim strIDPegawaiLokal As String
            strIDPegawaiLokal = strIDPegawai
            If URL = "http://dvlp.bpjs-kesehatan.go.id:8081/devWSLokalRest/" Then
                strIDPegawaiLokal = "88888888"
            End If
            
            strSQL = "SELECT Value FROM SettingGlobal WHERE Prefix='KodeRS'"
            Call msubRecFO(rs, strSQL)
            txtppkpelayanan.Text = rs(0).value
            
            strSQL = "SELECT Telepon FROM Pasien WHERE NoCM='" & txtNoCM.Text & "'"
            Call msubRecFO(rs, strSQL)
            txtNoTlpPasien.Text = IIf(IsNull(rs(0).value), "", Trim(rs(0).value))
            
'            strSQL = "SELECT dbo.getNamaPegawaiByIdPegawai('" & strIDPegawai & "')"
            strSQL = "SELECT NamaLengkap FROM DataPegawai WHERE IdPegawai='" & strIDPegawai & "'"
            Call msubRecFO(rs, strSQL)
            strNamaPegawai = rs(0).value
'            Diganti karena mstrNoSJPNew Harus String Array Ales
            
'            mstrNoSJPNew = context.UpdateSep(txtNoSJP.Text, kodeKelas, txtNoCM.Text, 1, Format(dtpTglDirujuk.value, "yyyy-mm-dd"), _
'                           txtNoRujukan.Text, ppkRujukan, txtCatatan.Text, strKdDiagnosa, IIf(mstrKdInstalasi, strKdDataPoliBPJS, ""), _
'                           0, 0, 0, 2, txtLokasiLakaLantas.Text, txtNoTlpPasien.Text, strNmPegawai)
                        
            Dim Temp As String
            Dim i As Integer
            For i = LBound(mstrNoSJPNew) To UBound(mstrNoSJPNew)
                Dim arr() As String
                arr = Split(mstrNoSJPNew(i), ":")
                Select Case arr(0)

                    Case txtNoSJP.Text
                       txtNoSJP.Text = arr(0)
                       Debug.Print "noSep : " & arr(0)
                       MsgBox "Update SEP Sukses"
                    Case "error"
                        MsgBox mstrNoSJPNew(i)
                        MsgBox "Generate SEP Gagal,,,,!!!"
                End Select
             Next i
                    
        Else
            'sp_JenisPasienJoinProgramBPJS
            MsgBox "Generate SEP Gagal,,,,!!!"
            typAsuransi.blnSuksesAsuransi = False
            bolGenerateSEPSukse = False
            txtNoSJP.Text = ""
        End If
End If
Exit Sub
SepEndPoint:
'     sp_JenisPasienJoinProgramBPJS
End Sub

Private Sub cmdUpdateTglPulangBPJS_Click()
On Error GoTo pesan
    Call msubRecFO(dbrs, "SELECT Value FROM SettingGlobal WHERE Prefix='KdJenisPasienBPJS'")
    If Not dbrs.EOF Then
        If dcPenjamin.BoundText = dbrs(0).value Then
        
            If Len(Trim(txtNoSJP.Text)) < 19 Then
                MsgBox "Silakan verifikasi NoSEP aktif", vbExclamation, "Validasi"
                Exit Sub
            End If
            
            If UpdateTglPulang(txtNoSJP.Text) = False Then Exit Sub
        Else
            MsgBox "Fitur Ini hanya untuk pasien BPJS", vbExclamation, "Validasi"
        End If
    End If
    
Exit Sub
pesan:
    Call msubPesanError
End Sub

Private Sub Command1_Click()
    chkLakalantas.value = False
    fraLakalantas.Visible = False
End Sub

Private Sub dcAsalRujukan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcAsalRujukan.MatchedWithList = True Then txtNoRujukan.SetFocus
        strSQL = "SELECT KdRujukanAsal, RujukanAsal FROM RujukanAsal where StatusEnabled='1' and (RujukanAsal LIKE '%" & dcAsalRujukan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcAsalRujukan.Text = ""
            Exit Sub
        End If
        dcAsalRujukan.BoundText = rs(0).value
        dcAsalRujukan.Text = rs(1).value
    End If
End Sub

Private Sub dcDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcDiagnosa.MatchedWithList = True Then cmdSimpan.SetFocus
        strSQL = "SELECT KdDiagnosa, NamaDiagnosa FROM Diagnosa where StatusEnabled='1'  and (NamaDiagnosa LIKE '%" & dcDiagnosa.Text & "%' or KdDiagnosa LIKE '%" & dcDiagnosa.Text & "%') ORDER BY NamaDiagnosa"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcDiagnosa.Text = ""
            Exit Sub
        End If
        dcDiagnosa.BoundText = rs(0).value
        dcDiagnosa.Text = rs(1).value
            
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
              
            result = context.RefDiagnosa(dcDiagnosa.Text)
        
        Dim i As Long
        For i = LBound(result) To UBound(result)
            Dim arr() As String
            arr = Split(result(i), ":")
            
            Select Case arr(0)
                    Case "kode"
                         blnKartuAktif = True
                         noKartu = arr(1)
                         Debug.Print "kode : " & arr(1)
            End Select
        Next i
        
        If UBound(result) = 0 Then
            blnKartuAktif = False
            MsgBox "Data SEP tidak ditemukan" & vbCrLf & Replace(result(0), "message:", ""), vbInformation, "Validasi"
            Debug.Print result(0)
            Exit Sub
        End If
    Else
        MsgBox "Sdk Bridging askes tidak di temukan"
    End If
Exit Sub
hell:
MsgBox "Koneksi Bridging Bermasalah"
End If
End Sub

Private Sub dcGolonganAsuransi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcGolonganAsuransi.MatchedWithList = True Then txtAlamatPA.SetFocus
        strSQL = "SELECT     KdGolongan, NamaGolongan FROM GolonganAsuransi where StatusEnabled='1' and (NamaGolongan LIKE '%" & dcGolonganAsuransi.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcGolonganAsuransi.Text = ""
            Exit Sub
        End If
        dcGolonganAsuransi.BoundText = rs(0).value
        dcGolonganAsuransi.Text = rs(1).value
    End If
End Sub

Private Sub dcHubungan_Change()
    txtAnakKe.Text = ""
    If dcHubungan.BoundText = "04" Then txtAnakKe.Enabled = True Else txtAnakKe.Enabled = False
End Sub

Private Sub dcJenisPasien_Change()
    On Error GoTo errLoad
    Set rs = Nothing
    rs.Open "select * from v_Penjaminpasien where KdKelompokPasien='" & dcJenisPasien.BoundText & "' and StatusEnabled='1' ORDER BY NamaPenjamin", dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcPenjamin.RowSource = rs
    dcPenjamin.BoundColumn = rs.Fields("idpenjamin").Name
    dcPenjamin.ListField = rs.Fields("namapenjamin").Name
    dcPenjamin.BoundText = ""

    Call msubRecFO(dbRst, "SELECT IdPenjamin FROM PenjaminKelompokPasien WHERE KdKelompokPasien = '" & dcJenisPasien.BoundText & "'")
    If dbRst(0).value = "2222222222" Then
        fraDataKartuPeserta.Enabled = False
        fraPemakaianAsuransi.Enabled = False
        fraDataRujukan.Enabled = False
        dcPerusahaan.Text = ""
    Else
        fraDataKartuPeserta.Enabled = True
        fraPemakaianAsuransi.Enabled = True
        fraDataRujukan.Enabled = True
    End If
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub dcKelasDitanggung_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcKelasDitanggung.BoundText
'    strSQL = "SELECT DISTINCT KdKelas, DeskKelas FROM V_KelasDitanggungPenjamin WHERE (IdPenjamin = '" & dcPenjamin.BoundText & "') AND KdKelompokPasien = '" & dcJenisPasien.BoundText & "'"
    strSQL = "SELECT DISTINCT KdKelas, DeskKelas FROM V_KelasDitanggungPenjamin"
    Call msubDcSource(dcKelasDitanggung, rs, strSQL)
    dcKelasDitanggung.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKelasDitanggung_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKelasDitanggung.MatchedWithList = True Then dcAsalRujukan.SetFocus
        strSQL = "SELECT KdKelas, DeskKelas FROM KelasPelayanan where KdKelas <>'04' and (DeskKelas LIKE '%" & dcKelasDitanggung.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcKelasDitanggung.Text = ""
            Exit Sub
        End If
        dcKelasDitanggung.BoundText = rs(0).value
        dcKelasDitanggung.Text = rs(1).value
    End If
End Sub

Private Sub dcHubungan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcHubungan.MatchedWithList = True Then If txtAnakKe.Enabled = True Then txtAnakKe.SetFocus Else txtNoSJP.SetFocus
        strSQL = "SELECT KdHubungan, NamaHubungan FROM HubunganPesertaAsuransi where StatusEnabled='1' and (NamaHubungan LIKE '%" & dcHubungan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcHubungan.Text = ""
            Exit Sub
        End If
        dcHubungan.BoundText = rs(0).value
        dcHubungan.Text = rs(1).value
    End If
End Sub

Private Sub dcJenisPasien_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then If fraDataKartuPeserta.Enabled = True Then dcPenjamin.SetFocus Else cmdSimpan.SetFocus
End Sub

Private Sub dcNamaAsalRujukan_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcNamaAsalRujukan.BoundText
    strSQL = "SELECT DetailRujukanAsal.KdDetailRujukanAsal, DetailRujukanAsal.DetailRujukanAsal" & _
    " FROM DetailRujukanAsal " & _
    " WHERE (KdRujukanAsal = '" & dcAsalRujukan.BoundText & "') and StatusEnabled='1'"
    Set rs = Nothing
    Call msubDcSource(dcNamaAsalRujukan, rs, strSQL)
    dcNamaAsalRujukan.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcNamaAsalRujukan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcNamaAsalRujukan.MatchedWithList = True Then dtpTglDirujuk.SetFocus
        strSQL = "SELECT KdDetailRujukanAsal, DetailRujukanAsal" & _
        " FROM DetailRujukanAsal " & _
        " WHERE (KdRujukanAsal = '" & dcAsalRujukan.BoundText & "') and (DetailRujukanAsal LIKE '%" & dcNamaAsalRujukan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcNamaAsalRujukan.Text = ""
            Exit Sub
        End If
        dcNamaAsalRujukan.BoundText = rs(0).value
        dcNamaAsalRujukan.Text = rs(1).value
    End If
End Sub

Private Sub dcNamaPerujuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcNamaPerujuk.MatchedWithList = True Then dcDiagnosa.SetFocus
        strSQL = "SELECT KodeDokter, NamaDokter FROM V_DaftarDokter where (NamaDokter LIKE '%" & dcNamaPerujuk.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcNamaPerujuk.Text = ""
            Exit Sub
        End If
        dcNamaPerujuk.BoundText = rs(0).value
        dcNamaPerujuk.Text = rs(1).value
    End If
End Sub

Private Sub dcPenjaminx()
    On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "SELECT dbo.AsuransiPasien.IdPenjamin, dbo.AsuransiPasien.IdAsuransi, dbo.AsuransiPasien.NoCM, dbo.AsuransiPasien.NamaPeserta, " & _
    " dbo.AsuransiPasien.IDPeserta, dbo.AsuransiPasien.KdGolongan, dbo.AsuransiPasien.TglLahir, dbo.AsuransiPasien.Alamat," & _
    " dbo.AsuransiPasien.KdInstitusiAsal, dbo.InstitusiAsalPasien.InstitusiAsal AS NamaPerusahaan, dbo.InstitusiAsalPasien.StatusEnabled" & _
    " FROM dbo.AsuransiPasien LEFT OUTER JOIN" & _
    " dbo.InstitusiAsalPasien ON dbo.AsuransiPasien.KdInstitusiAsal = dbo.InstitusiAsalPasien.KdInstitusiAsal INNER JOIN" & _
    " dbo.Penjamin ON dbo.AsuransiPasien.IdPenjamin = dbo.Penjamin.IdPenjamin " & _
    " WHERE (AsuransiPasien.NoCM = '" & txtNoCM.Text & "') AND (AsuransiPasien.IdPenjamin = '" & dcPenjamin.BoundText & "') and (dbo.InstitusiAsalPasien.StatusEnabled='1')"
    Call msubRecFO(rs, strSQL)

    
    If rs.EOF = False Then
    If chkDiriSendiri.value = Unchecked Then
        
        txtNoKartuPA.Text = IIf(IsNull(rs("IdAsuransi")), "", rs("IdAsuransi"))
        txtNamaPA.Text = IIf(IsNull(rs("NamaPeserta")), "", rs("NamaPeserta"))
        txtNipPA.Text = IIf(IsNull(rs("IDPeserta")), "-", rs("IDPeserta"))
        dcGolonganAsuransi.BoundText = IIf(IsNull(rs("KdGolongan")), "", rs("KdGolongan"))
        dtpTglLahirPA.value = IIf(IsNull(rs("TglLahir")), Now, rs("TglLahir"))
        txtAlamatPA.Text = IIf(IsNull(rs("Alamat")), "", rs("Alamat"))
        dcPerusahaan.Text = IIf(IsNull(rs("NamaPerusahaan")), "", rs("NamaPerusahaan"))
        Call subLoadPemakaianAsuransi(txtNoPendaftaran.Text, dcPenjamin.BoundText)
        dcHubungan.SetFocus
     
'    Else
'        txtNoKartuPA.Text = ""
'        txtNamaPA.Text = ""
'        txtNipPA.Text = ""
'        dcGolonganAsuransi.BoundText = ""
'        dtpTglLahirPA.value = Now
'        txtAlamatPA.Text = ""
    End If
    
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPenjamin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        If dcPenjamin.MatchedWithList = True Then dcPerusahaan.SetFocus
'        strSQL = "select * from v_Penjaminpasien where KdKelompokPasien='" & dcJenisPasien.BoundText & "' and StatusEnabled='1'  and (NamaPenjamin LIKE '%" & dcPenjamin.Text & "%')ORDER BY NamaPenjamin"
'        Set rs = Nothing
'        Call msubRecFO(rs, strSQL)
'        If rs.EOF = True Then
'            dcPenjamin.Text = ""
'            Exit Sub
'        Else
'            dcPenjamin.BoundText = rs(0).value
'            dcPenjamin.Text = rs(1).value
'        End If
       
'        Call dcPenjaminx
'        strSQL = "select * from asuransiPasien where NoCM='" & txtNoCM.Text & "' and IdPenjamin='" & dcPenjamin.BoundText & "' "
'        Set rs = Nothing
'        Call msubRecFO(rs, strSQL)
'        If (Not rs.EOF) Then
'            txtNoKartuPA.Text = rs("IdAsuransi").value
        End If
'    End If
End Sub

Private Sub dcPerusahaan_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    If KeyAscii = 13 Then
        If dcPerusahaan.MatchedWithList = True Then chkDiriSendiri.SetFocus
        strSQL = "SELECT  KdInstitusiAsal, InstitusiAsal FROM InstitusiAsalPasien WHERE (InstitusiAsal LIKE '" & dcPerusahaan.Text & "%') and StatusEnabled='1'"
        Set rs = Nothing
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcPerusahaan.Text = ""
            Exit Sub
        End If
        dcPerusahaan.BoundText = rs(0).value
        dcPerusahaan.Text = rs(1).value
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcUnitKerja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcUnitKerja.MatchedWithList = True Then dcKelasDitanggung.SetFocus
        strSQL = "SELECT KdRuangan, NamaRuangan FROM Ruangan where StatusEnabled='1'  and (NamaRuangan LIKE '%" & dcUnitKerja.Text & "%')ORDER BY NamaRuangan"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcUnitKerja.Text = ""
            Exit Sub
        End If
        dcUnitKerja.BoundText = rs(0).value
        dcUnitKerja.Text = rs(1).value
    End If
End Sub

Private Sub dtpTglDirujuk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcNamaPerujuk.SetFocus
End Sub

Private Sub dtpTglLahirPA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtNipPA.SetFocus
End Sub

Private Sub dtpTglSJP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtNoBP.SetFocus
End Sub

Private Sub fgDPJP_DblClick()
With fgDPJP
    If .rows <> 1 Then
        txtKdDPJP.Text = .TextMatrix(.row, 0)
        txtDPJP.Text = .TextMatrix(.row, 1)
        .Visible = False
    End If
End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    
        strSQL = "Select Value From SettingGlobal where Prefix ='StatusVclaim'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            StatusVclaim = rs(0).value
        End If
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
'    dtpTglLahirPA.value = Now

    strSQL2 = "SELECT TglLahir FROM Pasien WHERE NoCM='" & Right(mstrNoCM, 6) & "'"
    Call msubRecFO(rs2, strSQL2)

    dtpTglLahirPA.value = rs2(0).value
    dtpTglSJP.value = Now
    dtpTglPulang.value = Now
    dtpTglKejadian.value = Now
    dtpTglDirujuk.value = Now - 1
    
    txtNoBP.Text = ""
    txtNoKunjungan.Text = ""

    If mblnFormDaftarAntrian = True Then txtNoCM.Text = mstrNoCM
    txtNoPendaftaran = mstrNoPen
    If mblnAdmin = False Then
        dcJenisPasien.Enabled = False
    Else
        dcJenisPasien.Enabled = True
    End If

    Call subLoadDcSource

    dcJenisPasien.Text = "ASKES PNS"

    Set rs = Nothing
    If (dcJenisPasien.BoundText = "") Then
    End If
    rs.Open "select * from v_Penjaminpasien where KdKelompokPasien='" & dcJenisPasien.BoundText & "' and StatusEnabled='1' ORDER BY NamaPenjamin", dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcPenjamin.RowSource = rs
    dcPenjamin.BoundColumn = rs.Fields("idpenjamin").Name
    dcPenjamin.ListField = rs.Fields("namapenjamin").Name
    dcPenjamin.BoundText = ""
    bolGenerateSEPSukse = True
    ppkRujukan = ""
    
    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnTemp = True Then
        Unload Me
        If txtFormRegistrasiPengirim.Text = "frmRegistrasiAll" Then
            frmRegistrasiAll.Enabled = True
        Else
            frmRegistrasiRJPenunjang.Enabled = True
        End If
        mblnTemp = False
    Else
        Unload Me
    End If
End Sub

Private Sub MSFlexGrid1_DblClick()
    With MSFlexGrid1
        If .rows <> 1 Then
            If jnsBtn = "Propinsi" Then
                txtKdPropinsi.Text = .TextMatrix(.row, 0)
                txtPropinsi.Text = .TextMatrix(.row, 1)
                cmdKota.Enabled = True
            ElseIf jnsBtn = "Kota" Then
                txtKdKota.Text = .TextMatrix(.row, 0)
                txtKota.Text = .TextMatrix(.row, 1)
                cmdKec.Enabled = True
            ElseIf jnsBtn = "Kecamatan" Then
                txtKdKec.Text = .TextMatrix(.row, 0)
                txtKec.Text = .TextMatrix(.row, 1)
            End If
            
            .Visible = False
        End If
    End With
End Sub

Private Sub txtAlamatPA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcHubungan.SetFocus
End Sub

Private Sub txtAlamatPA_LostFocus()
    txtAlamatPA = StrConv(txtAlamatPA, vbProperCase)
End Sub

Private Sub txtAnakKe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNoSJP.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtLokasiLakaLantas_Change()

End Sub

Private Sub txtDPJP_KeyPress(KeyAscii As Integer)

On Error Resume Next
    
    If KeyAscii = 13 Then
    fgDPJP.Visible = True
    Dim strKdDataPoliBPJS As String
    Dim qryPoliBPJS As String
    Dim rsDataPoliBPJS As recordset
    
    qryPoliBPJS = "SELECT KodeExternal FROM dbo.Ruangan WHERE KdRuangan='" & mstrKdRuanganPasien & "'"
    Call msubRecFO(rsDataPoliBPJS, qryPoliBPJS)
    If rsDataPoliBPJS.EOF = False Then
        strKdDataPoliBPJS = IIf(IsNull(rsDataPoliBPJS(0)) = True, "", rsDataPoliBPJS(0))
    Else
        strKdDataPoliBPJS = "" 'nilai default kalau belum dimapping
    End If
            
        If (Dir("C:\SDK\vclaim\result.tlb") <> "") Then
            Dim context As ContextVclaim
            Set context = New ContextVclaim
            
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
            
'            kdOrNamaFaskes = Trim(txtPpkRujukan.Text)
'            If optFaskes1.value = True Then
'                jnsFaskes = "1"
'            Else
'                jnsFaskes = "2"
'            End If
            
            result = context.RefDokterDpjp(IIf(mstrKdInstalasi = "02", 2, 1), Format(dtpTglSJP, "yyyy-mm-dd"), IIf(strKdDataPoliBPJS = Null, "", strKdDataPoliBPJS))
            
            Dim strResult As String
            strResult = ""
            For n = LBound(result) To UBound(result)
            arr = Split(result(n), ":")
            strResult = strResult & vbCrLf & result(n)
            If arr(0) = "error" Then
                MsgBox strResult
                Exit Sub
            End If
'            If UCase(Trim(Right(arr(0), 4))) = "KODE" Then
'                mstrNoSJP = arr(1)
'                txtKdPPKRujukan.Text = arr(1)
''                Exit Sub
'            End If
'            If UCase(Trim(Right(arr(0), 4))) = "NAMA" Then
'                mstrNoSJP = arr(1)
'                txtPpkRujukan.Text = arr(1)
''                Exit Sub
'            End If

        Next n

            fgDPJP.Visible = True
            Call fillGridWithPropinsi(fgDPJP, result)
        End If
        
        fgDPJP.Visible = True
'        fgFaskes.Left = txtPPKRujukan.Left
'        fgFaskes.Top = 930
    ElseIf KeyAscii = 27 Then
        fgDPJP.Visible = False
        fgDPJP.Clear
    End If
End Sub

Private Sub txtNamaPA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglLahirPA.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtNamaPA_LostFocus()
    txtNamaPA = StrConv(txtNamaPA, vbProperCase)
End Sub

Private Sub txtNipPA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkCheckkartu.value = vbChecked Then
            If Len(Trim(txtNipPA.Text)) = 0 Then Exit Sub
            strJenisID = "NIK"
        End If
        
'        If StatusVclaim = 1 Then
            Call ValidateKartuPeserta
'        End If
        
        dcGolonganAsuransi.SetFocus
    End If
End Sub

Private Sub txtNipPA_LostFocus()
    txtNipPA = StrConv(txtNipPA, vbProperCase)
End Sub

Private Sub txtNoBP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcUnitKerja.SetFocus
End Sub

Private Sub txtNoKartuPA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    'If KeyAscii = 13 Then txtNamaPA.SetFocus
    If KeyAscii = 13 Then dcPenjamin.SetFocus
    strJenisID = "Kartu Peserta"
    
   '---- untuk cek kepesertaan BPJS-------------
    If KeyAscii = 13 Then
      If chkCheckkartu.value = vbChecked Then
        If dcJenisPasien.BoundText = "10" Then
'            If StatusVclaim = "Y" Then
            If optNoKartu.value = True Then
                Call ValidateKartuPeserta
            ElseIf optNoRujukan.value = True Then
                Call CariRujukanPcareByNoRujukan(txtNoKartuPA.Text)
            Else
                Call CariRujukanRSByNoKartu(txtNoKartuPA.Text)
            End If
'            Else
'                Call ValidateKartuPesertaV21
'                ValidateKartuPeserta 0
'            End If
            chkNoSJP.Enabled = True
            chkNoSJP.value = Checked
        End If
      End If
    End If
    '------------------------------------------
End Sub

Private Sub txtNoKartuPA_LostFocus()
'    On Error GoTo errLoad
'
'If optNoKartu.value = True Then
'
'    Dim strKdGolongan As String
'
''    strSQL = "SELECT * FROM AsuransiPasien " _
''    & "WHERE IdPenjamin='" & dcPenjamin.BoundText & "' AND IdAsuransi='" _
''    & txtNoKartuPA.Text & "' AND NoCM='" & txtNoCM.Text & "'"
'
'    strSQL = "SELECT * FROM AsuransiPasien " _
'    & "WHERE IdAsuransi='" _
'    & txtNoKartuPA.Text & "' AND NoCM<>'" & txtNoCM.Text & "'"
'
'    Set rs = Nothing
'    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
'   ' If rs.RecordCount = 0 Then Exit Sub
'
'    If rs.RecordCount <> 0 Then
'    '@Dimas 2014-05-13
'        If rs("NoCM").value <> mstrNoCM Then
'            MsgBox "Nomor sudah dipakai oleh pasien lain", vbCritical, "Validasi"
'                '   dcPenjamin.Enabled = False
'            cmdSimpan.Enabled = False
'            Exit Sub
'        End If
'    Else
'     dcPenjamin.Enabled = True
'     cmdSimpan.Enabled = True
'
'    End If
''
''
''    txtNamaPA.Text = rs.Fields("NamaPeserta").value
''    If Not IsNull(rs.Fields("IDPeserta").value) Then txtNipPA.Text = rs.Fields("IDPeserta").value
''
''    dcPenjamin.BoundText = rs.Fields("IdPenjamin").value
''
''    strSQL2 = "select * from v_Penjaminpasien where KdKelompokPasien='" & dcJenisPasien.BoundText & "' and IdPenjamin= '" & rs.Fields("IdPenjamin") & "'"
''    'Set dcPenjamin.RowSource = rs
''    Call msubRecFO(rs2, strSQL2)
''    dcPenjamin.Text = rs2.Fields("namapenjamin").value
''
''    dtpTglLahirPA.value = rs.Fields("TglLahir").value
''    strKdGolongan = rs.Fields("KdGolongan").value
''    If Not IsNull(rs.Fields("Alamat").value) Then txtAlamatPA.Text = rs.Fields("Alamat").value
''    strSQL = "SELECT NamaGolongan,KdGolongan FROM GolonganAsuransi WHERE KdGolongan='" & strKdGolongan & "' and StatusEnabled='1'"
''    Set rs = Nothing
''    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
''    dcKelasDitanggung.Text = rs.Fields(0).value
''    dcKelasDitanggung.BoundText = rs.Fields(1).value
''    Set rs = Nothing
''    txtNoKartuPA = StrConv(txtNoKartuPA, vbProperCase)
''
''
''
''    Exit Sub
''errLoad:
'
'On Error GoTo errLoad
'
''Dim strKdGolongan As String
'If chkCheckkartu.value = vbChecked Then
'        strSQL = "SELECT * FROM AsuransiPasien " _
'            & "WHERE IdAsuransi='" _
'            & txtNoKartuPA.Text & "' and NoCM='" & txtNoCM.Text & "'"
'        Call msubRecFO(rs1, strSQL)
'        If (rs1.EOF = True) Then
'            cmdSimpan.Enabled = True
'            strSQL = "select * from SettingGlobal where prefix='KdKelompokPasienBPJS' and value='" & dcJenisPasien.BoundText & "'"
'            Call msubRecFO(rs, strSQL)
'                If (rs.EOF = False) Then
'                    ValidateKartuPeserta
'                    chkNoSJP.Enabled = True
'                    chkNoSJP.value = Checked
'                End If
'        Else
'            cmdSimpan.Enabled = True
'            strSQL = "select * from SettingGlobal where prefix='KdKelompokPasienBPJS' and value='" & dcJenisPasien.BoundText & "'"
'            Call msubRecFO(rs, strSQL)
'                If (rs.EOF = False) Then
'                    ValidateKartuPeserta
'                    chkNoSJP.Enabled = True
'                    chkNoSJP.value = Checked
'
'                End If
'        End If
'
'Else
'    strSQL = "SELECT * FROM AsuransiPasien " _
'        & "WHERE IdPenjamin='" & dcPenjamin.BoundText & "' AND IdAsuransi='" _
'        & txtNoKartuPA.Text & "' AND NoCM='" & txtNoCM.Text & "'"
'    Set rs = Nothing
'    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
'    If rs.RecordCount = 0 Then Exit Sub
'    txtNamaPA.Text = rs.Fields("NamaPeserta").value
'    If Not IsNull(rs.Fields("IDPeserta").value) Then txtNipPA.Text = rs.Fields("IDPeserta").value
'    dtpTglLahirPA.value = rs.Fields("TglLahir").value
'    strKdGolongan = rs.Fields("KdGolongan").value
'    If Not IsNull(rs.Fields("Alamat").value) Then txtAlamatPA.Text = rs.Fields("Alamat").value
'    strSQL = "SELECT DISTINCT NamaGolongan,KdGolongan FROM GolonganAsuransi WHERE KdGolongan='" & strKdGolongan & "'"
'    Set rs = Nothing
'    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
'    dcKelasDitanggung.Text = rs.Fields(0).value
'    dcKelasDitanggung.BoundText = rs.Fields(1).value
'    Set rs = Nothing
'    txtNoKartuPA = StrConv(txtNoKartuPA, vbProperCase)
'End If
'
'ElseIf optNoRujukan.value = True Then
'    strJenisID = "No Rujukan"
'    strSQL = "select * from SettingGlobal where prefix='KdKelompokPasienBPJS' and value='" & dcJenisPasien.BoundText & "'"
'    Call msubRecFO(rs, strSQL)
'    If (rs.EOF = False) Then
'        Call CariRujukanPcareByNoRujukan(txtNoKartuPA.Text)
'    End If
'
'End If
'
'Exit Sub
'errLoad:
'    Call msubPesanError

End Sub

Private Sub txtNoKunjungan_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtNoRujukan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcNamaAsalRujukan.SetFocus
    
    If KeyAscii >= 65 And KeyAscii <= 90 Then
        Beep
        MsgBox "Harus Diisi Dengan Angka", vbCritical, "Validasi"
        KeyAscii = 0
    End If
End Sub


Private Sub txtNoSJP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dtpTglSJP.SetFocus
        
'        Call CariSEP
    End If
End Sub

Private Sub txtNoSJP_LostFocus()
    txtNoSJP = StrConv(txtNoSJP, vbProperCase)
End Sub

Private Function sp_AmbulNoKunjungan() As Boolean
    On Error GoTo errLoad
    sp_AmbulNoKunjungan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, dcPenjamin.BoundText)
        .Parameters.Append .CreateParameter("IdAsuransi", adChar, adParamInput, 15, Trim(txtNoKartuPA.Text))
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("KunjunganKe", adInteger, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("TglRujukanOut", adDate, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("TglPendaftaran", adDate, adParamInput, , Format(txtTglPendaftaran.Text, "yyyy/MM/dd hh:mm:ss"))
        .Parameters.Append .CreateParameter("NoSJPRujukan", adVarChar, adParamInput, 30, Trim(txtNoSJP.Text))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Check_NoRujukan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam pengambilan No Kunjungan", vbExclamation, "Validasi"
            sp_AmbulNoKunjungan = False
        Else
            txtNoKunjungan.Text = .Parameters("KunjunganKe").value
            If txtNoKunjungan.Text = "0" Then
                MsgBox "Masa berlaku No. Rujukan (SJP) sudah HABIS", vbExclamation, "Informasi"
                sp_AmbulNoKunjungan = False
            ElseIf Val(txtNoKunjungan.Text) > 3 Then
                MsgBox "Masa kunjungan No. Rujukan (SJP) sudah lebih dari 3 kali", vbExclamation, "Informasi"
                sp_AmbulNoKunjungan = False
            End If
            Call Add_HistoryLoginActivity("Check_NoRujukan")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    Call msubPesanError
    sp_AmbulNoKunjungan = False
End Function

Private Function sp_UpdateJenisPasienUmum(f_KdKelompokPasien As String, f_NoPendaftaran As String) As Boolean
    On Error GoTo errLoad
    sp_UpdateJenisPasienUmum = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKelompokpasien", adChar, adParamInput, 2, f_KdKelompokPasien)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_JenisPasienUmumNew"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_UpdateJenisPasienUmum = False
        Else
            Call Add_HistoryLoginActivity("Update_JenisPasienUmumNew")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    sp_UpdateJenisPasienUmum = False
    Call msubPesanError
End Function

Private Function sp_UpdateJenisPasienJaminan(f_KdKelompokPasien As String, f_NoPendaftaran As String) As Boolean
    On Error GoTo errLoad
    sp_UpdateJenisPasienJaminan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKelompokpasien", adChar, adParamInput, 2, f_KdKelompokPasien)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_JenisPasienJaminan"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_UpdateJenisPasienJaminan = False
        Else
            Call Add_HistoryLoginActivity("Update_JenisPasienJaminan")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    sp_UpdateJenisPasienJaminan = False
    Call msubPesanError
End Function

Private Sub ValidateKartuPeserta()
On Error GoTo hell
'If Periksa("datacombo", dcPenjamin, "Penjamin belum di isi") = False Then Exit Sub

    txtTempHakKelas.Text = ""
    blnKartuAktif = False
    fKdKelasDitanggung = ""
    
If (Dir("C:\SDK\Vclaim\result.tlb") <> "") Then
        Dim context As ContextVclaim
        Set context = New ContextVclaim
        Dim result() As String
        
        strSQL = "Select Value From SettingGlobal where Prefix In('ConsumerID','PasswordKey')"
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = False Then
            context.ConsumerID = rs(0).value
'            rs.MoveNext
'            context.KodeRumahSakit = rs(0).value
            rs.MoveNext
            context.PasswordKey = rs(0).value
        End If
        
        
         strSQL = "SELECT Value FROM SettingGlobal where Prefix='UrlGenerateSEP'"
            Call msubRecFO(rs, strSQL)
            Dim URL  As String
              If rs.EOF = False Then
                  URL = rs.Fields(0)
                  context.URL = URL
                 'Exit Sub
              End If
              
        
        If strJenisID = "NIK" Then
            'untuk yg pake NIK
            result = context.CariPesertaByNik(txtNipPA.Text, Format(Now, "yyyy-mm-dd"))
        Else
            'untuk yg pake NO Kartu BPJS
            result = context.CariPesertaByNoKartuBpjs(txtNoKartuPA.Text, Format(Now, "yyyy-mm-dd"))
            
        End If
        
        Dim i As Long
        For i = LBound(result) To UBound(result)
            Dim arr() As String
            arr = Split(result(i), ":")
            
            
            Select Case arr(0)
                Case "MR-NOKARTU"
                         blnKartuAktif = True
                         txtNoKartuPA.Text = arr(1)
                         noKartu = arr(1)
                         Debug.Print "NoKartu : " & arr(1)
                    Case "MR-NIK"
                         txtNipPA.Text = arr(1)
                         nik = arr(1)
                    Case "MR-NAMA"
                         txtNamaPA.Text = arr(1)
                         nama = arr(1)
                    Case "PROVUMUM-NMPROVIDER"
                         dcNamaAsalRujukan.Text = arr(1)
                         nmProvider = arr(1)
                    Case "STATUSPESERTA-TGLLAHIR"
                        If dtpTglLahirPA.value <> CDate(Split(arr(1), " ")(0)) Then
                            MsgBox "Tanggal lahir tidak sama!" & "Tanggal Lahir Peserta BPJS: " & arr(1) & vbCrLf & "Silakan cek data pasien", vbOKOnly, "Cek Kepesertaan"
                        End If
                        dtpTglLahirPA.value = CDate(Split(arr(1), " ")(0))
                        tgllahir = arr(1)
                    Case "PROVUMUM-KDPROVIDER"
                         ppkRujukan = arr(1)
                         kdProvider = arr(1)
                         txtPpkRujukan.Text = arr(1)
                    Case "HAKKELAS-KODE"
                        kdKelas = arr(1)
                    Case "HAKKELAS-KETERANGAN"
                        strSQL = "SELECT KdKelas FROM KelasPelayanan where NamaExternal='" & arr(1) & "'"
                        Call msubRecFO(rs, strSQL)
                     
                        If (rs.EOF = False) Then
                           fKdKelasDitanggung = rs(0).value
                           dcKelasDitanggung.BoundText = fKdKelasDitanggung
                        End If
                        
                        mstrKelasDitanggung = arr(1)
                        dcKelasDitanggung.Text = arr(1)
                        nmKelas = arr(1)
                   Case "JENISPESERTA-KETERANGAN"
                        txtJenisPasien = arr(1)
                        nmJenisPeserta = arr(1)
                    Case "MR-PISA"
                        Call msubRecFO(rs, "SELECT KdHubungan FROM dbo.HubunganPesertaAsuransi WHERE KodeExternal='" & arr(1) & "'")
                        If rs.EOF = False Then
                            dcHubungan.BoundText = rs(0)
                        End If
                        pisa = arr(1)
                    Case "PROVUMUM-SEX"
                         sex = arr(1)
                    Case "STATUSPESERTA-TGLCETAKKART"
                         tglCetakKartu = arr(1)
                    Case "STATUSPESERTA-KETERANGAN"
                         statusPeserta = arr(1)
                    Case "keluhan"
                         txtCatatan.Text = arr(1)
                    Case "kdCabang"
                         kdCabang = arr(1)
                    Case "nmCabang"
                         nmCabang = arr(1)
                    Case "NOKUNJUNGAN"
                         txtNoRujukan.Text = arr(1)
                    Case "TGLKUNJUNGAN"
                         dtpTglDirujuk.value = CDate(Split(arr(1), " ")(0))
                    Case "INFORMASI-PROLANISPRB"
                         potensiprb = arr(1)
            End Select
        Next i
        
        mstrTglVerifBPJS = Format(dtpTglLahirPA.value, "yyyy-MM-dd 00:00:00")
        mstrkartuPeserta = txtNoKartuPA.Text
        mstrnorujukan = txtNoRujukan.Text
        mstrJenisPeserta = txtJenisPasien.Text
        mstrJenisPesertaBpjs = txtJenisPasien.Text
        mstrNamaAsalRujukanMon = dcNamaAsalRujukan.Text
        mstrPpkRujukan = txtPpkRujukan.Text
        'mstrCatatatnBPJS = txtKeluhan.Text
        mstrNamaPeserta = txtNamaPA.Text
        
        If UBound(result) = 0 Then
            blnKartuAktif = False
            MsgBox "Data Peserta tidak ditemukan" & vbCrLf & Replace(result(0), "message:", ""), vbInformation, "Validasi"
            Debug.Print txtNoKartuPA.Text
            Debug.Print result(0)
            txtJenisPasien.Text = ""
            Exit Sub
        Else
            If sp_detailkartubpjs(noKartu, nik, nama, pisa, sex, tgllahir, tglCetakKartu, kdProvider, nmProvider, kdCabang, nmCabang, kdJenisPeserta, nmJenisPeserta, kdKelas, nmKelas, potensiprb) = False Then Exit Sub
        End If
        
'        If txtNamaPA.Text = "" Then
'            MsgBox "Data Peserta Tidak Ditemukan,,,,!!!", vbInformation, "Validasi"
'            txtNoKartuPA.Text = ""
'            'Set DgPasien2.DataSource = Nothing
'            'Call SubloadPasienBPJS
'            Exit Sub
'        End If
        
        'Call SubloadPasienBPJS
        'DgPasien2.SetFocus
    Else
        MsgBox "Sdk Bridging askes tidak di temukan"
    End If
Exit Sub
hell:
MsgBox "Koneksi Bridging Bermasalah"
End Sub

Private Sub GenerateSEPBPJS()
If chkNoSJP.value = vbChecked Then
             
          If (Dir("C:\SDK\Vclaim\result.tlb") <> "") Then
            Dim context As ContextVclaim
            Set context = New ContextVclaim
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
                 'Exit Sub
              End If
            
            mstrnorujukan = txtNoRujukan.Text
           
           Dim kodeKelas As Integer
           If (dcKelasDitanggung.Text = "Kelas I") Then
            kodeKelas = 1
           ElseIf (dcKelasDitanggung.Text = "Kelas II") Then
            kodeKelas = 2
           Else
            kodeKelas = 3
           End If
           
            strKdDiagnosa = Trim(dcDiagnosa.BoundText)
           
            Dim rsDataPoliBPJS As New ADODB.recordset
            Dim qryPoliBPJS As String
            Dim strKdDataPoliBPJS As String
            
            ''''''''' langsung ngambil kdruangan dri RS ''''''''''
            qryPoliBPJS = "SELECT KodeExternal FROM dbo.Ruangan WHERE KdRuangan='" & mstrKdRuanganPasien & "'"
            Call msubRecFO(rsDataPoliBPJS, qryPoliBPJS)
            If rsDataPoliBPJS.EOF = False Then
                strKdDataPoliBPJS = IIf(IsNull(rsDataPoliBPJS(0)) = True, "", rsDataPoliBPJS(0))
            Else
                strKdDataPoliBPJS = "" 'nilai default kalau belum dimapping
            End If
            
            '2+2=5, untuk testing generate ke DVLP IDPegawai panjangnnya tidak boleh lebih dari 9. Dipatok saja ke 888888888(9 Digit)
            Dim strIDPegawaiLokal As String
            strIDPegawaiLokal = strIDPegawai
            If URL = "http://dvlp.bpjs-kesehatan.go.id:8081/devWSLokalRest/" Then
                strIDPegawaiLokal = "88888888"
            End If
            
            strSQL = "SELECT Value FROM SettingGlobal WHERE Prefix='KodeRS'"
            Call msubRecFO(rs, strSQL)
            txtppkpelayanan.Text = rs(0).value
            
            strSQL = "SELECT Telepon FROM Pasien WHERE NoCM='" & txtNoCM.Text & "'"
            Call msubRecFO(rs, strSQL)
            txtNoTlpPasien.Text = IIf(IsNull(rs(0).value), "", Trim(rs(0).value))
            
'            strSQL = "SELECT dbo.getNamaPegawaiByIdPegawai('" & strIDPegawai & "')"
            strSQL = "SELECT NamaLengkap FROM DataPegawai WHERE IdPegawai='" & strIDPegawai & "'"
            Call msubRecFO(rs, strSQL)
            strNamaPegawai = rs(0).value
            
            
            lakalantas = IIf(chkLakalantas.value = False, "0", "1")
            lokasiLaka = IIf(Len(Trim(txtKec.Text)) = 0, "0", Trim(txtKec.Text))
            penjaminLakalantas = 0
            If chkJasaRaharja.value = vbChecked And chkBPJSKK.value = vbUnchecked And chkTaspen.value = vbUnchecked And chkAsabri.value = vbUnchecked Then
                penjaminLakalantas = "1"
            ElseIf chkJasaRaharja.value = vbUnchecked And chkBPJSKK.value = vbChecked And chkTaspen.value = vbUnchecked And chkAsabri.value = vbUnchecked Then
                penjaminLakalantas = "2"
            ElseIf chkJasaRaharja.value = vbUnchecked And chkBPJSKK.value = vbUnchecked And chkTaspen.value = vbChecked And chkAsabri.value = vbUnchecked Then
                penjaminLakalantas = "3"
            ElseIf chkJasaRaharja.value = vbUnchecked And chkBPJSKK.value = vbUnchecked And chkTaspen.value = vbUnchecked And chkAsabri.value = vbChecked Then
                penjaminLakalantas = "4"
            ElseIf chkJasaRaharja.value = vbChecked And chkBPJSKK.value = vbChecked And chkTaspen.value = vbUnchecked And chkAsabri.value = vbUnchecked Then
                penjaminLakalantas = "1,2"
            ElseIf chkJasaRaharja.value = vbChecked And chkBPJSKK.value = vbUnchecked And chkTaspen.value = vbChecked And chkAsabri.value = vbChecked Then
                penjaminLakalantas = "1,3"
            ElseIf chkJasaRaharja.value = vbChecked And chkBPJSKK.value = vbUnchecked And chkTaspen.value = vbUnchecked And chkAsabri.value = vbChecked Then
                penjaminLakalantas = "1,4"
            ElseIf chkJasaRaharja.value = vbUnchecked And chkBPJSKK.value = vbChecked And chkTaspen.value = vbChecked And chkAsabri.value = vbUnchecked Then
                penjaminLakalantas = "2,3"
            ElseIf chkJasaRaharja.value = vbUnchecked And chkBPJSKK.value = vbChecked And chkTaspen.value = vbUnchecked And chkAsabri.value = vbChecked Then
                penjaminLakalantas = "2,4"
            ElseIf chkJasaRaharja.value = vbUnchecked And chkBPJSKK.value = vbUnchecked And chkTaspen.value = vbChecked And chkAsabri.value = vbChecked Then
                penjaminLakalantas = "3,4"
            ElseIf chkJasaRaharja.value = vbChecked And chkBPJSKK.value = vbChecked And chkTaspen.value = vbChecked And chkAsabri.value = vbUnchecked Then
                penjaminLakalantas = "1,2,3"
            ElseIf chkJasaRaharja.value = vbChecked And chkBPJSKK.value = vbChecked And chkTaspen.value = vbUnchecked And chkAsabri.value = vbUnchecked Then
                penjaminLakalantas = "1,2,4"
            ElseIf chkJasaRaharja.value = vbChecked And chkBPJSKK.value = vbChecked And chkTaspen.value = vbChecked And chkAsabri.value = vbChecked Then
                penjaminLakalantas = "1,2,3,4"
            End If
            
            cob = 0
            If chkCob.value = vbChecked Then
                cob = "1"
            Else
                cob = 0
            End If
'            Diganti karena mstrNoSJPNew Harus String Array Ales
            
                mstrNoSJPNew = context.InsertSep(txtNoKartuPA.Text, Format(dtpTglSJP.value, "yyyy-mm-dd"), txtppkpelayanan.Text, IIf(mstrKdInstalasi = "03", "1", "2"), kodeKelas, txtNoCM.Text, _
                            1, Format(dtpTglDirujuk.value, "yyyy-mm-dd"), txtNoRujukan.Text, ppkRujukan, txtCatatan.Text, strKdDiagnosa, IIf(mstrKdInstalasi, strKdDataPoliBPJS, ""), _
                            0, cob, lakalantas, penjaminLakalantas, lokasiLaka, txtNoTlpPasien.Text, strNamaPegawai)
            
            Dim Temp As String
            Dim i As Integer
            For i = LBound(mstrNoSJPNew) To UBound(mstrNoSJPNew)
                arr = Split(mstrNoSJPNew(i), ":")
                Temp = Temp & vbCrLf & mstrNoSJPNew(i)
                If arr(0) = "error" Then
                    MsgBox Replace(mstrNoSJPNew(0), "error:", "HUBUNGI BPJS "), vbExclamation, "Generate SEP BPJS"
'                    GenerateSEPBPJS = False
                    txtNoSJP.Text = "-"
'                    Exit Function
                End If
                If UCase(Trim(Right(arr(0), 5))) = "NOSEP" Or UCase(Trim(Right(arr(0), 5))) = "UMUR-NOSEP" Then
                    mstrNoSJP = arr(1)
                    Exit For
                End If
'                Dim arr() As String
'                arr = Split(mstrNoSJPNew(i), ":")
'                Select Case arr(0)
'
'                    Case "noSep"
'                       txtNoSJP.Text = arr(1)
'                       Debug.Print "noSep : " & arr(1)
'                    Case "error"
'                        MsgBox mstrNoSJPNew(i)
'                        MsgBox "Generate SEP Gagal,,,,!!!"
'                End Select
             Next i
             
'             mstrNoSJP = Mid(mstrNoSJPNew(10), 1, 8)
                    If mstrNoSJP <> "" Then
                        
                        Call txtNoSJP_KeyPress(13)
                        txtNoSJP.Text = mstrNoSJP
                        cmdSimpan.Enabled = False
                        typAsuransi.blnSuksesAsuransi = True
                    Else
                    MsgBox mstrNoSJPNew   'Mid(Temp, 40, (Len(Temp) - Len(Right(Temp, 58))))  'mstrNoSJP
'            If txtFormRegistrasiPengirim.Text = "frmRegistrasiAll" Then
'                frmRegistrasiAll.Enabled = True
'                mstrNoSJP = context.GenerateSep(txtNoKartuPA.Text, Format(dtpTglDirujuk.value, "yyyy-MM-dd 00:00:00"), mstrnorujukan, ppkRujukan, IIf((mstrKdInstalasi = "02") Or (mstrKdInstalasi = "06") Or (mstrKdInstalasi = "01") Or (frmRegistrasiAll.dcRuangan.BoundText = "004"), "2", "1"), mstrCatatatnBPJS, Replace(dcDiagnosa.BoundText, ".", ""), frmRegistrasiAll.dcRuangan.BoundText, kodeKelas, strIDPegawai, txtNoCM.Text)
'            Else
'                frmRegistrasiRJPenunjang.Enabled = True
'                mstrNoSJP = context.GenerateSep(txtNoKartuPA.Text, Format(dtpTglDirujuk.value, "yyyy-MM-dd 00:00:00"), mstrnorujukan, ppkRujukan, IIf((mstrKdInstalasi = "02") Or (mstrKdInstalasi = "06") Or (mstrKdInstalasi = "01") Or (frmRegistrasiRJPenunjang.dcRuangan.BoundText = "004"), "2", "1"), mstrCatatatnBPJS, Replace(dcDiagnosa.BoundText, ".", ""), frmRegistrasiRJPenunjang.dcRuangan.BoundText, kodeKelas, strIDPegawai, txtNoCM.Text)
'            End If
           
'            mstrNoSJP = context.GenerateSep(txtNoKartuPA.Text, "2014-04-21 13:48:58", "090507080414Y000115", mstrKdRuanganORS, IIf(mstrKdInstalasi = "02", "2", "1"), mstrCatatatnBPJS, dcDiagnosa.BoundText, frmPasienBaru_Daftar.dcRuangan.BoundText, dcKelasDitanggung.BoundText, strIDPegawai, txtNoCM.Text)
             
'             mstrNoSJP = Mid(mstrNoSJP, 14, 19)
'             If (Mid(mstrNoSJP, 1, 8) = context.KodeRumahSakit) Then
'                    txtNoSJP.Text = mstrNoSJP
'                    Call txtNoSJP_KeyPress(13)
'                    If mbolSEP = True Then
'                        Exit Sub
'                    End If
'
'                    strNoSEPForSimpan = mstrNoSJP
'                    txtNoSJP.Text = mstrNoSJP
'                    'Call sp_JenisPasienJoinProgramBPJS
'                    typAsuransi.blnSuksesAsuransi = True
'                    cmdSimpan.Enabled = False
'                    bolGenerateSEPSukse = True
'                Else
'                    MsgBox mstrNoSJP
'                    bolGenerateSEPSukse = False
             End If
            
        Else
            'sp_JenisPasienJoinProgramBPJS
            MsgBox "Generate SEP Gagal,,,,!!!"
            typAsuransi.blnSuksesAsuransi = False
            bolGenerateSEPSukse = False
            txtNoSJP.Text = ""
        End If
End If
Exit Sub
SepEndPoint:
'     sp_JenisPasienJoinProgramBPJS
End Sub

'Private Function sp_JenisPasienJoinProgramBPJS() As Boolean
'On Error GoTo StatusErr
'
'    MousePointer = vbHourglass
'    sp_JenisPasienJoinProgramBPJS = True
'    Set dbcmd = New ADODB.Command
'    With dbcmd
'    .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
'    .Parameters.Append .CreateParameter("IdAsuransi", adVarChar, adParamInput, 15, txtNoKartuPA)
'    .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 15, txtNoCM)
'    .Parameters.Append .CreateParameter("TglPendaftaran", adDate, adParamInput, , Format(txtTglPendaftaran.Text, "yyyy/MM/dd hh:mm:ss"))
'    .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, dcJenisPasien.BoundText)
'    .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, dcPenjamin.BoundText)
'
'    .Parameters.Append .CreateParameter("NoSJP", adVarChar, adParamInput, 30, mstrNoSJP)
'
'    .Parameters.Append .CreateParameter("TglSJP", adDate, adParamInput, , Format(dtpTglSJP, "yyyy/MM/dd hh:mm:ss"))
'    .Parameters.Append .CreateParameter("OutputNoSJP", adVarChar, adParamOutput, 30, Null)
'    .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
'    .ActiveConnection = dbConn
'    .CommandText = "dbo.AUD_NoSEP"
'    .CommandType = adCmdStoredProc
'    .CommandTimeout = 120
'    .Execute
'        If .Parameters("return_value").value <> 0 Then
'            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
'            sp_JenisPasienJoinProgramBPJS = False
'        Else
'            txtNoSJP.Text = IIf(IsNull(.Parameters("OutputNoSJP")), "", .Parameters("OutputNoSJP"))
'            mstrNoSJP = IIf(IsNull(.Parameters("OutputNoSJP")), "", .Parameters("OutputNoSJP"))
'            'cmdSimpan.Enabled = False
'
'        End If
'        Call deleteADOCommandParameters(dbcmd)
'        Set dbcmd = Nothing
'    End With
'    MousePointer = vbDefault
'
'Exit Function
'StatusErr:
'    cmdSimpan.Enabled = True
'    MousePointer = vbDefault
'    sp_JenisPasienJoinProgramBPJS = False
'    Call msubPesanError("sp_JenisPasienJoinProgramBPJS")
'    MsgBox "Ulangi proses simpan ", vbCritical, "Validasi"
'End Function

Private Function sp_detailkartubpjs(f_noKartu As String, f_nik As String, f_nama As String, f_pisa As String, f_sex As String, f_TglLahir As String, f_tglCetakKartu As String, f_kdProvider As String, f_nmProvider As String, f_kdCabang As String, f_nmCabang As String, f_kdJenisPeserta As String, f_nmJenisPeserta As String, f_kdKelas As String, f_nmKelas As String, f_potensiprb As String)
On Error GoTo errLoad
    sp_detailkartubpjs = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("noKartu", adChar, adParamInput, 13, f_noKartu)
        .Parameters.Append .CreateParameter("nik", adChar, adParamInput, 16, f_nik)
        .Parameters.Append .CreateParameter("nama", adVarChar, adParamInput, 50, f_nama)
        .Parameters.Append .CreateParameter("pisa", adVarChar, adParamInput, 3, f_pisa)
        .Parameters.Append .CreateParameter("sex", adVarChar, adParamInput, 3, f_sex)
        .Parameters.Append .CreateParameter("tglLahir", adVarChar, adParamInput, 20, f_TglLahir)
        .Parameters.Append .CreateParameter("tglCetakKartu", adVarChar, adParamInput, 20, f_tglCetakKartu)
        .Parameters.Append .CreateParameter("kdProvider", adVarChar, adParamInput, 20, f_kdProvider)
        .Parameters.Append .CreateParameter("nmProvider", adVarChar, adParamInput, 50, f_nmProvider)
        .Parameters.Append .CreateParameter("kdCabang", adVarChar, adParamInput, 20, f_kdCabang)
        .Parameters.Append .CreateParameter("nmCabang", adVarChar, adParamInput, 50, f_nmCabang)
        .Parameters.Append .CreateParameter("kdJenisPeserta", adVarChar, adParamInput, 20, f_kdJenisPeserta)
        .Parameters.Append .CreateParameter("nmJenisPeserta", adVarChar, adParamInput, 50, f_nmJenisPeserta)
        .Parameters.Append .CreateParameter("kdKelas", adVarChar, adParamInput, 10, f_kdKelas)
        .Parameters.Append .CreateParameter("nmKelas", adVarChar, adParamInput, 20, f_nmKelas)
        .Parameters.Append .CreateParameter("potensiprb", adVarChar, adParamInput, 15, f_potensiprb)
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_DetailKartuBPJS"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute

        Call Add_HistoryLoginActivity("AUD_detailkartubpjs")
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    sp_detailkartubpjs = False
    Call msubPesanError
End Function

Private Sub HapusSEPVclaim()
On Error Resume Next
If (Dir("C:\SDK\Vclaim\result.tlb") <> "") Then
        Dim context As ContextVclaim
        Dim sep() As String
        Set context = New ContextVclaim
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
            context.URL = (URL)
        End If
        sep = context.DeleteSep(txtNoSJP.Text, Format(Now, "yyyy-mm-dd"))
'        If InStr(1, UCase(sep), "GAGAL") = 0 And InStr(1, UCase(sep), "ERROR") = 0 Then
        For n = LBound(sep) To UBound(sep)
            arr = Split(sep(n), ":")
            Select Case arr(0)
                Case "error"
                    MsgBox arr(1), vbExclamation, "Hapus SEP BPJS"
                    txtKetHapus.Text = ""
                    fraHapusSEP.Visible = False
                    Exit Sub
            End Select
        Next n
            strSQL = "UPDATE dbo.PemakaianAsuransi SET NoSJP='-' WHERE NoPendaftaran " & IIf(Trim(txtNoPendaftaran.Text) = "", " is NULL ", "='" & txtNoPendaftaran.Text & "'") & " AND NoSJP='" & txtNoSJP.Text & "'"
            Call msubRecFO(rs, strSQL)
    
    
            dbConn.Execute "INSERT INTO dbo.DaftarHapusSEP( NoPendaftaran ,NoSJP ,TglSJP ,KdLakaLantas ,PPKRujukan ,TglActivity,Keterangan,IDPegawai) " & _
                                "VALUES  ( '" & txtNoPendaftaran.Text & "' , '" & txtNoSJP.Text & "','" & Format(dtpTglSJP.value, "yyyy-MM-dd hh:mm:ss") & "' , " & IIf(chkLakalantas.value = False, 0, 1) & " , '" & txtPpkRujukan.Text & "' , GETDATE(),'" & txtKetHapus.Text & "','" & strIDPegawaiAktif & "' )"
            txtNoSJP.Text = "-"
        End If
    
        MsgBox sep, vbInformation, ""
'    Else
'        MsgBox "error", vbInformation, ""
'    End If
End Sub

Private Sub CariRujukanPcareByNoRujukan(vNoRujukan As String)
On Error GoTo mulih

    Dim context As ContextVclaim
    Set context = New ContextVclaim
    
    Dim result() As String
    Dim URL As String
    Dim i As Long
    
    strSQL = "select value from SettingGlobal where Prefix in ('ConsumerID','PasswordKey')"
    Call msubRecFO(rs, strSQL)
    
    If rs.EOF = False Then
        context.ConsumerID = rs(0).value
        rs.MoveNext
        context.PasswordKey = rs(0).value
    End If
    
    strSQL = "select value from SettingGlobal where Prefix='UrlGenerateSEP'"
    Call msubRecFO(rs, strSQL)
    
    If rs.EOF = False Then
        URL = rs.Fields(0)
        context.URL = (URL)
    End If
    
    txtNamaPA.Text = ""
    result = context.RujukanPcareByNoKartu(vNoRujukan)
    
    For i = LBound(result) To UBound(result)
        Debug.Print (result(i))
        Dim arr() As String
        arr = Split(result(i), ":")
        Select Case arr(0)
                    Case "PROVPERUJUK-TGLKUNJUNGAN"
                        dtpTglDirujuk.value = CDate(Split(arr(1), " ")(0))
                    Case "DIAGNOSA-KODE"
                        dcDiagnosa.Text = arr(1)
                        dcDiagnosa_KeyPress (13)
                    Case "DIAGNOSA-NOKUNJUNGAN"
                        txtNoRujukan.Text = arr(1)
                    Case "MR-NOKARTU"
                         blnKartuAktif = True
                         txtNoKartuPA.Text = arr(1)
                         noKartu = arr(1)
                         Debug.Print "NoKartu : " & arr(1)
                    Case "MR-NIK"
                         txtNipPA.Text = arr(1)
                         nik = arr(1)
                    Case "MR-NAMA"
                         txtNamaPA.Text = arr(1)
                         nama = arr(1)
                    Case "PROVPERUJUK-NAMA"
                         dcNamaAsalRujukan.Text = arr(1)
                         nmProvider = arr(1)
                    Case "STATUSPESERTA-TGLLAHIR"
                        If dtpTglLahirPA.value <> CDate(Split(arr(1), " ")(0)) Then
                            MsgBox "Tanggal lahir tidak sama!" & "Tanggal Lahir Peserta BPJS: " & arr(1) & vbCrLf & "Silakan cek data pasien", vbOKOnly, "Cek Kepesertaan"
                        End If
                        dtpTglLahirPA.value = CDate(Split(arr(1), " ")(0))
                        tgllahir = arr(1)
                    Case "PROVPERUJUK-KODE"
                         ppkRujukan = arr(1)
                         kdProvider = arr(1)
                         txtPpkRujukan.Text = arr(1)
                    Case "HAKKELAS-KODE"
                        kdKelas = arr(1)
                    Case "HAKKELAS-KETERANGAN"
                        strSQL = "SELECT KdKelas FROM KelasPelayanan where NamaExternal='" & arr(1) & "'"
                        Call msubRecFO(rs, strSQL)
                     
                        If (rs.EOF = False) Then
                           fKdKelasDitanggung = rs(0).value
                           dcKelasDitanggung.BoundText = fKdKelasDitanggung
                           If fKdKelasDitanggung = "01" Then
                                dcGolonganAsuransi.Text = "I"
                            ElseIf fKdKelasDitanggung = "02" Then
                                dcGolonganAsuransi.Text = "II"
                            ElseIf fKdKelasDitanggung = "03" Then
                                dcGolonganAsuransi.Text = "III"
                            End If
                           dcGolonganAsuransi_KeyPress (13)
                        End If
                        
                        mstrKelasDitanggung = arr(1)
                        dcKelasDitanggung.Text = arr(1)
                        nmKelas = arr(1)
                   Case "JENISPESERTA-KETERANGAN"
                        txtJenisPasien = arr(1)
                        nmJenisPeserta = arr(1)
                    Case "MR-PISA"
                        Call msubRecFO(rs, "SELECT KdHubungan FROM dbo.HubunganPesertaAsuransi WHERE KodeExternal='" & arr(1) & "'")
                        If rs.EOF = False Then
                            dcHubungan.BoundText = rs(0)
                        End If
                        pisa = arr(1)
                    Case "PROVUMUM-SEX"
                         sex = arr(1)
                    Case "STATUSPESERTA-TGLCETAKKART"
                         tglCetakKartu = arr(1)
                    Case "STATUSPESERTA-KETERANGAN"
                         statusPeserta = arr(1)
                    Case "keluhan"
                         txtCatatan.Text = arr(1)
                    Case "kdCabang"
                         kdCabang = arr(1)
                    Case "nmCabang"
                         nmCabang = arr(1)
                    Case "NOKUNJUNGAN"
                         txtNoRujukan.Text = arr(1)
                    Case "TGLKUNJUNGAN"
                         dtpTglDirujuk.value = CDate(Split(arr(1), " ")(0))
        End Select
    Next i
        mstrTglVerifBPJS = Format(dtpTglLahirPA.value, "yyyy-MM-dd 00:00:00")
        mstrkartuPeserta = txtNoKartuPA.Text
        mstrnorujukan = txtNoRujukan.Text
        mstrJenisPeserta = txtJenisPasien
        mstrNamaAsalRujukanMon = dcNamaAsalRujukan.Text
        mstrPpkRujukan = txtPpkRujukan.Text
        'mstrCatatatnBPJS = txtKeluhan.Text
        
'        If txtNamaPA.Text = "" Or UBound(result) = -1 Then
        If UBound(result) = 0 Then
'            MsgBox "Data Peserta Tidak Ditemukan,,,,!!!", vbInformation, "Validasi"
'            txtNoKartuPA.Text = ""
            blnKartuAktif = False
            MsgBox "Data Peserta Tidak Ditemukan,,,,!!!" & vbCrLf & Replace(result(0), "message:", ""), vbInformation, "Validasi"
            Debug.Print txtNoKartuPA.Text
            Debug.Print result(0)
'            txtNoKartuPA.Text = ""
'            txtJenisPasien.Text = ""
            
            
            'Set DgPasien2.DataSource = Nothing
            'Call SubloadPasienBPJS
            Exit Sub
        Else
            If sp_detailkartubpjs(noKartu, nik, nama, pisa, sex, tgllahir, tglCetakKartu, kdProvider, nmProvider, kdCabang, nmCabang, kdJenisPeserta, nmJenisPeserta, kdKelas, nmKelas, potensiprb) = False Then Exit Sub
        End If
        
        'Call SubloadPasienBPJS
        'DgPasien2.SetFocus
    'Else
    '    MsgBox "Sdk Bridging askes tidak di temukan"
   ' End If
   
'    If blnKartuAktif = True Then
'        MsgBox "Pasien tersebut aktif", vbInformation, "Cek Kepesertaan BPJS"
'    End If

Exit Sub
mulih:
MsgBox "Koneksi Bridging Bermasalah"
    
End Sub

Private Function PengajuanSEP(strNoKartu As String) As Boolean
On Error Resume Next
    
    PengajuanSEP = True
    If (Dir("C:\SDK\vclaim\result.tlb") <> "") Then
        Dim context As ContextVclaim
        Dim result() As String
        Set context = New ContextVclaim
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
        
        Dim strKodeInstalasiBPJS As String
            
            Select Case mstrKdInstalasi
                Case "03"
                    strKodeInstalasiBPJS = "1"
                Case "02"
                    strKodeInstalasiBPJS = "2"
                Case "01"
                    strKodeInstalasiBPJS = "2"
                Case "06"
                    strKodeInstalasiBPJS = "2"
                Case "22"
                    strKodeInstalasiBPJS = "2"
         End Select
         
'         strSQL = "SELECT dbo.getNamaPegawaiByIdPegawai('" & strIDPegawai & "')"
        strSQL = "SELECT NamaLengkap FROM DataPegawai WHERE IdPegawai='" & strIDPegawai & "'"
         Call msubRecFO(rs, strSQL)
         strNamaPegawai = rs(0).value
        
        result = context.PengajuanSEP(strNoKartu, Format(dtpTglSJP.value, "yyyy-MM-dd"), strKodeInstalasiBPJS, IIf(Len(Trim(txtCatatan.Text)) = 0, "-", Trim(txtCatatan.Text)), strNamaPegawai)
        
        For n = LBound(result) To UBound(result)
            arr = Split(result(n), ":")
            Select Case arr(0)
                Case "error"
                    MsgBox Replace(result(0), "error:", ""), vbExclamation, "Pengajuan SEP BPJS"
                    PengajuanSEP = False
                    fraPengajuanSEP.Visible = False
                    Exit Function
            End Select
        Next n

        MsgBox "Pengajuan SEP Berhasil" & vbCrLf & "No. Kartu: " & arr(1), vbInformation, "Pengajuan SEP BPJS"
    Else
        PengajuanSEP = False
        MsgBox "error", vbInformation, ""
    End If
End Function

Private Function sp_PengajuanSEPVclaim(f_status As String) As Boolean
On Error GoTo pesan
    
    Dim adoCommand As New ADODB.Command
    Set adoCommand = New ADODB.Command
    
    sp_PengajuanSEPVclaim = True
    MousePointer = vbHourglass
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, IIf(Len(Trim(txtNoPendaftaran.Text)) = 0, Null, txtNoPendaftaran.Text))
        .Parameters.Append .CreateParameter("NoKartu", adChar, adParamInput, 13, txtNoKartuPA.Text)
        .Parameters.Append .CreateParameter("NoSEP", adVarChar, adParamInput, 30, IIf(Len(Trim(txtNoSJP.Text)) = "", Null, txtNoSJP.Text))
        .Parameters.Append .CreateParameter("TglSEP", adDate, adParamInput, , Format(dtpTglSJP, "yyyy/MM/dd hh:mm:ss"))
        .Parameters.Append .CreateParameter("JnsPelayanan", adTinyInt, adParamInput, , CInt(mstrKdInstalasi))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 8000, IIf(Len(Trim(txtKetPengajuan.Text)) = 0, Null, Trim(txtKetPengajuan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("StatusPengajuan", adVarChar, adParamInput, 50, f_status)
        
        .ActiveConnection = dbConn
        .CommandText = "AU_PengajuanSEPVclaim"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data Persetujuan SEP", vbCritical, "Validasi"
            sp_PengajuanSEPVclaim = False
        Else
            Call Add_HistoryLoginActivity("AU_PengajuanSEPVclaim")
        End If
        
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    MousePointer = vbDefault
    
Exit Function
pesan:
    sp_PengajuanSEPVclaim = False
    Call msubPesanError
End Function

Private Function ApprovalPengajuanSEP(strNoKartu As String) As Boolean
On Error Resume Next
    
    ApprovalPengajuanSEP = True
    If (Dir("C:\SDK\vclaim\result.tlb") <> "") Then
        Dim context As ContextVclaim
        Dim result() As String
        Set context = New ContextVclaim
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
        
'        jnsPelayanan = IIf(mstrKdInstalasi = "03", "1", "2")
        
        Dim strKodeInstalasiBPJS As String
            
            Select Case mstrKdInstalasi
                Case "03"
                    strKodeInstalasiBPJS = "1"
                Case "02"
                    strKodeInstalasiBPJS = "2"
                Case "01"
                    strKodeInstalasiBPJS = "2"
                Case "06"
                    strKodeInstalasiBPJS = "2"
                Case "22"
                    strKodeInstalasiBPJS = "2"
         End Select
         
'         strSQL = "SELECT dbo.getNamaPegawaiByIdPegawai('" & strIDPegawai & "')"
         strSQL = "SELECT NamaLengkap FROM DataPegawai WHERE IdPegawai='" & strIDPegawai & "'"
         Call msubRecFO(rs, strSQL)
         strNamaPegawai = rs(0).value
        
        result = context.AprovalPengajuanSep(strNoKartu, Format(dtpTglSJP.value, "yyyy-MM-dd"), strKodeInstalasiBPJS, IIf(Len(Trim(txtCatatan.Text)) = 0, "-", Trim(txtCatatan.Text)), strNamaPegawai)
        
        For n = LBound(result) To UBound(result)
            arr = Split(result(n), ":")
            Select Case arr(0)
                Case "error"
                    MsgBox Replace(result(0), "error:", ""), vbExclamation, "Approval Pengajuan SEP BPJS"
                    ApprovalPengajuanSEP = False
                    Exit Function
            End Select
        Next n
        
        MsgBox "Approval Pengajuan SEP Berhasil" & vbCrLf & "No. Kartu: " & arr(1), vbInformation, "Approval Pengajuan SEP BPJS"
    Else
        ApprovalPengajuanSEP = False
        MsgBox "error", vbInformation, ""
    End If
End Function

Private Function UpdateTglPulang(strNoSEP As String) As Boolean
On Error Resume Next
    
    UpdateTglPulang = True
    If (Dir("C:\SDK\vclaim\result.tlb") <> "") Then
        Dim context As ContextVclaim
        Dim sep() As String
        Set context = New ContextVclaim
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
        
        strSQL = "SELECT dbo.getNamaPegawaiByIdPegawai('" & strIDPegawai & "')"
         Call msubRecFO(rs, strSQL)
         strNamaPegawai = rs(0).value
        
        
        sep = context.UpdateTgPulangSep(strNoSEP, Format(dtpTglPulang.value, "yyyy-MM-dd"), strNamaPegawai)
        
        For n = LBound(sep) To UBound(sep)
            arr = Split(sep(n), ":")
            Select Case arr(0)
                Case "error"
                    MsgBox arr(1), vbExclamation, "Update Tanggal Pulang Pasien BPJS"
                    UpdateTglPulang = False
                    Exit Function
            End Select
        Next n
        
        MsgBox "Update Tanggal Pulang berhasil" & vbCrLf & "Dengan No. SEP: " & arr(1), vbInformation, "Update Tanggal Pulang Pasien BPJS"
    Else
        MsgBox "error", vbCritical, ""
        UpdateTglPulang = False
    End If
End Function


Public Sub subLoadDataRujukan(sNoPen As String, sNoSEP As String)
On Error GoTo pesan
    
    With frmRujukanPasienBPJS
        strSQL = "Select NoRujukan From RujukanPasienVclaim where Nopendaftaran= '" & txtNoPendaftaran.Text & "'"
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = False Then
            .txtNoRujukan.Text = rs(0).value
        End If
        
        .txtNoCM.Text = txtNoCM.Text
        .txtNamaPasien.Text = txtNamaPasien.Text
        .txtNoPendaftaran.Text = txtNoPendaftaran.Text
        .txtNoKartu.Text = txtNoKartuPA.Text
        .txtJK.Text = txtJK.Text
        .txtJenisPasien.Text = txtJenisPasien.Text
        .txtPenjamin.Text = dcPenjamin.Text
        .txtJenisPasien.Text = txtJenisPasien.Text
        .txtKelas.Text = dcKelasDitanggung.Text
        .txtNoSEP.Text = txtNoSJP.Text
        .txtPelayanan.Text = mstrKdInstalasi
'        .txtDiagnosa.Text = txtDiagnosa.Text
        .txtdiagnosa.Text = dcDiagnosa.Text
        .dtpTglLahir.value = Format(dtpTglLahirPA.value, "dd/MM/yyyy")
        .dtpTglSEP.value = Format(dtpTglSJP.value, "dd/MM/yyyy")

    End With
    
Exit Sub
pesan:
    Call msubPesanError
End Sub

Private Sub CariSEP()
On Error GoTo hell
    
If (Dir("C:\SDK\Vclaim\result.tlb") <> "") Then
        Dim context As ContextVclaim
        Set context = New ContextVclaim
        Dim result() As String
        
        strSQL = "Select Value From SettingGlobal where Prefix In('ConsumerID','PasswordKey')"
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = False Then
            context.ConsumerID = rs(0).value
'            rs.MoveNext
'            context.KodeRumahSakit = rs(0).value
            rs.MoveNext
            context.PasswordKey = rs(0).value
        End If
        
        
         strSQL = "SELECT Value FROM SettingGlobal where Prefix='UrlGenerateSEP'"
            Call msubRecFO(rs, strSQL)
            Dim URL  As String
              If rs.EOF = False Then
                  URL = rs.Fields(0)
                  context.URL = URL
                 'Exit Sub
              End If
              
            result = context.CariSEP(txtNoSJP.Text)
        
        Dim i As Long
        For i = LBound(result) To UBound(result)
            Dim arr() As String
            arr = Split(result(i), ":")
            
            Select Case arr(0)
                    Case "noKartu"
                         blnKartuAktif = True
                         txtNoKartuPA.Text = arr(1)
                         noKartu = arr(1)
                         Debug.Print "NoKartu : " & arr(1)
                    Case "nik"
                         txtNipPA.Text = arr(1)
                         nik = arr(1)
                    Case "nama"
                         txtNamaPA.Text = arr(1)
                         nama = arr(1)
                    Case "nmProvider"
                         dcNamaAsalRujukan.Text = arr(1)
                         nmProvider = arr(1)
                    Case "tglLahir"
                         dtpTglLahirPA.value = CDate(Split(arr(1), " ")(0))
                         tgllahir = arr(1)
                    Case "kdProvider"
                         ppkRujukan = arr(1)
                         kdProvider = arr(1)
                         txtPpkRujukan.Text = arr(1)
                    Case "keterangan"
                         If txtTempHakKelas = "" Then
                         strSQL = "SELECT DeskKelas FROM KelasPelayanan where NamaExternal='" & arr(1) & "'"
                         Call msubRecFO(rs, strSQL)
                         If (rs.EOF = False) Then dcKelasDitanggung.BoundText = rs(0).value: fKdKelasDitanggung = rs(0).value
                         mstrKelasDitanggung = arr(1)
                         dcKelasDitanggung.Text = rs(0)
                         txtTempHakKelas.Text = arr(1)
                         nmKelas = arr(1)
'                             strSQL = "SELECT DISTINCT KdKelas, DeskKelas FROM V_KelasDitanggungPenjamin WHERE (IdPenjamin = '" & dcPenjamin.BoundText & "') AND KdKelompokPasien = '" & dcJenisPasien.BoundText & "' and DeskKelas = '" & dcKelasDitanggung.Text & "'"
'                             Call msubDcSource(dcKelasDitanggung, rs1, strSQL)
'                            dcKelasDitanggung.BoundText = rs1(0).value
'                            kdKelas = rs1(0).value
'                         Else
                         End If
                    Case "jenisPeserta"
                        Dim j As Long
                            arr = Split(result(i + 1), ":")
                            If arr(0) = "keterangan" Then
                                If txtJenisPasien.Text = "" Then
                                    txtJenisPasien = arr(1)
                                    keterangan = arr(1)
                                End If
                            End If
                    Case "pisa"
                         strSQL = "select NamaHubungan from HubunganPesertaAsuransi where KdHubungan= '0' + '" & arr(1) & "'"
                         Call msubRecFO(rs, strSQL)
                         If rs.EOF = False Then
                            dcHubungan.Text = rs(0).value
                         End If
                         pisa = arr(1)
                    Case "sex"
                         sex = arr(1)
                    Case "kdCabang"
                        kdCabang = arr(1)
                    Case "nmCabang"
                        nmCabang = arr(1)
                    Case "kdKelas"
                        kdKelas = arr(1)
'                    Case "nmKelas"
'                         strSQL = "SELECT KdKelas FROM KelasPelayanan where DeskKelas='" & arr(1) & "'"
'                          Call msubRecFO(rs, strSQL)
'                          If (rs.EOF = False) Then dcKelasDitanggung.BoundText = rs(0).value
'                         mstrKelasDitanggung = arr(1)
'                         dcKelasDitanggung.Text = arr(1)
'                         Case "nmJenisPeserta"
'                         'dcJenisPasien.Text = arr(1)
'                         txtJenisPasien.Text = arr(1)
            End Select
        Next i
        
        
        If UBound(result) = 0 Then
            blnKartuAktif = False
            MsgBox "Data SEP tidak ditemukan" & vbCrLf & Replace(result(0), "message:", ""), vbInformation, "Validasi"
            Debug.Print txtNoKartuPA.Text
            Debug.Print result(0)
            txtJenisPasien.Text = ""
            Exit Sub
        Else
            If sp_detailkartubpjs(noKartu, nik, nama, pisa, sex, tgllahir, tglCetakKartu, kdProvider, nmProvider, kdCabang, nmCabang, kdJenisPeserta, nmJenisPeserta, kdKelas, nmKelas, potensiprb) = False Then Exit Sub
        End If
        
'        If txtNamaPA.Text = "" Then
'            MsgBox "Data Peserta Tidak Ditemukan,,,,!!!", vbInformation, "Validasi"
'            txtNoKartuPA.Text = ""
'            'Set DgPasien2.DataSource = Nothing
'            'Call SubloadPasienBPJS
'            Exit Sub
'        End If
        
        'Call SubloadPasienBPJS
        'DgPasien2.SetFocus
    Else
        MsgBox "Sdk Bridging askes tidak di temukan"
    End If
Exit Sub
hell:
MsgBox "Koneksi Bridging Bermasalah"
End Sub

Sub fillGridWithPropinsi(vFG As MSFlexGrid, vResult() As String)
    With vFG
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .cols = 1
        .rows = 2
        
        Dim row, col, rows, cols, i As Integer
        
        row = 1
        For i = 0 To UBound(vResult)
            Dim arrResult() As String
            arrResult = Split(vResult(i), ":")
            col = isHeaderExist(vFG, arrResult(0))
            If col > -1 Then
                If .TextMatrix(row, col) <> "" Then
                    .rows = .rows + 1
                    row = .rows - 1
                End If
                .TextMatrix(row, col) = arrResult(1)
            Else
                col = .cols - 1
                .TextMatrix(0, col) = arrResult(0)
                .TextMatrix(row, col) = arrResult(1)
                .cols = .cols + 1
            End If
        Next i
        .ColWidth(col) = 7000
        .ColWidth(0) = 0
        .cols = .cols - 1
    End With
End Sub

Function isHeaderExist(vFG As MSFlexGrid, strHeader As String) As Integer
    isHeaderExist = -1
    With vFG
        Dim col As Integer
        For col = 0 To vFG.cols - 1
            If UCase(.TextMatrix(0, col)) = UCase(strHeader) Then
                isHeaderExist = col
                Exit Function
            End If
        Next col
    End With
End Function

Private Sub GenerateSEPBPJSNew()
If chkNoSJP.value = vbChecked Then
             
          If (Dir("C:\SDK\Vclaim\result.tlb") <> "") Then
            Dim context As ContextVclaim
            Set context = New ContextVclaim
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
            
            mstrnorujukan = txtNoRujukan.Text
           
           Dim kodeKelas As Integer
           If (dcKelasDitanggung.Text = "Kelas I") Then
            kodeKelas = 1
           ElseIf (dcKelasDitanggung.Text = "Kelas II") Then
            kodeKelas = 2
           Else
            kodeKelas = 3
           End If
           
            strKdDiagnosa = Trim(dcDiagnosa.BoundText)
           
            Dim rsDataPoliBPJS As New ADODB.recordset
            Dim qryPoliBPJS As String
            Dim strKdDataPoliBPJS As String
            
            qryPoliBPJS = "SELECT KodeExternal FROM dbo.Ruangan WHERE KdRuangan='" & mstrKdRuanganPasien & "'"
            Call msubRecFO(rsDataPoliBPJS, qryPoliBPJS)
            If rsDataPoliBPJS.EOF = False Then
                strKdDataPoliBPJS = IIf(IsNull(rsDataPoliBPJS(0)) = True, "", rsDataPoliBPJS(0))
            Else
                strKdDataPoliBPJS = "" 'nilai default kalau belum dimapping
            End If
            
            Dim strIDPegawaiLokal As String
            strIDPegawaiLokal = strIDPegawai
            If URL = "http://dvlp.bpjs-kesehatan.go.id:8081/devWSLokalRest/" Then
                strIDPegawaiLokal = "88888888"
            End If
            
            strSQL = "SELECT Value FROM SettingGlobal WHERE Prefix='KodeRS'"
            Call msubRecFO(rs, strSQL)
            txtppkpelayanan.Text = rs(0).value
            
            strSQL = "SELECT Telepon FROM Pasien WHERE NoCM='" & txtNoCM.Text & "'"
            Call msubRecFO(rs, strSQL)
            txtNoTlpPasien.Text = IIf(IsNull(rs(0).value), "", Trim(rs(0).value))
            
            strSQL = "SELECT NamaLengkap FROM DataPegawai WHERE IdPegawai='" & strIDPegawai & "'"
            Call msubRecFO(rs, strSQL)
            strNamaPegawai = rs(0).value
            
            
            lakalantas = IIf(chkLakalantas.value = vbChecked, "1", "0")
            lokasiLaka = IIf(Len(Trim(txtKec.Text)) = 0, "0", Trim(txtKec.Text))
            penjaminLakalantas = 0
            If chkJasaRaharja.value = vbChecked And chkBPJSKK.value = vbUnchecked And chkTaspen.value = vbUnchecked And chkAsabri.value = vbUnchecked Then
                penjaminLakalantas = "1"
            ElseIf chkJasaRaharja.value = vbUnchecked And chkBPJSKK.value = vbChecked And chkTaspen.value = vbUnchecked And chkAsabri.value = vbUnchecked Then
                penjaminLakalantas = "2"
            ElseIf chkJasaRaharja.value = vbUnchecked And chkBPJSKK.value = vbUnchecked And chkTaspen.value = vbChecked And chkAsabri.value = vbUnchecked Then
                penjaminLakalantas = "3"
            ElseIf chkJasaRaharja.value = vbUnchecked And chkBPJSKK.value = vbUnchecked And chkTaspen.value = vbUnchecked And chkAsabri.value = vbChecked Then
                penjaminLakalantas = "4"
            ElseIf chkJasaRaharja.value = vbChecked And chkBPJSKK.value = vbChecked And chkTaspen.value = vbUnchecked And chkAsabri.value = vbUnchecked Then
                penjaminLakalantas = "1,2"
            ElseIf chkJasaRaharja.value = vbChecked And chkBPJSKK.value = vbUnchecked And chkTaspen.value = vbChecked And chkAsabri.value = vbChecked Then
                penjaminLakalantas = "1,3"
            ElseIf chkJasaRaharja.value = vbChecked And chkBPJSKK.value = vbUnchecked And chkTaspen.value = vbUnchecked And chkAsabri.value = vbChecked Then
                penjaminLakalantas = "1,4"
            ElseIf chkJasaRaharja.value = vbUnchecked And chkBPJSKK.value = vbChecked And chkTaspen.value = vbChecked And chkAsabri.value = vbUnchecked Then
                penjaminLakalantas = "2,3"
            ElseIf chkJasaRaharja.value = vbUnchecked And chkBPJSKK.value = vbChecked And chkTaspen.value = vbUnchecked And chkAsabri.value = vbChecked Then
                penjaminLakalantas = "2,4"
            ElseIf chkJasaRaharja.value = vbUnchecked And chkBPJSKK.value = vbUnchecked And chkTaspen.value = vbChecked And chkAsabri.value = vbChecked Then
                penjaminLakalantas = "3,4"
            ElseIf chkJasaRaharja.value = vbChecked And chkBPJSKK.value = vbChecked And chkTaspen.value = vbChecked And chkAsabri.value = vbUnchecked Then
                penjaminLakalantas = "1,2,3"
            ElseIf chkJasaRaharja.value = vbChecked And chkBPJSKK.value = vbChecked And chkTaspen.value = vbUnchecked And chkAsabri.value = vbUnchecked Then
                penjaminLakalantas = "1,2,4"
            ElseIf chkJasaRaharja.value = vbChecked And chkBPJSKK.value = vbChecked And chkTaspen.value = vbChecked And chkAsabri.value = vbChecked Then
                penjaminLakalantas = "1,2,3,4"
            End If
            
            cob = 0
            If chkCob.value = vbChecked Then
                cob = "1"
            Else
                cob = 0
            End If
            
            katarak = 0
            If chkKatarak.value = vbChecked Then
                katarak = "1"
            Else
                katarak = 0
            End If
            
            suplesi = 0
            If chkSuplesi.value = vbChecked Then
                suplesi = "1"
            Else
                suplesi = 0
            End If
'            Diganti karena mstrNoSJPNew Harus String Array Ales
            
'                mstrNoSJPNew = context.InsertSep(txtNoKartuPA.Text, Format(dtpTglSJP.value, "yyyy-mm-dd"), txtppkpelayanan.Text, IIf(mstrKdInstalasi = "03", "1", "2"), kodeKelas, txtNoCM.Text, _
'                            1, Format(dtpTglDirujuk.value, "yyyy-mm-dd"), txtNoRujukan.Text, ppkRujukan, txtCatatan.Text, strKdDiagnosa, IIf(mstrKdInstalasi, strKdDataPoliBPJS, ""), _
'                            0, cob, lakalantas, penjaminLakalantas, lokasiLaka, txtNoTlpPasien.Text, strNamaPegawai)
            
            
                
                mstrNoSJPNew = context.InsertSepV1_1(txtNoKartuPA, Format(dtpTglSJP.value, "yyyy-mm-dd"), txtppkpelayanan.Text, IIf(mstrKdInstalasi = "03", "1", "2"), kodeKelas, txtNoCM.Text, _
                             IIf(mstrKdInstalasi = "03", 2, 1), Format(dtpTglDirujuk.value, "yyyy-mm-dd"), txtNoRujukan.Text, IIf(mstrKdInstalasi = "03", "0310r001", ppkRujukan), txtCatatan.Text, strKdDiagnosa, strKdDataPoliBPJS, _
                             0, cob, katarak, lakalantas, penjaminLakalantas, Format(dtpTglKejadian.value, "yyyy-mm-dd"), txtKet.Text, suplesi, txtSuplesi.Text, txtKdPropinsi.Text, txtKdKota.Text, txtKdKec.Text, _
                             txtNoSKDP, txtKdDPJP, txtNoTlpPasien.Text, strNamaPegawai)
                             
            Dim Temp As String
            Dim i As Integer
            For i = LBound(mstrNoSJPNew) To UBound(mstrNoSJPNew)
                arr = Split(mstrNoSJPNew(i), ":")
                Temp = Temp & vbCrLf & mstrNoSJPNew(i)
                If arr(0) = "error" Then
                    MsgBox Replace(mstrNoSJPNew(0), "error:", "HUBUNGI BPJS "), vbExclamation, "Generate SEP BPJS"
                    txtNoSJP.Text = "-"
                End If
                If UCase(Trim(Right(arr(0), 5))) = "NOSEP" Or UCase(Trim(Right(arr(0), 5))) = "UMUR-NOSEP" Then
                    mstrNoSJP = arr(1)
                    Exit For
                End If
             Next i
             
                    If mstrNoSJP <> "" Then
                        
                        Call txtNoSJP_KeyPress(13)
                        txtNoSJP.Text = mstrNoSJP
                        cmdSimpan.Enabled = False
                        typAsuransi.blnSuksesAsuransi = True
                    Else
                    MsgBox mstrNoSJPNew
             End If
            
        Else
            MsgBox "Generate SEP Gagal,,,,!!!"
            typAsuransi.blnSuksesAsuransi = False
            bolGenerateSEPSukse = False
            txtNoSJP.Text = ""
        End If
End If
Exit Sub
SepEndPoint:

End Sub

Private Sub CariRujukanRSByNoKartu(vNoRujukan As String)
On Error GoTo mulih

    Dim context As ContextVclaim
    Set context = New ContextVclaim
    
    Dim result() As String
    Dim URL As String
    Dim i As Long
    
    strSQL = "select value from SettingGlobal where Prefix in ('ConsumerID','PasswordKey')"
    Call msubRecFO(rs, strSQL)
    
    If rs.EOF = False Then
        context.ConsumerID = rs(0).value
        rs.MoveNext
        context.PasswordKey = rs(0).value
    End If
    
    strSQL = "select value from SettingGlobal where Prefix='UrlGenerateSEP'"
    Call msubRecFO(rs, strSQL)
    
    If rs.EOF = False Then
        URL = rs.Fields(0)
        context.URL = (URL)
    End If
    
    txtNamaPA.Text = ""
    result = context.RujukanRsByNoKartu(vNoRujukan)
    
    For i = LBound(result) To UBound(result)
        Debug.Print (result(i))
        Dim arr() As String
        arr = Split(result(i), ":")
        Select Case arr(0)
                    Case "PROVPERUJUK-TGLKUNJUNGAN"
                        dtpTglDirujuk.value = CDate(Split(arr(1), " ")(0))
                    Case "DIAGNOSA-KODE"
                        dcDiagnosa.Text = arr(1)
                        dcDiagnosa_KeyPress (13)
                    Case "DIAGNOSA-NOKUNJUNGAN"
                        txtNoRujukan.Text = arr(1)
                    Case "MR-NOKARTU"
                         blnKartuAktif = True
                         txtNoKartuPA.Text = arr(1)
                         noKartu = arr(1)
                         Debug.Print "NoKartu : " & arr(1)
                    Case "MR-NIK"
                         txtNipPA.Text = arr(1)
                         nik = arr(1)
                    Case "MR-NAMA"
                         txtNamaPA.Text = arr(1)
                         nama = arr(1)
                    Case "PROVPERUJUK-NAMA"
                         dcNamaAsalRujukan.Text = arr(1)
                         nmProvider = arr(1)
                    Case "STATUSPESERTA-TGLLAHIR"
                        If dtpTglLahirPA.value <> CDate(Split(arr(1), " ")(0)) Then
                            MsgBox "Tanggal lahir tidak sama!" & "Tanggal Lahir Peserta BPJS: " & arr(1) & vbCrLf & "Silakan cek data pasien", vbOKOnly, "Cek Kepesertaan"
                        End If
                        dtpTglLahirPA.value = CDate(Split(arr(1), " ")(0))
                        tgllahir = arr(1)
                    Case "PROVPERUJUK-KODE"
                         ppkRujukan = arr(1)
                         kdProvider = arr(1)
                         txtPpkRujukan.Text = arr(1)
                    Case "HAKKELAS-KODE"
                        kdKelas = arr(1)
                    Case "HAKKELAS-KETERANGAN"
                        strSQL = "SELECT KdKelas FROM KelasPelayanan where NamaExternal='" & arr(1) & "'"
                        Call msubRecFO(rs, strSQL)
                     
                        If (rs.EOF = False) Then
                           fKdKelasDitanggung = rs(0).value
                           dcKelasDitanggung.BoundText = fKdKelasDitanggung
                        End If
                        
                        mstrKelasDitanggung = arr(1)
                        dcKelasDitanggung.Text = arr(1)
                        nmKelas = arr(1)
                   Case "JENISPESERTA-KETERANGAN"
                        txtJenisPasien = arr(1)
                        nmJenisPeserta = arr(1)
                    Case "MR-PISA"
                        Call msubRecFO(rs, "SELECT KdHubungan FROM dbo.HubunganPesertaAsuransi WHERE KodeExternal='" & arr(1) & "'")
                        If rs.EOF = False Then
                            dcHubungan.BoundText = rs(0)
                        End If
                        pisa = arr(1)
                    Case "PROVUMUM-SEX"
                         sex = arr(1)
                    Case "STATUSPESERTA-TGLCETAKKART"
                         tglCetakKartu = arr(1)
                    Case "STATUSPESERTA-KETERANGAN"
                         statusPeserta = arr(1)
                    Case "keluhan"
                         txtCatatan.Text = arr(1)
                    Case "kdCabang"
                         kdCabang = arr(1)
                    Case "nmCabang"
                         nmCabang = arr(1)
                    Case "NOKUNJUNGAN"
                         txtNoRujukan.Text = arr(1)
                    Case "TGLKUNJUNGAN"
                         dtpTglDirujuk.value = CDate(Split(arr(1), " ")(0))
                    Case "INFORMASI-PROLANISPRB"
                         potensiprb = arr(1)
        End Select
    Next i
        mstrTglVerifBPJS = Format(dtpTglLahirPA.value, "yyyy-MM-dd 00:00:00")
        mstrkartuPeserta = txtNoKartuPA.Text
        mstrnorujukan = txtNoRujukan.Text
        mstrJenisPeserta = txtJenisPasien
        mstrNamaAsalRujukanMon = dcNamaAsalRujukan.Text
        mstrPpkRujukan = txtPpkRujukan.Text
        'mstrCatatatnBPJS = txtKeluhan.Text
        
'        If txtNamaPA.Text = "" Or UBound(result) = -1 Then
        If UBound(result) = 0 Then
'            MsgBox "Data Peserta Tidak Ditemukan,,,,!!!", vbInformation, "Validasi"
'            txtNoKartuPA.Text = ""
            blnKartuAktif = False
            MsgBox "Data Peserta Tidak Ditemukan,,,,!!!" & vbCrLf & Replace(result(0), "message:", ""), vbInformation, "Validasi"
            Debug.Print txtNoKartuPA.Text
            Debug.Print result(0)
'            txtNoKartuPA.Text = ""
'            txtJenisPasien.Text = ""
            
            
            'Set DgPasien2.DataSource = Nothing
            'Call SubloadPasienBPJS
            Exit Sub
        Else
            If sp_detailkartubpjs(noKartu, nik, nama, pisa, sex, tgllahir, tglCetakKartu, kdProvider, nmProvider, kdCabang, nmCabang, kdJenisPeserta, nmJenisPeserta, kdKelas, nmKelas, potensiprb) = False Then Exit Sub
        End If
        
        'Call SubloadPasienBPJS
        'DgPasien2.SetFocus
    'Else
    '    MsgBox "Sdk Bridging askes tidak di temukan"
   ' End If
   
'    If blnKartuAktif = True Then
'        MsgBox "Pasien tersebut aktif", vbInformation, "Cek Kepesertaan BPJS"
'    End If

Exit Sub
mulih:
MsgBox "Koneksi Bridging Bermasalah"
    
End Sub
