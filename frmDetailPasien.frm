VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDetailPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Detail Pasien"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDetailPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   9150
   Begin VB.TextBox txtFormPengirim 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      MaxLength       =   50
      TabIndex        =   73
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   53
      Top             =   8880
      Width           =   9135
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   6120
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   7560
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Alamat Keluarga Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      TabIndex        =   44
      Top             =   6360
      Width           =   9135
      Begin MSMask.MaskEdBox meRTRWKel 
         Height          =   330
         Left            =   5160
         TabIndex        =   23
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         Mask            =   "##/##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtAlamatKel 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         MaxLength       =   100
         TabIndex        =   22
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox txtKodePos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8040
         MaxLength       =   5
         TabIndex        =   29
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtTeleponKel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6000
         MaxLength       =   15
         TabIndex        =   24
         Top             =   600
         Width           =   2895
      End
      Begin MSDataListLib.DataCombo dcKota 
         Height          =   330
         Left            =   4320
         TabIndex        =   26
         Top             =   1320
         Width           =   4575
         _ExtentX        =   8070
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
      Begin MSDataListLib.DataCombo dcKecamatan 
         Height          =   330
         Left            =   240
         TabIndex        =   27
         Top             =   2040
         Width           =   4215
         _ExtentX        =   7435
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
      Begin MSDataListLib.DataCombo dcKelurahan 
         Height          =   330
         Left            =   4560
         TabIndex        =   28
         Top             =   2040
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSDataListLib.DataCombo dcPropinsi 
         Height          =   330
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   3975
         _ExtentX        =   7011
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Alamat Lengkap"
         Height          =   210
         Left            =   240
         TabIndex        =   52
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kota/Kabupaten"
         Height          =   210
         Left            =   4320
         TabIndex        =   51
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Propinsi"
         Height          =   210
         Left            =   240
         TabIndex        =   50
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Kecamatan"
         Height          =   210
         Left            =   240
         TabIndex        =   49
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Kelurahan"
         Height          =   210
         Left            =   4560
         TabIndex        =   48
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "RT/RW"
         Height          =   210
         Left            =   5160
         TabIndex        =   47
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Kode Pos"
         Height          =   210
         Left            =   8040
         TabIndex        =   46
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Telepon"
         Height          =   210
         Left            =   6000
         TabIndex        =   45
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Frame frpatient 
      Caption         =   "Data Keluarga Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      TabIndex        =   40
      Top             =   3840
      Width           =   9135
      Begin VB.TextBox txtKepalaKeluarga 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   71
         Top             =   2040
         Width           =   3735
      End
      Begin VB.TextBox txtNoKK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         MaxLength       =   50
         TabIndex        =   69
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox TxtIstriSuami 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6120
         MaxLength       =   50
         TabIndex        =   21
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox TxtIbu 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   20
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox TxtAyah 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         MaxLength       =   50
         TabIndex        =   19
         Top             =   1320
         Width           =   2775
      End
      Begin VB.ComboBox cboJnsKelaminKel 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "frmDetailPasien.frx":0CCA
         Left            =   3480
         List            =   "frmDetailPasien.frx":0CD4
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNamaKelPasien 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         MaxLength       =   50
         TabIndex        =   15
         Top             =   600
         Width           =   3135
      End
      Begin MSDataListLib.DataCombo dcPekerjaanKel 
         Height          =   330
         Left            =   4920
         TabIndex        =   17
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
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
      Begin MSDataListLib.DataCombo dcHubungan 
         Height          =   330
         Left            =   7080
         TabIndex        =   18
         Top             =   600
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
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Nama Kepala Keluarga"
         Height          =   210
         Left            =   3120
         TabIndex        =   72
         Top             =   1800
         Width           =   1785
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "No. KK"
         Height          =   210
         Left            =   240
         TabIndex        =   70
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Nama Suami / Istri"
         Height          =   210
         Left            =   6120
         TabIndex        =   67
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Nama Ibu"
         Height          =   210
         Left            =   3120
         TabIndex        =   66
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Nama Ayah"
         Height          =   210
         Left            =   240
         TabIndex        =   65
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pekerjaan"
         Height          =   210
         Left            =   4920
         TabIndex        =   54
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblTmpLhr 
         AutoSize        =   -1  'True
         Caption         =   "Hubungan"
         Height          =   210
         Left            =   7080
         TabIndex        =   43
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   3480
         TabIndex        =   42
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nama Keluarga"
         Height          =   210
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Data Detail Pasien"
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
      TabIndex        =   55
      Top             =   2040
      Width           =   9135
      Begin VB.TextBox txtNamaKeluarga 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         MaxLength       =   50
         TabIndex        =   6
         Top             =   600
         Width           =   3255
      End
      Begin VB.ComboBox cboGD 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "frmDetailPasien.frx":0CEE
         Left            =   3600
         List            =   "frmDetailPasien.frx":0CFE
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chkRhesus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cboStsPernikahan 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "frmDetailPasien.frx":0D0F
         Left            =   5280
         List            =   "frmDetailPasien.frx":0D1F
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cboWarganegara 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "frmDetailPasien.frx":0D40
         Left            =   6480
         List            =   "frmDetailPasien.frx":0D4A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1320
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo dcAgama 
         Height          =   330
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
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
      Begin MSDataListLib.DataCombo dcSuku 
         Height          =   330
         Left            =   2400
         TabIndex        =   12
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
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
      Begin MSDataListLib.DataCombo dcPendidikan 
         Height          =   330
         Left            =   4560
         TabIndex        =   13
         Top             =   1320
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
      Begin MSDataListLib.DataCombo dcPekerjaan 
         Height          =   330
         Left            =   6960
         TabIndex        =   10
         Top             =   600
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
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Pendidikan"
         Height          =   210
         Left            =   4560
         TabIndex        =   64
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Warga Negara"
         Height          =   210
         Left            =   6480
         TabIndex        =   63
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Suku Bangsa"
         Height          =   210
         Left            =   2400
         TabIndex        =   62
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Agama / Kepercayaan"
         Height          =   210
         Left            =   240
         TabIndex        =   61
         Top             =   1080
         Width           =   1785
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Status Pernikahan"
         Height          =   210
         Left            =   5280
         TabIndex        =   60
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nama Keluarga"
         Height          =   210
         Left            =   240
         TabIndex        =   59
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label lblRhesus 
         AutoSize        =   -1  'True
         Caption         =   "Rhesus"
         Height          =   210
         Left            =   4560
         TabIndex        =   58
         Top             =   360
         Width           =   570
      End
      Begin VB.Label lblGolDrh 
         AutoSize        =   -1  'True
         Caption         =   "Gol. Darah"
         Height          =   210
         Left            =   3600
         TabIndex        =   57
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pekerjaan"
         Height          =   210
         Left            =   6960
         TabIndex        =   56
         Top             =   360
         Width           =   795
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
      Height          =   975
      Left            =   0
      TabIndex        =   32
      Top             =   1080
      Width           =   9135
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         MaxLength       =   12
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4320
         MaxLength       =   9
         TabIndex        =   2
         Top             =   480
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
         Left            =   5520
         TabIndex        =   33
         Top             =   210
         Width           =   3375
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   480
            MaxLength       =   6
            TabIndex        =   3
            Top             =   188
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            MaxLength       =   6
            TabIndex        =   4
            Top             =   188
            Width           =   375
         End
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   2280
            MaxLength       =   6
            TabIndex        =   5
            Top             =   188
            Width           =   375
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Left            =   960
            TabIndex        =   36
            Top             =   240
            Width           =   285
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Left            =   1920
            TabIndex        =   35
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Left            =   2760
            TabIndex        =   34
            Top             =   240
            Width           =   165
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   1680
         TabIndex        =   38
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   4320
         TabIndex        =   37
         Top             =   240
         Width           =   1065
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
      Left            =   7320
      Picture         =   "frmDetailPasien.frx":0D60
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDetailPasien.frx":1AE8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDetailPasien.frx":44A9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDetailPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim j As Integer

Dim varPropinsi As String
Dim varKota As String
Dim varKecamatan As String
Dim varKelurahan As String
Private Sub cboGD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then chkRhesus.SetFocus
End Sub

Private Sub cboJnsKelaminKel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dcPekerjaanKel.SetFocus
End Sub

Private Sub cboStsPernikahan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dcPekerjaan.SetFocus
End Sub

Private Sub cboWarganegara_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtNamaKelPasien.SetFocus
End Sub

Private Sub chkRhesus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboStsPernikahan.SetFocus
End Sub

Private Sub cmdSimpan_Click()
    
    If cboStsPernikahan.Text <> "" Then
        If Periksa("combobox", cboStsPernikahan, "Status Pernikahan Tidak Terdaftar") = False Then Exit Sub
    End If
    
    If dcPekerjaan.Text <> "" Then
        If Periksa("datacombo", dcPekerjaan, "Pekerjaan Tidak Terdaftar") = False Then Exit Sub
    End If
    
    If dcAgama.Text <> "" Then
        If Periksa("datacombo", dcAgama, "Agama Tidak Terdaftar") = False Then Exit Sub
    End If
    
    If dcSuku.Text <> "" Then
        If Periksa("datacombo", dcSuku, "Suku Tidak Terdaftar") = False Then Exit Sub
    End If
        
    If dcPendidikan.Text <> "" Then
        If Periksa("datacombo", dcPendidikan, "Pendidikan Tidak Terdaftar") = False Then Exit Sub
    End If
    
    If cboWarganegara.Text <> "" Then
        If Periksa("combobox", cboWarganegara, "Warga Negara Tidak Terdaftar") = False Then Exit Sub
    End If
        
    If dcPekerjaanKel.Text <> "" Then
        If Periksa("datacombo", dcPekerjaanKel, "Pekerjaan Keluarga Tidak Terdaftar") = False Then Exit Sub
    End If
    
    If dcHubungan.Text <> "" Then
        If Periksa("datacombo", dcHubungan, "Hubungan Tidak Terdaftar") = False Then Exit Sub
    End If
    
    If dcKecamatan.Text <> "" Then
        If Periksa("datacombo", dcKecamatan, "Kecamatan Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcKelurahan.Text <> "" Then
        If Periksa("datacombo", dcKelurahan, "Kelurahan Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcKota.Text <> "" Then
        If Periksa("datacombo", dcKota, "Kota Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcPropinsi.Text <> "" Then
    If Periksa("datacombo", dcPropinsi, "Provinsi Tidak Terdaftar") = False Then Exit Sub
    End If
    
    If funcCekValidasi("DetailPasien") = False Then Exit Sub
    If Periksa("text", txtNamaKelPasien, "Keluarga Pasien Masih Kosong..") = False Then Exit Sub
    If txtNamaKelPasien.Text <> "" Then
        If funcCekValidasi("KeluargaPasien") = False Then Exit Sub
        Call sp_KeluargaPasien(dbcmd)
    End If
    Call sp_DetailPasien(dbcmd)
    MsgBox "Data detail pasien berhasil disimpan..", vbInformation, "Informasi"
    cmdSimpan.Enabled = False

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcAgama_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcAgama.BoundText
    strSQL = "SELECT Agama FROM Agama where StatusEnabled='1' Order By Agama"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcAgama.RowSource = rs
    dcAgama.ListField = rs.Fields(0).Name
    Set rs = Nothing
    dcAgama.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcAgama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then
        If dcAgama.MatchedWithList = True Then dcSuku.SetFocus
        strSQL = "SELECT kdAgama, agama FROM Agama where StatusEnabled='1' and (agama LIKE '%" & dcAgama.Text & "%')Order By Agama "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcAgama.Text = ""
            dcAgama.SetFocus
            Exit Sub
        End If
        dcAgama.BoundText = rs(0).value
        dcAgama.Text = rs(1).value
    End If
End Sub

Private Sub dcAgama_LostFocus()
    If dcAgama.Text = "" Then Exit Sub
    If dcAgama.MatchedWithList = False Then dcAgama.Text = "": dcAgama.SetFocus

End Sub
Private Sub dcHubungan_GotFocus()
    strSQL = "SELECT NamaHubungan FROM HubunganKeluarga where StatusEnabled='1'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcHubungan.RowSource = rs
    dcHubungan.ListField = rs.Fields(0).Name
    Set rs = Nothing
End Sub

Private Sub dcHubungan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then
        If dcHubungan.MatchedWithList = True Then TxtAyah.SetFocus
        strSQL = "SELECT Hubungan, NamaHubungan FROM HubunganKeluarga where StatusEnabled='1' and (NamaHubungan LIKE '%" & dcHubungan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcHubungan.Text = ""
            dcHubungan.SetFocus
            Exit Sub
        End If
        dcHubungan.BoundText = rs(0).value
        dcHubungan.Text = rs(1).value
    End If
End Sub

Private Sub dcHubungan_LostFocus()
    If dcHubungan.Text = "" Then Exit Sub
    If dcHubungan.MatchedWithList = False Then dcHubungan.Text = "": dcHubungan.SetFocus

End Sub

'Private Sub dcKecamatan_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Call subLoadDataWilayah("kecamatan")
'        dcKelurahan.SetFocus
'    End If
'End Sub
Private Sub dcKecamatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        j = 3
        dcKelurahan.Enabled = True
        Call subLoadDataWilayah("kecamatan")
        If dcKelurahan.Enabled = True Then
            dcKelurahan.SetFocus
        Else

        End If
    End If
End Sub

Private Sub dcKecamatan_LostFocus()
    dcKecamatan = Trim(StrConv(dcKecamatan, vbProperCase))
End Sub

'Private Sub dcKelurahan_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Call subLoadDataWilayah("desa")
'        txtKodePos.SetFocus
'    End If
'End Sub

Private Sub dcKelurahan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        j = 4
        Call subLoadDataWilayah("desa")
        txtKodePos.SetFocus
    End If
End Sub

Private Sub dcKelurahan_LostFocus()
    dcKelurahan = Trim(StrConv(dcKelurahan, vbProperCase))
End Sub

'Private Sub dcKota_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Call subLoadDataWilayah("kota")
'        dcKecamatan.SetFocus
'    End If
'End Sub

Private Sub dcKota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        j = 2
        dcKecamatan.Enabled = True
        dcKelurahan.Enabled = True
        Call subLoadDataWilayah("kota")
        If dcKecamatan.Enabled = True Then
            dcKecamatan.SetFocus
        End If
    End If
End Sub

Private Sub dcKota_LostFocus()
    dcKota = Trim(StrConv(dcKota, vbProperCase))
End Sub

Private Sub dcPekerjaan_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcPekerjaan.BoundText
    strSQL = "SELECT Pekerjaan FROM Pekerjaan where StatusEnabled='1' Order By Pekerjaan"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcPekerjaan.RowSource = rs
    dcPekerjaan.ListField = rs.Fields(0).Name
    Set rs = Nothing
    dcPekerjaan.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPekerjaan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then
        If dcPekerjaan.MatchedWithList = True Then dcAgama.SetFocus
        strSQL = "SELECT kdPekerjaan,pekerjaan FROM Pekerjaan where StatusEnabled='1'  and (Pekerjaan LIKE '%" & dcPekerjaan.Text & "%')Order By Pekerjaan"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcPekerjaan.Text = ""
            dcPekerjaan.SetFocus
            Exit Sub
        End If
        dcPekerjaan.BoundText = rs(0).value
        dcPekerjaan.Text = rs(1).value
    End If
End Sub

Private Sub dcPekerjaan_LostFocus()
    If dcPekerjaan.Text = "" Then Exit Sub
    If dcPekerjaan.MatchedWithList = False Then dcPekerjaan.Text = "": dcPekerjaan.SetFocus
End Sub

Private Sub dcPekerjaanKel_GotFocus()
    strSQL = "SELECT Pekerjaan FROM Pekerjaan where StatusEnabled='1'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcPekerjaanKel.RowSource = rs
    dcPekerjaanKel.ListField = rs.Fields(0).Name
    Set rs = Nothing
End Sub

Private Sub dcPekerjaanKel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then
        If dcPekerjaanKel.MatchedWithList = True Then dcHubungan.SetFocus
        strSQL = "SELECT kdPekerjaan, pekerjaan FROM Pekerjaan where StatusEnabled='1' and (pekerjaan LIKE '%" & dcPekerjaanKel.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcPekerjaanKel.Text = ""
            dcPekerjaanKel.SetFocus
            Exit Sub
        End If
        dcPekerjaanKel.BoundText = rs(0).value
        dcPekerjaanKel.Text = rs(1).value
    End If
End Sub

Private Sub dcPekerjaanKel_LostFocus()
    If dcPekerjaanKel.Text = "" Then Exit Sub
    If dcPekerjaanKel.MatchedWithList = False Then dcPekerjaanKel.Text = "": dcPekerjaanKel.SetFocus

End Sub

Private Sub dcPendidikan_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcPendidikan.BoundText
    strSQL = "SELECT Pendidikan FROM Pendidikan where StatusEnabled='1' order By Pendidikan"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcPendidikan.RowSource = rs
    dcPendidikan.ListField = rs.Fields(0).Name
    Set rs = Nothing
    dcPendidikan.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then
        If dcPendidikan.MatchedWithList = True Then cboWarganegara.SetFocus
        strSQL = "SELECT kdPendidikan, pendidikan FROM Pendidikan where StatusEnabled='1' and Pendidikan LIKE '%" & dcPendidikan.Text & "%'order By Pendidikan"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcPendidikan.Text = ""
            dcPendidikan.SetFocus
            Exit Sub
        End If
        dcPendidikan.BoundText = rs(0).value
        dcPendidikan.Text = rs(1).value
    End If
End Sub

Private Sub dcPendidikan_LostFocus()
    If dcPendidikan.Text = "" Then Exit Sub
    If dcPendidikan.MatchedWithList = False Then dcPendidikan.Text = "": dcPendidikan.SetFocus

End Sub

'Private Sub dcPropinsi_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Call subLoadDataWilayah("propinsi")
'        dcKota.SetFocus
'    End If
'End Sub
Private Sub dcPropinsi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        j = 1
        dcKota.Enabled = True
        dcKecamatan.Enabled = True
        dcKelurahan.Enabled = True
        Call subLoadDataWilayah("propinsi")
        If dcKota.Enabled = True Then
            dcKota.SetFocus
        End If
    End If
End Sub

Private Sub dcPropinsi_Click(Area As Integer)
    dcKota.Text = ""
    dcKecamatan.Text = ""
    dcKelurahan.Text = ""
    txtKodePos = ""
    CekPilihanWilayah "dcPropinsi", "Click"
End Sub

Private Sub CekPilihanWilayah(strItem As String, Optional strEvent As String)
    Dim X As Integer
    Dim Y

    X = 0
    Select Case strItem
        Case "dcPropinsi"
            Set dcKota.RowSource = Nothing
            Set dcKecamatan.RowSource = Nothing
            Set dcKelurahan.RowSource = Nothing
            dcKota.Text = ""
            dcKecamatan.Text = ""
            dcKelurahan.Text = ""
            txtKodePos = ""
            Select Case strEvent
                Case "Click"
                    subDcSource "Kota", " where kdPropinsi = '" & dcPropinsi.BoundText & "' order by NamaKotaKabupaten"
                Case "KeyPress"
                    If dcPropinsi.MatchedWithList = False Then
                        MsgBox "Pilih Propinsi"
                        X = 1
                        GoTo kosong
                        dcPropinsi.SetFocus
                    Else
                        subDcSource "Kota", " where kdPropinsi = '" & dcPropinsi.BoundText & "' order by NamaKotaKabupaten"
                        dcKota.SetFocus
                    End If
                Case "LostFocus"
                    If dcPropinsi.MatchedWithList = False Then
                        MsgBox "Pilih Propinsi"
                        X = 1
                        GoTo kosong
                        dcPropinsi.SetFocus
                    Else
                        subDcSource "Kota", " where kdPropinsi = '" & dcPropinsi.BoundText & "' order by NamaKotaKabupaten"
                        dcKota.SetFocus
                    End If
            End Select
        Case "dcKota"
            Set dcKecamatan.RowSource = Nothing
            Set dcKelurahan.RowSource = Nothing
            dcKecamatan.Text = ""
            dcKelurahan.Text = ""
            txtKodePos = ""
            If dcPropinsi.MatchedWithList = True Then
                Select Case strEvent
                    Case "Click"
                        If dcKota.Text = "" Then Exit Sub
                        subDcSource "Kecamatan", " where kdKotaKabupaten = '" & dcKota.BoundText & "' order by NamaKecamatan"
                    Case "KeyPress"
                        If dcKota.MatchedWithList = False Then
                           MsgBox "Pilih Kota"
                            X = 2
                            GoTo kosong
                            dcKota.SetFocus
                        Else
                            subDcSource "Kecamatan", " where kdKotaKabupaten = '" & dcKota.BoundText & "' order by NamaKecamatan"
                            dcKecamatan.SetFocus
                        End If
                    Case "LostFocus"
                        If dcKota.MatchedWithList = False Then
                            MsgBox "Pilih Kota"
                            X = 2
                            GoTo kosong
                            dcKota.SetFocus
                        Else
                            subDcSource "Kecamatan", " where kdKotaKabupaten = '" & dcKota.BoundText & "' order by NamaKecamatan"
                            dcKecamatan.SetFocus
                        End If
                End Select
            End If
        Case "dcKecamatan"
            Set dcKelurahan.RowSource = Nothing
            dcKelurahan.Text = ""
            txtKodePos = ""
            If dcKota.MatchedWithList = True Then
                Select Case strEvent
                    Case "Click"
                        If dcKecamatan.Text = "" Then Exit Sub
                        subDcSource "Kelurahan", " where kdkecamatan = '" & dcKecamatan.BoundText & "' order by NamaKelurahan"
                    Case "KeyPress"
                        If dcKecamatan.MatchedWithList = False Then
                            MsgBox "Pilih Kecamatan"
                            X = 3
                            GoTo kosong
                            dcKecamatan.SetFocus
                        Else
                            subDcSource "Kelurahan", " where kdkecamatan = '" & dcKecamatan.BoundText & "' order by NamaKelurahan"
                            dcKelurahan.SetFocus
                        End If
                    Case "LostFocus"
                        If dcKecamatan.MatchedWithList = False Then
                            MsgBox "Pilih Kecamatan"
                            X = 3
                            GoTo kosong
                            dcKecamatan.SetFocus
                        Else
                            subDcSource "Kelurahan", " where kdkecamatan = '" & dcKecamatan.BoundText & "' order by NamaKelurahan"
                            dcKelurahan.SetFocus
                        End If
                End Select
            End If
        Case "dcKelurahan"
            txtKodePos = ""
            If dcKecamatan.MatchedWithList = True Then
                Select Case strEvent
                    Case "KeyPress"
                        If dcKelurahan.MatchedWithList = False Then
                            MsgBox "Pilih Desa/Kelurahan"
                            X = 4
                            GoTo kosong
                            dcKelurahan.Text = ""
                            dcKelurahan.SetFocus
                        Else
                            txtKodePos.SetFocus
                        End If
                    Case "LostFocus"
                        If dcKelurahan.MatchedWithList = False Then
                            MsgBox "Pilih Desa/Kelurahan"
                            X = 4
                            GoTo kosong
                            dcKelurahan.SetFocus
                        End If
                End Select
            End If
    End Select

    Exit Sub

kosong:
    Y = MsgBox("Mulai lagi dari awal", vbYesNo, "Wilayah") ' vbYesNoCancel
    Select Case Y
        Case vbYes
            dcPropinsi.Text = ""
            dcKota.Text = ""
            dcKecamatan.Text = ""
            dcKelurahan.Text = ""
            dcPropinsi.SetFocus
        Case vbNo
            Exit Sub
'            Select Case X
'                Case 1
'                    dcPropinsi.SetFocus
'                Case 2
'                    dcKota.SetFocus
'                Case 3
'                    dcKecamatan.SetFocus
'                Case 4
'                    dcKelurahan.SetFocus
'            End Select
'        Case vbCancel
'            Exit Sub
    End Select
End Sub

Private Sub dcPropinsi_LostFocus()
    dcPropinsi = Trim(StrConv(dcPropinsi, vbProperCase))
End Sub

Private Sub dcSuku_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcSuku.BoundText
    strSQL = "SELECT Suku FROM Suku where StatusEnabled='1' Order By Suku"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcSuku.RowSource = rs
    dcSuku.ListField = rs.Fields(0).Name
    Set rs = Nothing
    dcSuku.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcSuku_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then
        If dcSuku.MatchedWithList = True Then dcPendidikan.SetFocus
        strSQL = "SELECT kdSuku, suku FROM Suku where StatusEnabled='1' and (Suku LIKE '%" & dcSuku.Text & "%')Order By Suku"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcSuku.Text = ""
            dcSuku.SetFocus
            Exit Sub
        End If
        dcSuku.BoundText = rs(0).value
        dcSuku.Text = rs(1).value
    End If
End Sub

Private Sub dcSuku_LostFocus()
    If dcSuku.Text = "" Then Exit Sub
    If dcSuku.MatchedWithList = False Then dcSuku.Text = "": dcSuku.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    cboGD.ListIndex = -1
    cboStsPernikahan.ListIndex = -1
    cboWarganegara.ListIndex = -1
    cboJnsKelaminKel.ListIndex = -1
    
    txtNoCM.MaxLength = "6"

    'Call subDcSource
    Call subDcDetSource
     subDcSource "Propinsi"
     subDcSource "Kota"
     subDcSource "Kecamatan"
     subDcSource "Kelurahan"
     
    If strPasien = "Baru" Then Call subLoadDataPasien(mstrNoCM)
    If strPasien = "Lama" Then Call subLoadDataPasien(mstrNoCM)
    If strPasien = "View" Then
        Call subLoadDataPasien(mstrNoCM)
    End If
End Sub

Private Sub meRTRWKel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTeleponKel.SetFocus
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtAlamatKel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then meRTRWKel.SetFocus
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub TxtAyah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then TxtIbu.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub TxtIbu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then TxtIstriSuami.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub TxtIstriSuami_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtNoKK.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtKepalaKeluarga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtAlamatKel.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtKodePos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      '  Call subLoadDataWilayah("kodepos")
        If cmdSimpan.Enabled = True Then cmdSimpan.SetFocus Else cmdTutup.SetFocus
    End If
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtNamaKelPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboJnsKelaminKel.SetFocus
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtNamaKeluarga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboGD.SetFocus
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtNoKK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtKepalaKeluarga.SetFocus
End Sub

Private Sub txtTeleponKel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcPropinsi.SetFocus
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToNumber(KeyAscii)
End Sub

'untuk cek validasi
Private Function funcCekValidasi(strChoose As String) As Boolean
    Select Case strChoose
        Case "DetailPasien"
            If txtNamaKeluarga.Text = "" Then
                MsgBox "Nama Keluarga harus diisi", vbExclamation, "Validasi"
                funcCekValidasi = False
                txtNamaKeluarga.SetFocus
                Exit Function
            End If
            If dcPekerjaan.Text = "" Then
                MsgBox "Pekerjaan Pasien harus diisi", vbExclamation, "Validasi"
                funcCekValidasi = False
                dcPekerjaan.SetFocus
                Exit Function
            End If
            If dcPendidikan.Text = "" Then
                MsgBox "Pendidikan Pasien harus diisi", vbExclamation, "Validasi"
                funcCekValidasi = False
                dcPendidikan.SetFocus
                Exit Function
            End If
            If cboWarganegara.Text = "" Then
                MsgBox "Warga Negara Pasien harus diisi", vbExclamation, "Validasi"
                funcCekValidasi = False
                cboWarganegara.SetFocus
                Exit Function
            End If
        Case "KeluargaPasien"
            If cboJnsKelaminKel.Text = "" Then
                MsgBox "Jenis Kelamin Keluarga Pasien harus diisi", vbExclamation, "Validasi"
                funcCekValidasi = False
                cboJnsKelaminKel.SetFocus
                Exit Function
            End If
            If dcHubungan.Text = "" Then
                MsgBox "Hubungan antara Pasien dengan Keluarga Pasien harus diisi", vbExclamation, "Validasi"
                funcCekValidasi = False
                dcHubungan.SetFocus
                Exit Function
            End If
    End Select
    funcCekValidasi = True
End Function

'Store procedure untuk mengisi detail pasien
Private Sub sp_DetailPasien(ByVal adoCommand As ADODB.Command)
    Dim strRhesus As String
    If chkRhesus.value = 1 Then
        strRhesus = "+"
    Else
        strRhesus = "-"
    End If
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        If txtNamaKeluarga.Text <> "" Then
            .Parameters.Append .CreateParameter("NamaKeluarga", adVarChar, adParamInput, 50, txtNamaKeluarga.Text)
        Else
            .Parameters.Append .CreateParameter("NamaKeluarga", adVarChar, adParamInput, 50, Null)
        End If
        .Parameters.Append .CreateParameter("WargaNegara", adChar, adParamInput, 1, Left(cboWarganegara.Text, 1))
        If cboGD.Text <> "" Then
            .Parameters.Append .CreateParameter("GolDarah", adVarChar, adParamInput, 2, cboGD.Text)
        Else
            .Parameters.Append .CreateParameter("GolDarah", adVarChar, adParamInput, 2, Null)
        End If
        .Parameters.Append .CreateParameter("Rhesus", adChar, adParamInput, 1, strRhesus)
        If cboStsPernikahan.Text <> "" Then
            .Parameters.Append .CreateParameter("StatusNikah", adVarChar, adParamInput, 10, cboStsPernikahan.Text)
        Else
            .Parameters.Append .CreateParameter("StatusNikah", adVarChar, adParamInput, 10, Null)
        End If
        .Parameters.Append .CreateParameter("Pekerjaan", adVarChar, adParamInput, 30, Trim(Left(dcPekerjaan.Text, 30)))
        If dcAgama.Text <> "" Then
            .Parameters.Append .CreateParameter("Agama", adVarChar, adParamInput, 20, Trim(Left(dcAgama.Text, 20)))
        Else
            .Parameters.Append .CreateParameter("Agama", adVarChar, adParamInput, 20, Null)
        End If
        If dcSuku.Text <> "" Then
            .Parameters.Append .CreateParameter("Suku", adVarChar, adParamInput, 20, Trim(Left(dcSuku.Text, 20)))
        Else
            .Parameters.Append .CreateParameter("Suku", adVarChar, adParamInput, 20, Null)
        End If
        .Parameters.Append .CreateParameter("Pendidikan", adVarChar, adParamInput, 25, IIf(dcPendidikan.Text = "", Null, dcPendidikan.Text))
        .Parameters.Append .CreateParameter("NamaAyah", adVarChar, adParamInput, 30, IIf(TxtAyah.Text = "", Null, TxtAyah.Text))
        .Parameters.Append .CreateParameter("NamaIbu", adVarChar, adParamInput, 30, IIf(TxtIbu.Text = "", Null, TxtIbu.Text))
        .Parameters.Append .CreateParameter("NamaIstriSuami", adVarChar, adParamInput, 30, IIf(TxtIstriSuami.Text = "", Null, TxtIstriSuami.Text))
        .Parameters.Append .CreateParameter("NoKK", adVarChar, adParamInput, 30, IIf(txtNoKK.Text = "", Null, txtNoKK.Text))
        .Parameters.Append .CreateParameter("NamaKepalaKeluarga", adVarChar, adParamInput, 30, IIf(txtKepalaKeluarga.Text = "", Null, txtKepalaKeluarga.Text))

        .ActiveConnection = dbConn
        .CommandText = "dbo.AU_DetailPasien "
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data Detail Pasien", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("AU_DetailPasien")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'Store procedure untuk mengisi detail pasien
Private Sub sp_KeluargaPasien(ByVal adoCommand As ADODB.Command)
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("NamaLengkap", adVarChar, adParamInput, 50, txtNamaKelPasien.Text)
        .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 1, Left(cboJnsKelaminKel.Text, 1))
        If txtAlamatKel.Text <> "" Then
            .Parameters.Append .CreateParameter("Alamat", adVarChar, adParamInput, 100, txtAlamatKel.Text)
        Else
            .Parameters.Append .CreateParameter("Alamat", adVarChar, adParamInput, 100, Null)
        End If
        If dcPekerjaanKel.Text <> "" Then
            .Parameters.Append .CreateParameter("Pekerjaan", adVarChar, adParamInput, 30, dcPekerjaanKel.Text)
        Else
            .Parameters.Append .CreateParameter("Pekerjaan", adVarChar, adParamInput, 30, Null)
        End If
        If txtTeleponKel.Text <> "" Then
            .Parameters.Append .CreateParameter("Telepon", adVarChar, adParamInput, 15, txtTeleponKel.Text)
        Else
            .Parameters.Append .CreateParameter("Telepon", adVarChar, adParamInput, 15, Null)
        End If
        If dcKelurahan.Text <> "" Then
            .Parameters.Append .CreateParameter("Kelurahan", adVarChar, adParamInput, 50, dcKelurahan.Text)
        Else
            .Parameters.Append .CreateParameter("Kelurahan", adVarChar, adParamInput, 50, Null)
        End If
        If meRTRWKel.Text <> "__/__" Then
            .Parameters.Append .CreateParameter("RTRW", adVarChar, adParamInput, 5, meRTRWKel.Text)
        Else
            .Parameters.Append .CreateParameter("RTRW", adVarChar, adParamInput, 5, Null)
        End If
        If dcKecamatan.Text <> "" Then
            .Parameters.Append .CreateParameter("Kecamatan", adVarChar, adParamInput, 50, dcKecamatan.Text)
        Else
            .Parameters.Append .CreateParameter("Kecamatan", adVarChar, adParamInput, 50, Null)
        End If
        If dcKota.Text <> "" Then
            .Parameters.Append .CreateParameter("Kota", adVarChar, adParamInput, 50, dcKota.Text)
        Else
            .Parameters.Append .CreateParameter("Kota", adVarChar, adParamInput, 50, Null)
        End If
        If dcPropinsi.Text <> "" Then
            .Parameters.Append .CreateParameter("Propinsi", adVarChar, adParamInput, 30, dcPropinsi.Text)
        Else
            .Parameters.Append .CreateParameter("Propinsi", adVarChar, adParamInput, 30, Null)
        End If
        If txtKodePos.Text <> "" Then
            .Parameters.Append .CreateParameter("KodePos", adChar, adParamInput, 5, txtKodePos.Text)
        Else
            .Parameters.Append .CreateParameter("KodePos", adChar, adParamInput, 5, Null)
        End If
        .Parameters.Append .CreateParameter("Hubungan", adVarChar, adParamInput, 20, dcHubungan.Text)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AU_KeluargaPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data Keluarga Pasien", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("AU_KeluargaPasien")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'untuk load data pasien yg sudah pernah didaftarkan
Private Sub subLoadDataPasien(strInput As String)
On Error Resume Next
    Dim strSQLLoadPasien As String
    Dim rsLoadPasien As New ADODB.recordset

    strSQLLoadPasien = "SELECT * FROM DetailPasien WHERE NoCM = '" & strInput & "'"
    Set rsLoadPasien = Nothing
    rsLoadPasien.Open strSQLLoadPasien, dbConn, adOpenForwardOnly, adLockReadOnly
    If rsLoadPasien.RecordCount > 0 Then 'Exit Sub
        txtNoCM.Text = mstrNoCM
        If Not IsNull(rsLoadPasien.Fields("NamaKeluarga").value) Then txtNamaKeluarga.Text = rsLoadPasien.Fields("NamaKeluarga").value
        If Not IsNull(rsLoadPasien.Fields("GolDarah").value) Then cboGD.Text = rsLoadPasien.Fields("GolDarah").value
        If rsLoadPasien.Fields("Rhesus").value = "+" Then
            chkRhesus.value = 1
        Else
            chkRhesus.value = 0
        End If
        If Not IsNull(rsLoadPasien.Fields("StatusNikah").value) Then cboStsPernikahan.Text = rsLoadPasien.Fields("StatusNikah").value
        dcPekerjaan.Text = rsLoadPasien.Fields("Pekerjaan").value
        If Not IsNull(rsLoadPasien.Fields("Agama").value) Then dcAgama.Text = rsLoadPasien.Fields("Agama").value
        If Not IsNull(rsLoadPasien.Fields("Suku").value) Then dcSuku.Text = rsLoadPasien.Fields("Suku").value
        TxtIbu.Text = IIf(IsNull(rsLoadPasien.Fields("NamaIbu").value), "", rsLoadPasien.Fields("NamaIbu").value)
        TxtAyah.Text = IIf(IsNull(rsLoadPasien.Fields("NamaAyah").value), "", rsLoadPasien.Fields("NamaAyah").value)
        TxtIstriSuami.Text = IIf(IsNull(rsLoadPasien.Fields("NamaSuamiIstri").value), "", rsLoadPasien.Fields("NamaSuamiIstri").value)
        dcPendidikan.Text = rsLoadPasien.Fields("Pendidikan").value
        If rsLoadPasien.Fields("Warganegara").value = "I" Then
            cboWarganegara.ListIndex = 0
        ElseIf rsLoadPasien.Fields("Warganegara").value = "A" Then
            cboWarganegara.ListIndex = 1
        End If
    End If

    strSQLLoadPasien = "SELECT * FROM KeluargaPasien WHERE NoCM = '" & strInput & "'"
    Set rsLoadPasien = Nothing
    rsLoadPasien.Open strSQLLoadPasien, dbConn, adOpenForwardOnly, adLockReadOnly
    If rsLoadPasien.RecordCount = 0 Then Set rsLoadPasien = Nothing: Exit Sub
    txtNamaKelPasien.Text = rsLoadPasien.Fields("NamaLengkap").value
    If rsLoadPasien.Fields("JenisKelamin").value = "L" Then
        cboJnsKelaminKel.ListIndex = 0
    ElseIf rsLoadPasien.Fields("JenisKelamin").value = "P" Then
        cboJnsKelaminKel.ListIndex = 1
    End If
    dcHubungan.Text = rsLoadPasien.Fields("Hubungan").value
    If Not IsNull(rsLoadPasien.Fields("Pekerjaan").value) Then dcPekerjaanKel.Text = rsLoadPasien.Fields("Pekerjaan").value
    If Not IsNull(rsLoadPasien.Fields("Alamat").value) Then txtAlamatKel.Text = rsLoadPasien.Fields("Alamat").value
    If Not IsNull(rsLoadPasien.Fields("RTRW").value) Then
        If Len(rsLoadPasien.Fields("RTRW").value) = 5 And InStr(1, rsLoadPasien.Fields("RTRW").value, "/") = 3 Then
            meRTRWKel.Text = rsLoadPasien.Fields("RTRW").value
        Else
            If InStr(1, rsLoadPasien.Fields("RTRW").value, "/") = 0 Then
                meRTRWKel.Text = Format(Left(rsLoadPasien.Fields("RTRW").value, Len(rsLoadPasien.Fields("RTRW").value) / 2), "00") & "/" & Format(Right(rsLoadPasien.Fields("RTRW").value, Len(rsLoadPasien.Fields("RTRW").value) / 2), "00")
            Else
                meRTRWKel.Text = Format(Left(rsLoadPasien.Fields("RTRW").value, InStr(1, rsLoadPasien.Fields("RTRW").value, "/") - 1), "00") & "/" & Format(Right(rsLoadPasien.Fields("RTRW").value, Len(rsLoadPasien.Fields("RTRW").value) - InStr(1, rsLoadPasien.Fields("RTRW").value, "/")), "00")
            End If
        End If
    End If
    If Not IsNull(rsLoadPasien.Fields("Telepon").value) Then txtTeleponKel.Text = rsLoadPasien.Fields("Telepon").value
    If Not IsNull(rsLoadPasien.Fields("Propinsi").value) Then dcPropinsi.Text = rsLoadPasien.Fields("Propinsi").value
    If Not IsNull(rsLoadPasien.Fields("Kota").value) Then dcKota.Text = rsLoadPasien.Fields("Kota").value
    If Not IsNull(rsLoadPasien.Fields("Kecamatan").value) Then dcKecamatan.Text = rsLoadPasien.Fields("Kecamatan").value
    If Not IsNull(rsLoadPasien.Fields("Kelurahan").value) Then dcKelurahan.Text = rsLoadPasien.Fields("Kelurahan").value
    If Not IsNull(rsLoadPasien.Fields("KodePos").value) Then txtKodePos.Text = rsLoadPasien.Fields("KodePos").value
    Set rsLoadPasien = Nothing

    Exit Sub

    Call msubPesanError
End Sub

Private Sub subDcDetSource()
    On Error GoTo errLoad

    strSQL = "SELECT KdAgama,Agama FROM Agama where StatusEnabled='1' Order By Agama"
    Call msubDcSource(dcAgama, rs, strSQL)

    strSQL = "SELECT Hubungan,NamaHubungan FROM HubunganKeluarga where StatusEnabled='1'"
    Call msubDcSource(dcHubungan, rs, strSQL)

    strSQL = "SELECT KdPekerjaan,Pekerjaan FROM Pekerjaan where StatusEnabled='1' Order By Pekerjaan"
    Call msubDcSource(dcPekerjaan, rs, strSQL)
    Call msubDcSource(dcPekerjaanKel, rs, strSQL)
    
    strSQL = "SELECT KdPendidikan,Pendidikan FROM Pendidikan where StatusEnabled='1' order By Pendidikan"
    Call msubDcSource(dcPendidikan, rs, strSQL)

    strSQL = "SELECT KdSuku,Suku FROM Suku where StatusEnabled='1' Order By Suku"
    Call msubDcSource(dcSuku, rs, strSQL)

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subDcSource(varstrPilihan As String, Optional varStrSQL As String)
    Select Case varstrPilihan

        Case "Propinsi"
            strSQL = "SELECT DISTINCT KdPropinsi, NamaPropinsi AS alias FROM V_Wilayah where StatusEnabled=1 order by NamaPropinsi"
            Set rsPropinsi = Nothing
            rsPropinsi.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dcPropinsi.RowSource = rsPropinsi
            dcPropinsi.BoundColumn = rsPropinsi(0).Name
            dcPropinsi.ListField = rsPropinsi(1).Name
        Case "Kota"
            strSQL = "SELECT DISTINCT KdKotaKabupaten, NamaKotaKabupaten AS alias FROM V_Wilayah " & varStrSQL & ""
            Set rsKota = Nothing
            rsKota.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dcKota.RowSource = rsKota
            dcKota.BoundColumn = rsKota(0).Name
            dcKota.ListField = rsKota(1).Name
        Case "Kecamatan"
            strSQL = "SELECT DISTINCT KdKecamatan, NamaKecamatan AS alias FROM V_Wilayah " & varStrSQL & ""
            Set rsKecamatan = Nothing
            rsKecamatan.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dcKecamatan.RowSource = rsKecamatan
            dcKecamatan.BoundColumn = rsKecamatan(0).Name
            dcKecamatan.ListField = rsKecamatan(1).Name
        Case "Kelurahan"
            strSQL = "SELECT DISTINCT KdKelurahan, NamaKelurahan AS alias FROM V_Wilayah " & varStrSQL & ""
            Set rsKelurahan = Nothing
            rsKelurahan.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dcKelurahan.RowSource = rsKelurahan
            dcKelurahan.BoundColumn = rsKelurahan(0).Name
            dcKelurahan.ListField = rsKelurahan(1).Name
    End Select

    Exit Sub
End Sub

Private Sub subLoadDataWilayah(strPencarian As String)
    'On Error GoTo errLoad
    On Error Resume Next
    Dim strTempSql As String

    Select Case strPencarian
        Case "propinsi"
            If Len(Trim(dcPropinsi.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaPropinsi LIKE '%" & dcPropinsi.Text & "%')and statusenabled=1"

        Case "kota"
            If Len(Trim(dcKota.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaPropinsi LIKE '%" & dcPropinsi.Text & "%') and (NamaKotaKabupaten LIKE '%" & dcKota.Text & "%')"

        Case "kecamatan"
            If Len(Trim(dcKecamatan.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaPropinsi LIKE '%" & dcPropinsi.Text & "%') and (NamaKotaKabupaten LIKE '%" & dcKota.Text & "%') and (NamaKecamatan LIKE '%" & dcKecamatan.Text & "%')"
        Case "desa"
            If Len(Trim(dcKelurahan.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaPropinsi LIKE '%" & dcPropinsi.Text & "%') and (NamaKotaKabupaten LIKE '%" & dcKota.Text & "%') and (NamaKecamatan LIKE '%" & dcKecamatan.Text & "%') and (NamaKelurahan LIKE '%" & dcKelurahan.Text & "%')"

        Case "kodepos"
            If Len(Trim(txtKodePos.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaPropinsi LIKE '%" & dcPropinsi.Text & "%') and (NamaKotaKabupaten LIKE '%" & dcKota.Text & "%') and (NamaKecamatan LIKE '%" & dcKecamatan.Text & "%') and (NamaKelurahan LIKE '%" & dcKelurahan.Text & "%') and (KodePos LIKE '%" & txtKodePos.Text & "%')"

    End Select

    strSQL = "SELECT DISTINCT ISNULL(NamaPropinsi, '') AS NamaPropinsi, ISNULL(NamaKotaKabupaten, '') AS NamaKotaKabupaten, ISNULL(NamaKecamatan, '')  AS NamaKecamatan, ISNULL(NamaKelurahan, '') AS NamaKelurahan, ISNULL(KodePos, '') AS KodePos" & _
    " FROM V_Wilayah" & _
    " " & strTempSql

    Call msubRecFO(rs, strSQL)
    If rs.EOF Then
        MsgBox "Data Wilayah Tidak Sesuai, Harap Cek Data Wilayah", vbInformation, "Validasi"

        dcPropinsi.BoundText = ""
        dcKota.BoundText = ""
        dcKecamatan.BoundText = ""
        dcKelurahan.BoundText = ""
        txtKodePos.Text = ""

    ElseIf j = 1 Then
        If rs(1).value = "" Then
            MsgBox "Data Kota/Kabupaten Belum Ada", vbInformation, "Validasi"
            dcKota.Enabled = False
            dcKecamatan.Enabled = False
            dcKelurahan.Enabled = False
        Else

        End If

    ElseIf j = 2 Then
        If rs(2).value = "" Then
            MsgBox "Data Kecamatan Belum Ada", vbInformation, "Validasi"
            dcKecamatan.Enabled = False
            dcKelurahan.Enabled = False
        Else

        End If

    ElseIf j = 3 Then
        If rs(3).value = "" Then
            MsgBox "Data Kelurahan Belum Ada", vbInformation, "Validasi"
            dcKelurahan.Enabled = False
        Else

        End If

    Else
        dcPropinsi.Text = rs("NamaPropinsi")
        dcKota.Text = rs("NamaKotaKabupaten")
        dcKecamatan.Text = rs("NamaKecamatan")
        dcKelurahan.Text = rs("NamaKelurahan")
        txtKodePos.Text = rs("KodePos")
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKota_Click(Area As Integer)
    dcKecamatan.Text = ""
    dcKelurahan.Text = ""
    txtKodePos = ""
    CekPilihanWilayah "dcKota", "Click"
End Sub

Private Sub dcKecamatan_Click(Area As Integer)
    dcKelurahan.Text = ""
    txtKodePos = ""
    CekPilihanWilayah "dcKecamatan", "Click"
End Sub

Private Sub dcKelurahan_Click(Area As Integer)
    txtKodePos = ""
    CekPilihanWilayah "dcKelurahan", "Click"
End Sub




