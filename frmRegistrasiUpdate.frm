VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegistrasiUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Ubah Registrasi Pasien"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegistrasiUpdate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   11310
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   27
      Top             =   4920
      Width           =   11295
      Begin VB.TextBox txtNoBKM 
         Height          =   495
         Left            =   240
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
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
         Left            =   7800
         TabIndex        =   18
         Top             =   240
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
         Left            =   9480
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "Cetak Label [F1]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   46
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Data Registrasi Baru"
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
      TabIndex        =   20
      Top             =   3120
      Width           =   11295
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   360
         Left            =   4200
         TabIndex        =   13
         Top             =   510
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
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   360
         Left            =   4680
         TabIndex        =   16
         Top             =   1245
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
      Begin MSDataListLib.DataCombo dcKelas 
         Height          =   360
         Left            =   2160
         TabIndex        =   15
         Top             =   1245
         Width           =   2415
         _ExtentX        =   4260
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
      Begin MSComCtl2.DTPicker dtpTglPendaftaran 
         Height          =   360
         Left            =   2160
         TabIndex        =   12
         Top             =   525
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
         Format          =   463667203
         UpDown          =   -1  'True
         CurrentDate     =   38061
      End
      Begin MSDataListLib.DataCombo dcJenisKelas 
         Height          =   360
         Left            =   7920
         TabIndex        =   14
         Top             =   480
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
      Begin MSDataListLib.DataCombo dcSubInstalasi 
         Height          =   360
         Left            =   7920
         TabIndex        =   17
         Top             =   1245
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
      Begin VB.Label Label4 
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
         Left            =   7920
         TabIndex        =   43
         Top             =   960
         Width           =   1845
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelas Pelayanan"
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
         Left            =   7920
         TabIndex        =   34
         Top             =   240
         Width           =   1860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Masuk"
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
         TabIndex        =   28
         Top             =   240
         Width           =   930
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
         Left            =   2160
         TabIndex        =   23
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Ruang Pemeriksaan"
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
         TabIndex        =   22
         Top             =   960
         Width           =   1695
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
         Left            =   4200
         TabIndex        =   21
         Top             =   240
         Width           =   1860
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Data Registrasi Lama"
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
      TabIndex        =   37
      Top             =   2040
      Width           =   11295
      Begin VB.TextBox txtKelompokPasienLama 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   550
         Width           =   2175
      End
      Begin VB.TextBox txtTglMasukLama 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   550
         Width           =   1935
      End
      Begin VB.TextBox txtRuanganLama 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   550
         Width           =   2415
      End
      Begin VB.TextBox txtKelasLama 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   550
         Width           =   1935
      End
      Begin VB.TextBox txtJenisKelasLama 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   550
         Width           =   1935
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Cara Bayar"
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
         TabIndex        =   42
         Top             =   285
         Width           =   945
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Masuk"
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
         TabIndex        =   41
         Top             =   280
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Ruang Pelayanan"
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
         Left            =   6360
         TabIndex        =   40
         Top             =   280
         Width           =   1470
      End
      Begin VB.Label Label10 
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
         Left            =   4320
         TabIndex        =   39
         Top             =   280
         Width           =   1380
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelas Pelayanan"
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
         TabIndex        =   38
         Top             =   280
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
      TabIndex        =   24
      Top             =   960
      Width           =   11295
      Begin VB.TextBox txtNoPendaftaran 
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
         Left            =   150
         MaxLength       =   10
         TabIndex        =   0
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
         Height          =   735
         Left            =   8520
         TabIndex        =   29
         Top             =   240
         Width           =   2535
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   360
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   6
            Top             =   330
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   360
            Left            =   960
            MaxLength       =   6
            TabIndex        =   5
            Top             =   330
            Width           =   375
         End
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   360
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6960
         MaxLength       =   9
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
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
         Left            =   150
         TabIndex        =   36
         Top             =   360
         Width           =   1380
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
         Left            =   6960
         TabIndex        =   33
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
         Left            =   1800
         TabIndex        =   26
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
         Left            =   3240
         TabIndex        =   25
         Top             =   360
         Width           =   1110
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   44
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
      Left            =   9480
      Picture         =   "frmRegistrasiUpdate.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRegistrasiUpdate.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRegistrasiUpdate.frx":30B0
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
      Left            =   0
      TabIndex        =   35
      Top             =   960
      Width           =   1605
   End
End
Attribute VB_Name = "frmRegistrasiUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilter As String
Dim intRowNow As Integer
Dim strSubInstalasi As String
Dim strNoAntrian As String

Dim subStrKdRuanganLama As String
Dim subStrKdKelasLama As String
Dim subStrKdInstalasiLama As String
Dim subStrStatusPasienLama As String

'Store procedure untuk mengisi struk billing pasien
Public Function sp_AddStrukBuktiKasMasuk() As Boolean
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
        .Parameters.Append .CreateParameter("JmlBayar", adCurrency, adParamInput, , mcurAll_HrsDibyr)
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
Public Function sp_AddStruk(ByVal adoCommand As ADODB.Command, strStsByr As String) As Boolean
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

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    If funcCekValidasi = False Then Exit Sub
    Call sp_UpdateRegistrasiMRS(dbcmd)
    Call subEnableButtonReg(True)

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan data ubah registrasi pasien", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub dcInstalasi_Change()
    dcJenisKelas.Text = ""
End Sub

Private Sub dcInstalasi_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String
    tempKode = dcInstalasi.BoundText
    strSQL = "SELECT DISTINCT KdInstalasi,NamaInstalasi FROM V_KelasPelayanan WHERE KdInstalasi <> '03' and StatusEnabled='1'"
    Call msubDcSource(dcInstalasi, rs, strSQL)
    dcInstalasi.BoundText = tempKode
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcInstalasi.MatchedWithList = True Then dcJenisKelas.SetFocus
        strSQL = "SELECT DISTINCT KdInstalasi, NamaInstalasi FROM V_KelasPelayanan WHERE KdInstalasi <> '03' and StatusEnabled='1' and (NamaInstalasi LIKE '%" & dcInstalasi.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcInstalasi.Text = ""
            Exit Sub
        End If
        dcInstalasi.BoundText = rs(0).value
        dcInstalasi.Text = rs(1).value
    End If
End Sub

Private Sub dcJenisKelas_Change()
    dcKelas.Text = ""
End Sub

Private Sub dcJenisKelas_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcJenisKelas.BoundText

    strSQL = "SELECT distinct KdDetailJenisJasaPelayanan,DetailJenisJasaPelayanan FROM V_KelasPelayanan where KdInstalasi='" & dcInstalasi.BoundText & "' and Expr1='1'"
    Call msubDcSource(dcJenisKelas, rs, strSQL)
    dcJenisKelas.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcJenisKelas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcJenisKelas.MatchedWithList = True Then dcKelas.SetFocus
        strSQL = "SELECT distinct KdDetailJenisJasaPelayanan,DetailJenisJasaPelayanan FROM V_KelasPelayanan where KdInstalasi='" & dcInstalasi.BoundText & "' and Expr1='1' and (DetailJenisJasaPelayanan LIKE '%" & dcJenisKelas.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcJenisKelas.Text = ""
            Exit Sub
        End If
        dcJenisKelas.BoundText = rs(0).value
        dcJenisKelas.Text = rs(1).value
    End If
End Sub

Private Sub dcKelas_Change()
    dcRuangan.Text = ""
End Sub

Private Sub dcKelas_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcKelas.BoundText

    strSQL = "SELECT distinct KdKelas,Kelas FROM V_KelasPelayanan WHERE KdDetailJenisJasaPelayanan='" & dcJenisKelas.BoundText & "' AND KdInstalasi='" & dcInstalasi.BoundText & "' and Expr2='1'"
    Call msubDcSource(dcKelas, rs, strSQL)

    dcKelas.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKelas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKelas.MatchedWithList = True Then dcRuangan.SetFocus
        strSQL = "SELECT distinct KdKelas,Kelas FROM V_KelasPelayanan WHERE KdDetailJenisJasaPelayanan='" & dcJenisKelas.BoundText & "' AND KdInstalasi='" & dcInstalasi.BoundText & "' and Expr2='1'and (Kelas LIKE '%" & dcKelas.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKelas.BoundText = rs(0).value
        dcKelas.Text = rs(1).value
    End If
End Sub

Private Sub dcRuangan_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcRuangan.BoundText

    If dcInstalasi.BoundText = "03" Then
        strSQL = "SELECT distinct KdRuangan,Ruangan FROM V_KamarRegRawatInap " _
        & "WHERE Kelas='" & dcKelas.Text & "' and StatusEnabled='1'"
    Else
        strSQL = "SELECT distinct KdRuangan,NamaRuangan FROM V_KelasPelayanan " _
        & "WHERE NamaInstalasi='" & dcInstalasi.Text & "' AND KdKelas='" & dcKelas.BoundText & "' and StatusEnabled='1'"
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
        If dcInstalasi.BoundText = "03" Then
            strSQL = "SELECT distinct KdRuangan,Ruangan FROM V_KamarRegRawatInap " _
            & "WHERE Kelas='" & dcKelas.Text & "' AND Ruangan LIKE '%" & dcRuangan.Text & "%' and StatusEnabled='1'"
        Else
            strSQL = "SELECT distinct KdRuangan,NamaRuangan FROM V_KelasPelayanan " _
            & "WHERE NamaInstalasi='" & dcInstalasi.Text & "' AND KdKelas='" & dcKelas.BoundText & "' AND NamaRuangan LIKE '%" & dcRuangan.Text & "%' and StatusEnabled='1'"
        End If
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcRuangan.BoundText = rs(0).value
        dcRuangan.Text = rs(1).value

        strSQL = "SELECT KdSubInstalasi FROM V_RegistrasiAll WHERE KdRuangan='" & dcRuangan.BoundText & "' and StatusEnabled='1'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then strSubInstalasi = rs(0).value Else strSubInstalasi = ""

        strSQL = "SELECT KdSubInstalasi, NamaSubInstalasi FROM  V_SubInstalasiRuangan WHERE (KdRuangan = '" & dcRuangan.BoundText & "') and StatusEnabled='1'"
        Call msubDcSource(dcSubInstalasi, rs, strSQL)
        If rs.EOF = False Then dcSubInstalasi.BoundText = rs(0).value
        dcSubInstalasi.SetFocus
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcSubInstalasi_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcSubInstalasi.BoundText
    strSQL = "SELECT KdSubInstalasi, NamaSubInstalasi FROM V_RegistrasiALL WHERE (KdRuangan = '" & dcRuangan.BoundText & "') and StatusEnabled='1'"
    Call msubDcSource(dcSubInstalasi, rs, strSQL)
    dcSubInstalasi.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcSubInstalasi_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    If KeyAscii = 13 Then
        strSQL = "SELECT KdSubInstalasi, NamaSubInstalasi FROM V_RegistrasiALL WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND (NamaSubInstalasi LIKE '%" & dcSubInstalasi.Text & "%') and StatusEnabled='1'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcSubInstalasi.BoundText = rs(0).value
        cmdSimpan.SetFocus
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dtpTglPendaftaran_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcInstalasi.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo hell:
    Select Case KeyCode
        Case vbKeyF1
            mstrNoPen = frmRegistrasiUpdate.txtnopendaftaran.Text
            mstrKdInstalasi = frmRegistrasiUpdate.dcInstalasi.BoundText
            mstrNoCM = txtNoCM.Text
            If cmdSimpan.Enabled = True Then Exit Sub
            If dcInstalasi.BoundText = "02" Or dcInstalasi.BoundText = "06" Or dcInstalasi.BoundText = "11" Then
            End If
            frm_cetak_label_viewer.Show
'            frm_cetak_label_viewer.Cetaklangsung
    End Select
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
    strRegistrasi = "RJ"
    If mblnCariPasien = True Then frmCariPasien.Enabled = False
    typAsuransi.blnSuksesAsuransi = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnCariPasien = True Then frmCariPasien.Enabled = True
End Sub

Public Sub txtNoPendaftaran_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call subClearData
        Call subEnableButtonReg(False)
        strSQL = "Select * from V_UbahRegistrasiMRS WHERE NoPendaftaran='" & txtnopendaftaran.Text & "'"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            mstrNoCM = ""
            mstrNoPen = ""
            Call subClearData
            Call subEnableButtonReg(False)
            cmdSimpan.Enabled = False
            Exit Sub
        End If

        txtNoCM.Text = rs("NoCM")

        mstrNoPen = txtnopendaftaran.Text
        mstrNoCM = txtNoCM.Text
        subStrKdRuanganLama = rs("KdRuangan")
        subStrKdKelasLama = rs("KdKelas")
        subStrKdInstalasiLama = IIf(IsNull(rs("KdSubInstalasi")), "", rs("KdSubInstalasi"))
        subStrStatusPasienLama = rs("StatusPasien")

        txtNamaPasien.Text = rs.Fields("Nama Pasien").value
        If rs.Fields("JK").value = "P" Then
            txtJK.Text = "Perempuan"
        ElseIf rs.Fields("JK").value = "L" Then
            txtJK.Text = "Laki-laki"
        End If
        txtThn.Text = rs.Fields("UmurTahun").value
        txtBln.Text = rs.Fields("UmurBulan").value
        txtHr.Text = rs.Fields("UmurHari").value

        txtTglMasukLama.Text = rs("TglMasuk")
        txtJenisKelasLama.Text = rs("DetailJenisJasaPelayanan")
        txtKelasLama.Text = rs("Kelas")
        txtRuanganLama.Text = rs("Ruangan")
        txtKelompokPasienLama.Text = rs("Jenis Pasien")
        Set rs = Nothing
        dtpTglPendaftaran.SetFocus
    End If
End Sub

'untuk enable/disable button reg
Private Sub subEnableButtonReg(blnStatus As Boolean)
    cmdSimpan.Enabled = Not blnStatus
    dtpTglPendaftaran.Enabled = Not blnStatus
    dcInstalasi.Enabled = Not blnStatus
    dcRuangan.Enabled = Not blnStatus
    dcKelas.Enabled = Not blnStatus
    dcJenisKelas.Enabled = Not blnStatus
    dcSubInstalasi.Enabled = Not blnStatus
End Sub

'Store procedure untuk mengisi registrasi pasien
Private Sub sp_UpdateRegistrasiMRS(ByVal adoCommand As ADODB.Command)
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("KdRuanganLama", adChar, adParamInput, 3, subStrKdRuanganLama)
        .Parameters.Append .CreateParameter("KdSubInstalasiLama", adChar, adParamInput, 3, IIf(subStrKdInstalasiLama = "", Null, subStrKdInstalasiLama))
        .Parameters.Append .CreateParameter("KdKelasLama", adChar, adParamInput, 2, subStrKdKelasLama)
        .Parameters.Append .CreateParameter("TglMasukLama", adDate, adParamInput, , Format(txtTglMasukLama.Text, "yyyy/MM/dd HH:mm:ss"))

        .Parameters.Append .CreateParameter("KdRuanganBaru", adChar, adParamInput, 3, dcRuangan.BoundText)
        .Parameters.Append .CreateParameter("KdSubInstalasiBaru", adChar, adParamInput, 3, dcSubInstalasi.BoundText)
        .Parameters.Append .CreateParameter("KdDetailJenisJasaPelayananBaru", adChar, adParamInput, 2, dcJenisKelas.BoundText)
        .Parameters.Append .CreateParameter("KdKelasBaru", adChar, adParamInput, 2, dcKelas.BoundText)
        .Parameters.Append .CreateParameter("TglMasukBaru", adDate, adParamInput, , Format(dtpTglPendaftaran.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("StatusPasienLama", adChar, adParamInput, 4, subStrStatusPasienLama)

        .ActiveConnection = dbConn
        .CommandText = "Update_RegistrasiMRS"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam update registrasi pasien MRS", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Update_RegistrasiMRS")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'untuk cek validasi
Private Function funcCekValidasi() As Boolean
    If txtNamaPasien.Text = "" Then
        MsgBox "No. CM Harus Diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        txtNoCM.SetFocus
        Exit Function
    End If
    If dcJenisKelas.Text = "" Then
        MsgBox "Pilihan Jenis Kelas Pelayanan Pasien harus Diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        dcJenisKelas.SetFocus
        Exit Function
    End If
    If dcKelas.Text = "" Then
        MsgBox "Pilihan Kelas Pelayanan Pasien harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        dcKelas.SetFocus
        Exit Function
    End If
    If dcInstalasi.Text = "" Then
        MsgBox "Pilihan Instalasi harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        dcInstalasi.SetFocus
        Exit Function
    End If
    If dcRuangan.Text = "" Then
        MsgBox "Pilihan Ruangan harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        dcRuangan.SetFocus
        Exit Function
    End If
    If dcSubInstalasi.Text = "" Then
        MsgBox "Pilihan Sub Instalasi () Kasus Penyakit harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        dcSubInstalasi.SetFocus
        Exit Function
    End If

    If dtpTglPendaftaran.value > Now Then
        MsgBox "Tanggal pendaftaran tidak boleh lebih dari sekarang", vbCritical, "Validasi"
        funcCekValidasi = False
        dtpTglPendaftaran.SetFocus
        Exit Function
    End If

    funcCekValidasi = True
End Function

'untuk membersihkan data pasien registrasi
Private Sub subClearData()
    txtNoCM.Text = ""
    txtNamaPasien.Text = ""
    txtJK.Text = ""
    txtThn.Text = ""
    txtBln.Text = ""
    txtHr.Text = ""
    dtpTglPendaftaran.MaxDate = #9/9/2999#
    dtpTglPendaftaran.value = Now
    dcInstalasi.Text = ""
    dcRuangan.Text = ""
    dcJenisKelas.Text = ""
    dcKelas.Text = ""

    txtTglMasukLama.Text = ""
    txtJenisKelasLama.Text = ""
    txtKelasLama.Text = ""
    txtRuanganLama.Text = ""
    txtKelompokPasienLama.Text = ""

    subStrKdRuanganLama = ""
    subStrKdInstalasiLama = ""
    subStrKdKelasLama = ""
    subStrStatusPasienLama = ""
End Sub

