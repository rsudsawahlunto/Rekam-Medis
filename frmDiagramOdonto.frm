VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDiagramOdonto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Odontogram"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14700
   Icon            =   "frmDiagramOdonto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   14700
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   0
      TabIndex        =   25
      Top             =   8400
      Width           =   14655
      Begin VB.CommandButton cmdCetakOdonto 
         Caption         =   "&Cetak"
         Height          =   465
         Left            =   11280
         TabIndex        =   389
         Top             =   180
         Width           =   1575
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   465
         Left            =   9600
         TabIndex        =   27
         Top             =   180
         Width           =   1575
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         Height          =   465
         Left            =   12960
         TabIndex        =   26
         Top             =   180
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   1800
         TabIndex        =   334
         Top             =   240
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
         Format          =   388235267
         UpDown          =   -1  'True
         CurrentDate     =   37823
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Periksa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   335
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   0
      TabIndex        =   24
      Top             =   6240
      Width           =   14655
      Begin VB.OptionButton optAksi 
         Caption         =   "Calculus"
         Height          =   375
         Index           =   14
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   445
         Top             =   1680
         Width           =   1815
      End
      Begin VB.PictureBox picNonVital 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DrawWidth       =   3
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   333
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picSisaAkar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DrawWidth       =   2
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8880
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   332
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picGigiTiruanLepas 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DrawWidth       =   3
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8880
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   331
         Top             =   1800
         Width           =   255
      End
      Begin VB.PictureBox picJembatan 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DrawWidth       =   3
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8880
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   330
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picGigiHilang 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DrawWidth       =   3
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8880
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   329
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picMNonLogam 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6000
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   153
         Top             =   1800
         Width           =   255
      End
      Begin VB.PictureBox picMLogam 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6000
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   152
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picTNonLogam 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6000
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   151
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picTLogam 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6000
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   150
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picKaries 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   149
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton optAksi 
         Caption         =   "Gigi Tiruan Lepas"
         Height          =   375
         Index           =   13
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CheckBox chkShowHideGigiHilang 
         Caption         =   "Sembunyikan Gigi Hilang"
         Height          =   375
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   146
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton optAksi 
         Caption         =   "Jembatan"
         Height          =   375
         Index           =   12
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   145
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton optAksi 
         Caption         =   "Gigi Hilang"
         Height          =   375
         Index           =   11
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optAksi 
         Caption         =   "Sisa Akar"
         Height          =   375
         Index           =   10
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optAksi 
         Caption         =   "Mahkota Non Logam"
         Height          =   375
         Index           =   9
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   1680
         Width           =   1815
      End
      Begin VB.OptionButton optAksi 
         Caption         =   "Mahkota Logam"
         Height          =   375
         Index           =   8
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton optAksi 
         Caption         =   "Tambalan Non Logam"
         Height          =   375
         Index           =   7
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optAksi 
         Caption         =   "Tambalan Logam"
         Height          =   375
         Index           =   6
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optAksi 
         Caption         =   "Non Vital"
         Height          =   375
         Index           =   5
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optAksi 
         Caption         =   "Karies"
         Height          =   375
         Index           =   4
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   1320
         Width           =   1815
      End
      Begin VB.OptionButton optAksi 
         Caption         =   "Anomali Bentuk"
         Height          =   375
         Index           =   3
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton optAksi 
         Caption         =   "Erupsi Sebagian"
         Height          =   375
         Index           =   2
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton optAksi 
         Caption         =   "Belum Erupsi"
         Height          =   375
         Index           =   1
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optAksi 
         Caption         =   "&Normal"
         Height          =   975
         Index           =   0
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblCalculus 
         AutoSize        =   -1  'True
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   480
         TabIndex        =   446
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label lblAnomaliBentuk 
         AutoSize        =   -1  'True
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   600
         TabIndex        =   156
         Top             =   1080
         Width           =   165
      End
      Begin VB.Label lblErupsiSebagian 
         AutoSize        =   -1  'True
         Caption         =   "PE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   480
         TabIndex        =   155
         Top             =   720
         Width           =   315
      End
      Begin VB.Label lblBelumErupsi 
         AutoSize        =   -1  'True
         Caption         =   "UE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   480
         TabIndex        =   154
         Top             =   360
         Width           =   330
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
      TabIndex        =   0
      Top             =   960
      Width           =   14655
      Begin VB.TextBox txtTglDaftar 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   12960
         TabIndex        =   14
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtJenisPasien 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   11520
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   6120
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2640
         TabIndex        =   11
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   9
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
         Left            =   7440
         TabIndex        =   2
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
            TabIndex        =   5
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
            TabIndex        =   4
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
            TabIndex        =   3
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Left            =   550
            TabIndex        =   8
            Top             =   277
            Width           =   285
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Left            =   1350
            TabIndex        =   7
            Top             =   277
            Width           =   240
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Left            =   2130
            TabIndex        =   6
            Top             =   270
            Width           =   165
         End
      End
      Begin VB.TextBox txtKls 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   9960
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pendaftaran"
         Height          =   210
         Left            =   12960
         TabIndex        =   21
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pasien"
         Height          =   210
         Left            =   11520
         TabIndex        =   20
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   6120
         TabIndex        =   19
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   2640
         TabIndex        =   18
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1560
         TabIndex        =   17
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Pelayanan"
         Height          =   210
         Left            =   9960
         TabIndex        =   15
         Top             =   360
         Width           =   1275
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   133
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
   Begin VB.Frame Frame1 
      Caption         =   "Odontogram"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   0
      TabIndex        =   22
      Top             =   2040
      Width           =   14655
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         MaxLength       =   200
         TabIndex        =   391
         Top             =   3840
         Width           =   13695
      End
      Begin VB.PictureBox picDiagramOdondo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   14415
         TabIndex        =   23
         Top             =   240
         Width           =   14415
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   1335
            Left            =   120
            ScaleHeight     =   1275
            ScaleWidth      =   1515
            TabIndex        =   148
            Top             =   1080
            Visible         =   0   'False
            Width           =   1575
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "JANGAN DIHAPUS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   615
               Left            =   120
               TabIndex        =   390
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   80
               Left            =   0
               TabIndex        =   323
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   79
               Left            =   0
               TabIndex        =   322
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   78
               Left            =   0
               TabIndex        =   321
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   77
               Left            =   0
               TabIndex        =   320
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   76
               Left            =   0
               TabIndex        =   319
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   70
               Left            =   0
               TabIndex        =   313
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   69
               Left            =   0
               TabIndex        =   312
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   68
               Left            =   0
               TabIndex        =   311
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   67
               Left            =   0
               TabIndex        =   310
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   66
               Left            =   0
               TabIndex        =   309
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   60
               Left            =   0
               TabIndex        =   303
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   59
               Left            =   0
               TabIndex        =   302
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   58
               Left            =   0
               TabIndex        =   301
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   57
               Left            =   0
               TabIndex        =   300
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   56
               Left            =   0
               TabIndex        =   299
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   50
               Left            =   0
               TabIndex        =   293
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   49
               Left            =   0
               TabIndex        =   292
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   40
               Left            =   0
               TabIndex        =   283
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   39
               Left            =   0
               TabIndex        =   282
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   30
               Left            =   0
               TabIndex        =   273
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   29
               Left            =   0
               TabIndex        =   272
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   20
               Left            =   0
               TabIndex        =   263
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   19
               Left            =   0
               TabIndex        =   262
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   10
               Left            =   0
               TabIndex        =   253
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   9
               Left            =   0
               TabIndex        =   252
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   8
               Left            =   0
               TabIndex        =   251
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   7
               Left            =   0
               TabIndex        =   250
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   6
               Left            =   0
               TabIndex        =   249
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   5
               Left            =   0
               TabIndex        =   248
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   4
               Left            =   0
               TabIndex        =   247
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   3
               Left            =   0
               TabIndex        =   246
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   2
               Left            =   0
               TabIndex        =   245
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   244
               Top             =   0
               Width           =   45
            End
            Begin VB.Label lblGigiAnomali 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Height          =   195
               Index           =   0
               Left            =   720
               TabIndex        =   243
               Top             =   960
               Width           =   45
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   80
               Left            =   0
               TabIndex        =   237
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   79
               Left            =   0
               TabIndex        =   236
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   78
               Left            =   0
               TabIndex        =   235
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   77
               Left            =   0
               TabIndex        =   234
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   76
               Left            =   0
               TabIndex        =   233
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   70
               Left            =   0
               TabIndex        =   227
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   69
               Left            =   0
               TabIndex        =   226
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   68
               Left            =   0
               TabIndex        =   225
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   67
               Left            =   0
               TabIndex        =   224
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   66
               Left            =   0
               TabIndex        =   223
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   60
               Left            =   0
               TabIndex        =   217
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   59
               Left            =   0
               TabIndex        =   216
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   58
               Left            =   0
               TabIndex        =   215
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   57
               Left            =   0
               TabIndex        =   214
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   56
               Left            =   0
               TabIndex        =   213
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   50
               Left            =   0
               TabIndex        =   207
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   49
               Left            =   0
               TabIndex        =   206
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   40
               Left            =   0
               TabIndex        =   197
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   39
               Left            =   0
               TabIndex        =   196
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   30
               Left            =   0
               TabIndex        =   187
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   29
               Left            =   0
               TabIndex        =   186
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   20
               Left            =   0
               TabIndex        =   177
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   19
               Left            =   0
               TabIndex        =   176
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   10
               Left            =   0
               TabIndex        =   167
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   9
               Left            =   0
               TabIndex        =   166
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   8
               Left            =   0
               TabIndex        =   165
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   7
               Left            =   0
               TabIndex        =   164
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   6
               Left            =   0
               TabIndex        =   163
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   5
               Left            =   0
               TabIndex        =   162
               Top             =   0
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   161
               Top             =   960
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   3
               Left            =   600
               TabIndex        =   160
               Top             =   960
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   2
               Left            =   480
               TabIndex        =   159
               Top             =   960
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   1
               Left            =   360
               TabIndex        =   158
               Top             =   960
               Width           =   75
            End
            Begin VB.Label lblGigi 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   157
               Top             =   960
               Width           =   75
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   85
            Left            =   2760
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   131
            Top             =   1920
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   85
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   132
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   85
                  Left            =   0
                  TabIndex        =   328
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   85
                  Left            =   0
                  TabIndex        =   242
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line172 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line171 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   84
            Left            =   3360
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   129
            Top             =   1920
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   84
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   130
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   84
                  Left            =   0
                  TabIndex        =   327
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   84
                  Left            =   0
                  TabIndex        =   241
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line170 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line169 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   83
            Left            =   3960
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   127
            Top             =   1920
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   83
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   128
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   83
                  Left            =   0
                  TabIndex        =   326
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   83
                  Left            =   0
                  TabIndex        =   240
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line168 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line167 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   82
            Left            =   4560
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   125
            Top             =   1920
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   82
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   126
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   82
                  Left            =   0
                  TabIndex        =   325
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   82
                  Left            =   0
                  TabIndex        =   239
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line166 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line165 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   81
            Left            =   5160
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   123
            Top             =   1920
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   81
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   124
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   81
                  Left            =   0
                  TabIndex        =   324
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   81
                  Left            =   0
                  TabIndex        =   238
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line164 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line163 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   75
            Left            =   8520
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   121
            Top             =   1920
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   75
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   122
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   75
                  Left            =   0
                  TabIndex        =   318
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   75
                  Left            =   0
                  TabIndex        =   232
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line152 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line151 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   74
            Left            =   7920
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   119
            Top             =   1920
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   74
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   120
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   74
                  Left            =   0
                  TabIndex        =   317
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   74
                  Left            =   0
                  TabIndex        =   231
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line150 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line149 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   73
            Left            =   7320
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   117
            Top             =   1920
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   73
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   118
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   73
                  Left            =   0
                  TabIndex        =   316
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   73
                  Left            =   0
                  TabIndex        =   230
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line148 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line147 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   72
            Left            =   6720
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   115
            Top             =   1920
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   72
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   116
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   72
                  Left            =   0
                  TabIndex        =   315
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   72
                  Left            =   0
                  TabIndex        =   229
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line146 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line145 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   71
            Left            =   6120
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   113
            Top             =   1920
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   71
               Left            =   120
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   114
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   71
                  Left            =   0
                  TabIndex        =   314
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   71
                  Left            =   0
                  TabIndex        =   228
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line144 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line143 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   65
            Left            =   8520
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   111
            Top             =   1080
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   65
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   112
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   65
                  Left            =   0
                  TabIndex        =   308
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   65
                  Left            =   0
                  TabIndex        =   222
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line132 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line131 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   64
            Left            =   7920
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   109
            Top             =   1080
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   64
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   110
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   64
                  Left            =   0
                  TabIndex        =   307
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   64
                  Left            =   0
                  TabIndex        =   221
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line130 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line129 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   63
            Left            =   7320
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   107
            Top             =   1080
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   63
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   108
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   63
                  Left            =   0
                  TabIndex        =   306
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   63
                  Left            =   0
                  TabIndex        =   220
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line128 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line127 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   62
            Left            =   6720
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   105
            Top             =   1080
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   62
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   106
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   62
                  Left            =   0
                  TabIndex        =   305
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   62
                  Left            =   0
                  TabIndex        =   219
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line126 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line125 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   61
            Left            =   6120
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   103
            Top             =   1080
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   61
               Left            =   120
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   104
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   61
                  Left            =   0
                  TabIndex        =   304
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   61
                  Left            =   0
                  TabIndex        =   218
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line124 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line123 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   55
            Left            =   2760
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   101
            Top             =   1080
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   55
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   102
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   55
                  Left            =   0
                  TabIndex        =   298
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   55
                  Left            =   0
                  TabIndex        =   212
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line112 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line111 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   54
            Left            =   3360
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   99
            Top             =   1080
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   54
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   100
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   54
                  Left            =   0
                  TabIndex        =   297
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   54
                  Left            =   0
                  TabIndex        =   211
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line110 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line109 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   53
            Left            =   3960
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   97
            Top             =   1080
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   53
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   98
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   53
                  Left            =   0
                  TabIndex        =   296
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   53
                  Left            =   0
                  TabIndex        =   210
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line108 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line107 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   52
            Left            =   4560
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   95
            Top             =   1080
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   52
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   96
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   52
                  Left            =   0
                  TabIndex        =   295
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   52
                  Left            =   0
                  TabIndex        =   209
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line106 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
            Begin VB.Line Line105 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   51
            Left            =   5160
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   92
            Top             =   1080
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   51
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   93
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   51
                  Left            =   0
                  TabIndex        =   294
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   51
                  Left            =   0
                  TabIndex        =   208
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line104 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line103 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   48
            Left            =   960
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   90
            Top             =   2760
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   48
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   91
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   48
                  Left            =   0
                  TabIndex        =   291
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   48
                  Left            =   0
                  TabIndex        =   205
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line98 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line97 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   47
            Left            =   1560
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   88
            Top             =   2760
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   47
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   89
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   47
                  Left            =   0
                  TabIndex        =   290
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   47
                  Left            =   0
                  TabIndex        =   204
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line96 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line95 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   46
            Left            =   2160
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   86
            Top             =   2760
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   46
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   87
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   46
                  Left            =   0
                  TabIndex        =   289
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   46
                  Left            =   0
                  TabIndex        =   203
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line94 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line93 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   45
            Left            =   2760
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   84
            Top             =   2760
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   45
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   85
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   45
                  Left            =   0
                  TabIndex        =   288
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   45
                  Left            =   0
                  TabIndex        =   202
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line92 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line91 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   44
            Left            =   3360
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   82
            Top             =   2760
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   44
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   83
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   44
                  Left            =   0
                  TabIndex        =   287
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   44
                  Left            =   0
                  TabIndex        =   201
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line90 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line89 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   43
            Left            =   3960
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   80
            Top             =   2760
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   43
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   81
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   43
                  Left            =   0
                  TabIndex        =   286
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   43
                  Left            =   0
                  TabIndex        =   200
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line88 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line87 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   42
            Left            =   4560
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   78
            Top             =   2760
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   42
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   79
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   42
                  Left            =   0
                  TabIndex        =   285
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   42
                  Left            =   0
                  TabIndex        =   199
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line86 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line85 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   510.968
               Y2              =   0
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   41
            Left            =   5160
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   76
            Top             =   2760
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   41
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   77
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   41
                  Left            =   0
                  TabIndex        =   284
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   41
                  Left            =   0
                  TabIndex        =   198
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line84 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line83 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   37
            Left            =   9720
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   72
            Top             =   2760
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   37
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   73
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   37
                  Left            =   0
                  TabIndex        =   280
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   37
                  Left            =   0
                  TabIndex        =   194
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line76 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line75 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   36
            Left            =   9120
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   70
            Top             =   2760
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   36
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   71
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   36
                  Left            =   0
                  TabIndex        =   279
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   36
                  Left            =   0
                  TabIndex        =   193
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line74 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line73 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   35
            Left            =   8520
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   68
            Top             =   2760
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   35
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   69
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   35
                  Left            =   0
                  TabIndex        =   278
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   35
                  Left            =   0
                  TabIndex        =   192
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line72 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line71 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   34
            Left            =   7920
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   66
            Top             =   2760
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   34
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   67
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   34
                  Left            =   0
                  TabIndex        =   277
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   34
                  Left            =   0
                  TabIndex        =   191
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line70 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line69 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   33
            Left            =   7320
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   64
            Top             =   2760
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   33
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   65
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   33
                  Left            =   0
                  TabIndex        =   276
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   33
                  Left            =   0
                  TabIndex        =   190
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line68 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line67 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   32
            Left            =   6720
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   62
            Top             =   2760
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   32
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   63
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   32
                  Left            =   0
                  TabIndex        =   275
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   32
                  Left            =   0
                  TabIndex        =   189
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line66 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line65 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   31
            Left            =   6120
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   60
            Top             =   2760
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   31
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   61
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   31
                  Left            =   0
                  TabIndex        =   274
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   31
                  Left            =   0
                  TabIndex        =   188
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line64 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line63 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   27
            Left            =   9720
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   56
            Top             =   240
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   27
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   57
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   27
                  Left            =   0
                  TabIndex        =   270
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   27
                  Left            =   0
                  TabIndex        =   184
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line56 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line55 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   26
            Left            =   9120
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   54
            Top             =   240
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   26
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   55
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   26
                  Left            =   0
                  TabIndex        =   269
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   26
                  Left            =   0
                  TabIndex        =   183
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line54 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line53 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   25
            Left            =   8520
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   52
            Top             =   240
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   25
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   53
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   25
                  Left            =   0
                  TabIndex        =   268
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   25
                  Left            =   0
                  TabIndex        =   182
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line52 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line51 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   24
            Left            =   7920
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   50
            Top             =   240
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   24
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   51
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   24
                  Left            =   0
                  TabIndex        =   267
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   24
                  Left            =   0
                  TabIndex        =   181
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line50 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line49 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   23
            Left            =   7320
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   48
            Top             =   240
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   23
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   49
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   23
                  Left            =   0
                  TabIndex        =   266
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   23
                  Left            =   0
                  TabIndex        =   180
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line48 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line47 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   22
            Left            =   6720
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   46
            Top             =   240
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   22
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   47
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   22
                  Left            =   0
                  TabIndex        =   265
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   22
                  Left            =   0
                  TabIndex        =   179
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line46 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line45 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   21
            Left            =   6120
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   44
            Top             =   240
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   21
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   45
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   21
                  Left            =   0
                  TabIndex        =   264
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   21
                  Left            =   0
                  TabIndex        =   178
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line44 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line43 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   17
            Left            =   1560
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   40
            Top             =   240
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   17
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   41
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   17
                  Left            =   120
                  TabIndex        =   260
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   17
                  Left            =   0
                  TabIndex        =   174
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line36 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line35 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   16
            Left            =   2160
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   38
            Top             =   240
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   16
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   39
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   16
                  Left            =   120
                  TabIndex        =   259
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   16
                  Left            =   0
                  TabIndex        =   173
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line34 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line33 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   15
            Left            =   2760
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   36
            Top             =   240
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   37
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   15
                  Left            =   120
                  TabIndex        =   258
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   15
                  Left            =   0
                  TabIndex        =   172
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line32 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line31 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   14
            Left            =   3360
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   34
            Top             =   240
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   35
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   14
                  Left            =   120
                  TabIndex        =   257
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   14
                  Left            =   0
                  TabIndex        =   171
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line30 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line29 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   13
            Left            =   3960
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   32
            Top             =   240
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   33
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   13
                  Left            =   50
                  TabIndex        =   256
                  Top             =   0
                  Width           =   75
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   13
                  Left            =   0
                  TabIndex        =   170
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line28 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line27 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   12
            Left            =   4560
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   30
            Top             =   240
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   31
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   12
                  Left            =   120
                  TabIndex        =   255
                  Top             =   0
                  Width           =   75
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   12
                  Left            =   0
                  TabIndex        =   169
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line26 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line25 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   11
            Left            =   5160
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   28
            Top             =   240
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   29
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   11
                  Left            =   0
                  TabIndex        =   254
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   11
                  Left            =   0
                  TabIndex        =   168
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line24 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line23 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   18
            Left            =   960
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   42
            Top             =   240
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   18
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   43
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   18
                  Left            =   120
                  TabIndex        =   261
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   18
                  Left            =   0
                  TabIndex        =   175
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line38 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line37 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   11
            Left            =   5160
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   393
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   12
            Left            =   4560
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   394
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   13
            Left            =   3960
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   395
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   14
            Left            =   3360
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   396
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   15
            Left            =   2760
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   397
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   16
            Left            =   2160
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   398
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   17
            Left            =   1560
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   399
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   18
            Left            =   960
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   400
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   21
            Left            =   6120
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   401
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   22
            Left            =   6720
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   402
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   23
            Left            =   7320
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   403
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   24
            Left            =   7920
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   404
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   25
            Left            =   8520
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   405
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   26
            Left            =   9120
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   406
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   61
            Left            =   6120
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   428
            Top             =   1080
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   62
            Left            =   6720
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   429
            Top             =   1080
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   63
            Left            =   7320
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   430
            Top             =   1080
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   64
            Left            =   7920
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   431
            Top             =   1080
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   65
            Left            =   8520
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   432
            Top             =   1080
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   31
            Left            =   6120
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   407
            Top             =   2760
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   32
            Left            =   6720
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   408
            Top             =   2760
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   33
            Left            =   7320
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   409
            Top             =   2760
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   34
            Left            =   7920
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   410
            Top             =   2760
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   35
            Left            =   8520
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   411
            Top             =   2760
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   36
            Left            =   9120
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   412
            Top             =   2760
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   41
            Left            =   5160
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   415
            Top             =   2760
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   42
            Left            =   4560
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   416
            Top             =   2760
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   43
            Left            =   3960
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   417
            Top             =   2760
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   44
            Left            =   3360
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   418
            Top             =   2760
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   45
            Left            =   2760
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   419
            Top             =   2760
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   46
            Left            =   2160
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   420
            Top             =   2760
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   47
            Left            =   1560
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   421
            Top             =   2760
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   48
            Left            =   960
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   422
            Top             =   2760
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   51
            Left            =   5160
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   423
            Top             =   1080
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   52
            Left            =   4560
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   424
            Top             =   1080
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   53
            Left            =   3960
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   425
            Top             =   1080
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   54
            Left            =   3360
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   426
            Top             =   1080
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   55
            Left            =   2760
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   427
            Top             =   1080
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   71
            Left            =   6120
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   433
            Top             =   1920
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   72
            Left            =   6720
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   434
            Top             =   1920
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   73
            Left            =   7320
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   435
            Top             =   1920
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   74
            Left            =   7920
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   436
            Top             =   1920
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   75
            Left            =   8520
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   437
            Top             =   1920
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   81
            Left            =   5160
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   438
            Top             =   1920
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   82
            Left            =   4560
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   439
            Top             =   1920
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   83
            Left            =   3960
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   440
            Top             =   1920
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   84
            Left            =   3360
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   441
            Top             =   1920
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   85
            Left            =   2760
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   442
            Top             =   1920
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   37
            Left            =   9720
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   413
            Top             =   2760
            Width           =   495
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   27
            Left            =   9720
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   443
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   28
            Left            =   10320
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   58
            Top             =   240
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   28
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   59
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   28
                  Left            =   0
                  TabIndex        =   271
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   28
                  Left            =   0
                  TabIndex        =   185
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line58 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line57 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   28
            Left            =   10320
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   444
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox picGigi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   38
            Left            =   10320
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   74
            Top             =   2760
            Width           =   495
            Begin VB.PictureBox picTengah 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   38
               Left            =   111
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   75
               Top             =   100
               Width           =   255
               Begin VB.Label lblGigiAnomali 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Height          =   195
                  Index           =   38
                  Left            =   0
                  TabIndex        =   281
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lblGigi 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   38
                  Left            =   0
                  TabIndex        =   195
                  Top             =   0
                  Width           =   75
               End
            End
            Begin VB.Line Line78 
               X1              =   0
               X2              =   495.484
               Y1              =   0
               Y2              =   495.484
            End
            Begin VB.Line Line77 
               X1              =   -15.484
               X2              =   495.484
               Y1              =   495.484
               Y2              =   -15.484
            End
         End
         Begin VB.PictureBox picBackGigi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   38
            Left            =   10320
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   414
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "85"
            Height          =   195
            Index           =   51
            Left            =   2880
            TabIndex        =   387
            Top             =   2400
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "84"
            Height          =   195
            Index           =   50
            Left            =   3480
            TabIndex        =   386
            Top             =   2400
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "83"
            Height          =   195
            Index           =   49
            Left            =   4080
            TabIndex        =   385
            Top             =   2400
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "82"
            Height          =   195
            Index           =   48
            Left            =   4680
            TabIndex        =   384
            Top             =   2400
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "81"
            Height          =   195
            Index           =   47
            Left            =   5280
            TabIndex        =   383
            Top             =   2400
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "75"
            Height          =   195
            Index           =   46
            Left            =   8640
            TabIndex        =   382
            Top             =   2400
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "74"
            Height          =   195
            Index           =   45
            Left            =   8040
            TabIndex        =   381
            Top             =   2400
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "73"
            Height          =   195
            Index           =   44
            Left            =   7440
            TabIndex        =   380
            Top             =   2400
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "72"
            Height          =   195
            Index           =   43
            Left            =   6840
            TabIndex        =   379
            Top             =   2400
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "71"
            Height          =   195
            Index           =   42
            Left            =   6240
            TabIndex        =   378
            Top             =   2400
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "65"
            Height          =   195
            Index           =   41
            Left            =   8640
            TabIndex        =   377
            Top             =   1560
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "64"
            Height          =   195
            Index           =   40
            Left            =   8040
            TabIndex        =   376
            Top             =   1560
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "63"
            Height          =   195
            Index           =   39
            Left            =   7440
            TabIndex        =   375
            Top             =   1560
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "62"
            Height          =   195
            Index           =   38
            Left            =   6840
            TabIndex        =   374
            Top             =   1560
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "61"
            Height          =   195
            Index           =   37
            Left            =   6240
            TabIndex        =   373
            Top             =   1560
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "55"
            Height          =   195
            Index           =   36
            Left            =   2880
            TabIndex        =   372
            Top             =   1560
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "54"
            Height          =   195
            Index           =   35
            Left            =   3480
            TabIndex        =   371
            Top             =   1560
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "53"
            Height          =   195
            Index           =   34
            Left            =   4080
            TabIndex        =   370
            Top             =   1560
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "52"
            Height          =   195
            Index           =   33
            Left            =   4680
            TabIndex        =   369
            Top             =   1560
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "51"
            Height          =   195
            Index           =   32
            Left            =   5280
            TabIndex        =   368
            Top             =   1560
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "48"
            Height          =   195
            Index           =   31
            Left            =   1080
            TabIndex        =   367
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "47"
            Height          =   195
            Index           =   30
            Left            =   1680
            TabIndex        =   366
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "46"
            Height          =   195
            Index           =   29
            Left            =   2280
            TabIndex        =   365
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "45"
            Height          =   195
            Index           =   28
            Left            =   2880
            TabIndex        =   364
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "44"
            Height          =   195
            Index           =   27
            Left            =   3480
            TabIndex        =   363
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "43"
            Height          =   195
            Index           =   26
            Left            =   4080
            TabIndex        =   362
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "42"
            Height          =   195
            Index           =   25
            Left            =   4680
            TabIndex        =   361
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "41"
            Height          =   195
            Index           =   24
            Left            =   5280
            TabIndex        =   360
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "38"
            Height          =   195
            Index           =   23
            Left            =   10440
            TabIndex        =   359
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "37"
            Height          =   195
            Index           =   22
            Left            =   9840
            TabIndex        =   358
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "36"
            Height          =   195
            Index           =   21
            Left            =   9240
            TabIndex        =   357
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "35"
            Height          =   195
            Index           =   20
            Left            =   8640
            TabIndex        =   356
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "34"
            Height          =   195
            Index           =   19
            Left            =   8040
            TabIndex        =   355
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "33"
            Height          =   195
            Index           =   18
            Left            =   7440
            TabIndex        =   354
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "32"
            Height          =   195
            Index           =   17
            Left            =   6840
            TabIndex        =   353
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "31"
            Height          =   195
            Index           =   16
            Left            =   6240
            TabIndex        =   352
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "28"
            Height          =   195
            Index           =   15
            Left            =   10440
            TabIndex        =   351
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "27"
            Height          =   195
            Index           =   14
            Left            =   9840
            TabIndex        =   350
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "26"
            Height          =   195
            Index           =   13
            Left            =   9240
            TabIndex        =   349
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "25"
            Height          =   195
            Index           =   12
            Left            =   8640
            TabIndex        =   348
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "24"
            Height          =   195
            Index           =   11
            Left            =   8040
            TabIndex        =   347
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "23"
            Height          =   195
            Index           =   10
            Left            =   7440
            TabIndex        =   346
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "22"
            Height          =   195
            Index           =   9
            Left            =   6840
            TabIndex        =   345
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "21"
            Height          =   195
            Index           =   8
            Left            =   6240
            TabIndex        =   344
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "18"
            Height          =   195
            Index           =   7
            Left            =   1080
            TabIndex        =   343
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "17"
            Height          =   195
            Index           =   6
            Left            =   1680
            TabIndex        =   342
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "16"
            Height          =   195
            Index           =   5
            Left            =   2280
            TabIndex        =   341
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "15"
            Height          =   195
            Index           =   4
            Left            =   2880
            TabIndex        =   340
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "14"
            Height          =   195
            Index           =   3
            Left            =   3480
            TabIndex        =   339
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "13"
            Height          =   195
            Index           =   2
            Left            =   4080
            TabIndex        =   338
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "12"
            Height          =   195
            Index           =   1
            Left            =   4680
            TabIndex        =   337
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblNoGigi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "11"
            Height          =   195
            Index           =   0
            Left            =   5280
            TabIndex        =   336
            Top             =   720
            Width           =   180
         End
      End
      Begin VB.PictureBox picTemp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   840
         ScaleHeight     =   3495
         ScaleWidth      =   10455
         TabIndex        =   388
         Top             =   240
         Visible         =   0   'False
         Width           =   10455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Catatan"
         Height          =   195
         Left            =   120
         TabIndex        =   392
         Top             =   3840
         Width           =   555
      End
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12720
      Picture         =   "frmDiagramOdonto.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDiagramOdonto.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDiagramOdonto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WARNA_TRANSPARAN = &H0
Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Declare Function GetDC Lib "user32" ( _
ByVal hWnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32" ( _
ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" ( _
ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
ByVal hDC As Long, ByVal nWidth As Long, _
ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" ( _
ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" ( _
ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" ( _
ByVal hDC As Long, ByVal wStartIndex As Long, _
ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) _
As Long
Private Declare Function CreatePalette Lib "gdi32" ( _
lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "gdi32" ( _
ByVal hDC As Long, ByVal hPalette As Long, _
ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" ( _
ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" ( _
ByVal hDCDest As Long, ByVal XDest As Long, _
ByVal YDest As Long, ByVal nWidth As Long, _
ByVal nHeight As Long, ByVal hDCSrc As Long, _
ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) _
As Long
Private Declare Function ReleaseDC Lib "user32" ( _
ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" ( _
ByVal hDC As Long) As Long
Private Declare Function OleCreatePictureIndirect _
Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, _
ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Private varStatusAksi As StatusAksi
Private varKondisiGigi(85) As KondisiGigi
Private varKondisiWarnaBagianGigi(85) As KondisiWarnaBagianGigi
Private adoCommand As New ADODB.Command

Private Enum StatusAksi
    NORMAL_A = 0
    BELUM_ERUPSI = 1
    ERUPSI_SEBAGIAN = 2
    ANOMALI_BENTUK = 3
    KARIES = 4
    NON_VITAL = 5
    TAMBALAN_LOGAM = 6
    TAMBALAN_NON_LOGAM = 7
    MAHKOTA_LOGAM = 8
    MAHKOTA_NON_LOGAM = 9
    SISA_AKAR = 10
    GIGI_HILANG = 11
    JEMBATAN_A = 12
    GIGI_TIRUAN_LEPAS = 13
    Calculus = 14
End Enum

Private Enum BagianGigi
    DEPAN_PART = 1
    KIRI_PART = 2
    KANAN_PART = 3
    ATAS_PART = 4
    BAWAH_PART = 5
End Enum

Private Type KondisiGigi
    AdaGigi As Boolean
    BelumErupsi As String
    ErupsiSebagian As String
    AnomaliBentuk As String
    KariesDepan As String
    KariesKiri As String
    KariesKanan As String
    KariesAtas As String
    KariesBawah As String
    NonVital As String
    TambalanLogamDepan As String
    TambalanLogamKiri As String
    TambalanLogamKanan As String
    TambalanLogamAtas As String
    TambalanLogamBawah As String
    TambalanNonLogamDepan As String
    TambalanNonLogamKiri As String
    TambalanNonLogamKanan As String
    TambalanNonLogamAtas As String
    TambalanNonLogamBawah As String
    MahkotaLogamDepan As String
    MahkotaLogamKiri As String
    MahkotaLogamKanan As String
    MahkotaLogamAtas As String
    MahkotaLogamBawah As String
    MahkotaNonLogamDepan As String
    MahkotaNonLogamKiri As String
    MahkotaNonLogamKanan As String
    MahkotaNonLogamAtas As String
    MahkotaNonLogamBawah As String
    SisaAkar As String
    GigiHilang As String
    Jembatan As String
    GigiTiruanLepas As String
    Calculus As String
End Type

Private Type KondisiWarnaBagianGigi
    Depan As Boolean
    Kiri As Boolean
    Kanan As Boolean
    Atas As Boolean
    Bawah As Boolean
End Type

Private Type PALETTEENTRY
    peRed   As Byte
    peGreen As Byte
    peBlue  As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion       As Integer
    palNumEntries    As Integer
    palPalEntry(255) As PALETTEENTRY
End Type

Private Type PicBmp
    Size As Long
Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Type GUID
    Data1    As Long
    Data2    As Integer
    Data3    As Integer
    Data4(7) As Byte
End Type

Private Type SumbuKoordinat
    SmbX As Single
    SmbY As Single
End Type

Public Function CreateBitmapPicture(ByVal hBmp As Long, _
    ByVal hPal As Long) As Picture
    Dim r As Long
    Dim Pic As PicBmp
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID

    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    With Pic
        .Size = Len(Pic)
        .Type = vbPicTypeBitmap
        .hBmp = hBmp
        .hPal = hPal
    End With

    r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

    Set CreateBitmapPicture = IPic
End Function

Public Function CaptureWindow(ByVal hWndSrc As Long, _
    ByVal bClient As Boolean, ByVal LeftSrc As Long, _
    ByVal TopSrc As Long, ByVal WidthSrc As Long, _
    ByVal HeightSrc As Long) As Picture
    Dim hDCMemory       As Long
    Dim hBmp            As Long
    Dim hBmpPrev        As Long
    Dim r               As Long
    Dim hDCSrc          As Long
    Dim hPal            As Long
    Dim hPalPrev        As Long
    Dim RasterCapsScrn  As Long
    Dim HasPaletteScrn  As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal          As LOGPALETTE

    If bClient Then
        hDCSrc = GetDC(hWndSrc)
    Else
        hDCSrc = GetWindowDC(hWndSrc)
    End If

    hDCMemory = CreateCompatibleDC(hDCSrc)

    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)   'Raster capabilities
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE       'Palette support
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) 'Palette size

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then

        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)

        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        r = RealizePalette(hDCMemory)
    End If

    r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
    LeftSrc, TopSrc, vbSrcCopy)

    hBmp = SelectObject(hDCMemory, hBmpPrev)

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If

    r = DeleteDC(hDCMemory)
    r = ReleaseDC(hWndSrc, hDCSrc)

    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function

Private Sub subWarnaiBagianGigi(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single, Warna As ColorConstants)
    Dim i As Integer
    Dim sngLTengah As Single, sngTTengah As Single
    Dim varWarna As ColorConstants

    If Button <> 1 Then Exit Sub
    sngLTengah = Me.picTengah(Index).Left + Me.picTengah(Index).Width
    sngTTengah = Me.picTengah(Index).Top + Me.picTengah(Index).Height
    If (X > 0 And X < Me.picTengah(Index).Left) And (y > Me.picTengah(Index).Top And y < sngTTengah) Then 'kiri
        Select Case varStatusAksi
            Case KARIES
                If varKondisiGigi(Index).KariesKiri = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).KariesKiri = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesKiri = "Y"
                        .TambalanLogamKiri = "T"
                        .TambalanNonLogamKiri = "T"
                        .MahkotaLogamKiri = "T"
                        .MahkotaNonLogamKiri = "T"
                    End With
                End If
            Case TAMBALAN_LOGAM
                If varKondisiGigi(Index).TambalanLogamKiri = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).TambalanLogamKiri = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesKiri = "T"
                        .TambalanLogamKiri = "Y"
                        .TambalanNonLogamKiri = "T"
                        .MahkotaLogamKiri = "T"
                        .MahkotaNonLogamKiri = "T"
                    End With
                End If
            Case TAMBALAN_NON_LOGAM
                If varKondisiGigi(Index).TambalanNonLogamKiri = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).TambalanNonLogamKiri = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesKiri = "T"
                        .TambalanLogamKiri = "T"
                        .TambalanNonLogamKiri = "Y"
                        .MahkotaLogamKiri = "T"
                        .MahkotaNonLogamKiri = "T"
                    End With
                End If
            Case MAHKOTA_LOGAM
                If varKondisiGigi(Index).MahkotaLogamKiri = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).MahkotaLogamKiri = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesKiri = "Y"
                        .TambalanLogamKiri = "T"
                        .TambalanNonLogamKiri = "T"
                        .MahkotaLogamKiri = "Y"
                        .MahkotaNonLogamKiri = "T"
                    End With
                End If
            Case MAHKOTA_NON_LOGAM
                If varKondisiGigi(Index).MahkotaNonLogamKiri = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).MahkotaNonLogamKiri = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesKiri = "T"
                        .TambalanLogamKiri = "T"
                        .TambalanNonLogamKiri = "T"
                        .MahkotaLogamKiri = "T"
                        .MahkotaNonLogamKiri = "Y"
                    End With
                End If
        End Select
        For i = 0 To Me.picTengah(Index).Left
            Me.picGigi(Index).Line (i, i)-(i, Me.picGigi(Index).ScaleHeight - i), varWarna
        Next
    ElseIf X > sngLTengah And X < Me.picGigi(Index).ScaleWidth Then 'kanan
        Select Case varStatusAksi
            Case KARIES
                If varKondisiGigi(Index).KariesKanan = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).KariesKanan = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesKanan = "Y"
                        .TambalanLogamKanan = "T"
                        .TambalanNonLogamKanan = "T"
                        .MahkotaLogamKanan = "T"
                        .MahkotaNonLogamKanan = "T"
                    End With
                End If
            Case TAMBALAN_LOGAM
                If varKondisiGigi(Index).TambalanLogamKanan = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).TambalanLogamKanan = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesKanan = "T"
                        .TambalanLogamKanan = "Y"
                        .TambalanNonLogamKanan = "T"
                        .MahkotaLogamKanan = "T"
                        .MahkotaNonLogamKanan = "T"
                    End With
                End If
            Case TAMBALAN_NON_LOGAM
                If varKondisiGigi(Index).TambalanNonLogamKanan = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).TambalanNonLogamKanan = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesKanan = "T"
                        .TambalanLogamKanan = "T"
                        .TambalanNonLogamKanan = "Y"
                        .MahkotaLogamKanan = "T"
                        .MahkotaNonLogamKanan = "T"
                    End With
                End If
            Case MAHKOTA_LOGAM
                If varKondisiGigi(Index).MahkotaLogamKanan = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).MahkotaLogamKanan = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesKanan = "T"
                        .TambalanLogamKanan = "T"
                        .TambalanNonLogamKanan = "T"
                        .MahkotaLogamKanan = "Y"
                        .MahkotaNonLogamKanan = "T"
                    End With
                End If
            Case MAHKOTA_NON_LOGAM
                If varKondisiGigi(Index).MahkotaNonLogamKanan = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).MahkotaNonLogamKanan = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesKanan = "T"
                        .TambalanLogamKanan = "T"
                        .TambalanNonLogamKanan = "T"
                        .MahkotaLogamKanan = "T"
                        .MahkotaNonLogamKanan = "Y"
                    End With
                End If
        End Select
        For i = sngLTengah To Me.picGigi(Index).ScaleWidth
            Me.picGigi(Index).Line (i, i)-(i, Me.picGigi(Index).ScaleHeight - i), varWarna
        Next
    ElseIf y > 0 And y < Me.picTengah(Index).Top Then 'atas
        Select Case varStatusAksi
            Case KARIES
                If varKondisiGigi(Index).KariesAtas = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).KariesAtas = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesAtas = "Y"
                        .TambalanLogamAtas = "T"
                        .TambalanNonLogamAtas = "T"
                        .MahkotaLogamAtas = "T"
                        .MahkotaNonLogamAtas = "T"
                    End With
                End If
            Case TAMBALAN_LOGAM
                If varKondisiGigi(Index).TambalanLogamAtas = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).TambalanLogamAtas = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesAtas = "T"
                        .TambalanLogamAtas = "Y"
                        .TambalanNonLogamAtas = "T"
                        .MahkotaLogamAtas = "T"
                        .MahkotaNonLogamAtas = "T"
                    End With
                End If
            Case TAMBALAN_NON_LOGAM
                If varKondisiGigi(Index).TambalanNonLogamAtas = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).TambalanNonLogamAtas = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesAtas = "T"
                        .TambalanLogamAtas = "T"
                        .TambalanNonLogamAtas = "Y"
                        .MahkotaLogamAtas = "T"
                        .MahkotaNonLogamAtas = "T"
                    End With
                End If
            Case MAHKOTA_LOGAM
                If varKondisiGigi(Index).MahkotaLogamAtas = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).MahkotaLogamAtas = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesAtas = "T"
                        .TambalanLogamAtas = "T"
                        .TambalanNonLogamAtas = "T"
                        .MahkotaLogamAtas = "Y"
                        .MahkotaNonLogamAtas = "T"
                    End With
                End If
            Case MAHKOTA_NON_LOGAM
                If varKondisiGigi(Index).MahkotaNonLogamAtas = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).MahkotaNonLogamAtas = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesAtas = "T"
                        .TambalanLogamAtas = "T"
                        .TambalanNonLogamAtas = "T"
                        .MahkotaLogamAtas = "T"
                        .MahkotaNonLogamAtas = "Y"
                    End With
                End If
        End Select
        For i = 0 To Me.picTengah(Index).Top
            Me.picGigi(Index).Line (i, i)-(Me.picGigi(Index).ScaleWidth - i, i), varWarna
        Next
    ElseIf y > sngTTengah And y < Me.picGigi(Index).ScaleHeight Then 'bawah
        Select Case varStatusAksi
            Case KARIES
                If varKondisiGigi(Index).KariesBawah = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).KariesBawah = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesBawah = "Y"
                        .TambalanLogamBawah = "T"
                        .TambalanNonLogamBawah = "T"
                        .MahkotaLogamBawah = "T"
                        .MahkotaNonLogamBawah = "T"
                    End With
                End If
            Case TAMBALAN_LOGAM
                If varKondisiGigi(Index).TambalanLogamBawah = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).TambalanLogamBawah = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesBawah = "T"
                        .TambalanLogamBawah = "Y"
                        .TambalanNonLogamBawah = "T"
                        .MahkotaLogamBawah = "T"
                        .MahkotaNonLogamBawah = "T"
                    End With
                End If
            Case TAMBALAN_NON_LOGAM
                If varKondisiGigi(Index).TambalanNonLogamBawah = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).TambalanNonLogamBawah = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesBawah = "T"
                        .TambalanLogamBawah = "T"
                        .TambalanNonLogamBawah = "Y"
                        .MahkotaLogamBawah = "T"
                        .MahkotaNonLogamBawah = "T"
                    End With
                End If
            Case MAHKOTA_LOGAM
                If varKondisiGigi(Index).MahkotaLogamBawah = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).MahkotaLogamBawah = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesBawah = "T"
                        .TambalanLogamBawah = "T"
                        .TambalanNonLogamBawah = "T"
                        .MahkotaLogamBawah = "Y"
                        .MahkotaNonLogamBawah = "T"
                    End With
                End If
            Case MAHKOTA_NON_LOGAM
                If varKondisiGigi(Index).MahkotaNonLogamBawah = "Y" Then
                    varWarna = vbWhite
                    varKondisiGigi(Index).MahkotaNonLogamBawah = "T"
                Else
                    varWarna = Warna
                    With varKondisiGigi(Index)
                        .KariesBawah = "T"
                        .TambalanLogamBawah = "T"
                        .TambalanNonLogamBawah = "T"
                        .MahkotaLogamBawah = "T"
                        .MahkotaNonLogamBawah = "Y"
                    End With
                End If
        End Select
        For i = sngTTengah To Me.picGigi(Index).ScaleHeight
            Me.picGigi(Index).Line (i, i)-(Me.picGigi(Index).ScaleWidth - i, i), varWarna
        Next
    End If
End Sub

Private Sub subSetBagianDepan(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single, Warna As ColorConstants)
    Dim varWarna As ColorConstants

    If Button <> 1 Then Exit Sub
    Select Case varStatusAksi
        Case KARIES
            If varKondisiGigi(Index).KariesDepan = "Y" Then
                varWarna = vbWhite
                varKondisiGigi(Index).KariesDepan = "T"
            Else
                varWarna = Warna
                With varKondisiGigi(Index)
                    .KariesDepan = "Y"
                    .TambalanLogamDepan = "T"
                    .TambalanNonLogamDepan = "T"
                    .MahkotaLogamDepan = "T"
                    .MahkotaNonLogamDepan = "T"
                End With
            End If
        Case TAMBALAN_LOGAM
            If varKondisiGigi(Index).TambalanLogamDepan = "Y" Then
                varWarna = vbWhite
                varKondisiGigi(Index).TambalanLogamDepan = "T"
            Else
                varWarna = Warna
                With varKondisiGigi(Index)
                    .KariesDepan = "T"
                    .TambalanLogamDepan = "Y"
                    .TambalanNonLogamDepan = "T"
                    .MahkotaLogamDepan = "T"
                    .MahkotaNonLogamDepan = "T"
                End With
            End If
        Case TAMBALAN_NON_LOGAM
            If varKondisiGigi(Index).TambalanNonLogamDepan = "Y" Then
                varWarna = vbWhite
                varKondisiGigi(Index).TambalanNonLogamDepan = "T"
            Else
                varWarna = Warna
                With varKondisiGigi(Index)
                    .KariesDepan = "T"
                    .TambalanLogamDepan = "T"
                    .TambalanNonLogamDepan = "Y"
                    .MahkotaLogamDepan = "T"
                    .MahkotaNonLogamDepan = "T"
                End With
            End If
        Case MAHKOTA_LOGAM
            If varKondisiGigi(Index).MahkotaLogamDepan = "Y" Then
                varWarna = vbWhite
                varKondisiGigi(Index).MahkotaLogamDepan = "T"
            Else
                varWarna = Warna
                With varKondisiGigi(Index)
                    .KariesDepan = "T"
                    .TambalanLogamDepan = "T"
                    .TambalanNonLogamDepan = "T"
                    .MahkotaLogamDepan = "Y"
                    .MahkotaNonLogamDepan = "T"
                End With
            End If
        Case MAHKOTA_NON_LOGAM
            If varKondisiGigi(Index).MahkotaNonLogamDepan = "Y" Then
                varWarna = vbWhite
                varKondisiGigi(Index).MahkotaNonLogamDepan = "T"
            Else
                varWarna = Warna
                With varKondisiGigi(Index)
                    .KariesDepan = "T"
                    .TambalanLogamDepan = "T"
                    .TambalanNonLogamDepan = "T"
                    .MahkotaLogamDepan = "T"
                    .MahkotaNonLogamDepan = "Y"
                End With
            End If
    End Select
    Me.picTengah(Index).BackColor = varWarna
    Call subRefreshGigiTengah(Index)
End Sub

Private Sub subResetArray()
    Dim i As Integer

    For i = 0 To 85
        With Me.lblGigi(i)
            .BackStyle = 0
        End With
        With Me.lblGigiAnomali(i)
            .BackStyle = 0
            .Left = 60
            .FontBold = True
        End With
        With varKondisiGigi(i)
            .AdaGigi = False
            .BelumErupsi = "T"
            .ErupsiSebagian = "T"
            .AnomaliBentuk = "T"
            .Calculus = "T"
            .KariesDepan = "T"
            .KariesKiri = "T"
            .KariesKanan = "T"
            .KariesAtas = "T"
            .KariesBawah = "T"
            .NonVital = "T"
            .TambalanLogamDepan = "T"
            .TambalanLogamKiri = "T"
            .TambalanLogamKanan = "T"
            .TambalanLogamAtas = "T"
            .TambalanLogamBawah = "T"
            .TambalanNonLogamDepan = "T"
            .TambalanNonLogamKiri = "T"
            .TambalanNonLogamKanan = "T"
            .TambalanNonLogamAtas = "T"
            .TambalanNonLogamBawah = "T"
            .MahkotaLogamDepan = "T"
            .MahkotaLogamKiri = "T"
            .MahkotaLogamKanan = "T"
            .MahkotaLogamAtas = "T"
            .MahkotaLogamBawah = "T"
            .MahkotaNonLogamDepan = "T"
            .MahkotaNonLogamKiri = "T"
            .MahkotaNonLogamKanan = "T"
            .MahkotaNonLogamAtas = "T"
            .MahkotaNonLogamBawah = "T"
            .SisaAkar = "T"
            .GigiHilang = "T"
            .Jembatan = "T"
            .GigiTiruanLepas = "T"
        End With
        With varKondisiWarnaBagianGigi(i)
            .Depan = False
            .Kiri = False
            .Kanan = False
            .Atas = False
            .Bawah = False
        End With
    Next
End Sub

Private Function Add_CatatanOdonto() As Boolean
    On Error GoTo errSimpan
    Set adoCommand = Nothing
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dtpTglPeriksa.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, mstrKdDokter)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("DiagramOdonto", adBinary, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(Trim(Len(Me.txtKeterangan.Text)) = 0, Null, Trim(Me.txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_CatatanOdonto"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data", vbCritical, "Validasi"
            Add_CatatanOdonto = False
        Else
            Add_CatatanOdonto = True
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Function
errSimpan:
    Call deleteADOCommandParameters(adoCommand)
    Set adoCommand = Nothing
    Call msubPesanError
End Function

Private Function Add_DetailCatatanOdonto(ByVal NoDiagramOdonto As String) As Boolean
    On Error GoTo errSimpan
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dtpTglPeriksa.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("NoDiagramOdonto", adTinyInt, adParamInput, , NoDiagramOdonto)
        .Parameters.Append .CreateParameter("BelumErupsi", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).BelumErupsi)
        .Parameters.Append .CreateParameter("ErupsiSebagian", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).ErupsiSebagian)
        .Parameters.Append .CreateParameter("AnomaliBentuk", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).AnomaliBentuk)

        .Parameters.Append .CreateParameter("Calculus", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).Calculus)

        .Parameters.Append .CreateParameter("KariesDepan", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).KariesDepan)
        .Parameters.Append .CreateParameter("KariesKiri", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).KariesKiri)
        .Parameters.Append .CreateParameter("KariesKanan", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).KariesKanan)
        .Parameters.Append .CreateParameter("KariesAtas", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).KariesAtas)
        .Parameters.Append .CreateParameter("KariesBawah", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).KariesBawah)
        .Parameters.Append .CreateParameter("NonVital", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).NonVital)
        .Parameters.Append .CreateParameter("TamblanLogamDepan", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).TambalanLogamDepan)
        .Parameters.Append .CreateParameter("TamblanLogamKiri", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).TambalanLogamKiri)
        .Parameters.Append .CreateParameter("TamblanLogamKanan", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).TambalanLogamKanan)
        .Parameters.Append .CreateParameter("TamblanLogamAtas", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).TambalanLogamAtas)
        .Parameters.Append .CreateParameter("TamblanLogamBawah", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).TambalanLogamBawah)
        .Parameters.Append .CreateParameter("TamblanNonLogamDepan", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).TambalanNonLogamDepan)
        .Parameters.Append .CreateParameter("TamblanNonLogamKiri", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).TambalanNonLogamKiri)
        .Parameters.Append .CreateParameter("TamblanNonLogamKanan", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).TambalanNonLogamKanan)
        .Parameters.Append .CreateParameter("TamblanNonLogamAtas", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).TambalanNonLogamAtas)
        .Parameters.Append .CreateParameter("TamblanNonLogamBawah", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).TambalanNonLogamBawah)
        .Parameters.Append .CreateParameter("MahkotaLogamDepan", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).MahkotaLogamDepan)
        .Parameters.Append .CreateParameter("MahkotaLogamKiri", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).MahkotaLogamKiri)
        .Parameters.Append .CreateParameter("MahkotaLogamKanan", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).MahkotaLogamKanan)
        .Parameters.Append .CreateParameter("MahkotaLogamAtas", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).MahkotaLogamAtas)
        .Parameters.Append .CreateParameter("MahkotaLogamBawah", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).MahkotaLogamBawah)
        .Parameters.Append .CreateParameter("MahkotaNonLogamDepan", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).MahkotaNonLogamDepan)
        .Parameters.Append .CreateParameter("MahkotaNonLogamKiri", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).MahkotaNonLogamKiri)
        .Parameters.Append .CreateParameter("MahkotaNonLogamKanan", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).MahkotaNonLogamKanan)
        .Parameters.Append .CreateParameter("MahkotaNonLogamAtas", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).MahkotaNonLogamAtas)
        .Parameters.Append .CreateParameter("MahkotaNonLogamBawah", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).MahkotaNonLogamBawah)
        .Parameters.Append .CreateParameter("SisaAkar", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).SisaAkar)
        .Parameters.Append .CreateParameter("GigiHilang", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).GigiHilang)
        .Parameters.Append .CreateParameter("Jembatan", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).Jembatan)
        .Parameters.Append .CreateParameter("GigiTiruanLepas", adChar, adParamInput, 1, varKondisiGigi(NoDiagramOdonto).GigiTiruanLepas)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DetailCatatanOdonto"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data", vbCritical, "Validasi"
            Add_DetailCatatanOdonto = False
        Else
            Add_DetailCatatanOdonto = True
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Function
errSimpan:
    Call deleteADOCommandParameters(adoCommand)
    Set adoCommand = Nothing
    msubPesanError
End Function

Private Function funcKonvertKeKoordinat(varBagianGigi As BagianGigi) As SumbuKoordinat
    With funcKonvertKeKoordinat
        Select Case varBagianGigi
            Case KIRI_PART
                .SmbX = 10
                .SmbY = Me.picTengah(11).ScaleHeight / 2
            Case KANAN_PART
                .SmbX = Me.picTengah(11).Left + Me.picTengah(11).Width + 10
                .SmbY = Me.picTengah(11).ScaleHeight / 2
            Case ATAS_PART
                .SmbX = (Me.picTengah(11).ScaleWidth / 2)
                .SmbY = 10
            Case BAWAH_PART
                .SmbX = (Me.picTengah(11).ScaleWidth / 2)
                .SmbY = Me.picTengah(11).Top + Me.picTengah(11).Height + 10
        End Select
    End With
End Function

Public Sub subLoadDetailCatatanOdonto()
    Dim idx As Integer
    Dim intX As Single, intY As Single

    strSQL = "select * from DetailCatatanOdonto where NoPendaftaran='" & Me.txtNoPendaftaran.Text & "'"

    Call msubRecFO(rs, strSQL)
    While Not rs.EOF
        idx = rs.Fields.Item("NoDiagramOdonto").value
        If rs.Fields.Item("BelumErupsi").value = "Y" Then
            varStatusAksi = BELUM_ERUPSI
            Call picGigi_MouseUp(idx, 1, 0, intX, intY)
        End If
        If rs.Fields.Item("ErupsiSebagian").value = "Y" Then
            varStatusAksi = ERUPSI_SEBAGIAN
            Call picGigi_MouseUp(idx, 1, 0, intX, intY)
        End If
        If rs.Fields.Item("AnomaliBentuk").value = "Y" Then
            varStatusAksi = ANOMALI_BENTUK
            Call picGigi_MouseUp(idx, 1, 0, intX, intY)
        End If
        If rs.Fields.Item("Calculus").value = "Y" Then
            varStatusAksi = Calculus
            Call picGigi_MouseUp(idx, 1, 0, intX, intY)
        End If
        If rs.Fields.Item("KariesDepan").value = "Y" Then
            varStatusAksi = KARIES
            Call picTengah_MouseUp(idx, 1, 0, 50, 50)
        End If
        If rs.Fields.Item("KariesKiri").value = "Y" Then
            varStatusAksi = KARIES
            With funcKonvertKeKoordinat(KIRI_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If
        If rs.Fields.Item("KariesKanan").value = "Y" Then
            varStatusAksi = KARIES
            With funcKonvertKeKoordinat(KANAN_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If
        If rs.Fields.Item("KariesAtas").value = "Y" Then
            varStatusAksi = KARIES
            With funcKonvertKeKoordinat(DEPAN_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If
        If rs.Fields.Item("KariesBawah").value = "Y" Then
            varStatusAksi = KARIES
            With funcKonvertKeKoordinat(BAWAH_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If
        If rs.Fields.Item("NonVital").value = "Y" Then
            varStatusAksi = NON_VITAL
            Call picTengah_MouseUp(idx, 1, 0, 10, 10)
        End If

        If rs.Fields.Item("TambalanLogamDepan").value = "Y" Then
            varStatusAksi = TAMBALAN_LOGAM
            Call picTengah_MouseUp(idx, 1, 0, 10, 10)
        End If
        If rs.Fields.Item("TambalanLogamKiri").value = "Y" Then
            varStatusAksi = TAMBALAN_LOGAM
            With funcKonvertKeKoordinat(KIRI_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If
        If rs.Fields.Item("TambalanLogamKanan").value = "Y" Then
            varStatusAksi = TAMBALAN_LOGAM
            With funcKonvertKeKoordinat(KANAN_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If
        If rs.Fields.Item("TambalanLogamAtas").value = "Y" Then
            varStatusAksi = TAMBALAN_LOGAM
            With funcKonvertKeKoordinat(ATAS_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If
        If rs.Fields.Item("TambalanLogamBawah").value = "Y" Then
            varStatusAksi = TAMBALAN_LOGAM
            With funcKonvertKeKoordinat(BAWAH_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If

        If rs.Fields.Item("TambalanNonLogamDepan").value = "Y" Then
            varStatusAksi = TAMBALAN_NON_LOGAM
            Call picTengah_MouseUp(idx, 1, 0, 10, 10)
        End If
        If rs.Fields.Item("TambalanNonLogamKiri").value = "Y" Then
            varStatusAksi = TAMBALAN_NON_LOGAM
            With funcKonvertKeKoordinat(KIRI_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If
        If rs.Fields.Item("TambalanNonLogamKanan").value = "Y" Then
            varStatusAksi = TAMBALAN_NON_LOGAM
            With funcKonvertKeKoordinat(KANAN_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If
        If rs.Fields.Item("TambalanNonLogamAtas").value = "Y" Then
            varStatusAksi = TAMBALAN_NON_LOGAM
            With funcKonvertKeKoordinat(ATAS_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If
        If rs.Fields.Item("TambalanNonLogamBawah").value = "Y" Then
            varStatusAksi = TAMBALAN_NON_LOGAM
            With funcKonvertKeKoordinat(BAWAH_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If

        If rs.Fields.Item("MahkotaLogamDepan").value = "Y" Then
            varStatusAksi = MAHKOTA_LOGAM
            Call picTengah_MouseUp(idx, 1, 0, 10, 10)
        End If
        If rs.Fields.Item("MahkotaLogamKiri").value = "Y" Then
            varStatusAksi = MAHKOTA_LOGAM
            With funcKonvertKeKoordinat(KIRI_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If
        If rs.Fields.Item("MahkotaLogamKanan").value = "Y" Then
            varStatusAksi = MAHKOTA_LOGAM
            With funcKonvertKeKoordinat(KANAN_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If
        If rs.Fields.Item("MahkotaLogamAtas").value = "Y" Then
            varStatusAksi = MAHKOTA_LOGAM
            With funcKonvertKeKoordinat(ATAS_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If
        If rs.Fields.Item("MahkotaLogamBawah").value = "Y" Then
            varStatusAksi = MAHKOTA_LOGAM
            With funcKonvertKeKoordinat(BAWAH_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If

        If rs.Fields.Item("MahkotaNonLogamDepan").value = "Y" Then
            varStatusAksi = MAHKOTA_NON_LOGAM
            Call picTengah_MouseUp(idx, 1, 0, 10, 10)
        End If
        If rs.Fields.Item("MahkotaNonLogamKiri").value = "Y" Then
            varStatusAksi = MAHKOTA_NON_LOGAM
            With funcKonvertKeKoordinat(KIRI_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If
        If rs.Fields.Item("MahkotaNonLogamKanan").value = "Y" Then
            varStatusAksi = MAHKOTA_NON_LOGAM
            With funcKonvertKeKoordinat(KANAN_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If
        If rs.Fields.Item("MahkotaNonLogamAtas").value = "Y" Then
            varStatusAksi = MAHKOTA_NON_LOGAM
            With funcKonvertKeKoordinat(ATAS_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If
        If rs.Fields.Item("MahkotaNonLogamBawah").value = "Y" Then
            varStatusAksi = MAHKOTA_NON_LOGAM
            With funcKonvertKeKoordinat(BAWAH_PART)
                Call picGigi_MouseUp(idx, 1, 0, .SmbX, .SmbY)
            End With
        End If

        If rs.Fields.Item("SisaAkar").value = "Y" Then
            varStatusAksi = SISA_AKAR
            Call picTengah_MouseUp(idx, 1, 0, 10, 10)
        End If
        If rs.Fields.Item("GigiHilang").value = "Y" Then
            varStatusAksi = GIGI_HILANG
            Call picTengah_MouseUp(idx, 1, 0, 10, 10)
        End If
        If rs.Fields.Item("Jembatan").value = "Y" Then
            varStatusAksi = JEMBATAN_A
            Call picTengah_MouseUp(idx, 1, 0, 10, 10)
        End If
        If rs.Fields.Item("GigiTiruanLepas").value = "Y" Then
            varStatusAksi = GIGI_TIRUAN_LEPAS
            Call picTengah_MouseUp(idx, 1, 0, 10, 10)
        End If
        rs.MoveNext
    Wend
    varStatusAksi = NORMAL_A
End Sub

Private Function subCaptureDesktop() As Boolean
    Dim hWndScreen As Long

    With Me.picDiagramOdondo
        Set Me.picTemp.Picture = CaptureWindow(.hWnd, True, 40, 0, _
        .ScaleX(Me.picTemp.ScaleWidth, .ScaleMode, vbPixels), _
        .ScaleY(Me.picTemp.ScaleHeight, .ScaleMode, vbPixels))
    End With

    SavePicture Me.picTemp.Image, App.path & "\tempbitmap.bmp"
    Set Me.picTemp.Picture = Nothing
End Function

Private Sub subNonVital(Index As Integer)
    Me.picTengah(Index).DrawWidth = 3
    Me.picTengah(Index).Line (0, Me.picTengah(Index).ScaleHeight)-(Me.picTengah(Index).ScaleWidth / 3, Me.picTengah(Index).ScaleHeight), vbRed
    Me.picTengah(Index).Line (Me.picTengah(Index).ScaleWidth / 3, Me.picTengah(Index).ScaleHeight)-((Me.picTengah(Index).ScaleWidth / 3) * 2, 0), vbRed
    Me.picTengah(Index).Line ((Me.picTengah(Index).ScaleWidth / 3) * 2, 0)-(Me.picTengah(Index).ScaleWidth, 0), vbRed
End Sub

Private Sub subSisaAkar(Index As Integer)
    Me.picTengah(Index).DrawWidth = 2
    Me.picTengah(Index).Line (Me.picTengah(Index).ScaleWidth / 2, 0)-(Me.picTengah(Index).ScaleWidth / 2, Me.picTengah(Index).ScaleHeight - 50), vbBlue
    Me.picTengah(Index).Line (0, Me.picTengah(Index).ScaleHeight - 50)-(Me.picTengah(Index).ScaleWidth, Me.picTengah(Index).ScaleHeight - 50), vbBlue
End Sub

Private Sub subGigiHilang(Index)
    Me.picTengah(Index).DrawWidth = 3
    Me.picTengah(Index).Line (0, 0)-(Me.picTengah(Index).ScaleWidth, Me.picTengah(Index).ScaleHeight), vbRed
    Me.picTengah(Index).Line (0, Me.picTengah(Index).ScaleHeight)-(Me.picTengah(Index).ScaleWidth, 0), vbRed
End Sub

Private Sub subJembatan(Index)
    Me.picTengah(Index).DrawWidth = 3
    Me.picTengah(Index).Line (0, 75)-(Me.picTengah(Index).ScaleWidth, 75), Me.picMLogam.BackColor
End Sub

Private Sub subGigiTiruanLepas(Index)
    Me.picTengah(Index).DrawWidth = 3
    Me.picTengah(Index).Line (0, 150)-(Me.picTengah(Index).ScaleWidth, 150), vbYellow
End Sub

Private Sub subRefreshGigiTengah(Index As Integer)
    With varKondisiGigi(Index)
        If .NonVital = "Y" Then Call subNonVital(Index)
        If .SisaAkar = "Y" Then Call subSisaAkar(Index)
        If .GigiHilang = "Y" Then Call subGigiHilang(Index)
        If .Jembatan = "Y" Then Call subJembatan(Index)
        If .GigiTiruanLepas = "Y" Then Call subGigiTiruanLepas(Index)
    End With
End Sub

Private Sub chkShowHideGigiHilang_Click()
    Dim i As Integer

    If Me.chkShowHideGigiHilang.value = 0 Then
        For i = 11 To 85
            If varKondisiGigi(i).AdaGigi Then
                If varKondisiGigi(i).GigiHilang = "Y" Then
                    Me.picGigi(i).Visible = True
                End If
            End If
        Next
    Else
        For i = 11 To 85
            If varKondisiGigi(i).AdaGigi Then
                If varKondisiGigi(i).GigiHilang = "Y" Then
                    Me.picGigi(i).Visible = False
                End If
            End If
        Next
    End If
End Sub

Private Sub cmdCetakOdonto_Click()
    Call subCaptureDesktop
    frmCetakOdontoGram.Show
End Sub

Private Sub cmdSimpan_Click()
    Dim i As Integer
    Dim blnBerhasil As Boolean

    If Not Add_CatatanOdonto Then Exit Sub
    For i = 11 To 85
        With varKondisiGigi(i)
            If .AdaGigi Then
                If Add_DetailCatatanOdonto(i) Then
                    blnBerhasil = True
                Else
                    blnBerhasil = False
                    Exit For
                End If
            End If
        End With
    Next
    If blnBerhasil Then
        MsgBox "Penyimpanan data berhasil!", vbInformation
        Me.cmdCetakOdonto.Enabled = True
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            Call subCaptureDesktop
    End Select
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)

    Call subResetArray
    Me.dtpTglPeriksa.value = Now

    Me.optAksi(0).value = True
    Me.picGigiHilang.Line (0, 0)-(Me.picGigiHilang.ScaleWidth, Me.picGigiHilang.ScaleHeight), vbRed
    Me.picGigiHilang.Line (0, Me.picGigiHilang.ScaleHeight)-(Me.picGigiHilang.ScaleWidth, 0), vbRed

    Me.picJembatan.Line (0, 75)-(Me.picJembatan.ScaleWidth, 75), Me.picMLogam.BackColor
    Me.picGigiTiruanLepas.Line (0, 150)-(Me.picJembatan.ScaleWidth, 150), vbYellow
    Me.picSisaAkar.Line (Me.picSisaAkar.ScaleWidth / 2, 0)-(Me.picSisaAkar.ScaleWidth / 2, Me.picSisaAkar.ScaleHeight - 50), vbBlue
    Me.picSisaAkar.Line (0, Me.picSisaAkar.ScaleHeight - 50)-(Me.picSisaAkar.ScaleWidth, Me.picSisaAkar.ScaleHeight - 50), vbBlue
    Me.picNonVital.Line (0, Me.picNonVital.ScaleHeight)-(Me.picNonVital.ScaleWidth / 3, Me.picNonVital.ScaleHeight), vbRed
    Me.picNonVital.Line (Me.picNonVital.ScaleWidth / 3, Me.picNonVital.ScaleHeight)-((Me.picNonVital.ScaleWidth / 3) * 2, 0), vbRed
    Me.picNonVital.Line ((Me.picNonVital.ScaleWidth / 3) * 2, 0)-(Me.picNonVital.ScaleWidth, 0), vbRed

    Dim i As Integer
    For i = 0 To 51
        Me.lblNoGigi(i).Left = Me.lblNoGigi(i).Left + 20
    Next

    Frame2.Height = 2175
    Frame4.Top = 8400
    Me.Height = 9690

End Sub

Private Sub lblGigi_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    Call picTengah_MouseUp(Index, Button, Shift, X, y)
End Sub

Private Sub lblGigiAnomali_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    Call picTengah_MouseUp(Index, Button, Shift, X, y)
End Sub

Private Sub optAksi_Click(Index As Integer)
    If Me.optAksi(Index).value = True Then
        varStatusAksi = Index
    End If
End Sub

Private Sub picGigi_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    Select Case varStatusAksi
        Case BELUM_ERUPSI
            Me.lblGigi(Index).Font.Size = 8 'direset
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            If varKondisiGigi(Index).BelumErupsi = "Y" Then
                Me.lblGigi(Index).Caption = ""
                varKondisiGigi(Index).BelumErupsi = "T"
            Else
                With Me.lblGigi(Index)
                    .Caption = Me.lblBelumErupsi.Caption
                    .ForeColor = Me.lblBelumErupsi.ForeColor
                End With
                varKondisiGigi(Index).BelumErupsi = "Y"
            End If
        Case ERUPSI_SEBAGIAN
            Me.lblGigi(Index).Font.Size = 8 'direset
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            If varKondisiGigi(Index).ErupsiSebagian = "Y" Then
                Me.lblGigi(Index).Caption = ""
                varKondisiGigi(Index).ErupsiSebagian = "T"
            Else
                With Me.lblGigi(Index)
                    .Caption = Me.lblErupsiSebagian.Caption
                    .ForeColor = Me.lblErupsiSebagian.ForeColor
                End With
                varKondisiGigi(Index).ErupsiSebagian = "Y"
            End If
        Case ANOMALI_BENTUK
            Me.lblGigi(Index).Font.Size = 8 'direset
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            If varKondisiGigi(Index).AnomaliBentuk = "Y" Then
                Me.lblGigiAnomali(Index).Caption = ""
                varKondisiGigi(Index).AnomaliBentuk = "T"
            Else
                With Me.lblGigiAnomali(Index)
                    .Caption = Me.lblAnomaliBentuk.Caption
                    .ForeColor = Me.lblAnomaliBentuk.ForeColor
                End With
                varKondisiGigi(Index).AnomaliBentuk = "Y"
            End If
        Case Calculus
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            If varKondisiGigi(Index).Calculus = "Y" Then
                Me.lblGigiAnomali(Index).Caption = ""
                varKondisiGigi(Index).Calculus = "T"
            Else
                With Me.lblGigiAnomali(Index)
                    .Caption = Me.lblCalculus.Caption
                    .Font.Size = 18 'dibesarkan
                    .ForeColor = Me.lblCalculus.ForeColor
                End With
                varKondisiGigi(Index).Calculus = "Y"
            End If
        Case KARIES
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            Call subWarnaiBagianGigi(Index, Button, Shift, X, y, Me.picKaries.BackColor)
        Case NON_VITAL
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            Call picTengah_MouseUp(Index, Button, Shift, X, y)
        Case TAMBALAN_LOGAM
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            Call subWarnaiBagianGigi(Index, Button, Shift, X, y, Me.picTLogam.BackColor)
        Case TAMBALAN_NON_LOGAM
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            Call subWarnaiBagianGigi(Index, Button, Shift, X, y, Me.picTNonLogam.BackColor)
        Case MAHKOTA_LOGAM
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            Call subWarnaiBagianGigi(Index, Button, Shift, X, y, Me.picMLogam.BackColor)
        Case MAHKOTA_NON_LOGAM
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            Call subWarnaiBagianGigi(Index, Button, Shift, X, y, Me.picMNonLogam.BackColor)
        Case SISA_AKAR
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            Call picTengah_MouseUp(Index, Button, Shift, X, y)
        Case GIGI_HILANG
            Call picTengah_MouseUp(Index, Button, Shift, X, y)
        Case JEMBATAN_A
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            Call picTengah_MouseUp(Index, Button, Shift, X, y)
        Case GIGI_TIRUAN_LEPAS
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            Call picTengah_MouseUp(Index, Button, Shift, X, y)
    End Select
    varKondisiGigi(Index).AdaGigi = True
    Exit Sub
jump:
    MsgBox "Gigi sudah hilang!", vbInformation
    varKondisiGigi(Index).AdaGigi = True
End Sub

Private Sub picTengah_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    Select Case varStatusAksi
        Case BELUM_ERUPSI
            Me.lblGigi(Index).Font.Size = 8 'direset
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            If varKondisiGigi(Index).BelumErupsi = "Y" Then
                Me.lblGigi(Index).Caption = ""
                varKondisiGigi(Index).BelumErupsi = "T"
            Else
                With Me.lblGigi(Index)
                    .Caption = Me.lblBelumErupsi.Caption
                    .ForeColor = Me.lblBelumErupsi.ForeColor
                End With
                varKondisiGigi(Index).BelumErupsi = "Y"
            End If
        Case ERUPSI_SEBAGIAN
            Me.lblGigi(Index).Font.Size = 8 'direset
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            If varKondisiGigi(Index).ErupsiSebagian = "Y" Then
                Me.lblGigi(Index).Caption = ""
                varKondisiGigi(Index).ErupsiSebagian = "T"
            Else
                With Me.lblGigi(Index)
                    .Caption = Me.lblErupsiSebagian.Caption
                    .ForeColor = Me.lblErupsiSebagian.ForeColor
                End With
                varKondisiGigi(Index).ErupsiSebagian = "Y"
            End If
        Case ANOMALI_BENTUK
            Me.lblGigi(Index).Font.Size = 8 'direset
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            If varKondisiGigi(Index).AnomaliBentuk = "Y" Then
                Me.lblGigiAnomali(Index).Caption = ""
                varKondisiGigi(Index).AnomaliBentuk = "T"
            Else
                With Me.lblGigiAnomali(Index)
                    .Caption = Me.lblAnomaliBentuk.Caption
                    .ForeColor = Me.lblAnomaliBentuk.ForeColor
                End With
                varKondisiGigi(Index).AnomaliBentuk = "Y"
            End If
        Case KARIES
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            Call subSetBagianDepan(Index, Button, Shift, X, y, Me.picKaries.BackColor)
        Case NON_VITAL
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            If varKondisiGigi(Index).NonVital = "Y" Then
                Me.picTengah(Index).Cls
                With varKondisiGigi(Index)
                    .NonVital = "T"
                    If .SisaAkar = "Y" Then Call subSisaAkar(Index)
                    If .GigiHilang = "Y" Then Call subGigiHilang(Index)
                    If .Jembatan = "Y" Then Call subJembatan(Index)
                    If .GigiTiruanLepas = "Y" Then Call subGigiTiruanLepas(Index)
                End With
            Else
                Call subNonVital(Index)
                With varKondisiGigi(Index)
                    .NonVital = "Y"
                    If .SisaAkar = "Y" Then
                    End If
                End With
            End If
        Case TAMBALAN_LOGAM
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            Call subSetBagianDepan(Index, Button, Shift, X, y, Me.picTLogam.BackColor)
        Case TAMBALAN_NON_LOGAM
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            Call subSetBagianDepan(Index, Button, Shift, X, y, Me.picTNonLogam.BackColor)
        Case MAHKOTA_LOGAM
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            Call subSetBagianDepan(Index, Button, Shift, X, y, Me.picMLogam.BackColor)
        Case MAHKOTA_NON_LOGAM
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            Call subSetBagianDepan(Index, Button, Shift, X, y, Me.picMNonLogam.BackColor)
        Case SISA_AKAR
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            If varKondisiGigi(Index).SisaAkar = "Y" Then
                Me.picTengah(Index).Cls
                With varKondisiGigi(Index)
                    .SisaAkar = "T"
                    If .NonVital = "Y" Then Call subNonVital(Index)
                    If .GigiHilang = "Y" Then Call subGigiHilang(Index)
                    If .Jembatan = "Y" Then Call subJembatan(Index)
                    If .GigiTiruanLepas = "Y" Then Call subGigiTiruanLepas(Index)
                End With
            Else
                Call subSisaAkar(Index)
                varKondisiGigi(Index).SisaAkar = "Y"
            End If
        Case GIGI_HILANG
            If varKondisiGigi(Index).GigiHilang = "Y" Then
                Me.picTengah(Index).Cls
                With varKondisiGigi(Index)
                    .GigiHilang = "T"
                    If .NonVital = "Y" Then Call subNonVital(Index)
                    If .SisaAkar = "Y" Then Call subSisaAkar(Index)
                    If .Jembatan = "Y" Then Call subJembatan(Index)
                    If .GigiTiruanLepas = "Y" Then Call subGigiTiruanLepas(Index)
                End With
            Else
                Call subGigiHilang(Index)
                varKondisiGigi(Index).GigiHilang = "Y"
                If Me.chkShowHideGigiHilang.value = 1 Then
                    Me.picGigi(Index).Visible = False
                End If
            End If
        Case JEMBATAN_A
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            If varKondisiGigi(Index).Jembatan = "Y" Then
                Me.picTengah(Index).Cls
                With varKondisiGigi(Index)
                    .Jembatan = "T"
                    If .NonVital = "Y" Then Call subNonVital(Index)
                    If .SisaAkar = "Y" Then Call subSisaAkar(Index)
                    If .GigiHilang = "Y" Then Call subGigiHilang(Index)
                    If .GigiTiruanLepas = "Y" Then Call subGigiTiruanLepas(Index)
                End With
            Else
                Call subJembatan(Index)
                varKondisiGigi(Index).Jembatan = "Y"
            End If
        Case GIGI_TIRUAN_LEPAS
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            If varKondisiGigi(Index).GigiTiruanLepas = "Y" Then
                Me.picTengah(Index).Cls
                With varKondisiGigi(Index)
                    .GigiTiruanLepas = "T"
                    If .NonVital = "Y" Then Call subNonVital(Index)
                    If .SisaAkar = "Y" Then Call subSisaAkar(Index)
                    If .GigiHilang = "Y" Then Call subGigiHilang(Index)
                    If .Jembatan = "Y" Then Call subJembatan(Index)
                End With
            Else
                Call subGigiTiruanLepas(Index)
                varKondisiGigi(Index).GigiTiruanLepas = "Y"
            End If
        Case Calculus
            If varKondisiGigi(Index).GigiHilang = "Y" Then GoTo jump
            If varKondisiGigi(Index).Calculus = "Y" Then
                Me.lblGigi(Index).Caption = ""
                Me.lblGigi(Index).Font.Size = 8 'direset
                varKondisiGigi(Index).Calculus = "T"
            Else
                With Me.lblGigi(Index)
                    .Font.Size = 18 'dibesarkan
                    .Caption = Me.lblCalculus.Caption
                    .ForeColor = Me.lblCalculus.ForeColor
                End With
                varKondisiGigi(Index).Calculus = "Y"
            End If
    End Select
    varKondisiGigi(Index).AdaGigi = True
    Exit Sub
jump:
    MsgBox "Gigi sudah hilang!", vbInformation
    varKondisiGigi(Index).AdaGigi = True
End Sub

