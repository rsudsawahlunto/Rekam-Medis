VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCatatanKlinisPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Catatan Klinis Pasien"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCatatanKlinisPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   9495
   Begin VB.Frame fraDokter 
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
      Height          =   2895
      Left            =   120
      TabIndex        =   49
      Top             =   120
      Visible         =   0   'False
      Width           =   9135
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   2295
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   8535
         _ExtentX        =   15055
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
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   36
      Top             =   5520
      Width           =   9495
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   465
         Left            =   5400
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   465
         Left            =   7440
         TabIndex        =   23
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Klinis Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   0
      TabIndex        =   24
      Top             =   2040
      Width           =   9495
      Begin VB.TextBox txtGCSTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   21
         Text            =   "0"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox txtGCSM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3120
         MaxLength       =   3
         TabIndex        =   20
         Text            =   "0"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox txtGCSF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "0"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox txtGCSE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "0"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox txtPemeriksa 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2280
         MaxLength       =   150
         TabIndex        =   8
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox txtNadi 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1680
         MaxLength       =   150
         TabIndex        =   13
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtSuhu 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6240
         MaxLength       =   150
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtPernafasan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1680
         MaxLength       =   150
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtTekananDarah 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1680
         MaxLength       =   150
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1680
         MaxLength       =   150
         TabIndex        =   17
         Top             =   2400
         Width           =   7575
      End
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   118751235
         UpDown          =   -1  'True
         CurrentDate     =   38076
      End
      Begin MSDataListLib.DataCombo dcKesadaran 
         Height          =   330
         Left            =   6240
         TabIndex        =   16
         Top             =   2040
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin MSMask.MaskEdBox meBeratTingi 
         Height          =   330
         Left            =   6240
         TabIndex        =   15
         Top             =   1680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   7
         Mask            =   "###/###"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo dcPerawat 
         Height          =   330
         Left            =   6120
         TabIndex        =   10
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   210
         Index           =   27
         Left            =   3840
         TabIndex        =   56
         Top             =   2760
         Width           =   420
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "M"
         Height          =   210
         Index           =   26
         Left            =   3360
         TabIndex        =   55
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "F"
         Height          =   210
         Index           =   25
         Left            =   2640
         TabIndex        =   54
         Top             =   2760
         Width           =   90
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "E"
         Height          =   210
         Index           =   24
         Left            =   1920
         TabIndex        =   53
         Top             =   2760
         Width           =   105
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "GCS"
         Height          =   210
         Index           =   23
         Left            =   1170
         TabIndex        =   52
         Top             =   3000
         Width           =   330
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Paramedis Pemeriksa"
         Height          =   210
         Index           =   22
         Left            =   6120
         TabIndex        =   51
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "/Menit"
         Height          =   210
         Index           =   21
         Left            =   3120
         TabIndex        =   50
         Top             =   2115
         Width           =   525
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   9360
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Kg/Cm"
         Height          =   210
         Index           =   20
         Left            =   7185
         TabIndex        =   48
         Top             =   1740
         Width           =   540
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "C"
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
         Index           =   19
         Left            =   7800
         TabIndex        =   47
         Top             =   1440
         Width           =   120
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "o"
         Height          =   210
         Index           =   18
         Left            =   7650
         TabIndex        =   46
         Top             =   1320
         Width           =   105
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "/Menit"
         Height          =   210
         Index           =   17
         Left            =   3120
         TabIndex        =   45
         Top             =   1740
         Width           =   525
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "MmHg"
         Height          =   210
         Index           =   16
         Left            =   3120
         TabIndex        =   44
         Top             =   1395
         Width           =   510
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Kesadaran"
         Height          =   210
         Index           =   15
         Left            =   5280
         TabIndex        =   43
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Berat/Tinggi Badan"
         Height          =   210
         Index           =   14
         Left            =   4560
         TabIndex        =   42
         Top             =   1680
         Width           =   1560
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Suhu"
         Height          =   210
         Index           =   13
         Left            =   5640
         TabIndex        =   41
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Pernafasan"
         Height          =   210
         Index           =   12
         Left            =   615
         TabIndex        =   40
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Dokter/Perawat Pemeriksa"
         Height          =   210
         Index           =   11
         Left            =   2280
         TabIndex        =   39
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Periksa"
         Height          =   210
         Index           =   10
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nadi"
         Height          =   210
         Index           =   8
         Left            =   1155
         TabIndex        =   37
         Top             =   2040
         Width           =   345
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   210
         Index           =   9
         Left            =   555
         TabIndex        =   35
         Top             =   2400
         Width           =   945
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Tekanan Darah"
         Height          =   210
         Index           =   7
         Left            =   270
         TabIndex        =   34
         Top             =   1320
         Width           =   1230
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
      TabIndex        =   25
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
         TabIndex        =   26
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
            TabIndex        =   29
            Top             =   277
            Width           =   285
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Index           =   5
            Left            =   1350
            TabIndex        =   28
            Top             =   277
            Width           =   240
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Index           =   6
            Left            =   2130
            TabIndex        =   27
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
         TabIndex        =   33
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Index           =   2
         Left            =   3000
         TabIndex        =   32
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Index           =   1
         Left            =   1800
         TabIndex        =   31
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   1335
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   57
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
      Picture         =   "frmCatatanKlinisPasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmCatatanKlinisPasien.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmCatatanKlinisPasien.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmCatatanKlinisPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilterDokter As String

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    If Periksa("text", txtPemeriksa, "Nama pemeriksa kosong") = False Then Exit Sub
    If Periksa("text", txtTekananDarah, "Tekanan Darah kosong") = False Then Exit Sub
    If Periksa("text", txtNadi, "Nadi kosong") = False Then Exit Sub
    If Periksa("text", txtPernafasan, "Pernapasan kosong") = False Then Exit Sub
    If Periksa("text", txtSuhu, "Suhu kosong") = False Then Exit Sub
    If Len(Trim(dcPerawat.Text)) > 0 Then If Periksa("datacombo", dcPerawat, "Nama perawat kosong") = False Then Exit Sub
    If Len(Trim(dcKesadaran.Text)) = 0 Then If Periksa("datacombo", dcKesadaran, "Kesadaran Masih Kosong dan harus di isi") = False Then Exit Sub

    If mstrKdDokter = "" Then
        MsgBox "Pilih dulu Pemeriksa yang akan menangani Pasien", vbExclamation, "Validasi"
        txtPemeriksa.SetFocus
        Exit Sub
    End If

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dtpTglPeriksa.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("IdPegawai", adVarChar, adParamInput, 10, mstrKdDokter)

        .Parameters.Append .CreateParameter("TekananDarah", adVarChar, adParamInput, 20, IIf(Len(Trim(txtTekananDarah.Text)) = 0, Null, Trim(txtTekananDarah.Text)))
        .Parameters.Append .CreateParameter("Pernafasan", adVarChar, adParamInput, 20, IIf(Len(Trim(txtPernafasan.Text)) = 0, Null, Trim(txtPernafasan.Text)))
        .Parameters.Append .CreateParameter("Nadi", adVarChar, adParamInput, 20, IIf(txtNadi.Text = "", Null, txtNadi.Text))
        .Parameters.Append .CreateParameter("Suhu", adVarChar, adParamInput, 20, IIf(Len(Trim(txtSuhu.Text)) = 0, Null, Trim(txtSuhu.Text)))
        .Parameters.Append .CreateParameter("BeratTinggiBadan", adVarChar, adParamInput, 20, IIf(meBeratTingi.Text = "__/___", Null, meBeratTingi.Text))
        .Parameters.Append .CreateParameter("KdKesadaran", adChar, adParamInput, 2, IIf(dcKesadaran.BoundText = "", Null, dcKesadaran.BoundText))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 1000, IIf(Len(Trim(txtKeterangan.Text)) = 0, Null, Trim(txtKeterangan.Text)))

        .Parameters.Append .CreateParameter("GCSE", adTinyInt, adParamInput, , Val(txtGCSE.Text))
        .Parameters.Append .CreateParameter("GCSF", adTinyInt, adParamInput, , Val(txtGCSF.Text))
        .Parameters.Append .CreateParameter("GCSM", adTinyInt, adParamInput, , Val(txtGCSM.Text))

        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("IdParamedis", adChar, adParamInput, 10, IIf(dcPerawat.BoundText = "", Null, dcPerawat.BoundText))

        .ActiveConnection = dbConn
        .CommandText = "dbo.AU_CatatanKlinisPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
        Else
            MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
            Call Add_HistoryLoginActivity("AU_CatatanKlinisPasien")
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    cmdSimpan.Enabled = False
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    If cmdSimpan.Enabled = True Then
        If txtNoPendaftaran.Text <> "" Then
            If MsgBox("Simpan catatan klinis ", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
                Call cmdSimpan_Click
                Exit Sub
            End If
        End If
    End If
    Unload Me
End Sub

Private Sub dcKesadaran_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub dcKesadaran_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKesadaran.MatchedWithList = True Then txtKeterangan.SetFocus
        strSQL = " SELECT KdKesadaran, NamaKesadaran" & _
        " From Kesadaran where StatusEnabled='1'" & _
        " and (NamaKesadaran LIKE '%" & dcKesadaran.Text & "%')ORDER BY NamaKesadaran"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcKesadaran.Text = ""
            Exit Sub
        End If
        dcKesadaran.BoundText = rs(0).value
        dcKesadaran.Text = rs(1).value
    End If
End Sub

Private Sub dcPerawat_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    If KeyAscii = 13 Then
        If Len(Trim(dcPerawat.Text)) > 0 Then
            strSQL = "SELECT  IdPegawai, [Nama Pemeriksa]" & _
            " From V_DaftarPemeriksaPasien" & _
            " WHERE ([Nama Pemeriksa] LIKE '%" & dcPerawat.Text & "%')"
            Call msubRecFO(rs, strSQL)
            dcPerawat.Text = ""
            If rs.EOF = False Then dcPerawat.BoundText = rs(0).value: txtTekananDarah.SetFocus
        Else
            txtTekananDarah.SetFocus
        End If
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dgDokter.ApproxCount = 0 Then Exit Sub
        txtPemeriksa.Text = dgDokter.Columns(1).value
        mstrKdDokter = dgDokter.Columns(0).value
        If mstrKdDokter = "" Then
            MsgBox "Pilih dulu Pemeriksa yang akan menangani Pasien", vbCritical, "Validasi"
            txtPemeriksa.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        fraDokter.Visible = False
        dcPerawat.SetFocus
    End If
    If KeyAscii = 27 Then
        fraDokter.Visible = False
    End If
End Sub

Private Sub meBeratTingi_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtGCSE_Change()
    Call subHitungGCS
End Sub

Private Sub txtGCSE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtGCSF.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtGCSE_LostFocus()
    txtGCSE.Text = Val(txtGCSE.Text)
End Sub

Private Sub txtGCSF_Change()
    Call subHitungGCS
End Sub

Private Sub txtGCSF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtGCSM.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtGCSF_LostFocus()
    txtGCSF.Text = Val(txtGCSF.Text)
End Sub

Private Sub txtGCSM_Change()
    Call subHitungGCS
End Sub

Private Sub txtGCSM_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    If KeyAscii = 13 Then If cmdSimpan.Enabled = True Then cmdSimpan.SetFocus Else cmdTutup.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
    Exit Sub
errLoad:
End Sub

Private Sub txtGCSM_LostFocus()
    txtGCSM.Text = Val(txtGCSM.Text)
End Sub

Private Sub txtGCSTotal_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then If cmdSimpan.Enabled = False Then cmdTutup.SetFocus Else cmdSimpan.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtNadi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtSuhu.SetFocus
End Sub

Private Sub txtNadi_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtPemeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errLoad
    Select Case KeyCode
        Case 13
            If fraDokter.Visible = True Then
                dgDokter.SetFocus
            Else
                dcPerawat.SetFocus
            End If
        Case vbKeyEscape
            fraDokter.Visible = False
    End Select
    Exit Sub
    If KeyCode = vbKeyDown Then
        If fraDokter.Visible = False Then Exit Sub
        dgDokter.SetFocus
    End If
    Call SetKeyPressToChar(KeyCode)
errLoad:
    Call msubPesanError
End Sub

Private Sub dtpTglPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtPemeriksa.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglPeriksa.value = Now
    txtPemeriksa.Text = ""
    txtTekananDarah.Text = ""
    txtPernafasan.Text = ""
    txtNadi.Text = ""
    txtSuhu.Text = ""
    meBeratTingi.Text = "___/___"
    dcKesadaran.BoundText = ""
    txtKeterangan.Text = ""
    dcPerawat.BoundText = ""

    Call subLoadDcSource

    strSQL = "SELECT  IdPegawai, [Nama Pemeriksa]" & _
    " FROM V_DaftarPemeriksaPasien " & _
    " WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
    Call msubRecFO(dbRst, strSQL)
    If rs.EOF = False Then
        txtPemeriksa.Text = dbRst(1).value
        mstrKdDokter = dbRst(0).value
    Else
        mstrKdDokter = ""
        txtPemeriksa.Text = ""
    End If
    fraDokter.Visible = False

    With frmTransaksiPasien
        txtNoPendaftaran = .txtNoPendaftaran.Text
        txtNoCM = .txtNoCM.Text
        txtNamaPasien = .txtNamaPasien.Text
        txtSex.Text = .txtSex.Text
        txtThn = .txtThn.Text
        txtBln = .txtBln.Text
        txtHari = .txtHr.Text
    End With

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmTransaksiPasien.Enabled = True
    Call frmTransaksiPasien.subLoadRiwayatCatatanKlinis
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    strSQL = "SELECT KdKesadaran, NamaKesadaran" & _
    " From Kesadaran where StatusEnabled='1'" & _
    " ORDER BY NamaKesadaran"
    Call msubDcSource(dcKesadaran, rs, strSQL)
    If rs.EOF = False Then dcKesadaran.BoundText = rs(0).value

    strSQL = "SELECT  IdPegawai, [Nama Pemeriksa]" & _
    " From V_DaftarPemeriksaPasien" & _
    " ORDER BY  [Nama Pemeriksa]"
    Call msubDcSource(dcPerawat, rs, strSQL)
    If rs.EOF = False Then dcPerawat.BoundText = strIDPegawaiAktif

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDokter()
    On Error GoTo errLoad

    strSQL = "SELECT IdPegawai AS [Kode Pemeriksa], [Nama Pemeriksa],JK,[Jenis Pemeriksa] " & _
    " FROM V_DaftarDokterdanPemeriksaPasien " & strFilterDokter
    Call msubRecFO(rs, strSQL)
    Set dgDokter.DataSource = rs
    With dgDokter
        .Columns(0).Width = 1500
        .Columns(1).Width = 4000
        .Columns(2).Width = 400
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 2000
    End With
    fraDokter.Left = 240
    fraDokter.Top = 3000
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub meBeratTingi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcKesadaran.SetFocus
End Sub

Private Sub txtKeterangan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtGCSE.SetFocus
End Sub

Private Sub txtPemeriksa_Change()
    strFilterDokter = "WHERE [Nama Pemeriksa] like '%" & txtPemeriksa.Text & "%'"
    mstrKdDokter = ""
    fraDokter.Visible = True
    Call subLoadDokter
End Sub

Private Sub txtPernafasan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtNadi.SetFocus
End Sub

Private Sub txtPernafasan_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtSuhu_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then meBeratTingi.SetFocus
End Sub

Private Sub txtSuhu_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtTekananDarah_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtPernafasan.SetFocus
End Sub

Private Sub subHitungGCS()
    On Error GoTo errLoad

    txtGCSTotal.Text = Val(txtGCSE.Text) + Val(txtGCSF.Text) + Val(txtGCSM.Text)

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtTekananDarah_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

