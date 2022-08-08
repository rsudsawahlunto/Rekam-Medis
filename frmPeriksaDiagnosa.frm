VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPeriksaDiagnosa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pencatatan Diagnosa Pasien "
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPeriksaDiagnosa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   11475
   Begin VB.TextBox txtNamaFormPengirim 
      Height          =   375
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkICD9 
      Caption         =   "Data Diagnosa Tindakan Pasien [ICD 9]"
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
      Left            =   120
      TabIndex        =   19
      Top             =   5520
      Width           =   3735
   End
   Begin VB.Frame fraICD9 
      Caption         =   "Diagnosa Tindakan Pasien [ICD 9]"
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
      Height          =   2775
      Left            =   0
      TabIndex        =   39
      Top             =   5520
      Width           =   11415
      Begin VB.TextBox txtKodeICD9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtDiagnosaTindakan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   6375
      End
      Begin MSComctlLib.ListView lvwDiagnosaTindakan 
         Height          =   1815
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nama Diagnosa"
            Object.Width           =   13229
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcJenisDiagnosaTindakan 
         Height          =   330
         Left            =   6600
         TabIndex        =   21
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Diagnosa Tindakan"
         Height          =   210
         Left            =   6600
         TabIndex        =   41
         Top             =   240
         Width           =   1965
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nama Diagnosa Tindakan"
         Height          =   210
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   2025
      End
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
      Height          =   1815
      Left            =   1800
      TabIndex        =   36
      Top             =   3840
      Visible         =   0   'False
      Width           =   8775
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   1335
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2355
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
      Height          =   735
      Left            =   0
      TabIndex        =   37
      Top             =   8160
      Width           =   11415
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   7200
         TabIndex        =   25
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   9360
         TabIndex        =   26
         Top             =   240
         Width           =   1935
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
      TabIndex        =   27
      Top             =   960
      Width           =   11415
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7440
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3240
         TabIndex        =   2
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   1455
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
         Left            =   8880
         TabIndex        =   28
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   29
            Top             =   270
            Width           =   150
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   7440
         TabIndex        =   35
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   3240
         TabIndex        =   34
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1680
         TabIndex        =   33
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   38
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
      Caption         =   "Data Diagnosa Pasien [ICD 10]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   0
      TabIndex        =   43
      Top             =   1920
      Width           =   11415
      Begin VB.TextBox txtNamaDiagnosa 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2280
         TabIndex        =   8
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txtKodeICD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   2160
      End
      Begin MSDataListLib.DataCombo dcJenisDiagnosa 
         Height          =   330
         Left            =   6120
         TabIndex        =   10
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   124977155
         UpDown          =   -1  'True
         CurrentDate     =   38077
      End
      Begin MSComctlLib.ListView lvwDiagnosa 
         Height          =   1335
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   2355
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nama Diagnosa"
            Object.Width           =   13229
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcPenyebabDiagnosa 
         Height          =   330
         Left            =   5400
         TabIndex        =   13
         Top             =   1080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcInfeksiNosokomial 
         Height          =   330
         Left            =   8400
         TabIndex        =   14
         Top             =   1080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcPenyebabInfeksiNosokomial 
         Height          =   330
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcKetunaanKelainan 
         Height          =   330
         Left            =   6240
         TabIndex        =   17
         Top             =   1680
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcMorfologiNeoplasma 
         Height          =   330
         Left            =   3240
         TabIndex        =   16
         Top             =   1680
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pemeriksaan"
         Height          =   210
         Left            =   240
         TabIndex        =   52
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Diagnosa"
         Height          =   210
         Left            =   6120
         TabIndex        =   51
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nama Diagnosa"
         Height          =   210
         Left            =   240
         TabIndex        =   50
         Top             =   840
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Pemeriksa"
         Height          =   210
         Left            =   2280
         TabIndex        =   49
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Penyebab Diagnosa"
         Height          =   210
         Left            =   5400
         TabIndex        =   48
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Infeksi Nosokomial"
         Height          =   210
         Left            =   8400
         TabIndex        =   47
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Penyebab Infeksi Nosokomial"
         Height          =   210
         Left            =   240
         TabIndex        =   46
         Top             =   1440
         Width           =   2355
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Morfologi Neoplasma"
         Height          =   210
         Left            =   3240
         TabIndex        =   45
         Top             =   1440
         Width           =   1665
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Ketunaan Kelainan"
         Height          =   210
         Left            =   6240
         TabIndex        =   44
         Top             =   1440
         Width           =   1500
      End
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9600
      Picture         =   "frmPeriksaDiagnosa.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPeriksaDiagnosa.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPeriksaDiagnosa.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "frmPeriksaDiagnosa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilterDokter As String
Dim intJmlDiagDipilih, intICD9Diplh, intICD9DiplhBaru As Integer
Dim strKdDiagnosa() As String
Dim strKdDiagnosaTindakan() As String
Dim bolvalICD10 As Boolean
Dim i As Integer
Dim itemAll As Object
Dim j As Integer

Private Sub chkICD9_Click()
    If chkICD9.value = vbChecked Then
        fraICD9.Enabled = True
    Else
        fraICD9.Enabled = False
    End If
End Sub

Private Sub chkICD9_KeyPress(KeyAscii As Integer)
    If chkICD9.value = vbChecked Then
        fraICD9.Enabled = True
        txtDiagnosaTindakan.SetFocus
    Else
        fraICD9.Enabled = False
        chkICD9.SetFocus
    End If
End Sub

Private Sub cmdCetak_Click()
    frm_cetak_info_diag_viewer.Show
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If dcJenisDiagnosa.Text = "" Then
        MsgBox "Jenis Diagnosa Belum Diisi ", vbCritical, "Validasi"
        dcJenisDiagnosa.SetFocus
        Exit Sub
    ElseIf mstrKdDokter = "" Then
        MsgBox "Pilih dulu Dokter yang akan menangani Pasien ", vbCritical, "Validasi"
        If txtDokter.Enabled = False Then Exit Sub
        txtDokter.SetFocus
        Exit Sub
    ElseIf intJmlDiagDipilih = 0 Then
        MsgBox "Diagnosa Belum Dipilih ", vbCritical, "Validasi"
        txtNamaDiagnosa.SetFocus
        Exit Sub
    End If

    If lvwDiagnosa.Enabled = True Then
        strSQL = "Select * from PeriksaDiagnosa where NoPendaftaran='" & mstrNoPen & "' and KdJenisDiagnosa='" & dcJenisDiagnosa.BoundText & "'"
        Set dbRst2 = Nothing
        Call msubRecFO(dbRst2, strSQL)
        If dbRst2.EOF = False Then
            MsgBox "Jenis Diagnosa utama tidak boleh lebih dari satu ", vbExclamation, "Validasi"
            Exit Sub
        End If

        If bolvalICD10 = False Then
            For i = 1 To intJmlDiagDipilih
                strSQL = ""
                strSQL = "Select KdDiagnosa From PeriksaDiagnosa Where NoPendaftaran = '" & txtNoPendaftaran.Text & "' And KdDiagnosa = '" & strKdDiagnosa(i) & "'"
                Set rs = Nothing
                rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
                If rs.EOF = False Then
                    MsgBox "KdDiagnosa " & strKdDiagnosa(i) & " sudah diinput sebelumnya", vbExclamation, "Validasi"
                    GoTo valKdDiagnosa
                End If
                If sp_PeriksaDiagnosa(dbcmd, strKdDiagnosa(i)) = False Then Exit Sub
valKdDiagnosa:
            Next i
            bolvalICD10 = True
        End If
    End If

    'INA DRG
    If chkICD9.value = vbChecked Then
        If dcJenisDiagnosaTindakan.Text = "" Then
            MsgBox "Jenis Diagnosa Tindakan Belum Diisi ", vbCritical, "Validasi"
            dcJenisDiagnosaTindakan.SetFocus
            Exit Sub
        ElseIf intICD9Diplh = 0 Then
            MsgBox "Diagnosa Tindakan Belum Dipilih ", vbCritical, "Validasi"
            txtDiagnosaTindakan.SetFocus
            Exit Sub
        ElseIf dcJenisDiagnosaTindakan.BoundText = "05" Then
            If intICD9Diplh > 1 Then
                MsgBox "Diagnosa Utama tidak bisa lebih dari satu", vbCritical, "Validasi"
                txtDiagnosaTindakan.SetFocus
                Exit Sub
            End If
        End If

        strSQL = ""
        strSQL = "Delete From DetailPeriksaDiagnosa Where NoPendaftaran = '" & txtNoPendaftaran.Text & "' And KdDiagnosa = '" & mstrKdDiagnosa & "' And TglPeriksa = '" & Format(dtpTglPeriksa.value, "yyyy/MM/dd HH:mm:ss") & "' And KdJenisDiagnosa='" & dcJenisDiagnosaTindakan.BoundText & "'"
        Set rs = Nothing
        Call msubRecFO(rs, strSQL)

        For i = 1 To intICD9Diplh
            If bolEditDiagnosa = True Then
                If sp_AUDDetailPeriksaDiagnosa(dbcmd, mstrKdDiagnosa, Right(strKdDiagnosaTindakan(i), 5), "A") = False Then Exit Sub
            Else
                If sp_AUDDetailPeriksaDiagnosa(dbcmd, strKdDiagnosa(1), Right(strKdDiagnosaTindakan(i), 5), "A") = False Then Exit Sub
            End If
        Next i
    End If

    If txtNamaFormPengirim.Text = "frmTransaksiPasien" Then
        frmTransaksiPasien.subLoadRiwayatDiagnosa (False)
    End If
    MsgBox "Data diagnosa berhasil disimpan ", vbInformation, "Informasi"

    Call Add_HistoryLoginActivity("Add_PeriksaDiagnosa+AUD_DetailPeriksaDiagnosa")
    cmdSimpan.Enabled = False
    mstrKdDokter = ""
    intJmlDokter = 0
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    On Error GoTo hell
    If cmdSimpan.Enabled = True Then
        If txtKodeICD.Text <> "" Then
            If MsgBox("Simpan data diagnosa ", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
                Call cmdSimpan_Click
                Exit Sub
            End If
        End If
    End If
    Unload Me

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcInfeksiNosokomial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcInfeksiNosokomial.MatchedWithList = True Then dcPenyebabInfeksiNosokomial.SetFocus
        strSQL = "Select KdInfeksiNosokomial, InfeksiNosokomial From InfeksiNosokomial where StatusEnabled='1'  and (InfeksiNosokomial LIKE '%" & dcInfeksiNosokomial.Text & "%')order by InfeksiNosokomial "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcInfeksiNosokomial.Text = ""
            dcPenyebabInfeksiNosokomial.SetFocus
            Exit Sub
        End If
        dcInfeksiNosokomial.BoundText = rs(0).value
        dcInfeksiNosokomial.Text = rs(1).value
    End If
End Sub

Private Sub dcJenisDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcJenisDiagnosa.MatchedWithList = True Then txtNamaDiagnosa.SetFocus
        strSQL = "select kdjenisdiagnosa, jenisdiagnosa from jenisdiagnosa where StatusEnabled='1'  and (jenisdiagnosa LIKE '%" & dcJenisDiagnosa.Text & "%')order by JenisDiagnosa"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcJenisDiagnosa.Text = ""
            txtNamaDiagnosa.SetFocus
            Exit Sub
        End If
        dcJenisDiagnosa.BoundText = rs(0).value
        dcJenisDiagnosa.Text = rs(1).value
    End If
End Sub

Private Sub dcJenisDiagnosaTindakanx()
    lvwDiagnosaTindakan.ListItems.Clear
    mstrKdJenisDiagnosaTindakan = ""
    mstrKdJenisDiagnosaTindakan = dcJenisDiagnosaTindakan.BoundText
    intICD9DiplhBaru = 0
    Call subLoadLvwICD9
End Sub

Private Sub dcJenisDiagnosaTindakan_Change()
    Call dcJenisDiagnosaTindakan_KeyPress(13)
End Sub

Private Sub dcJenisDiagnosaTindakan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcJenisDiagnosaTindakan.MatchedWithList = True Then cmdSimpan.SetFocus
        strSQL = "select kdjenisdiagnosa, jenisdiagnosa from jenisdiagnosa where kdjenisdiagnosa in('05','11') and StatusEnabled='1'   and (jenisdiagnosa LIKE '%" & dcJenisDiagnosaTindakan.Text & "%') "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcJenisDiagnosaTindakan.Text = ""
            cmdSimpan.SetFocus
            Exit Sub
        End If
        dcJenisDiagnosaTindakan.BoundText = rs(0).value
        dcJenisDiagnosaTindakan.Text = rs(1).value
        Call dcJenisDiagnosaTindakanx
    End If
End Sub

Private Sub dcKetunaanKelainan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKetunaanKelainan.MatchedWithList = True Then chkICD9.SetFocus
        strSQL = "select KdKetunaanKelainan, KetunaanKelainan from KetunaanKelainan where StatusEnabled='1'   and (KetunaanKelainan LIKE '%" & dcKetunaanKelainan.Text & "%')order by KetunaanKelainan "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcKetunaanKelainan.Text = ""
            chkICD9.SetFocus
            Exit Sub
        End If
        dcKetunaanKelainan.BoundText = rs(0).value
        dcKetunaanKelainan.Text = rs(1).value
    End If
End Sub

Private Sub dcMorfologiNeoplasma_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcMorfologiNeoplasma.MatchedWithList = True Then dcKetunaanKelainan.SetFocus
        strSQL = "Select KdMorfologiNeoplasma, MorfologiNeoplasma From MorfologiNeoplasma where StatusEnabled='1'  and (MorfologiNeoplasma LIKE '%" & dcMorfologiNeoplasma.Text & "%')order by MorfologiNeoplasma "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcMorfologiNeoplasma.Text = ""
            dcKetunaanKelainan.SetFocus
            Exit Sub
        End If
        dcMorfologiNeoplasma.BoundText = rs(0).value
        dcMorfologiNeoplasma.Text = rs(1).value
    End If
End Sub

Private Sub dcPenyebabDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcPenyebabDiagnosa.MatchedWithList = True Then dcInfeksiNosokomial.SetFocus
        strSQL = "Select KdPenyebabDiagnosa, PenyebabDiagnosa From PenyebabDiagnosa where StatusEnabled='1'and (PenyebabDiagnosa LIKE '%" & dcInfeksiNosokomial.Text & "%')order by PenyebabDiagnosa "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcPenyebabDiagnosa.Text = ""
            dcInfeksiNosokomial.SetFocus
            Exit Sub
        End If
        dcPenyebabDiagnosa.BoundText = rs(0).value
        dcPenyebabDiagnosa.Text = rs(1).value
    End If
End Sub

Private Sub dcPenyebabInfeksiNosokomial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcPenyebabInfeksiNosokomial.MatchedWithList = True Then dcMorfologiNeoplasma.SetFocus
        strSQL = "Select KdPenyebabIN, PenyebabIN From PenyebabInfeksiNosokomial where StatusEnabled='1'  and (PenyebabIN LIKE '%" & dcPenyebabInfeksiNosokomial.Text & "%')order by PenyebabIN "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcPenyebabInfeksiNosokomial.Text = ""
            dcMorfologiNeoplasma.SetFocus
            Exit Sub
        End If
        dcPenyebabInfeksiNosokomial.BoundText = rs(0).value
        dcPenyebabInfeksiNosokomial.Text = rs(1).value
    End If
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        fraDokter.Visible = False
    End If
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        txtDokter.Text = dgDokter.Columns(1).value
        mstrKdDokter = dgDokter.Columns(0).value
        If mstrKdDokter = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter.Text = ""
            dcJenisDiagnosa.SetFocus
            Exit Sub
        End If
        fraDokter.Visible = False
    End If
End Sub

Private Sub dtpTglPeriksa_Change()
    dtpTglPeriksa.MaxDate = Now
End Sub

Private Sub dtpTglPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtDokter.SetFocus
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo hell
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglPeriksa.value = Now

    Set rs = Nothing
    rs.Open "select * from jenisdiagnosa where StatusEnabled='1' order by JenisDiagnosa DESC", dbConn, adOpenDynamic, adLockOptimistic 'kdJenisDiagnosa not in ('02', '08', '09') and
    Set dcJenisDiagnosa.RowSource = rs
    dcJenisDiagnosa.ListField = rs.Fields(1).Name
    dcJenisDiagnosa.BoundColumn = rs.Fields(0).Name

    'INA DRG
    Set rs = Nothing
    rs.Open "select * from jenisdiagnosa where kdjenisdiagnosa in('05','11') and StatusEnabled='1'", dbConn, adOpenDynamic, adLockOptimistic
    Set dcJenisDiagnosaTindakan.RowSource = rs
    dcJenisDiagnosaTindakan.ListField = rs.Fields(1).Name
    dcJenisDiagnosaTindakan.BoundColumn = rs.Fields(0).Name

    Call subLoadDcSource

    Set rs = Nothing
    intJmlDiagDipilih = 0
    If bolEditDiagnosa = True Then
        subLoadDataDiagnosa
    Else
        subLoadLvw
    End If

    'icd 9
    intICD9Diplh = 0
    intICD9DiplhBaru = 0
    subLoadLvwICD9
    bolvalICD10 = False

hell:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If txtNamaFormPengirim.Text = "frmDaftarPasienRJRIIGD" Then frmDaftarPasienRJRIIGD.Enabled = True
    If txtNamaFormPengirim.Text = "frmTransaksiPasien" Then frmTransaksiPasien.Enabled = True: Call frmTransaksiPasien.subLoadRiwayatDiagnosa(False)
End Sub

Private Sub lvwDiagnosa_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim blnSelected As Boolean
    If Item.Checked = True Then
        intJmlDiagDipilih = intJmlDiagDipilih + 1
        ReDim Preserve strKdDiagnosa(intJmlDiagDipilih)
        strKdDiagnosa(intJmlDiagDipilih) = Item.Key
    Else
        blnSelected = False
        For i = 1 To intJmlDiagDipilih
            If strKdDiagnosa(i) = Item.Key Then blnSelected = True
            If blnSelected = True Then
                If i = intJmlDiagDipilih Then
                    strKdDiagnosa(i) = ""
                Else
                    strKdDiagnosa(i) = strKdDiagnosa(i + 1)
                End If
            End If
        Next i
        intJmlDiagDipilih = intJmlDiagDipilih - 1
    End If
    txtKodeICD.Text = lvwDiagnosa.SelectedItem.Key
End Sub

'add icd9
Private Sub lvwDiagnosaTindakan_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim blnSelectedICD9 As Boolean
    If Item.Checked = True Then
        strSQL = ""
        strSQL = "Select DetailPeriksaDiagnosa.KdDiagnosaTindakan,JenisDiagnosa.JenisDiagnosa From DetailPeriksaDiagnosa Inner Join JenisDiagnosa On DetailPeriksaDiagnosa.KdJenisDiagnosa=JenisDiagnosa.KdJenisDiagnosa " & _
        "Where DetailPeriksaDiagnosa.NoPendaftaran = '" & txtNoPendaftaran.Text & "' And DetailPeriksaDiagnosa.KdDiagnosa = '" & mstrKdDiagnosa & "' And DetailPeriksaDiagnosa.TglPeriksa = '" & Format(dtpTglPeriksa.value, "yyyy/MM/dd HH:mm:ss") & "' " & _
        "And DetailPeriksaDiagnosa.KdDiagnosaTindakan = '" & Right(Item.Key, 5) & "'"
        Set rs = Nothing
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            MsgBox "Diagnosa tindakan yg dipilih sudah diinput pada jenis " & IIf(IsNull(rs(1)), "", rs(1)), vbExclamation, "Validasi"
            Item.Checked = False
            Exit Sub
        End If

        intICD9Diplh = intICD9Diplh + 1
        ReDim Preserve strKdDiagnosaTindakan(intICD9Diplh)
        strKdDiagnosaTindakan(intICD9Diplh) = Right(Item.Key, 5)

        intICD9DiplhBaru = intICD9DiplhBaru + 1
    Else
        blnSelectedICD9 = False
        For i = 1 To intICD9Diplh
            If Right(strKdDiagnosaTindakan(i), 5) = Right(Item.Key, 5) Then blnSelectedICD9 = True
            If blnSelectedICD9 = True Then
                strSQL = ""
                strSQL = "Select KdDiagnosaTindakan From DetailPeriksaDiagnosa Where NoPendaftaran = '" & txtNoPendaftaran.Text & "' And KdDiagnosa = '" & mstrKdDiagnosa & "' And TglPeriksa = '" & Format(dtpTglPeriksa.value, "yyyy/MM/dd HH:mm:ss") & "' And KdDiagnosaTindakan = '" & Right(strKdDiagnosaTindakan(i), 5) & "'"
                Set rs = Nothing
                Call msubRecFO(rs, strSQL)
                If rs.EOF = False Then
                    If sp_AUDDetailPeriksaDiagnosa(dbcmd, mstrKdDiagnosa, Right(strKdDiagnosaTindakan(i), 5), "D") = False Then Exit Sub
                End If
                If i = intICD9Diplh Then
                    strKdDiagnosaTindakan(i) = ""
                Else
                    strKdDiagnosaTindakan(i) = strKdDiagnosaTindakan(i + 1)
                End If
            End If
        Next i
        intICD9Diplh = intICD9Diplh - 1
        If intICD9DiplhBaru = 0 Then Exit Sub
        intICD9DiplhBaru = intICD9DiplhBaru - 1
    End If
    txtKodeICD9.Text = lvwDiagnosaTindakan.SelectedItem.Key
End Sub

Private Sub lvwDiagnosa_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo errLoad
    txtKodeICD.Text = lvwDiagnosa.SelectedItem.Key
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub lvwDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcPenyebabDiagnosa.SetFocus
End Sub

Private Sub lvwDiagnosaTindakan_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo errLoad
    txtKodeICD9.Text = Right(lvwDiagnosaTindakan.SelectedItem.Key, 5)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub lvwDiagnosaTindakan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJenisDiagnosaTindakan.SetFocus
End Sub

Private Sub txtDiagnosaTindakan_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then lvwDiagnosaTindakan.SetFocus
End Sub

Private Sub txtDokter_Change()
    strFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
    mstrKdDokter = ""
    fraDokter.Visible = True
    fraDokter.Top = 3000
    fraDokter.Left = 240
    Call subLoadDokter
End Sub

Private Sub txtDokter_GotFocus()
    txtDokter.SelStart = 0
    txtDokter.SelLength = Len(txtDokter.Text)
    fraDokter.Visible = True
    strFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
    Call subLoadDokter
End Sub

Private Sub txtDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If fraDokter.Visible = False Then Exit Sub
        dgDokter.SetFocus
    End If
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        dgDokter.SetFocus
    End If
    If KeyAscii = 27 Then
        fraDokter.Visible = False
    End If
    If KeyAscii = 39 Then KeyAscii = 0
hell:
End Sub

Private Sub txtNamaDiagnosa_Change()
    subLoadLvw "AND (NamaDiagnosa LIKE '%" & txtNamaDiagnosa.Text & "%' or KdDiagnosa LIKE '%" & txtNamaDiagnosa.Text & "%') "
End Sub

Private Sub txtDiagnosaTindakan_Change()
    Call subLoadLvwICD9
End Sub

Private Sub txtNamaDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then lvwDiagnosa.SetFocus
End Sub

'tambahan field di tabel periksa diagnosa
Private Sub subLoadDcSource()
    On Error GoTo hell
    Call msubDcSource(dcPenyebabDiagnosa, rs, "Select KdPenyebabDiagnosa, PenyebabDiagnosa From PenyebabDiagnosa where StatusEnabled='1' order by PenyebabDiagnosa")
    Call msubDcSource(dcInfeksiNosokomial, rs, "Select KdInfeksiNosokomial, InfeksiNosokomial From InfeksiNosokomial where StatusEnabled='1' order by InfeksiNosokomial")
    Call msubDcSource(dcPenyebabInfeksiNosokomial, rs, "Select KdPenyebabIN, PenyebabIN From PenyebabInfeksiNosokomial where StatusEnabled='1' order by PenyebabIN")
    Call msubDcSource(dcMorfologiNeoplasma, rs, "Select KdMorfologiNeoplasma, MorfologiNeoplasma From MorfologiNeoplasma where StatusEnabled='1' order by MorfologiNeoplasma")
    Call msubDcSource(dcKetunaanKelainan, rs, "select KdKetunaanKelainan, KetunaanKelainan from KetunaanKelainan where StatusEnabled='1' order by KetunaanKelainan")
    Exit Sub
hell:
    Call msubPesanError
End Sub

'untuk loading data dokter
Private Sub subLoadDokter()
    On Error Resume Next
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & strFilterDokter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDokter = rs.RecordCount
    Set dgDokter.DataSource = rs
    With dgDokter
        .Columns(0).Width = 1300
        .Columns(1).Width = 3000
        .Columns(2).Width = 500
        .Columns(3).Width = 2500
    End With
End Sub

Private Sub subLoadDataDiagnosa()
    On Error GoTo hell
    Dim X As Integer
    strSQL = "SELECT * FROM V_Diagnosa WHERE (KdDiagnosa = '" & mstrKdDiagnosa & "')"
    Set rs = Nothing
    msubRecFO rs, strSQL
    lvwDiagnosa.ListItems.Clear
    For X = 1 To rs.RecordCount
        Set itemAll = lvwDiagnosa.ListItems.Add(, rs("KdDiagnosa").value, rs("NamaDiagnosa").value)
        lvwDiagnosa.ListItems(X).Checked = True
        lvwDiagnosa.ListItems(X).ForeColor = vbBlue
        lvwDiagnosa.ListItems(X).Bold = True

        intJmlDiagDipilih = X
        rs.MoveNext
    Next X
    Exit Sub
hell:
    msubPesanError
End Sub

'untuk loading data listview diagnosa
Private Sub subLoadLvw(Optional strKriteria As String)
    On Error GoTo hell
    strSQL = "SELECT * FROM V_Diagnosa WHERE (KdDiagnosa LIKE '%" & txtNamaDiagnosa.Text & "%') OR (NamaDiagnosa LIKE '%" & txtNamaDiagnosa.Text & "%') and StatusEnabled='1' ORDER BY NamaDiagnosa"
    msubRecFO rs, strSQL
    lvwDiagnosa.ListItems.Clear
    Do While rs.EOF = False
        Set itemAll = lvwDiagnosa.ListItems.Add(, rs("KdDiagnosa").value, rs("NamaDiagnosa").value)
        rs.MoveNext
    Loop

    If intJmlDiagDipilih = 0 Then Exit Sub
    For i = 1 To lvwDiagnosa.ListItems.Count
        For j = 1 To intJmlDiagDipilih
            If lvwDiagnosa.ListItems(i).Key = strKdDiagnosa(j) Then
                lvwDiagnosa.ListItems(i).Checked = True
                lvwDiagnosa.ListItems(i).ForeColor = vbBlue
                lvwDiagnosa.ListItems(i).Bold = True
            End If
        Next j
    Next i
    lvwDiagnosa.View = lvwList
    Exit Sub
hell:
End Sub

'ICD 9 INA DRG
Private Sub subLoadLvwICD9(Optional strKriteria As String)
    Dim X As Integer
    strSQL = ""
    strSQL = "SELECT * FROM DiagnosaTindakan WHERE (KdDiagnosaTindakan LIKE '%" & txtDiagnosaTindakan.Text & "%') OR (DiagnosaTindakan LIKE '%" & txtDiagnosaTindakan.Text & "%') and StatusEnabled='1' ORDER BY DiagnosaTindakan"
    msubRecFO rs, strSQL
    lvwDiagnosaTindakan.ListItems.Clear
    Do While rs.EOF = False
        Set itemAll = lvwDiagnosaTindakan.ListItems.Add(, "A" & rs("KdDiagnosaTindakan").value, rs("DiagnosaTindakan").value)
        rs.MoveNext
    Loop

    strSQL = ""
    strSQL = "Select KdDiagnosaTindakan From DetailPeriksaDiagnosa Where NoPendaftaran = '" & mstrNoPen & "' And KdDiagnosa = '" & mstrKdDiagnosa & "' And KdJenisDiagnosa = '" & mstrKdJenisDiagnosaTindakan & "'"
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    intICD9Diplh = rs.RecordCount + intICD9DiplhBaru
    ReDim Preserve strKdDiagnosaTindakan(intICD9Diplh)
    For X = 1 To rs.RecordCount
        strKdDiagnosaTindakan(X) = rs("KdDiagnosaTindakan").value
        rs.MoveNext
    Next X
    If intICD9Diplh = 0 Then Exit Sub
    For i = 1 To lvwDiagnosaTindakan.ListItems.Count
        For j = 1 To intICD9Diplh
            If Right(lvwDiagnosaTindakan.ListItems(i).Key, 5) = Right(strKdDiagnosaTindakan(j), 5) Then
                lvwDiagnosaTindakan.ListItems(i).Checked = True
                lvwDiagnosaTindakan.ListItems(i).ForeColor = vbBlue
                lvwDiagnosaTindakan.ListItems(i).Bold = True
            End If
        Next j
    Next i
    lvwDiagnosaTindakan.View = lvwList
End Sub

'untuk menyimpan data diagnosa pasien
Private Function sp_PeriksaDiagnosa(adoCommand As ADODB.Command, strKodeDiagnosa As String) As Boolean
    On Error GoTo errSimpan
    Set adoCommand = New ADODB.Command
    sp_PeriksaDiagnosa = False
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, Right(txtNoCM.Text, 6))
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        If txtNamaFormPengirim.Text = "frmDaftarPasienRJRIIGD" Then
            .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuanganPasien)
        Else
            .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        End If
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, mstrKdDokter)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dtpTglPeriksa.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdDiagnosa", adVarChar, adParamInput, 7, strKodeDiagnosa)
        .Parameters.Append .CreateParameter("KdJenisDiagnosa", adChar, adParamInput, 2, dcJenisDiagnosa.BoundText)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("KdPenyebabDiagnosa", adSmallInt, adParamInput, , IIf(dcPenyebabDiagnosa.BoundText = "", Null, dcPenyebabDiagnosa.BoundText))
        .Parameters.Append .CreateParameter("KdInfeksiNosokomial", adVarChar, adParamInput, 2, IIf(dcInfeksiNosokomial.BoundText = "", Null, dcInfeksiNosokomial.BoundText))
        .Parameters.Append .CreateParameter("KdPenyebabIN", adSmallInt, adParamInput, , IIf(dcPenyebabInfeksiNosokomial.BoundText = "", Null, dcPenyebabInfeksiNosokomial.BoundText))
        .Parameters.Append .CreateParameter("KdMorfologiNeoplasma", adTinyInt, adParamInput, , IIf(dcMorfologiNeoplasma.BoundText = "", Null, dcMorfologiNeoplasma.BoundText))
        .Parameters.Append .CreateParameter("KdKetunaanKelainan", adTinyInt, adParamInput, , IIf(dcKetunaanKelainan.BoundText = "", Null, dcKetunaanKelainan.BoundText))

        .ActiveConnection = dbConn
        .CommandText = "Add_PeriksaDiagnosa"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data", vbCritical, "Validasi"
        Else
            sp_PeriksaDiagnosa = True
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

Private Function sp_AUDDetailPeriksaDiagnosa(adoCommand As ADODB.Command, strKdICD10 As String, strKdICD9 As String, f_Status As String) As Boolean
    On Error GoTo hell
    Set adoCommand = New ADODB.Command
    sp_AUDDetailPeriksaDiagnosa = False
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdDiagnosa", adVarChar, adParamInput, 7, strKdICD10)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dtpTglPeriksa.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdDiagnosaTindakan", adVarChar, adParamInput, 7, strKdICD9)
        .Parameters.Append .CreateParameter("KdJenisDiagnosaTindakan", adChar, adParamInput, 2, dcJenisDiagnosaTindakan.BoundText)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_DetailPeriksaDiagnosa"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data", vbCritical, "Validasi"
        Else
            sp_AUDDetailPeriksaDiagnosa = True
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Function
hell:
    sp_AUDDetailPeriksaDiagnosa = False
    Call deleteADOCommandParameters(adoCommand)
    Set adoCommand = Nothing
    Call msubPesanError
End Function

