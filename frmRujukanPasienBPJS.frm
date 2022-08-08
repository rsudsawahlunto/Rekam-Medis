VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRujukanPasienBPJS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Rujukan Pasien BPJS"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRujukanPasienBPJS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   9645
   Begin VB.Timer tmrJalan 
      Interval        =   120
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame fraFaskes 
      Caption         =   "Daftar Faskes"
      Height          =   3495
      Left            =   1560
      TabIndex        =   56
      Top             =   480
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdTutupDaftarRS 
         BackColor       =   &H0000FF00&
         Caption         =   "&Tutup"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   2760
         Width           =   6255
      End
      Begin MSFlexGridLib.MSFlexGrid fgFaskes 
         Height          =   2175
         Left            =   240
         TabIndex        =   57
         Top             =   360
         Visible         =   0   'False
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3836
         _Version        =   393216
         AllowUserResizing=   3
         Appearance      =   0
      End
   End
   Begin MSDataGridLib.DataGrid gridDiagnosa 
      Height          =   2895
      Left            =   1320
      TabIndex        =   51
      Top             =   -2640
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
   Begin VB.Frame fraDataPasien 
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
      Height          =   2775
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   9615
      Begin VB.OptionButton optRi 
         Caption         =   "RI"
         Height          =   210
         Left            =   8640
         TabIndex        =   61
         Top             =   1800
         Width           =   495
      End
      Begin VB.OptionButton optRj 
         Caption         =   "RJ"
         Height          =   210
         Left            =   7800
         TabIndex        =   60
         Top             =   1800
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox txtPelayanan 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Height          =   330
         Left            =   7800
         TabIndex        =   49
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         TabIndex        =   39
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtPenjamin 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5760
         TabIndex        =   36
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtJenisPasien 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1920
         TabIndex        =   34
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtDiagnosa 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         TabIndex        =   29
         Top             =   2280
         Width           =   9135
      End
      Begin VB.TextBox txtKelas 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7800
         TabIndex        =   27
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         MaxLength       =   12
         TabIndex        =   23
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   22
         Top             =   480
         Width           =   4335
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6360
         MaxLength       =   9
         TabIndex        =   21
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtNoKartu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtNoSEP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Height          =   330
         Left            =   2400
         TabIndex        =   16
         Top             =   1680
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker dtpTglSEP 
         Height          =   330
         Left            =   6120
         TabIndex        =   38
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   58589187
         UpDown          =   -1  'True
         CurrentDate     =   37813
      End
      Begin MSComCtl2.DTPicker dtpTglLahir 
         Height          =   330
         Left            =   7800
         TabIndex        =   42
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   58589187
         UpDown          =   -1  'True
         CurrentDate     =   37813
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Jenis Rujukan"
         Height          =   270
         Index           =   13
         Left            =   7800
         TabIndex        =   50
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Lahir"
         Height          =   210
         Index           =   10
         Left            =   7800
         TabIndex        =   41
         Top             =   240
         Width           =   750
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Index           =   9
         Left            =   240
         TabIndex        =   40
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kelompok Pasien"
         Height          =   210
         Index           =   8
         Left            =   5760
         TabIndex        =   37
         Top             =   840
         Width           =   1365
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Peserta"
         Height          =   210
         Index           =   7
         Left            =   1920
         TabIndex        =   35
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnosa (Penyakit)"
         Height          =   210
         Index           =   4
         Left            =   240
         TabIndex        =   30
         Top             =   2040
         Width           =   1620
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kelas"
         Height          =   210
         Index           =   3
         Left            =   7800
         TabIndex        =   28
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. RM"
         Height          =   210
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   1920
         TabIndex        =   25
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   6360
         TabIndex        =   24
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kartu"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. SEP"
         Height          =   210
         Index           =   1
         Left            =   6120
         TabIndex        =   18
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. SEP"
         Height          =   210
         Index           =   0
         Left            =   2400
         TabIndex        =   17
         Top             =   1440
         Width           =   660
      End
   End
   Begin VB.Frame fraRujukan 
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
      Height          =   2535
      Left            =   0
      TabIndex        =   4
      Top             =   4080
      Width           =   9615
      Begin VB.OptionButton optFaskes2 
         Caption         =   "Faskes 2 (RS)"
         Height          =   210
         Left            =   2520
         TabIndex        =   55
         Top             =   840
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optFaskes1 
         Caption         =   "Faskes 1"
         Height          =   210
         Left            =   1320
         TabIndex        =   54
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtKdDiagnosa 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         TabIndex        =   52
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtNoRujukan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txtKdPPKRujukan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2520
         TabIndex        =   43
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtPPKRujukan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         TabIndex        =   31
         Top             =   1080
         Width           =   6135
      End
      Begin VB.Frame Frame6 
         Caption         =   "Tipe Rujukan"
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
         Left            =   5760
         TabIndex        =   12
         Top             =   240
         Width           =   3615
         Begin VB.OptionButton optRujukBalik 
            Caption         =   "Rujuk Balik"
            Height          =   210
            Left            =   2280
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optPartial 
            Caption         =   "Partial"
            Height          =   210
            Left            =   1320
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optPenuh 
            Caption         =   "Penuh"
            Height          =   210
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.TextBox txtCatatan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4920
         MaxLength       =   30
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "-"
         Top             =   2040
         Width           =   4455
      End
      Begin VB.TextBox txtDiagnosaRujukan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   4575
      End
      Begin MSComCtl2.DTPicker dtpTglRujukan 
         Height          =   330
         Left            =   4080
         TabIndex        =   5
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   58589187
         UpDown          =   -1  'True
         CurrentDate     =   37813
      End
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   330
         Left            =   6480
         TabIndex        =   45
         Top             =   1080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Diagnosa"
         Height          =   210
         Index           =   14
         Left            =   2880
         TabIndex        =   53
         Top             =   1440
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Rujukan"
         Height          =   210
         Index           =   12
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode PPK Rujukan"
         Height          =   210
         Index           =   11
         Left            =   960
         TabIndex        =   44
         Top             =   120
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Poli Rujukan"
         Height          =   210
         Index           =   6
         Left            =   6480
         TabIndex        =   33
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirujuk Ke"
         Height          =   210
         Index           =   5
         Left            =   240
         TabIndex        =   32
         Top             =   840
         Width           =   825
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnosa (Penyakit) Rujukan"
         Height          =   210
         Index           =   22
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   2325
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan Rujukan"
         Height          =   210
         Index           =   32
         Left            =   4920
         TabIndex        =   10
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Rujukan"
         Height          =   210
         Index           =   2
         Left            =   4080
         TabIndex        =   6
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   6600
      Width           =   9615
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   3720
         TabIndex        =   59
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   495
         Left            =   5160
         TabIndex        =   48
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   6600
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   9960
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
   Begin VB.Label lblJalan 
      Alignment       =   2  'Center
      Caption         =   "PASTIKAN RS RUJUKAN DAN KODE RUJUKAN BENAR !!!!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   62
      Top             =   1080
      Width           =   9615
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRujukanPasienBPJS.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7800
      Picture         =   "frmRujukanPasienBPJS.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRujukanPasienBPJS.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmRujukanPasienBPJS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim result() As String
Dim arr() As String
Dim n As Integer
Dim sJnsPelayanan As String
Dim spoliRujukan As String
Dim stipeRujukan As String
Dim sdiagnosaRujukan As String
Dim jnsFaskes As String
Dim kdOrNamaFaskes As String
Dim strNamaPegawai As String

Public sKdPPKRujukan As String

Private Sub cmdCetak_Click()
    On Error GoTo Bandung
    strSQL = "Select * From RujukanPasienVclaim where Nopendaftaran ='" & txtNoPendaftaran.Text & "' "
    Call msubRecFO(rs, strSQL)
   
        If rs.EOF = False Then
                   
                   mstrNoSepRujukan = txtNoSEP.Text
                   
                   frmViewerRujukanLuar.Show
        Exit Sub
        End If
Bandung:

End Sub

Private Sub cmdHapus_Click()
On Error GoTo pesan
    
    If MsgBox("Yakin akan menghapus rujukan " & vbCrLf & "dengan No. Rujukan: " & txtNoRujukan.Text & " ?", vbQuestion + vbYesNo, "Hapus Rujukan Pasien BPJS") = vbNo Then Exit Sub
    If Len(Trim(txtNoRujukan.Text)) < 19 Then Exit Sub
    
    If DeleteRujukanVclaim = False Then Exit Sub
    
    If sp_RujukanPasienVclaim("D") = False Then Exit Sub
    cmdHapus.Enabled = False
    
Exit Sub
pesan:
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo pesan
        
    If Periksa("text", txtPpkRujukan, "Tujukan Rujukan belum diisi") = False Then Exit Sub
    If Periksa("text", txtDiagnosaRujukan, "Diagnosa Rujukan belum diisi") = False Then Exit Sub
    
    If Len(Trim(txtNoRujukan.Text)) = 0 Then
        If InsertRujukanVclaim = False Then Exit Sub
    Else
        If UpdateRujukanVclaim = False Then Exit Sub
    End If
    
    If sp_RujukanPasienVclaim("A") = False Then Exit Sub
    
    cmdSimpan.Enabled = False
Exit Sub
pesan:
End Sub

Private Sub cmdTutup_Click()
    sdiagnosaRujukan = ""
    stipeRujukan = ""
    spoliRujukan = ""
    sJnsPelayanan = ""
'    frmUbahJenisPasien.Enabled = True
    Unload Me
End Sub

Private Sub cmdTutupDaftarRS_Click()
    fraFaskes.Visible = False
End Sub

Private Sub fgFaskes_DblClick()
    Call fgFaskes_KeyPress(13)
End Sub

Public Sub fgFaskes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With fgFaskes
            If fgFaskes.rows <> 1 Then
                txtKdPPKRujukan.Text = .TextMatrix(.row, 0)
                txtPpkRujukan.Text = .TextMatrix(.row, 1)
                
                .Visible = False
            End If
        End With
    End If
End Sub

Private Sub Form_Load()
On Error GoTo pesan

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    
    dtpTglRujukan.value = Format(Now, "dd/MM/yyyy")
    dtpTglLahir.value = Format(Now, "dd/MM/yyyy")
    
    Call msubDcSource(dcRuangan, rs, "SELECT KdRuangan,NamaRuangan FROM V_RuanganPelayanan ORDER BY NamaRuangan")

Exit Sub
pesan:
    Call msubPesanError
End Sub

Public Function InsertRujukanVclaim() As Boolean
On Error GoTo pesan
    
    InsertRujukanVclaim = True
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
        
'        If txtPelayanan.Text = "R.Jalan" Then
'            sJnsPelayanan = 2
'        ElseIf txtPelayanan.Text = "R.Inap" Then
'            sJnsPelayanan = 1
'        End If
        
'        Select Case txtPelayanan.Text
'                Case "03"
'                    sJnsPelayanan = "1"
'                Case "02"
'                    sJnsPelayanan = "2"
'                Case "01"
'                    sJnsPelayanan = "2"
'                Case "06"
'                    sJnsPelayanan = "2"
'                Case "22"
'                    sJnsPelayanan = "2"
'            End Select
        
        If optRj.value = True Then
            sJnsPelayanan = "2"
        ElseIf optRi.value = True Then
            sJnsPelayanan = "1"
        End If

        
        If optPenuh.value = True Then
            stipeRujukan = 0
        ElseIf optPartial.value = True Then
            stipeRujukan = 1
        ElseIf optRujukBalik.value = True Then
            stipeRujukan = 2
        End If
        
        Call msubRecFO(dbrs, "SELECT KodeExternal FROM Ruangan WHERE KdRuangan='" & dcRuangan.BoundText & "'")
        If Not dbrs.EOF Then
            spoliRujukan = Trim(dbrs(0).value)
        End If
        
'        strSQL = "SELECT dbo.getNamaPegawaiByIdPegawai('" & strIDPegawai & "')"
        strSQL = "SELECT NamaLengkap FROM DataPegawai WHERE IdPegawai='" & strIDPegawai & "'"
            Call msubRecFO(rs, strSQL)
            strNamaPegawai = rs(0).value
        
        result = context.InsertRujukan(txtNoSEP.Text, Format(dtpTglRujukan.value, "yyyy-MM-dd"), txtKdPPKRujukan.Text, sJnsPelayanan, _
                 Trim(txtCatatan.Text), sdiagnosaRujukan, stipeRujukan, spoliRujukan, strNamaPegawai)
        
        For n = LBound(result) To UBound(result)
            arr = Split(result(n), ":")
            Select Case arr(0)
                Case "error"
                    MsgBox arr(1), vbExclamation, "Rujukan Pasien BPJS"
                    InsertRujukanVclaim = False
                    Exit Function
                Case "DIAGNOSA-NORUJUKAN"
                    txtNoRujukan.Text = arr(1)
                    Exit For
            End Select
        Next n
        
        MsgBox "Insert Rujukan Berhasil" & vbCrLf & "Dengan No. Rujukan: " & txtNoRujukan.Text, vbInformation, "Rujukan Pasien BPJS"
        
    End If
    
Exit Function
pesan:
    InsertRujukanVclaim = False
    Call msubPesanError
End Function

Public Function UpdateRujukanVclaim() As Boolean
On Error GoTo pesan
    
    UpdateRujukanVclaim = True
    If (Dir("C:\SDK\askes\result.tlb") <> "") Then
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
        
'        If txtPelayanan.Text = "R.Jalan" Then
'            sJnsPelayanan = 2
'        ElseIf txtPelayanan.Text = "R.Inap" Then
'            sJnsPelayanan = 1
'        End If
        
        If optRj.value = True Then
            sJnsPelayanan = "2"
        ElseIf optRi.value = True Then
            sJnsPelayanan = "1"
        End If

        
        If optPenuh.value = True Then
            stipeRujukan = 0
        ElseIf optPartial.value = True Then
            stipeRujukan = 1
        ElseIf optRujukBalik.value = True Then
            stipeRujukan = 2
        End If
        
        Call msubRecFO(dbrs, "SELECT KodeExternal FROM Ruangan WHERE KdRuangan='" & dcRuangan.BoundText & "'")
        If Not dbrs.EOF Then
            spoliRujukan = Trim(dbrs(0).value)
        End If
        
        sdiagnosaRujukan = Trim(txtKdDiagnosa.Text)
        
        result = context.UpdateRujukan(txtNoRujukan.Text, txtKdPPKRujukan.Text, stipeRujukan, sJnsPelayanan, Trim(txtCatatan.Text), _
                sdiagnosaRujukan, stipeRujukan, spoliRujukan, strIDPegawaiAktif)

        For n = LBound(result) To UBound(result)
            arr = Split(result(n), ":")
            Select Case arr(0)
                Case "error"
                    MsgBox arr(1), vbExclamation, "Rujukan Pasien BPJS"
                    UpdateRujukanVclaim = False
                    Exit Function
            End Select
        Next n
        
        MsgBox "Update Rujukan Berhasil" & vbCrLf & "Dengan No. Rujukan: " & Replace(result(0), "Sukses:", ""), vbInformation, "Rujukan Pasien BPJS"
    End If
    
Exit Function
pesan:
    UpdateRujukanVclaim = False
    Call msubPesanError
End Function

Public Function DeleteRujukanVclaim() As Boolean
On Error GoTo pesan
    
    DeleteRujukanVclaim = True
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
        
        If txtPelayanan.Text = "R.Jalan" Then
            sJnsPelayanan = 2
        ElseIf txtPelayanan.Text = "R.Inap" Then
            sJnsPelayanan = 1
        End If
        
        result = context.DeleteRujukan(txtNoRujukan.Text, strIDPegawaiAktif)
        
        For n = LBound(result) To UBound(result)
            arr = Split(result(n), ":")
            Select Case arr(0)
                Case "error"
                    MsgBox arr(1), vbExclamation, "Rujukan Pasien BPJS"
                    DeleteRujukanVclaim = False
                    Exit Function
            End Select
        Next n
        
        MsgBox "No. Rujukan: " & arr(1) & " Berhasil Dihapus", vbInformation, "Rujukan Pasien BPJS"
    End If
    
Exit Function
pesan:
    DeleteRujukanVclaim = False
    Call msubPesanError
End Function

Private Function sp_RujukanPasienVclaim(f_status As String) As Boolean
On Error GoTo pesan
    
    sp_RujukanPasienVclaim = True
    Dim adoCommand As New ADODB.Command
    Set adoCommand = New ADODB.Command
    MousePointer = vbHourglass
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("NoRujukan", adVarChar, adParamInput, 30, txtNoRujukan.Text)
        .Parameters.Append .CreateParameter("KdRujukanAsal", adChar, adParamInput, 2, Null)
        .Parameters.Append .CreateParameter("SubRujukanAsal", adVarChar, adParamInput, 100, Null)
        .Parameters.Append .CreateParameter("NamaPerujuk", adVarChar, adParamInput, 100, Null)
        .Parameters.Append .CreateParameter("TglRujukan", adDate, adParamInput, , Format(dtpTglRujukan, "yyyy/MM/dd hh:mm:ss"))
        
        sdiagnosaRujukan = IIf(txtKdDiagnosa.Text = "", sdiagnosaRujukan, Trim(txtKdDiagnosa.Text))
        .Parameters.Append .CreateParameter("DiagnosaRujukan", adVarChar, adParamInput, 100, sdiagnosaRujukan)
        
        .Parameters.Append .CreateParameter("NoSEP", adVarChar, adParamInput, 30, txtNoSEP.Text)
        .Parameters.Append .CreateParameter("TglSEP", adDate, adParamInput, , Format(dtpTglSEP, "yyyy/MM/dd hh:mm:ss"))
        
        Call msubRecFO(dbrs, "SELECT KodeExternal FROM Ruangan WHERE KdRuangan='" & dcRuangan.BoundText & "'")
        If Not dbrs.EOF Then
            spoliRujukan = Trim(dbrs(0).value)
        End If
        .Parameters.Append .CreateParameter("PoliTujukan", adVarChar, adParamInput, 30, dcRuangan.Text)
        .Parameters.Append .CreateParameter("KdPPKRujukan", adVarChar, adParamInput, 20, Trim(txtKdPPKRujukan.Text))
        .Parameters.Append .CreateParameter("TujukanRujukan", adVarChar, adParamInput, 500, Trim(txtPpkRujukan.Text))
        .Parameters.Append .CreateParameter("CatatanRujukan", adVarChar, adParamInput, 8000, Trim(txtCatatan.Text))
        
        If optPenuh.value = True Then
            stipeRujukan = 0
        ElseIf optPartial.value = True Then
            stipeRujukan = 1
        ElseIf optRujukBalik.value = True Then
            stipeRujukan = 2
        End If
        .Parameters.Append .CreateParameter("TipeRujukan", adTinyInt, adParamInput, , stipeRujukan)
        
'        If txtPelayanan.Text = "R.Jalan" Then
'            sJnsPelayanan = 2
'        ElseIf txtPelayanan.Text = "R.Inap" Then
'            sJnsPelayanan = 1
'        End If
        .Parameters.Append .CreateParameter("JnsPelayanan", adTinyInt, adParamInput, , CInt(sJnsPelayanan))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("status", adChar, adParamInput, 1, f_status)
        
        .ActiveConnection = dbConn
        .CommandText = "AUD_RujukanPasienVclaim"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data Rujukan Pasien", vbCritical, "Validasi"
            sp_RujukanPasienVclaim = False
        Else
            Call Add_HistoryLoginActivity("AUD_RujukanPasienVclaim")
        End If
        
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    MousePointer = vbDefault
    
Exit Function
pesan:
    sp_RujukanPasienVclaim = False
    Call msubPesanError
End Function

Private Sub gridDiagnosa_Click()
    WheelHook.WheelUnHook
    Set MyProperty = gridDiagnosa
    WheelHook.WheelHook gridDiagnosa
End Sub

Private Sub gridDiagnosa_DblClick()
    Call gridDiagnosa_KeyPress(13)
End Sub

Private Sub gridDiagnosa_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 13 Then
        sdiagnosaRujukan = gridDiagnosa.Columns(0).value
        txtKdDiagnosa.Text = sdiagnosaRujukan
        txtDiagnosaRujukan.Text = gridDiagnosa.Columns(1).value
        gridDiagnosa.Visible = False
        cmdSimpan.SetFocus
    End If
    
    If KeyAscii = 27 Then
        gridDiagnosa.Visible = False
        txtDiagnosaRujukan.SetFocus
    End If
Exit Sub
hell:
End Sub

Private Sub txtDiagnosaRujukan_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 13 Then
        gridDiagnosa.Visible = True
        gridDiagnosa.Top = 2520
        gridDiagnosa.Left = txtDiagnosaRujukan.Left
        
        Set rs = Nothing
        strSQL = "select KdDiagnosa, NamaDiagnosa from V_Diagnosa where (NamaDiagnosa like '%" & txtDiagnosaRujukan.Text & "%' or KdDiagnosa LIKE '%" & txtDiagnosaRujukan.Text & "%') order by NamaDiagnosa"
        rs.Open strSQL, dbConn, adOpenDynamic, adLockOptimistic
        Set gridDiagnosa.DataSource = rs
        With gridDiagnosa
            .Columns(0).Caption = "Kode Diagnosa"
            .Columns(0).Width = 1300
            .Columns(1).Caption = "Nama Diagnosa"
            .Columns(1).Width = 8000
        End With
        Set rs = Nothing
        gridDiagnosa.SetFocus
    End If
    If KeyAscii = 27 Then
        gridDiagnosa.Visible = False
    End If
    
Exit Sub
hell:
End Sub

Private Sub txtNoSEP_KeyPress(KeyAscii As Integer)
On Error GoTo Bandung
    If KeyAscii = 13 Then

If Periksa("text", txtNoSEP, "") = False Then
        MsgBox "Silakan isi No SEP", vbCritical, "No SEP"
        Exit Sub
End If
        strSQL = "Select  * from V_CetakSuratJaminanPelayanan where NoSJP='" & txtNoSEP.Text & "'"
        Call msubRecFO(rs, strSQL)
        
                If rs.EOF = True Then
                    MsgBox "DATA YANG DIMASUKAN TIDAK VALID", vbCritical, "SEP SALAH"
                    txtNoCM.Text = ""
                    txtNamaPasien.Text = ""
                    txtNoPendaftaran.Text = ""
                    txtNoKartu.Text = ""
                    txtJK.Text = ""
                    txtJenisPasien.Text = ""
                    txtPenjamin.Text = ""
                    txtKelas.Text = ""
                    txtDiagnosa.Text = ""
                    txtNoRujukan.Text = ""
                    txtKdPPKRujukan.Text = ""
                    txtPpkRujukan.Text = ""
                    txtKdDiagnosa.Text = ""
                    txtDiagnosaRujukan.Text = ""
                    txtCatatan.Text = ""
                    Exit Sub
                End If
       

        txtNoCM.Text = rs("NoCM").value
        txtNamaPasien.Text = rs("NamaPeserta").value
        txtNoPendaftaran.Text = rs("NoPendaftaran").value
        txtNoKartu.Text = rs("NoKartuPeserta").value
        txtJK.Text = rs("JK").value
        txtJenisPasien.Text = rs("JenisPasien").value
        txtPenjamin.Text = "BPJS"
        txtKelas.Text = rs("KelasBPJS").value
        txtDiagnosa.Text = rs("NamaDiagnosa").value
        dtpTglLahir.value = Format(rs("TglLahir").value, "dd/MM/yyyy")
        dtpTglSEP.value = Format(rs("TglSJP").value, "dd/MM/yyyy")

        
        Call msubRecFO(dbrs, "SELECT * FROM RujukanPasienVclaim WHERE  NoSEP='" & txtNoSEP.Text & "' ")
        If Not dbrs.EOF Then
            txtNoRujukan.Text = dbrs("NoRujukan").value
            dtpTglRujukan.value = Format(dbrs("TglRujukan").value, "dd/MM/yyyy")
            txtKdPPKRujukan.Text = dbrs("KdPPKRujukan").value
            txtPpkRujukan.Text = dbrs("TujuanRujukan").value

            Call msubRecFO(rs, "SELECT KdRuangan,NamaRuangan FROM Ruangan WHERE KodeExternal='" & dbrs("PoliTujuan").value & "'")
            If Not rs.EOF Then
                dcRuangan.BoundText = rs("KdRuangan").value
                dcRuangan.Text = rs("NamaRuangan").value
            End If

            Call msubRecFO(rs, "SELECT KdDiagnosa,NamaDiagnosa FROM V_Diagnosa WHERE KdDiagnosa='" & dbrs("DiagnosaRujukan").value & "'")
            If Not rs.EOF Then
                txtKdDiagnosa.Text = rs(0).value
                txtDiagnosaRujukan.Text = rs(1).value
            End If

            txtCatatan.Text = dbrs("CatatanRujukan").value
            If dbrs("TipeRujukan").value = "0" Then
                optPenuh.value = True
            ElseIf dbrs("TipeRujukan").value = "1" Then
                optPartial.value = True
            ElseIf dbrs("TipeRujukan").value = "2" Then
                optRujukBalik.value = True
            End If
        Exit Sub
        End If
        
        txtNoRujukan.Text = ""
        txtKdPPKRujukan.Text = ""
        txtPpkRujukan.Text = ""
        txtKdDiagnosa.Text = ""
        txtDiagnosaRujukan.Text = ""
        txtCatatan.Text = ""
        
End If
Exit Sub
Bandung:
End Sub

Private Sub txtPPKRujukan_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
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
            
            kdOrNamaFaskes = Trim(txtPpkRujukan.Text)
            If optFaskes1.value = True Then
                jnsFaskes = "1"
            Else
                jnsFaskes = "2"
            End If
            
            result = context.RefFasilitasKesehatan(kdOrNamaFaskes, jnsFaskes)
            
            Dim strResult As String
            strResult = ""
            For n = LBound(result) To UBound(result)
            arr = Split(result(n), ":")
            strResult = strResult & vbCrLf & result(n)
            If arr(0) = "error" Then
                MsgBox Replace(result(0), "error:", ""), vbExclamation, "Lembar Pengajuan Klaim"
                Exit Sub
            End If
            If UCase(Trim(Right(arr(0), 4))) = "KODE" Then
                mstrNoSJP = arr(1)
                txtKdPPKRujukan.Text = arr(1)
'                Exit Sub
            End If
            If UCase(Trim(Right(arr(0), 4))) = "NAMA" Then
                mstrNoSJP = arr(1)
                txtPpkRujukan.Text = arr(1)
'                Exit Sub
            End If

        Next n

            fraFaskes.Visible = True
            Call fillGridWithFaskes(fgFaskes, result)
        End If
        
        fgFaskes.Visible = True
'        fgFaskes.Left = txtPPKRujukan.Left
'        fgFaskes.Top = 930
    ElseIf KeyAscii = 27 Then
        fgFaskes.Visible = False
        fgFaskes.Clear
    End If
End Sub

Sub fillGridWithFaskes(vFG As MSFlexGrid, vResult() As String)
    With vFG
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .cols = 1
        .rows = 2
'        Call subSetGrid
        Dim row As Integer
        Dim col As Integer
        Dim rows As Integer
        Dim cols As Integer
        Dim i As Integer
        row = 1
        For i = 0 To UBound(vResult)
            Dim arrResult() As String
'            Debug.Assert i <> 9
            arrResult = Split(vResult(i), ":")
            col = isHeaderExist(vFG, arrResult(0))
            If col > -1 Then
                'col = isHeaderExist(vFG, arrResult(0))
                
                If .TextMatrix(row, col) <> "" Then 'KALO TEXTMATRIX TARGET SUDAH ADA ISI BERARTI KITA HARUS TAMBAH
                                                    'ROWS/PINDAH KE BARIS SELANJUTNYA
                                                    'diharapkan kolom pertama adalah kolom yang selalu memiliki nilai
                    .rows = .rows + 1
                    row = .rows - 1
                End If
                .TextMatrix(row, col) = arrResult(1)
            Else 'KALAU ADA KOLOM BARU
                col = .cols - 1
                .TextMatrix(0, col) = arrResult(0) 'BERI HEADER BARU
                .TextMatrix(row, col) = arrResult(1)
                .cols = .cols + 1
                
            End If
        Next i
        .ColWidth(col) = 7000
        .cols = .cols - 1 'MENGHILANGKAN KOLOM YG KELEBIHAN
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


Private Sub lblJalan_DblClick()
tmrJalan.Enabled = False
End Sub

Private Sub lblJalan_Click()
tmrJalan.Enabled = True
End Sub

Private Sub tmrJalan_Timer()
If (lblJalan.Left + lblJalan.Width) <= 0 Then
    lblJalan.Left = Me.Width
End If
lblJalan.Left = lblJalan.Left - 100
End Sub

