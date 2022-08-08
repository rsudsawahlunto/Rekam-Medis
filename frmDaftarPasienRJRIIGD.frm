VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDaftarPasienRJRIIGD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienRJRIIGD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   14910
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   8610
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5212
            Text            =   "Cetak Lembar Masuk RI (Ctrl+R)"
            TextSave        =   "Cetak Lembar Masuk RI (Ctrl+R)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5212
            Text            =   "Ubah Penanggung Jawab (Ctrl+U)"
            TextSave        =   "Ubah Penanggung Jawab (Ctrl+U)"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   5212
            Text            =   "Cetak Surat Keterangan (Ctrl+S)"
            TextSave        =   "Cetak Surat Keterangan (Ctrl+S)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   6535
            Text            =   "Cetak Daftar Pasien ( F9 )"
            TextSave        =   "Cetak Daftar Pasien ( F9 )"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5212
            Text            =   "Refresh ( F5 )"
            TextSave        =   "Refresh ( F5 )"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5212
            Text            =   "Cetak SJP ( F9)"
            TextSave        =   "Cetak SJP ( F9)"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5212
            Text            =   "Cetak Label f(11)"
            TextSave        =   "Cetak Label f(11)"
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
   Begin VB.Frame fraCari 
      Caption         =   "Cari Data Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   0
      TabIndex        =   17
      Top             =   7680
      Width           =   14895
      Begin VB.CommandButton cmdUbahRuang 
         Appearance      =   0  'Flat
         Caption         =   "Ubah Ruangan"
         Height          =   495
         Left            =   3840
         TabIndex        =   33
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdAsKep 
         Appearance      =   0  'Flat
         Caption         =   "&Asuhan Keperawatan"
         Height          =   495
         Left            =   3810
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdUbahJenisPasien 
         Appearance      =   0  'Flat
         Caption         =   "Ubah &Jenis Pasien"
         Height          =   495
         Left            =   7710
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdDiagnosa 
         Caption         =   "Periksa Dia&gnosa"
         Height          =   495
         Left            =   5760
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdDataPasien 
         Caption         =   "&Data Pasien"
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
         Left            =   9660
         TabIndex        =   14
         ToolTipText     =   "Perbaiki data pasien"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdTP 
         Caption         =   "&Riwayat Pemeriksaan"
         Height          =   495
         Left            =   11250
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   13200
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   440
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukan Nama Pasien /  No.CM / Ruangan"
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   195
         Width           =   3450
      End
   End
   Begin VB.Frame fraDaftar 
      Caption         =   "Daftar Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   0
      TabIndex        =   18
      Top             =   960
      Width           =   14895
      Begin VB.Frame fraCetakLabel 
         Caption         =   "Jumlah Baris Label"
         Height          =   1335
         Left            =   11400
         TabIndex        =   44
         Top             =   4920
         Visible         =   0   'False
         Width           =   2655
         Begin VB.TextBox txtJml 
            Height          =   375
            Left            =   240
            TabIndex        =   48
            Text            =   "1"
            Top             =   480
            Width           =   975
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   375
            Left            =   1200
            TabIndex        =   47
            Top             =   480
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Value           =   1
            Max             =   100
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Batal"
            Height          =   375
            Left            =   1560
            TabIndex        =   46
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton cmdCetakLabel 
            Caption         =   "Cetak"
            Height          =   375
            Left            =   1560
            TabIndex        =   45
            Top             =   240
            Width           =   855
         End
      End
      Begin MSDataListLib.DataCombo dcKasusPenyakit 
         Height          =   330
         Left            =   840
         TabIndex        =   42
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cetak Daftar Pasien"
         Height          =   3495
         Left            =   4560
         TabIndex        =   34
         Top             =   1680
         Visible         =   0   'False
         Width           =   6015
         Begin VB.Frame Frame3 
            Caption         =   "Tipe Laporan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3135
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   5775
            Begin VB.Frame Frame4 
               Height          =   1215
               Left            =   120
               TabIndex        =   38
               Top             =   840
               Width           =   5535
               Begin VB.OptionButton Option1 
                  Caption         =   "Berdasarkan Wilayah dan Jenis"
                  Height          =   450
                  Left            =   480
                  TabIndex        =   40
                  Top             =   480
                  Value           =   -1  'True
                  Width           =   2175
               End
               Begin VB.OptionButton Option2 
                  Caption         =   "Berdasarkan Wilayah dan Jenis Kelamin"
                  Height          =   450
                  Left            =   3000
                  TabIndex        =   39
                  Top             =   480
                  Width           =   2175
               End
            End
            Begin VB.CommandButton Command1 
               Caption         =   "TUTU&P"
               Height          =   615
               Left            =   2880
               TabIndex        =   37
               Top             =   2280
               Width           =   1575
            End
            Begin VB.CommandButton Command2 
               Caption         =   "&CETAK"
               Height          =   615
               Left            =   1200
               TabIndex        =   36
               Top             =   2280
               Width           =   1575
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "JUMLAH PASIEN RS"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   41
               Top             =   360
               Width           =   5535
            End
         End
      End
      Begin VB.CheckBox chkDiagnosaKosong 
         Caption         =   "Diagnosa Kosong"
         Height          =   255
         Left            =   12240
         TabIndex        =   31
         Top             =   1080
         Width           =   2115
      End
      Begin VB.CheckBox chkPasienSudahPulang 
         Caption         =   "Tampilkan pasien sudah pulang"
         Height          =   255
         Left            =   9120
         TabIndex        =   30
         Top             =   1080
         Width           =   3075
      End
      Begin MSDataListLib.DataCombo dcAsalPasien 
         Height          =   330
         Left            =   6885
         TabIndex        =   6
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Frame Frame1 
         Caption         =   "Periode"
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
         Left            =   9000
         TabIndex        =   19
         Top             =   150
         Width           =   5775
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   158334979
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   156172291
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   21
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataListLib.DataCombo dcJenisPasien 
         Height          =   330
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcKelas 
         Height          =   330
         Left            =   5115
         TabIndex        =   5
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcPenjamin 
         Height          =   330
         Left            =   3000
         TabIndex        =   4
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   330
         Left            =   3000
         TabIndex        =   7
         Top             =   1080
         Width           =   3800
         _ExtentX        =   6694
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   330
         Left            =   6885
         TabIndex        =   8
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid dgDaftarPasienRJ 
         Height          =   4575
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   8070
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         Caption         =   "Kasus Penyakit"
         Height          =   255
         Left            =   840
         TabIndex        =   43
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan"
         Height          =   210
         Index           =   4
         Left            =   6885
         TabIndex        =   29
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instalasi"
         Height          =   210
         Index           =   3
         Left            =   3000
         TabIndex        =   28
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penjamin"
         Height          =   210
         Index           =   2
         Left            =   3000
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Kecamatan"
         Height          =   255
         Left            =   6840
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kelas"
         Height          =   210
         Index           =   1
         Left            =   5160
         TabIndex        =   25
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Pasien"
         Height          =   210
         Index           =   0
         Left            =   840
         TabIndex        =   24
         Top             =   240
         Width           =   960
      End
      Begin VB.Label LblJumData 
         AutoSize        =   -1  'True
         Caption         =   "10 / 100 Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   1155
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   32
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
      Left            =   13080
      Picture         =   "frmDaftarPasienRJRIIGD.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPasienRJRIIGD.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmDaftarPasienRJRIIGD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dTglMasuk As Date
Dim mstrfilterTambahan As String
Dim printLabel As Printer
Dim Barcode39 As clsBarCode39
Dim X As String

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    Call msubDcSource(dcJenisPasien, rs, "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where StatusEnabled='1' order by JenisPasien")
    Call msubDcSource(dcKelas, rs, "SELECT KdKelas, DeskKelas FROM KelasPelayanan where StatusEnabled='1'")
    Call msubDcSource(dcAsalPasien, rs, "Select KdKecamatan, NamaKecamatan From Kecamatan where StatusEnabled='1' order by NamaKecamatan")
    Call msubDcSource(dcPenjamin, rs, " Select IdPenjamin, NamaPenjamin From Penjamin where StatusEnabled='1'")
    Call msubDcSource(dcInstalasi, rs, "Select KdInstalasi, NamaInstalasi from Instalasi where StatusEnabled='1'")
    Call msubDcSource(dcRuangan, rs, "Select KdRuangan, NamaRuangan from Ruangan where StatusEnabled='1'")
    Call msubDcSource(dcKasusPenyakit, rs, "Select kdsubinstalasi, NamaSubInstalasi from SubInstalasi where StatusEnabled='1'")

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdAsKep_Click()
    If dgDaftarPasienRJ.Columns("NoKamar").value = "Null" Then
        MsgBox "Asuhan Keperawatan ini hanya untuk pasien Rawat Inap ", vbInformation, "Validasi"
        Exit Sub
    End If
    With frmAsuhanKeperawatan
        mstrNoPen = dgDaftarPasienRJ.Columns("No. Registrasi")
        mstrNoCM = dgDaftarPasienRJ.Columns("NoCM")
        .txtNoPendaftaran = mstrNoPen
        .txtNoCM = mstrNoCM
        .txtNamaPasien = dgDaftarPasienRJ.Columns("NamaPasien")

        If Left(.txtSex, 1) = "P" Then
            .txtSex.Text = "Perempuan"
        Else
            .txtSex.Text = "Laki-laki"
        End If
        .txtThn = dgDaftarPasienRJ.Columns("UmurTahun")
        .txtBln = dgDaftarPasienRJ.Columns("UmurBulan")
        .txthari = dgDaftarPasienRJ.Columns("UmurHari")
        .Show
    End With
End Sub

Private Sub chkDiagnosaKosong_Click()
    If chkDiagnosaKosong.value = Checked Then
        mstrFilter = " AND dbo.CekDiagnosaUtama(Nopendaftaran,NoCM,KdRuangan,KdSubInstalasi,'KdDiagnosaUtama') = 0"
    Else
        mstrFilter = ""
    End If
End Sub

Public Sub cmdCari_Click()
    On Error GoTo errLoad
    LblJumData.Caption = "0/0"
    MousePointer = vbHourglass

    If chkPasienSudahPulang.value = Unchecked Then
        strSQL = "SELECT NamaInstalasi,RuanganPerawatan, NoPendaftaran, NoCM, NamaPasien, JK, Umur, JenisPasien, NamaPenjamin, Kelas, TglMasuk, TglKeluar, StatusKeluar, KondisiPulang, KasusPenyakit, NoKamar, NoBed, Kecamatan,Kelurahan,Alamat, DokterPemeriksa, UmurTahun, UmurBulan, UmurHari, KdKelas, KdJenisTarif, KdSubInstalasi, KdRuangan, IdDokter,dbo.CekDiagnosaUtama(Nopendaftaran,NoCM,KdRuangan,KdSubInstalasi,'KdDiagnosaUtama') AS Diagnosa,dbo.Ambil_KodeWilayah(Kecamatan,'W') AS Judul, '1' as Jml, '" & Format(dtpAwal, "dd-mm-yyyy") & " s/d " & Format(dtpAkhir, "dd-mm-yyyy") & "' as Periode" & _
        " FROM V_DaftarInfoPasienAll " & _
        " WHERE (NamaPasien like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR RuanganPerawatan like '%" & txtParameter.Text & "%' OR NoKamar like '%" & txtParameter.Text & "%' or Alamat like '%" & txtParameter.Text & "%') and ((TglMasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "')  OR TglKeluar is NULL) and (Kecamatan Like '%" & dcAsalPasien.Text & "%' AND JenisPasien LIKE '%" & dcJenisPasien.Text & "%' and NamaPenjamin like '%" & dcPenjamin.Text & "%' and NamaInstalasi like '%" & dcInstalasi.Text & "%' and RuanganPerawatan like '%" & dcRuangan.Text & "%' AND Kelas LIKE '%" & dcKelas.Text & "%' AND NoPendaftaran LIKE '%%' and KasusPenyakit like '%" & dcKasusPenyakit.Text & "%')" & _
        " " & mstrFilter
    Else
        strSQL = "SELECT NamaInstalasi,RuanganPerawatan, NoPendaftaran, NoCM, NamaPasien, JK, Umur, JenisPasien, NamaPenjamin, Kelas, TglMasuk, TglKeluar, StatusKeluar, KondisiPulang, KasusPenyakit, NoKamar, NoBed, Kecamatan,Kelurahan,Alamat, DokterPemeriksa, UmurTahun, UmurBulan, UmurHari, KdKelas, KdJenisTarif, KdSubInstalasi, KdRuangan, IdDokter,dbo.CekDiagnosaUtama(Nopendaftaran,NoCM,KdRuangan,KdSubInstalasi,'KdDiagnosaUtama') AS Diagnosa,dbo.Ambil_KodeWilayah(Kecamatan,'W') AS Judul, '1' as Jml, '" & Format(dtpAwal, "dd-mm-yyyy") & " s/d " & Format(dtpAkhir, "dd-mm-yyyy") & "' as Periode" & _
        " FROM V_DaftarInfoPasienAll_Pulang " & _
        " WHERE (NamaPasien like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR RuanganPerawatan like '%" & txtParameter.Text & "%' OR NoKamar like '%" & txtParameter.Text & "%' or Alamat like '%" & txtParameter.Text & "%') and ((TglKeluar between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "')) and (Kecamatan Like '%" & dcAsalPasien.Text & "%' AND JenisPasien LIKE '%" & dcJenisPasien.Text & "%' and NamaPenjamin like '%" & dcPenjamin.Text & "%' and NamaInstalasi like '%" & dcInstalasi.Text & "%' and RuanganPerawatan like '%" & dcRuangan.Text & "%' AND Kelas LIKE '%" & dcKelas.Text & "%' AND NoPendaftaran LIKE '%%' and KasusPenyakit like '%" & dcKasusPenyakit.Text & "%')" & _
        " " & mstrFilter
    End If
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
   
    Set dgDaftarPasienRJ.DataSource = rs
    Call SetGridPasienRJ
    LblJumData.Caption = "1 / " & dgDaftarPasienRJ.ApproxCount & " Data"
    MousePointer = vbDefault
    Exit Sub
errLoad:
    MousePointer = vbDefault
End Sub

Private Sub cmdCetakLabel_Click()
Dim i As Integer

For i = 1 To txtJml.Text
    Call printerLabel
Next i

fraCetakLabel.Visible = False
End Sub

Private Sub cmdDataPasien_Click()
    On Error GoTo hell
   
    
    strPasien = "View"
    strRegistrasi = "DaftarPasienRIRJIGD"
'    mstrNoCM = dgDaftarPasienRJ.Columns("NoCM").value
    mstrNoCM = Right(dgDaftarPasienRJ.Columns("NoCM").value, 6)
    'frmPasienBaru.Show
    
    AntrianForDataPasien = True
    frmDataPasien.Show
    Exit Sub
hell:
End Sub

Private Sub cmdDiagnosa_Click()
    On Error GoTo errLoad
    Dim subKdDokterTemp As String

    If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
    Me.Enabled = False
    With frmPeriksaDiagnosa
        .Show
        .txtNamaFormPengirim.Text = Me.Name
        .txtNoPendaftaran = dgDaftarPasienRJ.Columns("No. Registrasi")
        .txtNoCM = Right(dgDaftarPasienRJ.Columns("NoCM"), 6)
        .txtNamaPasien = dgDaftarPasienRJ.Columns("NamaPasien")
        If Trim(dgDaftarPasienRJ.Columns("JK")) = "P" Then
            .txtSex.Text = "Perempuan"
        Else
            .txtSex.Text = "Laki-laki"
        End If
        .txtThn = dgDaftarPasienRJ.Columns("UmurTahun")
        .txtBln = dgDaftarPasienRJ.Columns("UmurBulan")
        .txthari = dgDaftarPasienRJ.Columns("UmurHari")

        mstrKdRuanganPasien = dgDaftarPasienRJ.Columns("KdRuangan")
        mstrKdSubInstalasi = dgDaftarPasienRJ.Columns("KdSubinstalasi")
        subKdDokterTemp = dgDaftarPasienRJ.Columns("IdDokter")
        .txtDokter = dgDaftarPasienRJ.Columns("DokterPemeriksa")
        mstrKdDokter = subKdDokterTemp
        .fraDokter.Visible = False
        .Show
    End With
    Exit Sub
errLoad:
    Call msubPesanError
    frmPeriksaDiagnosa.Show
End Sub

Private Sub cmdLembarMasukKeluar_Click()
    On Error GoTo errLoad
    If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub

    mstrNoPen = dgDaftarPasienRJ.Columns("No. Registrasi")
    strSQL = "SELECT * FROM V_LembarMasukDanKeluarRI WHERE NoPendaftaran = '" & mstrNoPen & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then Exit Sub
    frmCetakLembarMasukDanKeluarV2.Show
    Exit Sub
errLoad:
End Sub

Private Sub cmdTP_Click()
    Call subLoadFormTP
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdUbahJenisPasien_Click()
    On Error GoTo errLoad
    mstrKdRuanganPasien = dgDaftarPasienRJ.Columns("KdRuangan")
    If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
    strSQL = "SELECT KdInstalasi FROM dbo.Ruangan WHERE KdRuangan = '" & dgDaftarPasienRJ.Columns("KdRuangan") & "'"
    Call msubRecFO(rs, strSQL)
    mstrKdInstalasi = rs.Fields("KdInstalasi")
'    If mstrKdInstalasi = "02" Or mstrKdInstalasi = "06" Or mstrKdInstalasi = "11" Then
'        MsgBox "Pasien yang sudah masuk klinik " & vbCrLf & "[Sudah Di Periksa - Sudah Bayar] Tidak Bisa Di Ubah Jenis pasien", vbExclamation, "Validasi"
'        Exit Sub
'    End If
    Call subLoadFormJP
    Exit Sub
errLoad:
End Sub

Private Sub cmdUbahRuang_Click()
With frmEditRuangPelayanan
    .Show
    .txtNoPendaftaran.Text = dgDaftarPasienRJ.Columns(["No. Registrasi"]).value
End With
End Sub

Private Sub Command1_Click()
Frame2.Visible = False
End Sub

Private Sub Command2_Click()
    Call cmdCari_Click
    If Option1.value = True Then
        strCetak = "WilayahJenis"
    Else
        strCetak = "WilayahJekel"
    End If
    frmCetakDaftarPasienS.Show
End Sub



Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
fraCetakLabel.Visible = False
End Sub

Private Sub dcAsalPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcAsalPasien.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "Select KdKecamatan, NamaKecamatan From Kecamatan where StatusEnabled='1' and (NamaKecamatan LIKE '%" & dcAsalPasien.Text & "%')order by NamaKecamatan "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcAsalPasien.Text = ""
            dcInstalasi.SetFocus
            Exit Sub
        End If
        dcAsalPasien.BoundText = rs(0).value
        dcAsalPasien.Text = rs(1).value
    End If
End Sub

Private Sub dcInstalasi_Change()
    dcRuangan.BoundText = ""
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcInstalasi.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "Select KdInstalasi, NamaInstalasi from Instalasi where StatusEnabled='1' and (Namainstalasi LIKE '%" & dcInstalasi.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcInstalasi.Text = ""
            dcRuangan.SetFocus
            Exit Sub
        End If
        dcInstalasi.BoundText = rs(0).value
        dcInstalasi.Text = rs(1).value
    End If
End Sub


Private Sub dcJenisPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcJenisPasien.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where StatusEnabled='1' and (JenisPasien LIKE '%" & dcJenisPasien.Text & "%')order by JenisPasien"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcJenisPasien.Text = ""
            dcPenjamin.SetFocus
            Exit Sub
        End If
        dcJenisPasien.BoundText = rs(0).value
        dcJenisPasien.Text = rs(1).value
    End If
End Sub

Private Sub dcJenisPasien_Change()
    dcPenjamin.BoundText = ""
End Sub
'^_^Pipit Ermita 2015-04-06
Private Sub dcKasusPenyakit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If dcKasusPenyakit.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "SELECT kdsubinstalasi, NamaSubInstalasi FROM SubInstalasi where StatusEnable='1' and (NamaSubInstalasi LIKE '%" & dcKasusPenyakit.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcKasusPenyakit.Text = ""
            dcKasusPenyakit.SetFocus
            Exit Sub
        End If
        dcKasusPenyakit.BoundText = rs(0).value
        dcKasusPenyakit.Text = rs(1).value
    End If
End Sub

Private Sub dcKelas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKelas.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "SELECT KdKelas, DeskKelas FROM KelasPelayanan where StatusEnabled='1' and (DeskKelas LIKE '%" & dcKelas.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcKelas.Text = ""
            dcAsalPasien.SetFocus
            Exit Sub
        End If
        dcKelas.BoundText = rs(0).value
        dcKelas.Text = rs(1).value
    End If
End Sub

Private Sub dcPenjamin_GotFocus()
    Call msubDcSource(dcPenjamin, rs, "select  distinct a.idpenjamin, b.namapenjamin from PenjaminKelompokPasien a " & _
    " inner join Penjamin b on a.idpenjamin = b.idpenjamin " & _
    " inner join KelompokPasien c on a.kdkelompokpasien = c.kdkelompokpasien " & _
    " where   a.kdkelompokpasien like '%" & dcJenisPasien.BoundText & "%' " & _
    " order by b.namapenjamin ")
End Sub

Private Sub dcPenjamin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcPenjamin.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = " Select IdPenjamin, NamaPenjamin From Penjamin where StatusEnabled='1' and (NamaPenjamin LIKE '%" & dcPenjamin.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcPenjamin.Text = ""
            dcKelas.SetFocus
            Exit Sub
        End If
        dcPenjamin.BoundText = rs(0).value
        dcPenjamin.Text = rs(1).value
    End If
End Sub

Private Sub dcRuangan_GotFocus()
    Call msubDcSource(dcRuangan, rs, "select  distinct a.kdinstalasi, b.namaruangan from Instalasi a " & _
    " inner join Ruangan b on a.kdinstalasi = b.kdinstalasi " & _
    " where a.kdinstalasi like '%" & dcInstalasi.BoundText & "%' " & _
    " order by b.namaruangan ")
End Sub

Private Sub dcRuangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcRuangan.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "Select KdRuangan, NamaRuangan from Ruangan where StatusEnabled='1' and (NamaRuangan LIKE '%" & dcRuangan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcRuangan.Text = ""
            cmdCari.SetFocus
            Exit Sub
        End If
        dcRuangan.BoundText = rs(0).value
        dcRuangan.Text = rs(1).value
    End If
End Sub

Private Sub dgDaftarPasienRJ_Click()
fraCetakLabel.Visible = False
End Sub

Private Sub dgDaftarPasienRJ_HeadClick(ByVal ColIndex As Integer)
    Select Case ColIndex
        Case 0
            mstrFilter = " Order By RuanganPerawatan"
        Case 1
            mstrFilter = " Order By NoPendaftaran"
        Case 2
            mstrFilter = " Order By NoCM"
        Case 3
            mstrFilter = " Order By NamaPasien"
        Case 4
            mstrFilter = " Order By JK"
        Case 5
            mstrFilter = " Order By Umur"
        Case 6
            mstrFilter = " Order By JenisPasien"
        Case 7
            mstrFilter = " Order By Kelas"
        Case 8
            mstrFilter = " Order By TglMasuk"
        Case 9
            mstrFilter = " Order By TglKeluar"
        Case Else
            mstrFilter = ""
    End Select
    Call cmdCari_Click
End Sub

Private Sub dgDaftarPasienRJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTP.SetFocus
End Sub

Private Sub dgDaftarPasienRJ_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    LblJumData.Caption = dgDaftarPasienRJ.Bookmark & " / " & dgDaftarPasienRJ.ApproxCount & " Data"
    If dgDaftarPasienRJ.Columns("NoKamar") = "" Then
        cmdUbahRuang.Enabled = False
    Else
        cmdUbahRuang.Enabled = True
    End If
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errLoad
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    
    strSQL = ""
    Select Case KeyCode
        Case vbKeyF1
            Frame2.Visible = True
        Case vbKeyF10
            If dcAsalPasien.Text = "" Then
                Call cmdCari_Click
                frmCtkDaftarPasien.Show 'Buat Semua Kecamatan
            Else
                Call cmdCari_Click
                frmCtkDaftarPasien2.Show 'Buat Kecamatan yang dipilih
            End If
        Case vbKeyR
            If strCtrlKey = 4 Then
                If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
                If dgDaftarPasienRJ.Columns("RuanganPerawatan").value = "Gawat Darurat" Or dgDaftarPasienRJ.Columns("RuanganPerawatan").value = "VK Bersalin" Or Left(dgDaftarPasienRJ.Columns("RuanganPerawatan").value, 4) = "Poli" Then Exit Sub
                mstrNoPen = dgDaftarPasienRJ.Columns("No. Registrasi")
                strSQL = "SELECT * FROM V_LembarMasukDanKeluarRI WHERE NoPendaftaran = '" & mstrNoPen & "'"
                Call msubRecFO(rs, strSQL)
                If rs.EOF = True Then
                    MsgBox "Pasien Sudah Keluar dari Rawat Inap"
                    Exit Sub
                End If
                frmCetakLembarMasukDanKeluarV2.Show
            End If
'        Case vbKeyS
'            If strCtrlKey = 4 Then
'                If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
'                mstrNoPen = dgDaftarPasienRJ.Columns("No. Registrasi")
'                mdTglMasuk = Format(Now, "yyyy/MM/dd")
'                strSQL = "SELECT * FROM V_InfoSuratKeterangan WHERE NoPendaftaran = '" & mstrNoPen & "'"
'                Call msubRecFO(rs, strSQL)
'                If rs.EOF = True Then Exit Sub
'                frmCetakSuratKeterangan.Show
'            Else
'                Me.Caption = KeyCode
'            End If

        Case vbKeyF5
            Call cmdCari_Click
        
        Case vbKeyU
           If strCtrlKey = 4 Then
            Dim strSQLLoadPasien As String
            Dim rsLoadPasien As New ADODB.recordset
            'If strCtrlKey = 4 Then
                mstrNoPen = dgDaftarPasienRJ.Columns("No. Registrasi")

                strSQLLoadPasien = "SELECT * FROM PenanggungJawab WHERE NoPendaftaran = '" & mstrNoPen & "'"
                Set rsLoadPasien = Nothing
                rsLoadPasien.Open strSQLLoadPasien, dbConn, adOpenForwardOnly, adLockReadOnly

               ' If dgDaftarPasienRJ.Columns("RuanganPerawatan").value = "Gawat Darurat" Or dgDaftarPasienRJ.Columns("RuanganPerawatan").value = "VK Bersalin" Or Left(dgDaftarPasienRJ.Columns("RuanganPerawatan").value, 4) = "Poli" Then Exit Sub

                With frmUbahPenanggungJawab
                    .Show
                    .txtNoPendaftaran.Text = mstrNoPen
                    .txtNamaPasien.Text = dgDaftarPasienRJ.Columns("NamaPasien")
                    If dgDaftarPasienRJ.Columns("JK") = "L" Then
                        .txtJK.Text = "Laki-laki"
                    Else
                        .txtJK.Text = "Perempuan"
                    End If
                    .txtNoCM.Text = dgDaftarPasienRJ.Columns("NoCM")
                    .txtNamaRI.Text = rsLoadPasien.Fields("NamaPJ").value
                    If IsNull(rsLoadPasien.Fields("Hubungan").value) Then
                        .dcHubungan.BoundText = ""
                    Else
                        .dcHubungan.BoundText = rsLoadPasien.Fields("Hubungan").value
                    End If
                    .dcPekerjaanPJ.Text = rsLoadPasien.Fields("Pekerjaan").value
                    .txtAlamatRI.Text = rsLoadPasien.Fields("AlamatPJ").value
                    .dcPropinsiPJ.Text = rsLoadPasien.Fields("Propinsi").value
                    .dcKotaPJ.Text = rsLoadPasien.Fields("Kota").value
                    .dcKecamatanPJ.Text = rsLoadPasien.Fields("Kecamatan").value
                    .dcKelurahanPJ.Text = rsLoadPasien.Fields("Kelurahan").value
                    .meRTRWPJ.Text = rsLoadPasien.Fields("RTRW").value
                    .txtKodePos.Text = rsLoadPasien.Fields("KodePos").value
                    .txtTlpRI.Text = rsLoadPasien.Fields("TeleponPJ").value
                End With
            End If
        Case vbKeyF9
            mstrNoPen = dgDaftarPasienRJ.Columns("No. Registrasi")
            strSQL = "select *  from SettingGlobal where Prefix = 'KdKelompokPasienUmum'"

            Call msubRecFO(rsCek, strSQL)
            If rsCek.EOF = False Then
                
                strSQL1 = "SELECT TOP 1 * FROM PemakaianAsuransi where NoPendaftaran = '" & mstrNoPen & "' ORDER BY TglSJP DESC"
                Call msubRecFO(rs1, strSQL1)
                mstrNoSJP = rs1("NoSJP")
                If mstrNoSJP = "" Then
                    MsgBox "No SJP kosong", vbExclamation, "Validasi"
                    Exit Sub
                End If
                vLaporan = "view"
                frmViewerSJP.Show
            End If
        Case vbKeyF11
            txtJml.Text = 1
            UpDown1.value = 1
            fraCetakLabel.Visible = True
                        
    End Select
    Exit Sub
errLoad:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call subLoadDcSource
    dtpAwal.value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.value = Now
    mstrFilter = ""
    If mblnAdmin = False Then
        cmdUbahJenisPasien.Enabled = False
    Else
        cmdUbahJenisPasien.Enabled = True
    End If
    Call cmdCari_Click
    mblnForm = True
    LblJumData.Caption = "0/0"

    If dgDaftarPasienRJ.Columns("NoKamar") = "" Then
        cmdUbahRuang.Enabled = False
    Else
        cmdUbahRuang.Enabled = True
    End If
End Sub

Sub SetGridPasienRJ()
    On Error Resume Next

    Dim i As Integer

    With dgDaftarPasienRJ
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i
        .Columns("RuanganPerawatan").Width = 1700
        .Columns("NoPendaftaran").Width = 1200
        .Columns("NoPendaftaran").Caption = "No. Registrasi"
        .Columns("NoCM").Width = 1800
        .Columns("NamaPasien").Width = 1800
        .Columns("JK").Width = 300
        .Columns("Umur").Width = 1000

        .Columns("JenisPasien").Width = 1600
        .Columns("NamaPenjamin").Width = 1600
        .Columns("Kelas").Width = 1600
        .Columns("TglMasuk").Width = 1590
        .Columns("TglKeluar").Width = 1590
        .Columns("StatusKeluar").Width = 2250
        .Columns("KondisiPulang").Width = 1800
        .Columns("KasusPenyakit").Width = 2100
        .Columns("NoKamar").Width = 1000
        .Columns("NoBed").Width = 700
        .Columns("Kelurahan").Width = 3000
        .Columns("Alamat").Width = 3000
        .Columns("DokterPemeriksa").Width = 2000
'
'        If IsNull(rs("NoKamar").value) = False Then
'           .Columns("NoKamar").value = "BDU"
'        Else
'           .Columns("NoKamar").value = rs("NoKamar").value
'        End If
'         'If IsNull(.Columns(7).value) Then txtTptLhr.Text = "" Else txtTptLhr.Text = .Columns(7).value
     
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnForm = False
End Sub

Private Sub Frame5_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub txtjml_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
        Beep
        KeyAscii = 0
End If
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call cmdCari_Click
        If dgDaftarPasienRJ.ApproxCount = 0 Then
            txtParameter.SetFocus
            Exit Sub
        Else
            cmdDiagnosa.SetFocus
        End If
    End If
    Exit Sub
errLoad:
End Sub

Private Sub subLoadFormTP()
    On Error GoTo hell
    mstrNoPen = dgDaftarPasienRJ.Columns("No. Registrasi").value
    mstrNoCM = dgDaftarPasienRJ.Columns("NoCM").value

    mstrKdRuanganPasien = dgDaftarPasienRJ.Columns("KdRuangan").value 'Kode Ruangan Pasien
    mstrNamaRuanganPasien = dgDaftarPasienRJ.Columns("RuanganPerawatan").value 'Nama Ruangan Pasien

    With frmTransaksiPasien
        .Show
        .txtNoPendaftaran.Text = dgDaftarPasienRJ.Columns("No. Registrasi").value
        .txtNoCM.Text = dgDaftarPasienRJ.Columns("NoCM").value
        .txtNamaPasien.Text = dgDaftarPasienRJ.Columns("NamaPasien").value
        If dgDaftarPasienRJ.Columns("JK").value = "L" Then
            .txtSex.Text = "Laki-Laki"
        Else
            .txtSex.Text = "Perempuan"
        End If
        .txtThn.Text = dgDaftarPasienRJ.Columns("UmurTahun").value
        .txtBln.Text = dgDaftarPasienRJ.Columns("UmurBulan").value
        .txtHr.Text = dgDaftarPasienRJ.Columns("UmurHari").value
        .txtKls.Text = dgDaftarPasienRJ.Columns("Kelas").value
        .txtJenisPasien.Text = dgDaftarPasienRJ.Columns("JenisPasien").value
        .txtTglDaftar.Text = dgDaftarPasienRJ.Columns("TglMasuk").value
    End With

    mdTglMasuk = dgDaftarPasienRJ.Columns("TglMasuk").value
    mstrKdKelas = dgDaftarPasienRJ.Columns("KdKelas").value
    mstrKelas = dgDaftarPasienRJ.Columns("Kelas").value
    'mstrKdRuangan = dgDaftarPasienRJ.Columns("KdRuangan").value
    mstrKdSubInstalasi = dgDaftarPasienRJ.Columns("KdSubInstalasi").value
    mstrKdDokter = dgDaftarPasienRJ.Columns("IdDokter").value
    mstrNamaDokter = dgDaftarPasienRJ.Columns("DokterPemeriksa").value

    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If
    Exit Sub
hell:
End Sub

'untuk load data pasien di form ubah jenis pasien
Private Sub subLoadFormJP()
    On Error GoTo errLoad
    mstrNoPen = dgDaftarPasienRJ.Columns("No. Registrasi").value
    mstrNoCM = dgDaftarPasienRJ.Columns("NoCM").value
    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If
    With frmUbahJenisPasien
        .Show
        .txtNamaFormPengirim.Text = Me.Name
        .txtNoCM.Text = Right(dgDaftarPasienRJ.Columns("NoCM").value, 6)
        .txtNamaPasien.Text = dgDaftarPasienRJ.Columns("NamaPasien").value
        If dgDaftarPasienRJ.Columns("JK").value = "P" Then
            .txtJK.Text = "Perempuan"
        Else
            .txtJK.Text = "Laki-laki"
        End If
        .txtThn.Text = dgDaftarPasienRJ.Columns("UmurTahun").value
        .txtBln.Text = dgDaftarPasienRJ.Columns("UmurBulan").value
        .txtHr.Text = dgDaftarPasienRJ.Columns("UmurHari").value
        .txttglpendaftaran.Text = dgDaftarPasienRJ.Columns("TglMasuk").value
        .lblNoPendaftaran.Visible = False
        .txtNoPendaftaran.Visible = False
        .dcJenisPasien.BoundText = mstrKdJenisPasien
        .dcPenjamin.BoundText = mstrKdPenjaminPasien
        strSQL1 = "SELECT IdAsuransi,NoSJP FROM PemakaianAsuransi Where NoPendaftaran = '" & mstrNoPen & "'"
        Set rs1 = Nothing
        Call msubRecFO(rs1, strSQL1)
        If rs1.RecordCount > 0 Then
        .txtNoKartuPA = rs1(0)
        .txtNoSJP = rs1(1)
        End If
        
        
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub printerLabel()
    
    If dgDaftarPasienRJ.ApproxCount = 0 Then Exit Sub
    Dim tempPrint As String
    Dim KertasLabel As String
    
    tempPrint = ReadINI("Default Printer", "Printer Label", "", "C:\Setting.ini")
    KertasLabel = ReadINI("Default Printer", "Kertas Label", "", "C:\Setting.ini")
    
    For Each printLabel In Printers
        If Right(printLabel.DeviceName, Len(tempPrint)) = tempPrint Then X = tempPrint: Exit For
    Next
    
    If X = "" Then MsgBox "Printer label tidak terdeteksi, harap periksa lagi", vbInformation, "Informasi": Exit Sub
    
    If printLabel.DeviceName = X Then
        Set Printer = printLabel
    Else
        MsgBox "Printer label tidak terdeteksi, harap periksa lagi", vbInformation, "Informasi"
        Exit Sub
    End If
    Printer.Font = "Arial Narrow"
    If KertasLabel = "2" Then
    '    Printer.CurrentY = 0
    '    Printer.FontName = "Tahoma"
        Printer.FontSize = 9
        Printer.FontBold = True
    '
        Printer.CurrentY = 0 + 170
        Printer.CurrentX = 100
        Printer.Print "      RSUD SAWAHLUNTO"
        Printer.CurrentY = 0 + 170
        Printer.CurrentX = 3100
        Printer.Print "      RSUD SAWAHLUNTO"
        
        Printer.CurrentY = 300 + 170
        Printer.CurrentX = 100
        Printer.FontSize = 9
        Printer.FontBold = True
        Printer.Print Left(dgDaftarPasienRJ.Columns("NoCM").value, 8)
    
        Printer.CurrentY = 300 + 170
        Printer.CurrentX = 3100
        Printer.Print Left(dgDaftarPasienRJ.Columns("NoCM").value, 8)
        
        Set Barcode39 = New clsBarCode39
        With Barcode39
            .CurrentX = 0  '400 - 150
            .CurrentY = 400 + 170 'sip ' jarak barcode dari atas ke bawah makin dikit makin ke atas
            
            .NarrowX = 12 'Val(txtNarrowX.Text)
            .BarcodeHeight = 300 'Val(txtHeight.Text)
            .ShowBox = 0
            .Barcode = Right(dgDaftarPasienRJ.Columns("NoCM").value, 6)
            If .ErrNumber <> 0 Then
                MsgBox "Error: It contain invalid barcode charater", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            .Draw Printer
            
        End With
        
        With Barcode39
            .CurrentX = 3000  '400 - 150
            .CurrentY = 400 + 170 'sip ' jarak barcode dari atas ke bawah makin dikit makin ke atas
            
            .NarrowX = 12 'Val(txtNarrowX.Text)
            .BarcodeHeight = 300 'Val(txtHeight.Text)
            .ShowBox = 0
            .Barcode = Right(dgDaftarPasienRJ.Columns("NoCM").value, 6)
            If .ErrNumber <> 0 Then
                MsgBox "Error: It contain invalid barcode charater", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            .Draw Printer
            
        End With
        
        Printer.CurrentY = 850 + 170
        Printer.CurrentX = 100
        Printer.FontSize = 8
        Printer.FontBold = False
        Printer.Print dgDaftarPasienRJ.Columns("NamaPasien").value & " (" & dgDaftarPasienRJ.Columns("JK").value & ")"
    
        Printer.CurrentY = 850 + 170
        Printer.CurrentX = 3100
        Printer.Print dgDaftarPasienRJ.Columns("NamaPasien").value & " (" & dgDaftarPasienRJ.Columns("JK").value & ")"
        
        Set rs1 = Nothing
        strSQL1 = "SELECT TglLahir FROM Pasien WHERE NoCM='" & Right(dgDaftarPasienRJ.Columns("NoCM").value, 6) & "'"
        Call msubRecFO(rs1, strSQL1)
        
        Printer.CurrentY = 1030 + 170
        Printer.CurrentX = 100
        Printer.Print "Tgl. Lahir " & rs1(0).value & " (" & dgDaftarPasienRJ.Columns("UmurTahun").value & " th)"
    
        Printer.CurrentY = 1030 + 170
        Printer.CurrentX = 3100
        Printer.Print "Tgl. Lahir " & rs1(0).value & " (" & dgDaftarPasienRJ.Columns("UmurTahun").value & " th)"
        Printer.EndDoc
        
    Else
    
        Printer.CurrentY = 0 + 180
        Printer.CurrentX = 50
        Printer.FontSize = 7
        Printer.FontBold = True
        Printer.Print Left(dgDaftarPasienRJ.Columns("NoCM").value, 8)
    
        Printer.CurrentY = 0 + 180
        Printer.CurrentX = 2050
        Printer.Print Left(dgDaftarPasienRJ.Columns("NoCM").value, 8)
        
        Printer.CurrentY = 0 + 180
        Printer.CurrentX = 4050
        Printer.Print Left(dgDaftarPasienRJ.Columns("NoCM").value, 8)
        
        Printer.CurrentY = 250 + 140
        Printer.CurrentX = 50
        Printer.FontSize = 6
        Printer.FontBold = True
        Printer.Print dgDaftarPasienRJ.Columns("NamaPasien").value & " (" & dgDaftarPasienRJ.Columns("JK").value & ")"
    
        Printer.CurrentY = 250 + 140
        Printer.CurrentX = 2050
        Printer.Print dgDaftarPasienRJ.Columns("NamaPasien").value & " (" & dgDaftarPasienRJ.Columns("JK").value & ")"
        
        Printer.CurrentY = 250 + 140
        Printer.CurrentX = 4050
        Printer.Print dgDaftarPasienRJ.Columns("NamaPasien").value & " (" & dgDaftarPasienRJ.Columns("JK").value & ")"
        
        Set rs1 = Nothing
        strSQL1 = "SELECT TglLahir FROM Pasien WHERE NoCM='" & Right(dgDaftarPasienRJ.Columns("NoCM").value, 6) & "'"
        Call msubRecFO(rs1, strSQL1)
        
'        Printer.FontBold = False
        Printer.CurrentY = 400 + 140
        Printer.CurrentX = 50
        Printer.Print "Tgl. Lahir " & rs1(0).value & " (" & dgDaftarPasienRJ.Columns("UmurTahun").value & " th)"
    
        Printer.CurrentY = 400 + 140
        Printer.CurrentX = 2050
        Printer.Print "Tgl. Lahir " & rs1(0).value & " (" & dgDaftarPasienRJ.Columns("UmurTahun").value & " th)"
        
        Printer.CurrentY = 400 + 140
        Printer.CurrentX = 4050
        Printer.Print "Tgl. Lahir " & rs1(0).value & " (" & dgDaftarPasienRJ.Columns("UmurTahun").value & " th)"
        
        Set Barcode39 = New clsBarCode39
        
        With Barcode39
            
            .CurrentX = 0  '400 - 150
            .CurrentY = 500 + 100 'sip ' jarak barcode dari atas ke bawah makin dikit makin ke atas
            
            .NarrowX = 12 'Val(txtNarrowX.Text)
            .BarcodeHeight = 150 'Val(txtHeight.Text)
            
            .ShowBox = 0
            .Barcode = Right(dgDaftarPasienRJ.Columns("NoCM").value, 6)
            If .ErrNumber <> 0 Then
                MsgBox "Error: It contain invalid barcode charater", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            .Draw Printer
            
        End With
        
        With Barcode39
            .CurrentX = 2000  '400 - 150
            .CurrentY = 500 + 100 'sip ' jarak barcode dari atas ke bawah makin dikit makin ke atas
            
            .NarrowX = 12 'Val(txtNarrowX.Text)
            .BarcodeHeight = 150 'Val(txtHeight.Text)
           
            .ShowBox = 0
            .Barcode = Right(dgDaftarPasienRJ.Columns("NoCM").value, 6)
            If .ErrNumber <> 0 Then
                MsgBox "Error: It contain invalid barcode charater", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            .Draw Printer
            
        End With
        
        With Barcode39
            .CurrentX = 4000  '400 - 150
            .CurrentY = 500 + 100 'sip ' jarak barcode dari atas ke bawah makin dikit makin ke atas
            
            .NarrowX = 12 'Val(txtNarrowX.Text)
            .BarcodeHeight = 150 'Val(txtHeight.Text)
           
            .ShowBox = 0
            .Barcode = Right(dgDaftarPasienRJ.Columns("NoCM").value, 6)
            If .ErrNumber <> 0 Then
                MsgBox "Error: It contain invalid barcode charater", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            .Draw Printer
            
        End With
        Printer.EndDoc
        End If
End Sub

Private Sub UpDown1_Change()
    txtJml.Text = UpDown1.value
End Sub
