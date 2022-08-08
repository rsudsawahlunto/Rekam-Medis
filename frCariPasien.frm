VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCariPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pencarian Data Pasien"
   ClientHeight    =   8520
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   14910
   Visible         =   0   'False
   Begin VB.ComboBox cbojnsPrinter 
      Height          =   330
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame Frame4 
      Height          =   825
      Left            =   0
      TabIndex        =   17
      Top             =   7320
      Width           =   14895
      Begin VB.CommandButton cmdAsuransi 
         Appearance      =   0  'Flat
         Caption         =   "Asuransi Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4680
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdTP 
         Caption         =   "R&iwayat Pemeriksaan"
         Height          =   465
         Left            =   10800
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdHapus 
         Appearance      =   0  'Flat
         Caption         =   "&Hapus Data"
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
         Height          =   465
         Left            =   4680
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdRegRJ 
         Appearance      =   0  'Flat
         Caption         =   "&Registrasi Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6720
         TabIndex        =   10
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
         Height          =   465
         Left            =   8760
         TabIndex        =   11
         ToolTipText     =   "Perbaiki data pasien"
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdKeluar 
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
         Height          =   465
         Left            =   12840
         TabIndex        =   13
         ToolTipText     =   "Tutup aplikasi"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   5295
      Left            =   0
      TabIndex        =   16
      Top             =   2040
      Width           =   14895
      Begin VB.Frame fraDataRiwayatPemeriksaanPasien 
         Caption         =   "Riwayat Data Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   2160
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   10455
         Begin VB.CommandButton cmdCetak 
            Caption         =   "&Cetak"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   7320
            TabIndex        =   26
            ToolTipText     =   "Tutup aplikasi"
            Top             =   2880
            Width           =   1455
         End
         Begin VB.CommandButton cmdClose 
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
            Height          =   345
            Left            =   8880
            TabIndex        =   25
            ToolTipText     =   "Tutup aplikasi"
            Top             =   2880
            Width           =   1455
         End
         Begin MSDataGridLib.DataGrid dgData 
            Height          =   2535
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   4471
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
      Begin MSDataGridLib.DataGrid dgpasien 
         Height          =   4815
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   8493
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   22
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
            Size            =   12
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
   End
   Begin VB.Frame Frame1 
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
      TabIndex        =   14
      Top             =   960
      Width           =   14895
      Begin VB.Frame frtipenama 
         Caption         =   "Pencarian Berdasarkan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   19
         Top             =   150
         Width           =   3975
         Begin VB.OptionButton opt_pnocm 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   0
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton opt_pnama 
            Caption         =   "Nama / Alamat Pasien"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1560
            TabIndex        =   1
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Jenis Kelamin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Left            =   10680
         TabIndex        =   18
         Top             =   195
         Visible         =   0   'False
         Width           =   3015
         Begin VB.OptionButton optlaki2 
            Caption         =   "Pria"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   280
            Width           =   615
         End
         Begin VB.OptionButton optwanita 
            Caption         =   "Wanita"
            Height          =   255
            Left            =   960
            TabIndex        =   5
            Top             =   280
            Width           =   975
         End
         Begin VB.OptionButton optsemua 
            Caption         =   "Semua"
            Height          =   255
            Left            =   2040
            TabIndex        =   6
            Top             =   280
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "&Cari"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   13800
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   525
         Width           =   855
      End
      Begin VB.TextBox txtAlamat 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7200
         TabIndex        =   3
         Top             =   555
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.TextBox txtsearchparameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4320
         TabIndex        =   2
         Top             =   555
         Width           =   2775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Alamat Pasien"
         Height          =   210
         Left            =   7200
         TabIndex        =   28
         Top             =   240
         Width           =   1965
      End
      Begin VB.Label lblNamaNoCM 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Nama Pasien / No. CM"
         Height          =   210
         Left            =   4320
         TabIndex        =   15
         Top             =   240
         Width           =   2640
      End
   End
   Begin MSComctlLib.StatusBar stbInformasi 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   8145
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8784
            MinWidth        =   2646
            Text            =   "Cetak Kartu (F1)"
            TextSave        =   "Cetak Kartu (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   8167
            MinWidth        =   1764
            Text            =   "Blanko Catatan Medis Pasien (F5)"
            TextSave        =   "Blanko Catatan Medis Pasien (F5)"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8678
            Text            =   "Resume Pelayanan Medis Pasien (F8)"
            TextSave        =   "Resume Pelayanan Medis Pasien (F8)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8678
            Text            =   "Cetak Gelang Pasien (F9)"
            TextSave        =   "Cetak Gelang Pasien (F9)"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   22
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
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frCariPasien.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmCariPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsb As New ADODB.recordset
Dim filter As String
Dim sex As String
Dim subPrinterZebra As Printer
Dim subPrinterGelang As Printer
Dim X As String
Dim z As String

Private Sub cmdAsuransi_Click()
On Error GoTo errLoad
    With frmAsuransi
        .Show
        .txtNoCM.Text = dgpasien.Columns("No. CM")
        .txtNamaPeserta.Text = dgpasien.Columns("Nama Lengkap")
        strSQL = "Select TgLLahir from Pasien where NoCM ='" & .txtNoCM.Text & "'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            .dtpTglLahir.value = rs.Fields(0)
        Else
            .dtpTglLahir.value = Now
        End If
        .txtAlamat.Text = dgpasien.Columns("Alamat")
    End With
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    If dgData.ApproxCount = 0 Then Exit Sub
    mstrNoPen = dgData.Columns("NoPendaftaran")
    mstrNoCM = dgData.Columns("NoCM")
    vLaporan = ""
    strSQL = "SELECT * FROM V_DataDetailRiwayatPemeriksaanPasien where NoPendaftaran='" & mstrNoPen & "' AND NoCM='" & mstrNoCM & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
       ' If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
        vLaporan = "view"
        frmCetakDataRiwayatPemeriksaanPasien.Show
    Else
        MsgBox "Riwayat Hasil Pemeriksaan pasien tersebut belum diisi", vbInformation, "Informasi"
        'Exit Sub
    End If
    
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdclose_Click()
    fraDataRiwayatPemeriksaanPasien.Visible = False
    frtipenama.Enabled = True
    txtsearchparameter.Enabled = True
    txtAlamat.Enabled = True
    cmdsearch.Enabled = True
    dgpasien.Enabled = True
    cmdDataPasien.Enabled = True
    cmdHapus.Enabled = True
    cmdRegRJ.Enabled = True
    cmdTP.Enabled = True
    cmdKeluar.Enabled = True
End Sub

Private Sub cmdDataPasien_Click()
    On Error GoTo hell
    strPasien = "Lama"
    mstrNoCM = dgpasien.Columns(0).value
    boltampil = True
    'boltampil = False
    'frmPasienBaru.Show
    AntrianForDataPasien = True
    frmDataPasien2.Show
    Exit Sub
hell:
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    If dgpasien.ApproxCount = 0 Then Exit Sub
    If MsgBox("Yakin akan menghapus data pasien non aktif", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub

    If sp_PasienTemporary(dgpasien.Columns("No. CM").value) = False Then Exit Sub

    MsgBox "No. CM " & dgpasien.Columns("No. CM").value & " berhasil dihapus..", vbInformation, "Informasi"
    Call cmdsearch_Click
    Exit Sub
errHapus:
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdRegRJ_Click()
    On Error GoTo hell
    If dgpasien.ApproxCount = 0 Then Exit Sub
    
    '   @503hendri (2014-11-16)
    '   cek pasien di Antrian IGD
        strSQL = "SELECT NoCM " & _
        " FROM V_DaftarAntrianPasienMRS " & _
        " WHERE (NoCM = '" & dgpasien.Columns(0).value & "') AND (DAY(TglMasuk) = '" & Day(Now) & "') AND (MONTH(TglMasuk) = '" & Month(Now) & "') AND (YEAR(TglMasuk) = '" & Year(Now) & "')" & _
        " AND KdRuangan = '001' AND [Status Periksa]='Belum'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            MsgBox "Pasien tersebut Masih di Antrian IGD", vbInformation, "Informasi"
            Exit Sub
        End If
 
    'cek pasien IGD
    strSQL = "SELECT NoCM FROM V_DaftarPasienIGDAktif WHERE (NoCM = '" & dgpasien.Columns(0).value & "')"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        MsgBox "Pasien tersebut belum keluar dari IGD", vbInformation, "Informasi"
        Exit Sub
    End If

    'cek pasien RI
    strSQL = "SELECT dbo.RegistrasiRI.NoCM, dbo.Ruangan.NamaRuangan FROM dbo.RegistrasiRI INNER JOIN dbo.Ruangan ON dbo.RegistrasiRI.KdRuangan = dbo.Ruangan.KdRuangan WHERE (NoCM = '" & dgpasien.Columns(0).value & "') AND StatusPulang = 'T'"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        MsgBox "Pasien tersebut belum keluar dari Rawat Inap," & vbNewLine & "Ruangan " & rs("NamaRuangan") & "", vbInformation, "Informasi"
        Exit Sub
    End If

    'cek pasien
    strSQL = "SELECT NoCM " & _
    " FROM PasienMasukRumahSakit " & _
    " WHERE (NoCM = '" & dgpasien.Columns(0).value & "') AND (DAY(TglMasuk) = '" & Day(Now) & "') AND (MONTH(TglMasuk) = '" & Month(Now) & "') AND (YEAR(TglMasuk) = '" & Year(Now) & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        If MsgBox("Pasien tersebut sudah terdaftar di Rumah Sakit," & vbNewLine & "Lanjutkan", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    End If
    
    If MsgBox("Pasien mau ke RAWAT JALAN atau PENUNJANG ?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
        strPasien = "Lama"
        mstrNoCM = dgpasien.Columns(0).value
        With frmRegistrasiRJPenunjang
            .Show
            .txtNoCM.Text = Right(mstrNoCM, strBanyakNoCM)
            .txtnocmterm = dgpasien.Columns(1).value
            .txtNamaPasien.Text = dgpasien.Columns(2).value
            If dgpasien.Columns(3).value = "L" Then
                .cboJK.Text = "Laki-Laki"
            ElseIf dgpasien.Columns(3).value = "P" Then
                .cboJK.Text = "Perempuan"
            End If
            .txtThn.Text = dgpasien.Columns(6).value 'tahun
            .txtBln.Text = dgpasien.Columns(7).value
            .txtHr.Text = dgpasien.Columns(8).value
        End With
    Else
        strPasien = "Lama"
        mstrNoCM = Right(dgpasien.Columns(0).value, 6)
        With frmRegistrasiAll
            .Show
            .txtNoCM.Text = mstrNoCM
            .txtnocmterm.Text = dgpasien.Columns(1).value
            .txtNamaPasien.Text = dgpasien.Columns(2).value
            If dgpasien.Columns(3).value = "L" Then
                .cboJK.Text = "Laki-Laki"
            ElseIf dgpasien.Columns(3).value = "P" Then
                .cboJK.Text = "Perempuan"
            End If
            .txtThn.Text = dgpasien.Columns(6).value
            .txtBln.Text = dgpasien.Columns(7).value
            .txtHr.Text = dgpasien.Columns(8).value
        End With
    End If

    Unload Me
    Exit Sub
hell:
End Sub

Public Sub cmdsearch_Click()
    On Error Resume Next
    If opt_pnocm.value = True Then
        If Len(Trim(txtsearchparameter.Text)) = 0 Then
            MsgBox "Parameter pencarian harus diisi", vbExclamation, "Informasi"
            Exit Sub
        Else
            filter = " WHERE [No. CM] like '%" & txtsearchparameter.Text & "%'"
        End If
    ElseIf opt_pnama.value = True Then
        If Len(Trim(txtsearchparameter.Text)) = 0 And Len(Trim(txtAlamat.Text)) = 0 Then
            MsgBox "Parameter pencarian harus diisi", vbExclamation, "Informasi"
            Exit Sub
        Else
            If (txtAlamat.Text = "") And txtsearchparameter.Text <> "" Then
                filter = " WHERE [Nama Lengkap] like '%" & txtsearchparameter.Text & "%' "
            ElseIf (txtAlamat.Text <> "") And (txtsearchparameter.Text = "") Then
                filter = " WHERE  (Alamat like '%" & txtAlamat.Text & "%')"
            Else
                filter = " WHERE [Nama Lengkap] like '%" & txtsearchparameter.Text & "%' AND (Alamat like '%" & txtAlamat.Text & "%')"
            End If
            
        End If
    End If
    Me.MousePointer = vbHourglass
    Set rsb = Nothing
    strSQL = "Select TOP 100 * from v_CariPasien " & filter & sex
    rsb.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    Call subLoadDataPasien
    Me.MousePointer = vbDefault
    If rsb.RecordCount = 0 Then Exit Sub
'    dgpasien.SetFocus
End Sub

Public Sub subLoadCariPasien()
    Set rsb = Nothing
    strSQL = "Select TOP 100 * from v_CariPasien " & filter & sex
    rsb.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    Call subLoadDataPasien
    Me.MousePointer = vbDefault
    If rsb.RecordCount = 0 Then Exit Sub
End Sub

Private Sub cmdTP_Click()
    If dgpasien.ApproxCount = 0 Then Exit Sub
    With frmRiwayatPasien
        .Show
        .txtNoCM.Text = dgpasien.Columns("No. CM")
        .txtNoCM_KeyPress (13)
    End With
End Sub

Private Sub dgData_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgData
    WheelHook.WheelHook dgData
    
    cmdCetak.Enabled = True
End Sub

Private Sub dgpasien_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgpasien
    WheelHook.WheelHook dgpasien
    'Call Form_KeyDown
End Sub

Private Sub dgpasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdRegRJ.SetFocus
End Sub

Private Sub Form_Activate()
    Call centerForm(Me, MDIUtama)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errLoad
    Select Case KeyCode
        Case vbKeyF1
            If dgpasien.ApproxCount = 0 Then Exit Sub
            If IsNull(rsb.Fields("Kota").value) Then
                    dgpasien.Columns(10).value = "-"
                Else
                    dgpasien.Columns(10).value = rsb.Fields("Kota").value
                End If
                
                If IsNull(rsb.Fields("Alamat").value) Then
                    dgpasien.Columns(5).value = "-"
                Else
                    dgpasien.Columns(5).value = rsb.Fields("Alamat").value
                End If
                
'            Call SubPrinterBarcodeNew(dgpasien.Columns(0).value, dgpasien.Columns(1).value, dgpasien.Columns(2).value, dgpasien.Columns(14).value, dgpasien.Columns(4).value, dgpasien.Columns(10).value)
                
'            Call subPrintRegistrasiBarcode(dgpasien.Columns(1).value, dgpasien.Columns(0).value, dgpasien.Columns(4).value, dgpasien.Columns(10).value)

            Call subPrintRegistrasiBarcode
        Case vbKeyF5
            If dgpasien.ApproxCount = 0 Then Exit Sub
            mstrNoCM = dgpasien.Columns(0).value
            frmCetakCatatanMedis.Show
        Case vbKeyF8
            If dgpasien.ApproxCount = 0 Then Exit Sub
            mstrNoCM = dgpasien.Columns(0).value
            strSQL = "SELECT TglMasuk, NoPendaftaran, NoCM, RuanganPerawatan,JenisPasien, KasusPenyakit FROM V_DaftarPasienAll WHERE NoCM='" & mstrNoCM & "' ORDER BY TglMasuk"
            Set dbRst = Nothing
            Call msubRecFO(dbRst, strSQL)
            Set dgData.DataSource = dbRst
            With dgData
                .Columns("TglMasuk").Width = 1590
                .Columns("NoPendaftaran").Width = 1300
                .Columns("NoCM").Width = 800
                .Columns("RuanganPerawatan").Width = 2500
                .Columns("JenisPasien").Width = 1000
                .Columns("KasusPenyakit").Width = 2200
            End With
            frtipenama.Enabled = False
            txtsearchparameter.Enabled = False
            txtAlamat.Enabled = False
            cmdsearch.Enabled = False
            dgpasien.Enabled = False
            cmdDataPasien.Enabled = False
            cmdHapus.Enabled = False
            cmdRegRJ.Enabled = False
            cmdTP.Enabled = False
            fraDataRiwayatPemeriksaanPasien.Visible = True
            cmdCetak.Enabled = False
            cmdKeluar.Enabled = False
        Case vbKeyF9
            
            Call PrintGelangPasien
            
    End Select
    Exit Sub
errLoad:
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    blnCariPasien = True
    opt_pnama.value = True
    mblnCariPasien = True

    strSQL = "Select TOP 100 * from v_CariPasien"
    rsb.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    Call subLoadDataPasien

    If mblnAdmin = True Then cmdHapus.Enabled = True Else cmdHapus.Enabled = False
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsb = Nothing
    blnCariPasien = False
    mblnCariPasien = False
End Sub

Private Sub opt_pnama_Click()
    lblNamaNoCM.Caption = "Masukkan Nama Pasien"
    Label2.Visible = True
    txtAlamat.Visible = True
    txtAlamat.MaxLength = 100
    txtAlamat.Enabled = True

    txtsearchparameter.MaxLength = 50
    txtsearchparameter.Text = ""
    txtAlamat.Text = ""
End Sub

Private Sub opt_pnama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtsearchparameter.SetFocus
End Sub

Private Sub opt_pnocm_Click()
    lblNamaNoCM.Caption = "Masukkan No. CM"
    
    txtAlamat.Text = ""
    Label2.Visible = False
    txtAlamat.Visible = False
'    txtAlamat.Enabled = False

    txtsearchparameter.MaxLength = 12
    txtsearchparameter.Text = ""
End Sub

Private Sub opt_pnocm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtsearchparameter.SetFocus
End Sub

Private Sub optlaki2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdsearch_Click
End Sub

Private Sub OptSemua_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdsearch_Click
End Sub

Private Sub optwanita_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdsearch_Click
End Sub

Private Sub txtAlamat_Change()
    Call cmdsearch_Click
End Sub

Private Sub txtAlamat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdsearch_Click
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtAlamat_LostFocus()
    txtAlamat.Text = StrConv(txtAlamat.Text, vbProperCase)
End Sub

Private Sub txtsearchparameter_Change()
    Call cmdsearch_Click
End Sub

Private Sub txtsearchparameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdsearch_Click
    If opt_pnocm.value = True Then
        Call SetKeyPressToNumber(KeyAscii)
     Else
         Call SetKeyPressToChar(KeyAscii)
    End If
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtsearchparameter_LostFocus()
    txtsearchparameter.Text = StrConv(txtsearchparameter.Text, vbProperCase)
End Sub

'untuk load data pasien
Private Sub subLoadDataPasien()
    Set dgpasien.DataSource = rsb
    With dgpasien
        .Columns(0).Width = 0 '800
        .Columns(1).Width = 1200 '800
        .Columns(2).Width = 3500
        .Columns(3).Width = 400
        .Columns(4).Width = 2000
        .Columns(5).Width = 2000
        .Columns(6).Width = 4000
        .Columns(7).Width = 0
        .Columns(8).Width = 0
        .Columns(9).Width = 0 '1300
        .Columns(10).Width = 0 '1300
        .Columns(11).Width = 0
        .Columns(12).Width = 2500
        .Columns(13).Width = 0
        .Columns(14).Width = 0
        .Columns(15).Width = 0
    End With
End Sub

Private Function sp_PasienTemporary(f_NoCM As String) As Boolean
    On Error GoTo errLoad
    sp_PasienTemporary = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, f_NoCM)
        .Parameters.Append .CreateParameter("TglPembatalan", adDate, adParamInput, , Format(Now, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PasienTemporary"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_PasienTemporary = False
        Else
            Call Add_HistoryLoginActivity("Add_PasienTemporary")
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing

    Exit Function
errLoad:
    sp_PasienTemporary = False
    Call msubPesanError
End Function

Private Sub SubPrinterBarcodeNew(NoCM As String, NamaPasien As String, jk As String, tgllahir As Date, Alamat As String, Kota As String)
Dim result As Long

    result = Shell("C:\Program Files (x86)\PT. Jasamedika Saranatama\Kartu Pasien\KartuPasien.exe export:pdf nocm:" & Chr(34) & txtNoCM.Text & Chr(34) & " namapasien:" & Chr(34) & txtNamaPasien.Text & Chr(34) & " jenisKelamin:" & Chr(34) & txtJenisKelamin.Text & Chr(34) & " tglLahir:" & Chr(34) & txtTanggalLahir.Text & Chr(34) & " alamat:" & Chr(34) & txtAlamat.Text & Chr(34) & " kota:" & Chr(34) & txtKota.Text & Chr(34) & " jenisKartu:0 namaPrinter:" & Chr(34) & "Send To OneNote 2010" & Chr(34), vbNormalFocus)
End Sub

''Private Sub subPrintRegistrasiBarcode()
'Public Sub subPrintRegistrasiBarcode(NamaPasien As String, NoCM As String, Alamat As String, Kota As String)
'    On Error GoTo errLoad
'    Dim PosAwal, PosTamb, Hal As Double
'    Dim mstrNoCMBar As String
'    Dim tmpXY As String
'
''    If dgpasien.ApproxCount = 0 Then Exit Sub
''    Call msubRecFO(rs, "SELECT NamaPrinterBarcode FROM MasterDataPendukung")
''    If IsNull(rs("NamaPrinterBarcode")) Then
''        MsgBox "Nama printer barcode kosong", vbExclamation, "Informasi"
''        Exit Sub
''    End If
'
'    cbojnsPrinter.Clear
'    Dim tempPrint As String
'    tempPrint = ReadINI("Default Printer", "PrinterBarcode", "", "C:\SettingPrinter.ini")
'
''    cbojnsPrinter.clear
''    For Each subPrinterZebra In Printers
''        cbojnsPrinter.AddItem subPrinterZebra.DeviceName
''        If Right(subPrinterZebra.DeviceName, Len(rs("NamaPrinterBarcode"))) = rs("NamaPrinterBarcode") Then '"Zebra P330i USB Card Printer" Then
''            X = rs("NamaPrinterBarcode")
''            Exit For
''        End If
''    Next
'    intLen = Len(dgpasien.Columns(0))
'    For Each subPrinterZebra In Printers
'        Dim printer_temp As String
'        printer_temp = subPrinterZebra.DeviceName
'        If printer_temp = tempPrint Then
'            X = tempPrint
'            Exit For
'        End If
'    Next
'    If X = "" Then Exit Sub
'
'    mstrServerPrinterBarcode = tempPrint 'GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "ServerPrinterBarcode")
'
'    If Len(Trim(mstrServerPrinterBarcode)) = 0 Or LCase(mstrServerPrinterBarcode) = "error" Then
'        frmSetPrinter.Show vbModal
'        Exit Sub
'    End If
'    tmpXY = X
'    X = "\\" & mstrServerPrinterBarcode & "\" & X
'
'    If subPrinterZebra.DeviceName = X Then
'        Set Printer = subPrinterZebra
'    ElseIf subPrinterZebra.DeviceName = tmpXY Then
'        Set Printer = subPrinterZebra
'    Else
'        MsgBox "Printer barcode tidak terdeteksi, harap periksa lagi", vbInformation, "Informasi"
'        Exit Sub
'    End If
'
'    Set Barcode39 = New clsBarCode39
'    PosAwal = 100 'pos awal ???
'    PosTamb = 0
'    Hal = 1
'
''    Printer.CurrentY = PosTamb
''
''    'print nama rs
''    Printer.Print ""
''    Printer.CurrentX = 100
''    Printer.FontName = "Tahoma"
''    Printer.Font.Size = 12
''    Printer.Font.Bold = True
''    Printer.Print ""
''
''    'print jalan rs
'''    Printer.CurrentX = 100
'''    Printer.FontName = "Tahoma"
'''    Printer.Font.Size = 9
'''    Printer.Font.Bold = False
'''    Printer.Print ""
''
''    'print telp
'''    Printer.CurrentX = 100
'''    Printer.FontName = "Tahoma"
'''    Printer.Font.Size = 9
'''    Printer.Font.Bold = False
'''    Printer.Print ""
''
''    Printer.CurrentX = 100
''    Printer.FontName = "Tahoma"
''    Printer.Font.Size = 9
''    Printer.Font.Bold = False
''    Printer.Print ""
''
''    Printer.CurrentX = 100
''    Printer.FontName = "Tahoma"
''    Printer.Font.Size = 9
''    Printer.Font.Bold = False
''    Printer.Print ""
''
''    Printer.CurrentX = 100
''    Printer.FontName = "Tahoma"
''    Printer.Font.Size = 9
''    Printer.Font.Bold = False
''    Printer.Print ""
''
''    'print NamaPasien
''    Printer.CurrentX = 500
''    Printer.Font.Name = "Tahoma"
''    Printer.Font.Size = 10
''    Printer.Font.Bold = True
''    Printer.Print NamaPasien 'dgpasien.Columns("Nama Lengkap").value
''
'''    mstrNoCMBar = Left(dgpasien.Columns("No. CM"), 2) & "-" & Mid(dgpasien.Columns("No. CM"), 3, 2) & "-" & Right(dgpasien.Columns("No. CM"), 2)
''    'add by Bangkit
''    If intLen = 6 Then
''        mstrNoCMBar = Left(NoCM, 2) & "-" & Mid(NoCM, 3, 2) & "-" & Right(NoCM, 2)
''    ElseIf intLen = 10 Then
''         mstrNoCMBar = Left(NoCM, 3) & "-" & Mid(NoCM, 4, 3) & "-" & Right(NoCM, 4)
''    ElseIf intLen = 12 Then
''
''        mstrNoCMBar = Left(NoCM, 4) & "-" & Mid(NoCM, 5, 4) & "-" & Right(NoCM, 4)
''
''    End If
''
''
''    Printer.CurrentX = 500
''    Printer.Font.Name = "Tahoma"
''    Printer.Font.Size = 10
''    Printer.FontBold = False
''    Printer.Print mstrNoCMBar
''
''    Printer.Print ""
''
''    Printer.CurrentX = 500
''    Printer.Font.Name = "Tahoma"
''    Printer.Font.Size = 12
''    Printer.FontBold = False
''    PosTamb = Printer.CurrentY
''
''    With Barcode39
''        .CurrentX = 500 - 150
''        .CurrentY = 2275 'sip
''
''        .NarrowX = 15 'Val(txtNarrowX.Text)
''        .BarcodeHeight = 400 'Val(txtHeight.Text)
''        .ShowBox = 0
''        .Barcode = NoCM 'dgpasien.Columns("No. CM").value
''        If .ErrNumber <> 0 Then
''            MsgBox "Error: It contain invalid barcode charater", vbOKOnly + vbCritical, "Error"
''            Exit Sub
''        End If
''        .Draw Printer
''
''    End With
''    Printer.EndDoc
'    Exit Sub
'errLoad:
'    Call msubPesanError
'    Printer.KillDoc
'End Sub

'@Dimas 2014-05-10
Private Sub subPrintRegistrasiBarcode()

On Error GoTo errLoad
Dim PosAwal, PosTamb, Hal As Double
Dim mstrNoCMBar As String
Dim tmpXY As String
    
    If dgpasien.ApproxCount = 0 Then Exit Sub
    Call msubRecFO(rs, "SELECT NamaPrinterBarcode FROM MasterDataPendukung")
    
    If IsNull(rs("NamaPrinterBarcode")) Then
        MsgBox "Nama printer barcode kosong", vbExclamation, "Informasi"
        Exit Sub
    End If
    
    
    cbojnsPrinter.Clear
    z = "Zebra P330i Card Printer USB"
    For Each subPrinterZebra In Printers
        cbojnsPrinter.AddItem subPrinterZebra.DeviceName
        If Right(subPrinterZebra.DeviceName, Len(rs("NamaPrinterBarcode"))) = rs("NamaPrinterBarcode") Then
            X = rs("NamaPrinterBarcode")
            Exit For
        ElseIf Right(subPrinterZebra.DeviceName, Len(rs("NamaPrinterBarcode"))) = z Then
            X = z
            Exit For
        End If
    Next
    
    If X = "" Then Exit Sub
    
    tmpXY = X
    
    If subPrinterZebra.DeviceName = X Then
        Set Printer = subPrinterZebra
    ElseIf subPrinterZebra.DeviceName = tmpXY Then
        Set Printer = subPrinterZebra
    Else
        MsgBox "Printer barcode tidak terdeteksi, harap periksa lagi", vbInformation, "Informasi"
        Exit Sub
    End If
    
    Set Barcode39 = New clsBarCode39
    PosAwal = 100 'pos awal ???
    PosTamb = 0
    Hal = 1
    
    Printer.CurrentY = PosTamb
    
    'print nama rs
    Printer.Print ""
    Printer.CurrentX = 100
    Printer.FontName = "Tahoma"
    Printer.Font.Size = 12
    Printer.Font.Bold = True
    Printer.Print ""
    
    'print jalan rs
    Printer.CurrentX = 100
    Printer.FontName = "Tahoma"
    Printer.Font.Size = 9
    Printer.Font.Bold = False
    Printer.Print ""
    
    'print telp
    Printer.CurrentX = 100
    Printer.FontName = "Tahoma"
    Printer.Font.Size = 12
    Printer.Font.Bold = False
    Printer.Print ""
    
    Printer.CurrentX = 100
    Printer.FontName = "Tahoma"
    Printer.Font.Size = 10
    Printer.Font.Bold = False
    Printer.Print ""
    
'    Printer.CurrentX = 100
'    Printer.FontName = "Tahoma"
'    Printer.Font.Size = 9
'    Printer.Font.Bold = False
'    Printer.Print ""
    
'    Printer.CurrentX = 100
'    Printer.FontName = "Tahoma"
'    Printer.Font.Size = 9
'    Printer.Font.Bold = False
'    Printer.Print ""
'
'     Printer.Print ""
'    mstrNoCMBar = Left(dgpasien.Columns("No. CM"), 2) & "-" & Mid(dgpasien.Columns("No. CM"), 3, 2) & "-" & Right(dgpasien.Columns("No. CM"), 2)
'    mstrNoCMBar = dgpasien.Columns("No. CM").value
    mstrNoCMBar = dgpasien.Columns("NoCM Baru").value
    Printer.CurrentX = 550
    Printer.Font.Name = "Tahoma"
    Printer.Font.Size = 14
    Printer.FontBold = True
    Printer.Print mstrNoCMBar
    
   
    
    Printer.CurrentX = 300
    Printer.Font.Name = "Tahoma"
    Printer.Font.Size = 12
    Printer.FontBold = False
    PosTamb = Printer.CurrentY
    
    With Barcode39
        .CurrentX = 275  '400 - 150
        .CurrentY = 1450 'sip ' jarak barcode dari atas ke bawah makin dikit makin ke atas
        
        .NarrowX = 12 'Val(txtNarrowX.Text)
        .BarcodeHeight = 400 'Val(txtHeight.Text)
        .ShowBox = 0
        .Barcode = dgpasien.Columns("No. CM").value
        If .ErrNumber <> 0 Then
            MsgBox "Error: It contain invalid barcode charater", vbOKOnly + vbCritical, "Error"
            Exit Sub
        End If
        .Draw Printer
        
    End With
     
    Printer.CurrentX = 500
    Printer.Font.Name = "Tahoma"
    Printer.Font.Size = 5
    
     Printer.Print ""
    'print NamaPasien
    Printer.CurrentX = 300
    Printer.Font.Name = "Tahoma"
    Printer.Font.Size = 8
    Printer.Font.Bold = True
    Printer.Print UCase(dgpasien.Columns("Nama Lengkap").value)
    
    'print Alamat Pasien
    Printer.CurrentX = 300
    Printer.Font.Name = "Tahoma"
    Printer.Font.Size = 5
    'Printer.Font.Bold = True
    Printer.Print dgpasien.Columns("Alamat").value & " " & "KEC. " & dgpasien.Columns("Kecamatan").value
'    Printer.Print dgpasien.Columns("Kelurahan").value & " " & "RT/RW : " & dgpasien.Columns("RT/RW").value & " " & "KEC. " & dgpasien.Columns("Kecamatan").value
    
    'print Kota Pasien
    Printer.CurrentX = 300
    Printer.Font.Name = "Tahoma"
    Printer.Font.Size = 5
    'Printer.Font.Bold = True
    Printer.Print dgpasien.Columns("Kota").value

    Printer.EndDoc
Exit Sub
errLoad:
    Call msubPesanError
    Printer.KillDoc
End Sub


Private Sub PrintGelangPasien()
Dim tempPrint As String

On Error GoTo errLoad
    If dgpasien.ApproxCount = 0 Then Exit Sub
    Set rs = Nothing
    strSQL = "SELECT Title FROM Pasien WHERE NoCM='" & dgpasien.Columns(0).value & "'"
    Call msubRecFO(rs, strSQL)
    
    If rs(0).value = "An." Or rs(0).value = "Bayi" Then
        If dgpasien.Columns(3).value = "L" Then
            tempPrint = ReadINI("Default Printer", "PrinterBarcodeAL", "", "C:\Setting.ini")
        Else
            tempPrint = ReadINI("Default Printer", "PrinterBarcodeAP", "", "C:\Setting.ini")
        End If
    Else
        If dgpasien.Columns(3).value = "L" Then
            tempPrint = ReadINI("Default Printer", "PrinterBarcodeDL", "", "C:\Setting.ini")
        Else
            tempPrint = ReadINI("Default Printer", "PrinterBarcodeDP", "", "C:\Setting.ini")
        End If
    End If
    
    For Each subPrinterGelang In Printers
        If Right(subPrinterGelang.DeviceName, Len(tempPrint)) = tempPrint Then X = tempPrint: Exit For
    Next
    
    If X = "" Then MsgBox "Printer barcode tidak terdeteksi, harap periksa lagi", vbInformation, "Informasi": Exit Sub
    
    If subPrinterGelang.DeviceName = X Then
        Set Printer = subPrinterGelang
    Else
        MsgBox "Printer barcode tidak terdeteksi, harap periksa lagi", vbInformation, "Informasi"
        Exit Sub
    End If
    
    Printer.CurrentY = 150
    
    Printer.FontName = "Tahoma"
'    Printer.FontBold = True
    Printer.FontSize = 10
    
    Printer.CurrentX = 600
    Printer.Print UCase(dgpasien.Columns(2).value)
    Printer.CurrentX = 600
    Printer.Print "TGL " & dgpasien.Columns(4).value
    Printer.CurrentX = 600
    Printer.Print "MR " & dgpasien.Columns(1).value
    
    Printer.EndDoc
Exit Sub
errLoad:
    Call msubPesanError
    Printer.KillDoc
End Sub

