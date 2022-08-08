VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmVerifikasiDataAsuransiPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifist2000 - Verifikasi Data Pemakaian Asuransi Pasien"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerifikasiDataAsuransiPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   11790
   Begin MSDataGridLib.DataGrid dgDaftarPasien 
      Height          =   3495
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6165
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
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   11775
      Begin VB.Frame Frame4 
         Caption         =   "Tanggal Pendaftaran"
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
         Left            =   5160
         TabIndex        =   9
         Top             =   120
         Width           =   6495
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   840
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   1560
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
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
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   126812163
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   4200
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
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
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   126812163
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3840
            TabIndex        =   10
            Top             =   315
            Width           =   255
         End
      End
      Begin VB.Label lblJumData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data 0/0"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   5640
      Width           =   11775
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   6480
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdUbahKelPasien 
         Caption         =   "&Jenis Pasien"
         Height          =   495
         Left            =   8175
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   9885
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   12
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmVerifikasiDataAsuransiPasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9960
      Picture         =   "frmVerifikasiDataAsuransiPasien.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmVerifikasiDataAsuransiPasien.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmVerifikasiDataAsuransiPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cmdCari_Click()
    MousePointer = vbHourglass
    Call subLoadDataPasien
    MousePointer = vbDefault
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo errLoad
    If dgDaftarPasien.ApproxCount = 0 Then
        MsgBox "Tidak ada data", vbExclamation, "Informasi"
        mdTglAwal = dtpAwal.value
        mdTglAkhir = dtpAkhir.value
    Else
        mdTglAwal = dtpAwal.value
        mdTglAkhir = dtpAkhir.value
        vLaporan = ""
        If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
        frmCetakDaftarPasienVerifikasiPemakaianAsuransi.Show
    End If
errLoad:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdUbahKelPasien_Click()
    On Error GoTo hell
    Call subLoadFormJP
    Exit Sub
hell:
End Sub

Private Sub dgDaftarPasien_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDaftarPasien
    WheelHook.WheelHook dgDaftarPasien
End Sub

Private Sub dgDaftarPasien_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    lblJumData.Caption = "Data " & dgDaftarPasien.Bookmark & "/" & dgDaftarPasien.ApproxCount
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

Private Sub Form_Activate()
    cmdCari_Click
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpAwal.value = Format(Now, "dd MMMM yyyy 00:00:00")
    dtpAkhir.value = Format(Now, "dd MMMM yyyy 23:59:59")
    mstrFilter = ""
    Call subLoadDataPasien
End Sub

'untuk load data pasien
Private Sub subLoadDataPasien()
    On Error GoTo errLoad

    strSQL = "SELECT * FROM V_VerifikasiPemakaianAsuransiPasien " & _
    " WHERE TglPendaftaran BETWEEN '" & Format(dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "' "

    Call msubRecFO(rs, strSQL)
    Set dgDaftarPasien.DataSource = rs
    With dgDaftarPasien
        .Columns("NoPendaftaran").Width = 1500
        .Columns("NoCM").Width = 1000
        .Columns("NamaPasien").Width = 2500
        .Columns("TglPendaftaran").Width = 2000
        .Columns("RuanganPelayanan").Width = 1750
        .Columns("JenisPasien").Width = 1500
        .Columns("NamaPenjamin").Width = 2000
        .Columns("IdPenjamin").Width = 0
        .Columns("KdKelompokPasien").Width = 0
        .Columns("JK").Width = 0
        .Columns("Umur").Width = 0
        .Columns("UmurTahun").Width = 0
        .Columns("UmurBulan").Width = 0
        .Columns("UmurHari").Width = 0
    End With
    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnFrmCariPasien = False
End Sub

Private Sub subLoadFormJP()
    On Error GoTo hell
    mstrNoPen = dgDaftarPasien.Columns("NoPendaftaran").value
    strSQL = "SELECT KdKelompokPasien, IdPenjamin, NoCM FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
        mstrNoCM = rs("NoCM").value
    End If

    With frmUbahJenisPasien
        .Show
        .txtNamaFormPengirim.Text = Me.Name
        Call msubRecFO(rs, "SELECT KdInstalasi FROM Ruangan WHERE (NamaRuangan = '" & dgDaftarPasien.Columns("RuanganPelayanan") & "')")
        .txtKdInstalasi.Text = rs("KdInstalasi")
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = dgDaftarPasien.Columns("NamaPasien").value
        If dgDaftarPasien.Columns("JK").value = "P" Then
            .txtJK.Text = "Perempuan"
        Else
            .txtJK.Text = "Laki-laki"
        End If
        .txtThn.Text = dgDaftarPasien.Columns("UmurTahun")
        .txtBln.Text = dgDaftarPasien.Columns("UmurBulan")
        .txtHr.Text = dgDaftarPasien.Columns("UmurHari")
        .lblNoPendaftaran.Visible = False
        .txtNoPendaftaran.Visible = False
        .dcJenisPasien.BoundText = mstrKdJenisPasien
        .dcPenjamin.BoundText = mstrKdPenjaminPasien
    End With
hell:
End Sub

