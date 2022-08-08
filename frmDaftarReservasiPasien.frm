VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDaftarReservasiPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Reservasi"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarReservasiPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   15795
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   8145
      Width           =   15795
      _ExtentX        =   27861
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Visible         =   0   'False
            Object.Width           =   14111
            MinWidth        =   14111
            Text            =   "F1 - Cetak Antrian"
            TextSave        =   "F1 - Cetak Antrian"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Visible         =   0   'False
            Object.Width           =   14111
            MinWidth        =   14111
            Text            =   "F9 - Cetak Daftar Pasien Reservasi"
            TextSave        =   "F9 - Cetak Daftar Pasien Reservasi"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      TabIndex        =   5
      Top             =   7200
      Width           =   15735
      Begin VB.CheckBox ChkCari 
         Caption         =   "Ruangan"
         Height          =   240
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdRegRJ 
         Appearance      =   0  'Flat
         Caption         =   "&Registrasi Pasien"
         Height          =   465
         Left            =   12240
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdHapusRegistrasi 
         Caption         =   "&Hapus Data"
         Height          =   450
         Left            =   10440
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2760
         TabIndex        =   2
         Top             =   480
         Width           =   3855
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   450
         Left            =   13980
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo DcRuangan 
         Height          =   360
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.CommandButton CmdUpdateRM 
         Caption         =   "&Update No RM"
         Height          =   450
         Left            =   10440
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukan Nama Pasien /  No. CM / Ruangan"
         Height          =   240
         Index           =   0
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   3735
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
      Height          =   6255
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   15735
      Begin VB.CheckBox chkCetak 
         Caption         =   "Daftar Pasien Cetak Antrian"
         Height          =   375
         Left            =   1920
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   3975
      End
      Begin MSDataGridLib.DataGrid dgCetak 
         Height          =   5055
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   8916
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgDaftarReservasiPasien 
         Height          =   5055
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   8916
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "Pasien Yang Sudah Di Registrasi"
         Height          =   375
         Left            =   6240
         TabIndex        =   20
         Top             =   480
         Width           =   3135
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
         Left            =   9855
         TabIndex        =   8
         Top             =   165
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
            TabIndex        =   1
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   16
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   127401987
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   17
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   127401987
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   9
            Top             =   360
            Width           =   255
         End
      End
      Begin MSDataListLib.DataCombo dcStatusPeriksa 
         Height          =   360
         Left            =   6360
         TabIndex        =   0
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin VB.Label LblJumData 
         AutoSize        =   -1  'True
         Caption         =   "Data 0 / 0"
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
         TabIndex        =   11
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Status Periksa"
         Height          =   240
         Index           =   1
         Left            =   6360
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12480
      Picture         =   "frmDaftarReservasiPasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3315
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarReservasiPasien.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarReservasiPasien.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14055
   End
End
Attribute VB_Name = "frmDaftarReservasiPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intJumlahPrint As Integer
Dim tempPembayaranKe As Integer


'===============================Tambah Pencarian Ruang Untuk melihat antrian (Cyber 30 Mei 2012)========================
Private Sub ChkCari_Click()
    If ChkCari = vbChecked Then
        txtParameter.Enabled = False
        strSQL = "Select * from Ruangan where KdInstalasi = '02' and StatusEnabled='1'"
        Call msubDcSource(dcRuangan, rs, strSQL)
        If rs.EOF = False Then dcRuangan.BoundText = rs(0).value
    Else
        txtParameter.Enabled = True
        dcStatusPeriksa.Enabled = True
        dcRuangan.Text = ""
    End If
End Sub

'------------------------------------------------------------- Edited By DAYZ ------------------------------------------------------------------------

'--------------------------------------------------------------- 4/02/2013 ------------------------------------------------------------------------
Private Sub chkCetak_Click()
If chkCetak.value = vbChecked Then
   dgCetak.Visible = True
   ChkCari.Enabled = False
   CmdUpdateRM.Enabled = False
   cmdRegRJ.Enabled = False
   chkStatus.value = False
   cmdHapusRegistrasi.Enabled = False
   cmdCari.Enabled = True
   chkStatus.Enabled = False

   strSQL = "Select NoAntrian,TglMasuk,NoCM as [No.RM],NamaLengkap,NamaRuangan,Keterangan,NamaOperator,NamaDokter from Reservasi_Temp Where " & _
             "(NamaLengkap like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR NamaRuangan like '%" & txtParameter.Text & "%') " & _
             "And TglMasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' " & _
             "and '" & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "' "
Else
   dgCetak.Visible = False
   ChkCari.Enabled = True
   CmdUpdateRM.Enabled = True
   cmdRegRJ.Enabled = True
   cmdHapusRegistrasi.Enabled = True
   cmdCari.Enabled = True
   chkStatus.Enabled = True
End If

Set rsb = Nothing
rsb.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
Set dgCetak.DataSource = rsb
Call SetGridCetak
lblJumData.Caption = "Data 0 / " & dgCetak.ApproxCount
End Sub
Private Sub chkStatus_Click()
If chkStatus.value = vbChecked Then
    cmdRegRJ.Enabled = False
    'cmdHapusRegistrasi.Enabled = False
    Call cmdCari_Click
Else
     cmdRegRJ.Enabled = True
     'cmdHapusRegistrasi.Enabled = True
     Call cmdCari_Click
End If
End Sub

'------------------------------------------------------------- Edited By DAYZ ------------------------------------------------------------------------

'--------------------------------------------------------------- 23/01/2013 ------------------------------------------------------------------------
Public Sub cmdCari_Click()
On Error GoTo errLoad
    lblJumData.Caption = "Data 0 / 0"

    If chkStatus = vbUnchecked Then
            strSQL = "select NoAntrian, [Tgl Pesan], TglMasuk, NoCM , NamaLengkap, KdRuangan, NamaRuangan, NoTlp, " & _
                     "Keterangan, JenisKelamin, NoReservasi, TglLahir, UmurTahun, IdDokter, NoReservasi,NamaKamar,NoBed,KdInstalasi,NamaDokter,JenisKelamin " & _
                     "from V_DaftarReservasiPasien " & _
                     "where NamaRuangan like '%" & dcRuangan.Text & "%' And (NamaLengkap like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR NamaRuangan like '%" & txtParameter.Text & "%') and TglMasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "' " & _
                     "And StatusDaftar = 'T'  and StatusReservasi = 'Y' " & _
                     "order by NamaRuangan, NoAntrian"
                    
    Else
            strSQL = "select NoAntrian, [Tgl Pesan], TglMasuk, NoCM , NamaLengkap, KdRuangan, NamaRuangan, NoTlp, " & _
                     "Keterangan, JenisKelamin, NoReservasi, TglLahir, UmurTahun, IdDokter, NoReservasi,NamaKamar,NoBed,KdInstalasi,NamaDokter,JenisKelamin " & _
                     "from V_DaftarReservasiPasien " & _
                     "where NamaRuangan like '%" & dcRuangan.Text & "%' And (NamaLengkap like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR NamaRuangan like '%" & txtParameter.Text & "%') and TglMasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "' " & _
                     "And StatusDaftar = 'Y' " & _
                     "order by NamaRuangan, NoAntrian"
    
    End If

    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    
    Set dgDaftarReservasiPasien.DataSource = rs
    Call SetGridAntrianPasien
    lblJumData.Caption = "Data 0 / " & dgDaftarReservasiPasien.ApproxCount
    If dgDaftarReservasiPasien.ApproxCount > 0 Then
        dgDaftarReservasiPasien.SetFocus
    Else
        cmdTutup.SetFocus
    End If
    
Exit Sub
errLoad:
End Sub

Private Sub cmdHapusRegistrasi_Click()
On Error GoTo errLoad
    If (chkStatus.value = vbChecked) Then
        MsgBox "Data Reservasi tidak bisa dihapus karena sudah terdaftar"
        Exit Sub
    End If

    If dgDaftarReservasiPasien.ApproxCount = 0 Then Exit Sub
    
    If dgDaftarReservasiPasien.Columns(16).value = "" Then Exit Sub
    If MsgBox("Anda yakin akan menghapus data Reservasi pasien", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    
    strSQL = "delete ReservasiPasien where Noreservasi='" & dgDaftarReservasiPasien.Columns(14).value & "'"
    
    Call msubRecFO(rs, strSQL)
    

Call cmdCari_Click

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdRegRJ_Click()
On Error GoTo hell
Dim rsKdRuanganReservasi As String
Dim rsIdDokterReservasi As String
Dim strNoKamar As String
Dim strNoBed As String

'If MsgBox("Yakin ingin REGISTRASI PASIEN Ini? " & vbNewLine & "PASIEN ini akan terhapus dari tabel RESERVASI !", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    
    bolPasienReservasi = True
    
   
    
    If dgDaftarReservasiPasien.ApproxCount = 0 Then Exit Sub
     strnoAntrianPasien = dgDaftarReservasiPasien.Columns(0).value
    'cek pasien IGD
    strSQL = "SELECT NoCM FROM V_DaftarPasienIGDAktif WHERE (NoCM = '" & dgDaftarReservasiPasien.Columns(3).value & "')"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        MsgBox "Pasien tersebut belum keluar dari IGD", vbInformation, "Informasi"
        Exit Sub
    End If
    
    'cek pasien RI
    strSQL = "SELECT dbo.RegistrasiRI.NoCM, dbo.Ruangan.NamaRuangan FROM dbo.RegistrasiRI INNER JOIN dbo.Ruangan ON dbo.RegistrasiRI.KdRuangan = dbo.Ruangan.KdRuangan WHERE (NoCM = '" & dgDaftarReservasiPasien.Columns(6).value & "') AND StatusPulang = 'T'"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        MsgBox "Pasien tersebut belum keluar dari Rawat Inap," & vbNewLine & "Ruangan " & rs("NamaRuangan") & "", vbInformation, "Informasi"
        Exit Sub
    End If
    
    strSQL = "SELECT NoCM " & _
        " FROM PasienMasukRumahSakit " & _
        " WHERE (NoCM = '" & dgDaftarReservasiPasien.Columns(3).value & "') AND (DAY(TglMasuk) = '" & Day(Now) & "') AND (MONTH(TglMasuk) = '" & Month(Now) & "') AND (YEAR(TglMasuk) = '" & Year(Now) & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        If MsgBox("Pasien tersebut sudah terdaftar di Rumah Sakit, " & vbNewLine & "Lanjutkan", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Else
'        dbConn.Execute "Update ReservasiPasien Set StatusReservasi='T' Where StatusReservasi ='Y' and Noreservasi = '" & dgDaftarReservasiPasien.Columns("NoReservasi").value & "' "
    End If
    
  
   
    
    If Trim(dgDaftarReservasiPasien.Columns(3)) = "" Then
        strReservasi = "Aktif"
        strNoReservasi = dgDaftarReservasiPasien.Columns(10)
        strKdRuanganReservasi = dgDaftarReservasiPasien.Columns(6)
        strPasien = "Reservasi"
        
        With frmPasienBaru
            .Show
            .txtNamaPasien.Text = dgDaftarReservasiPasien.Columns(4).value
            .txtFormPengirim.Text = Me.Name
            If dgDaftarReservasiPasien.Columns(19) = "01" Then
                If dgDaftarReservasiPasien.Columns(12) < 15 Then
                    .cboNamaDepan.Text = "An."
                ElseIf dgDaftarReservasiPasien.Columns(12) > "16" And dgDaftarReservasiPasien.Columns(12) < "30" Then
                    .cboNamaDepan.Text = "Sdr."
                Else
                    .cboNamaDepan.Text = "Tn."
                End If
                .cboJnsKelaminPasien.Text = "Laki-laki"
            Else
                If dgDaftarReservasiPasien.Columns(12) < 15 Then
                    .cboNamaDepan.Text = "An."
                ElseIf dgDaftarReservasiPasien.Columns(12) > "16" And dgDaftarReservasiPasien.Columns(12) < "30" Then
                    .cboNamaDepan.Text = "Nn."
                Else
                    .cboNamaDepan.Text = "Ny."
                End If
                .cboJnsKelaminPasien.Text = "Perempuan"
            End If
            .meTglLahir = dgDaftarReservasiPasien.Columns(11).value
            .txtTelepon = dgDaftarReservasiPasien.Columns(7).value
            .txtNoReservasi = dgDaftarReservasiPasien.Columns(10).value
            
        End With
    Else
        strReservasi = "Aktif"
'********************** Tambah pemanggilan NoReservasi, KdRuangan untuk menghidari kesalahan data (Cyber 30 sept 2012)********************

        strSQL = "Select NoReservasi, NoCM, NamaLengkap, JenisKelamin, TglLahir, [Tgl Pesan],TglMasuk, " & _
                 "NoTlp, NoAntrian, KdRuangan, UmurTahun, UmurBulan, UmurHari,KdKelas,Kelas,NamaKamar,NoBed,idDokter,NamaDokter " & _
                 "from v_daftarReservasiPasien " & _
                 "Where NoReservasi =  '" & dgDaftarReservasiPasien.Columns(14) & "' and NoCm = '" & dgDaftarReservasiPasien.Columns(3) & "' and KdRuangan = '" & dgDaftarReservasiPasien.Columns(5) & "'"

'********************** Tambah pemanggilan NoReservasi, KdRuangan untuk menghidari kesalahan data (Cyber 30 sept 2012)********************
        Call msubRecFO(rsReservasi, strSQL)

        rsKdRuanganReservasi = rsReservasi(9).value
        strNoReservasi = rsReservasi(0).value
   
   If dgDaftarReservasiPasien.Columns(17).value = "03" Then
        With frmRegistrasiAll
                .Show
                .txtFormPengirim.Text = Me.Name
                .txtKdAntrian.Text = rsReservasi.Fields(8)
                .txtNoCM.Text = rsReservasi.Fields(1).value
                .CariData
                .dcRujukanRI.SetFocus
                .txtNoCM.Enabled = False
                'Add
                .txtNoReservasi.Text = strNoReservasi
                .txtNamaPasien.Text = rsReservasi.Fields(2).value
                .dtpTglPendaftaran.value = dgDaftarReservasiPasien.Columns(2).value
                If rsReservasi(3).value = "01" Then
                    .cboJK.Text = "Laki-laki"
                Else
                    .cboJK.Text = "Perempuan"
                End If
                .txtThn.Text = rsReservasi(10).value
                .txtBln.Text = rsReservasi(11).value
                .txtHr.Text = rsReservasi(12).value
                .dcNoKamarRI.Text = strNoKamar
                .dcNoBedRI.Text = strNoBed
                .txtKdDokter.Text = rsReservasi.Fields("IdDokter")
                .txtDokter.Text = rsReservasi.Fields("NamaDokter")
                .fraDokter.Visible = False
                .dcRujukanRI.SetFocus
                .txtNoCM.Enabled = False
                .dcKelasKamarRI.Text = strNoKamar
                .dcNoBedRI.Text = strNoBed
                'untuk mengisi Instalasi Pemeriksaan
                strSQL = ""
                strSQL = "select dbo.Ruangan.KdInstalasi, dbo.Instalasi.NamaInstalasi from dbo.Ruangan inner join dbo.instalasi " & _
                         "on dbo.Ruangan.KdInstalasi = dbo.Instalasi.KdInstalasi " & _
                         "where dbo.Ruangan.kdRuangan = '" & rsKdRuanganReservasi & "' "
                Call msubDcSource(.dcInstalasi, rsReservasi, strSQL)
                .dcInstalasi.BoundText = rsReservasi(0).value
                    
               
                'untuk mengisi Jenis Kelas Pelayanan
    '            Set rsReservasi = Nothing
                strSQL = ""
                strSQL = "Select distinct KdDetailJenisJasaPelayanan, DetailJenisJasaPelayanan " & _
                         "from V_KelasPelayanan " & _
                         "where KdRuangan = '" & rsKdRuanganReservasi & "' "
                Call msubDcSource(.dcJenisKelas, rsReservasi, strSQL)
                .dcJenisKelas.BoundText = rsReservasi(0).value
    '            .dcJenisKelas.BoundText = rs(1).value
                
                'untuk mengisi kelas pelayanan
                strSQL = ""
                strSQL = "Select KdKelas, Kelas from V_KelasPelayanan " & _
                         "where KdRuangan = '" & rsKdRuanganReservasi & "' "
                Call msubDcSource(.dcKelas, rsReservasi, strSQL)
                .dcKelas.BoundText = rsReservasi(0).value
                
                'untuk mengisi Ruangan Pelayanan
                strSQL = ""
                strSQL = "Select KdRuangan, NamaRuangan from V_KelasPelayanan " & _
                         "where KdRuangan = '" & rsKdRuanganReservasi & "' "
                Call msubDcSource(.dcRuangan, rsReservasi, strSQL)
                .dcRuangan.BoundText = rsReservasi(0).value
                
                'untuk mengisi Ruangan Poli (Cyber 30 Juli 2012)
    '            .CboNamaPoli.Text = dgDaftarReservasiPasien.Columns(8).value
                
                'untuk mengisi RuanganPoli Penyakit
                strSQL = ""
                strSQL = "Select KdSubInstalasi, NamaSubInstalasi from V_RegistrasiALL where Kdruangan = '" & rsKdRuanganReservasi & "' "
                Call msubDcSource(.dcSubInstalasi, rsReservasi, strSQL)
                .dcSubInstalasi.BoundText = rsReservasi(0).value
                
                'untuk mengisi SMF/Kasus Penyakit
                strSQL = ""
                strSQL = "Select KdSubInstalasi, NamaSubInstalasi from V_RegistrasiALL where Kdruangan = '" & rsKdRuanganReservasi & "' "
                Call msubDcSource(.dcSubInstalasi, rsReservasi, strSQL)
                .dcSubInstalasi.BoundText = rsReservasi(0).value
                
                strSQL = ""
                strSQL = "select KdRujukanAsal from RujukanAsal where KdRujukanAsal = '01' "
                Call msubRecFO(rsReservasi, strSQL)
                .dcRujukanRI.BoundText = rsReservasi(0).value
                
                strSQL = ""
                strSQL = "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where KdKelompokPasien = '01' "
                Call msubRecFO(rsReservasi, strSQL)
                .dcKelompokPasien.BoundText = rsReservasi(0).value
                
                strSQL = ""
                strSQL = "Select KodeDokter, NamaDokter from V_DaftarDokter where KodeDokter = '" & rsIdDokterReservasi & "' "
                Call msubRecFO(rsReservasi, strSQL)
                .txtKdDokter.Text = rsReservasi(0).value
                .txtDokter.Text = rsReservasi(1).value
    
    '************************************************* Untuk RI Set Focus ke dcCaraMasuk (Cyber 09 Agust 2012) ***************************************
    
                If .dcInstalasi.BoundText = "02" Or .dcInstalasi.BoundText = "32" Then
                    .cmdSimpan.SetFocus
                Else
                    .dcCaraMasukRI.SetFocus
                End If
    '************************************************* Untuk RI Set Focus ke dcCaraMasuk (Cyber 09 Agust 2012) ***************************************
            
        End With
Else
            With frmRegistrasiRJPenunjang
            .Show
            .txtFormPengirim.Text = Me.Name
            .txtNoCM.Text = rsReservasi.Fields(1).value
            .CariData
            .dcRujukanRI.SetFocus
            .txtNoCM.Enabled = False
            'Add
            .txtNoReservasi.Text = strNoReservasi
            
            .txtNamaPasien.Text = rsReservasi.Fields(2).value
            .dtpTglPendaftaran.value = dgDaftarReservasiPasien.Columns(2).value
            If rsReservasi("JenisKelamin").value = "01" Then
                .cboJK.Text = "Laki-laki"
            Else
                .cboJK.Text = "Perempuan"
            End If
              'untuk mengisi Instalasi Pemeriksaan
                strSQL = ""
                strSQL = "select dbo.Ruangan.KdInstalasi, dbo.Instalasi.NamaInstalasi from dbo.Ruangan inner join dbo.instalasi " & _
                         "on dbo.Ruangan.KdInstalasi = dbo.Instalasi.KdInstalasi " & _
                         "where dbo.Ruangan.kdRuangan = '" & rsKdRuanganReservasi & "' "
                Call msubDcSource(.dcInstalasi, rsReservasi, strSQL)
                .dcInstalasi.BoundText = rsReservasi(0).value
                    
               
                'untuk mengisi Jenis Kelas Pelayanan
    '            Set rsReservasi = Nothing
                strSQL = ""
                strSQL = "Select distinct KdDetailJenisJasaPelayanan, DetailJenisJasaPelayanan " & _
                         "from V_KelasPelayanan " & _
                         "where KdRuangan = '" & rsKdRuanganReservasi & "' "
                Call msubDcSource(.dcJenisKelas, rsReservasi, strSQL)
                .dcJenisKelas.BoundText = rsReservasi(0).value
    '            .dcJenisKelas.BoundText = rs(1).value
                
                'untuk mengisi kelas pelayanan
                strSQL = ""
                strSQL = "Select KdKelas, Kelas from V_KelasPelayanan " & _
                         "where KdRuangan = '" & rsKdRuanganReservasi & "' "
                Call msubDcSource(.dcKelas, rsReservasi, strSQL)
                .dcKelas.BoundText = rsReservasi(0).value
                
                'untuk mengisi Ruangan Pelayanan
                strSQL = ""
                strSQL = "Select KdRuangan, NamaRuangan from V_KelasPelayanan " & _
                         "where KdRuangan = '" & rsKdRuanganReservasi & "' "
                Call msubDcSource(.dcRuangan, rsReservasi, strSQL)
                .dcRuangan.BoundText = rsReservasi(0).value
                
                'untuk mengisi Ruangan Poli (Cyber 30 Juli 2012)
    '            .CboNamaPoli.Text = dgDaftarReservasiPasien.Columns(8).value
                
                'untuk mengisi RuanganPoli Penyakit
                strSQL = ""
                strSQL = "Select KdSubInstalasi, NamaSubInstalasi from V_RegistrasiALL where Kdruangan = '" & rsKdRuanganReservasi & "' "
                Call msubDcSource(.dcSubInstalasi, rsReservasi, strSQL)
                .dcSubInstalasi.BoundText = rsReservasi(0).value
                
                'untuk mengisi SMF/Kasus Penyakit
                strSQL = ""
                strSQL = "Select KdSubInstalasi, NamaSubInstalasi from V_RegistrasiALL where Kdruangan = '" & rsKdRuanganReservasi & "' "
                Call msubDcSource(.dcSubInstalasi, rsReservasi, strSQL)
                .dcSubInstalasi.BoundText = rsReservasi(0).value
            .txtThn.Text = rsReservasi(10).value
            .txtBln.Text = rsReservasi(11).value
            .txtHr.Text = rsReservasi(12).value
'            .dcDokter.Text = rsReservasi(15).value
            .dcKelasKamarRI.BoundText = rsReservasi.Fields("KdKelas")
            .dcKelasKamarRI.Text = rsReservasi.Fields("Kelas")
            .dcNoKamarRI.Text = rsReservasi("NamaKamar").value
            .dcNoBedRI.Text = rsReservasi("NoBed").value
            
            
            
            'untuk mengisi Instalasi Pemeriksaan
            strSQL = ""
            strSQL = "select dbo.Ruangan.KdInstalasi, dbo.Instalasi.NamaInstalasi from dbo.Ruangan inner join dbo.instalasi " & _
                     "on dbo.Ruangan.KdInstalasi = dbo.Instalasi.KdInstalasi " & _
                     "where dbo.Ruangan.kdRuangan = '" & rsKdRuanganReservasi & "' "
            Call msubDcSource(.dcInstalasi, rsReservasi, strSQL)
            .dcInstalasi.BoundText = rsReservasi(0).value
                
           
            'untuk mengisi Jenis Kelas Pelayanan
'            Set rsReservasi = Nothing
            strSQL = ""
            strSQL = "Select distinct KdDetailJenisJasaPelayanan, DetailJenisJasaPelayanan " & _
                     "from V_KelasPelayanan " & _
                     "where KdRuangan = '" & rsKdRuanganReservasi & "' "
            Call msubDcSource(.dcJenisKelas, rsReservasi, strSQL)
            .dcJenisKelas.BoundText = rsReservasi(0).value
'            .dcJenisKelas.BoundText = rs(1).value
            
            'untuk mengisi kelas pelayanan
            strSQL = ""
            strSQL = "Select KdKelas, Kelas from V_KelasPelayanan " & _
                     "where KdRuangan = '" & rsKdRuanganReservasi & "' "
            Call msubDcSource(.dcKelas, rsReservasi, strSQL)
            .dcKelas.BoundText = rsReservasi(0).value
            
            'untuk mengisi Ruangan Pelayanan
            strSQL = ""
            strSQL = "Select KdRuangan, NamaRuangan from V_KelasPelayanan " & _
                     "where KdRuangan = '" & rsKdRuanganReservasi & "' "
            Call msubDcSource(.dcRuangan, rsReservasi, strSQL)
            .dcRuangan.BoundText = rsReservasi(0).value
            
            'untuk mengisi Ruangan Poli (Cyber 30 Juli 2012)
'            .CboNamaPoli.Text = dgDaftarReservasiPasien.Columns(8).value
            
            'untuk mengisi RuanganPoli Penyakit
            strSQL = ""
            strSQL = "Select KdSubInstalasi, NamaSubInstalasi from V_RegistrasiALL where Kdruangan = '" & rsKdRuanganReservasi & "' "
            Call msubDcSource(.dcSubInstalasi, rsReservasi, strSQL)
            .dcSubInstalasi.BoundText = rsReservasi(0).value
            
            'untuk mengisi SMF/Kasus Penyakit
            strSQL = ""
            strSQL = "Select KdSubInstalasi, NamaSubInstalasi from V_RegistrasiALL where Kdruangan = '" & rsKdRuanganReservasi & "' "
            Call msubDcSource(.dcSubInstalasi, rsReservasi, strSQL)
            .dcSubInstalasi.BoundText = rsReservasi(0).value
            
            strSQL = ""
            strSQL = "select KdRujukanAsal from RujukanAsal where KdRujukanAsal = '01' "
            Call msubRecFO(rsReservasi, strSQL)
            .dcRujukanRI.BoundText = rsReservasi(0).value
            
            strSQL = ""
            strSQL = "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where KdKelompokPasien = '01' "
            Call msubRecFO(rsReservasi, strSQL)
            .dcKelompokPasien.BoundText = rsReservasi(0).value
            .dcRujukanRI.SetFocus
            .txtNoCM.Enabled = False
            strSQL = ""
            strSQL = "Select KodeDokter, NamaDokter from V_DaftarDokter where KodeDokter = '" & rsIdDokterReservasi & "' "
            Call msubRecFO(rsReservasi, strSQL)
'            .dcDokter.BoundText = rsReservasi(0).value
'            .dcDokter.Text = rsReservasi(1).value

'************************************************* Untuk RI Set Focus ke dcCaraMasuk (Cyber 09 Agust 2012) ***************************************

            If .dcInstalasi.BoundText = "02" Or .dcInstalasi.BoundText = "32" Then
                .cmdSimpan.SetFocus
            Else
                .dcCaraMasukRI.SetFocus
            End If
'************************************************* Untuk RI Set Focus ke dcCaraMasuk (Cyber 09 Agust 2012) ***************************************
        
    End With

    End If
'            dbConn.Execute "Update ReservasiPasien Set StatusReservasi='T' Where StatusReservasi ='Y' and Noreservasi = '" & dgDaftarReservasiPasien.Columns("NoReservasi").value & "' "

End If
Exit Sub
hell:
End Sub

Private Sub cmdTutup_Click()
    bolPasienReservasi = False
    Unload Me
End Sub

Private Sub dcRuangan_Change()
    Call cmdCari_Click
End Sub

Private Sub dcRuangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcRuangan.MatchedWithList = True Then dtpAwal.SetFocus
        strSQL = "Select kdruangan, NamaRuangan From Ruangan Where StatusEnabled='1' and (NamaRuangan LIKE '%" & dcRuangan.Text & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcRuangan.Text = ""
            Exit Sub
        End If
        dcRuangan.BoundText = rs(0).value
        dcRuangan.Text = rs(1).value
    End If
End Sub

Private Sub dgDaftarReservasiPasien_KeyPress(KeyAscii As Integer)
 '   If KeyAscii = 13 Then cmdRegRJ.SetFocus
End Sub

Private Sub dgDaftarReservasiPasien_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    lblJumData.Caption = "Data " & dgDaftarReservasiPasien.Bookmark & " / " & dgDaftarReservasiPasien.ApproxCount
End Sub

Private Sub dtpAkhir_Change()
'    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
'    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Activate()
'    Call cmdCari_Click
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errLoad
Select Case KeyCode
        Case vbKeyF1
            If chkCetak.value = vbChecked Then
                MsgBox "Data tersebut tidak dapat di cetak", vbInformation, "Medifirst2000"
                chkCetak.SetFocus
            Else
                strSQL = "Select NoCM,NamaLengkap,NoAntrian,NamaRuangan " & _
                         "From Reservasi_Temp where NoCM = '" & dgDaftarReservasiPasien.Columns(3).value & "' " & _
                         "And NamaLengkap = '" & dgDaftarReservasiPasien.Columns(4).value & "' " & _
                         "And NoAntrian = '" & dgDaftarReservasiPasien.Columns(0).value & "' " & _
                         "And NamaRuangan = '" & dgDaftarReservasiPasien.Columns(7).value & "' "
                Call msubRecFO(rs, strSQL)
                If Not rs.EOF Then
                     MsgBox "Data Tersebut Sudah Pernah Di Cetak", vbInformation, "Medifirst2000"
                Else
                    frmCetakAntrianReservasi.Show
                End If
            End If
        Case vbKeyF9
                frmCetakDaftarPasienReservasi.Show
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
    
    mblnFormDaftarAntrian = True
    
    dtpAwal.value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.value = Now
'    dcStatusPeriksa.BoundText = ""
    
    If mblnAdmin = True Or mblnVerifikator = True Then
        cmdHapusRegistrasi.Enabled = True
        CmdUpdateRM.Enabled = True
    Else
        cmdHapusRegistrasi.Enabled = False
        CmdUpdateRM.Enabled = False
    End If
    
    Call subLoadDcSource
    Call cmdCari_Click
   
End Sub

Private Sub subLoadDcSource()
On Error GoTo errLoad
    strSQL = "Select * From StatusPeriksaPasien Where StatusEnabled='1'"
    Call msubDcSource(dcStatusPeriksa, rs, strSQL)
    If rs.EOF = False Then dcStatusPeriksa.BoundText = rs(0).value
Exit Sub
errLoad:
    Call msubPesanError
End Sub
Sub SetGridCetak()
    With dgCetak
 
 
  .Columns(0).Width = 1200
  .Columns(0).Caption = "No. Antrian"
  .Columns(0).Alignment = dbgCenter
  
  .Columns(1).Width = 2000
  .Columns(1).Caption = "Tgl Pesan"
  
  .Columns(2).Width = 900
  .Columns(2).Caption = "No. Rekam Medis"

  .Columns(3).Width = 3000
  .Columns(3).Caption = "Nama Pasien"

'  .Columns(4).Width = 1300
'  .Columns(4).Caption = "Nama Poli"

  .Columns(4).Width = 1500
  .Columns(4).Caption = "Nama Ruangan"

  
  .Columns(5).Width = 2500
  .Columns(5).Caption = "Keterangan"

  .Columns(6).Width = 0
  .Columns(6).Caption = "Nama Operator"
  
  
  .Columns(7).Width = 2750
  .Columns(7).Caption = "Nama Dokter"
  
 End With
End Sub

Sub SetGridAntrianPasien()
 With dgDaftarReservasiPasien
 
 
  .Columns(0).Width = 1200
  .Columns(0).Caption = "No. Antrian"
  .Columns(0).Alignment = dbgCenter
  
  .Columns(1).Width = 2000
  .Columns(1).Caption = "Tgl Pesan"
  
  .Columns(2).Width = 2000
  .Columns(2).Caption = "Tgl Pemeriksaan"

  .Columns(3).Width = 1500
  .Columns(3).Caption = "No.CM"

  .Columns(4).Width = 3000
  .Columns(4).Caption = "Nama Pasien"


  .Columns(5).Width = 0
  .Columns(5).Caption = "KdRuangan"
  
  .Columns(6).Width = 1500
  .Columns(6).Caption = "Ruangan"

  
'  .Columns(8).Width = 2500
'  .Columns(8).Caption = "Alamat"
  
  
  .Columns(7).Width = 1500
  .Columns(7).Caption = "No Telp"
  
  .Columns(8).Width = 2000
  .Columns(8).Caption = "Keterangan"
  
  .Columns(9).Width = 0
  .Columns(10).Width = 0
  .Columns(11).Width = 0
  .Columns(12).Width = 0
  .Columns(13).Width = 0
  .Columns(14).Width = 0
  .Columns(15).Width = 1200
  .Columns(15).Caption = "No. Kamar"
  .Columns(16).Width = 1200
  .Columns(16).Caption = "No. Bed"
  .Columns(17).Width = 0
  .Columns(17).Caption = "Kode Instalasi"
  .Columns(18).Width = 0
  .Columns(18).Caption = "Nama Dokter"
  .Columns(19).Width = 0

  

  
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnFormDaftarAntrian = False
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdCari_Click
'        If dgDaftarReservasiPasien.ApproxCount > 0 Then
'            dgDaftarReservasiPasien.SetFocus
'        Else
        txtParameter.SetFocus
'        End If
    End If
End Sub

'untuk load data pasien di form ubah jenis pasien
'Private Sub subLoadFormJP()
'
'    mstrNoPen = dgDaftarReservasiPasien.Columns("No. Registrasi").value
'    mstrNoCM = dgDaftarReservasiPasien.Columns("No .CM").value
'    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
'    Call msubRecFO(rs, strSQL)
'    If rs.EOF = False Then
'        mstrKdJenisPasien = rs("KdKelompokPasien").value
'        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
'    End If
'
'    With frmUbahJenisPasien
'        .Show
'        .txtNamaFormPengirim.Text = Me.Name
'        .txtNoCM.Text = dgDaftarReservasiPasien.Columns("No .CM").value
'        .txtNamaPasien.Text = dgDaftarReservasiPasien.Columns("Nama Pasien").value
'        If dgDaftarReservasiPasien.Columns("JK").value = "P" Then
'            .txtJK.Text = "Perempuan"
'        Else
'            .txtJK.Text = "Laki-laki"
'        End If
'        .txtThn.Text = dgDaftarReservasiPasien.Columns("UmurTahun").value
'        .txtBln.Text = dgDaftarReservasiPasien.Columns("UmurBulan").value
'        .txtHr.Text = dgDaftarReservasiPasien.Columns("UmurHari").value
'        .txtTglPendaftaran.Text = dgDaftarReservasiPasien.Columns("TglMasuk").value
'        .lblNoPendaftaran.Visible = False
'        .txtNoPendaftaran.Visible = False
'        .dcJenisPasien.BoundText = mstrKdJenisPasien
'        .dcPenjamin.BoundText = mstrKdPenjaminPasien
'    End With
'End Sub
