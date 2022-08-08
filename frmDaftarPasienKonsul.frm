VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDaftarPasienKonsul 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Konsul"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienKonsul.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   14910
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
      Height          =   840
      Left            =   0
      TabIndex        =   7
      Top             =   7200
      Width           =   14895
      Begin VB.CommandButton cmdTransaksiPelayanan 
         Caption         =   "&Transaksi Pelayanan"
         Height          =   450
         Left            =   10320
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   600
         TabIndex        =   4
         Top             =   400
         Width           =   3735
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   450
         Left            =   12600
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukan Nama Pasien /  No.CM / Ruangan"
         Height          =   240
         Index           =   0
         Left            =   600
         TabIndex        =   8
         Top             =   165
         Width           =   3675
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
      TabIndex        =   9
      Top             =   960
      Width           =   14895
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
         TabIndex        =   10
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
            Format          =   107020291
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
            Format          =   127467523
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   11
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgDaftarPasienKonsul 
         Height          =   5175
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   9128
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   19
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
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
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
         TabIndex        =   13
         Top             =   720
         Width           =   1155
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   8070
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   13097
            Text            =   "Cetak Label Konsul (F1)"
            TextSave        =   "Cetak Label Konsul (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   13097
            Text            =   "Refresh Data (F5)"
            TextSave        =   "Refresh Data (F5)"
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
      TabIndex        =   14
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
      Picture         =   "frmDaftarPasienKonsul.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPasienKonsul.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPasienKonsul.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDaftarPasienKonsul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intJumlahPrint As Integer

Private Sub cmdCari_Click()
    On Error GoTo errLoad
    LblJumData.Caption = "Data 0 / 0"
    If dtpAwal.Day <> dtpAkhir.Day Or dtpAwal.Month <> dtpAkhir.Month Or dtpAwal.Year <> dtpAkhir.Year Then
        strSQL = "SELECT TOP 100 [Ruangan Tujuan], [Ruangan Perujuk], TglDirujuk, [No. Urut], NoPendaftaran, NoCM, [Nama Pasien], JK, Umur, Kelas, JenisPasien, KdRuanganTujuan, KdRuanganAsal, UmurTahun, UmurBulan, UmurHari, KdKelas, KdSubInstalasi, [Dokter Perujuk]" & _
        " from V_DaftarPasienKonsul " & _
        " where ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR [Ruangan Tujuan] like '%" & txtParameter.Text & "%') and TglDirujuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'" & _
        " order by [Ruangan Tujuan], TglDirujuk, [No. Urut]"
    Else
        strSQL = "SELECT TOP 100 PERCENT [Ruangan Tujuan], [Ruangan Perujuk], TglDirujuk, [No. Urut], NoPendaftaran, NoCM, [Nama Pasien], JK, Umur, Kelas, JenisPasien, KdRuanganTujuan, KdRuanganAsal, UmurTahun, UmurBulan, UmurHari, KdKelas, KdSubInstalasi, [Dokter Perujuk]" & _
        " from V_DaftarPasienKonsul " & _
        " where ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR [Ruangan Tujuan] like '%" & txtParameter.Text & "%') and TglDirujuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'" & _
        " order by [Ruangan Tujuan], TglDirujuk, [No. Urut]"
    End If
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic

    Set dgDaftarPasienKonsul.DataSource = rs
    Call SetGridAntrianPasien
    LblJumData.Caption = "Data 0 / " & dgDaftarPasienKonsul.ApproxCount
    If dgDaftarPasienKonsul.ApproxCount > 0 Then
        dgDaftarPasienKonsul.SetFocus
    Else
        dtpAwal.SetFocus
    End If
    If mblnAdmin = False Then
        cmdTransaksiPelayanan.Visible = False
    Else
        cmdTransaksiPelayanan.Visible = True
    End If

    Exit Sub
errLoad:
End Sub

Private Sub cmdTransaksiPelayanan_Click()
    Call subLoadFormTP
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgDaftarPasienKonsul_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDaftarPasienKonsul
    WheelHook.WheelHook dgDaftarPasienKonsul
End Sub

Private Sub dgDaftarPasienKonsul_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTransaksiPelayanan.SetFocus
End Sub

Private Sub dgDaftarPasienKonsul_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    LblJumData.Caption = dgDaftarPasienKonsul.Bookmark & " / " & dgDaftarPasienKonsul.ApproxCount & " Data"
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
    Call cmdCari_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errLoad
    Dim strShiftKey As String
    strShiftKey = (Shift + vbShiftMask)
    Select Case KeyCode
        Case vbKeyF1
            If dgDaftarPasienKonsul.ApproxCount = 0 Then Exit Sub
            intJumlahPrint = 0
            If intJumlahPrint = 0 Then
                intJumlahPrint = 1
                mstrNoPen = dgDaftarPasienKonsul.Columns("No. Registrasi").value
                frmCetakStrukKonsuldrDaftarKonsul.Show
            Else
                intJumlahPrint = 0
            End If
        Case vbKeyF4
            If strShiftKey = 2 Then
                strPasien = "Lama"
                mstrNoCM = dgDaftarPasienKonsul.Columns(4).value
                frmPasienBaru.Show
            End If
        Case vbKeyF5
            Call cmdCari_Click
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

    Call cmdCari_Click
End Sub

Sub SetGridAntrianPasien()
    With dgDaftarPasienKonsul
        .Columns(0).Width = 1800
        .Columns(1).Width = 1800
        .Columns(2).Width = 1590
        .Columns(3).Width = 800
        .Columns(4).Caption = "No. Registrasi"
        .Columns(4).Width = 1200
        .Columns(5).Caption = "No .CM"
        .Columns(5).Width = 1500
        .Columns(6).Width = 2000
        .Columns(7).Width = 400
        .Columns(8).Width = 1700
        .Columns(9).Width = 1000
        .Columns(10).Width = 1500
        .Columns("KdRuanganTujuan").Width = 0
        .Columns("KdRuanganAsal").Width = 0

        .Columns("UmurTahun").Width = 0
        .Columns("UmurBulan").Width = 0
        .Columns("UmurHari").Width = 0

        .Columns("KdKelas").Width = 0
        .Columns("KdSubInstalasi").Width = 0
        .Columns("Dokter Perujuk").Width = 0

        .Columns(3).Alignment = dbgCenter
        .Columns(8).Alignment = dbgCenter
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnFormDaftarAntrian = False
End Sub


Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdCari_Click
        txtParameter.SetFocus
    End If
End Sub

Private Sub subLoadFormTP()
    On Error GoTo hell
    mstrNoPen = dgDaftarPasienKonsul.Columns("No. Registrasi").value
    mstrNoCM = dgDaftarPasienKonsul.Columns("No .CM").value

    mstrKdRuanganPasien = dgDaftarPasienKonsul.Columns("KdRuanganTujuan").value 'Kode Ruangan Pasien
    mstrNamaRuanganPasien = dgDaftarPasienKonsul.Columns("Ruangan Tujuan").value 'Nama Ruangan Pasien

    With frmTransaksiPasien
        .Show
        .txtnopendaftaran.Text = dgDaftarPasienKonsul.Columns("No. Registrasi").value
        .txtNoCM.Text = dgDaftarPasienKonsul.Columns("No .CM").value
        .txtNamaPasien.Text = dgDaftarPasienKonsul.Columns("Nama Pasien").value
        If dgDaftarPasienKonsul.Columns("JK").value = "L" Then
            .txtSex.Text = "Laki-Laki"
        Else
            .txtSex.Text = "Perempuan"
        End If
        .txtThn.Text = dgDaftarPasienKonsul.Columns("UmurTahun").value
        .txtBln.Text = dgDaftarPasienKonsul.Columns("UmurBulan").value
        .txtHr.Text = dgDaftarPasienKonsul.Columns("UmurHari").value
        .txtKls.Text = dgDaftarPasienKonsul.Columns("Kelas").value
        .txtJenisPasien.Text = dgDaftarPasienKonsul.Columns("JenisPasien").value
        .txtTglDaftar.Text = dgDaftarPasienKonsul.Columns("Tgldirujuk").value
    End With

    mstrKdKelas = dgDaftarPasienKonsul.Columns("KdKelas").value
    mstrKelas = dgDaftarPasienKonsul.Columns("Kelas").value
    mstrKdSubInstalasi = dgDaftarPasienKonsul.Columns("KdSubInstalasi").value
    mstrNamaDokter = dgDaftarPasienKonsul.Columns("Dokter Perujuk").value

    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If
    Exit Sub
hell:
End Sub

