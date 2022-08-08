VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmKondisiBarangNM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kondisi Barang Non Medis"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKondisiBarangNM.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   7230
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
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
      Left            =   1680
      TabIndex        =   8
      Top             =   7440
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dgCariBarang 
      Height          =   2535
      Left            =   360
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   15
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
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   0
      TabIndex        =   17
      Top             =   2760
      Width           =   7215
      Begin VB.TextBox txtCariBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   7
         Top             =   3960
         Width           =   3240
      End
      Begin MSDataGridLib.DataGrid dgKondisiBarang 
         Height          =   3495
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6165
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   15
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
            Locked          =   -1  'True
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Cari Barang"
         Height          =   210
         Index           =   6
         Left            =   255
         TabIndex        =   19
         Top             =   4005
         Width           =   900
      End
      Begin VB.Label lblJmlData 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Jumlah Barang"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   5865
         TabIndex        =   18
         Top             =   4020
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
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
      Left            =   2760
      TabIndex        =   9
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
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
      Left            =   3855
      TabIndex        =   10
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
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
      Left            =   4920
      TabIndex        =   11
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
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
      Left            =   6000
      TabIndex        =   12
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Frame fraBarang 
      Height          =   1695
      Left            =   0
      TabIndex        =   13
      Top             =   960
      Width           =   7215
      Begin VB.TextBox txtNoReg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   24
         Top             =   360
         Width           =   2400
      End
      Begin VB.TextBox txtStok 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   4
         Top             =   1080
         Width           =   1320
      End
      Begin VB.TextBox txtKdBarang 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   20
         Text            =   "txtkdbarang"
         Top             =   0
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.TextBox txtNamaBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Top             =   360
         Width           =   2880
      End
      Begin MSDataListLib.DataCombo dcAsalBarang 
         Height          =   330
         Left            =   1560
         TabIndex        =   2
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcKondisiBarang 
         Height          =   330
         Left            =   4680
         TabIndex        =   3
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtJmlBarang 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4680
         MaxLength       =   25
         TabIndex        =   5
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Stok Ruangan"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Kondisi"
         Height          =   210
         Index           =   8
         Left            =   3840
         TabIndex        =   21
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nama Barang"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Asal Barang"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   930
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah"
         Height          =   210
         Index           =   3
         Left            =   3840
         TabIndex        =   14
         Top             =   1080
         Width           =   555
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   23
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
      Left            =   5400
      Picture         =   "frmKondisiBarangNM.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmKondisiBarangNM.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmKondisiBarangNM.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmKondisiBarangNM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tempbolTampil As Boolean
Dim NoRegister As String

Private Sub cmdBatal_Click()
On Error GoTo Errload

    Call subKosong
    Call subLoadDcSource
    Call subLoadGridSource
    txtNamaBarang.SetFocus

Exit Sub
Errload:
End Sub

Private Sub cmdCetak_Click()
On Error Resume Next
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakKondisiBarang.Show
End Sub

Private Sub cmdHapus_Click()
On Error GoTo Errload

    If txtKdBarang.Text = "" Then
        MsgBox "Nama barang kosong", vbExclamation, "Validasi": txtNamaBarang.SetFocus: Exit Sub
    End If
    
    If MsgBox("Anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    dbConn.Execute "DELETE KondisiBarangNonMedis WHERE KdRuangan = '" & mstrKdRuangan & "' AND KdBarang ='" & txtKdBarang.Text & "' AND KdAsal='" & dcAsalBarang.BoundText & "' AND KdKondisi='" & dcKondisiBarang.BoundText & "' AND JmlBarang='" & txtJmlBarang.Text & "' AND NoRegisterAsset ='" & NoRegister & "' "
    Call cmdBatal_Click
    MsgBox "Penghapusan data berhasil", vbInformation, "Informasi"

Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo Errload

    If txtNamaBarang.Text = "" Then
        MsgBox "Nama barang kosong", vbExclamation, "Validasi": txtNamaBarang.SetFocus: Exit Sub
    ElseIf txtJmlBarang.Text = "0" Then
      MsgBox "Jumlah barang Tidak boleh Kosong", vbExclamation, "Validasi": txtJmlBarang.SetFocus: Exit Sub
    
    End If
    If Periksa("datacombo", dcAsalBarang, "Asal barang kosong") = False Then Exit Sub
    If Periksa("datacombo", dcKondisiBarang, "Kondisi barang kosong") = False Then Exit Sub
    
    If CDbl(txtJmlBarang.Text) > CDbl(txtStok.Text) Then
        MsgBox "Jumlah Barang tidak boleh melebihi Stok Ruangan", vbCritical, "Validasi"
        Exit Sub
    End If
    
    If sp_KondisiBarangNonMedis() = False Then Exit Sub
    MsgBox "Penyimpanan Kondisi Barang Berhasil", vbInformation, "Informasi"
    Call cmdBatal_Click

Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcAsalBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If txtStok.Enabled = False Then txtJmlBarang.SetFocus
End Sub

Private Sub dcBahanBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcKondisiBarang.SetFocus
End Sub

Private Sub dcKondisiBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJmlBarang.SetFocus
End Sub

Private Sub dgCariBarang_Click()
WheelHook.WheelUnHook
        Set MyProperty = dgCariBarang
        WheelHook.WheelHook dgCariBarang
End Sub

Private Sub dgCariBarang_DblClick()
On Error GoTo Errload

    With dgCariBarang
        If .ApproxCount = 0 Then Exit Sub
       txtStok.Text = .Columns("jmlStok")
        
        dcAsalBarang.BoundText = .Columns("KdAsal")
        txtKdBarang.Text = .Columns("KdBarang")
        txtNoReg.Text = .Columns("NoRegisterAsset")
        NoRegister = .Columns("NoRegisterAsset")
        
        txtNamaBarang.Text = .Columns("Nama Barang")
        
        
        
        .Visible = False
    End With
        
'    strSQL = "SELECT JmlStok FROM V_InfoStokGudangUmumFIFO WHERE KdRuangan = '" & mstrKdRuangan & "' AND KdBarang = '" & txtKdBarang.Text & "' AND KdAsal = '" & dcAsalBarang.BoundText & "' "
'    Call msubRecFO(rs, strSQL)
'    If rs.EOF = False Then txtStok.Text = rs(0).Value Else txtStok.Text = 0
'
    dcKondisiBarang.SetFocus

Exit Sub
Errload:
End Sub

Private Sub dgCariBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgCariBarang_DblClick
End Sub

Private Sub dgKondisiBarang_Click()
WheelHook.WheelUnHook
        Set MyProperty = dgKondisiBarang
        WheelHook.WheelHook dgKondisiBarang
End Sub

Private Sub dgKondisiBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaBarang.SetFocus
End Sub

Private Sub dgKondisiBarang_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Errload

    With dgKondisiBarang
        If .ApproxCount = 0 Then Exit Sub
        txtKdBarang.Text = .Columns("KdBarang")
        
        dcAsalBarang.BoundText = .Columns("KdAsal")
        dcKondisiBarang.BoundText = .Columns("KdKondisi")
        txtJmlBarang.Text = .Columns("JmlBarang")
        Call msubRecFO(rs, "select dbo.FB_TakeStokBrgNonMedis('" & mstrKdRuangan & "', '" & txtKdBarang & "','" & dcAsalBarang.BoundText & "') as stok")
        txtStok.Text = IIf(IsNull(rs(0)), 0, rs(0))
        txtNoReg.Text = .Columns("NoRegisterAsset")
        txtNamaBarang.Text = .Columns("Nama Barang")
    End With
    dgCariBarang.Visible = False
    lblJmlData.Caption = dgKondisiBarang.Bookmark & " / " & dgKondisiBarang.ApproxCount & " Data"

Exit Sub
Errload:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo Errload
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call subKosong
    Call subLoadDcSource
    Call subLoadGridSource
Exit Sub
Errload:
End Sub

Private Sub txtCariBarang_Change()
On Error GoTo Errload

    Call subLoadGridSource

Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub txtLokasi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then dgKondisiBarang.SetFocus
End Sub

Private Sub txtJmlBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtJmlBarang_LostFocus()
    txtJmlBarang.Text = IIf(Val(txtJmlBarang) = 0, 0, Format(txtJmlBarang, "#,###"))
End Sub

Private Sub txtLokasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtNamaBarang_Change()
On Error GoTo Errload
    
    If tempbolTampil = True Then Exit Sub
    Call subCariBarang

Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub txtNamaBarang_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyDown Then If dgCariBarang.Visible = True Then dgCariBarang.SetFocus
End Sub

Private Sub txtNamaBarang_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then If dgCariBarang.Visible = True Then dgCariBarang.SetFocus Else dcAsalBarang.SetFocus
    If KeyAscii = 27 Then If dgCariBarang.Visible = True Then dgCariBarang.Visible = False
End Sub

Private Sub subKosong()
    txtKdBarang.Text = ""
    txtNamaBarang.Text = ""
    txtCariBarang.Text = ""
    dcAsalBarang.BoundText = ""
    dcKondisiBarang.BoundText = ""
    txtNoReg.Text = ""
    txtJmlBarang.Text = 0
    txtStok.Text = 0
    dgCariBarang.Visible = False
End Sub

Private Sub subLoadDcSource()
On Error GoTo Errload

    Call msubDcSource(dcAsalBarang, rs, "SELECT KdAsal, NamaAsal FROM AsalBarang where StatusEnabled='1' ORDER BY NamaAsal")
    Call msubDcSource(dcKondisiBarang, rs, "SELECT KdKondisi, Kondisi FROM KondisiBarang where StatusEnabled='1' ORDER BY Kondisi")

Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub subCariBarang()
On Error GoTo Errload

    strsql = "SELECT  distinct [Nama Barang],AsalBarang,NoRegisterAsset , jmlStok,DetailJenisBarang AS [Jenis Barang], KdBarang, KdAsal,Satuan FROM V_CariBarangNonMedisx " & _
        " WHERE kdruangan='" & mstrKdRuangan & "' AND [Nama Barang] LIKE '%" & txtNamaBarang.Text & "%' " & _
        " ORDER BY [Nama Barang]"
        
'    strsql = "SELECT  distinct [Nama Barang],AsalBarang,NoRegisterAsset , jmlStok,DetailJenisBarang AS [Jenis Barang], KdBarang, KdAsal,Satuan FROM V_CariBarangNonMedis " & _
'        " WHERE ((NoRegisterAsset <> '0000000' AND NoRegisterAsset <> '0' AND NoRegisterAsset <>'000000') OR KdJenisAset is null) And kdruangan='" & mstrKdRuangan & "' AND [Nama Barang] LIKE '%" & txtNamaBarang.Text & "%' " & _
'        " ORDER BY [Nama Barang]"
    
    Call msubRecFO(rs, strsql)
    Set dgCariBarang.DataSource = rs
    With dgCariBarang
        .Columns("Nama Barang").Width = 2900
        .Columns("Satuan").Width = 500
        .Columns("AsalBarang").Width = 1500
        .Columns("Jenis Barang").Width = 1440
        .Columns("KdBarang").Width = 0
        .Columns("KdAsal").Width = 0
        .Columns("NoRegisterAsset").Width = 1700
        
        
        .Height = 2390
        .Top = 1680
        .Left = 0
        .Visible = True
    End With

Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()
On Error GoTo Errload
Dim i As Integer

    tempbolTampil = True
    strsql = "SELECT NamaBarang AS [Nama Barang], NamaAsal AS [Asal], DetailJenisBarang AS [Jenis Barang], Kondisi, JmlBarang, KdBarang, KdAsal, KdDetailJenisBarang, KdRuangan, KdKondisi,NoRegisterAsset " & _
        " FROM V_KondisiBarangNonMedisnew " & _
        " WHERE kdruangan='" & mstrKdRuangan & "' AND NamaBarang LIKE '%" & txtCariBarang & "%'"
    Call msubRecFO(rs, strsql)
    Set dgKondisiBarang.DataSource = rs
    With dgKondisiBarang
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i
        .Columns("Nama Barang").Width = 2200
        .Columns("Asal").Width = 1000
        
        .Columns("Jenis Barang").Width = 900
        .Columns("Kondisi").Width = 1200
        .Columns("JmlBarang").Width = 1000
        .Columns("NoRegisterAsset").Width = 1000
        
    End With
    lblJmlData.Caption = 0 & " / " & dgKondisiBarang.ApproxCount & " Data"
    tempbolTampil = False
    
Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Function sp_KondisiBarangNonMedis() As Boolean
On Error GoTo Errload

    sp_KondisiBarangNonMedis = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, txtKdBarang.Text)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, dcAsalBarang.BoundText)
        .Parameters.Append .CreateParameter("KdKondisi", adChar, adParamInput, 2, dcKondisiBarang.BoundText)
        .Parameters.Append .CreateParameter("JmlBarang", adInteger, adParamInput, , txtJmlBarang.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")
        .Parameters.Append .CreateParameter("NoRegistrasiAsset", adVarChar, adParamInput, 15, NoRegister)
    
        .ActiveConnection = dbConn
        .CommandText = "AUD_KondisiBarangNonMedisNew"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_KondisiBarangNonMedis = False
        End If
    End With

Exit Function
Errload:
    Call msubPesanError
End Function


