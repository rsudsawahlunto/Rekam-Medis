VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBukuRegisterPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Buku Register Pasien Masuk"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBukuRegisterPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   14115
   Begin VB.Frame Frame2 
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
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   7680
      Width           =   14085
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   450
         Width           =   2655
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   10440
         TabIndex        =   7
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   12240
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo cboRuangan 
         Height          =   360
         Left            =   3240
         TabIndex        =   17
         Top             =   450
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcRuanganPerujuk 
         Height          =   360
         Left            =   6720
         TabIndex        =   30
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan Perujuk"
         Height          =   210
         Left            =   6720
         TabIndex        =   29
         Top             =   240
         Width           =   2700
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan"
         Height          =   210
         Left            =   3240
         TabIndex        =   18
         Top             =   195
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan  Nama Pasien / No.CM"
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   195
         Width           =   2640
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buku Register"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6675
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   14115
      Begin VB.ComboBox cbStatusPasienRuangan 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "FrmBukuRegisterPasien.frx":0CCA
         Left            =   2280
         List            =   "FrmBukuRegisterPasien.frx":0CCC
         TabIndex        =   34
         Top             =   1200
         Width           =   2055
      End
      Begin VB.ComboBox cbStatusPasienRS 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "FrmBukuRegisterPasien.frx":0CCE
         Left            =   120
         List            =   "FrmBukuRegisterPasien.frx":0CD0
         TabIndex        =   33
         Top             =   1200
         Width           =   2055
      End
      Begin VB.ComboBox cbJK 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "FrmBukuRegisterPasien.frx":0CD2
         Left            =   11400
         List            =   "FrmBukuRegisterPasien.frx":0CD4
         TabIndex        =   32
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Frame Frame3 
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
         Left            =   8160
         TabIndex        =   11
         Top             =   150
         Width           =   5775
         Begin VB.CommandButton cmdCari 
            Caption         =   "Cari"
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
            TabIndex        =   4
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPickerAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   2
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   132120579
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker DTPickerAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   3
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   132120579
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   12
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgData 
         Height          =   4815
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   8493
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin MSDataListLib.DataCombo dcJenisPasien 
         Height          =   330
         Left            =   3360
         TabIndex        =   1
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin MSDataListLib.DataCombo dcAsalPasien 
         Height          =   330
         Left            =   6000
         TabIndex        =   19
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcKelas 
         Height          =   330
         Left            =   4560
         TabIndex        =   24
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcKondisiPulang 
         Height          =   330
         Left            =   6840
         TabIndex        =   26
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcKasusPenyakit 
         Height          =   330
         Left            =   9120
         TabIndex        =   28
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label13 
         Caption         =   "Jenis Kelamin"
         Height          =   255
         Left            =   11400
         TabIndex        =   31
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "Kasus Penyakit"
         Height          =   255
         Left            =   9120
         TabIndex        =   27
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "Kondisi Pulang"
         Height          =   255
         Left            =   6840
         TabIndex        =   25
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Kelas"
         Height          =   255
         Left            =   4560
         TabIndex        =   23
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "StatusPasienRuangan"
         Height          =   255
         Left            =   2280
         TabIndex        =   22
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "StatusPasienRS"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Kecamatan"
         Height          =   255
         Left            =   6000
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pasien"
         Height          =   210
         Left            =   3360
         TabIndex        =   15
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama Instalasi"
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1140
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   16
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
      Left            =   12240
      Picture         =   "FrmBukuRegisterPasien.frx":0CD6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "FrmBukuRegisterPasien.frx":1A5E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "FrmBukuRegisterPasien.frx":441F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "FrmBukuRegisterPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFilter As String

Private Sub cboRuangan_GotFocus()
    If dcInstalasi.BoundText = "02" Then
        strSQL = "SELECT kdRuangan, NamaRuangan FROM Ruangan  WHERE KdInstalasi IN ('02','11','06') and StatusEnabled='1' order by NamaRuangan"
        Call msubDcSource(cboRuangan, rs, strSQL)
    Else
        Call msubDcSource(cboRuangan, rs, "SELECT kdRuangan, NamaRuangan FROM Ruangan where kdInstalasi = '" & dcInstalasi.BoundText & "' and StatusEnabled='1' order by NamaRuangan")
    End If
End Sub

Private Sub dcRuanganPerujuk_GotFocus()
    If dcInstalasi.BoundText = "02" Then
        strSQL = "SELECT kdRuangan, NamaRuangan FROM Ruangan  WHERE KdInstalasi IN ('02','11','06') and StatusEnabled='1' order by NamaRuangan"
        Call msubDcSource(dcRuanganPerujuk, rs, strSQL)
    Else
        Call msubDcSource(dcRuanganPerujuk, rs, "SELECT kdRuangan, NamaRuangan FROM Ruangan where kdInstalasi = '" & dcInstalasi.BoundText & "' and StatusEnabled='1' order by NamaRuangan")
    End If
End Sub

Private Sub dcAsalPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcAsalPasien.MatchedWithList = True Then dcAsalPasien.SetFocus
        strSQL = "Select KdKecamatan, NamaKecamatan From Kecamatan where StatusEnabled='1' and (NamaKecamatan LIKE '%" & dcAsalPasien.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcAsalPasien.Text = ""
            Exit Sub
        End If
        dcAsalPasien.BoundText = rs(0).value
        dcAsalPasien.Text = rs(1).value
    End If
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcInstalasi.MatchedWithList = True Then dcInstalasi.SetFocus
        strSQL = "SELECT KdInstalasi, NamaInstalasi FROM Instalasi WHERE KdInstalasi IN ('01','02','03') and StatusEnabled='1' and (Namainstalasi LIKE '%" & dcInstalasi.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcInstalasi.Text = ""
            Exit Sub
        End If
        dcInstalasi.BoundText = rs(0).value
        dcInstalasi.Text = rs(1).value
    End If
End Sub

Private Sub cboRuangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If cboRuangan.MatchedWithList = True Then cboRuangan.SetFocus
        If dcInstalasi.BoundText = "02" Then
            strSQL = "SELECT kdRuangan, NamaRuangan FROM Ruangan  WHERE KdInstalasi IN ('02','11','06') and StatusEnabled='1' and (NamaRuangan LIKE '%" & cboRuangan.Text & "%') order by NamaRuangan"
        Else
            strSQL = "SELECT kdRuangan, NamaRuangan FROM Ruangan where kdInstalasi = '" & dcInstalasi.BoundText & "' and StatusEnabled='1'and (NamaRuangan LIKE '%" & cboRuangan.Text & "%') order by NamaRuangan"
        End If
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            cboRuangan.Text = ""
            Exit Sub
        End If
        cboRuangan.BoundText = rs(0).value
        cboRuangan.Text = rs(1).value
        Call cmdCari_Click
    End If
End Sub

Private Sub dcRuanganPerujuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcRuanganPerujuk.MatchedWithList = True Then dcRuanganPerujuk.SetFocus
        If dcInstalasi.BoundText = "02" Then
            strSQL = "SELECT kdRuangan, NamaRuangan FROM Ruangan  WHERE KdInstalasi IN ('02','11','06') and StatusEnabled='1' and (NamaRuangan LIKE '%" & dcRuanganPerujuk.Text & "%') order by NamaRuangan"
        Else
            strSQL = "SELECT kdRuangan, NamaRuangan FROM Ruangan where kdInstalasi = '" & dcInstalasi.BoundText & "' and StatusEnabled='1'and (NamaRuangan LIKE '%" & dcRuanganPerujuk.Text & "%') order by NamaRuangan"
        End If
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcRuanganPerujuk.Text = ""
            Exit Sub
        End If
        dcRuanganPerujuk.BoundText = rs(0).value
        dcRuanganPerujuk.Text = rs(1).value
        Call cmdCari_Click
    End If
End Sub

Private Sub cmdCari_Click()
    On Error GoTo hell

    strStatus = ""
    strFilter = ""
    strStatuspasienRS = ""
    strStatuspasienRuangan = ""
    strStatusJenisKelamin = ""

    
    If cbStatusPasienRS.Text = "Baru" Then
        strStatuspasienRS = "AND StatusPasienRS ='Baru'"
    ElseIf cbStatusPasienRS.Text = "Lama" Then
        strStatuspasienRS = "AND StatusPasienRS ='Lama'"
    Else
        strStatuspasienRS = ""
    End If
    
    If cbStatusPasienRuangan.Text = "Baru" Then
        strStatuspasienRuangan = "AND StatusPasienRuangan ='Baru'"
     ElseIf cbStatusPasienRuangan.Text = "Lama" Then
        strStatuspasienRuangan = "AND StatusPasienRuangan ='Lama'"
    Else
        strStatuspasienRuangan = ""
    End If
    
    If cbJK.Text = "L" Then
        strStatusJenisKelamin = "AND JenisKelamin ='L'"
    ElseIf cbJK.Text = "P" Then
        strStatusJenisKelamin = "AND JenisKelamin ='P'"
    Else
        strStatusJenisKelamin = ""
    End If
    
    
    If dcKondisiPulang.Text = "" Then
        strStatusKondisiPulang = "and KondisiPulang is null"
    Else
        strStatusKondisiPulang = "and KondisiPulang like '%" & dcKondisiPulang.Text & "%'"
    End If
    
       
    If dcRuanganPerujuk.Text = "" Then
        strstatusRuanganPerujuk = " and RuanganPerujuk is null"
    Else
        strstatusRuanganPerujuk = " and RuanganPerujuk like '%" & dcRuanganPerujuk.Text & "%'"
    End If
    
    strFilter = " (NoCM like '%" & txtParameter.Text & "%' OR NamaLengkap like '%" & txtParameter.Text & "%') AND TglPendaftaran BETWEEN '" & _
    Format(DTPickerAwal.value, "yyyy/MM/dd HH:mm:00") & "' AND '" & _
    Format(DTPickerAkhir.value, "yyyy/MM/dd HH:mm:59") & "'" & strStatus & strStatuspasienRS & strStatuspasienRuangan & strStatusJenisKelamin & strStatusKondisiPulang & strstatusRuanganPerujuk
    strFilter = strFilter & " and KdInstalasi = '" & dcInstalasi.BoundText & "' and KelasPelayanan like '%" & dcKelas.Text & "%' and NamaSubInstalasi like '%" & dcKasusPenyakit.Text & "%' and JenisKelamin like '%" & cbJK.Text & "%' and Kecamatan Like '%" & dcAsalPasien.Text & "%' and JenisPasien like '%" & dcJenisPasien.Text & "%' and NamaRuangan like '%" & cboRuangan.Text & "%'"
    
    subLoadData strFilter
    'lblJumData.Caption = "Data 0/" & rs.RecordCount
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo Errload

    If dgData.ApproxCount <> 0 Then
        vLaporan = ""
        If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"

            strSQL = "SELECT * FROM BukuRegisterALL_V WHERE " _
            & strFilter

        FrmViewerLaporanNew.Show
        cmdCetak.Enabled = True
    Else
        MsgBox "Tidak ada data", vbInformation, "information"
        Exit Sub
    End If
    Exit Sub
Errload:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcInstalasi_Change()
    On Error GoTo hell
    
    cboRuangan.Text = ""
    dcRuanganPerujuk.Text = ""
    Set cboRuangan.RowSource = Nothing
    Set dcRuanganPerujuk.RowSource = Nothing
    txtParameter = ""

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcJenisPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcJenisPasien.MatchedWithList = True Then dcJenisPasien.SetFocus
        strSQL = "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where StatusEnabled='1' and (JenisPasien LIKE '%" & dcJenisPasien.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcJenisPasien.Text = ""
            Exit Sub
        End If
        dcJenisPasien.BoundText = rs(0).value
        dcJenisPasien.Text = rs(1).value
    End If
End Sub

Private Sub dgData_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgData
    WheelHook.WheelHook dgData
End Sub

Private Sub dgData_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    lblJumData.Caption = "Data " & dgData.Bookmark & "/" & dgData.ApproxCount
End Sub

Private Sub DTPickerAkhir_Change()
    DTPickerAkhir.MaxDate = Now
End Sub

Private Sub DTPickerAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub DTPickerAwal_Change()
    DTPickerAwal.MaxDate = Now
End Sub

Private Sub DTPickerAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DTPickerAkhir.SetFocus
End Sub

Private Sub DTPickerAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then DTPickerAkhir.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo hell
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    With Me
        .DTPickerAwal.value = Format(Now, "dd MMM yyyy 00:00:00")
        .DTPickerAkhir.value = Now
    End With
    strSQL = "SELECT KdInstalasi, NamaInstalasi FROM Instalasi WHERE KdInstalasi IN ('01','02','03') and StatusEnabled='1'"
    Call msubDcSource(dcInstalasi, dbRst, strSQL)
    Call msubDcSource(dcJenisPasien, rs, "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where StatusEnabled='1' order by JenisPasien")
    Call msubDcSource(dcAsalPasien, rs, "Select KdKecamatan, NamaKecamatan From Kecamatan where StatusEnabled='1' order by NamaKecamatan")
    Call msubDcSource(dcKelas, rs, "Select KdKelas, DeskKelas From KelasPelayanan where StatusEnabled='1'")
    Call msubDcSource(dcKondisiPulang, rs, "Select KdKondisiPulang, KondisiPulang From KondisiPulang where StatusEnabled='1'")
    Call msubDcSource(dcKasusPenyakit, rs, "Select KdSubInstalasi, NamaSubInstalasi From SubInstalasi order by KdSubInstalasi ")

    dcInstalasi.BoundText = "01"
    
    cbStatusPasienRS.AddItem "Baru"
    cbStatusPasienRS.AddItem "Lama"
    
    cbStatusPasienRuangan.AddItem "Baru"
    cbStatusPasienRuangan.AddItem "Lama"
    
    cbJK.AddItem "L"
    cbJK.AddItem "P"
    
    Call cmdCari_Click

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subLoadData(strFilter As String)
    On Error GoTo Errload

        strSQL = "SELECT NoPendaftaran,TglPendaftaran,NoCM,NamaLengkap as NamaPasien,JenisKelamin,JenisPasien,StatusPasienRS, StatusPasienRuangan, CaraMasuk, RuanganPerujuk, StatusKeluar," _
        & "StatusPulang, KondisiPulang, TglKeluar, NoKamarNoBed, KelasPelayanan, NamaSubInstalasi as KasusPenyakit " _
        & "FROM BukuRegisterALL_V " _
        & "WHERE " & strFilter

    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgData.DataSource = rs
    subSetGrid
    Exit Sub
Errload:
    msubPesanError
End Sub

Private Sub subSetGrid()
    With dgData

        .Columns(0).Caption = "NoPendaftaran"
        .Columns(0).Width = 1500
        .Columns(1).Caption = "Tgl Pendaftaran"
        .Columns(1).Width = 2200
        .Columns(2).Caption = "No.CM"
        .Columns(2).Width = 1500
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Caption = "Nama Pasien"
        .Columns(3).Width = 2500
        .Columns(4).Caption = "JK"
        .Columns(4).Width = 400
        .Columns(5).Caption = "JenisPasien"
        .Columns(5).Width = 1500
        .Columns(6).Caption = "StatusPasienRS"
        .Columns(6).Width = 1500
        .Columns(7).Caption = "StatusPasienRuangan"
        .Columns(7).Width = 1800
        .Columns(8).Caption = "CaraMasuk"
        .Columns(8).Width = 1500
        .Columns(9).Caption = "RuanganPerujuk"
        .Columns(9).Width = 1500
        .Columns(10).Caption = "StatusKeluar"
        .Columns(10).Width = 1500
        .Columns(11).Caption = "StatusPulang"
        .Columns(11).Width = 1500
        .Columns(12).Caption = "KondisiPulang"
        .Columns(12).Width = 1500
        .Columns(13).Caption = "Tgl Keluar"
        .Columns(13).Width = 1500
        .Columns(14).Caption = "NoKamarNoBed"
        .Columns(14).Width = 1500
        .Columns(15).Caption = "KelasPelayanan"
        .Columns(15).Width = 1500
        .Columns(16).Caption = "KasusPenyakit"
        .Columns(16).Width = 2500

    End With
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then Call cmdCari_Click
End Sub

