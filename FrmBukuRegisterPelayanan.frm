VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmBukuRegisterPelayanan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Buku Register Pelayanan Pasien"
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
   Icon            =   "FrmBukuRegisterPelayanan.frx":0000
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
      TabIndex        =   8
      Top             =   7680
      Width           =   14085
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   450
         Width           =   2655
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   10440
         TabIndex        =   6
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   12240
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan  Nama Pasien / No.CM"
         Height          =   210
         Left            =   120
         TabIndex        =   12
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
      TabIndex        =   9
      Top             =   960
      Width           =   14115
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
         TabIndex        =   10
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
            TabIndex        =   3
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPickerAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   136577027
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker DTPickerAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   2
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   136577027
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
      Begin MSDataGridLib.DataGrid dgData 
         Height          =   4815
         Left            =   120
         TabIndex        =   4
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
      Begin MSDataListLib.DataCombo dcJenisPasien 
         Height          =   330
         Left            =   3360
         TabIndex        =   0
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
      Begin MSDataListLib.DataCombo dcKelas 
         Height          =   330
         Left            =   6000
         TabIndex        =   16
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcNamaPelayanan 
         Height          =   330
         Left            =   3360
         TabIndex        =   17
         ToolTipText     =   "Nama Item"
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcJenisPelayanan 
         Height          =   330
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Jenis Item"
         Top             =   1200
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcNamaDokter 
         Height          =   330
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label3 
         Caption         =   "Nama Dokter"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Jenis Pelayanan"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Nama Pelayanan"
         Height          =   255
         Left            =   3360
         TabIndex        =   18
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label9 
         Caption         =   "Kelas"
         Height          =   255
         Left            =   6000
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pasien"
         Height          =   210
         Left            =   3360
         TabIndex        =   13
         Top             =   240
         Width           =   960
      End
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
      Left            =   12240
      Picture         =   "FrmBukuRegisterPelayanan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "FrmBukuRegisterPelayanan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "FrmBukuRegisterPelayanan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "FrmBukuRegisterPelayanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFilter As String

Private Sub cmdCari_Click()
    On Error GoTo hell
    
    strFilter = ""

    strFilter = " (NoCM like '%" & txtParameter.Text & "%' OR NamaPasien like '%" & txtParameter.Text & "%') AND TglPelayanan BETWEEN '" & _
    Format(DTPickerAwal.value, "yyyy/MM/dd HH:mm:00") & "' AND '" & _
    Format(DTPickerAkhir.value, "yyyy/MM/dd HH:mm:59") & "'"
    strFilter = strFilter & " and Kelas like '%" & dcKelas.Text & "%' and JenisPelayanan like '%" & dcJenisPelayanan.Text & "%' and NamaPelayanan like '%" & dcNamaPelayanan.Text & "%' and DokterOperator Like '%" & dcNamaDokter.Text & "%' and JenisPasien like '%" & dcJenisPasien.Text & "%'"
    
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

            strSQL = "SELECT * FROM BukuRegisterPelayanan_V WHERE " _
            & strFilter

        FrmViewerLaporanforBukuRegisterPelayanan.Show
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

Private Sub dcNamaDokter_Change()
   Call cmdCari_Click
End Sub

Private Sub dcNamaDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcNamaDokter.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "Select KdJenisPegawai, NamaLengkap from DataPegawai where KdJenisPegawai = '001' and NamaLengkap like '%" & dcNamaDokter.Text & "%'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcNamaDokter.Text = ""
            cmdCari.SetFocus
            Exit Sub
        End If
        dcNamaDokter.BoundText = rs(0).value
        dcNamaDokter.Text = rs(1).value
    End If
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
'    strSQL = "SELECT KdInstalasi, NamaInstalasi FROM Instalasi WHERE KdInstalasi IN ('01','02','03') and StatusEnabled='1'"
'    Call msubDcSource(dcInstalasi, dbRst, strSQL)
    Call msubDcSource(dcJenisPasien, rs, "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where StatusEnabled='1' order by JenisPasien")
    Call msubDcSource(dcKelas, rs, "Select KdKelas, DeskKelas From KelasPelayanan where StatusEnabled='1'")
    Call msubDcSource(dcJenisPelayanan, rs, "SELECT KdJnsPelayanan, Deskripsi FROM JenisPelayanan where StatusEnabled=1")
    Call msubDcSource(dcNamaPelayanan, rs, "SELECT KdPelayananRS, NamaPelayanan FROM ListPelayananRS where StatusEnabled=1")
    Call msubDcSource(dcNamaDokter, rs, "Select KdJenisPegawai, NamaLengkap from DataPegawai where KdJenisPegawai = '001'   order by NamaLengkap")
    
    Call cmdCari_Click

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subLoadData(strFilter As String)
    On Error GoTo Errload

        strSQL = "SELECT NoPendaftaran,NoCM,NamaPasien,JK,JenisPasien,Kelas,TglPelayanan,JenisPelayanan,[R/K],NamaPelayanan,AsalPelayanan," _
        & "Qty,HargaSatuan,HargaCito,HargaService,TotalBiaya,JmlHutangPenjamin,JmlTanggunganRS,JmlDiskon,TotalHarusDibayar,DokterOperator,DokterAnastesi, DokterPendamping,Ruangan,TglBkm " _
        & "FROM BukuRegisterPelayanan_V " _
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
        .Columns(1).Caption = "No.CM"
        .Columns(1).Width = 1500
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Caption = "Nama Pasien"
        .Columns(2).Width = 2500
        .Columns(3).Caption = "JK"
        .Columns(3).Width = 400
        .Columns(4).Caption = "JenisPasien"
        .Columns(4).Width = 1500
        .Columns(5).Caption = "Kelas"
        .Columns(5).Width = 1500
        .Columns(6).Caption = "TglPelayanan"
        .Columns(6).Width = 2200
        .Columns(7).Caption = "JenisPelayanan"
        .Columns(7).Width = 1800
        .Columns(8).Caption = "R/K"
        .Columns(8).Width = 1000
        .Columns(9).Caption = "NamaPelayanan"
        .Columns(9).Width = 1800
        .Columns(10).Caption = "AsalPelayanan"
        .Columns(10).Width = 1500
        .Columns(11).Caption = "QTY"
        .Columns(11).Width = 1500
        .Columns(12).Caption = "HargaSatuan"
        .Columns(12).Width = 1500
        .Columns(13).Caption = "HargaCito"
        .Columns(13).Width = 1500
        .Columns(14).Caption = "HargaService"
        .Columns(14).Width = 1500
        .Columns(15).Caption = "TotalBiaya"
        .Columns(15).Width = 1500
        .Columns(16).Caption = "JmlHutangPenjamin"
        .Columns(16).Width = 1500
        .Columns(17).Caption = "JmlTanggunganRS"
        .Columns(17).Width = 1500
        .Columns(18).Caption = "JmlDiskon"
        .Columns(18).Width = 1500
        .Columns(19).Caption = "TotalHarusDibayar"
        .Columns(19).Width = 1500
        .Columns(20).Caption = "DokterOperator"
        .Columns(20).Width = 2000
        .Columns(21).Caption = "Dokter Anastesi"
        .Columns(21).Width = 2000
        .Columns(22).Caption = "Dokter Pendamping"
        .Columns(22).Width = 2000
        .Columns(23).Caption = "Ruangan"
        .Columns(23).Width = 1500
        .Columns(24).Caption = "Tgl BKM"
        .Columns(24).Width = 1500
        
    
    End With
End Sub


Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then Call cmdCari_Click
End Sub

