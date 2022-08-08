VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLapRkpPsnJenisPeriksa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Rekapitulasi Pasien Per JenisPeriksa"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLapRkpPsnJenisPeriksa.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   10155
   Begin VB.Frame fraButton 
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
      Left            =   0
      TabIndex        =   8
      Top             =   6600
      Width           =   10125
      Begin VB.CommandButton cmdGrafik 
         Caption         =   "&Grafik"
         Height          =   375
         Left            =   6480
         TabIndex        =   6
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Spreadsheet"
         Height          =   375
         Left            =   4680
         TabIndex        =   5
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   8280
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraPeriode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5595
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   10125
      Begin VB.Frame Frame4 
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
         Left            =   4200
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
            TabIndex        =   3
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
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
            OLEDropMode     =   1
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   126943235
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   2
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
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   126943235
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3120
            TabIndex        =   11
            Top             =   315
            Width           =   255
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   4395
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   7752
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   330
         Left            =   360
         TabIndex        =   0
         Top             =   480
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Instalasi Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   1755
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   13
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
      Left            =   8280
      Picture         =   "frmLapRkpPsnJenisPeriksa.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmLapRkpPsnJenisPeriksa.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLapRkpPsnJenisPeriksa.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmLapRkpPsnJenisPeriksa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iRowNow As Integer
Dim rsTemp1 As ADODB.recordset
Dim rsTemp2 As ADODB.recordset

Private Sub cmdCari_Click()
    On Error GoTo errLoad

    Dim intJmlRow As Integer
    Dim intNo As Integer
    Dim intJmlPria As Integer
    Dim intJmlWanita As Integer
    Dim intJmlTotal As Integer

    If Periksa("datacombo", dcInstalasi, "Nama instalasi kosong") = False Then Exit Sub

    Call subSetGrid

    Call msubRecFO(rs, "SELECT KdRuangan FROM Ruangan WHERE KdInstalasi = '" & dcInstalasi.BoundText & "' ")
    If rs.EOF = True Then Exit Sub Else mstrKdRuangan = rs(0).value

    'u/ mempercepat
    fgData.Visible = False
    MousePointer = vbHourglass
    intNo = 0
    iRowNow = 0
    intJmlPria = 0
    intJmlWanita = 0
    intJmlTotal = 0
    'Hitung jumlah row dari data yang hendak ditampilkan
    strSQL = "SELECT NamaRuangan,JenisPeriksa,JenisPasien," _
    & "SUM(JmlPasienPria) AS TJmlPasienPria," _
    & "SUM(JmlPasienWanita) AS TJmlPasienWanita," _
    & "SUM(Total) AS TTotal From V_RekapitulasiKunjunganPasienBJenisPeriksa " _
    & "WHERE (TglPelayanan BETWEEN '" _
    & Format(dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" _
    & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "') " _
    & " AND KdRuangan = '" & mstrKdRuangan & "' " _
    & "GROUP BY NamaRuangan,JenisPeriksa,JenisPasien ORDER BY NamaRuangan,JenisPeriksa,JenisPasien"
    msubRecFO rs, strSQL
    'jika tidak ada data
    If rs.EOF = True Then
        fgData.Visible = True: MousePointer = vbNormal
        dcInstalasi.SetFocus
        Exit Sub
    End If
    intJmlRow = rs.RecordCount + 1
    strSQL = "SELECT NamaRuangan,COUNT(NamaRuangan) AS JmlRuangan," _
    & "SUM(JmlPasienPria) AS JmlPasienPria,SUM(JmlPasienWanita) AS JmlPasienWanita," _
    & "SUM(Total) AS Total From V_RekapitulasiKunjunganPasienBJenisPeriksa " _
    & "WHERE TglPelayanan BETWEEN '" _
    & Format(dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" _
    & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "' " _
    & " AND KdRuangan = '" & mstrKdRuangan & "' " _
    & "GROUP BY NamaRuangan ORDER BY NamaRuangan"
    msubRecFO rsTemp1, strSQL
    'Tambahkan jumlah row dengan jumlah subtotal
    intJmlRow = intJmlRow + rsTemp1.RecordCount

    strSQL = "SELECT NamaRuangan,JenisPeriksa,COUNT(JenisPeriksa) AS JmlJenisPeriksa," _
    & "SUM(JmlPasienPria) AS JmlPasienPria,SUM(JmlPasienWanita) AS JmlPasienWanita," _
    & "SUM(Total) AS Total From V_RekapitulasiKunjunganPasienBJenisPeriksa " _
    & "WHERE TglPelayanan BETWEEN '" _
    & Format(dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" _
    & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "' " _
    & " AND KdRuangan = '" & mstrKdRuangan & "'  " _
    & "GROUP BY NamaRuangan,JenisPeriksa ORDER BY NamaRuangan,JenisPeriksa"
    msubRecFO rsTemp2, strSQL
    'Tambahkan jumlah row dengan jumlah subtotal
    intJmlRow = intJmlRow + rsTemp2.RecordCount

    'u/ menampilkan yang di group by
    With fgData
        'jml baris akhir
        .Rows = intJmlRow
        While rs.EOF = False
            'baris u/ sub total
            iRowNow = iRowNow + 1
            intNo = intNo + 1
            .TextMatrix(iRowNow, 0) = intNo
            .TextMatrix(iRowNow, 1) = rs("NamaRuangan").value
            .TextMatrix(iRowNow, 2) = rs("JenisPeriksa").value
            .TextMatrix(iRowNow, 3) = rs("JenisPasien").value
            .TextMatrix(iRowNow, 4) = rs("TJmlPasienPria").value
            .TextMatrix(iRowNow, 5) = rs("TJmlPasienWanita").value
            .TextMatrix(iRowNow, 6) = rs("TTotal").value
            intJmlPria = intJmlPria + rs("TJmlPasienPria").value
            intJmlWanita = intJmlWanita + rs("TJmlPasienWanita").value
            intJmlTotal = intJmlTotal + rs("TTotal").value
            rs.MoveNext
            If rs.EOF = True Then GoTo stepJenisPeriksa
            If rsTemp2("NamaRuangan").value = rs("NamaRuangan").value And rsTemp2("JenisPeriksa").value <> rs("JenisPeriksa").value Then
stepJenisPeriksa:
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = .TextMatrix(iRowNow - 1, 1)
                .TextMatrix(iRowNow, 2) = .TextMatrix(iRowNow - 1, 2)
                .TextMatrix(iRowNow, 3) = "Sub Total"
                .TextMatrix(iRowNow, 4) = rsTemp2("JmlPasienPria").value
                .TextMatrix(iRowNow, 5) = rsTemp2("JmlPasienWanita").value
                .TextMatrix(iRowNow, 6) = rsTemp2("Total").value
                subSetSubTotalRow iRowNow, 3, vbBlackness, vbWhite
                rsTemp2.MoveNext
            ElseIf rsTemp2("NamaRuangan").value <> rs("NamaRuangan").value Then
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = .TextMatrix(iRowNow - 1, 1)
                .TextMatrix(iRowNow, 2) = .TextMatrix(iRowNow - 1, 2)
                .TextMatrix(iRowNow, 3) = "Sub Total"
                .TextMatrix(iRowNow, 4) = rsTemp2("JmlPasienPria").value
                .TextMatrix(iRowNow, 5) = rsTemp2("JmlPasienWanita").value
                .TextMatrix(iRowNow, 6) = rsTemp2("Total").value
                subSetSubTotalRow iRowNow, 3, vbBlackness, vbWhite
                rsTemp2.MoveNext
            End If
            If rs.EOF = True Then GoTo stepNamaRuangan
            If rsTemp1("NamaRuangan").value <> rs("NamaRuangan").value Then
stepNamaRuangan:
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = "Total"
                .TextMatrix(iRowNow, 4) = rsTemp1("JmlPasienPria").value
                .TextMatrix(iRowNow, 5) = rsTemp1("JmlPasienWanita").value
                .TextMatrix(iRowNow, 6) = rsTemp1("Total").value
                subSetSubTotalRow iRowNow, 1, vbBlue, vbWhite
                rsTemp1.MoveNext
            End If
        Wend
    End With
    fgData.Visible = True
    MousePointer = vbNormal

    Exit Sub
errLoad:
    Call msubPesanError
    fgData.Visible = True: MousePointer = vbNormal
End Sub

Private Sub cmdCetak_Click()
    cmdCetak.Enabled = False
    mdTglAwal = dtpAwal.value
    mdTglAkhir = dtpAkhir.value
    mblnGrafik = False
    strSQL = "SELECT NamaRuangan FROM V_RekapitulasiKunjunganPasienBJenisPeriksa " _
    & "WHERE (TglPelayanan BETWEEN '" _
    & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
    & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " _
    & " AND KdRuangan = '" & mstrKdRuangan & "' "
    msubRecFO rs, strSQL
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbExclamation, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    frmCetakLaporanJenisPeriksa.Show
    frmCetakLaporanJenisPeriksa.Caption = "Medifirst2000 - Laporan Rekapitulasi Pasien Per JenisPeriksa"
    cmdCetak.Enabled = True
End Sub

Private Sub cmdgrafik_Click()
    cmdCetak.Enabled = False
    mdTglAwal = dtpAwal.value
    mdTglAkhir = dtpAkhir.value
    mblnGrafik = True
    strSQL = "SELECT NamaRuangan FROM V_RekapitulasiKunjunganPasienBJenisPeriksa " _
    & "WHERE (TglPelayanan BETWEEN '" _
    & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
    & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " _
    & " AND KdRuangan = '" & mstrKdRuangan & "' "
    msubRecFO rs, strSQL
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbExclamation, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    frmCetakLaporanJenisPeriksa.Show
    frmCetakLaporanJenisPeriksa.Caption = "Medifirst2000 - Grafik Laporan Rekapitulasi Pasien Per JenisPeriksa"
    cmdCetak.Enabled = True
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcInstalasi.MatchedWithList = True Then dtpAwal.SetFocus
        strSQL = "SELECT KdInstalasi, NamaInstalasi FROM  Instalasi WHERE (KdInstalasi IN ('09', '10', '16'))  and (Namainstalasi LIKE '%" & dcInstalasi.Text & "%')ORDER BY NamaInstalasi"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcInstalasi.Text = ""
            Exit Sub
        End If
        dcInstalasi.BoundText = rs(0).value
        dcInstalasi.Text = rs(1).value
    End If
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    With Me
        .dtpAwal.value = Now
        .dtpAkhir.value = Now
    End With
    Call msubDcSource(dcInstalasi, rs, "SELECT KdInstalasi, NamaInstalasi FROM  Instalasi WHERE (KdInstalasi IN ('09', '10', '16')) ORDER BY NamaInstalasi")
    Call subSetGrid
End Sub

'Untuk setting grid
Private Sub subSetGrid()
    With fgData
        .Visible = False
        .clear
        .Cols = 7
        .Rows = 2
        .Row = 0

        For i = 0 To .Cols - 1
            .Col = i
            .CellFontBold = True
            .RowHeight(0) = 300
            .CellAlignment = flexAlignCenterCenter
        Next

        .MergeCells = 1
        .MergeCol(1) = True
        .MergeCol(2) = True

        .TextMatrix(0, 0) = "No."
        .TextMatrix(0, 1) = "Ruangan"
        .TextMatrix(0, 2) = "JenisPeriksa"
        .TextMatrix(0, 3) = "JenisPasien"
        .TextMatrix(0, 4) = "Laki-Laki"
        .TextMatrix(0, 5) = "Perempuan"
        .TextMatrix(0, 6) = "Total"

        .ColWidth(0) = 500
        .ColWidth(1) = 1750
        .ColWidth(2) = 2850
        .ColWidth(3) = 1200
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100
        .ColWidth(6) = 1100

        .Visible = True
        iRowNow = 0
    End With
End Sub

'Untuk mensetting grid di row subtotal
Private Sub subSetSubTotalRow(iRowNow As Integer, iColMulai As Integer, vbBackColor, vbForeColor)
    Dim i As Integer
    With fgData
        'tampilan Black & White
        For i = iColMulai To .Cols - 1
            .Col = i
            .Row = iRowNow
            .CellBackColor = vbBackColor
            .CellForeColor = vbForeColor
            .CellFontBold = True
        Next
    End With
End Sub

