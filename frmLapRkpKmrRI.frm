VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmLapRkpKmrRI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Rekapitulasi Kamar Rawat Inap"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10245
   Icon            =   "frmLapRkpKmrRI.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10245
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
      Height          =   5475
      Left            =   0
      TabIndex        =   7
      Top             =   930
      Width           =   10245
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
         Left            =   6960
         TabIndex        =   8
         Top             =   150
         Width           =   3135
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
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   62521347
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   4275
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   7541
         _Version        =   393216
         Appearance      =   0
      End
   End
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
      TabIndex        =   6
      Top             =   6360
      Width           =   10245
      Begin VB.CommandButton cmdGrafik 
         Caption         =   "&Grafik"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   4
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Spreadsheet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   3
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
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
      Left            =   8400
      Picture         =   "frmLapRkpKmrRI.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmLapRkpKmrRI.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLapRkpKmrRI.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmLapRkpKmrRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iRowNow As Integer
Dim rsTemp1 As ADODB.recordset
Dim rsTemp2 As ADODB.recordset

Private Sub cmdCari_Click()
On Error GoTo hell

Dim intJmlRow As Integer
Dim intNo As Integer
Dim intJmlBedTerisi As Integer
Dim intJmlBedKosong As Integer
Dim intJmlTotal As Integer
    Call subSetGrid
    'u/ mempercepat
    fgData.Visible = False
    MousePointer = vbHourglass
    intNo = 0
    iRowNow = 0
    intJmlBedTerisi = 0
    intJmlBedKosong = 0
    intJmlTotal = 0
    'Hitung jumlah row dari data yang hendak ditampilkan
    strSQL = "SELECT Ruangan,Kelas,NoKamar," _
        & "AVG(JmlBedTerisi) AS TJmlBedTerisi," _
        & "AVG(JmlBedKosong) AS TJmlBedKosong," _
        & "AVG(TotalBed) AS TTotalBed From V_RekapitulasiKamarRawatInap WHERE" _
        & " DAY(TglHitung) = '" & dtpAwal.Day & "' " _
        & " AND MONTH(TglHitung) = '" & dtpAwal.Month & "' " _
        & " AND YEAR(TglHitung) = '" & dtpAwal.Year & "' " _
        & "  " _
        & " GROUP BY Ruangan,Kelas,NoKamar ORDER BY Ruangan,Kelas,NoKamar"
    msubRecFO rs, strSQL
    'jika tidak ada data
    If rs.EOF = True Then
        fgData.Visible = True: MousePointer = vbNormal
        dtpAwal.SetFocus
        Exit Sub
    End If
    intJmlRow = rs.RecordCount + 1
    strSQL = "SELECT Ruangan,COUNT(Ruangan) AS JmlRuangan," _
        & "SUM(JmlBedTerisi) AS TJmlBedTerisi," _
        & "SUM(JmlBedKosong) AS TJmlBedKosong," _
        & "SUM(TotalBed) AS TTotalBed From V_RekapitulasiKamarRawatInap WHERE" _
        & " DAY(TglHitung) = '" & dtpAwal.Day & "' " _
        & " AND MONTH(TglHitung) = '" & dtpAwal.Month & "' " _
        & " AND YEAR(TglHitung) = '" & dtpAwal.Year & "' " _
        & "  " _
        & "GROUP BY Ruangan ORDER BY Ruangan"
    msubRecFO rsTemp1, strSQL
    'Tambahkan jumlah row dengan jumlah subtotal
    intJmlRow = intJmlRow + rsTemp1.RecordCount
    
    strSQL = "SELECT Ruangan,Kelas,COUNT(Kelas) AS JmlKelas," _
        & "SUM(JmlBedTerisi) AS TJmlBedTerisi," _
        & "SUM(JmlBedKosong) AS TJmlBedKosong," _
        & "SUM(TotalBed) AS TTotalBed From V_RekapitulasiKamarRawatInap WHERE" _
        & " DAY(TglHitung) = '" & dtpAwal.Day & "' " _
        & " AND MONTH(TglHitung) = '" & dtpAwal.Month & "' " _
        & " AND YEAR(TglHitung) = '" & dtpAwal.Year & "' " _
        & "  " _
        & "GROUP BY Ruangan,Kelas ORDER BY Ruangan,Kelas"
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
            .TextMatrix(iRowNow, 1) = rs("Ruangan").Value
            .TextMatrix(iRowNow, 2) = rs("Kelas").Value
            .TextMatrix(iRowNow, 3) = rs("NoKamar").Value
            .TextMatrix(iRowNow, 4) = rs("TJmlBedTerisi").Value
            .TextMatrix(iRowNow, 5) = rs("TJmlBedKosong").Value
            .TextMatrix(iRowNow, 6) = rs("TTotalBed").Value
            intJmlBedTerisi = intJmlBedTerisi + rs("TJmlBedTerisi").Value
            intJmlBedKosong = intJmlBedKosong + rs("TJmlBedKosong").Value
            intJmlTotal = intJmlTotal + rs("TTotalBed").Value
            rs.MoveNext
            If rs.EOF = True Then GoTo stepDokter
            If rsTemp2("Ruangan").Value = rs("Ruangan").Value And rsTemp2("Kelas").Value <> rs("Kelas").Value Then
stepDokter:
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = .TextMatrix(iRowNow - 1, 1)
                .TextMatrix(iRowNow, 2) = .TextMatrix(iRowNow - 1, 2)
                .TextMatrix(iRowNow, 3) = "Sub Total"
                .TextMatrix(iRowNow, 4) = rsTemp2("TJmlBedTerisi").Value
                .TextMatrix(iRowNow, 5) = rsTemp2("TJmlBedKosong").Value
                .TextMatrix(iRowNow, 6) = rsTemp2("TTotalBed").Value
                subSetSubTotalRow iRowNow, 3, vbBlackness, vbWhite
                rsTemp2.MoveNext
            ElseIf rsTemp2("Ruangan").Value <> rs("Ruangan").Value Then
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = .TextMatrix(iRowNow - 1, 1)
                .TextMatrix(iRowNow, 2) = .TextMatrix(iRowNow - 1, 2)
                .TextMatrix(iRowNow, 3) = "Sub Total"
                .TextMatrix(iRowNow, 4) = rsTemp2("TJmlBedTerisi").Value
                .TextMatrix(iRowNow, 5) = rsTemp2("TJmlBedKosong").Value
                .TextMatrix(iRowNow, 6) = rsTemp2("TTotalBed").Value
                subSetSubTotalRow iRowNow, 3, vbBlackness, vbWhite
                rsTemp2.MoveNext
            End If
            If rs.EOF = True Then GoTo stepRuangan
            If rsTemp1("Ruangan").Value <> rs("Ruangan").Value Then
stepRuangan:
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = "Total"
                .TextMatrix(iRowNow, 4) = rsTemp1("TJmlBedTerisi").Value
                .TextMatrix(iRowNow, 5) = rsTemp1("TJmlBedKosong").Value
                .TextMatrix(iRowNow, 6) = rsTemp1("TTotalBed").Value
                subSetSubTotalRow iRowNow, 1, vbBlue, vbWhite
                rsTemp1.MoveNext
            End If
        Wend
    End With
    fgData.Visible = True
    MousePointer = vbNormal
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdCetak_Click()
On Error GoTo hell
    cmdCetak.Enabled = False
    mdTglAwal = dtpAwal.Value
'    mdTglAkhir = dtpAkhir.Value
    mblnGrafik = False
    strSQL = "SELECT Ruangan FROM V_RekapitulasiKamarRawatInap WHERE " _
        & " DAY(TglHitung) = '" & dtpAwal.Day & "' " _
        & " AND MONTH(TglHitung) = '" & dtpAwal.Month & "' " _
        & " AND YEAR(TglHitung) = '" & dtpAwal.Year & "' " _
        & " "
    msubRecFO rs, strSQL
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    frmCetakLaporanRkpKmrRI.Show
    frmCetakLaporanRkpKmrRI.Caption = "Medifirst2000 - Rekapitulasi Kamar Rawat Inap"
    cmdCetak.Enabled = True
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdgrafik_Click()
On Error GoTo hell
    cmdCetak.Enabled = False
    mdTglAwal = dtpAwal.Value
'    mdTglAkhir = dtpAkhir.Value
    mblnGrafik = True
    strSQL = "SELECT Ruangan FROM V_RekapitulasiKamarRawatInap WHERE " _
        & " DAY(TglHitung) = '" & dtpAwal.Day & "' " _
        & " AND MONTH(TglHitung) = '" & dtpAwal.Month & "' " _
        & " AND YEAR(TglHitung) = '" & dtpAwal.Year & "' " _
        & " "
    msubRecFO rs, strSQL
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    frmCetakLaporanRkpKmrRI.Show
    frmCetakLaporanRkpKmrRI.Caption = "Medifirst2000 - Grafik Laporan Rekapitulasi Kamar Rawat Inap"
    cmdCetak.Enabled = True
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus  'dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    With Me
        .dtpAwal.Value = Now
        '.dtpAkhir.Value = Now
    End With
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
        .TextMatrix(0, 2) = "Kelas"
        .TextMatrix(0, 3) = "NoKamar"
        .TextMatrix(0, 4) = "Bed Isi"
        .TextMatrix(0, 5) = "Bed Kosong"
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
