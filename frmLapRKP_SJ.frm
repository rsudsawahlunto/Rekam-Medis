VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLapRKP_SJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kunjungan Pasien Berdasarkan Status & Jenis Pasien"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLapRKP_SJ.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   9435
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
      TabIndex        =   7
      Top             =   2040
      Width           =   9405
      Begin VB.CommandButton cmdGrafik 
         Caption         =   "&Grafik"
         Height          =   375
         Left            =   3840
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Spreadsheet"
         Height          =   375
         Left            =   5640
         TabIndex        =   4
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   7440
         TabIndex        =   6
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
      Height          =   1035
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   9405
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
         TabIndex        =   9
         Top             =   150
         Width           =   5055
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
            Left            =   0
            TabIndex        =   3
            Top             =   600
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   120
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
            Format          =   46137347
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   2760
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
            Format          =   46137347
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
            Left            =   2400
            TabIndex        =   10
            Top             =   315
            Width           =   255
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   4395
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   7752
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   360
         Left            =   360
         TabIndex        =   0
         Top             =   480
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
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
   Begin VB.Image Image1 
      Height          =   930
      Left            =   -790
      Picture         =   "frmLapRKP_SJ.frx":08CA
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "frmLapRKP_SJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iRowNow As Integer
Dim rsTemp1 As ADODB.recordset
Dim rsTemp2 As ADODB.recordset

Private Sub cmdcari_Click()
On Error GoTo errLoad

Dim intJmlRow As Integer
Dim intNo As Integer
Dim intJmlPria As Integer
Dim intJmlWanita As Integer
Dim intJmlTotal As Integer
    
    If Periksa("datacombo", dcInstalasi, "Nama instalasi kosong") = False Then Exit Sub
        
    Call subSetGrid
    
'    Call msubRecFO(rs, "SELECT KdRuangan FROM Ruangan WHERE KdInstalasi = '" & dcInstalasi.BoundText & "' ")
'    If rs.EOF = True Then Exit Sub Else mstrKdRuangan = rs(0).Value
        
    'u/ mempercepat
    fgData.Visible = False
    MousePointer = vbHourglass
    intNo = 0
    iRowNow = 0
    intJmlPria = 0
    intJmlWanita = 0
    intJmlTotal = 0
    'Hitung jumlah row dari data yang hendak ditampilkan
    strSQL = "SELECT * FROM KelompokPasien"
    msubRecFO rsb, strSQL
    Dim strTMP As String
    strTMP = ""
    For i = 1 To rsb.RecordCount
        strTMP = strTMP & "SUM([JPL " & rsb("JenisPasien").Value & "]) AS [JPL " & rsb("JenisPasien").Value & "]"
        strTMP = strTMP & ","
        strTMP = strTMP & "SUM([JPP " & rsb("JenisPasien").Value & "]) AS [JPP " & rsb("JenisPasien").Value & "]"
        strTMP = strTMP & ","
        strTMP = strTMP & "SUM([JPT " & rsb("JenisPasien").Value & "]) AS [JPT " & rsb("JenisPasien").Value & "]"
        If i <> rsb.RecordCount Then strTMP = strTMP & ","
        rsb.MoveNext
    Next i
    
    
    strSQL = "SELECT NamaRuangan," _
        & "SUM(SLL) AS SLL," _
        & "SUM(SLP) AS SLP," _
        & "SUM(SLT) AS SLT," _
        & "SUM(SBL) AS SBL," _
        & "SUM(SBP) AS SBP," _
        & "SUM(SBT) AS SBT," _
        & strTMP _
        & " FROM v_S_RekapitulasiKStatusJenis " _
        & "WHERE (TglPendaftaran BETWEEN '" _
        & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "') " _
        & "GROUP BY NamaRuangan "
'        & " AND KdRuangan = '" & mstrKdRuangan & "' "
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
        & "SUM(Total) AS Total From v_S_RekapitulasiKStatusJenis " _
        & "WHERE TglPelayanan BETWEEN '" _
        & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " _
        & " AND KdRuangan = '" & mstrKdRuangan & "' " _
        & "GROUP BY NamaRuangan ORDER BY NamaRuangan"
    msubRecFO rsTemp1, strSQL
    'Tambahkan jumlah row dengan jumlah subtotal
    intJmlRow = intJmlRow + rsTemp1.RecordCount
    
    strSQL = "SELECT NamaRuangan,JenisPeriksa,COUNT(JenisPeriksa) AS JmlJenisPeriksa," _
        & "SUM(JmlPasienPria) AS JmlPasienPria,SUM(JmlPasienWanita) AS JmlPasienWanita," _
        & "SUM(Total) AS Total From v_S_RekapitulasiKStatusJenis " _
        & "WHERE TglPelayanan BETWEEN '" _
        & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " _
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
            .TextMatrix(iRowNow, 1) = rs("NamaRuangan").Value
            .TextMatrix(iRowNow, 2) = rs("JenisPeriksa").Value
            .TextMatrix(iRowNow, 3) = rs("JenisPasien").Value
            .TextMatrix(iRowNow, 4) = rs("TJmlPasienPria").Value
            .TextMatrix(iRowNow, 5) = rs("TJmlPasienWanita").Value
            .TextMatrix(iRowNow, 6) = rs("TTotal").Value
            intJmlPria = intJmlPria + rs("TJmlPasienPria").Value
            intJmlWanita = intJmlWanita + rs("TJmlPasienWanita").Value
            intJmlTotal = intJmlTotal + rs("TTotal").Value
            rs.MoveNext
            If rs.EOF = True Then GoTo stepJenisPeriksa
            If rsTemp2("NamaRuangan").Value = rs("NamaRuangan").Value And rsTemp2("JenisPeriksa").Value <> rs("JenisPeriksa").Value Then
stepJenisPeriksa:
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = .TextMatrix(iRowNow - 1, 1)
                .TextMatrix(iRowNow, 2) = .TextMatrix(iRowNow - 1, 2)
                .TextMatrix(iRowNow, 3) = "Sub Total"
                .TextMatrix(iRowNow, 4) = rsTemp2("JmlPasienPria").Value
                .TextMatrix(iRowNow, 5) = rsTemp2("JmlPasienWanita").Value
                .TextMatrix(iRowNow, 6) = rsTemp2("Total").Value
                subSetSubTotalRow iRowNow, 3, vbBlackness, vbWhite
                rsTemp2.MoveNext
            ElseIf rsTemp2("NamaRuangan").Value <> rs("NamaRuangan").Value Then
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = .TextMatrix(iRowNow - 1, 1)
                .TextMatrix(iRowNow, 2) = .TextMatrix(iRowNow - 1, 2)
                .TextMatrix(iRowNow, 3) = "Sub Total"
                .TextMatrix(iRowNow, 4) = rsTemp2("JmlPasienPria").Value
                .TextMatrix(iRowNow, 5) = rsTemp2("JmlPasienWanita").Value
                .TextMatrix(iRowNow, 6) = rsTemp2("Total").Value
                subSetSubTotalRow iRowNow, 3, vbBlackness, vbWhite
                rsTemp2.MoveNext
            End If
            If rs.EOF = True Then GoTo stepNamaRuangan
            If rsTemp1("NamaRuangan").Value <> rs("NamaRuangan").Value Then
stepNamaRuangan:
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = "Total"
                .TextMatrix(iRowNow, 4) = rsTemp1("JmlPasienPria").Value
                .TextMatrix(iRowNow, 5) = rsTemp1("JmlPasienWanita").Value
                .TextMatrix(iRowNow, 6) = rsTemp1("Total").Value
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
    If Periksa("datacombo", dcInstalasi, "Data instalasi kosong") = False Then Exit Sub
    
    'cmdCetak.Enabled = False
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    mstrInstalasi = dcInstalasi.BoundText
    mblnGrafik = False
    If strCetak = "LapRekapKPSJ" Then
        strSQL = "SELECT NamaRuangan FROM v_S_RekapKunjunganPsnSJ " _
            & "WHERE (TglPendaftaran BETWEEN '" _
            & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
            & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') and KdInstalasi='" & dcInstalasi.BoundText & "' GROUP BY NamaRuangan"
    ElseIf strCetak = "LapRekapKPSR" Then
        strSQL = "SELECT NamaRuangan FROM v_S_RekapKunjunganPsnSR " _
            & "WHERE (TglPendaftaran BETWEEN '" _
            & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
            & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') and KdInstalasi='" & dcInstalasi.BoundText & "' GROUP BY NamaRuangan"
    End If
    msubRecFO rs, strSQL
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbExclamation, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
'    frmCetakLapRkpPsnSJ.Show
    frmCtkLapRekap_Viewer.Show
    frmCtkLapRekap_Viewer.Caption = "Medifirst2000 - Laporan Rekapitulasi Pasien Per JenisPeriksa"
    cmdCetak.Enabled = True
End Sub

Private Sub cmdgrafik_Click()
    cmdCetak.Enabled = False
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    mblnGrafik = True
    strSQL = "SELECT NamaRuangan FROM v_S_RekapitulasiKStatusJenis " _
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
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    With Me
        .dtpAwal.Value = Now
        .dtpAkhir.Value = Now
    End With
    Call subDcSource
    'Call subSetGrid
End Sub

'Untuk setting grid
Private Sub subSetGrid()
Dim i, j, k As Integer
    If strCetak = "LapRekapKPSJ" Then
        With fgData
            .Visible = False
            .clear
            .Cols = 8
            .Rows = 4
            .Row = 0
            
            .MergeCells = 1
            .MergeCol(0) = True
            .MergeCol(1) = True
            .MergeCol(2) = True
            
            .MergeRow(0) = True
            .MergeRow(1) = True
            
            .FixedRows = 3
            .TextMatrix(0, 0) = "No."
            .TextMatrix(1, 0) = "No."
            .TextMatrix(2, 0) = "No."
            .TextMatrix(0, 1) = "Ruangan"
            .TextMatrix(1, 1) = "Ruangan"
            .TextMatrix(2, 1) = "Ruangan"
            .TextMatrix(0, 2) = "Status Kunjungan"
            .TextMatrix(0, 3) = "Status Kunjungan"
            .TextMatrix(0, 4) = "Status Kunjungan"
            .TextMatrix(1, 2) = "Lama"
            .TextMatrix(1, 3) = "Lama"
            .TextMatrix(1, 4) = "Lama"
            .TextMatrix(2, 2) = "L"
            .TextMatrix(2, 3) = "P"
            .TextMatrix(2, 4) = "Total"
            .TextMatrix(0, 5) = "Status Kunjungan"
            .TextMatrix(0, 6) = "Status Kunjungan"
            .TextMatrix(0, 7) = "Status Kunjungan"
            .TextMatrix(1, 5) = "Baru"
            .TextMatrix(1, 6) = "Baru"
            .TextMatrix(1, 7) = "Baru"
            .TextMatrix(2, 5) = "L"
            .TextMatrix(2, 6) = "P"
            .TextMatrix(2, 7) = "Total"
            
            .ColWidth(0) = 500
            .ColWidth(1) = 1750
            .ColWidth(2) = 1000
            .ColWidth(3) = 1000
            .ColWidth(4) = 1100
            .ColWidth(5) = 1000
            .ColWidth(6) = 1000
            .ColWidth(7) = 1100
            
            j = 7
            strSQL = "SELECT * FROM KelompokPasien"
            msubRecFO rsb, strSQL
            .Cols = 8 + rsb.RecordCount * 3
            For i = 1 To rsb.RecordCount * 3
                .TextMatrix(0, j + i) = "Jenis Pasien"
            Next i
            k = 0
            For i = 1 To rsb.RecordCount
                .TextMatrix(1, j + i + k) = rsb("JenisPasien").Value
                .TextMatrix(2, j + i + k) = "L"
                .ColWidth(j + i + k) = 1000
                k = k + 1
                .TextMatrix(1, j + i + k) = rsb("JenisPasien").Value
                .TextMatrix(2, j + i + k) = "P"
                .ColWidth(j + i + k) = 1000
                k = k + 1
                .TextMatrix(1, j + i + k) = rsb("JenisPasien").Value
                .TextMatrix(2, j + i + k) = "Total"
                .ColWidth(j + i + k) = 1100
                rsb.MoveNext
            Next i
            
            For j = 0 To 2
                .Row = j
                For i = 0 To .Cols - 1
                    .Col = i
                    .CellFontBold = True
                    .RowHeight(0) = 300
                    .CellAlignment = flexAlignCenterCenter
                Next i
            Next j
            
            .Visible = True
            iRowNow = 0
        End With
    ElseIf strCetak = "LapRekapKPSR" Then
        
    End If
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

Private Sub subDcSource()
    Call msubDcSource(dcInstalasi, rs, "SELECT * FROM V_InstalasiLaporan1")
End Sub

