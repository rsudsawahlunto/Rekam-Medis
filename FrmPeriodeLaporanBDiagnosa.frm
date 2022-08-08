VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmPeriodeLaporanBDiagnosa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPeriodeLaporanBDiagnosa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11070
   Begin VB.Frame Frame2 
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
      Top             =   6480
      Width           =   11055
      Begin VB.CommandButton cmdgrafik 
         Caption         =   "&Grafik"
         Height          =   375
         Left            =   7440
         TabIndex        =   6
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Spreadsheet"
         Height          =   375
         Left            =   5640
         TabIndex        =   5
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   9240
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5475
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   11055
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
         Left            =   4800
         TabIndex        =   10
         Top             =   200
         Width           =   6075
         Begin VB.CommandButton cmdcari 
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
         Begin MSComCtl2.DTPicker DTPickerAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   1
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
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
            Format          =   127336451
            UpDown          =   -1  'True
            CurrentDate     =   37956
         End
         Begin MSComCtl2.DTPicker DTPickerAkhir 
            Height          =   375
            Left            =   3600
            TabIndex        =   2
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
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
            Format          =   127336451
            UpDown          =   -1  'True
            CurrentDate     =   37956
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3240
            TabIndex        =   11
            Top             =   307
            Width           =   255
         End
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   360
         Left            =   1560
         TabIndex        =   0
         Top             =   575
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
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
      Begin MSFlexGridLib.MSFlexGrid fgdata 
         Height          =   4305
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   7594
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Instalasi Pelayanan"
         Height          =   210
         Left            =   1560
         TabIndex        =   12
         Top             =   300
         Width           =   1500
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
      Left            =   9240
      Picture         =   "FrmPeriodeLaporanBDiagnosa.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "FrmPeriodeLaporanBDiagnosa.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "FrmPeriodeLaporanBDiagnosa.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "FrmPeriodeLaporanBDiagnosa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iRowNow As Integer
Dim rsdiagnosa As ADODB.recordset
Dim iRowNow2 As Integer
Dim rsa As New ADODB.recordset

Private Sub cmdCari_Click()
    On Error GoTo errLoad

    Dim intJmlRow As Integer
    Dim intJmlPria As Integer
    Dim intJmlWanita As Integer
    Dim intJmlTotal As Integer

    If Periksa("datacombo", dcInstalasi, "Data instalasi kosong") = False Then Exit Sub

    'Panggil Desain
    Call subSetGrid
    'u/ mempercepat
    fgData.Visible = False: MousePointer = vbHourglass

    strSQL = "SELECT RuanganPelayanan, SUM(JmlPria) AS JmlPria, SUM(JmlWanita) AS JmlWanita, SUM(JmlPria + JmlWanita) AS Total" & _
    " FROM V_DataDiagnosaPasienPH_LakiPerempuan " & _
    " WHERE TglPeriksa BETWEEN '" & Format(DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' AND kdinstalasi = '" & dcInstalasi.BoundText & "'" & _
    " GROUP BY RuanganPelayanan" & _
    " ORDER BY RuanganPelayanan, SUM(JmlPria + JmlWanita)"

    Call msubRecFO(rsa, strSQL)
    intJmlRow = rsa.RecordCount

    Set rs = Nothing
    strSQL = "SELECT RuanganPelayanan, KdDiagnosa, Diagnosa, SUM(JmlPria) AS TJmlPria, SUM(JmlWanita) AS TJmlWanita, SUM(JmlPria + JmlWanita) AS TTotal " & _
    " From V_DataDiagnosaPasienPH_LakiPerempuan " & _
    " WHERE (TglPeriksa BETWEEN '" & Format(DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "') AND kdinstalasi = '" & dcInstalasi.BoundText & "'" & _
    " GROUP BY RuanganPelayanan, KdDiagnosa, Diagnosa " & _
    " ORDER BY RuanganPelayanan, SUM(JmlPria + JmlWanita) DESC, KdDiagnosa"

    Call msubRecFO(rs, strSQL)
    intJmlRow = intJmlRow + rs.RecordCount

    fgData.Rows = intJmlRow + 2
    intRowNow = 0
    intJmlPria = 0: intJmlWanita = 0: intJmlTotal = 0

    For i = 1 To rs.RecordCount
        intRowNow = intRowNow + 1
        For j = 1 To fgData.Cols - 1
            fgData.TextMatrix(intRowNow, j) = rs(j - 1).value
        Next j
        rs.MoveNext
        'sub total per Diagnosa
        If rs.EOF = True Then GoTo stepSubTotalDiagnosa
        If rs("RuanganPelayanan").value <> rsa("RuanganPelayanan").value Then
stepSubTotalDiagnosa:
            intRowNow = intRowNow + 1
            fgData.TextMatrix(intRowNow, 1) = fgData.TextMatrix(intRowNow - 1, 1)
            fgData.TextMatrix(intRowNow, 3) = "Sub Total"
            fgData.TextMatrix(intRowNow, 4) = rsa("JmlPria").value
            fgData.TextMatrix(intRowNow, 5) = rsa("JmlWanita").value
            fgData.TextMatrix(intRowNow, 6) = rsa("Total").value
            Call subSetSubTotalRow(Me, intRowNow, 2, vbBlackness, vbWhite)

            'disimpan u/ jml total
            intJmlPria = intJmlPria + rsa("JmlPria").value
            intJmlWanita = intJmlWanita + rsa("JmlWanita").value
            intJmlTotal = intJmlTotal + rsa("Total").value

            If rsa.EOF Then Exit Sub
            rsa.MoveNext
        End If
    Next i
    intRowNow = intRowNow + 1
    fgData.TextMatrix(intRowNow, 1) = "Total"
    fgData.TextMatrix(intRowNow, 4) = Format(intJmlPria, "#,###")
    fgData.TextMatrix(intRowNow, 5) = Format(intJmlWanita, "#,###")
    fgData.TextMatrix(intRowNow, 6) = Format(intJmlTotal, "#,###")
    Call subSetSubTotalRow(Me, intRowNow, 1, vbBlue, vbWhite)
    fgData.Visible = True: MousePointer = vbNormal

    cmdCetak.SetFocus
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdCetak_Click()
    cmdCetak.Enabled = False
    mdTglAwal = DTPickerAwal.value
    mdTglAkhir = DTPickerAkhir.value
    strSQL = " SELECT KdRuanganPelayanan, RuanganPelayanan, KdDiagnosa, Diagnosa, SUM(JmlPria + JmlWanita) AS Total, SUM(JmlPria) AS JmlPria, SUM(JmlWanita) AS JmlWanita " & _
    " From V_DataDiagnosaPasienPH_LakiPerempuan" & _
    " WHERE (TglPeriksa BETWEEN '" & Format(DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "') AND KdInstalasi = '" & dcInstalasi.BoundText & "'" & _
    " GROUP BY KdRuanganPelayanan, RuanganPelayanan, KdDiagnosa, Diagnosa" & _
    " ORDER BY RuanganPelayanan, SUM(JmlPria + JmlWanita) DESC"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    cetak = "Rekapdiagnosa"
    FrmViewerLaporan.Show
    FrmViewerLaporan.Caption = "Medifirst2000 - Grafik Rekapitulasi Berdasarkan Diagnosa Pasien"
    cmdCetak.Enabled = True

End Sub

Private Sub cmdgrafik_Click()
    cmdCetak.Enabled = False

    strSQL = "SELECT * FROM V_DataDiagnosaPasienPH_LakiPerempuan " _
    & "WHERE (TglPeriksa BETWEEN '" _
    & Format(DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
    & Format(DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "') ORDER BY RuanganPelayanan,diagnosa"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If

    cetak = "RekapdiagnosaGrafik"
    FrmViewerLaporan.Show
    FrmViewerLaporan.Caption = "Medifirst2000 - Grafik Rekapitulasi Berdasarkan Diagnosa Pasien"
    cmdCetak.Enabled = True

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcInstalasi.MatchedWithList = True Then DTPickerAwal.SetFocus
        strSQL = "SELECT KdInstalasi, NamaInstalasi " & _
        " From instalasi" & _
        " WHERE (KdInstalasi IN ('01', '02', '03', '06', '08')) and(NamaInstalasi LIKE '%" & dcInstalasi.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcInstalasi.Text = ""
            Exit Sub
        End If
        dcInstalasi.BoundText = rs(0).value
        dcInstalasi.Text = rs(1).value
    End If
End Sub

Private Sub DTPickerAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdcari.SetFocus
End Sub

Private Sub DTPickerAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DTPickerAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    txtJmlPria = "0": txtJmlTotal = "0": txtJmlWanita = "0"

    With Me
        .DTPickerAwal.value = Now
        .DTPickerAkhir.value = Now
    End With

    Call subDcSource
    Call subSetGrid
End Sub

Private Sub subSetGrid()
    With fgData
        .Visible = False
        .clear
        .Cols = 7
        .Rows = 2
        .Row = 0

        For i = 1 To .Cols - 1
            .Col = i
            .CellFontBold = True
            .RowHeight(0) = 300
            .CellAlignment = flexAlignCenterCenter
        Next

        .MergeCells = 1
        .MergeCol(1) = True

        .TextMatrix(0, 1) = "Ruangan"
        .TextMatrix(0, 2) = "ICD X"
        .TextMatrix(0, 3) = "Nama Diagnosa"
        .TextMatrix(0, 4) = "Laki-Laki"
        .TextMatrix(0, 5) = "Perempuan"
        .TextMatrix(0, 6) = "Total"

        .ColWidth(0) = 500
        .ColWidth(1) = 2850
        .ColWidth(2) = 1000
        .ColWidth(3) = 2850
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100
        .ColWidth(6) = 1100

        .Visible = True
        iRowNow = 0
    End With
End Sub

Private Sub subDcSource()
    On Error GoTo errLoad
    strSQL = "SELECT KdInstalasi, NamaInstalasi " & _
    " From instalasi" & _
    " WHERE (KdInstalasi IN ('01', '02', '03', '06', '08'))"
    Call msubDcSource(dcInstalasi, rs, strSQL)
    If rs.EOF = False Then dcInstalasi.BoundText = rs(0)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

