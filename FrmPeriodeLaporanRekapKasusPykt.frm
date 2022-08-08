VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPeriodeLaporanRekapKasusPykt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPeriodeLaporanRekapKasusPykt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   10230
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
      Height          =   6315
      Left            =   0
      TabIndex        =   8
      Top             =   930
      Width           =   10215
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
         Left            =   3960
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
            CustomFormat    =   "dd MMMM, yyyy"
            Format          =   47906819
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
            CustomFormat    =   "dd MMMM, yyyy"
            Format          =   47906819
            UpDown          =   -1  'True
            CurrentDate     =   37956
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3240
            TabIndex        =   11
            Top             =   322
            Width           =   255
         End
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
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
      Begin MSFlexGridLib.MSFlexGrid fgdata 
         Height          =   4905
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   9945
         _ExtentX        =   17542
         _ExtentY        =   8652
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Instalasi Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   1890
      End
   End
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
      TabIndex        =   7
      Top             =   7320
      Width           =   10245
      Begin VB.CommandButton cmdgrafik 
         Caption         =   "&Grafik"
         Height          =   375
         Left            =   6150
         TabIndex        =   5
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Spreadsheet"
         Height          =   375
         Left            =   4320
         TabIndex        =   4
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   7980
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Image Image2 
      Height          =   930
      Left            =   0
      Picture         =   "FrmPeriodeLaporanRekapKasusPykt.frx":08CA
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "FrmPeriodeLaporanRekapKasusPykt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iRowNow As Integer
Dim rswilayah As New ADODB.recordset
Dim iRowNow2 As Integer
Dim rsA As New ADODB.recordset

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
    fgdata.Visible = False: MousePointer = vbHourglass
    
    strSQL = "SELECT Ruangan, SUM(JmlPasienPria) AS JmlPria, SUM(JmlPasienWanita) AS JmlWanita, SUM(Total) AS Total" & _
        " FROM V_RekapitulasiPasienBKasusPenyakit " & _
        " WHERE TglPendaftaran BETWEEN '" & Format(DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND kdinstalasi = '" & dcInstalasi.BoundText & "'" & _
        " GROUP BY Ruangan" & _
        " ORDER BY Ruangan"
    Call msubRecFO(rsA, strSQL)
    intJmlRow = rsA.RecordCount
    
    If intJmlRow = 0 Then fgdata.Visible = True: MousePointer = vbDefault: dcInstalasi.SetFocus: Exit Sub
    
    strSQL = "SELECT Ruangan,KasusPenyakit,SUM(JmlPasienPria) AS TJmlPasienPria, SUM(JmlPasienWanita) AS TJmlPasienWanita, SUM(Total) AS TTotal " & _
        " From V_RekapitulasiPasienBKasusPenyakit " & _
        " WHERE (TglPendaftaran BETWEEN '" & Format(DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "') AND kdinstalasi = '" & dcInstalasi.BoundText & "'" & _
        " GROUP BY Ruangan, KasusPenyakit " & _
        " ORDER BY Ruangan, KasusPenyakit"
    Call msubRecFO(rs, strSQL)
    intJmlRow = intJmlRow + rs.RecordCount
    
    fgdata.Rows = intJmlRow + 2
    intRowNow = 0
    intJmlPria = 0: intJmlWanita = 0: intJmlTotal = 0

    For i = 1 To rs.RecordCount
        intRowNow = intRowNow + 1
        For j = 1 To fgdata.Cols - 1
            fgdata.TextMatrix(intRowNow, j) = rs(j - 1).Value
        Next j
        rs.MoveNext
        'sub total per KasusPenyakit
        If rs.EOF = True Then GoTo stepSubTotalKasusPenyakit
        If rs("Ruangan").Value <> rsA("Ruangan").Value Then
stepSubTotalKasusPenyakit:
            intRowNow = intRowNow + 1
            fgdata.TextMatrix(intRowNow, 1) = fgdata.TextMatrix(intRowNow - 1, 1)
            fgdata.TextMatrix(intRowNow, 2) = "Sub Total"
            fgdata.TextMatrix(intRowNow, 3) = rsA("JmlPria").Value
            fgdata.TextMatrix(intRowNow, 4) = rsA("JmlWanita").Value
            fgdata.TextMatrix(intRowNow, 5) = rsA("Total").Value
            Call subSetSubTotalRow(Me, intRowNow, 2, vbBlackness, vbWhite)
            
            'disimpan u/ jml total
            intJmlPria = intJmlPria + rsA("JmlPria").Value
            intJmlWanita = intJmlWanita + rsA("JmlWanita").Value
            intJmlTotal = intJmlTotal + rsA("Total").Value
            
            If rsA.EOF Then Exit Sub
            rsA.MoveNext
        End If
    Next i
    intRowNow = intRowNow + 1
    fgdata.TextMatrix(intRowNow, 1) = "Total"
    fgdata.TextMatrix(intRowNow, 3) = Format(intJmlPria, "#,###")
    fgdata.TextMatrix(intRowNow, 4) = Format(intJmlWanita, "#,###")
    fgdata.TextMatrix(intRowNow, 5) = Format(intJmlTotal, "#,###")
    Call subSetSubTotalRow(Me, intRowNow, 1, vbBlue, vbWhite)
    fgdata.Visible = True: MousePointer = vbNormal

    cmdCetak.SetFocus
    Exit Sub
errLoad:
    Call msubPesanError
End Sub


Private Sub cmdCetak_Click()
    cmdCetak.Enabled = False
'    mdTglAwal = dtpTglAwal.Value
'    mdTglAkhir = dtpTglAkhir.Value
    strSQL = "SELECT * FROM V_RekapitulasiPasienBKasusPenyakit " _
        & "WHERE (TglPendaftaran BETWEEN '" _
        & Format(DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "') " _
        & " and kdinstalasi = '" & dcInstalasi.BoundText & "' ORDER BY ruangan,kasuspenyakit"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    cetak = "Rekapkasuspenyakit"
    FrmViewerLaporan.Show
    FrmViewerLaporan.Caption = "Medifirst2000 - Rekapitulasi Berdasarkan Kasus Penyakit IGD"
    cmdCetak.Enabled = True

End Sub

Private Sub cmdgrafik_Click()
    cmdCetak.Enabled = False
    
    strSQL = "SELECT * FROM V_RekapitulasiPasienBKasusPenyakit " _
        & "WHERE (TglPendaftaran BETWEEN '" _
        & Format(DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "') " _
        & " and kdinstalasi = '" & dcInstalasi.BoundText & "' ORDER BY ruangan,kasuspenyakit"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    
    cetak = "grafikRekapkasuspenyakit"
    FrmViewerLaporan.Show
    FrmViewerLaporan.Caption = "Medifirst2000 - Grafik Rekapitulasi Berdasarkan Kasus Penyakit IGD"
    cmdCetak.Enabled = True

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub dcInstalasi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DTPickerAwal.SetFocus
End Sub

Private Sub DTPickerAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdcari.SetFocus
End Sub

Private Sub DTPickerAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DTPickerAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    txtJmlPria = "0": txtJmlTotal = "0": txtJmlWanita = "0"
    
    With Me
        .DTPickerAwal.Value = Now
        .DTPickerAkhir.Value = Now
    End With
    
    Call subDcSource
    Call subSetGrid
End Sub

Private Sub subSetGrid()
    With fgdata
        .Visible = False
        .clear
        .Cols = 6
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
        .TextMatrix(0, 2) = "Kasus Penyakit"
        .TextMatrix(0, 3) = "Laki-Laki"
        .TextMatrix(0, 4) = "Perempuan"
        .TextMatrix(0, 5) = "Total"
        
        .ColWidth(0) = 500
        .ColWidth(1) = 2850
        .ColWidth(2) = 2850
        .ColWidth(3) = 1100
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100
        
        .Visible = True
        iRowNow = 0
    End With
End Sub

Private Sub subDcSource()
    strSQL = "SELECT KdInstalasi, NamaInstalasi " & _
        " From instalasi" & _
        " WHERE (KdInstalasi IN ('01', '02', '03', '04', '06', '08', '09', '10', '16'))"
    Call msubDcSource(dcInstalasi, rs, strSQL)
End Sub


