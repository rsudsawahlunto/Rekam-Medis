VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmPeriodeIndikatorRS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Indikator Pelayanan Rumah Sakit"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPeriodeIndikatorRS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10515
   Begin VB.Frame Frame1 
      Caption         =   "Indikator Pelayanan Rumah Sakit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5115
      Left            =   0
      TabIndex        =   9
      Top             =   930
      Width           =   10515
      Begin VB.ComboBox cboKriteria 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmPeriodeIndikatorRS.frx":0CCA
         Left            =   240
         List            =   "FrmPeriodeIndikatorRS.frx":0CD4
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   550
         Width           =   2805
      End
      Begin VB.Frame Frame5 
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
         Left            =   4560
         TabIndex        =   14
         Top             =   150
         Width           =   5775
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
            CustomFormat    =   "dd MMMM, yyyy"
            Format          =   127336451
            UpDown          =   -1  'True
            CurrentDate     =   37956
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
            Format          =   127336451
            UpDown          =   -1  'True
            CurrentDate     =   37956
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   15
            Top             =   315
            Width           =   255
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgdata 
         Height          =   3825
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   6747
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Kriteria Indikator"
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   280
         Width           =   1335
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
      TabIndex        =   8
      Top             =   6000
      Width           =   10510
      Begin VB.CommandButton cmdgrafik 
         Caption         =   "&Grafik"
         Height          =   375
         Left            =   4320
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   7080
         TabIndex        =   5
         Top             =   240
         Width           =   1545
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   8760
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Jumlah        Pria           Wanita         Total"
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   6120
      Visible         =   0   'False
      Width           =   4155
      Begin VB.TextBox txtJmlTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2970
         TabIndex        =   13
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox txtJmlWanita 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1890
         TabIndex        =   12
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox txtJmlPria 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   810
         TabIndex        =   11
         Top             =   240
         Width           =   1000
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   17
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
      Left            =   8640
      Picture         =   "FrmPeriodeIndikatorRS.frx":0CF0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "FrmPeriodeIndikatorRS.frx":1A78
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "FrmPeriodeIndikatorRS.frx":30D6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "FrmPeriodeIndikatorRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iRowNow As Integer
Dim rsstatusPasien As ADODB.recordset
Dim rsstatusPasien1 As ADODB.recordset
Dim iRowNow2 As Integer

Private Sub cboKriteria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then DTPickerAwal.SetFocus
End Sub

Private Sub cmdCari_Click()
    On Error GoTo hell

    Dim intJmlRow As Integer
    Dim intJmlPria As Integer
    Dim intJmlWanita As Integer
    Dim intJmlTotal As Integer
    If cboKriteria.Text = "" Then MsgBox "Pilih Kriteria Indikator", vbCritical, "Validasi": Exit Sub
    Call subSetGrid
    'u/ mempercepat
    fgData.Visible = False: MousePointer = vbHourglass

    If cboKriteria.Text = "Per Ruangan" Then
        'Hitung jumlah row dari data yang hendak ditampilkan
        strSQL = "SELECT COUNT(tglhitung) AS JmlRow " & _
        " FROM v_S_RekapIndikatorPlyn " & _
        " WHERE tglhitung BETWEEN " & _
        " '" & Format(DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' "

        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        'jika tidak ada data
        If rs(0).value = 0 Then
            fgData.Visible = True: MousePointer = vbNormal
            MsgBox "Tidak ada Data"
            txtJmlPria = "0": txtJmlTotal = "0": txtJmlWanita = "0"
            Exit Sub
        End If

        intJmlRow = rs("JmlRow").value

        strSQL = "SELECT namaruangan, " & _
        " round(AVG(JmlTOI),2) AS JmlTOI,round(AVG(JmlBOR),2) AS JmlBOR," & _
        " round(AVG(JmlBTO),2) AS JmlBTO, round(AVG(JmlLOS),2)AS JmlLOS," & _
        " round(AVG(JmlGDR),2) AS JmlGDR, round(AVG(JmlNDR),2) AS JmlNDR, SUM(JmlPasien) AS JmlPasien From v_S_RekapIndikatorPlyn " & _
        " WHERE TglHitung BETWEEN " & _
        " '" & Format(DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' " & _
        " GROUP BY namaruangan"

        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        With fgData
            'jml baris akhir
            .Rows = intJmlRow
            While rs.EOF = False
                'baris u/ sub total
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = rs("namaruangan").value
                .TextMatrix(iRowNow, 2) = rs("JmlTOI").value & " " & "hari"
                .TextMatrix(iRowNow, 3) = rs("JmlBOR").value & " " & "%"
                .TextMatrix(iRowNow, 4) = rs("JmlBTO").value & " " & "kali"
                .TextMatrix(iRowNow, 5) = rs("JmlLOS").value & " " & "hari"
                .TextMatrix(iRowNow, 6) = rs("JmlGDR").value & " " & "‰ "
                .TextMatrix(iRowNow, 7) = rs("JmlNDR").value & " " & "‰ "
                .TextMatrix(iRowNow, 8) = rs("JmlPasien").value
                rs.MoveNext
            Wend
            'banyak baris berdasarkan irownow
            .Rows = iRowNow + 2

            .Col = 1
            For i = 1 To .Rows - 1
                .Row = i
                .CellFontBold = True
            Next

            .Visible = True: MousePointer = vbNormal
        End With

    ElseIf cboKriteria.Text = "Per Kelas" Then
        'Hitung jumlah row dari data yang hendak ditampilkan
        strSQL = "SELECT COUNT(tglhitung) AS JmlRow " & _
        " FROM v_S_RekapIndikatorPlyn " & _
        " WHERE tglhitung BETWEEN " & _
        " '" & Format(DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' "

        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        'jika tidak ada data
        If rs(0).value = 0 Then
            fgData.Visible = True: MousePointer = vbNormal
            MsgBox "Tidak ada Data"
            '        MsgBox "Tidak ada data antara tanggal  '" & Format(DTPickerAwal.Value, "dd - MMMM - yyyy") & "' dan '" & Format(dtpTglAkhir.Value, "dd - MMMM - yyyy") & "' ", vbExclamation, "Validasi"
            txtJmlPria = "0": txtJmlTotal = "0": txtJmlWanita = "0"
            Exit Sub
        End If

        intJmlRow = rs("JmlRow").value

        strSQL = "SELECT DeskKelas, " & _
        " round(AVG(JmlTOI),2) AS JmlTOI,round(AVG(JmlBOR),2) AS JmlBOR," & _
        " round(AVG(JmlBTO),2) AS JmlBTO, round(AVG(JmlLOS),2)AS JmlLOS," & _
        " round(AVG(JmlGDR),2) AS JmlGDR, round(AVG(JmlNDR),2) AS JmlNDR, SUM(JmlPasien) AS JmlPasien From v_S_RekapIndikatorPlyn " & _
        " WHERE TglHitung BETWEEN " & _
        " '" & Format(DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' " & _
        " GROUP BY DeskKelas"

        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        With fgData
            'jml baris akhir
            .Rows = intJmlRow
            While rs.EOF = False
                'baris u/ sub total
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = rs("DeskKelas").value
                .TextMatrix(iRowNow, 2) = rs("JmlTOI").value & " " & "hari"
                .TextMatrix(iRowNow, 3) = rs("JmlBOR").value & " " & "%"
                .TextMatrix(iRowNow, 4) = rs("JmlBTO").value & " " & "kali"
                .TextMatrix(iRowNow, 5) = rs("JmlLOS").value & " " & "hari"
                .TextMatrix(iRowNow, 6) = rs("JmlGDR").value & " " & "‰"
                .TextMatrix(iRowNow, 7) = rs("JmlNDR").value & " " & "‰"
                .TextMatrix(iRowNow, 8) = rs("JmlPasien").value
                rs.MoveNext
            Wend
            'banyak baris berdasarkan irownow
            .Rows = iRowNow + 2

            .Col = 1
            For i = 1 To .Rows - 1
                .Row = i
                .CellFontBold = True
            Next

            .Visible = True: MousePointer = vbNormal
        End With

    ElseIf cboKriteria.Text = "Semua" Then
        '    'Hitung jumlah row dari data yang hendak ditampilkan
        strSQL = "SELECT COUNT(tglhitung) AS JmlRow " & _
        " FROM RekapitulasiIndikatorPelayananRS " & _
        " WHERE tglhitung BETWEEN " & _
        " '" & Format(DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' "

        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        'jika tidak ada data
        If rs(0).value = 0 Then
            fgData.Visible = True: MousePointer = vbNormal
            MsgBox "Tidak ada Data"
            Exit Sub
        End If

        intJmlRow = rs("JmlRow").value

        strSQL = "SELECT  round(AVG(JmlTOI),2) AS JmlTOI,round(AVG(JmlBOR),2) AS JmlBOR," & _
        " round(AVG(JmlBTO),2) AS JmlBTO, round(AVG(JmlLOS),2)AS JmlLOS," & _
        " round(AVG(JmlGDR),2) AS JmlGDR, round(AVG(JmlNDR),2) AS JmlNDR From RekapitulasiIndikatorPelayananRS " & _
        " WHERE TglHitung BETWEEN " & _
        " '" & Format(DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' "

        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        With fgData
            'jml baris akhir
            .Rows = intJmlRow '+ 1
            While rs.EOF = False
                'baris u/ sub total
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = rs("JmlTOI").value & " " & "hari"
                .TextMatrix(iRowNow, 2) = rs("JmlBOR").value & " " & "%"
                .TextMatrix(iRowNow, 3) = rs("JmlBTO").value & " " & "kali"
                .TextMatrix(iRowNow, 4) = rs("JmlLOS").value & " " & "hari"
                .TextMatrix(iRowNow, 5) = rs("JmlGDR").value & " " & "‰"
                .TextMatrix(iRowNow, 6) = rs("JmlNDR").value & " " & "‰"
                rs.MoveNext
            Wend
            'banyak baris berdasarkan irownow
            .Rows = iRowNow + 2

            .Col = 1
            For i = 1 To .Rows - 1
                .Row = i
                .CellFontBold = True
            Next

            .Visible = True: MousePointer = vbNormal
        End With

    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    If cboKriteria.Text = "Semua" Then
        strSQL = "SELECT AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR, SUM(JmlPasien) AS JmlPasien " _
        & "FROM RekapitulasiIndikatorPelayananRS " _
        & "WHERE TglHitung BETWEEN '" _
        & Format(FrmPeriodeIndikatorRS.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(FrmPeriodeIndikatorRS.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "' "
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount = 0 Then
            MsgBox "Tidak ada data", vbCritical, "Validasi"
            cmdCetak.Enabled = True
            Exit Sub
        End If
        cetak = "Semua"
        frmUtilitasRS2.Show
    ElseIf (cboKriteria.Text = "Per Ruangan") Then
        strSQL = "SELECT NamaRuangan AS Ruangan,AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR, SUM(JmlPasien) AS JmlPasien " _
        & "FROM dbo.v_S_RekapIndikatorPlyn " _
        & "WHERE TglHitung BETWEEN '" _
        & Format(FrmPeriodeIndikatorRS.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(FrmPeriodeIndikatorRS.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "' " _
        & "GROUP BY NamaRuangan"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount = 0 Then
            MsgBox "Tidak ada data", vbCritical, "Validasi"
            cmdCetak.Enabled = True
            Exit Sub
        End If
        cetak = "PerRuangan"
        frmUtilitasRS.Show
    ElseIf (cboKriteria.Text = "Per Kelas") Then
        strSQL = "SELECT DeskKelas AS Kelas,AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR, SUM(JmlPasien) AS JmlPasien " _
        & "FROM dbo.v_S_RekapIndikatorPlyn " _
        & "WHERE TglHitung BETWEEN '" _
        & Format(FrmPeriodeIndikatorRS.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(FrmPeriodeIndikatorRS.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "' " _
        & "GROUP BY DeskKelas"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount = 0 Then
            MsgBox "Tidak ada data", vbCritical, "Validasi"
            cmdCetak.Enabled = True
            Exit Sub
        End If
        cetak = "PerKelas"
        frmUtilitasRS.Show
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdgrafik_Click()
    On Error GoTo hell
    If cboKriteria.Text = "Semua" Then
        strSQL = "SELECT AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR " _
        & "FROM RekapitulasiIndikatorPelayananRS " _
        & "WHERE TglHitung BETWEEN '" _
        & Format(FrmPeriodeIndikatorRS.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(FrmPeriodeIndikatorRS.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "' "
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount = 0 Then
            MsgBox "Tidak ada data", vbCritical, "Validasi"
            cmdCetak.Enabled = True
            Exit Sub
        End If
        cetak = "GrafikSemua"
        frmUtilitasRS2.Show
    ElseIf (cboKriteria.Text = "Per Ruangan") Then
        strSQL = "SELECT NamaRuangan AS Ruangan,AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR " _
        & "FROM dbo.v_S_RekapIndikatorPlyn " _
        & "WHERE TglHitung BETWEEN '" & Format(FrmPeriodeIndikatorRS.DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(FrmPeriodeIndikatorRS.DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' " _
        & "GROUP BY NamaRuangan"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount = 0 Then
            MsgBox "Tidak ada data", vbCritical, "Validasi"
            cmdCetak.Enabled = True
            Exit Sub
        End If
        cetak = "GrafikPerRuangan"
        frmUtilitasRS.Show
    ElseIf (cboKriteria.Text = "Per Kelas") Then
        strSQL = "SELECT DeskKelas AS Kelas,AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR " _
        & "FROM dbo.v_S_RekapIndikatorPlyn " _
        & "WHERE TglHitung BETWEEN '" & Format(FrmPeriodeIndikatorRS.DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(FrmPeriodeIndikatorRS.DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' " _
        & "GROUP BY DeskKelas"
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount = 0 Then
            MsgBox "Tidak ada data", vbCritical, "Validasi"
            cmdCetak.Enabled = True
            Exit Sub
        End If
        cetak = "GrafikPerkelas"
        frmUtilitasRS.Show
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
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

    Call subSetGrid
End Sub

Private Sub subSetGrid()
    If cboKriteria.Text = "Per Ruangan" Then
        With fgData
            .Visible = False
            .clear
            .Cols = 9
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

            .TextMatrix(0, 1) = "Ruang Pelayanan"
            .TextMatrix(0, 2) = "TOI"
            .TextMatrix(0, 3) = "BOR"
            .TextMatrix(0, 4) = "BTO"
            .TextMatrix(0, 5) = "LOS"
            .TextMatrix(0, 6) = "GDR"
            .TextMatrix(0, 7) = "NDR"
            .TextMatrix(0, 8) = "Jml Pasien"

            .ColWidth(0) = 500
            .ColWidth(1) = 2850
            .ColWidth(2) = 1100
            .ColWidth(3) = 1100
            .ColWidth(4) = 1100
            .ColWidth(5) = 1100
            .ColWidth(6) = 1100
            .ColWidth(7) = 1100
            .ColWidth(8) = 1100

            .Visible = True
            iRowNow = 0
        End With
    ElseIf cboKriteria.Text = "Per Kelas" Then
        With fgData
            .Visible = False
            .clear
            .Cols = 9
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

            .TextMatrix(0, 1) = "Kelas Pelayanan"
            .TextMatrix(0, 2) = "TOI"
            .TextMatrix(0, 3) = "BOR"
            .TextMatrix(0, 4) = "BTO"
            .TextMatrix(0, 5) = "LOS"
            .TextMatrix(0, 6) = "GDR"
            .TextMatrix(0, 7) = "NDR"
            .TextMatrix(0, 8) = "Jml Pasien"

            .ColWidth(0) = 500
            .ColWidth(1) = 2850
            .ColWidth(2) = 1100
            .ColWidth(3) = 1100
            .ColWidth(4) = 1100
            .ColWidth(5) = 1100
            .ColWidth(6) = 1100
            .ColWidth(7) = 1100
            .ColWidth(8) = 1100

            .Visible = True
            iRowNow = 0
        End With

    ElseIf cboKriteria.Text = "Semua" Then
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

            .TextMatrix(0, 1) = "TOI"
            .TextMatrix(0, 2) = "BOR"
            .TextMatrix(0, 3) = "BTO"
            .TextMatrix(0, 4) = "LOS"
            .TextMatrix(0, 5) = "GDR"
            .TextMatrix(0, 6) = "NDR"

            .ColWidth(0) = 1100
            .ColWidth(1) = 1100
            .ColWidth(2) = 1100
            .ColWidth(3) = 1100
            .ColWidth(4) = 1100
            .ColWidth(5) = 1100
            .ColWidth(6) = 1100

            .Visible = True
            iRowNow = 0
        End With

    End If
End Sub

