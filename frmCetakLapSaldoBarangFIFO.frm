VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakLapSaldoBarangFIFO 
   Caption         =   "Medifirst 2000-Laporan Saldo Barang "
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5865
   Icon            =   "frmCetakLapSaldoBarangFIFO.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   5865
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakLapSaldoBarangFIFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crLapSaldoBarangFIFO
Dim Judul1 As String
Dim total As Currency
Dim adocomd As New ADODB.Command

Private Sub Form_Load()
    On Error GoTo errLoad
    Set adocomd = Nothing
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    If strCetak = "Hari" Or strCetak = "" Then
        Call LaporanPerHari
        Report.txtPeriode.SetText "Periode : " & Format(mdTglAwal, "dd MMMM yyyy") & " s/d " & Format(mdTglAkhir, " dd MMMM yyyy")
    ElseIf strCetak = "Bulan" Then
        Call LaporanPerBulan

        If Month(mdTglAwal) = Month(mdTglAkhir) Then
            Report.txtPeriode.SetText "Bulan : " & Format(mdTglAwal, " MMMM yyyy")
        Else
            Report.txtPeriode.SetText "Bulan : " & Format(mdTglAwal, "MMMM yyyy") & " s/d " & Format(mdTglAkhir, "MMMM yyyy")
        End If
    ElseIf strcetsk = "Tahun" Then
        Call LaporanPerTahun

        If Year(mdTglAwal) = Year(mdTglAkhir) Then
            Report.txtPeriode.SetText "Tahun : " & Format(mdTglAwal, "yyyy")
        Else
            Report.txtPeriode.SetText "Tahun : " & Format(mdTglAwal, "yyyy") & " s/d " & Format(mdTglAkhir, "yyyy")
        End If
    End If
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .usNamaBarang.SetUnboundFieldSource ("{ado.NamaBarang}")
        .unAwal.SetUnboundFieldSource ("{ado.SaldoAwal}")
        .unMasuk.SetUnboundFieldSource ("{ado.JmlTerima}")
        .unKeluar.SetUnboundFieldSource ("{ado.JmlKeluar}")
        .unSaldo.SetUnboundFieldSource Format(("{ado.SaldoAkhir}"), "###,##0")
        .ucNetto.SetUnboundFieldSource Format(("{ado.HargaNetto}"), "###,##0")
        If sKriteria = "JenisBarang" Then
            .txtJudulKriteria.SetText "Jenis Barang"
            .usJenisBarang.SetUnboundFieldSource ("{ado.JenisBarang}")
        ElseIf sKriteria = "KategoryBarang" Then
            .txtJudulKriteria.SetText "Kategory Barang"
            .usJenisBarang.SetUnboundFieldSource ("{ado.KategoryBarang}")
        ElseIf sKriteria = "GolonganBarang" Then
            .txtJudulKriteria.SetText "Golongan Barang"
            .usJenisBarang.SetUnboundFieldSource ("{ado.GolonganBarang}")
        ElseIf sKriteria = "StatusBarang" Then
            .txtJudulKriteria.SetText "Status Barang"
            .usJenisBarang.SetUnboundFieldSource ("{ado.StatusBarang}")
        ElseIf sKriteria = "NamaPabrik" Then
            .txtJudulKriteria.SetText "Nama Pabrik"
            .usJenisBarang.SetUnboundFieldSource ("{ado.Pabrik}")
        End If
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        .txtJudul.SetText Judul1
        .txtUser.SetText strNmPegawai
        .txtRuanganLogin.SetText mstrNamaRuangan
    End With
    Screen.MousePointer = vbHourglass

    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault

    Exit Sub
errLoad:
    Screen.MousePointer = vbDefault
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strSQL = ""
    strSQL = "Delete from LaporanSaldoBarangMedis_T where KdRuangan Like '%" & mstrKdRuangan & "%'"
    Call msubRecFO(rs, strSQL)
    Set frmCetakLapSaldoBarangFIFO = Nothing
End Sub

Private Sub LaporanPerHari()
    adocomd.ActiveConnection = dbConn
    Judul1 = "LAPORAN SALDO BARANG MEDIS (PER HARI)"
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText
End Sub

Private Sub LaporanPerBulan()
    adocomd.ActiveConnection = dbConn
    Judul1 = "LAPORAN SALDO BARANG MEDIS (PERBULAN)"
    tgl = ""
    tgl = funcHitungHari(Month(mdTglAkhir), Year(mdTglAkhir))
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText
End Sub

Private Sub LaporanPerTahun()
    adocomd.ActiveConnection = dbConn
    Judul1 = "LAPORAN SALDO BARANG MEDIS (PERTAHUN)"
    tgl = ""
    tgl = funcHitungHari(Month(mdTglAkhir), Year(mdTglAkhir))
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText
End Sub
