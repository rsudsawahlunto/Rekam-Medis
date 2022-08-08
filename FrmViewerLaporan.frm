VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmViewerLaporan1 
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   Icon            =   "FrmViewerLaporan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   5850
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
End
Attribute VB_Name = "FrmViewerLaporan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim reportTopten As New crDiagnosaTopTen21
Dim reporrtoptengrafik As New crDiagnosaTopTenGrafik1
Dim adocomd As New ADODB.Command
Dim tanggal As String

Private Sub Form_Load()
    Me.WindowState = 2

    Set adocomd = New ADODB.Command
    adocomd.ActiveConnection = dbConn

    Dim tanggal As String

    Select Case cetak
        Case "RekapTopten"
            Call RekapTopTen
        Case "RekapToptenGrafik"
            Call RekapTopTenGrafik
    End Select
End Sub

Private Sub RekapTopTen()

    Set reportTopten = New crDiagnosaTopTen21

    If FrmPeriodeLaporanTopTen1.optKodeDiagnosa.value = True Then
        adocomd.CommandText = "SELECT top " & Val(FrmPeriodeLaporanTopTen1.txtJmlData) & " Diagnosa," & _
        " dbo.JmlPasienPerDiagnosa('" & Format(FrmPeriodeLaporanTopTen1.DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "','" & Format(FrmPeriodeLaporanTopTen1.DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "',KdDiagnosa,'" & mstrFilter2 & "') AS [JmlPasien], sum(jumlahpasien) as [JmlKunjungan] " & _
        " FROM V_RekapitulasiDiagnosaTopTenALL " & _
        " WHERE TglPeriksa BETWEEN " & _
        " '" & Format(FrmPeriodeLaporanTopTen1.DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(FrmPeriodeLaporanTopTen1.DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND kdinstalasi " & mstrFilter & " AND KdRuangan LIKE '%" & FrmPeriodeLaporanTopTen1.dcRuangPoli.BoundText & "%' AND " & _
        " JenisPasien LIKE '%" & FrmPeriodeLaporanTopTen1.dcJenisPasien & "%' group by KdDiagnosa,Diagnosa order by Diagnosa asc"
    Else
        adocomd.CommandText = "SELECT top " & Val(FrmPeriodeLaporanTopTen1.txtJmlData) & " Diagnosa," & _
        " dbo.JmlPasienPerDiagnosa('" & Format(FrmPeriodeLaporanTopTen1.DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "','" & Format(FrmPeriodeLaporanTopTen1.DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "',KdDiagnosa,'" & mstrFilter2 & "') AS [JmlPasien], sum(jumlahpasien) as [JmlKunjungan] " & _
        " FROM V_RekapitulasiDiagnosaTopTenALL " & _
        " WHERE TglPeriksa BETWEEN " & _
        " '" & Format(FrmPeriodeLaporanTopTen1.DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(FrmPeriodeLaporanTopTen1.DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND kdinstalasi " & mstrFilter & " AND KdRuangan LIKE '%" & FrmPeriodeLaporanTopTen1.dcRuangPoli.BoundText & "%' AND " & _
        " JenisPasien LIKE '%" & FrmPeriodeLaporanTopTen1.dcJenisPasien & "%' group by KdDiagnosa,Diagnosa order by [JmlPasien] desc"
    End If

    adocomd.CommandType = adCmdText
    reportTopten.Database.AddADOCommand dbConn, adocomd

    If Format(FrmPeriodeLaporanTopTen1.DTPickerAwal, "dd MMMM yyyy") = Format(FrmPeriodeLaporanTopTen1.DTPickerAkhir, "dd MMMM yyyy") Then
        tanggal = "Tanggal Kunjungan  : " & " " & Format(FrmPeriodeLaporanTopTen1.DTPickerAwal, "dd MMMM yyyy") '& " S/d " & Format(FrmPeriodeLaporanTopTen1.DTPickerAkhir, "dd MMMM yyyy")
    Else
        tanggal = "Periode Kunjungan  : " & " " & Format(FrmPeriodeLaporanTopTen1.DTPickerAwal, "dd MMMM yyyy") & " S/d " & Format(FrmPeriodeLaporanTopTen1.DTPickerAkhir, "dd MMMM yyyy")
    End If

    With reportTopten
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        .txtJudul.SetText "DAFTAR " & FrmPeriodeLaporanTopTen1.txtJmlData & " BESAR PENYAKIT PASIEN"

        If FrmPeriodeLaporanTopTen1.optKodeDiagnosa.value = True Then
            .txtKet.SetText "BERDASARKAN DIAGNOSA"
        Else
            .txtKet.SetText "BERDASARKAN JUMLAH PASIEN"
        End If

        .txtPeriode2.SetText tanggal
        .txtInstalasi.SetText FrmPeriodeLaporanTopTen1.dcInstalasi.Text
        .txtRuangInstalasi.SetText FrmPeriodeLaporanTopTen1.dcRuangPoli.Text
        .txtJenisPasien.SetText FrmPeriodeLaporanTopTen1.dcJenisPasien
        .usDiagnosa.SetUnboundFieldSource ("{ado.Diagnosa}")
        .unJmlKunj.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .unJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")
        .SelectPrinter sDriver, sPrinter, vbNull
        settingreport reportTopten, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = reportTopten
            .ViewReport
            .Zoom (100)
        End With
    Else
        reportTopten.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub RekapTopTenGrafik()

    Set reporrtoptengrafik = New crDiagnosaTopTenGrafik1

    If FrmPeriodeLaporanTopTen1.optKodeDiagnosa.value = True Then
        adocomd.CommandText = "SELECT top " & Val(FrmPeriodeLaporanTopTen1.txtJmlData) & " Diagnosa, sum(jumlahpasien) as [JmlPasien]" & _
        " FROM V_RekapitulasiDiagnosaTopTen " & _
        " WHERE TglPeriksa BETWEEN " & _
        " '" & Format(FrmPeriodeLaporanTopTen1.DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(FrmPeriodeLaporanTopTen1.DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND kdinstalasi " & mstrFilter & " AND KdRuangan LIKE '%" & FrmPeriodeLaporanTopTen1.dcRuangPoli.BoundText & "%' AND " & _
        " JenisPasien LIKE '%" & FrmPeriodeLaporanTopTen1.dcJenisPasien & "%' group by Diagnosa order by Diagnosa asc"
    Else
        adocomd.CommandText = "SELECT top " & Val(FrmPeriodeLaporanTopTen1.txtJmlData) & " Diagnosa, sum(jumlahpasien) as [JmlPasien]" & _
        " FROM V_RekapitulasiDiagnosaTopTen " & _
        " WHERE TglPeriksa BETWEEN " & _
        " '" & Format(FrmPeriodeLaporanTopTen1.DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(FrmPeriodeLaporanTopTen1.DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND kdinstalasi " & mstrFilter & " AND KdRuangan LIKE '%" & FrmPeriodeLaporanTopTen1.dcRuangPoli.BoundText & "%' AND " & _
        " JenisPasien LIKE '%" & FrmPeriodeLaporanTopTen1.dcJenisPasien & "%' group by Diagnosa order by [JmlPasien] desc"
    End If

    adocomd.CommandType = adCmdText
    reporrtoptengrafik.Database.AddADOCommand dbConn, adocomd

    If Format(FrmPeriodeLaporanTopTen1.DTPickerAwal, "dd MMMM yyyy") = Format(FrmPeriodeLaporanTopTen1.DTPickerAkhir, "dd MMMM yyyy") Then
        tanggal = "Tanggal Kunjungan  : " & " " & Format(FrmPeriodeLaporanTopTen1.DTPickerAwal, "dd MMMM yyyy") '& " S/d " & Format(FrmPeriodeLaporanTopTen1.DTPickerAkhir, "dd MMMM yyyy")
    Else
        tanggal = "Periode Kunjungan  : " & " " & Format(FrmPeriodeLaporanTopTen1.DTPickerAwal, "dd MMMM yyyy") & " S/d " & Format(FrmPeriodeLaporanTopTen1.DTPickerAkhir, "dd MMMM yyyy")
    End If

    With reporrtoptengrafik
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        .txtJudul.SetText "DAFTAR '" & FrmPeriodeLaporanTopTen1.txtJmlData & "' BESAR PENYAKIT PASIEN"
        .Graph1.Data2LabelFont.Size = 6
        .Graph1.Data2TitleFont.Size = 6
        .Graph1.DataLabelFont.Size = 6
        .Graph1.DataTitleFont.Size = 6
        .Graph1.FootnoteFont.Size = 6
        .Graph1.GroupLabelFont.Size = 6
        .Graph1.GroupTitleFont.Size = 6
        .Graph1.SeriesLabelFont.Size = 6
        .Graph1.SeriesTitleFont.Size = 6
        .Graph1.SubTitleFont.Size = 6
        .Graph1.TitleFont.Size = 6
        .txtPeriode2.SetText tanggal
        .txtInstalasi.SetText FrmPeriodeLaporanTopTen1.dcInstalasi.Text
        .txtRuangInstalasi.SetText FrmPeriodeLaporanTopTen1.dcRuangPoli.Text
        .txtJenisPasien.SetText FrmPeriodeLaporanTopTen1.dcJenisPasien
        .usDiagnosa.SetUnboundFieldSource ("{ado.diagnosa}")
        .unJumlahPasien.SetUnboundFieldSource ("{ado.JmlPasien}")
        .SelectPrinter sDriver, sPrinter, vbNull
        settingreport reporrtoptengrafik, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = reporrtoptengrafik
            .ViewReport
            .Zoom (100)
        End With
    Else
        reporrtoptengrafik.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmViewerLaporan = Nothing
End Sub

