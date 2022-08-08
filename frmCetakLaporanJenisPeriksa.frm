VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakLaporanJenisPeriksa 
   Caption         =   "Medifirst2000 - Cetak"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakLaporanJenisPeriksa.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
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
Attribute VB_Name = "frmCetakLaporanJenisPeriksa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rptRkpPsnJenisPeriksa As New crRkpPsnJenisPeriksa
Dim rptRkpPsnJenisPeriksaGrafik As New crRkpPsnJenisPeriksaGrafik

Private Sub Form_Load()
Dim adocomd As New ADODB.Command
    On Error GoTo errLoad
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    'Laporan Rekapitulasi Pasien Per JenisPeriksa
    If mblnGrafik = False Then
        With rptRkpPsnJenisPeriksa
            .txtNamaRS.SetText strNNamaRS
            .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
            .txtAlamat2.SetText strWebSite & ", " & strEmail
            
            If Format(mdTglAwal, "dd MMMM yyyy") = Format(mdTglAkhir, "dd MMMM yyyy") Then
               .txtTanggal.SetText "Tanggal Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") '& " S/d " & Format(frmregister.DTPickerAkhir, "dd MMMM yyyy")
            Else
               .txtTanggal.SetText "Periode Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") & " S/d " & Format(mdTglAkhir, "dd MMMM yyyy")
            End If
            
            Set adocomd.ActiveConnection = dbConn
            adocomd.CommandText = "SELECT * FROM V_RekapitulasiKunjunganPasienBJenisPeriksa " _
                & "WHERE (TglPelayanan BETWEEN '" _
                & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
                & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " _
                & " AND KdRuangan = '" & mstrKdRuangan & "'  " _
                & "ORDER BY NamaRuangan,JenisPeriksa,JenisPasien"
            adocomd.CommandType = adCmdUnknown
            
            .Database.AddADOCommand dbConn, adocomd
'            .usKdRuangan.SetUnboundFieldSource ("{ado.KdRuangan}")
'            .usruangan.SetUnboundFieldSource ("{ado.NamaRuangan}")
'            .usDokter.SetUnboundFieldSource ("{ado.JenisPeriksa}")
'            .usJnsPsn.SetUnboundFieldSource ("{ado.JenisPasien}")
'            .unPria.SetUnboundFieldSource ("{ado.JmlPasienPria}")
'            .unWanita.SetUnboundFieldSource ("{ado.JmlPasienWanita}")
'            .unTotal.SetUnboundFieldSource ("{ado.Total}")
            
            settingreport rptRkpPsnJenisPeriksa, sPrinter, sDriver, sUkuranKertas, sDuplex, crPortrait
            CRViewer1.ReportSource = rptRkpPsnJenisPeriksa
        End With
    Else
        With rptRkpPsnJenisPeriksaGrafik
            .txtNamaRS.SetText strNNamaRS
            .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
            .txtAlamat2.SetText strWebSite & ", " & strEmail
            
            If Format(mdTglAwal, "dd MMMM yyyy") = Format(mdTglAkhir, "dd MMMM yyyy") Then
               .txtTanggal.SetText "Tanggal Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") '& " S/d " & Format(frmregister.DTPickerAkhir, "dd MMMM yyyy")
            Else
               .txtTanggal.SetText "Periode Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") & " S/d " & Format(mdTglAkhir, "dd MMMM yyyy")
            End If
            
            Set adocomd.ActiveConnection = dbConn
            adocomd.CommandText = "SELECT * FROM V_RekapitulasiKunjunganPasienBJenisPeriksa " _
                & "WHERE (TglPelayanan BETWEEN '" _
                & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
                & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " _
                & " AND KdRuangan = '" & mstrKdRuangan & "'  " _
                & "ORDER BY NamaRuangan,JenisPeriksa,JenisPasien"
            adocomd.CommandType = adCmdUnknown
            
            .Database.AddADOCommand dbConn, adocomd
            .usKdRuangan.SetUnboundFieldSource ("{ado.KdRuangan}")
            .usruangan.SetUnboundFieldSource ("{ado.NamaRuangan}")
            .usDokter.SetUnboundFieldSource ("{ado.JenisPeriksa}")
            .unJmlPria.SetUnboundFieldSource ("{ado.JmlPasienPria}")
            .unJmlWanita.SetUnboundFieldSource ("{ado.JmlPasienWanita}")
'                .unTotal.SetUnboundFieldSource ("{ado.Total}")
            
            settingreport rptRkpPsnJenisPeriksaGrafik, sPrinter, sDriver, sUkuranKertas, sDuplex, crPortrait
            CRViewer1.ReportSource = rptRkpPsnJenisPeriksaGrafik
        End With
    End If
    
    With CRViewer1
        .Zoom 1 ' Set the zoom level to fit the page width to the viewer window
        .ViewReport ' Set the viewer to view the report
'        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
errLoad:
    Screen.MousePointer = vbDefault
    msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakLaporanJenisPeriksa = Nothing
    mblnGrafik = False
End Sub
