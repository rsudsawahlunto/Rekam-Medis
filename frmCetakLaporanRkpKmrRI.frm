VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakLaporanRkpKmrRI 
   Caption         =   "Medifirst2000 - Cetak"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakLaporanRkpKmrRI.frx":0000
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
Attribute VB_Name = "frmCetakLaporanRkpKmrRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReportRekKamarRI As New cr_RekKamarRI
Dim ReportRekKamarRIGrafik As New cr_RekKamarRIGrafik

Private Sub Form_Load()
Dim tanggal As String
Dim laporan As String
Dim adocomd As New ADODB.Command
    On Error GoTo errLoad
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    'Laporan Rekapitulasi Pasien Per Dokter
    If mblnGrafik = False Then
        strSQL = "select * from V_RekapitulasiKamarRawatInap " _
            & " WHERE DAY(TglHitung) = '" & frmLapRkpKmrRI.dtpAwal.Day & "' " _
            & " AND MONTH(TglHitung) = '" & frmLapRkpKmrRI.dtpAwal.Month & "' " _
            & " AND YEAR(TglHitung) = '" & frmLapRkpKmrRI.dtpAwal.Year & "' " _
            & " "
                
        Set rs = Nothing
        rs.Open strSQL, dbConn, , adLockOptimistic
        If rs.EOF Then
            MsgBox "Data Tidak Ada", vbInformation, "Informasi"
            Exit Sub
        End If

        Call openConnection
            adocomd.ActiveConnection = dbConn
            adocomd.CommandText = "SELECT DISTINCT NoKamar,TglHitung,Ruangan,Kelas,jmlbedterisi,jmlbedkosong,totalbed from v_rekapitulasikamarrawatinap" _
                & " WHERE DAY(TglHitung) = '" & frmLapRkpKmrRI.dtpAwal.Day & "' " _
                & " AND MONTH(TglHitung) = '" & frmLapRkpKmrRI.dtpAwal.Month & "' " _
                & " AND YEAR(TglHitung) = '" & frmLapRkpKmrRI.dtpAwal.Year & "' "
        
           adocomd.CommandType = adCmdText
           ReportRekKamarRI.Database.AddADOCommand dbConn, adocomd
           tanggal = "Tanggal : " & " " & Format(frmLapRkpKmrRI.dtpAwal.Value, "dd MMMM yyyy")
        
        With ReportRekKamarRI
            .Text1.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .Text2.SetText "KABUPATEN " & strNKotaRS
            .Text3.SetText strNAlamatRS & " " & "Telp." & " " & strNTeleponRS
            .txtTanggal.SetText tanggal
            .TxtRuangan.SetText strNNamaRuangan
            .usKelas.SetUnboundFieldSource ("{ado.Kelas}")
            .usNoKamar.SetUnboundFieldSource ("{ado.NoKamar}")
            .udTglHitung.SetUnboundFieldSource ("{ado.TglHitung}")
            .unJmlBedIsi.SetUnboundFieldSource ("{ado.JmlBedTerisi}")
            .unJmlBedKosong.SetUnboundFieldSource ("{ado.JmlBedKosong}")
            .unTotalBed.SetUnboundFieldSource ("{ado.TotalBed}")
            '.Text9.SetText NamaPegawai
            .SelectPrinter sDriver, sPrinter, vbNull
            settingreport ReportRekKamarRI, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
        End With
        CRViewer1.ReportSource = ReportRekKamarRI
    Else
        With ReportRekKamarRIGrafik
            .txtNamaRS.SetText strNNamaRS
            .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
            .txtAlamat2.SetText strWebsite & ", " & strEmail
            
            .txtTanggal.SetText "Tanggal Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy")
            
            Set adocomd.ActiveConnection = dbConn
            adocomd.CommandText = "SELECT * FROM V_RekapitulasiKamarRawatInap" _
                & " WHERE DAY(TglHitung) = '" & frmLapRkpKmrRI.dtpAwal.Day & "' " _
                & " AND MONTH(TglHitung) = '" & frmLapRkpKmrRI.dtpAwal.Month & "' " _
                & " AND YEAR(TglHitung) = '" & frmLapRkpKmrRI.dtpAwal.Year & "' " _
                & "ORDER BY Ruangan,Kelas,NoKamar"
            adocomd.CommandType = adCmdUnknown
            
            .Database.AddADOCommand dbConn, adocomd
            .usKdRuangan.SetUnboundFieldSource ("{ado.KdRuangan}")
            .usRuangan.SetUnboundFieldSource ("{ado.Ruangan}")
            .usDokter.SetUnboundFieldSource ("{ado.Kelas}")
            .unJmlPria.SetUnboundFieldSource ("{ado.JmlBedTerisi}")
            .unJmlWanita.SetUnboundFieldSource ("{ado.JmlBedKosong}")
'                .unTotal.SetUnboundFieldSource ("{ado.Total}")
            
            settingreport ReportRekKamarRIGrafik, sPrinter, sDriver, sUkuranKertas, sDuplex, crPortrait
            CRViewer1.ReportSource = ReportRekKamarRIGrafik
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
    Set frmCetakLaporanRkpKmrRI = Nothing
    mblnGrafik = False
End Sub
