VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakLaporanRIperSMF 
   Caption         =   "Medifirst2000 - Daftar Pasien Rawat Jalan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakLaporanRIperSMF.frx":0000
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
      EnableGroupTree =   -1  'True
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
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakLaporanRIperSMF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crLaporanRIperSMF

Private Sub Form_Load()
    Dim intJmlRow As Integer
    Dim intJmlHidup As Integer
    Dim intJmlMati As Integer
    Dim intJmlTotal As Integer
    Dim adocomd As New ADODB.Command
    Call openConnection
    Set frmCetakLaporanRIperSMF = Nothing
    Me.WindowState = 2

    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = "SELECT NamaSubInstalasi, SUM(JmlMati) AS JmlMati, SUM(JmlHidup) AS JmlHidup From V_LaporanPasienRIPerSMFv2 WHERE (TglPulang BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') GROUP BY NamaSubInstalasi"
    adocomd.CommandType = adCmdText

    If mblnGrafik = False Then
        With Report
            .txtNamaRS.SetText strNNamaRS
            .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
            .txtWebsite.SetText strWebsite & ", " & strEmail

            If Format(mdTglAwal, "dd MMMM yyyy") = Format(mdTglAkhir, "dd MMMM yyyy") Then
                .txtTanggal.SetText "Tanggal Pulang  : " & " " & Format(mdTglAwal, "dd MMMM yyyy")
            Else
                .txtTanggal.SetText "Periode Pulang  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy")
            End If

            .Database.AddADOCommand dbConn, adocomd

            .usSMF.SetUnboundFieldSource ("{Ado.NamaSubInstalasi}")
            .unHidup.SetUnboundFieldSource ("{Ado.JmlHidup}")
            .unMati.SetUnboundFieldSource ("{Ado.JmlMati}")
        End With
        CRViewer1.ReportSource = Report
        CRViewer1.EnableGroupTree = False
    Else
        With reportgrafik
            .txtNamaRS.SetText strNNamaRS
            .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
            .txtWebsite.SetText strWebsite & ", " & strEmail

            If Format(mdTglAwal, "dd MMMM yyyy") = Format(mdTglAkhir, "dd MMMM yyyy") Then
                .txtTanggal.SetText "Tanggal Pulang  : " & " " & Format(mdTglAwal, "dd MMMM yyyy")
            Else
                .txtTanggal.SetText "Periode Pulang  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy")
            End If

            .Database.AddADOCommand dbConn, adocomd

            .usSMF.SetUnboundFieldSource ("{Ado.NamaSubInstalasi}")
            .unHidup.SetUnboundFieldSource ("{Ado.JmlHidup}")
            .unMati.SetUnboundFieldSource ("{Ado.JmlMati}")
        End With
        CRViewer1.ReportSource = reportgrafik
        CRViewer1.EnableGroupTree = False
    End If

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
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakLaporanRIperSMF = Nothing
    mblnGrafik = False
End Sub
