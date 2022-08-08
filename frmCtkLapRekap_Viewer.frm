VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCtkLapRekap_Viewer 
   Caption         =   "CETAK"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCtkLapRekap_Viewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
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
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCtkLapRekap_Viewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As String, b As String, c As String
Dim Report As CRAXDRT.Report

Private Sub Form_Load()
Dim iRet As Integer
'    On Error GoTo errPrint
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Select Case strCetak
    Case "LapRekapKPSJ"
        Set Report = CROpenMSQLReport(App.Path & "\report\cr_LapKunjunganPasienSJ.rpt", strServerName, strDatabaseName, "v_S_RekapKunjunganPsnSJ", True)
        Report.RecordSelectionFormula = "{v_S_RekapKunjunganPsnSJ.KdInstalasi} = '" & mstrInstalasi & "' and {v_S_RekapKunjunganPsnSJ.TglPendaftaran} IN DateTime (" & Year(mdTglAwal) & ", " & Month(mdTglAwal) & ", " & Day(mdTglAwal) & ", 0, 0, 0) TO DateTime (" & Year(mdTglAkhir) & ", " & Month(mdTglAkhir) & ", " & Day(mdTglAkhir) & ", 23, 59, 59)"
                                        '{v_S_RekapKunjunganPsnSJ.TglPendaftaran} in DateTime (2004, 09, 20, 16, 27, 29) to DateTime (2005, 10, 28, 13, 53, 21) and {v_S_RekapKunjunganPsnSJ.KdInstalasi} = "02"
        Call subDataRSU
    Case "LapRekapKPSR"
        Set Report = CROpenMSQLReport(App.Path & "\report\cr_LapKunjunganPasienSR.rpt", strServerName, strDatabaseName, "v_S_RekapKunjunganPsnSR", True)
        Report.RecordSelectionFormula = "{v_S_RekapKunjunganPsnSR.KdInstalasi} = '" & mstrInstalasi & "' and {v_S_RekapKunjunganPsnSR.TglPendaftaran} IN DateTime (" & Year(mdTglAwal) & ", " & Month(mdTglAwal) & ", " & Day(mdTglAwal) & ", 0, 0, 0) TO DateTime (" & Year(mdTglAkhir) & ", " & Month(mdTglAkhir) & ", " & Day(mdTglAkhir) & ", 23, 59, 59)"
        Call subDataRSU
    End Select
'    iRet = frmPrinter.SelectPrinter(Report)
'    Report.PrintOut False*
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
    CRViewer1.Zoom (100)
    Screen.MousePointer = 0
    Exit Sub
errPrint:
    MsgBox "Error cetak!" & vbNewLine & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_Resize()
    CRViewer1.Width = Me.ScaleWidth
    CRViewer1.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCtkLapRekap_Viewer = Nothing
End Sub

Private Sub subDataRSU()
    On Error Resume Next
    modReport.CRFormula(Report, "NamaRS") = strNNamaRS
    modReport.CRFormula(Report, "AlamatRS") = strNAlamatRS & " " & "Telp." & " " & strNTeleponRS
    modReport.CRFormula(Report, "Tanggal") = Format(mdTglAwal, "dd/mm/yyyy") & " s.d. " & Format(mdTglAkhir, "dd/mm/yyyy")
    modReport.CRFormula(Report, "User") = strUser
End Sub
