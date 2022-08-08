VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakPemanfaatanRS 
   Caption         =   "LAPORAN SENSUS HARIAN"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
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
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakPemanfaatanRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'by splakuk 2008-12-03
Dim Report As New crCetakDataPemanfaatanRS

Private Sub Form_Load()

    Dim adocomd As New ADODB.Command
    adocomd.ActiveConnection = dbConn

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    
    Set adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, adocomd
    
    With Report
       
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strWebsite & ", " & strEmail
        .txtPeriode.SetText "Periode: " & Format(frmLapDataPemanfaatanRS.dtpAwal.Value, "yyyy")
                
        .txtjudul.SetText "LAPORAN DATA PEMANFAATAN RUMAH SAKIT (RAWAT INAP)"
        
        .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
        .usBulan.SetUnboundFieldSource ("{ado.Bulan}")
        .usjudul.SetUnboundFieldSource ("{ado.Judul}")
        .unDetailJudul.SetUnboundFieldSource ("{ado.DetailJudul}")
        
    End With
    
'    If sUkuranKertas = "" Then
'    sUkuranKertas = "5"
'    sOrientasKertas = "2"
'    sDuplex = "0"
'    End If

    'settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    
    CRViewer1.ReportSource = Report
    
    With CRViewer1
        .EnableGroupTree = False
        .ViewReport
        .Zoom 1
    End With
    Screen.MousePointer = vbDefault
    Set adocomd = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakPemanfaatanRS = Nothing
End Sub
