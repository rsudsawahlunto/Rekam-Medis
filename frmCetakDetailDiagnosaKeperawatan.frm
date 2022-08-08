VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakDetailDiagnosaKeperawatan 
   Caption         =   "Medifirst2000 - Master Diagnosa Keperawatan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakDetailDiagnosaKeperawatan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
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
Attribute VB_Name = "frmCetakDetailDiagnosaKeperawatan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crDetailDiagnosaKeperawatan

Private Sub Form_Load()
    Call openConnection

    Set frmCetakDetailDiagnosaKeperawatan = Nothing
    Set dbcmd = New ADODB.Command

    With dbcmd
        .ActiveConnection = dbConn
        .CommandText = "SELECT * FROM DetailDiagnosaKeperawatan"
        .CommandType = adCmdText
    End With

    strSQL = "SELECT * FROM DetailDiagnosaKeperawatan"
    Call msubRecFO(rs, strSQL)

    With Report
        .Database.AddADOCommand dbConn, dbcmd

        .usKdDetailAskep.SetUnboundFieldSource ("{ado.KdDetailAskep}")
        .usDetailAskep.SetUnboundFieldSource ("{ado.DetailAskep}")

        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtWebsite.SetText strWebsite & ", " & strEmail
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .EnableGroupTree = False
        .Zoom 1
    End With

    Screen.MousePointer = vbDefault

    Set dbcmd = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakDetailDiagnosaKeperawatan = Nothing
End Sub
