VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_cetak_PengunjungPasien 
   Caption         =   "Laporan Cetak Pengunjung Pasien"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frm_cetak_PengunjungPasien.frx":0000
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
Attribute VB_Name = "frm_cetak_PengunjungPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crPengunjungPasien

Private Sub Form_Load()

    Me.WindowState = 2
    Set frm_cetak_PengunjungPasien = Nothing

    Dim adocomd As New ADODB.Command

    Call openConnection

    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText
    
    Report.Database.AddADOCommand dbConn, adocomd
    Report.usjudul.SetUnboundFieldSource ("{ado.judul}")
    Report.unJumlah.SetUnboundFieldSource ("{ado.jumlah}")

    strSQL = "select * " & _
    " from V_Koders  "

    Call msubRecFO(rs, strSQL)

    With Report
        .Text2.SetText rs("NO1")
        .Text3.SetText rs("NO2")
        .Text4.SetText rs("NO3")
        .Text5.SetText rs("NO4")
        .Text6.SetText rs("NO5")
        .Text7.SetText rs("NO6")
        .Text8.SetText rs("NO7")
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom 1
        End With
    Else
        Report.PrintOut False
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
    Set frm_cetak_PengunjungPasien = Nothing
End Sub

