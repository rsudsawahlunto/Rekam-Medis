VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmViewerRujukanLuar 
   Caption         =   "Medifirst2000 - Cetak No SJP"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3840
   Icon            =   "frmViewerRujukanLuar.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   3840
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   4125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3765
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
Attribute VB_Name = "frmViewerRujukanLuar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim Report As New crCetakSJP
Dim Report As New crCetakRujukanLuar

Private Sub Form_Load()
  
    Set FrmViewerLaporan = Nothing
    Set dbcmd = New ADODB.Command
    
    strSQL = "select * " & _
        " from V_RujukanKeluarRS  where " & _
        " Nosep='" & mstrNoSepRujukan & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        MsgBox "Data Tidak Ada"
        Exit Sub
    End If
    With dbcmd
        .ActiveConnection = dbConn
        .CommandText = strSQL
        .CommandType = adCmdText
    End With
    
    With Report
        .Database.AddADOCommand dbConn, dbcmd
      
       
        .txtNomorKartuAskes.SetText IIf(IsNull(rs("IdAsuransi")), "-", rs("IdAsuransi"))
        .txtdiagnosa.SetText IIf(IsNull(rs("DiagnosaRujukan")), "-", rs("DiagnosaRujukan"))
       
        .txtNamaPasien.SetText IIf(IsNull(rs("NamaLengkap")), "-", rs("NamaLengkap"))
        .txtCOB.SetText strNNamaRS
        .txtRuangan.SetText IIf(IsNull(rs("PoliTujuan")), "-", rs("PoliTujuan"))
        
        .txtkelamin.SetText IIf(IsNull(rs("JenisKelamin")), "-", rs("JenisKelamin"))
        
        If rs("TipeRujukan").value = 0 Then
          .txtJenisrawat.SetText "RUJUKAN PENUH"
        ElseIf rs("TipeRujukan").value = 1 Then
          .txtJenisrawat.SetText "RUJUKAN PARTIAL"
        Else
          .txtJenisrawat.SetText "RUJUKAN BALIK"
        End If
       
        If rs("jnsPelayanan").value = 1 Then
          .txtJenisrawat.SetText "RAWAT INAP"
        Else
          .txtJenisrawat.SetText "RAWAT JALAN"
        End If
        .txtNamaRsTujuan.SetText IIf(IsNull(rs("TujuanRujukan")), "-", rs("TujuanRujukan"))
        .txtPeserta.SetText IIf(IsNull(rs("NoRujukan")), "-", rs("NoRujukan"))
        .txtKeterangan.SetText IIf(IsNull(rs("CatatanRujukan")), "-", rs("CatatanRujukan"))
        
        
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
    Set dbcmd = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmViewerRujukanLuar = Nothing
End Sub
