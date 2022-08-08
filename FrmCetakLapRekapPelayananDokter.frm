VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmCetakLapRekapPelayananDokter 
   Caption         =   "FrmCetakLapRekapPelayananDokter"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   Icon            =   "FrmCetakLapRekapPelayananDokter.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   10365
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   10215
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
Attribute VB_Name = "FrmCetakLapRekapPelayananDokter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rptRekapPelayananDokter As New crRekapPelayananDokter

Private Sub Form_Load()
Dim adocmd As New ADODB.Command

On Error GoTo errLoad
 
Screen.MousePointer = vbHourglass
Me.WindowState = 2

    Set adocmd.ActiveConnection = dbConn
    strSQL = "SELECT Distinct Tgl,NoCM, NamaPasien, NamaInstalasi, NamaRuangan FROM V_RekapitulasiPelayananDokterPerPasien " & _
             "where KdInstalasi like '%" & FrmRekapPelayananDokter.dcInstalasi.BoundText & "%' " & _
             "and KdRuangan like '%" & FrmRekapPelayananDokter.dcRuangan.BoundText & "%' " & _
             "and IdDokter like '%" & FrmRekapPelayananDokter.dcDokter.BoundText & "%' AND TglPelayanan BETWEEN '" & Format(FrmRekapPelayananDokter.dtpAwal.Value, "yyyy/MM/dd hh:mm:00") & "' AND '" & Format(FrmRekapPelayananDokter.dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "'"
             
With rptRekapPelayananDokter
    .txtNamaRS.SetText strNNamaRS
    .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    .txtAlamat2.SetText strWebsite & ", " & strEmail
    

    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
    .Database.AddADOCommand dbConn, adocmd
    
    If Format(FrmRekapPelayananDokter.dtpAkhir, "yyyy MM dd HH:mm:ss") = Format(FrmRekapPelayananDokter.dtpAwal, "yyyy MM dd HH:mm:ss") Then
            .txtPeriode.SetText Format(FrmRekapPelayananDokter.dtpAwal, "dd MMMM YYYY")
        Else
            .txtPeriode.SetText Format(FrmRekapPelayananDokter.dtpAwal, "dd MMMM YYYY HH:mm:ss") & " S/D " & Format(FrmRekapPelayananDokter.dtpAkhir, "dd MMMM YYYY HH:mm:ss")
    End If
    
    .UnTgl.SetUnboundFieldSource ("{ado.Tgl}")
    '.udtTanggal.SetUnboundFieldSource ("{ado.Tgl}")
    .UsNoCM.SetUnboundFieldSource ("{ado.NoCM}")
    .UsNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
    .usInstalasi.SetUnboundFieldSource ("{ado.NamaInstalasi}")
    .UsRuangan.SetUnboundFieldSource ("{ado.NamaRuangan}")
    .txtNamaDokter.SetText strNamaDokter
    '.UsDokter.SetUnboundFieldSource ("{ado.NamaDokter}")
End With

CRViewer1.ReportSource = rptRekapPelayananDokter
CRViewer1.Zoom 1
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
Exit Sub

errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmCetakLapRekapPelayananDokter = Nothing
End Sub
