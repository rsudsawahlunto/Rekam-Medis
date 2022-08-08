VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakKegiatanPembedahan 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakKegiatanPembedahan.frx":0000
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
Attribute VB_Name = "frmCetakKegiatanPembedahan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crKegiatanPembedahan
Dim adocomd As New ADODB.Command

Private Sub Form_Load()
    Set adocomd = New ADODB.Command
    Set adocomd = Nothing
    Dim tanggal As String

    adocomd.ActiveConnection = dbConn

    Me.WindowState = 2

    Set Report = New crKegiatanPembedahan

    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText

    With Report
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS
        .Text1.SetText " LAPORAN KEGIATAN PEMBEDAHAN "

        .Database.AddADOCommand dbConn, adocomd

        .txtTanggal.SetText Format(frmKegiatanPembedahan.DTPickerAwal.value, "dd/MM/yyyy")
        .txtPeriode.SetText Format(frmKegiatanPembedahan.DTPickerAkhir.value, "dd/MM/yyyy")

        .usSpesialisasi.SetUnboundFieldSource "{ado.Spesialisasi}"
        .unJml.SetUnboundFieldSource "{ado.Total}"
        .unKBedah.SetUnboundFieldSource "{ado.Khusus_KBedah}"
        .unUDarurat.SetUnboundFieldSource "{ado.Khusus_IGD}"
        .unKBedah1.SetUnboundFieldSource "{ado.Besar_KBedah}"
        .unUDarurat1.SetUnboundFieldSource "{ado.Besar_IGD}"
        .unKBedah2.SetUnboundFieldSource "{ado.Sedang_KBedah}"
        .unUDarurat2.SetUnboundFieldSource "{ado.Sedang_IGD}"
        .unKBedah4.SetUnboundFieldSource "{ado.Kecil_KBedah}"
        .unUDarurat4.SetUnboundFieldSource "{ado.Kecil_IGD}"
        .unPoliKlinik.SetUnboundFieldSource "{ado.Kecil_RJ}"
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakKegiatanPembedahan = Nothing
End Sub
