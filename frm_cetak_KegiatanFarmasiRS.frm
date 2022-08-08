VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_cetak_KegiatanFarmasiRS 
   Caption         =   "Laporan Cetak Kegiatan Farmasi Rumah Sakit"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   Icon            =   "frm_cetak_KegiatanFarmasiRS.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   5850
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
Attribute VB_Name = "frm_cetak_KegiatanFarmasiRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crKegiatanFarmasiRumahSakit

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Set frm_cetak_KegiatanFarmasiRS = Nothing

    Dim adocomd As New ADODB.Command
    Call openConnection

    adocomd.ActiveConnection = dbConn

    adocomd.CommandText = strSQL

    adocomd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, adocomd

    With Report
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS

        .txtTanggal.SetText Format(frmKegiatanFarmasi.DTPickerAwal.value, "dd/MM/yyyy")
        .txtTanggal2.SetText Format(frmKegiatanFarmasi.DTPickerAkhir.value, "dd/MM/yyyy")

        .txtJml1.SetText iJml1
        .txtJml2.SetText iJml2
        .txtJml3.SetText iJml3
        .txtJml4.SetText cJml4
        .txtJml5.SetText cJml5
        .txtJml6.SetText cJml6
        .txtJml7.SetText cJml7
        .txtJml8.SetText cJml8
        .txtJml9.SetText cJml9
    End With

    CRViewer1.ReportSource = Report

    Screen.MousePointer = vbHourglass

    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 1
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
    Set frm_cetak_KegiatanFarmasiRS = Nothing
End Sub
