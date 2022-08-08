VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form md 
   Caption         =   "Laporan Cetak Stok Barang"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frm_cetak_stokBarang.frx":0000
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
Attribute VB_Name = "md"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New cr_stokBarang

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    adocomd.ActiveConnection = dbConn
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, adocomd

    With Report
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        .namaRuangan.SetText mstrNamaRuangan
        .usNamaBarang.SetUnboundFieldSource ("{ado.Nama Barang}")
        .usAsalBarang.SetUnboundFieldSource ("{ado.Asal Barang}")
        .unJmlMin.SetUnboundFieldSource ("{ado.JmlMinimum}")
        .usStock.SetUnboundFieldSource ("{ado.jmlStok}")
        .usJenisBarang.SetUnboundFieldSource ("{ado.Jenis Barang}")
        .usLokasi.SetUnboundFieldSource ("{ado.Lokasi}")
        .usSatuan.SetUnboundFieldSource ("{ado.Satuan}")
    End With

    With CRViewer1
        .EnableExportButton = True
        .EnableGroupTree = True
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
'    Set frm_cetak_stokBarang = Nothing
End Sub
