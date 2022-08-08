VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakStokOpname 
   Caption         =   "Medifrst2000 - Stok Opname"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   Icon            =   "frmCetakStokOpname.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4635
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7005
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   5805
      DisplayGroupTree=   0   'False
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
      EnableAnimationControl=   0   'False
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
Attribute VB_Name = "frmCetakStokOpname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crStokOpname

Private Sub Form_Load()
    On Error GoTo errLoad
    Dim adocomd As New ADODB.Command

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Call openConnection

    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = "sELECT * from V_DataStokBarangMedisRekap WHERE KdRuangan = '" & mstrKdRuangan & "' AND TglClosing = '" & Format(mdtglclosing, "yyyy/MM/dd") & "'"
    adocomd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, adocomd

    With Report
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail

        .usNamaBarang.SetUnboundFieldSource ("{ado.NamaBarang}")
        .usJenisBarang.SetUnboundFieldSource ("{ado.JenisBarang}")
        .usAsalBarang.SetUnboundFieldSource ("{ado.AsalBarang}")
        .usKekuatan.SetUnboundFieldSource ("{ado.KeKuatan}")
        .unStokSystem.SetUnboundFieldSource ("{ado.StokSystem}")
        .unStokReal.SetUnboundFieldSource ("{ado.StokReal}")
        .ucHargaNetto1.SetUnboundFieldSource Format(("{ado.HargaNetto1}"), "##,###,##0")
        .ucHargaNetto2.SetUnboundFieldSource Format(("{ado.HargaNetto2}"), "##,###,##0")
        .ucDiscount.SetUnboundFieldSource Format(("{ado.Discount}"), "##,###,##0")
        .ucTotalNetto1.SetUnboundFieldSource Format(("{ado.TotalNetto1}"), "##,###,##0")
        .ucTotalNetto2.SetUnboundFieldSource Format(("{ado.TotalNetto2}"), "##,###,##0")

        .txtRuanganLogin.SetText mstrNamaRuangan
        .txtUser.SetText petugas
    End With
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 1
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
errLoad:
    Screen.MousePointer = vbDefault
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakStokOpname = Nothing
End Sub
