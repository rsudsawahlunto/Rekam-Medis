VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakLapSaldo 
   Caption         =   "Medifirst2000"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   Icon            =   "frmCetakLapSaldo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   6855
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
Attribute VB_Name = "frmCetakLapSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crLapSaldo

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    Dim adocomd As New ADODB.Command

    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn

    Me.Caption = "Medifirst2000 - Cetak Laporan Saldo Barang"
    Set Report = New crLapSaldo

    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText
    With Report
        .Database.AddADOCommand dbConn, dbcmd

        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail

        .txtPeriode.SetText Format(mdTglAwal, "MMM yyyy")

        .usNamaBarang.SetUnboundFieldSource ("{ado.NamaBarang}")
        .usAsalBarang.SetUnboundFieldSource ("{ado.AsalBarang}")
        .unSaldoAwal.SetUnboundFieldSource ("{ado.SaldoAwal}")
        .unJmlTerima.SetUnboundFieldSource ("{ado.JmlTerima}")
        .unJmlKeluar.SetUnboundFieldSource ("{ado.JmlKeluar}")
        .unSaldoAkhir.SetUnboundFieldSource ("{ado.SaldoAkhir}")
        .ucHarga.SetUnboundFieldSource ("{ado.HargaNetto}")

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
    Call frmLaporanSaldoBarangNM_v3.DeleteTable
    Set frmCetakLapSaldo = Nothing
End Sub
