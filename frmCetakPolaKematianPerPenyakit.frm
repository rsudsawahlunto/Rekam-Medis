VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakPolaKematianPerPenyakit 
   Caption         =   "Medifirst2000 - Cetak"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakPolaKematianPerPenyakit.frx":0000
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
Attribute VB_Name = "frmCetakPolaKematianPerPenyakit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rptCetakDaftarPasienMeninggal As crPolaKematianPerPenyakit

Private Sub Form_Load()
    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn
    
    Me.Caption = "Medifirst2000 - Cetak Laporan Pola Kematian Penyakit"
    Set rptCetakDaftarPasienMeninggal = New crPolaKematianPerPenyakit

    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText
    With rptCetakDaftarPasienMeninggal
        .Database.AddADOCommand dbConn, dbcmd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS
        .txtAlamat2.SetText strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        
        .txtBulan.SetText MonthName(frmDaftarPolaKematianPenyakit.dtpAwal.Month) & "  " & frmDaftarPolaKematianPenyakit.dtpAwal.Year
'        .udTanggalMeninggal.SetUnboundFieldSource ("{ado.TglMeninggal}")
        .usNamaPenyakit.SetUnboundFieldSource ("{ado.NamaDiagnosa}")
        .unJml1Thn.SetUnboundFieldSource ("{ado.< 1 Tahun}")
        .unJml4Thn.SetUnboundFieldSource ("{ado.1 - 4 Tahun}")
        .unJmlSemuaUmur.SetUnboundFieldSource ("{ado.SemuaUmur}")
'        .unUmur.SetUnboundFieldSource ("{ado.UmurTahun}")
    
    End With
    CRViewer1.ReportSource = rptCetakDaftarPasienMeninggal
    
    With CRViewer1
        .EnableGroupTree = False
        .EnableExportButton = True
        .ViewReport
        .Zoom 1
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    With CRViewer1
        .Top = 0
        .Left = 0
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakPolaKematianPerPenyakit = Nothing
End Sub
