VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakDaftarPasienMeninggal 
   Caption         =   "Medifirst2000 - Cetak"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakDaftarPasienMeninggal.frx":0000
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
Attribute VB_Name = "frmCetakDaftarPasienMeninggal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rptCetakDaftarPasienMeninggal As crCetakDaftarPasienMeninggal

Private Sub Form_Load()
    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn
    
    Me.Caption = "Medifirst2000 - Cetak Laporan Sensus Pelayanan"
    Set rptCetakDaftarPasienMeninggal = New crCetakDaftarPasienMeninggal
    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText
    With rptCetakDaftarPasienMeninggal
        .Database.AddADOCommand dbConn, dbcmd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS
        .txtAlamat2.SetText strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtPeriode.SetText Format(frmDaftarPasienMeninggal.dtpAwal.Value, "dd MMMM yyyy") & " s/d " & Format(frmDaftarPasienMeninggal.dtpAkhir.Value, "dd MMMM yyyy")
        
        .usNoPendaftaran.SetUnboundFieldSource ("{ado.NoPendaftaran}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.Nama Pasien}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .usUmur.SetUnboundFieldSource ("{ado.UmurTahun}")
        .usAlamat.SetUnboundFieldSource ("{ado.Alamat}")
        .udTglPendaftaran.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .udTglMeninggal.SetUnboundFieldSource ("{ado.TglMeninggal}")
        .usKasusPenyakit.SetUnboundFieldSource ("{ado.NamaSubInstalasi}")
        .usPenyebab.SetUnboundFieldSource ("{ado.Penyebab}")
        .usDiagnosa.SetUnboundFieldSource ("{ado.NamaDiagnosa}")
        .usTempatMeninggal.SetUnboundFieldSource ("{ado.Tempat Meninggal}")
        .usDokterPemeriksa.SetUnboundFieldSource ("{ado.Dokter Pemeriksa}")
         settingreport rptCetakDaftarPasienMeninggal, sPrinter, sDriver, sUkuranKertas, sDuplex, crLandscape
    End With
    CRViewer1.ReportSource = rptCetakDaftarPasienMeninggal
    
    With CRViewer1
        .EnableGroupTree = True
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
    Set frmCetakDaftarPasienMeninggal = Nothing
End Sub












