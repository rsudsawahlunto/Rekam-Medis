VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakUreqRekapHarianPasienIGD 
   Caption         =   "Medifirst2000 - Cetak"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakUreqRekapHarianPasienIGD.frx":0000
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
Attribute VB_Name = "frmCetakUreqRekapHarianPasienIGD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReportBulan2 As New crUreqRekapHarianPasienIGD2

Private Sub Form_Load()
On Error GoTo errLoad

Dim tanggal As String
Dim laporan As String
Dim adocomd As New ADODB.Command


    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
      
            Call openConnection
        
            adocomd.ActiveConnection = dbConn
            adocomd.CommandText = "select distinct * from VPelIGD " _
            & " WHERE TanggalS Between '" & frmRekapHarianPasienIGD.dtpAwal.Day & " ' and '" & frmRekapHarianPasienIGD.dtpAkhir.Day & "' " & _
            " and BulanS between '" & frmRekapHarianPasienIGD.dtpAwal.Month & "' and '" & frmRekapHarianPasienIGD.dtpAkhir.Month & "'" & _
            " AND TahunS between '" & frmRekapHarianPasienIGD.dtpAwal.Year & "' and '" & frmRekapHarianPasienIGD.dtpAkhir.Year & "'" & _
            " AND KdRuangan = '" & frmRekapHarianPasienIGD.dcRuangan.BoundText & "' ORDER BY TanggalS"
            
            
           adocomd.CommandType = adCmdText
'           adocomd.CommandTimeout = 120
           
           tanggal = Format(frmRekapHarianPasienIGD.dtpAwal.Value, "MMMM yyyy")
        
        With ReportBulan2
            .Database.AddADOCommand dbConn, adocomd
        
            .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .txtAlamat.SetText "KABUPATEN " & strNKotaRS
            .txtAlamat2.SetText strNAlamatRS & " " & "Telp." & " " & strNTeleponRS
            
            '.txtPeriode.SetText tanggal
            .txtRuangRawat.SetText strNNamaRuangan
            
            .txtTanggalPilih1.SetText Format(frmRekapHarianPasienIGD.dtpAwal.Value, "dd/MM/yyyy")
            .txtTanggalPilih2.SetText Format(frmRekapHarianPasienIGD.dtpAkhir.Value, "dd/MM/yyyy")

            .usJenisPelayanan.SetUnboundFieldSource ("{ado.NamaSubInstalasi}")
            .unAsalRujukan.SetUnboundFieldSource ("{ado.JmlAsalRujukan}")
            .unNonrujukan.SetUnboundFieldSource ("{ado.JmlAsalNonRujukan}")
            .undirawat.SetUnboundFieldSource ("{ado.JmlDirawat}")
            .unDirujuk.SetUnboundFieldSource ("{ado.JmlDirujuk}")
            .unPulang.SetUnboundFieldSource ("{ado.JmlPulang}")
            .unMati.SetUnboundFieldSource ("{ado.JmlMati}")
            .udtTanggalS.SetUnboundFieldSource ("{ado.TanggalS}")
            
        End With
        CRViewer1.ReportSource = ReportBulan2
    With CRViewer1
        .Zoom 1
        .ViewReport

    End With
    Screen.MousePointer = vbDefault
    
Exit Sub
errLoad:
    Screen.MousePointer = vbDefault
    msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakUreqRekapHarianPasienIGD = Nothing
End Sub
