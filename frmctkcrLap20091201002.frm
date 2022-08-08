VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmctkcrLap20091201002 
   Caption         =   "Medifirst2000 - Kunjungan Pasien Berdasarkan Diagnosa"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
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
Attribute VB_Name = "frmctkcrLap20091201002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReportBulan2 As New crLap20091201002

Private Sub Form_Load()
On Error GoTo errLoad

Dim tanggal As String
Dim laporan As String
Dim adocomd As New ADODB.Command

    Set frmctkcrLap20091201002 = Nothing '#
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
      
'            Call openConnection
            Set adocomd = New ADODB.Command '#
            Set adocomd = Nothing '#
            adocomd.ActiveConnection = dbConn
            adocomd.CommandText = "select Periode = cast(Tahun as char(4))+'/'+ cast(Bulan as char(2))+'/'+cast(Tanggal as char(2)), JmlAskes, JmlNonAskes, JmlKontrak, JmlBebasGratis, JmlTidakMampuGratis, Totalnya= isnull(JmlAskes,0)+isnull(JmlNonAskes,0)+isnull(JmlKontrak,0)+isnull(JmlBebasGratis,0)+isnull(JmlTidakMampuGratis,0) from v_S_TingkatPemanfaatanRSUD " _
            & " WHERE Tanggal Between '" & frmLap20091201002.dtpAwal.Day & " ' and '" & frmLap20091201002.dtpAkhir.Day & "' " & _
            " and Bulan between '" & frmLap20091201002.dtpAwal.Month & "' and '" & frmLap20091201002.dtpAkhir.Month & "'" & _
            " AND Tahun between '" & frmLap20091201002.dtpAwal.Year & "' and '" & frmLap20091201002.dtpAkhir.Year & "'" & _
            " ORDER BY Tahun, Bulan, Tanggal"
            
            
           adocomd.CommandType = adCmdText
'           adocomd.CommandTimeout = 120
           
           tanggal = Format(frmLap20091201002.dtpAwal.Value, "MMMM yyyy")
        
        With ReportBulan2
            .Database.AddADOCommand dbConn, adocomd
        
            .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .txtAlamat.SetText "KABUPATEN " & strNKotaRS
            .txtAlamat2.SetText strNAlamatRS & " " & "Telp." & " " & strNTeleponRS
            
            '.txtPeriode.SetText tanggal
'            .txtRuangRawat.SetText strNNamaRuangan
            
            .txtTanggalPilih1.SetText Format(frmLap20091201002.dtpAwal.Value, "dd/MM/yyyy")
            .txtTanggalPilih2.SetText Format(frmLap20091201002.dtpAkhir.Value, "dd/MM/yyyy")

            .usPeriode.SetUnboundFieldSource ("{ado.Periode}")
            .unAskes.SetUnboundFieldSource ("{ado.JmlAskes}")
            .unNonAskes.SetUnboundFieldSource ("{ado.JmlNonAskes}")
            .unKontrak.SetUnboundFieldSource ("{ado.JmlKontrak}")
            .unBebasGratis.SetUnboundFieldSource ("{ado.JmlBebasGratis}")
            .unTidakMampuGratis.SetUnboundFieldSource ("{ado.JmlTidakMampuGratis}")
            .unTotalnya.SetUnboundFieldSource ("{ado.Totalnya}")
            
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
    Set frmctkcrLap20091201002 = Nothing
End Sub
