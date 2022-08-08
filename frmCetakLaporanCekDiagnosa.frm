VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakLaporanCekDiagnosa 
   Caption         =   "Medifirst2000 - Cetak"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakLaporanCekDiagnosa.frx":0000
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
Attribute VB_Name = "frmCetakLaporanCekDiagnosa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''splakuk 2009-06-10
Dim ReportBulan As New crlaporancekdiagnosa
'Dim ReportTriWulan As New crUreqRekapHarianPasienRI3Wulan


Private Sub Form_Load()
On Error GoTo errLoad

Dim tanggal As String
Dim tanggal2 As String
Dim laporan As String
Set dbcmd = New ADODB.Command

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
        Call msubRecFO(rs, strSQL)
                
            dbcmd.ActiveConnection = dbConn
            
           If frmLaporanCekDiagnosa.optDaftar.Value = True Then
                dbcmd.CommandText = "select * from V_RiwayatDiagnosaNull " _
                & " where " _
                & " KdInstalasi = '" & frmLaporanCekDiagnosa.dcInstalasi.BoundText & "' and TglPendaftaran between '" & Format(frmLaporanCekDiagnosa.dtpAwal, "yyyy-MM-dd HH:MM:ss") & "' and '" & Format(frmLaporanCekDiagnosa.dtpAkhir, "yyyy-MM-dd HH:MM:ss") & "'"
           End If

           If frmLaporanCekDiagnosa.optValidasi.Value = True Then
                 dbcmd.CommandText = "select * from V_RiwayatDiagnosaNull " _
                & " where " _
                & " KdInstalasi = '" & frmLaporanCekDiagnosa.dcInstalasi.BoundText & "' and TglPulang between '" & Format(frmLaporanCekDiagnosa.dtpAwal, "yyyy-MM-dd HH:MM:ss") & "' and '" & Format(frmLaporanCekDiagnosa.dtpAkhir, "yyyy-MM-dd HH:MM:ss") & "'"
           End If

'        dbcmd.CommandTimeout = 0
           
           tanggal = Format(frmLaporanCekDiagnosa.dtpAwal.Value, "dd MMMM yyyy")
           tanggal2 = Format(frmLaporanCekDiagnosa.dtpAkhir.Value, "dd MMMM yyyy")
        
        With ReportBulan
            .Database.AddADOCommand dbConn, dbcmd
        
            .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .txtAlamat.SetText "KABUPATEN " & strNKotaRS
            .txtAlamat2.SetText strNAlamatRS & " " & "Telp." & " " & strNTeleponRS
            
            .txtPeriode.SetText tanggal2
            .txtTanggal.SetText tanggal
            .txtRuangRawat.SetText strNNamaRuangan
            
            .udTglPendaftaran.SetUnboundFieldSource ("{ado.TglPendaftaran}")
            .unnoPendaftaran.SetUnboundFieldSource ("{ado.NoPendaftaran}")
            .unnoCM.SetUnboundFieldSource ("{ado.NoCM}")
            .usNmPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
            .udTglPulang.SetUnboundFieldSource ("{ado.TglPulang}")
            .usSMF.SetUnboundFieldSource ("{ado.SMF}")
            .usDokter.SetUnboundFieldSource ("{ado.Dokter}")
            .usKelas.SetUnboundFieldSource ("{ado.Kelas}")
            .usNmInstalasi.SetUnboundFieldSource ("{ado.NamaInstalasi}")
            
        End With
        CRViewer1.ReportSource = ReportBulan
   
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
    Set frmCetakLaporanCekDiagnosa = Nothing
End Sub
