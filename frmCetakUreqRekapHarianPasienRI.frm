VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakUreqRekapHarianPasienRI 
   Caption         =   "Medifirst2000 - Cetak"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakUreqRekapHarianPasienRI.frx":0000
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
Attribute VB_Name = "frmCetakUreqRekapHarianPasienRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''splakuk 2009-06-10

Dim ReportBulan As New crUreqDataKegiatanRS

Private Sub Form_Load()
On Error GoTo errLoad

Dim tanggal As String
Dim laporan As String

Set dbcmd = New ADODB.Command


    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
        strSQL = "Select sum(jlhBed) as JmlBed from NoKamar where KdRuangan='" & frmRekapitulasiHarianPasienRI.dcRuangan.BoundText & "' "
        Call msubRecFO(rs, strSQL)
        
        'Call openConnection
        
            dbcmd.ActiveConnection = dbConn
            dbcmd.CommandText = "select * from VRekapHarianPelPasienRI " _
            & " WHERE BulanS = " & frmRekapitulasiHarianPasienRI.dtpAwal.Month & " " _
            & " AND TahunS = " & frmRekapitulasiHarianPasienRI.dtpAwal.Year & " " _
            & " AND KdRuangan = '" & frmRekapitulasiHarianPasienRI.dcRuangan.BoundText & "' AND KdJenisPelayanan like '" & frmRekapitulasiHarianPasienRI.dcSubInstalasi.BoundText & "' ORDER BY TanggalS"
        
        
           'dbcmd.CommandType = adCmdText
'           dbcmd.CommandTimeout = 0
           
           tanggal = Format(frmRekapitulasiHarianPasienRI.dtpAwal.Value, "MMMM yyyy")
        
        With ReportBulan
            .Database.AddADOCommand dbConn, dbcmd
        
            .txtBed.SetText rs.Fields("JmlBed").Value
            .txtNamaRS.SetText strNNamaRS & " " & strKelasRS & " " & strKetKelasRS
            .txtAlamat.SetText "KABUPATEN " & strNKotaRS
            .txtAlamat2.SetText strNAlamatRS & " " & "Telp." & " " & strNTeleponRS
            
            .txtPeriode.SetText tanggal
            .txtRuangRawat.SetText strNNamaRuangan
            .txtSubInstalasi.SetText strNNamaSubInstalasi
            
            .UnboundNumber1.SetUnboundFieldSource ("{ado.TanggalS}")
            .UnboundNumber2.SetUnboundFieldSource ("{ado.Saldo}")
            .UnboundNumber3.SetUnboundFieldSource ("{ado.Masuk}")
            .UnboundNumber4.SetUnboundFieldSource ("{ado.Pindah}")
            
            .UnboundNumber6.SetUnboundFieldSource ("{ado.Dipindahkan}")
            .UnboundNumber7.SetUnboundFieldSource ("{ado.KeluarH}")
            
            .UnboundNumber9.SetUnboundFieldSource ("{ado.KeluarM1}")
            .UnboundNumber10.SetUnboundFieldSource ("{ado.KeluarM2}")
            
            .UnboundNumber12.SetUnboundFieldSource ("{ado.LamaRawat}")
            .UnboundNumber13.SetUnboundFieldSource ("{ado.KMSama}")
            .UnboundNumber14.SetUnboundFieldSource ("{ado.Sisa}")
            
            .UnboundNumber15.SetUnboundFieldSource ("{ado.VVIP}")
            .UnboundNumber16.SetUnboundFieldSource ("{ado.VIP}")
            .UnboundNumber17.SetUnboundFieldSource ("{ado.DeluxeSuper}")
            .UnboundNumber18.SetUnboundFieldSource ("{ado.Deluxe}")
            .UnboundNumber19.SetUnboundFieldSource ("{ado.Suite}")
            .UnboundNumber20.SetUnboundFieldSource ("{ado.SY}")
            .UnboundNumber21.SetUnboundFieldSource ("{ado.Standar}")
            .UnboundNumber22.SetUnboundFieldSource ("{ado.Intensif}")
            
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
    Set frmCetakUreqRekapHarianPasienRI = Nothing
End Sub
