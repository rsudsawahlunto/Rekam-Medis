VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakDaftarPemeriksaanPasien 
   Caption         =   "Medifirst2000 - Daftar Pasien Rawat Jalan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakDaftarPemeriksaanPasien.frx":0000
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
Attribute VB_Name = "frmCetakDaftarPemeriksaanPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crDaftarPemeriksaanPasien

Private Sub Form_Load()
Dim adocomd As New ADODB.Command
    Call openConnection
    Set frmCetakDaftarPemeriksaan = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
'     " select * from V_LaporanDatapelayananDokter where Tgl between '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' and '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' " '& _
'     " ORDER BY TglKeluar DESC, [Cara Pulang] DESC "
    adocomd.CommandType = adCmdText
    
    With Report
        .txtNamaRS.SetText strNNamaRS
        .txtAlamatRS.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        '.txtWebsiteRS.SetText strWebsite & ", " & strEmail
        
'        If Format(mdTglAwal, "dd MMMM yyyy") = Format(mdTglAkhir, "dd MMMM yyyy") Then
'           .txtTanggal.SetText "Tanggal Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy")
'        Else
'           .txtTanggal.SetText "Periode Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") & " S/d " & Format(mdTglAkhir, "dd MMMM yyyy")
'        End If

        .txttanggal1.SetText Format(frmDaftarPemeriksaanPasien.DTPickerAwal, "dd MMMM yyyy")
        .txtTanggal2.SetText Format(frmDaftarPemeriksaanPasien.DTPickerAkhir, "dd MMMM yyyy")
          
        .Database.AddADOCommand dbConn, adocomd
        
        .udtanggal.SetUnboundFieldSource ("{ado.TglPelayanan}")
        .usnoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNmPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
        .usNmDokter.SetUnboundFieldSource ("{ado.DokterPemeriksa}")
        .usInstalasi.SetUnboundFieldSource ("{ado.NamaInstalasi}")
        .usUmur.SetUnboundFieldSource ("{ado.Umur}")
        .usDiagnosa.SetUnboundFieldSource ("{ado.NamaDiagnosa}")
        .usJenisPelayanan.SetUnboundFieldSource ("{ado.JenisPelayanan}")
        .usTindakan.SetUnboundFieldSource ("{ado.Tindakan}")
        
'        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .EnableGroupTree = False
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
    Set frmCetakDaftarPemeriksaanPasien = Nothing
End Sub

