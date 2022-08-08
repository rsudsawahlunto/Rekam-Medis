VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmMorbiditasRJ2 
   Caption         =   "Morbiditas"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   Icon            =   "frmMorbiditasRawatJalan2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   8325
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
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
Attribute VB_Name = "frmMorbiditasRJ2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crMorbiditasRawatInap2
Dim adoCommand As New ADODB.Command
'Dim strSQL As String

Private Sub Form_Load()
On Error GoTo hell

    openConnection
    Set frmMorbiditasRJ2 = Nothing
            
    adoCommand.CommandText = strSQL
    adoCommand.CommandType = adCmdText
        
    If Format(mdTglMasuk, "dd MMMM yyyy") = Format(mdTglAkhir, "dd MMMM yyyy") Then
        tanggal = "Tanggal Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") '& " S/d " & Format(mdtglakhir, "dd MMMM yyyy")
    Else
        tanggal = "Periode Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") & " S/d " & Format(mdTglAkhir, "dd MMMM yyyy")
    End If

 With Report
        .Database.AddADOCommand dbConn, adoCommand
        
        If frmLapMorbiditas.dcInstalasi.BoundText = "01" Then
            .txtJudul.SetText "DATA KEADAAN MORBIDITAS PASIEN GAWAT DARURAT"
            
            .txtJudul2.SetText "FORMULIR RL 2b"
            .txtPeriode.SetText tanggal
            .usNoDTD.SetUnboundFieldSource ("{ado.NoDTD}")
            .usGroup.SetUnboundFieldSource ("{ado.Grup}")
            .usNoDT.SetUnboundFieldSource ("{ado.NoDTerperinci}")
            .usNamaDTD.SetUnboundFieldSource ("{ado.NamaDTD}")
            .unKel1.SetUnboundFieldSource ("{ado.Kel_Umur1}")
            .unKel2.SetUnboundFieldSource ("{ado.Kel_Umur2}")
            .unKel3.SetUnboundFieldSource ("{ado.Kel_Umur3}")
            .unKel4.SetUnboundFieldSource ("{ado.Kel_Umur4}")
            .unKel5.SetUnboundFieldSource ("{ado.Kel_Umur5}")
            .unKel6.SetUnboundFieldSource ("{ado.Kel_Umur6}")
            .unKel7.SetUnboundFieldSource ("{ado.Kel_Umur7}")
            .unKel8.SetUnboundFieldSource ("{ado.Kel_Umur8}")
            .unKelL.SetUnboundFieldSource ("{ado.Kel_L}")
            .unKelP.SetUnboundFieldSource ("{ado.Kel_P}")
            .unKelH.SetUnboundFieldSource ("{ado.Total}")
            .txtJmlPasien.SetText "Jumlah Kasus Baru Menurut Seks"
            .txtJmlPasienH.SetText "Jumlah Kunj Pasien Total"
            .txtJmlPasienM.Suppress = True
            .unKelM.Suppress = True
            .Text36.Suppress = True
            .Line32.Suppress = True
            .Line33.Suppress = True
            .Line35.Right = 10830
            .Line37.Right = 10830
            .Field4.Suppress = True
            '.unKelM.SetUnboundFieldSource ("{ado.Kel_M}")
            .Text1.SetText strNNamaRS
            .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
            .Text3.SetText strWebsite & ", " & strEmail
'            .SelectPrinter sDriver, sPrinter, vbNull
'            settingreport Report, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
            
    ElseIf frmLapMorbiditas.dcInstalasi.BoundText = "02" Then
            .txtJudul.SetText "DATA KEADAAN MORBIDITAS PASIEN RAWAT JALAN"
            .txtJudul2.SetText "FORMULIR RL 2b1"
            .txtPeriode.SetText tanggal
            .usNoDTD.SetUnboundFieldSource ("{ado.NoDTD}")
            .usGroup.SetUnboundFieldSource ("{ado.Grup}")
            .usNoDT.SetUnboundFieldSource ("{ado.NoDTerperinci}")
            .usNamaDTD.SetUnboundFieldSource ("{ado.NamaDTD}")
            .unKel1.SetUnboundFieldSource ("{ado.Kel_Umur1}")
            .unKel2.SetUnboundFieldSource ("{ado.Kel_Umur2}")
            .unKel3.SetUnboundFieldSource ("{ado.Kel_Umur3}")
            .unKel4.SetUnboundFieldSource ("{ado.Kel_Umur4}")
            .unKel5.SetUnboundFieldSource ("{ado.Kel_Umur5}")
            .unKel6.SetUnboundFieldSource ("{ado.Kel_Umur6}")
            .unKel7.SetUnboundFieldSource ("{ado.Kel_Umur7}")
            .unKel8.SetUnboundFieldSource ("{ado.Kel_Umur8}")
            .unKelL.SetUnboundFieldSource ("{ado.Kel_L}")
            .unKelP.SetUnboundFieldSource ("{ado.Kel_P}")
            
            .unKelH.SetUnboundFieldSource ("{ado.Kel_Kunj}")
            .txtJmlPasien.SetText "Jumlah Kasus Baru Menurut Seks"
            .txtJmlPasienH.SetText "Jumlah Kunj Pasien Total"
            .txtJmlPasienM.Suppress = True
            .unKelM.Suppress = True
            .Text36.Suppress = True
            .Line32.Suppress = True
            .Line33.Suppress = True
            .Line35.Right = 10830
            .Line37.Right = 10830
            .Field4.Suppress = True
    '        .unKelM.SetUnboundFieldSource ("{ado.Kel_M}")
            .Text1.SetText strNNamaRS
            .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
            .Text3.SetText strWebsite & ", " & strEmail
'            .SelectPrinter sDriver, sPrinter, vbNull
'            settingreport Report, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
     Else
            .txtJudul.SetText "DATA KEADAAN MORBIDITAS PASIEN RAWAT INAP"
            .txtJudul2.SetText "FORMULIR RI 2a"
            .txtPeriode.SetText tanggal
            .usNoDTD.SetUnboundFieldSource ("{ado.NoDTD}")
            .usGroup.SetUnboundFieldSource ("{ado.Grup}")
            .usNoDT.SetUnboundFieldSource ("{ado.NoDTerperinci}")
            .usNamaDTD.SetUnboundFieldSource ("{ado.NamaDTD}")
            .unKel1.SetUnboundFieldSource ("{ado.Kel_Umur1}")
            .unKel2.SetUnboundFieldSource ("{ado.Kel_Umur2}")
            .unKel3.SetUnboundFieldSource ("{ado.Kel_Umur3}")
            .unKel4.SetUnboundFieldSource ("{ado.Kel_Umur4}")
            .unKel5.SetUnboundFieldSource ("{ado.Kel_Umur5}")
            .unKel6.SetUnboundFieldSource ("{ado.Kel_Umur6}")
            .unKel7.SetUnboundFieldSource ("{ado.Kel_Umur7}")
            .unKel8.SetUnboundFieldSource ("{ado.Kel_Umur8}")
            .unKelL.SetUnboundFieldSource ("{ado.Kel_L}")
            .unKelP.SetUnboundFieldSource ("{ado.Kel_P}")
            .unKelH.SetUnboundFieldSource ("{ado.Kel_H}")
            .unKelM.SetUnboundFieldSource ("{ado.Kel_M}")
            .txtJmlPasien.SetText "Pasien Keluar(Hidup&Mati)"
            .txtJmlPasienH.SetText "Jumlah Pasien Keluar (13+14)"
            .txtJmlPasienM.SetText "Jumlah Pasien Keluar Mati"
            .txtJmlPasienM.Suppress = False
            .unKelM.Suppress = False
            .Line32.Suppress = False
            .Line33.Suppress = False
            .Line35.Right = 11430
            .Line37.Right = 11430
            .Field4.Suppress = False
            
            '.unKelM.SetUnboundFieldSource ("{ado.Kel_M}")
            .Text1.SetText strNNamaRS
            .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
            .Text3.SetText strWebsite & ", " & strEmail
'            .SelectPrinter sDriver, sPrinter, vbNull
'            settingreport Report, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
           
        End If
        End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .EnableGroupTree = True
        .ViewReport
        .Zoom 1
    End With
    Screen.MousePointer = vbDefault
Exit Sub
hell:
    Call msubPesanError

End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMorbiditasRJ2 = Nothing
End Sub
