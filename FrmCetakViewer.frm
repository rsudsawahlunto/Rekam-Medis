VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmcetakviewer 
   Caption         =   "Medifirst2000 - Cetak"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmCetakViewer.frx":0000
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
Attribute VB_Name = "frmcetakviewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crRkpPsnJenisPeriksa
Dim rpttahun As New CrTndakandanPeriksaPTahun
Private Sub Form_Load()
Set Report = Nothing
Set rpttahun = Nothing
Set frmcetakviewer = Nothing
Select Case StrCetak2
    Case "LapKunjunganJenisPeriksahari"
        Select Case StrCetak3
        Case "JenisPeriksaBInstalasiAsal"
        Call LapKunjunganJenisPeriksaHariperAsalInstalasi
        Case "JenisPeriksaBJenispasien"
        Call lapkunjunganperhariperjenispasien
    End Select
    Case "LapKunjunganJenisPeriksabulan"
        Select Case StrCetak3
        Case "LapKunjunganJenisPeriksaBulanInstalasiAsal"
        Call lakunjunganbulaninstalasiasal
        Case "LapKunjunganJenisPeriksaBulanJenisPasien"
        Call lapkunjunganperBulanperjenispasien
    End Select
    Case "LapKunjunganJenisPeriksaTahun"
        Select Case StrCetak3
        Case "LapKunjunganJenisPeriksaJenisPasienTahun"
        Call LapKunjunganJenisPeriksaJenisPasienTahun
        Case "LapKunjunganJenisPeriksaInstalasiAsalTahun"
        Call LapKunjunganJenisPeriksaInstalasiAsalTahun
    End Select
'========================================================
    Case "LapKunjunganJenisTindakanHari"
        Select Case StrCetak3
        Case "LapKunjunganJenisTindakanBinstalasiAsal"
        Call LapKunjunganJenisTindakanBinstalasiAsal
        Case "LapKunjunganJenisTindakanBJenisPasienHari"
        Call LapKunjunganJenisTindakanBJenisPasienHari
        End Select
    Case "LapKunjunganJenisTindakanBulan"
        Select Case StrCetak3
        Case "LapKunjunganJenisTindakanBinstalasiAsalBulan"
        Call LapKunjunganJenisTindakanBinstalasiAsalBulan
        Case "LapKunjunganJenisTindakanBJenisPasienBulan"
        Call LapKunjunganJenisTindakanBJenisPasienBulan
        End Select
    Case "LapKunjunganJenisTindakantahun"
        Select Case StrCetak3
        Case "LapKunjunganJenisTindakanBinstalasiAsaltahun"
        Call LapKunjunganJenisTindakanBinstalasiAsaltahun
        Case "LapKunjunganJenisTindakanBJenisPtahun"
        Call LapKunjunganJenisTindakanBJenisPtahun
        End Select
  
    
End Select
End Sub
Private Sub lapkunjunganperBulanperjenispasien()
Call openConnection
    Dim adocomd As New ADODB.Command
     Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
   With Report
     .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "MMMM-yyyy")) = CStr(Format(mdTglAkhir, "MMMM-yyyy")) Then
          .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
          .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If
    .txtNamaRS.SetText strNNamaRS
    .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    .txtAlamat2.SetText strEmail
    .txtinstalasi.SetText MstrInstalasi2
    .txtjudul.SetText "LAPORAN KUNJUNGAN PASIEN BERDASARKAN JENIS PERIKSA"
    .UdPeriode.SetUnboundFieldSource ("{ado.TglPelayanan}")
    .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
    .UsJnsPelayanan.SetUnboundFieldSource ("{ado.JenisPeriksa}")
    .UsInstasal.SetUnboundFieldSource ("{ado.JenisPasien}")
    .UsJK.SetUnboundFieldSource ("{ado.JK}")
    .txtjenis.SetText ("Jenis Pasien")
    .JMlPasien.SetUnboundFieldSource ("{ado.JmlPelayanan}")
    If sUkuranKertas = "" Then
    sUkuranKertas = "3"
    sOrientasKertas = "2"
    End If
    settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    Call ukurkertas
   End With
    
Screen.MousePointer = vbHourglass
With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (75)
End With
Screen.MousePointer = vbDefault
End Sub
Private Sub LapKunjunganJenisPeriksaJenisPasienTahun()
Call openConnection
    Dim adocomd As New ADODB.Command
     Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With rpttahun
    .Database.AddADOCommand dbConn, adocomd
    .txtNamaRS.SetText strNNamaRS
    .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    .txtAlamat2.SetText strEmail
    .txtinstalasi.SetText MstrInstalasi2
    .txtjudul.SetText ("LAPORAN REKAPITULASI KUNJUNGAN PASIEN BERDASARKAN JENIS TINDAKAN ")
    
        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
             .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
             .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If

    .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
    .Udate.SetUnboundFieldSource ("{ado.TglPelayanan}")
    .usJenisPelayanan.SetUnboundFieldSource ("{ado.JenisPeriksa}")
    .usinstalasiasal.SetUnboundFieldSource ("{ado.JenisPasien}")
    .UsJK.SetUnboundFieldSource ("{ado.JK}")
    .txtpilihan.SetText ("Jenis Pasien")
    .UJml.SetUnboundFieldSource ("{ado.JmlPelayanan}")
'    .txtjudul.SetText Judul7
     settingreport rpttahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
   End With
    
Screen.MousePointer = vbHourglass
With CRViewer1
        .ReportSource = rpttahun
        .ViewReport
        .Zoom (75)
End With
Screen.MousePointer = vbDefault
Set frmcetakviewer = Nothing
End Sub
Private Sub LapKunjunganJenisTindakanBinstalasiAsaltahun()
Call openConnection
    Dim adocomd As New ADODB.Command
     Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With rpttahun
    .Database.AddADOCommand dbConn, adocomd
    .txtNamaRS.SetText strNNamaRS
    .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    .txtAlamat2.SetText strEmail
    .txtinstalasi.SetText MstrInstalasi2
    
    .txtjudul.SetText ("LAPORAN REKAPITULASI KUNJUNGAN PASIEN BERDASARKAN JENIS TINDAKAN ")
            If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
             .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
            Else
             .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
            End If
   
    .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
    .Udate.SetUnboundFieldSource ("{ado.TglPelayanan}")
    .usJenisPelayanan.SetUnboundFieldSource ("{ado.JenisPelayanan}")
    .usinstalasiasal.SetUnboundFieldSource ("{ado.InstalasiAsal}")
    .UsJK.SetUnboundFieldSource ("{ado.JK}")
    .txtpilihan.SetText ("Instalasi Asal")
    .UJml.SetUnboundFieldSource ("{ado.JmlPelayanan}")
'    .txtjudul.SetText Judul7

     settingreport rpttahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
   End With
    
Screen.MousePointer = vbHourglass
With CRViewer1
        .ReportSource = rpttahun
        .ViewReport
        .Zoom (75)
End With
Screen.MousePointer = vbDefault
Set frmcetakviewer = Nothing

End Sub
Private Sub LapKunjunganJenisTindakanBJenisPtahun()
Call openConnection
    Dim adocomd As New ADODB.Command
     Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With rpttahun
    .Database.AddADOCommand dbConn, adocomd
    .txtNamaRS.SetText strNNamaRS
    .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    .txtAlamat2.SetText strEmail
    .txtinstalasi.SetText MstrInstalasi2
    .txtjudul.SetText ("LAPORAN REKAPITULASI KUNJUNGAN PASIEN BERDASARKAN JENIS TINDAKAN ")
        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
             .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
             .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
   
    .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
    .Udate.SetUnboundFieldSource ("{ado.TglPelayanan}")
    .usJenisPelayanan.SetUnboundFieldSource ("{ado.JenisPelayanan}")
    .usinstalasiasal.SetUnboundFieldSource ("{ado.JenisPasien}")
    .UsJK.SetUnboundFieldSource ("{ado.JK}")
    .txtpilihan.SetText ("Jenis Pasien")
    .UJml.SetUnboundFieldSource ("{ado.JmlPelayanan}")
'    .txtjudul.SetText Judul7
     settingreport rpttahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
   End With
    
Screen.MousePointer = vbHourglass
With CRViewer1
        .ReportSource = rpttahun
        .ViewReport
        .Zoom (75)
End With
Screen.MousePointer = vbDefault
Set frmcetakviewer = Nothing
End Sub
Private Sub LapKunjunganJenisPeriksaInstalasiAsalTahun()
Call openConnection
    Dim adocomd As New ADODB.Command
     Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
   adocomd.CommandText = strSQL
    With rpttahun
    .Database.AddADOCommand dbConn, adocomd
    .txtNamaRS.SetText strNNamaRS
    .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    .txtAlamat2.SetText strEmail
    .txtinstalasi.SetText MstrInstalasi2
    .txtjudul.SetText ("LAPORAN REKAPITULASI KUNJUNGAN PASIEN BERDASARKAN JENIS TINDAKAN ")
        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
             .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
             .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
    .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
    .Udate.SetUnboundFieldSource ("{ado.TglPelayanan}")
    .usJenisPelayanan.SetUnboundFieldSource ("{ado.JenisPeriksa}")
    .usinstalasiasal.SetUnboundFieldSource ("{ado.InstalasiAsal}")
    .UsJK.SetUnboundFieldSource ("{ado.JK}")
    .txtpilihan.SetText ("Instalasi Asal")
    .UJml.SetUnboundFieldSource ("{ado.JmlPelayanan}")
    If sUkuranKertas = "" Then
    sUkuranKertas = "3"
    sOrientasKertas = "2"
    sDuplex = "1"
    sDriver = "1"
    End If
     settingreport rpttahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    Call ukukertasperTahuan
   End With
    
Screen.MousePointer = vbHourglass
With CRViewer1
        .ReportSource = rpttahun
        .ViewReport
        .Zoom (75)
End With
Screen.MousePointer = vbDefault
Set frmcetakviewer = Nothing
End Sub
Private Sub ukukertasperTahuan()
With rpttahun
Select Case sUkuranKertas
   '-----------------------------
   Case "0" 'Depault Setting/legal
     
     .Priode.Left = 12000
     .txtjudul.Width = 13080
     .Priode.Width = 3240
     .Halaman.Width = 3240
   Case "11"
    .Priode.Left = 12000
    .txtjudul.Width = 20000
    .Halaman.Width = 20000
   Case "1" 'Letter
    .Priode.Left = 11000
     .txtjudul.Width = 18000
     .Priode.Width = 6000
     .Halaman.Width = 3240
   '-----------------------------
   Case "3" 'legal
    .Priode.Left = 12000
    .txtjudul.Width = 18840
    .Halaman.Width = 18840
   '----------------------------
   Case "6" ' A 3
     .Priode.Left = 12000
     .txtjudul.Width = 13080
     .Priode.Width = 3240
     .Halaman.Width = 3240
  '-----------------------------
    Case "7" ' A 4
    .Priode.Left = 12000
   .txtjudul.Width = 18840
   .Halaman.Width = 18840
End Select
End With
End Sub
Private Sub LapKunjunganJenisTindakanBJenisPasienBulan()
Call openConnection
    Dim adocomd As New ADODB.Command
     Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
   With Report
    .Database.AddADOCommand dbConn, adocomd
    .txtNamaRS.SetText strNNamaRS
    .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    .txtAlamat2.SetText strEmail
    .txtinstalasi.SetText MstrInstalasi2
    .txtjudul.SetText ("LAPORAN REKAPITULASI KUNJUNGAN PASIEN BERDASARKAN JENIS TINDAKAN ")
        If CStr(Format(mdTglAwal, "MMMM-yyyy")) = CStr(Format(mdTglAkhir, "MMMM-yyyy")) Then
              .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
              .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If
    .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
    .UdPeriode.SetUnboundFieldSource ("{ado.TglPelayanan}")
    .UsJnsPelayanan.SetUnboundFieldSource ("{ado.JenisPelayanan}")
    .UsInstasal.SetUnboundFieldSource ("{ado.JenisPasien}")
    .UsJK.SetUnboundFieldSource ("{ado.JK}")
    .txtjenis.SetText ("Jenis pasien")
    .JMlPasien.SetUnboundFieldSource ("{ado.JmlPelayanan}")
'    .txtjudul.SetText Judul7
    If sUkuranKertas = "" Then
    sUkuranKertas = "1"
    sOrientasKertas = "2"
    End If
     settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
   End With
    
Screen.MousePointer = vbHourglass
With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (75)
End With
Screen.MousePointer = vbDefault
Set frmcetakviewer = Nothing
End Sub
Private Sub LapKunjunganJenisTindakanBJenisPasienHari()
Call openConnection
    Dim adocomd As New ADODB.Command
     Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
   With Report
    .Database.AddADOCommand dbConn, adocomd
    .txtNamaRS.SetText strNNamaRS
    .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    .txtAlamat2.SetText strEmail
    .txtinstalasi.SetText MstrInstalasi2
    .txtjudul.SetText ("LAPORAN REKAPITULASI KUNJUNGAN PASIEN BERDASARKAN JENIS TINDAKAN ")
    If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
       .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
    Else
       .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
    End If
    
    '******** Jika Total *************
    If Strcetak4 = "Total" Then
        .CrossTab1.Suppress = True
        .CrossTab2.Suppress = False
        .TtxtPriode.SetText "Ruangan"
        .TxtRuangan.SetText "Jenis Pelayanan"
        .txtjenis.SetText "Instalasi Asal"
        .TtxtPriode.Left = .TtxtPriode.Left + 80
        .TxtRuangan.Left = .TxtRuangan.Left + 700
        .txtjenis.Left = .txtjenis.Left - 1600
        .TxtPelayanan.Suppress = True
        .LineV4.Suppress = True
        .LineV2.Left = 2510
        .LineV3.Left = 5390
    End If

    .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
    .UdPeriode.SetUnboundFieldSource ("{ado.TglPelayanan}")
    .UsJnsPelayanan.SetUnboundFieldSource ("{ado.JenisPelayanan}")
    .UsInstasal.SetUnboundFieldSource ("{ado.JenisPasien}")
    .UsJK.SetUnboundFieldSource ("{ado.JK}")
    .txtjenis.SetText ("Jenis Pasien")
    .JMlPasien.SetUnboundFieldSource ("{ado.JmlPelayanan}")
'    .txtjudul.SetText Judul7
    If sUkuranKertas = "" Then
    sUkuranKertas = "1"
    sOrientasKertas = "2"
    End If
     settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
     Call ukurkertas
   End With
    
Screen.MousePointer = vbHourglass
With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (75)
End With
Screen.MousePointer = vbDefault
End Sub
Private Sub LapKunjunganJenisTindakanBinstalasiAsalBulan()
Call openConnection
    Dim adocomd As New ADODB.Command
     Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
   With Report
    .Database.AddADOCommand dbConn, adocomd
    .txtNamaRS.SetText strNNamaRS
    .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    .txtAlamat2.SetText strEmail
    .txtinstalasi.SetText MstrInstalasi2
    .txtjudul.SetText ("LAPORAN REKAPITULASI KUNJUNGAN PASIEN BERDASARKAN JENIS TINDAKAN ")
        If CStr(Format(mdTglAwal, "MMMM-yyyy")) = CStr(Format(mdTglAkhir, "MMMM-yyyy")) Then
              .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
              .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If

    .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
    .UdPeriode.SetUnboundFieldSource ("{ado.TglPelayanan}")
    .UsJnsPelayanan.SetUnboundFieldSource ("{ado.JenisPelayanan}")
    .UsInstasal.SetUnboundFieldSource ("{ado.InstalasiAsal}")
    .UsJK.SetUnboundFieldSource ("{ado.JK}")
    .txtjenis.SetText ("Instalasi Asal")
    .JMlPasien.SetUnboundFieldSource ("{ado.JmlPelayanan}")
    If sUkuranKertas = "" Then
    sUkuranKertas = "1"
    sOrientasKertas = "2"
    End If
    settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    Call ukurkertas
   End With
    
Screen.MousePointer = vbHourglass
With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (75)
End With
Screen.MousePointer = vbDefault
End Sub
Private Sub LapKunjunganJenisTindakanBinstalasiAsal()
Call openConnection
    Dim adocomd As New ADODB.Command
     Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    
   With Report
    .Database.AddADOCommand dbConn, adocomd
    .txtNamaRS.SetText strNNamaRS
    .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    .txtAlamat2.SetText strEmail
    .txtinstalasi.SetText MstrInstalasi2
    .txtjudul.SetText ("LAPORAN REKAPITULASI KUNJUNGAN PASIEN BERDASARKAN JENIS TINDAKAN ")
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
          .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
          .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
    
    '******** Jika Total *************
    If Strcetak4 = "Total" Then
        .CrossTab1.Suppress = True
        .CrossTab2.Suppress = False
        .TtxtPriode.SetText "Ruangan"
        .TxtRuangan.SetText "Jenis Pelayanan"
        .txtjenis.SetText "Instalasi Asal"
        .TtxtPriode.Left = .TtxtPriode.Left + 80
        .TxtRuangan.Left = .TxtRuangan.Left + 700
        .txtjenis.Left = .txtjenis.Left - 1600
        .TxtPelayanan.Suppress = True
        .LineV4.Suppress = True
        .LineV2.Left = 2510
        .LineV3.Left = 5390
    End If
    
    .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
    .UdPeriode.SetUnboundFieldSource ("{ado.TglPelayanan}")
    .UsJnsPelayanan.SetUnboundFieldSource ("{ado.JenisPelayanan}")
    .UsInstasal.SetUnboundFieldSource ("{ado.InstalasiAsal}")
    .UsJK.SetUnboundFieldSource ("{ado.JK}")
    .txtjenis.SetText ("Instalasi Asal")
    .JMlPasien.SetUnboundFieldSource ("{ado.JmlPelayanan}")
    If sUkuranKertas = "" Then
    sUkuranKertas = "1"
    sOrientasKertas = "2"
    End If
    settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    Call ukurkertas
   End With
    
Screen.MousePointer = vbHourglass
With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (75)
End With
Screen.MousePointer = vbDefault
End Sub
Private Sub lakunjunganbulaninstalasiasal()
 Call openConnection
    Dim adocomd As New ADODB.Command
     Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
   With Report
     .Database.AddADOCommand dbConn, adocomd
    
        If CStr(Format(mdTglAwal, "MMMM-yyyy")) = CStr(Format(mdTglAkhir, "MMMM-yyyy")) Then
              .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
              .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If
    .txtNamaRS.SetText strNNamaRS
    .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    .txtAlamat2.SetText strEmail
    .txtinstalasi.SetText MstrInstalasi2
    .UdPeriode.SetUnboundFieldSource ("{ado.TglPelayanan}")
    .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
    .UsJnsPelayanan.SetUnboundFieldSource ("{ado.JenisPeriksa}")
    .UsInstasal.SetUnboundFieldSource ("{ado.InstalasiAsal}")
    .UsJK.SetUnboundFieldSource ("{ado.JK}")
    .txtjenis.SetText ("Instalasi Asal")
    .JMlPasien.SetUnboundFieldSource ("{ado.JmlPelayanan}")
'    .txtjudul.SetText Judul7
    If sUkuranKertas = "" Then
    sUkuranKertas = "1"
    sOrientasKertas = "2"
    End If
    settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    Call ukurkertas
   End With
    
Screen.MousePointer = vbHourglass
With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (75)
End With
Screen.MousePointer = vbDefault
End Sub
Private Sub LapKunjunganJenisPeriksaHariperAsalInstalasi()
 Call openConnection
    Set frmCetakDaftarPasienRawatJalan = Nothing
    Dim adocomd As New ADODB.Command
     Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
   With Report
     .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
          .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
          .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
    .txtNamaRS.SetText strNNamaRS
    .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    .txtAlamat2.SetText strEmail
    .txtinstalasi.SetText MstrInstalasi2
    .txtjudul.SetText "LAPORAN KUNJUNGAN PASIEN BERDASARKAN JENIS PERIKSA"
    .UdPeriode.SetUnboundFieldSource ("{ado.TglPelayanan}")
    .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
    .UsJnsPelayanan.SetUnboundFieldSource ("{ado.JenisPeriksa}")
    .UsInstasal.SetUnboundFieldSource ("{ado.InstalasiAsal}")
    .UsJK.SetUnboundFieldSource ("{ado.JK}")
     .txtjenis.SetText ("Instalasi Asal")
    .JMlPasien.SetUnboundFieldSource ("{ado.JmlPelayanan}")
'    .txtjudul.SetText Judul7
    If sUkuranKertas = "" Then
    sUkuranKertas = "3"
    sOrientasKertas = "2"
    End If
    settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    Call ukurkertas
    
   End With
    
Screen.MousePointer = vbHourglass
With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (75)
End With
Screen.MousePointer = vbDefault
End Sub
Private Sub ukurkertas()
With Report
Select Case sUkuranKertas
   '-----------------------------
   Case "0" 'Depault Setting/legal
     .txtTanggal.Left = 12000
     .txtjudul.Width = 13080
     .txtTanggal.Top = 1560
     .txtTanggal.Width = 3240
     .Halaman.Width = 3240
   Case "11"
    .txtTanggal.Left = 12000
    .txtjudul.Width = 20000
    .Halaman.Width = 20000
   Case "1" 'Letter
    .txtTanggal.Left = 11000
     .txtjudul.Width = 18000
     .txtTanggal.Top = 1560
     .txtTanggal.Width = 6000
     .Halaman.Width = 3240
   '-----------------------------
   Case "3" 'legal
    .txtTanggal.Left = 15840
    .txtjudul.Width = 18840
    .Halaman.Width = 18840
   '----------------------------
   Case "6" ' A 3
     .txtTanggal.Left = 12000
     .txtjudul.Width = 13080
     .txtTanggal.Top = 1560
     .txtTanggal.Width = 3240
     .Halaman.Width = 3240
  '-----------------------------
    Case "7" ' A 4
    .txtTanggal.Left = 12000
   .txtjudul.Width = 18840
   .Halaman.Width = 18840
End Select
End With
End Sub
Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub


Private Sub lapkunjunganperhariperjenispasien()

 Call openConnection
    Set frmCetakDaftarPasienRawatJalan = Nothing
    Dim adocomd As New ADODB.Command
     Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
   With Report
     .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
          .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
          .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
    .txtNamaRS.SetText strNNamaRS
    .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    .txtAlamat2.SetText strEmail
    .txtinstalasi.SetText MstrInstalasi2
    .UdPeriode.SetUnboundFieldSource ("{ado.TglPelayanan}")
    .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
    .UsJnsPelayanan.SetUnboundFieldSource ("{ado.JenisPeriksa}")
    .UsInstasal.SetUnboundFieldSource ("{ado.JenisPasien}")
    .UsJK.SetUnboundFieldSource ("{ado.JK}")
    .txtjenis.SetText ("Jenis Pasien")
    .JMlPasien.SetUnboundFieldSource ("{ado.JmlPelayanan}")
    If sUkuranKertas = "" Then
    sUkuranKertas = "3"
    sOrientasKertas = "2"
    End If
    settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    Call ukurkertas
   End With
    
Screen.MousePointer = vbHourglass
With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (75)
End With
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmcetakviewer = Nothing
sUkuranKertas = ""
sOrientasKertas = ""
sDuplex = ""
sDriver = ""
End Sub
