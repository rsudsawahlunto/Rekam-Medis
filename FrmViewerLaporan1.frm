VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmViewerLaporan 
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   Icon            =   "FrmViewerLaporan1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   5850
   WindowState     =   2  'Maximized
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
      EnablePrintButton=   0   'False
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
Attribute VB_Name = "FrmViewerLaporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim reportTopten As New crDiagnosaTopTen2
Dim reporrtoptengrafik As New crDiagnosaTopTenGrafik
Dim reportBukuBesar As New crBukuBesar
Dim rptDPMeninggal As New crCetakDaftarPasienMeninggal
Dim rptPmeninggal As New crSuratKeteranganMeninggal
Dim adocomd As New ADODB.Command
Dim tanggal As String
Dim p As Printer
Dim tempPrint1 As String
Dim strDeviceName As String
Dim strDriverName As String
Dim strPort As String
Dim Posisi, z, Urutan As Integer




Private Sub Form_Load()
    Set adocomd = New ADODB.Command
    Set adocomd = Nothing
    adocomd.ActiveConnection = dbConn

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Dim tanggal As String

    Select Case cetak
        Case "RekapTopten"
            Call RekapTopTen
        Case "RekapToptenGrafik"
            Call RekapTopTenGrafik
        Case "BkRegisterIGD"
            Call BkRegisterIGD
        Case "BkRegisterRJ"
            Call BkRegisterRJ
        Case "BkRegisterRI"
            Call BkRegisterRI
        Case "DPMeninggal"
            Call DPmeninggal
        Case "PMeninggal"
            Call Pmeninggal
    End Select
End Sub

Private Sub RekapTopTen()
    adocomd.CommandText = "SELECT * FROM V_RekapitulasiDiagnosaTopTen " _
    & "WHERE (TglPeriksa BETWEEN '" _
    & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
    & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "') " _
    & " and kdinstalasi " & mstrFilter & " AND KdRuangan LIKE '%" & FrmPeriodeLaporanTopTen.dcRuangPoli.BoundText & "%' AND JenisPasien LIKE '%" & FrmPeriodeLaporanTopTen.dcJenisPasien & "%' ORDER BY instalasi,diagnosa"

    adocomd.CommandType = adCmdText
    reportTopten.Database.AddADOCommand dbConn, adocomd

    If Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") = Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy") Then
        tanggal = "Tanggal Kunjungan  : " & " " & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") '& " S/d " & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy")
    Else
        tanggal = "Periode Kunjungan  : " & " " & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") & " S/d " & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy")
    End If

    With reportTopten
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        .txtPeriode2.SetText tanggal
        .txtInstalasi.SetText FrmPeriodeLaporanTopTen.dcInstalasi.Text
        .txtRuangInstalasi.SetText FrmPeriodeLaporanTopTen.dcRuangPoli.Text
        .txtJenisPasien.SetText FrmPeriodeLaporanTopTen.dcJenisPasien
        .usSMF.SetUnboundFieldSource ("{ado.instalasi}")
        .usDiagnosa.SetUnboundFieldSource ("{ado.diagnosa}")
        .unJumlahPasien.SetUnboundFieldSource ("{ado.jumlahpasien}")
        .SelectPrinter sDriver, sPrinter, vbNull
        settingreport reportTopten, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = reportTopten
            .ViewReport
            .Zoom (100)
        End With
    Else
        reportTopten.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub RekapTopTenGrafik()
    adocomd.CommandText = "SELECT * FROM V_RekapitulasiDiagnosaTopTen " _
    & "WHERE (TglPeriksa BETWEEN '" _
    & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
    & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "')  " _
    & " and kdinstalasi " & mstrFilter & " AND KdRuangan LIKE '%" & FrmPeriodeLaporanTopTen.dcRuangPoli.BoundText & "%' AND JenisPasien LIKE '%" & FrmPeriodeLaporanTopTen.dcJenisPasien & "%' ORDER BY instalasi,diagnosa"

    adocomd.CommandType = adCmdText
    reporrtoptengrafik.Database.AddADOCommand dbConn, adocomd

    If Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") = Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy") Then
        tanggal = "Tanggal Kunjungan  : " & " " & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") '& " S/d " & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy")
    Else
        tanggal = "Periode Kunjungan  : " & " " & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") & " S/d " & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy")
    End If

    With reporrtoptengrafik
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        .Graph1.Data2LabelFont.Size = 6
        .Graph1.Data2TitleFont.Size = 6
        .Graph1.DataLabelFont.Size = 6
        .Graph1.DataTitleFont.Size = 6
        .Graph1.FootnoteFont.Size = 6
        .Graph1.GroupLabelFont.Size = 6
        .Graph1.GroupTitleFont.Size = 6
        .Graph1.SeriesLabelFont.Size = 6
        .Graph1.SeriesTitleFont.Size = 6
        .Graph1.SubTitleFont.Size = 6
        .Graph1.TitleFont.Size = 6
        .txtPeriode2.SetText tanggal
        .txtInstalasi.SetText FrmPeriodeLaporanTopTen.dcInstalasi.Text
        .txtRuangInstalasi.SetText FrmPeriodeLaporanTopTen.dcRuangPoli.Text
        .txtJenisPasien.SetText FrmPeriodeLaporanTopTen.dcJenisPasien
        .usSMF.SetUnboundFieldSource ("{ado.instalasi}")
        .usDiagnosa.SetUnboundFieldSource ("{ado.diagnosa}")
        .unJumlahPasien.SetUnboundFieldSource ("{ado.jumlahpasien}")
        .SelectPrinter sDriver, sPrinter, vbNull
        settingreport reporrtoptengrafik, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    If vLaporan = "view" Then
        Screen.MousePointer = reporrtoptengrafik
        With CRViewer1
            .ReportSource = reporrtoptengrafik
            .ViewReport
            .Zoom (100)
        End With
    Else
        reporrtoptengrafik.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub BkRegisterIGD()
'    strSQL = "SELECT * FROM V_BukuRegisterPasienIGD  " _
'    & "WHERE (TglMasuk BETWEEN '" _
'    & Format(FrmBukuRegister.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
'    & Format(FrmBukuRegister.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "')"
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText
    With reportBukuBesar
        .Text16.SetText strNNamaRS
        .Text18.SetText strNAlamatRS
        .txtJudul.SetText "DAFTAR PASIEN MASUK GAWAT DARURAT"
        .Text19.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
        .Database.AddADOCommand dbConn, adocomd
        .txtTgl.SetText Format(FrmBukuRegister.DTPickerAwal, "dd/MM/yyyy") & "  s/d  " & Format(FrmBukuRegister.DTPickerAkhir, "dd/MM/yyyy")
        .UnRuangan.SetUnboundFieldSource ("{ado.Ruangan}")
        .udtTglMasuk.SetUnboundFieldSource "{ado.tglmasuk}"
        .usNoReg.SetUnboundFieldSource "{ado.NoRegister}"
        .usCM.SetUnboundFieldSource "{ado.nocm}"
        .usPasien.SetUnboundFieldSource "{ado.namapasien}"
        .usAlamat.SetUnboundFieldSource "{ado.alamat}"

        If FrmBukuRegister.dcAsalPasien.Text = "" Then
            .usKecamatan.Suppress = True
            .txtKecamatan.Suppress = False
            .txtKecamatan.SetText "Semua Kecamatan"
        Else
            .usKecamatan.Suppress = False
            .usKecamatan.SetUnboundFieldSource "{ado.Kecamatan}"
        End If

        .usAgama.SetUnboundFieldSource "{ado.agama}"
        .usUmur.SetUnboundFieldSource "{ado.umur}"
        .usJK.SetUnboundFieldSource "{ado.jk}"
        .usStatus.SetUnboundFieldSource "{ado.statusPasien}"
        .usRujukan.SetUnboundFieldSource "{ado.AsalRujukan}"
        .usDiagnosa.SetUnboundFieldSource "{ado.diagnosa}"
        .usKlpkPasien.SetUnboundFieldSource "{ado.jenispasien}"

        If FrmBukuRegister.dcPenjamin.Text = "" Then
            .usPenjamin.Suppress = True
            .txtPenjamin.Suppress = False
            .txtPenjamin.SetText "Semua Penjamin"
        Else
            .usPenjamin.Suppress = False
            .usPenjamin.SetUnboundFieldSource "{ado.NamaPenjamin}"
        End If

        .usDokter.SetUnboundFieldSource "{ado.DokterPerawat}"
'        .SelectPrinter sDriver, sPrinter, vbNull
'        settingreport reportBukuBesar, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
    End With
    
sPrinter5 = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "Printer5")
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
        .ReportSource = reportBukuBesar
        .ViewReport
        .Zoom (100)

         Screen.MousePointer = vbDefault
        End With
    Else
        Urutan = 0
        For z = 1 To Len(sPrinter5)
            If Mid(sPrinter5, z, 1) = ";" Then
                Urutan = Urutan + 1
                Posisi = z
                ReDim Preserve arrPrinter(Urutan)
                arrPrinter(Urutan).intUrutan = Urutan
                arrPrinter(Urutan).intPosisi = Posisi
                If Urutan = 1 Then
                    arrPrinter(Urutan).strNamaPrinter = Mid(sPrinter5, 1, z - 1)
                Else
                    arrPrinter(Urutan).strNamaPrinter = Mid(sPrinter5, arrPrinter(Urutan - 1).intPosisi + 1, z - arrPrinter(Urutan - 1).intPosisi - 1)
                End If
             
             
            For Each p In Printers
                    strDeviceName = arrPrinter(Urutan).strNamaPrinter
                    strDriverName = p.DriverName
                    strPort = p.Port
        
                    reportBukuBesar.SelectPrinter strDriverName, strDeviceName, strPort
                    reportBukuBesar.PrintOut False
                    Screen.MousePointer = vbDefault

            Exit For
            
            Next
        End If
    Next z
      Unload Me
    End If
End Sub

Private Sub BkRegisterRI()
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText
    With reportBukuBesar
        .Text16.SetText strNNamaRS
        .Text18.SetText strNAlamatRS
        .txtJudul.SetText "DAFTAR PASIEN MASUK RAWAT INAP"
        .Text19.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos
        .Database.AddADOCommand dbConn, adocomd
        .txtTgl.SetText Format(FrmBukuRegister.DTPickerAwal, "dd/MM/yyyy") & "  s/d  " & Format(FrmBukuRegister.DTPickerAkhir, "dd/MM/yyyy")
        .UnRuangan.SetUnboundFieldSource ("{ado.NamaRuangan}")
        .udtTglMasuk.SetUnboundFieldSource "{ado.tglmasuk}"
        .usNoReg.SetUnboundFieldSource "{ado.NoRegister}"
        .usCM.SetUnboundFieldSource "{ado.nocm}"
        .usPasien.SetUnboundFieldSource "{ado.namapasien}"
        .usAlamat.SetUnboundFieldSource "{ado.alamat}"
        If FrmBukuRegister.dcAsalPasien.Text = "" Then
            .usKecamatan.Suppress = True
            .txtKecamatan.Suppress = False
            .txtKecamatan.SetText "Semua Kecamatan"
        Else
            .usKecamatan.Suppress = False
            .usKecamatan.SetUnboundFieldSource "{ado.Kecamatan}"
        End If
        .usAgama.SetUnboundFieldSource "{ado.pekerjaan}"
        .usUmur.SetUnboundFieldSource "{ado.umur}"
        .usJK.SetUnboundFieldSource "{ado.jk}"
        .usStatus.SetUnboundFieldSource "{ado.status}"
        .usRujukan.SetUnboundFieldSource "{ado.AsalRujukan}"
        .usDiagnosa.SetUnboundFieldSource "{ado.CaraMasuk}"
        .usKet.SetUnboundFieldSource "{ado.Kelas}"
        .usKlpkPasien.SetUnboundFieldSource "{ado.jenispasien}"

        If FrmBukuRegister.dcPenjamin.Text = "" Then
            .usPenjamin.Suppress = True
            .txtPenjamin.Suppress = False
            .txtPenjamin.SetText "Semua Penjamin"
        Else
            .usPenjamin.Suppress = False
            .usPenjamin.SetUnboundFieldSource "{ado.NamaPenjamin}"
        End If

        .usDokter.SetUnboundFieldSource "{ado.DokterPerawat}"
        .TxtCaraMasuk.SetText "Cara Masuk"
        .txtKelas.SetText "Kelas"
        .TxtPekerjaan.SetText "Pekerjaan"
'        .SelectPrinter sDriver, sPrinter, vbNull
'        settingreport reportBukuBesar, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
    End With
    
sPrinter5 = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "Printer5")
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
        .ReportSource = reportBukuBesar
        .ViewReport
        .Zoom (100)

         Screen.MousePointer = vbDefault
        End With
    Else
        Urutan = 0
        For z = 1 To Len(sPrinter5)
            If Mid(sPrinter5, z, 1) = ";" Then
                Urutan = Urutan + 1
                Posisi = z
                ReDim Preserve arrPrinter(Urutan)
                arrPrinter(Urutan).intUrutan = Urutan
                arrPrinter(Urutan).intPosisi = Posisi
                If Urutan = 1 Then
                    arrPrinter(Urutan).strNamaPrinter = Mid(sPrinter5, 1, z - 1)
                Else
                    arrPrinter(Urutan).strNamaPrinter = Mid(sPrinter5, arrPrinter(Urutan - 1).intPosisi + 1, z - arrPrinter(Urutan - 1).intPosisi - 1)
                End If
             
             
            For Each p In Printers
                    strDeviceName = arrPrinter(Urutan).strNamaPrinter
                    strDriverName = p.DriverName
                    strPort = p.Port
        
                    reportBukuBesar.SelectPrinter strDriverName, strDeviceName, strPort
                    reportBukuBesar.PrintOut False
                    Screen.MousePointer = vbDefault

            Exit For
            
            Next
        End If
    Next z
      Unload Me
    End If
End Sub

' =============== [ Buku Regiter RJ ] ==================
Private Sub BkRegisterRJ()
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText
    With reportBukuBesar
        .Text16.SetText strNNamaRS
        .Text18.SetText strNAlamatRS
        .txtJudul.SetText "DAFTAR PASIEN MASUK RAWAT JALAN"
        .Text19.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos
        .Database.AddADOCommand dbConn, adocomd
        .txtTgl.SetText Format(FrmBukuRegister.DTPickerAwal, "dd/MM/yyyy") & "  s/d  " & Format(FrmBukuRegister.DTPickerAkhir, "dd/MM/yyyy")
        .UnRuangan.SetUnboundFieldSource ("{Ado.ruangan}")
        .udtTglMasuk.SetUnboundFieldSource "{ado.tglmasuk}"
        .usNoReg.SetUnboundFieldSource "{ado.NoRegister}"
        .usCM.SetUnboundFieldSource "{ado.nocm}"
        .usPasien.SetUnboundFieldSource "{ado.namapasien}"
        .usAlamat.SetUnboundFieldSource "{ado.alamat}"
        If FrmBukuRegister.dcAsalPasien.Text = "" Then
            .usKecamatan.Suppress = True
            .txtKecamatan.Suppress = False
            .txtKecamatan.SetText "Semua Kecamatan"
        Else
            .usKecamatan.Suppress = False
            .usKecamatan.SetUnboundFieldSource "{ado.Kecamatan}"
        End If
        .usAgama.SetUnboundFieldSource "{ado.agama}"
        .usUmur.SetUnboundFieldSource "{ado.umur}"
        .usJK.SetUnboundFieldSource "{ado.jk}"
        .usStatus.SetUnboundFieldSource "{ado.statusPasien}"
        .usRujukan.SetUnboundFieldSource "{ado.AsalRujukan}"
        .usDiagnosa.SetUnboundFieldSource "{ado.diagnosa}"
        .usKlpkPasien.SetUnboundFieldSource "{ado.jenispasien}"

        If FrmBukuRegister.dcPenjamin.Text = "" Then
            .usPenjamin.Suppress = True
            .txtPenjamin.Suppress = False
            .txtPenjamin.SetText "Semua Penjamin"
        Else
            .usPenjamin.Suppress = False
            .usPenjamin.SetUnboundFieldSource "{ado.NamaPenjamin}"
        End If

        .usDokter.SetUnboundFieldSource "{ado.DokterPerawat}"
'        .SelectPrinter sDriver, sPrinter, vbNull
'        settingreport reportBukuBesar, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
    End With
    
    
sPrinter5 = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "Printer5")

        If vLaporan = "view" Then
            Screen.MousePointer = vbHourglass
            With CRViewer1
            .ReportSource = reportBukuBesar
            .ViewReport
            .Zoom (100)

             Screen.MousePointer = vbDefault
            End With
        Else
            Urutan = 0
            For z = 1 To Len(sPrinter5)
                If Mid(sPrinter5, z, 1) = ";" Then
                    Urutan = Urutan + 1
                    Posisi = z
                    ReDim Preserve arrPrinter(Urutan)
                    arrPrinter(Urutan).intUrutan = Urutan
                    arrPrinter(Urutan).intPosisi = Posisi
                    If Urutan = 1 Then
                        arrPrinter(Urutan).strNamaPrinter = Mid(sPrinter5, 1, z - 1)
                    Else
                        arrPrinter(Urutan).strNamaPrinter = Mid(sPrinter5, arrPrinter(Urutan - 1).intPosisi + 1, z - arrPrinter(Urutan - 1).intPosisi - 1)
                    End If
                 
                 
                For Each p In Printers
                        strDeviceName = arrPrinter(Urutan).strNamaPrinter
                        strDriverName = p.DriverName
                        strPort = p.Port
            
                        reportBukuBesar.SelectPrinter strDriverName, strDeviceName, strPort
                        reportBukuBesar.PrintOut False
                        Screen.MousePointer = vbDefault
    
                Exit For
                
                Next
            End If
        Next z
          Unload Me
        End If
End Sub

'========= [ Daftar PAsien Meninggal ] =====================
Private Sub DPmeninggal()
    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Me.Caption = "Medifirst2000 - Cetak Laporan Sensus Pelayanan"
    Set rptDPMeninggal = New crCetakDaftarPasienMeninggal
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText
    With rptDPMeninggal
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS
        .txtAlamat2.SetText strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtPeriode.SetText Format(frmDaftarPasienMeninggal.dtpAwal.value, "dd MMMM yyyy") & " s/d " & Format(frmDaftarPasienMeninggal.dtpAkhir.value, "dd MMMM yyyy")
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
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = rptDPMeninggal
        .ViewReport
        .Zoom (100)
    End With
    Screen.MousePointer = vbDefault
End Sub

'========= [ pasien Meninggal ] ====================
Private Sub Pmeninggal()
    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Me.Caption = "Medifirst2000 - Cetak Laporan Sensus Pelayanan"
    Set rs = Nothing
    Set rs = dbConn.Execute(strSQL)
    With rptPmeninggal
        .txtNamaRS.SetText strNNamaRS
        .txtNamaRSDetail.SetText strNNamaRS
        .Text1.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .txtKotaKodyaKab.SetText strNKotaRS
        .txtTahun.SetText Format(Now, "YYYY")
        .txtNoCM.SetText rs("NOCM")
        .txtNama.SetText rs("Nama Pasien")
        .txtumur.SetText rs("Umur")
        .txtJK.SetText IIf(rs("JK") = "L", "Laki-Laki", "Perempuan")
        .TxtPekerjaan.SetText IIf(IsNull(rs("Pekerjaan")), "-", rs("Pekerjaan"))
        .txtAlamat.SetText IIf(IsNull(rs("Alamat")), "-", rs("Alamat"))
        .TxtJam.SetText Format(rs("TglMeninggal"), "HH : MM : SS")
        .txtTgl.SetText Format(rs("TglMeninggal"), "DD - MMMM - YYYY")
        .txtDokter.SetText rs("Dokter Pemeriksa")
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = rptPmeninggal
        .ViewReport
        .Zoom (100)
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
    strStatus = ""
    strFilter = ""
    Set FrmViewerLaporan = Nothing
End Sub

