VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakKunjunganRS 
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8835
   Icon            =   "frmCetakKunjunganRS.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   8835
   WindowState     =   2  'Maximized
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakKunjunganRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crPengunjungRJJenisPembayaran
Dim Report1 As New crRekapUTD
Dim Report2 As New crDaftarPasienKecelakaan
Dim report3 As New crDataPasienRJByKodeWilayah
Dim report4 As New rptWilayahJekelRI
Dim Report5 As New rptWilayahRI
Dim Report6 As New crJumlahPasienEKGUSG
Dim Report7 As New crRekapFisioterapi
Dim Report8 As New crPengunjungRJJenisPembayaran2
Dim Report9 As New crKunjPasienRJ
Dim Report10 As New cr_DPJP
Dim Report11 As New crKDRS
Dim Report12 As New crIndeksDiagnosa

Private Sub Form_Load()
    Set frmCetakKunjunganRS = Nothing
    Select Case strCetak
        Case "DPJP"
            Call DPJP
        Case "Jenis Pembayaran RJGD"
            Call JenisBayarRJGD
        Case "Jenis Pembayaran RJGD2"
            Call JenisBayarRJGD2
        Case "Rekapitulasi Jumlah Pemeriksaan UTD"
            Call RekapUTD
        Case "Daftar Pasien Kecelakaan"
            Call DaftarPasienKecelakaan
        Case "Data Pasien RJ By Kode Wilayah"
            If frmDataPengunjung.dcInstalasi.BoundText = "01" Or frmDataPengunjung.dcInstalasi.BoundText = "02" Or frmDataPengunjung.dcInstalasi.BoundText = "06" Then
                Call DataPasienRIJByKodeWilayah
            Else
                Call WilayahJekelRI2
            End If
        Case "WilayahJekelRI"
            Call WilayahJekelRI
        Case "Jumlah Pasien EKG"
            Call JumlahPasienEKG
        Case "Jumlah Pasien USG"
            Call JumlahPasienUSG
        Case "Rekap Fisioterapi"
            Call RekapFisioterapi
        Case "KunjRJ"
            Call KunjPasienRJ
        Case "TindakanOperasi"
            Call TindakanOperasi
        Case "KDRS"
            Call KDRS
        Case "IndeksPeny"
            Call IndeksDiagnosa
    End Select
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

Private Sub TindakanOperasi()
    Set frmCetakKunjunganRS = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    
    With Report10
        If frmDataPengunjung.dtpAwal <> frmDataPengunjung.dtpAkhir Then
            If frmDataPengunjung.optGroupBy(1).value = True Then
                .txtPeriode.SetText "Periode " & frmDataPengunjung.dtpAwal.Month & "-" & frmDataPengunjung.dtpAwal.Year & " s/d " & frmDataPengunjung.dtpAkhir.Month & "-" & frmDataPengunjung.dtpAkhir.Year
            Else
                .txtPeriode.SetText "Periode " & frmDataPengunjung.dtpAwal.Year & " s/d " & frmDataPengunjung.dtpAkhir.Year
            End If
        Else
            If frmDataPengunjung.optGroupBy(1).value = True Then
                .txtPeriode.SetText "Periode " & frmDataPengunjung.dtpAwal.Month & "-" & frmDataPengunjung.dtpAwal.Year
            Else
                .txtPeriode.SetText "Periode " & frmDataPengunjung.dtpAwal.Year
            End If
        End If
        .Text1.SetText ("JUMLAH TINDAKAN OPERASI PASIEN")
        .txtTgl.SetText Format(Now, "dd mmmm yyyy")
        .txtUser.SetText strNmPegawai
        .Database.AddADOCommand dbConn, adocomd
        .usNamaDokter.SetUnboundFieldSource ("{ado.NamaLengkap}")
        .usNamaPenjamin.SetUnboundFieldSource ("{ado.NamaPenjamin}")
        .unJml.SetUnboundFieldSource ("{ado.Jml}")
    End With
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = Report10
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub DPJP()
    Set frmCetakKunjunganRS = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    
    With Report10
        If frmDataPengunjung.dtpAwal <> frmDataPengunjung.dtpAkhir Then
            If frmDataPengunjung.optGroupBy(1).value = True Then
                .txtPeriode.SetText "Periode " & frmDataPengunjung.dtpAwal.Month & "-" & frmDataPengunjung.dtpAwal.Year & " s/d " & frmDataPengunjung.dtpAkhir.Month & "-" & frmDataPengunjung.dtpAkhir.Year
            Else
                .txtPeriode.SetText "Periode " & frmDataPengunjung.dtpAwal.Year & " s/d " & frmDataPengunjung.dtpAkhir.Year
            End If
        Else
            If frmDataPengunjung.optGroupBy(1).value = True Then
                .txtPeriode.SetText "Periode " & frmDataPengunjung.dtpAwal.Month & "-" & frmDataPengunjung.dtpAwal.Year
            Else
                .txtPeriode.SetText "Periode " & frmDataPengunjung.dtpAwal.Year
            End If
        End If
        .Text1.SetText ("JUMLAH PASIEN MENURUT DOKTER")
        .txtTgl.SetText Format(Now, "dd mmmm yyyy")
        .txtUser.SetText strNmPegawai
        .Database.AddADOCommand dbConn, adocomd
        .usNamaDokter.SetUnboundFieldSource ("{ado.NamaLengkap}")
        .usNamaPenjamin.SetUnboundFieldSource ("{ado.NamaPenjamin}")
        .unJml.SetUnboundFieldSource ("{ado.jml}")
    End With
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = Report10
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub KunjPasienRJ()
    Set frmCetakKunjunganRS = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    
    With Report9
        If frmDataPengunjung.dtpAwal <> frmDataPengunjung.dtpAkhir Then
            .txtPeriode.SetText "Tahun " & frmDataPengunjung.dtpAwal.Year & " s/d " & frmDataPengunjung.dtpAkhir.Year
        Else
            .txtPeriode.SetText "Tahun " & frmDataPengunjung.dtpAwal.Year
        End If
        .Database.AddADOCommand dbConn, adocomd
        .usBulan2.SetUnboundFieldSource ("{ado.BlnTglMasuk}")
        .usBulan.SetUnboundFieldSource ("{ado.Bln}")
        .usRuangan.SetUnboundFieldSource ("{ado.NamaRuangan}")
        .usStatus.SetUnboundFieldSource ("{ado.StatusPasien}")
        .unJml.SetUnboundFieldSource ("{ado.Jml}")
    End With
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = Report9
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub RekapFisioterapi()
    Set frmCetakKunjunganRS = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report7
        .txtTgl.SetText Format(Now, "dd mmmm yyyy")
        .txtUser.SetText strNmPegawai
        .Text2.SetText "RSUD SAWAHLUNTO TAHUN " & frmDataPengunjung.dtpAwal.Year
        .Database.AddADOCommand dbConn, adocomd
        .usNamaPelayanan.SetUnboundFieldSource ("{ado.NamaPelayanan}")
        .unBulan.SetUnboundFieldSource ("{ado.bulan1}")
        .usBulan.SetUnboundFieldSource ("{ado.bulan}")
        .usInstalasi.SetUnboundFieldSource ("{ado.Instalasi}")
        .unJumlah.SetUnboundFieldSource ("{ado.jumlah}")
    End With
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = Report7
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub JenisBayarRJGD()
    Set frmCetakKunjunganRS = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd

'        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
''            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
'        Else
''            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
'        End If
        If frmDataPengunjung.dcRuangan.Text = "" Then
            .txtJudul.SetText "DATA PENGUNJUNG " & frmDataPengunjung.dcInstalasi.Text & " DI RSUD SAWAHLUNTO"
        Else
            .txtJudul.SetText "DATA PENGUNJUNG " & UCase(frmDataPengunjung.dcInstalasi.Text) & " RUANGAN " & UCase(frmDataPengunjung.dcRuangan.Text) & " DI RSUD SAWAHLUNTO"
        End If
        
        .strBulan2.SetUnboundFieldSource ("{ado.Bulan2}")
        .strBulan.SetUnboundFieldSource ("{ado.Bulan}")
        .strTahun.SetUnboundFieldSource ("{ado.Tahun}")
        .strPenjamin.SetUnboundFieldSource ("{ado.KelompokPasien}")
        .strWilayah.SetUnboundFieldSource ("{ado.Kriteria}")
        .unJumlah.SetUnboundFieldSource ("{ado.JmlPasien}")
    End With
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub JenisBayarRJGD2()
    Set frmCetakKunjunganRS = Nothing
    Dim adocmd As New ADODB.Command
    Set adocmd = Nothing
    Me.WindowState = 2
    adocmd.ActiveConnection = dbConn
    adocmd.CommandText = strSQL
    
    With Report8
        If frmDataPengunjung.optGroupBy(0).value = True Then
            .txtPeriode.SetText "Periode : " & Format(frmDataPengunjung.dtpAwal, "dd MMM yyyy") & " s/d " & Format(frmDataPengunjung.dtpAkhir, "dd MMM yyyy")
        ElseIf frmDataPengunjung.optGroupBy(1).value = True Then
            .txtPeriode.SetText "Periode : " & Format(frmDataPengunjung.dtpAwal, "MMM yyyy") & " s/d " & Format(frmDataPengunjung.dtpAkhir, "MMM yyyy")
        Else
            .txtPeriode.SetText "Periode : " & Format(frmDataPengunjung.dtpAwal, "yyyy") & " s/d " & Format(frmDataPengunjung.dtpAkhir, "yyyy")
        End If
        
        .Database.AddADOCommand dbConn, adocmd
        .usPenjamin.SetUnboundFieldSource ("{ado.KelompokPasien}")
        .usDesk.SetUnboundFieldSource ("{ado.Kriteria}")
        .usKelas.SetUnboundFieldSource ("{ado.DeskKelas}")
        .unJml.SetUnboundFieldSource ("{ado.Jml}")
    End With
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = Report8
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub RekapUTD()
    Set frmCetakKunjunganRS = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report1
        .txtTgl.SetText Format(Now, "dd mmmm yyyy")
        .txtUser.SetText strNmPegawai
        .Text2.SetText "RSUD SAWAHLUNTO TAHUN " & frmDataPengunjung.dtpAwal.Year
        .Database.AddADOCommand dbConn, adocomd
        .usBulan2.SetUnboundFieldSource ("{ado.Bulan2}")
        .usBulan.SetUnboundFieldSource ("{ado.Bulan}")
        .usTahun.SetUnboundFieldSource ("{ado.Tahun}")
        .usNamaPelayanan.SetUnboundFieldSource ("{ado.NamaPelayanan}")
        .unJumlahPelayanan.SetUnboundFieldSource ("{ado.JumlahPelayanan}")
    End With
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = Report1
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub DaftarPasienKecelakaan()
    Set frmCetakKunjunganRS = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report2
        .txtTgl.SetText Format(Now, "dd mmmm yyyy")
        .txtUser.SetText strNmPegawai
        .Database.AddADOCommand dbConn, adocomd
        .usNama.SetUnboundFieldSource ("{ado.NamaLengkap}")
        .usUmur.SetUnboundFieldSource ("{ado.Umur}")
        .udtTanggal.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .usDiagnosa.SetUnboundFieldSource ("{ado.Diagnosa}")
        .usInstalasi.SetUnboundFieldSource ("{ado.NamaInstalasi}")
    End With
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = Report2
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub DataPasienRIJByKodeWilayah()
    Set frmCetakKunjunganRS = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With report3
        If frmDataPengunjung.dcInstalasi.BoundText <> "01" Then
            .txtJudul.SetText "DATA PASIEN " & UCase(frmDataPengunjung.dcInstalasi.Text)
        Else
            .txtJudul.SetText "DATA PASIEN " & UCase(frmDataPengunjung.dcInstalasi.Text) & " RUANGAN " & UCase(frmDataPengunjung.dcRuangan.Text)
        End If
        .txtTgl.SetText Format(Now, "dd mmmm yyyy")
        .txtUser.SetText strNmPegawai
        .Text8.SetText frmDataPengunjung.dtpAwal.Year
        .Database.AddADOCommand dbConn, adocomd
        .usBulan1.SetUnboundFieldSource ("{ado.Bulan1}")
        .usBulan.SetUnboundFieldSource ("{ado.Bulan}")
        
        .usTahun.SetUnboundFieldSource ("{ado.Tahun}")
        .unDalam.SetUnboundFieldSource ("{ado.Dalam}")
        .unLuar.SetUnboundFieldSource ("{ado.Luar}")
    End With
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = report3
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub WilayahJekelRI()
    Set frmCetakKunjunganRS = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With report4
        .txtTgl.SetText Format(Now, "dd mmmm yyyy")
        .txtUser.SetText strNmPegawai
'        .Text8.SetText frmDataPengunjung.dtpAwal.Year
        .Database.AddADOCommand dbConn, adocomd
        .usRuangan.SetUnboundFieldSource ("{ado.NamaRuangan}")
        .usWilayah.SetUnboundFieldSource ("{ado.Kriteria}")
        .usKriteria.SetUnboundFieldSource ("{ado.DeskKelas}")
        .usJenisKelamin.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .unJumlah.SetUnboundFieldSource ("{ado.Jml}")
    End With
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = report4
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub WilayahJekelRI2()
    Set frmCetakKunjunganRS = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report5
        .Text2.SetText "TAHUN " & frmDataPengunjung.dtpAwal.Year
        .txtUser.SetText strNmPegawai
        .txtTgl.SetText Format(Now, "dd mmmm yyyy")
        .Database.AddADOCommand dbConn, adocomd
        .unBulan.SetUnboundFieldSource ("{ado.Bulan1}")
        .usBulan.SetUnboundFieldSource ("{ado.Bulan}")
        .usWilayah.SetUnboundFieldSource ("{ado.Kriteria}")
        .usKelas.SetUnboundFieldSource ("{ado.DeskKelas}")
        .unJml.SetUnboundFieldSource ("{ado.Jml}")
    End With
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = Report5
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub JumlahPasienEKG()
    Set frmCetakKunjunganRS = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report6
        .Text7.SetText "JUMLAH PASIEN EKG"
        .Text8.SetText "DI RUANG " & UCase(frmDataPengunjung.dcRuangan.Text) & " RSUD SAWAHLUNTO"
        .Text9.SetText "TAHUN " & frmDataPengunjung.dtpAwal.Year
'        .txtUser.SetText strNmPegawai
'        .txtTgl.SetText Format(Now, "dd mmmm yyyy")
        .Database.AddADOCommand dbConn, adocomd
        .unBulan1.SetUnboundFieldSource ("{ado.Bulan1}")
        .usBulan.SetUnboundFieldSource ("{ado.Bulan}")
        .usJenisPembayaran.SetUnboundFieldSource ("{ado.NamaPenjamin}")
        .unJumlah.SetUnboundFieldSource ("{ado.Jumlah}")
    End With
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = Report6
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub JumlahPasienUSG()
    Set frmCetakKunjunganRS = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report6
        .Text7.SetText "JUMLAH PASIEN USG"
        .Text8.SetText "DI RUANG " & UCase(frmDataPengunjung.dcRuangan.Text) & " RSUD SAWAHLUNTO"
        .Text9.SetText "TAHUN " & frmDataPengunjung.dtpAwal.Year
'        .txtUser.SetText strNmPegawai
'        .txtTgl.SetText Format(Now, "dd mmmm yyyy")
        .Database.AddADOCommand dbConn, adocomd
        .unBulan1.SetUnboundFieldSource ("{ado.Bulan1}")
        .usBulan.SetUnboundFieldSource ("{ado.Bulan}")
        .usJenisPembayaran.SetUnboundFieldSource ("{ado.NamaPenjamin}")
        .unJumlah.SetUnboundFieldSource ("{ado.Jumlah}")
    End With
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = Report6
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub KDRS()
    Set frmCetakKunjunganRS = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report11
        .txtPeriode.SetText Format(frmDataPengunjung.dtpAwal.value, "dd mmmm yyyy") & " s/d " & Format(frmDataPengunjung.dtpAkhir.value, "dd mmmm yyyy")
        .txtTanggal.SetText "Sawahlunto " & Format(Now, "dd mmmm yyyy")
        .txtUser.SetText strNmPegawai
        
        .Database.AddADOCommand dbConn, adocomd
        .usDiagnosa.SetUnboundFieldSource ("{ado.KdDiagnosa}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaLengkap}")
        .unUmur.SetUnboundFieldSource ("{ado.Umur}")
        .usJenisKelamin.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .usNamaOrtu.SetUnboundFieldSource ("{ado.NamaIbu}")
        .usAlamat.SetUnboundFieldSource ("{ado.Alamat}")
        .udTglMasuk.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .udTglKeluar.SetUnboundFieldSource ("{ado.TglPulang}")
        .usKeadaanPulang.SetUnboundFieldSource ("{ado.KondisiPulang}")
        .usNamaRuangan.SetUnboundFieldSource ("{ado.NamaRuangan}")
    End With
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = Report11
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub IndeksDiagnosa()
    Set frmCetakKunjunganRS = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report12
        .txtPeriode.SetText Format(frmDataPengunjung.dtpAwal.value, "dd mmmm yyyy") & " s/d " & Format(frmDataPengunjung.dtpAkhir.value, "dd mmmm yyyy")
        .txtKdDiagnosa.SetText frmDataPengunjung.txtdiagnosa.Text
        .txtTanggal.SetText "Sawahlunto " & Format(Now, "dd mmmm yyyy")
        .txtUser.SetText strNmPegawai
        
        .Database.AddADOCommand dbConn, adocomd
        .usKdDiagnosa.SetUnboundFieldSource ("{ado.KdDiagnosa}")
        .usDiagnosa.SetUnboundFieldSource ("{ado.NamaDiagnosa}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaLengkap}")
        .udTglPendaftaran.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .udTglPulang.SetUnboundFieldSource ("{ado.TglPulang}")
        .unLos.SetUnboundFieldSource ("{ado.LOS}")
        .unUmur.SetUnboundFieldSource ("{ado.Umur}")
        .usRuangan.SetUnboundFieldSource ("{ado.NamaRuangan}")
        .unBiaya.SetUnboundFieldSource ("{ado.TotalBiaya}")
    End With
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = Report12
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

