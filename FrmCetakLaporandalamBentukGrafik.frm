VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmCetakLaporandalamBentukGrafik 
   Caption         =   "Cetak Laporan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmCetakLaporandalamBentukGrafik.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
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
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FrmCetakLaporandalamBentukGrafik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New Cr_RekapGrafik
Dim ReportPerTahun As New Cr_RekapGrafikPerTahun
Dim RptTotal As New Cr_RekapGrafikPerTotal
Dim Judul1 As String

Private Sub Form_Load()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Select Case strCetak2
        Case "LapKunjunganJenisStatusHari"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS PASIEN "
            Call KunjunganBjenisBStatusHari

        Case "LapKunjunganJenisStatusBulan"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS PASIEN "
            Call KunjunganBjenisBStatusBulan

        Case "LapKunjunganJenisStatusTahun"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS PASIEN "
            Call KunjunganBjenisBStatusTahun

        Case "LapKunjunganJenisStatusTotal"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KASUS PENYAKIT  "
            Call RekapKunjunganPerTotal

            '======================================
        Case "LapKunjunganSt_PnyktPsnHari"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KASUS PENYAKIT "
            Call KunjunganBjenisBStatusHari

        Case "LapKunjunganSt_PnyktPsnBulan"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KASUS PENYAKIT "
            Call KunjunganBjenisBStatusBulan

        Case "LapKunjunganSt_PnyktPsnTahun"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KASUS PENYAKIT  "
            Call KunjunganBjenisBStatusTahun

        Case "LapKunjunganSt_PnyktPsnTotal"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KASUS PENYAKIT  "
            Call RekapKunjunganPerTotal

            '==========================================
        Case "LapKunjunganBwilayahHari"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN WILAYAH  "
            Call KunjunganBjenisBStatusHari

        Case "LapKunjunganBwilayahBulan"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS WILAYAH  "
            Call KunjunganBjenisBStatusBulan

        Case "LapKunjunganBwilayahTahun"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS WILAYAH "
            Call KunjunganBjenisBStatusTahun

        Case "LapKunjunganBwilayahTotal"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS WILAYAH "
            Call RekapKunjunganPerTotal

            '=======================================
        Case "LapKunjunganKelasStatusHari"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KELAS  "
            Call KunjunganBjenisBStatusHari

        Case "LapKunjunganKelasStatusBulan"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KELAS  "
            Call KunjunganBjenisBStatusBulan

        Case "LapKunjunganKelasStatusTahun"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KELAS "
            Call KunjunganBjenisBStatusTahun

        Case "LapKunjunganKelasStatusTotal"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KELAS  "
            Call RekapKunjunganPerTotal

            '=======================================
        Case "LapKunjunganRujukanBStatusHari"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN RUJUKAN "
            Call KunjunganBjenisBStatusHari
        Case "LapKunjunganRujukanBStatusBulan"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN RUJUKAN"
            Call KunjunganBjenisBStatusBulan
        Case "LapKunjunganRujukanBStatusTahun"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN RUJUKAN"
            Call KunjunganBjenisBStatusTahun

        Case "LapKunjunganRujukanBStatusTotal"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN RUJUKAN "
            Call RekapKunjunganPerTotal

            '=======================================
        Case "LapKunjunganKonPulang_StatusHari"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KONDISI PULANG  "
            Call LapKunjunganKonPulang_StatusHari

        Case "LapKunjunganKonPulang_StatusBulan"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KONDISI PULANG  "
            Call LapKunjunganKonPulang_StatusBulan

        Case "LapKunjunganKonPulang_StatusTahun"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KONDISI PULANG "
            Call LapKunjunganKonPulang_StatusTahun

        Case "LapKunjunganKonPulang_StatusTotal"
            '         Call LapKunjunganKonPulang_StatusTotal
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KONDISI PULANG "
            '=======================================
        Case "LapKunjunganJenisOperasi_StatusHari"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS OPERASI "
            Call KunjunganBjenisBStatusHari

        Case "LapKunjunganJenisOperasi_StatusBulan"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS OPERASI "
            Call KunjunganBjenisBStatusBulan

        Case "LapKunjunganJenisOperasi_StatusTahun"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS OPERASI "
            Call KunjunganBjenisBStatusTahun

        Case "LapKunjunganJenisOperasi_StatusTotal"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS OPERASI "
            Call KunjunganBjenisBStatusTotal

            '================================================
        Case "LapKunjunganBjenisTindakanHari"
            Call KunjunganBjenisBStatusHari
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS OPERASI "

            '==================================================
        Case "LapKunjunganTriaseStatusHari"
            Call KunjunganBjenisBStatusHari
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN TRIASE  "
        Case "LapKunjunganTriaseStatusBulan"
            Call KunjunganBjenisBStatusBulan
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN TRIASE  "
        Case "LapKunjunganTriaseStatusTahun"
            Call KunjunganBjenisBStatusTahun
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN TRIASE "
        Case "LapKunjunganTriaseStatusTotal"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN TRIASE  "
            Call RekapKunjunganPerTotal

    End Select
End Sub

Private Sub RekapKunjunganPerTotal()
    Set frmCetakDaftarPasienRawatJalan = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With RptTotal
        If rs.RecordCount < 100 Then
            .Graph1.Width = 9240
            .txtJudul.Width = 11280
            .Periode.Left = 6720
            .txtfootjudul.Width = 11280
            If sUkuranKertas = "" Then
                sUkuranKertas = "5"
                sOrientasKertas = "1"
                sDuplex = "0"
            End If
        Else
            .Graph1.Width = 18720
            .Graph1.Left = 120
            If sUkuranKertas = "" Then
                sUkuranKertas = "5"
                sOrientasKertas = "2"
                sDuplex = "0"
            End If
        End If

        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .usjudul.SetUnboundFieldSource ("{ado.Judul}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.Detail}")
        .usJK.SetUnboundFieldSource ("{ado.Jk}")
        .unJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtJudul.SetText Judul1
        settingreport RptTotal, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = RptTotal
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub KunjunganBjenisBStatusHari()
    Set frmCetakDaftarPasienRawatJalan = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        If rs.RecordCount < 100 Then
            .Graph1.Width = 9240
            .txtJudul.Width = 11280
            .Periode.Left = 6720
            .txtfootjudul.Width = 11280
            If sUkuranKertas = "" Then
                sUkuranKertas = "5"
                sOrientasKertas = "1"
                sDuplex = "0"
            End If
        Else
            .Graph1.Width = 18720
            .Graph1.Left = 120
            If sUkuranKertas = "" Then
                sUkuranKertas = "5"
                sOrientasKertas = "2"
                sDuplex = "0"
            End If
        End If

        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .usjudul.SetUnboundFieldSource ("{ado.Judul}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.Detail}")
        .usJK.SetUnboundFieldSource ("{ado.Jk}")
        .unJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtJudul.SetText Judul1

        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub KunjunganBjenisBStatusBulan()
    Call openConnection
    Set frmCetakDaftarPasienRawatJalan = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With Report
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .usjudul.SetUnboundFieldSource ("{ado.Judul}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.Detail}")
        .usJK.SetUnboundFieldSource ("{ado.Jk}")
        .unJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub KunjunganBjenisBStatusTahun()
    Call openConnection
    Set frmCetakDaftarPasienRawatJalan = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With ReportPerTahun
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .usjudul.SetUnboundFieldSource ("{ado.Judul}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.Detail}")
        .usJK.SetUnboundFieldSource ("{ado.Jk}")
        .unJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport ReportPerTahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = ReportPerTahun
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub KunjunganBjenisBStatusTotal()
    Call openConnection
    Set frmCetakDaftarPasienRawatJalan = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .usjudul.SetUnboundFieldSource ("{ado.Judul}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.Detail}")
        .usJK.SetUnboundFieldSource ("{ado.Jk}")
        .unJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport ReportPerTahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganKonPulang_StatusHari()
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
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If

        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.tglkeluar}")
        .usjudul.SetUnboundFieldSource ("{ado.Judul}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.Detail}")
        .usJK.SetUnboundFieldSource ("{ado.Jk}")
        .unJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If

        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganKonPulang_StatusBulan()
    Call openConnection
    Set frmCetakDaftarPasienRawatJalan = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If

        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.tglkeluar}")
        .usjudul.SetUnboundFieldSource ("{ado.Judul}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.Detail}")
        .usJK.SetUnboundFieldSource ("{ado.Jk}")
        .unJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganKonPulang_StatusTahun()
    Call openConnection
    Set frmCetakDaftarPasienRawatJalan = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With ReportPerTahun
        .Database.AddADOCommand dbConn, adocomd
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If

        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.tglkeluar}")
        .usjudul.SetUnboundFieldSource ("{ado.Judul}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.Detail}")
        .usJK.SetUnboundFieldSource ("{ado.Jk}")
        .unJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")

        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport ReportPerTahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = ReportPerTahun
        .ViewReport
        .Zoom (98)
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
    Set FrmCetakLaporandalamBentukGrafik = Nothing
    sUkuranKertas = ""
End Sub
