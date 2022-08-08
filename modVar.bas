Attribute VB_Name = "modVar"
Public jnsPelayanan As String
Public mstrNoSepRujukan As String

Public PrinterName As String
Public jmlPrint As Integer
Public strDeviceName As String
Public strDriverName As String
Public strPort As String

Public bolDetailPasienTombol As Boolean
Public tempkodeAntrian As String
Public strNoAntrian As String
Public strStsPasien As String
Public strNmPegawai As String
Public strKdSubInstalasi As String
Public strNNamaSubInstalasi As String
Public substrKdPegawai As String
Public strSQLIdentifikasi As String
Public strUserName As String
Public strPassword As String
Public TempKodeKelas As String
Public mstrKdAsalRujukan As String
Public strNoTerima As String
Public strjudulRuangan As String
Public flagJikaPasienSudahRujukan As String


Public AntrianForDataPasien As Boolean
Public ppkRujukan As String
Public mstrKelasDitanggung As String
Public mstrTglVerifBPJS As String
Public mstrkartuPeserta As String
Public mstrnorujukan As String
Public mstrJenisPeserta As String
Public mstrNamaAsalRujukanMon As String
Public mstrNamaPeserta As String

Public intCetakKe As Integer
Public strCetakKendaliLaporan As String
Public NamaCR As String
Public mstrKdDokterPerujuk As String
Public strKdRuanganReservasi As String
Public strNoReservasi As String
Public strKdDokterReservasi As String
Public strReservasi As String
Public intTahun As Integer
Public intBulan As Integer
Public intTanggal As Integer
Public blnMeTglLahEdi As Boolean
Public blnTambah As Boolean
Public blnOptBelum As Boolean
Public substrNoOrder As String
Public Periode As String
Public BlnAwal As String
Public BlnAkhir As String
Public ThnAwal As String
Public ThnAkhir As String
Public TglAwal As String
Public TglAkhir As String
Public strjudul As String
Public bolAntrian As Boolean

Public strNamaHostLocal As String
Public strKdAplikasi As String
Public dTglLogin As Date
Public dTglLogout As Date

Public dNTglBerlaku As Date
Public strNStsCITO As String
'Tuk Konvensional
Public fStatusPiutang As String
Public fStatusBayarSemua As String
'
'var Tuk cetak Kitir[F1]
Public dTglDaftar As Date
Public sRuangPeriksa As String
Public sNamaPasien As String
Public sJK As String
Public sUmur As String
Public sAlamat As String
Public sPenjamin As String
Public sKelas As String
Public sNoBed As String
Public tempKode As String
Public iNoAntrian As Integer
Public strhari As String
'end var kitir

Public strPathLogo As String
Public strNamaFileLogoRS As String
Public strNNamaRS As String
Public strNAlamatRS As String
Public strNTeleponRS As String
Public strNKotaRS As String
Public strNKodepos As String
Public strWebsite As String
Public strEmail As String
Public strKelasRS As String
Public strKetKelasRS As String
Public strPropinsi As String
Public blnError As Boolean

Public strNKdJenisTarif As String
Public strNKdJenisTarif2 As String
Public mstrKdInstalasi As String
Public mstrNamaInstalasi As String
Public kdsubinstalasi As String
Public idpegawai As String
Public strNIdPejamin As String
Public blnStatusAsuransi As Boolean
Public strNKdRuangan As String
Public strNNamaRuangan As String
Public strNNamaInstalasi As String
Public mstrCatatatnBPJS As String
Public strNoSEPForSimpan As String


Public strCetak As String
Public strCetak2 As String
Public strCetak3 As String
Public mstrCetak2 As String

Public mstrKdSubInstalasi As String
Public mstrKdRuangan As String
Public mstrKdRuanganLogin As String
Public mstrNamaRuangan As String
Public mstrKdRuanganPasien As String
Public mblnPsnMati As Boolean
Public mstrFilterDokter As String
Public mintJmlDokter As Integer
Public mintRowNow As Double
Public mdTglBerlaku As Date
Public mstrKdKelas As String
Public mdTglMasuk As Date

Public mstrNamaBarang As String
Public strPasien As String
Public mblnForm As Boolean
Public cetak As String
Type typeBarang
    strkdbarang As String
    strNamaBarang As String
    strKdAsal As String
    intJmlTerkecil As Double
    intJmlTempTotal As Double
End Type

Public mblnFormDaftarPasienRI As Boolean

Public typBarang() As typeBarang
Public ctk As String

Public boltampil As Boolean

Public mblnAdmin As Boolean
Public mblnVerifikator As Boolean

Public mstrLaporan As String
Public mstrKdDokter As String
Public mstrNamaDokter As String
Public strKodePelayananRS As String
Public mstrKdJenisLaporan As String
Public mstrStatusBayar As String
Public mstrNamaKomponenTarif As String
Public mstrFilterData As String
Public mstrNoLabRad As String

Public mstrKdJenisPasien As String
Public mstrNamaJenisPasien As String

Public mstrKdPenjaminPasien As String
Public mstrNamaPenjaminPasien As String
Public mstrFilter As String
Public mstrFilter2 As String
Public mstrFilterRuangan As String
Public mstrKriteria As String
Public mstrNoOrder As String
Public mstrGroup As String
Public mstrCetak As String
Public mstrNama As String
Public mstrIsiGroup As String

Public mstrKdKelompokBarang As String
Public mstrKdInstalasiNonMedis As String

Public strTampung As String

Public editpoli As Boolean
Public grafikkelompok As Boolean
Public grafikklinik As Boolean
Public nike As Boolean
Public noidpegawai As String
Public petugas As String
Public blnAdmin As Boolean
Public strUser As String
Public strPass As String
Public strPass2 As String
Public strPassEn As String
Public strStatus As String
Public strFilter As String
Public intLenUser As Integer
Public strIDPegawai As String
Public strIDPegawaiAktif As String
Public tgl As String
Public varcounter As Boolean
Public darah As Boolean
Public Alamat As Boolean
Public dadang As Boolean
Public dadang1 As Boolean
Public mstrNoCM2 As String
Public strSQL As String
Public strSQLX As String
Public strSQL1 As String
Public strSQL2 As String
Public strSQL3 As String
Public strSQL4 As String
Public strSQL5 As String
Public strSQL6 As String
Public strSQL7 As String
Public strSQL8 As String
Public strSQL9 As String
Public strSQL10 As String
Public strSQL11 As String
Public strSQL12 As String
Public blnEditPoli As Boolean
Public strJK As String
Public strNoStruk() As String
Public intJmlStruk As Integer
Public varNIPCek As String
Public varJmlHarusDibayarCek As String
Public varJenisPasienCek As String
Public varPembayaranKeCek As String
Public varNoStruk As String
Public varNoBKM As String
Public varNamaDokterCek As String
Public varAnPasienCek As String
Public mstrTglKirim As Date
'Kode Aplikasi yang sedang dijalankan, ganti sesuai keperluan
'**************************************
Public Const intAplikasi = 1
'**************************************

'variabel global koneksi & recordset ke db

'Hilangkan bila telah dideklarasikan sebelumnya
'**************************************
'Public dbConn As New ADODB.Connection
'**************************************

Public rsPegawai As New ADODB.recordset
Public rsPegawaiCount As New ADODB.recordset
Public rslogin As New ADODB.recordset
Public rsLoginApp As New ADODB.recordset
Public rsAplikasi As New ADODB.recordset
Public rsAplikasiCount As New ADODB.recordset
Public rsLoginCompare As New ADODB.recordset
Public strQuery As String

Public rsReservasi As New ADODB.recordset
Public rsb As New ADODB.recordset
Public rs As New ADODB.recordset
Public rsdiagnosa As New ADODB.recordset
Public rsProfil As New ADODB.recordset
Public rs2 As New ADODB.recordset
Public rsx As New ADODB.recordset
Public rsy As New ADODB.recordset
Public rs1 As New ADODB.recordset
Public rs3 As New ADODB.recordset
Public rs4 As New ADODB.recordset
Public rs5 As New ADODB.recordset
Public rsCek As New ADODB.recordset
Public rsWew As New ADODB.recordset
Public rsDuwew As New ADODB.recordset
Public rsDokumen As New ADODB.recordset
Public rsD As New ADODB.recordset

Public dbrs As New ADODB.recordset
Public adors As New ADODB.recordset
Public dbRst As New ADODB.recordset
Public dbRst2 As New ADODB.recordset
Public dbRst3 As New ADODB.recordset
Public dbRst4 As New ADODB.recordset
Public dbRst5 As New ADODB.recordset
Public dbRst6 As New ADODB.recordset
Public dbRst7 As New ADODB.recordset
Public dbRst8 As New ADODB.recordset
Public dbRst9 As New ADODB.recordset
Public dbConn As New ADODB.connection 'dipakai hampir disemua modul
Public dbcmd As New ADODB.Command
Public dbCmdSubReport As New ADODB.Command
Public adoComm As New ADODB.Command
Public dmParam As New ADODB.Parameter
Public dbcomm As New ADODB.Command

Public crxReport As CRAXDDRT.Report         'dipakai untuk modul crystal reports
'
'variabel servername, databasename, namarumahsakit dibuat global,
'karena nilai dari variabel2 ini akan dipakai untuk fungsi getsetting.
'
Public mstrNoPen As String
Public mstrDiagnosaUtama As String
Public mstrTglDiagnosaUtama As Date
Public mstrNoCM As String
Public mstrNoCMTerm As String
Public mstrIdPegawai As String
Public mstrTglKeluar As Date
Public strNamaRuangan As String
Public instalasi As Integer
Public KdInstalasi As String
Public SubInstalasi As Integer
Public ServerName As String
Public DatabaseName As String
Public UserID As String
Public UserName As String
Public NamaRumahSakit As String
'
Public NoCM As String
Public NamaPasien As String
Public isFindFirst As Boolean
Public enableEdit As Boolean
Public message As String
Public Umur As udt_umur
Public prmTgl As String
Public continue As Boolean
Public queryString As String
Public JenisKelamin As String
Public KodeInstalasi As String
Public KodeSubInstalasi As String
Public DataString As String
'
Public TtlHari As Integer

Type udt_umur
    tahun As Integer
    bulan As Integer
    hari As Integer
End Type

Public mdTglAwal As Date
Public mdTglAkhir As Date
Public mblnGrafik As Boolean
Public strDatabaseName As String
Public strServerName As String
Public YOC_intYear As Integer
Public YOC_intMonth As Integer
Public YOC_intDay As Integer

'setting printer
Public prn As Printer
Public sizepaper As CRPaperSize
Public duplexpaper As CRPrinterDuplexType
Public sPrinter As String, sDriver As String, sUkuranKertas As String
Public sDuplex As String, sOrientasKertas As String
Public tmpOrien As String

Type RecPrinter
    intUrutan As Integer
    intPosisi As Integer
    strNamaPrinter As String
End Type

Public arrPrinter() As RecPrinter

Public intTimerPrinter As String
Public sPrinter2 As String
Public sPrinter3 As String
Public sPrinter4 As String
Public sPrinter5 As String
Public sPrinterLabel1 As String
Public sPrinterLabel2 As String

'setting printer

Public mcurAll_TBP As Currency
Public mcurAll_TP As Currency
Public mcurAll_TRS As Currency
Public mcurAll_Pemb As Currency
Public mcurAll_HrsDibyr As Currency
Public mcurTM_TBP As Currency
Public mcurTM_TP As Currency
Public mcurTM_TRS As Currency
Public mcurTM_Pemb As Currency
Public mcurTM_Discount As Currency
Public mcurTM_HrsDibyr As Currency
Public mcurTM_HrsDibyrNow As Currency
Public mcurTM_JmlByr As Currency
Public mcurTM_ST As Currency
Public mcurOA_TBP As Currency
Public mcurOA_TP As Currency
Public mcurOA_TRS As Currency
Public mcurOA_Pemb As Currency
Public mcurOA_Discount As Currency
Public mcurOA_HrsDibyr As Currency
Public mcurOA_HrsDibyrNow As Currency
Public mcurOA_JmlByr As Currency
Public mcurOA_ST As Currency
Public mblnTM As Boolean
Public mblnOA As Boolean
Public mstrKdPenjamin As String
Public mcurBayar As Currency
Public mcurPembebasan As Currency
Public mstrNoStruk As String

Public strRegistrasi As String
Public blnCariPasien As Boolean
Public intJmlDokter  As Integer
Public mblnFormDaftarPasienIGD As Boolean
Public mblnFormDaftarAntrian As Boolean
Public vLaporan As String
Public mstrKdRuanganORS As String
Public mstrNoIBS As String
Public mblnTP As String
Public mstrKdJenisOperasi As String
Public mstrJenisOperasi  As String
Public subTanggalTerakhir As Integer
Public mblnOperator As Boolean
Public mstrKdInstalasiLogin As String
Public mstrNoHasilLab As String
Public mstrNoBKM As String
Public mstrNamaRuanganPerujuk As String
Public mstrKdInstalasiPerujuk As String
Public mstrNamaKelas As String
Public mstrNoLab  As String
Public mblnKonsul As Boolean
Public dTglRujukan As Date
Public mstrNoBKK As String
Public mstrNoRad As String
Public mstrJenisPasien As String
Public mstrRuanganPerujuk As String
Public mstrNoHasilRad As String
Public mblnCariPasien As Boolean
Public strNHargaSatuan As Double
Public strNTotal As Double
Public mstrServerPrinterBarcode As String
Public miRowNow As Double
Public mstrNamaRuanganPasien As String
Public mstrKelas As String
Public blnFrmCariPasien As Boolean
Public mstrKdRuanganKasir As String
Public vJudul As String
Public mblnTemp As Boolean

Public boolStafSIMRS As Boolean
Public FlagStatusPulang As String

Type PenjaminSisaTagihan
    strNamaLengkap As String
    dTglLahir As Date
    strJenisKelamin As String
    strNoIdentitas As String
    strHubungan As String
    strAlamat As String
    strTelepon As String
    strPropinsi As String
    strKota As String
    strKecamatan As String
    strKelurahan As String
    strRTRW As String
    strKodePos As String
    blnStatus As Boolean
End Type

Public typPenjaminSisaTagihan As PenjaminSisaTagihan
Public typPenjaminSisaTagihanApotik As PenjaminSisaTagihan
Public mblnAdaPlynTdkDibyr As Boolean
Public mcurDiscount As Currency
Public mblnTindakanKasir As Boolean
Public mLapPerParameter As String

Type typeSettingDataPendukung
    intTerminBayarFaktur As Integer
    realPPn As Double
    realLimitDiscount As Double
    realJasaPenulisResep As Double
    intJmlPembulatanHarga As Integer
End Type

Public typSettingDataPendukung As typeSettingDataPendukung
Public mstrValid As String
Public mstrNoTerima As String

Type Asuransi
    strIdPenjamin As String
    strIdAsuransi As String
    strNoCm As String
    strNamaPeserta As String
    strIdPeserta As String

    strKdKelasDitanggung As String
    strKdGolongan As String
    dTglLahir As Date
    strAlamat As String
    strNoPendaftaran As String
    strHubungan As String

    strNoSJP As String
    dTglSJP As Date
    strNoBp As String
    intNoKunjungan As Integer

    strStatusNoSJP As String
    intAnakKe As Integer
    strUnitBagian As String
    strKdPaket As String

    strNoRujukan As String
    strKdRujukanAsal As String
    strDetailRujukanAsal As String
    strKdDetailRujukanAsal As String
    strNamaPerujuk As String

    dTglDirujuk As String
    strDiagnosaRujukan As String
    strKdDiagnosa As String

    blnSuksesAsuransi As Boolean
    strKdKelompokPasien As String

    strPerusahaanPenjamin As String
End Type

Public typAsuransi As Asuransi
Public mstrFormPengirim As String
Public mstrNoKirim As String
Public mdAwal As Date
Public mdAkhir As Date
Public mSup As String
Public mblnFormDaftarPasienRJ As Boolean
Public DbRec As New ADODB.recordset

Public strKdKelompokPasien As String
Public mstrNoSJP As String

Public mintJmlBarisGrafik As Integer ' number of rows needed in Chart data
Public mintJmlKolomGrafik As Integer ' number of colomns needed in Chart data
Public arrGrafik() ' Chart data
Public JnsKriteria() ' criteria
Public mstrGrafik As String

Public mstrKdJenisTarif As String
Public mstrJenisTarif As String
Public intJmlPelayanan As Integer

Type TypePembayaranNonPaket
    curHarga As Currency
    curTanggunganPjm As Currency
    curCostSharing As Currency
    curTanggunganRS As Currency
    curHrsDibyrPsn As Currency
End Type

Type TypePembayaranUmum
    curHrsDibyrPsn As Currency
End Type

Type TypePembayaranPaket
    curHarga As Currency
    curTanggunganPjm As Currency
    curCostSharing As Currency
    curTanggunganRS As Currency
    curHrsDibyrPsn As Currency
    strKdPaket As String
    strKdJnsPelayanan As String
End Type

Public mintJmlPktPlyn As Integer
Public TypByrPkt() As TypePembayaranPaket
Public TypByrPktTotal As TypePembayaranNonPaket
Public TypByrNonPkt As TypePembayaranNonPaket
Public TypByrUmum As TypePembayaranUmum
Public TypByrAll As TypePembayaranNonPaket
Public TypByrOA As TypePembayaranNonPaket
Public TypByrTM As TypePembayaranNonPaket

Public mstrKriteriaLaporan As String
Public intRowNow As Integer
Public mstrNoValidasi As String
Public VFilter As String

Public mstrKdDiagnosa As String
Public mstrKdJenisDiagnosaTindakan As String
Public bolEditDiagnosa As Boolean
Public mdtglclosing As Date
Public strNamaDokter As String
Public strKriteria As String
Public subTampil As Boolean
Public mblnStatusCetakRD As Boolean
Public iJml1 As Integer
Public iJml2 As Integer
Public iJml3 As Integer
Public cJml4 As Currency
Public cJml5 As Currency
Public cJml6 As Currency
Public cJml7 As Currency
Public cJml8 As Currency
Public cJml9 As Currency

Public iJmlRawatJalan As Integer
Public iJmlUGD As Integer
Public iJmlRawatInap As Integer
Public iJmlResep As Integer
Public iJmlTotal As Integer
Public iJmlResep2 As Integer
Public iJmlTotal2 As Integer
Public iJml8RJ As Integer
Public iJmlUGD2 As Integer
Public iJml10 As Integer
Public iJml11 As Integer
Public iJml12 As Integer
Public iJml13 As Integer
Public iJml14 As Integer
Public iJml15 As Integer
Public iJml16 As Integer
Public iJml17 As Integer
Public iJml18 As Integer
Public iJml19 As Integer
Public iJml20 As Integer
Public iJml21 As Integer

Public varStatusBulanan As Boolean
Public varTglCetak As Date
Public varTahun As Integer
Public varBulan As Integer
Public varHari As Integer
Public dTglHitung As Date
Public varNoCetak As String

Public subKdParamedis As String
Public strSQLinst As String
Public rsint As New ADODB.recordset
Public bolStatusFIFO As Boolean
Public bolStatusFrmAntrian As Boolean

Public rsPropinsi As New ADODB.recordset
Public rsKota As New ADODB.recordset
Public rsKecamatan As New ADODB.recordset
Public rsKelurahan As New ADODB.recordset

Public strnonmedis As Boolean
Public MstrKdRuanganAntrian As String
Public MstrKdIstalasiAntrian As String
Public MstrJenisKelasAntrian As String
Public MstrKelasAntrian As String
Public bolPasienReservasi As Boolean
Public bolDetailPasien As Boolean
Public strnoAntrianPasien As String

Public strStatuspasienRS As String
Public strStatuspasienRuangan As String
Public strStatusJenisKelamin As String
Public strStatusKondisiPulang As String
Public strstatusRuanganPerujuk As String

Public strKdRuanganRekamMedis As String
Public strKdRuanganRegistrasiRJ As String
Public strKdRuanganRegistrasiRI As String


Public strBanyakNoCM As String '-- tambahan untuk variabel jumlah panjang NoCm dari settingglobal by Richoe 02/05/2014

'Vclaim Ales
Public mstrNoSJPNew() As String

Public strKdInstalasiRJ As String
Public strKdInstalasiRI As String
Public strKdInstalasiIGD As String
Public mstrJenisPesertaBpjs As String

