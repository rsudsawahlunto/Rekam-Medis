Attribute VB_Name = "modFungsiDB"
'Konversi dari Function : formatNomor
Public Function f_FormatNomor(fNomor As String, fPanjang As Integer) As String
    Dim ffNomor As String
    Dim fi As Integer

    fi = 1
    ffNomor = "0"
    While (fi < fPanjang)
        ffNomor = ffNomor + "0"
        fi = fi + 1
    Wend
    f_FormatNomor = (Left(ffNomor, (Len(ffNomor) - Len(fNomor))) + fNomor)
End Function

'Konversi dari SP: Update_JenisPasienJoinProgramAskes
Public Function f_UpdateJenisPasienJoinProgramAskes(fIdPenjamin As String, fIdAsuransi As String, fNoCM As String, fNamaPeserta As String, fIDPeserta As Variant, fKdGolongan As String, fTglLahir As Date, fAlamat As Variant, fNoPendaftaran As String, fHubungan As String, fNoSJP As String, fTglSJP As Date, fIdUser As String, fNoBP As String, fKunjunganKe As Integer, fStatusNoSJP As String, fAnakKe As Integer, fUnitBagian As Variant, fKdPaket As Variant, fNoRujukan As String, fKdRujukanAsal As String, fDetailRujukanAsal As String, fKdDetailRujukanAsal As String, fNamaPerujuk As Variant, fTglDirujuk As Date, fDiagnosaRujukan As Variant, fKdDiagnosa As Variant, fKdKelompokPasien As String) As String
    'Allow Null: fIDPeserta,fAlamat,fNoSJP,fNoBP,fUnitBagian,fKdPaket,fDiagnosaRujukan,fKdDiagnosa
    'fStatusNoSJP: O=Otomatis; M=Manual
    'fKdDetailRujukanAsal: Jika fStatusNoSJP='O' Not Null
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String
    Dim fintTemp As Integer
    Dim fTempNoSJP As String
    Dim fbln As String
    Dim fthn As String
    Dim fKdRS As String
    Dim fx As Double
    Dim fSingkatanHub As String
    Dim fJK As String
    Dim fKdRuanganAskes As String
    Dim fKdDetailRujukanAsalAskes As String
    Dim fNoSJPTemp As String
    Dim fTglMasuk As Date
    Dim fKdKelas As String
    Dim fKdJenisTarif As String
    Dim fKdSubInstalasi As String
    Dim fKdRuangan As String
    Dim fTempNoSJPNonAskes As String
    Dim fi As Double
    Dim fhr As String

    fthn = Right(Year(fTglSJP), 2)
    fbln = f_FormatNomor(Month(fTglSJP), 2)
    fhr = f_FormatNomor(Day(fTglSJP), 2)
    Set fRS = Nothing
    fQuery = "select TglPendaftaran,KdKelasAkhir,KdRuanganAkhir  from PasienDaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTglMasuk = IIf(IsNull(fRS("TglPendaftaran").value), "", fRS("TglPendaftaran").value) Else fTglMasuk = ""
    If fRS.EOF = False Then fKdKelas = IIf(IsNull(fRS("KdKelasAkhir").value), "", fRS("KdKelasAkhir").value) Else fKdKelas = ""
    If fRS.EOF = False Then fKdRuangan = IIf(IsNull(fRS("KdRuanganAkhir").value), "", fRS("KdRuanganAkhir").value) Else fKdRuangan = ""
    Set fRS = Nothing
    fQuery = "select KdJenisTarif from KelompokPasien where KdKelompokPasien='" & fKdKelompokPasien & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisTarif = IIf(IsNull(fRS("KdJenisTarif").value), "01", fRS("KdJenisTarif").value) Else fKdJenisTarif = "01"
    'cek apakah pasien sudah ada asuransi
    Set fRS = Nothing
    fQuery = "select count(*) as JmlCount from AsuransiPasien where IdPenjamin='" & fIdPenjamin & "' and IdAsuransi='" & fIdAsuransi & "' and NoCM='" & fNoCM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fintTemp = IIf(IsNull(fRS("JmlCount").value), 0, fRS("JmlCount").value) Else fintTemp = 0
    If fintTemp = 0 Then
        Set fRS = Nothing
        fQuery = "insert into AsuransiPasien values('" & fIdPenjamin & "','" & fIdAsuransi & "','" & fNoCM & "','" & fNamaPeserta & "'," & IIf(fIDPeserta = "", "null", "'" & fIDPeserta & "'") & ",'" & fKdGolongan & "','" & Format(fTglLahir, "yyyy/MM/dd") & "'," & IIf(fAlamat = "", "null", "'" & fAlamat & "'") & ")"
        Call msubRecFO(fRS, fQuery)
    Else
        Set fRS = Nothing
        fQuery = "update AsuransiPasien set NamaPeserta='" & fNamaPeserta & "',IDPeserta=" & IIf(fIDPeserta = "", "null", "'" & fIDPeserta & "'") & ",KdGolongan='" & fKdGolongan & "',TglLahir='" & Format(fTglLahir, "yyyy/MM/dd") & "',Alamat=" & IIf(fAlamat = "", "null", "'" & fAlamat & "'") & " where IdPenjamin='" & fIdPenjamin & "' and IdAsuransi='" & fIdAsuransi & "' and NoCM='" & fNoCM & "'"
        Call msubRecFO(fRS, fQuery)
    End If
    'aktifkan kode ini jika join dengan askes
    '   If UCase(fStatusNoSJP) = "O" Then
    '        fKdRS = "1301R002"
    '        Set fRS = Nothing
    '        fQuery = "select JenisKelamin from Pasien where NoCM='" & fNoCM & "'"
    '        Call msubRecFO(fRS, fQuery)
    '        If fRS.EOF = False Then fJK = IIf(IsNull(fRS("JenisKelamin").Value), "", fRS("JenisKelamin").Value) Else fJK = ""
    '        Set fRS = Nothing
    '        fQuery = "select Singkatan from HubunganPesertaAsuransi where Hubungan='" & fHubungan & "'"
    '        Call msubRecFO(fRS, fQuery)
    '        If fRS.EOF = False Then fSingkatanHub = IIf(IsNull(fRS("Singkatan").Value), "", fRS("Singkatan").Value) Else fSingkatanHub = ""
    '        If UCase(fSingkatanHub) = "A" Then
    '            fSingkatanHub = CStr(fAnakKe)
    '        Else
    '            fSingkatanHub = fSingkatanHub
    '        End If
    '        Set fRS = Nothing
    '        fQuery = "select KdRuanganAskes from ConvertRuanganToAskes where KdRuangan='" & fKdRuangan & "'"
    '        Call msubRecFO(fRS, fQuery)
    '        If fRS.EOF = False Then fKdRuanganAskes = IIf(IsNull(fRS("KdRuanganAskes").Value), "", fRS("KdRuanganAskes").Value) Else fKdRuanganAskes = ""
    '        Set fRS = Nothing
    '        fQuery = "select KdDetailRujukanAsalAskes from ConvertDetailRujukanAsalToAskes where KdDetailRujukanAsal='" & fKdDetailRujukanAsal & "'"
    '        Call msubRecFO(fRS, fQuery)
    '        If fRS.EOF = False Then fKdDetailRujukanAsalAskes = IIf(IsNull(fRS("KdDetailRujukanAsalAskes").Value), "", fRS("KdDetailRujukanAsalAskes").Value) Else fKdDetailRujukanAsalAskes = ""
    '        If fKdKelompokPasien = "02" Or fKdKelompokPasien = "10" Or fKdKelompokPasien = "11" Then
    '            Set fRS = Nothing
    '            fQuery = "select max(cast(right(ASKESRS.dbo.DatSJP.NoSJP,6) as bigint)) as NoSJPMax from ASKESRS.dbo.DatSJP where (left(ASKESRS.dbo.DatSJP.NoSJP,8)='" & fKdRS & "') and (substring(ASKESRS.dbo.DatSJP.NoSJP,9,2)='" & fbln & "') and (substring(ASKESRS.dbo.DatSJP.NoSJP,11,2)='" & fthn & "') and (substring(ASKESRS.dbo.DatSJP.NoSJP,13,1)='Y')"
    '            Call msubRecFO(fRS, fQuery)
    '            If fRS.EOF = False Then fTempNoSJP = CStr(IIf(IsNull(fRS("NoSJPMax").Value), 0, fRS("NoSJPMax").Value)) Else fTempNoSJP = CStr(0)
    '            If fTempNoSJP = "0" Then
    '                fTempNoSJP = fKdRS + fbln + fthn + "Y" + "000001"
    '            Else
    '                fx = CDbl(Right(fTempNoSJP, 6)) + 1
    '                fTempNoSJP = fKdRS + fbln + fthn + "Y" + f_FormatNomor(CStr(fx), 6)
    '            End If
    '            fNoSJP = fTempNoSJP
    '            'insert ke DB Askes
    '            Set fRS = Nothing
    '            fQuery = "select ASKESRS.dbo.DatSJP.NoSJP from AskesRS.dbo.DatSJP where ASKESRS.dbo.DatSJP.NoSJP=" & fNoSJP & ""
    '            Call msubRecFO(fRS, fQuery)
    '            If fRS.EOF = True Then
    '                Set fRS = Nothing
    '                fQuery = "insert into AskesRS.dbo.DatPeserta values(null,'" & fIdAsuransi & "','" & fSingkatanHub & "','" & fNamaPeserta & "','" & Format(fTglLahir, "yyyy/MM/dd") & "','" & fJK & "',null,null,null,null)"
    '                Call msubRecFO(fRS, fQuery)
    '                Set fRS = Nothing
    '                fQuery = "insert into AskesRS.dbo.DatMR values('" & fKdDetailRujukanAsalAskes & "','" & fIdAsuransi & "','" & fSingkatanHub & "','" & fNoCM & "')"
    '                Call msubRecFO(fRS, fQuery)
    '                Set fRS = Nothing
    '                fQuery = "insert into AskesRS.dbo.DatSJP values(null," & fNoSJP & ",'" & Format(fTglSJP, "yyyy/MM/dd") & "',null,null,'" & fIdAsuransi & "','" & fSingkatanHub & "',null,'" & fNoCM & "','" & Format(fTglMasuk, "yyyy/MM/dd") & "',null,null,null,null," & fKdDiagnosa & ",'" & fKdRuanganAskes & "',null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null)"
    '                Call msubRecFO(fRS, fQuery)
    '            End If
    '        End If
    '        If fKdKelompokPasien = "05" Then
    '            Set fRS = Nothing
    '            fQuery = "select max(cast(right(ASKESRSGakin.dbo.DatSJP.NoSJP,6) as bigint)) as NoSJPMax from ASKESRSGakin.dbo.DatSJP where (left(ASKESRSGakin.dbo.DatSJP.NoSJP,8)='" & fKdRS & "') and (substring(ASKESRSGakin.dbo.DatSJP.NoSJP,9,2)='" & fbln & "') and (substring(ASKESRSGakin.dbo.DatSJP.NoSJP,11,2)='" & fthn & "') and (substring(ASKESRSGakin.dbo.DatSJP.NoSJP,13,1)='Y')"
    '            Call msubRecFO(fRS, fQuery)
    '            If fRS.EOF = False Then fTempNoSJP = CStr(IIf(IsNull(fRS("NoSJPMax").Value), 0, fRS("NoSJPMax").Value)) Else fTempNoSJP = CStr(0)
    '            If fTempNoSJP = "0" Then
    '                fTempNoSJP = fKdRS + fbln + fthn + "Y" + "000001"
    '            Else
    '                fx = CDbl(Right(fTempNoSJP, 6)) + 1
    '                fTempNoSJP = fKdRS + fbln + fthn + "Y" + f_FormatNomor(CStr(fx), 6)
    '            End If
    '            fNoSJP = fTempNoSJP
    '            'insert ke DB Askes
    '            Set fRS = Nothing
    '            fQuery = "select ASKESRSGakin.dbo.DatSJP.NoSJP from ASKESRSGakin.dbo.DatSJP where ASKESRSGakin.dbo.DatSJP.NoSJP=" & fNoSJP & ""
    '            Call msubRecFO(fRS, fQuery)
    '            If fRS.EOF = True Then
    '                Set fRS = Nothing
    '                fQuery = "insert into ASKESRSGakin.dbo.DatPeserta values(null,'" & fIdAsuransi & "','" & fSingkatanHub & "','" & fNamaPeserta & "','" & Format(fTglLahir, "yyyy/MM/dd") & "','" & fJK & "',null,null,null,null)"
    '                Call msubRecFO(fRS, fQuery)
    '                Set fRS = Nothing
    '                fQuery = "insert into ASKESRSGakin.dbo.DatMR values('" & fKdDetailRujukanAsalAskes & "','" & fIdAsuransi & "','" & fSingkatanHub & "','" & fNoCM & "')"
    '                Call msubRecFO(fRS, fQuery)
    '                Set fRS = Nothing
    '                fQuery = "insert into ASKESRSGakin.dbo.DatSJP values(null," & fNoSJP & ",'" & Format(fTglSJP, "yyyy/MM/dd") & "',null,null,'" & fIdAsuransi & "','" & fSingkatanHub & "',null,'" & fNoCM & "','" & Format(fTglMasuk, "yyyy/MM/dd HH:mm:ss") & "',null,null,null,null," & fKdDiagnosa & ",'" & fKdRuanganAskes & "',null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null)"
    '                Call msubRecFO(fRS, fQuery)
    '            End If
    '        End If
    '        If fKdPaket <> "" Then
    '            Set fRS = Nothing
    '            fQuery = "insert into PelayananSJP values('" & fNoPendaftaran & "'," & fNoSJP & "," & fKdPaket & ")"
    '            Call msubRecFO(fRS, fQuery)
    '        End If
    '   Else
    If fNoSJP = "" Then
        Set fRS = Nothing
        fQuery = "select max(cast(right(NoSJP,4) as integer)) as NoSJPMax from PemakaianAsuransi where (left(NoSJP,2)='" & fKdKelompokPasien & "') and (substring(NoSJP,3,2)='" & fhr & "') and (substring(NoSJP,5,2)='" & fbln & "') and (substring(NoSJP,7,2)='" & fthn & "')"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fTempNoSJPNonAskes = CStr(IIf(IsNull(fRS("NoSJPMax").value), 0, fRS("NoSJPMax").value)) Else fTempNoSJPNonAskes = CStr(0)
        If fTempNoSJPNonAskes = "0" Then
            fTempNoSJPNonAskes = fKdKelompokPasien + fhr + fbln + fthn + "0001"
        Else
            fi = CInt(Right(fTempNoSJPNonAskes, 4)) + 1
            fTempNoSJPNonAskes = fKdKelompokPasien + fhr + fbln + fthn + f_FormatNomor(CStr(fi), 4)
        End If
        fNoSJP = fTempNoSJPNonAskes
    End If
    '   End If
    'cek apakah pasien sudah pakai asuransi
    Set fRS = Nothing
    fQuery = "select count(*)  as JmlCount from PemakaianAsuransi where NoPendaftaran=" & fNoPendaftaran & ""
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fintTemp = IIf(IsNull(fRS("JmlCount").value), 0, fRS("JmlCount").value) Else fintTemp = 0
    If fintTemp = 0 Then
        Set fRS = Nothing
        fQuery = "insert into PemakaianAsuransi values('" & fIdPenjamin & "','" & fIdAsuransi & "','" & fNoCM & "','" & fNoPendaftaran & "','" & fHubungan & "'," & fNoSJP & ",'" & Format(fTglSJP, "yyyy/MM/dd HH:mm:ss") & "'," & fNoBP & ",'" & fKunjunganKe & "'," & fUnitBagian & "," & fAnakKe & ")"
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "insert into Rujukan values('" & fNoPendaftaran & "','" & fNoCM & "','" & fNoRujukan & "','" & fKdRujukanAsal & "','" & fDetailRujukanAsal & "','" & fNamaPerujuk & "','" & Format(fTglDirujuk, "yyyy/MM/dd HH:mm:ss") & "'," & fDiagnosaRujukan & ")"
        Call msubRecFO(fRS, fQuery)
    Else
        Set fRS = Nothing
        fQuery = "update PemakaianAsuransi set IdPenjamin='" & fIdPenjamin & "',IdAsuransi='" & fIdAsuransi & "',NoCM='" & fNoCM & "',KdHubungan='" & fHubungan & "',TglSJP='" & Format(fTglSJP, "yyyy/MM/dd HH:mm:ss") & "',NoBP=" & fNoBP & ",KunjunganKe=" & fKunjunganKe & ",UnitBagian=" & fUnitBagian & ",AnakKe=" & fAnakKe & " where NoPendaftaran='" & fNoPendaftaran & "'  and NoSJP=" & fNoSJP & ""
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "update Rujukan set NoRujukan='" & fNoRujukan & "',KdRujukanAsal='" & fKdRujukanAsal & "',SubRujukanAsal='" & fDetailRujukanAsal & "',NamaPerujuk='" & fNamaPerujuk & "',DiagnosaRujukan=" & fDiagnosaRujukan & "  where NoPendaftaran='" & fNoPendaftaran & "' and TglDirujuk='" & Format(fTglDirujuk, "yyyy/MM/dd HH:mm:ss") & "'"
        Call msubRecFO(fRS, fQuery)
    End If
    'update kelompok pasien di pasien daftar
    Set fRS = Nothing
    fQuery = "update PasienDaftar set KdKelompokPasien='" & fKdKelompokPasien & "' where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    'update jenis tarif pasien di biaya pelayanan
    Set fRS = Nothing
    fQuery = "update BiayaPelayanan set KdJenisTarif='" & fKdJenisTarif & "' where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    'execute reports Jenis Pasien Lama
    Call f_UpdateReportsOAOnUbahJenisPasienLama(fNoPendaftaran)
    Call f_UpdateReportsTMOnUbahJenisPasienLama(fNoPendaftaran)
    'execute reports Jenis Pasien Baru
    Call f_UpdateBiayaPelayananOnUbahJenisPasien(fNoPendaftaran)
    'debug
    Call f_AddDetailBiayaPelayananOnUbahJenisPasien(fNoPendaftaran)
    Call f_UpdatePemakaianAlkesOnUbahJenisPasien(fNoPendaftaran)
    Call f_AddDetailPemakaianObatAlkesOnUbahJenisPasien(fNoPendaftaran)
    f_UpdateJenisPasienJoinProgramAskes = fNoSJP
End Function

'Konversi dari SP: Add_TempHargaKomponen
Public Function f_AddTempHargaKomponen(fNoPendaftaran As String, fKdRuangan As String, fTglPelayanan As Date, fKdPelayananRS As String, fKdKelas As String, fKdJenisTarif As String, fTarifCito As Double, fJmlPelayanan As Integer, fStatusCito As String, fIdPegawai As String)
    'Public Function f_AddTempHargaKomponen(fNoPendaftaran As String, fKdRuangan As String, fTglPelayanan As Date, fKdPelayananRS As String, fKdKelas As String, fKdJenisTarif As String, fTarifCito As Double, fJmlPelayanan As Integer, fStatusCito As String, fIdDokter As String)
    Dim fKdKomponen As String
    Dim fHarga As Currency
    Dim fTotalTarif As Currency
    Dim fKdKomponenTarifTotal As String
    Dim fKdKomponenTarifCito As String
    Dim fTarifTotal As Currency
    Dim fIdDokter As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fIdPegawai1 As String
    Dim fIdPegawai2 As Variant
    Dim fIdPegawai3 As Variant
    Dim fKdJenisPegawai1 As String
    Dim fKdJenisPegawai2 As String
    Dim fKdJenisPegawai3 As String
    Dim fJmlPembebasanPerKomp As Currency
    Dim fJmlHutangPerKomp As Currency
    Dim fJmlTanggunganPerKomp As Currency
    Dim fTarifKelasPenjaminDB As Currency
    Dim fJmlHutangPenjaminDB As Currency
    Dim fJmlTanggunganRSDB As Currency
    Dim fJmlPembebasanDB As Currency
    Dim fTotalTarifPenjamin As Currency
    Dim fKdRuanganAsal As String
    Dim fNoLab_Rad As Variant

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select IdPegawai,IdPegawai2,IdPegawai3,TarifKelasPenjamin,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,NoLab_Rad from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPegawai1 = fRS("IdPegawai").value
        fIdPegawai2 = fRS("IdPegawai2").value
        fIdPegawai3 = fRS("IdPegawai3").value
        fTarifKelasPenjaminDB = IIf(IsNull(fRS("TarifKelasPenjamin").value), 0, fRS("TarifKelasPenjamin").value)
        fJmlHutangPenjaminDB = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRSDB = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanDB = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fNoLab_Rad = fRS("NoLab_Rad").value

    End If

    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "','" & fNoLab_Rad & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = fRS("KdRuanganAsal").value
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai1 & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisPegawai1 = fRS("KdJenisPegawai").value Else fKdJenisPegawai1 = ""
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai2 & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisPegawai2 = fRS("KdJenisPegawai").value Else fKdJenisPegawai2 = ""
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai3 & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisPegawai3 = fRS("KdJenisPegawai").value Else fKdJenisPegawai3 = ""
    fTotalTarifPenjamin = fTarifKelasPenjaminDB + fTarifCito
    Set fRS = Nothing
    fQuery = "select KdDetailJenisJasaPelayanan from PasienDaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdDetailJenisJasaPelayanan = fRS("KdDetailJenisJasaPelayanan").value Else fKdDetailJenisJasaPelayanan = ""
    If fKdJenisPegawai1 = "001" Then
        fIdDokter = fIdPegawai
    Else
        fIdDokter = ""
    End If
    Set fRS = Nothing
    fQuery = "select KdPelayananRS from ConvertPelayananToJasaDokter where KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdPelayananRS='" & fKdPelayananRS & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "'"
    Else
        If (fIdDokter = "") Then
            fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen not in ('02','04','14')"
        End If
        If (fIdPegawai2 = "") And (fIdPegawai3 = "") And (fIdDokter <> "") Then
            fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen not in ('04','14')"
        End If
        If (fIdPegawai2 <> "") And (fIdPegawai3 = "") And (fIdDokter <> "") Then
            fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen<>'14'"
        End If
        If (fIdPegawai2 <> "") And (fIdPegawai3 <> "") And (fIdDokter <> "") Then
            fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "'"
        End If
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)

    While fRS.EOF = False
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        Set fRS2 = Nothing
        fQuery2 = "select dbo.FB_NewTakeTarifBPTMK('" & fNoPendaftaran & "', '" & fKdPelayananRS & "', '" & fKdKelas & "', '" & fKdJenisTarif & "', '" & fKdKomponen & "') as Harga"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fHarga = fRS2("Harga").value Else fHarga = 0
        fJmlPembebasanPerKomp = 0
        If fTarifKelasPenjaminDB = 0 Then
            fJmlHutangPerKomp = 0
            fJmlTanggunganPerKomp = 0
        Else
            fJmlHutangPerKomp = (CDec(fHarga) / CDec(fTotalTarifPenjamin)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKomp = (CDec(fHarga) / CDec(fTotalTarifPenjamin)) * CDec(fJmlTanggunganRSDB)
        End If
        Set fRS2 = Nothing
        fQuery2 = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then
            If fKdKomponen <> "04" And fKdKomponen <> "14" Then
                fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "'," & fHarga & "," & fJmlPelayanan & ", null," & IIf(fIdPegawai1 = "", "null", "'" & fIdPegawai1 & "'") & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKomp)) & ",null)"
            End If
            If fKdKomponen = "04" Then
                fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "'," & fHarga & "," & fJmlPelayanan & ", null," & IIf(fIdPegawai2 = "", "null", "'" & fIdPegawai2 & "'") & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKomp)) & ",null)"
            End If
            If fKdKomponen = "14" Then
                fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "'," & fHarga & "," & fJmlPelayanan & ", null," & IIf(fIdPegawai3 = "", "null", "'" & fIdPegawai3 & "'") & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKomp)) & ",null)"
            End If
        Else
            If fKdKomponen <> "04" And fKdKomponen <> "14" Then
                fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fHarga & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
            End If
            If fKdKomponen = "04" Then
                fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fHarga & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai=" & fIdPegawai2 & ",JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
            End If
            If fKdKomponen = "14" Then
                fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fHarga & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai=" & fIdPegawai3 & ",JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
            End If
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
        Call f_AMDataPelayananTMPasienPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, "A")
        If fKdJenisPegawai1 = "001" And fKdKomponen <> "04" And fKdKomponen <> "14" And fKdKomponen <> "01" Then
            Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, fIdPegawai1, "A")
        End If
        If fKdJenisPegawai2 = "001" And fKdKomponen = "04" Then
            Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, fIdPegawai2, "A")
        End If
        If fKdJenisPegawai3 = "001" And fKdKomponen = "14" Then
            Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, fIdPegawai3, "A")
        End If
        fRS.MoveNext
    Wend

    '--begin Tarif Total
    Set fRS = Nothing
    fQuery = "select KdKomponenTarifTotalTM from MasterDataPendukung"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then fKdKomponenTarifTotal = "12" Else fKdKomponenTarifTotal = fRS("KdKomponenTarifTotalTM").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTM('" & fNoPendaftaran & "', '" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdJenisTarif & "','" & fStatusCito & "','" & fIdPegawai1 & "','" & fIdPegawai2 & "','" & fIdPegawai3 & "', 'T') as Harga"
    Call msubRecFO(fRS, fQuery)
    fTarifTotal = fRS("Harga").value
    Set fRS = Nothing
    fQuery = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifTotal & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponenTarifTotal & "','" & fKdJenisTarif & "'," & fTarifTotal & "," & fJmlPelayanan & ", null," & IIf(fIdPegawai1 = "", "null", "'" & fIdPegawai1 & "'") & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminDB)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSDB)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanDB)) & ",null)"
    Else
        fQuery = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & msubKonversiKomaTitik(CStr(fTarifTotal)) & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai=" & IIf(fIdPegawai1 = "", "null", "'" & fIdPegawai1 & "'") & ",JmlHutangPenjamin=" & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminDB)) & ",JmlTanggunganRS=" & msubKonversiKomaTitik(CStr(fJmlTanggunganRSDB)) & ",JmlPembebasan=" & msubKonversiKomaTitik(CStr(fJmlPembebasanDB)) & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifTotal & "' and NoStruk is null"
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
    Call f_AMDataPelayananTMPasienPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponenTarifTotal, fTarifTotal, fJmlHutangPenjaminDB, fJmlTanggunganRSDB, fJmlPembebasanDB, fKdKelas, "A")
    'end Tarif Total

    'begin Tarif Cito
    If fStatusCito = "1" Then
        Set fRS = Nothing
        fQuery = "select KdKomponenTarifCito from MasterDataPendukung"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then fKdKomponenTarifCito = "07" Else fKdKomponenTarifCito = fRS("KdKomponenTarifCito").value
        fJmlPembebasanPerKomp = 0
        If fTarifKelasPenjaminDB = 0 Then
            fJmlHutangPerKomp = 0
            fJmlTanggunganPerKomp = 0
        Else
            fJmlHutangPerKomp = (CDec(fTarifCito) / CDec(fTotalTarifPenjamin)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKomp = (CDec(fTarifCito) / CDec(fTotalTarifPenjamin)) * CDec(fJmlTanggunganRSDB)
        End If
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifCito & "' and NoStruk is null"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            fQuery = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponenTarifCito & "','" & fKdJenisTarif & "'," & fTarifCito & "," & fJmlPelayanan & ", null," & fIdPegawai1 & "," & fJmlHutangPerKomp & "," & fJmlTanggunganPerKomp & "," & fJmlPembebasanPerKomp & ",null)"
        Else
            fQuery = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fTarifCito & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifCito & "' and NoStruk is null"
        End If
        Set fRS = Nothing
        Call msubRecFO(fRS, fQuery)
        Call f_AMDataPelayananTMPasienPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponenTarifCito, CCur(fTarifCito), fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, "A")
        If fKdJenisPegawai1 = "001" Then
            Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponenTarifCito, CCur(fTarifCito), fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, fIdPegawai1, "A")
        End If
    End If
    'end Tarif Cito
End Function

'fungsi ini tidak berlaku untuk RSU Haji
'Konversi dari SP: Add_TempHargaKomponenForIBS
Public Function f_AddTempHargaKomponenForIBS(fNoPendaftaran As String, fKdRuangan As String, fTglPelayanan As Date, fKdPelayananRS As String, fKdKelas As String, fKdJenisTarif As String, fJmlPelayanan As Integer)
    Dim fKdKomponen As String
    Dim fHarga As Currency
    Dim fTotalTarif As Currency
    Dim fKdKomponenTarifTotal As String
    Dim fKdKomponenTarifCito As String
    Dim fTarifTotal As Currency
    Dim fIdDokter As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fIdPegawai1 As String
    Dim fIdPegawai2 As Variant
    Dim fIdPegawai3 As Variant
    Dim fKdJenisPegawai1 As String
    Dim fKdJenisPegawai2 As String
    Dim fKdJenisPegawai3 As String
    Dim fJmlPembebasanPerKomp As Currency
    Dim fJmlHutangPerKomp As Currency
    Dim fJmlTanggunganPerKomp As Currency
    Dim fTarifKelasPenjaminDB As Currency
    Dim fJmlHutangPenjaminDB As Currency
    Dim fJmlTanggunganRSDB As Currency
    Dim fJmlPembebasanDB As Currency
    Dim fTotalTarifPenjamin As Currency
    Dim fTarifCito As Currency
    Dim fKdRuanganAsal As String
    Dim fNoLab_Rad As Variant

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select IdPegawai,IdPegawai2,IdPegawai3,TarifKelasPenjamin,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,TarifCito,NoLab_Rad from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPegawai1 = fRS("IdPegawai").value
        fIdPegawai2 = fRS("IdPegawai2").value
        fIdPegawai3 = fRS("IdPegawai3").value
        fTarifKelasPenjaminDB = IIf(IsNull(fRS("TarifKelasPenjamin").value), 0, fRS("TarifKelasPenjamin").value)
        fJmlHutangPenjaminDB = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRSDB = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanDB = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fTarifCito = IIf(IsNull(fRS("TarifCito").value), 0, fRS("TarifCito").value)
        fNoLab_Rad = fRS("NoLab_Rad").value
    End If
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "','" & fNoLab_Rad & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = fRS("KdRuanganAsal").value
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai1 & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisPegawai1 = fRS("KdJenisPegawai").value Else fKdJenisPegawai1 = ""
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai=" & fIdPegawai2 & ""
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisPegawai2 = fRS("KdJenisPegawai").value Else fKdJenisPegawai2 = ""
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai=" & fIdPegawai3 & ""
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisPegawai3 = fRS("KdJenisPegawai").value Else fKdJenisPegawai3 = ""
    fTotalTarifPenjamin = fTarifKelasPenjaminDB + fTarifCito
    Set fRS = Nothing
    fQuery = "select KdDetailJenisJasaPelayanan from PasienDaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdDetailJenisJasaPelayanan = fRS("KdDetailJenisJasaPelayanan").value Else fKdDetailJenisJasaPelayanan = ""
    Set fRS = Nothing
    If (fIdPegawai1 = "") Then
        fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen not in ('02','04','14','20')"
    End If
    If (fIdPegawai2 = "") And (fIdPegawai3 = "") And (fIdPegawai1 <> "") Then
        fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen not in ('04','14','20')"
    End If
    If (fIdPegawai2 <> "") And (fIdPegawai3 = "") And (fIdPegawai1 <> "") Then
        fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen<>'14'"
    End If
    If (fIdPegawai2 <> "") And (fIdPegawai3 <> "") And (fIdPegawai1 <> "") Then
        fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "'"
    End If
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        Set fRS2 = Nothing
        fQuery2 = "select dbo.FB_NewTakeTarifBPTMK('" & fNoPendaftaran & "', '" & fKdPelayananRS & "', '" & fKdKelas & "', '" & fKdJenisTarif & "', '" & fKdKomponen & "') as Harga"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fHarga = fRS2("Harga").value Else fHarga = 0
        fJmlPembebasanPerKomp = 0
        If fTarifKelasPenjaminDB = 0 Then
            fJmlHutangPerKomp = 0
            fJmlTanggunganPerKomp = 0
        Else
            fJmlHutangPerKomp = (CDec(fHarga) / CDec(fTotalTarifPenjamin)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKomp = (CDec(fHarga) / CDec(fTotalTarifPenjamin)) * CDec(fJmlTanggunganRSDB)
        End If
        If fJmlHutangPerKomp = "" Then fJmlHutangPerKomp = 0
        If fJmlTanggunganPerKomp = "" Then fJmlTanggunganPerKomp = 0
        If fKdKomponen = "04" And fIdPegawai2 = "" Then fKdKomponen = "26"
        Set fRS2 = Nothing
        fQuery2 = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then
            If fKdKomponen <> "04" And fKdKomponen <> "14" Then
                fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "'," & fHarga & "," & fJmlPelayanan & ", null," & fIdPegawai1 & "," & fJmlHutangPerKomp & "," & fJmlTanggunganPerKomp & "," & fJmlPembebasanPerKomp & ",null)"
            End If
            If fKdKomponen = "04" Then
                fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "'," & fHarga & "," & fJmlPelayanan & ", null," & fIdPegawai2 & "," & fJmlHutangPerKomp & "," & fJmlTanggunganPerKomp & "," & fJmlPembebasanPerKomp & ",null)"
            End If
            If fKdKomponen = "14" Then
                fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "'," & fHarga & "," & fJmlPelayanan & ", null," & fIdPegawai3 & "," & fJmlHutangPerKomp & "," & fJmlTanggunganPerKomp & "," & fJmlPembebasanPerKomp & ",null)"
            End If
        Else
            If fKdKomponen <> "04" And fKdKomponen <> "14" Then
                fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fHarga & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
            End If
            If fKdKomponen = "04" Then
                fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fHarga & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai=" & fIdPegawai2 & ",JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
            End If
            If fKdKomponen = "14" Then
                fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fHarga & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai=" & fIdPegawai3 & ",JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
            End If
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
        Call f_AMDataPelayananTMPasienPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, "A")
        If fKdJenisPegawai1 = "001" And fKdKomponen <> "04" And fKdKomponen <> "14" And fKdKomponen <> "01" Then
            Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, fIdPegawai1, "A")
        End If
        If fKdJenisPegawai2 = "001" And fKdKomponen = "04" Then
            Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, fIdPegawai2, "A")
        End If
        If fKdJenisPegawai3 = "001" And fKdKomponen = "14" Then
            Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, fIdPegawai3, "A")
        End If
        fRS.MoveNext
    Wend

    '--begin Tarif Total
    Set fRS = Nothing
    fQuery = "select KdKomponenTarifTotalTM from MasterDataPendukung"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then fKdKomponenTarifTotal = "12" Else fKdKomponenTarifTotal = fRS("KdKomponenTarifTotalTM").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTM('" & fNoPendaftaran & "', '" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdJenisTarif & "','" & fStatusCito & "'," & fIdPegawai1 & "," & fIdPegawai2 & "," & fIdPegawai3 & ", 'T') as Harga"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTarifTotal = fRS("Harga").value Else fTarifTotal = 0
    Set fRS = Nothing
    fQuery = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & fTglPelayanan & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifTotal & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponenTarifTotal & "','" & fKdJenisTarif & "'," & fTarifTotal & "," & fJmlPelayanan & ", null," & fIdPegawai1 & "," & fJmlHutangPenjaminDB & "," & fJmlTanggunganRSDB & "," & fJmlPembebasanDB & ",null)"
    Else
        fQuery = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fTarifTotal & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin=" & fJmlHutangPenjaminDB & ",JmlTanggunganRS=" & fJmlTanggunganRSDB & ",JmlPembebasan=" & fJmlPembebasanDB & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifTotal & "' and NoStruk is null"
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
    Call f_AMDataPelayananTMPasienPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponenTarifTotal, fTarifTotal, fJmlHutangPenjaminDB, fJmlTanggunganRSDB, fJmlPembebasanDB, fKdKelas, "A")
    'end Tarif Total

    'begin Tarif Cito
    If fTarifCito <> 0 Then
        Set fRS = Nothing
        fQuery = "select KdKomponenTarifCito from MasterDataPendukung"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then fKdKomponenTarifCito = "07" Else fKdKomponenTarifCito = fRS("KdKomponenTarifCito").value
        fJmlPembebasanPerKomp = 0
        If fTarifKelasPenjaminDB = 0 Then
            fJmlHutangPerKomp = 0
            fJmlTanggunganPerKomp = 0
        Else
            fJmlHutangPerKomp = (CDec(fTarifCito) / CDec(fTotalTarifPenjamin)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKomp = (CDec(fTarifCito) / CDec(fTotalTarifPenjamin)) * CDec(fJmlTanggunganRSDB)
        End If
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifCito & "' and NoStruk is null"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            fQuery = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponenTarifCito & "','" & fKdJenisTarif & "'," & fTarifCito & "," & fJmlPelayanan & ", null," & fIdPegawai1 & "," & fJmlHutangPerKomp & "," & fJmlTanggunganPerKomp & "," & fJmlPembebasanPerKomp & ",null)"
        Else
            fQuery = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fTarifCito & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifCito & "' and NoStruk is null"
        End If
        Set fRS = Nothing
        Call msubRecFO(fRS, fQuery)
        Call f_AMDataPelayananTMPasienPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponenTarifCito, fTarifCito, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, "A")
        If fKdJenisPegawai1 = "001" Then
            Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponenTarifCito, fTarifCito, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, fIdPegawai1, "A")
        End If
    End If
    'end Tarif Cito
End Function

'Konversi dari SP: Add_TempHargaKomponenForPenunjangM
Public Function f_AddTempHargaKomponenForPenunjangM(fNoPendaftaran As String, fKdRuangan As String, fTglPelayanan As Date, fKdPelayananRS As String, fKdKelas As String, fKdJenisTarif As String, fTarifCito As Currency, fJmlPelayanan As Integer, fStatusCito As String, fKdLaboratory As String)
    Dim fKdKomponen As String
    Dim fHarga As Currency
    Dim fTotalTarif As Currency
    Dim fKdKomponenTarifTotal As String
    Dim fKdKomponenTarifCito As String
    Dim fTarifTotal As Currency
    Dim fIdDokter As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fIdPegawai1 As String
    Dim fIdPegawai2 As Variant
    Dim fIdPegawai3 As Variant
    Dim fKdJenisPegawai1 As String
    Dim fKdJenisPegawai2 As String
    Dim fKdJenisPegawai3 As String
    Dim fJmlPembebasanPerKomp As Currency
    Dim fJmlHutangPerKomp As Currency
    Dim fJmlTanggunganPerKomp As Currency
    Dim fTarifKelasPenjaminDB As Currency
    Dim fJmlHutangPenjaminDB As Currency
    Dim fJmlTanggunganRSDB As Currency
    Dim fJmlPembebasanDB As Currency
    Dim fTotalTarifPenjamin As Currency
    Dim fKdRuanganAsal As String
    Dim fNoLab_Rad As Variant

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select IdPegawai,IdPegawai2,IdPegawai3,TarifKelasPenjamin,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,NoLab_Rad from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPegawai1 = fRS("IdPegawai").value
        fIdPegawai2 = fRS("IdPegawai2").value
        fIdPegawai3 = fRS("IdPegawai3").value
        fTarifKelasPenjaminDB = IIf(IsNull(fRS("TarifKelasPenjamin").value), 0, fRS("TarifKelasPenjamin").value)
        fJmlHutangPenjaminDB = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRSDB = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanDB = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fNoLab_Rad = fRS("NoLab_Rad").value
    End If
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "','" & fNoLab_Rad & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = fRS("KdRuanganAsal").value
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai1 & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisPegawai1 = fRS("KdJenisPegawai").value Else fKdJenisPegawai1 = ""
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai=" & fIdPegawai2 & ""
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisPegawai2 = fRS("KdJenisPegawai").value Else fKdJenisPegawai2 = ""
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai=" & fIdPegawai3 & ""
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisPegawai3 = fRS("KdJenisPegawai").value Else fKdJenisPegawai3 = ""
    fTotalTarifPenjamin = fTarifKelasPenjaminDB + fTarifCito
    Set fRS = Nothing
    fQuery = "select KdDetailJenisJasaPelayanan from PasienDaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdDetailJenisJasaPelayanan = fRS("KdDetailJenisJasaPelayanan").value Else fKdDetailJenisJasaPelayanan = ""
    If fKdJenisPegawai1 = "001" Then
        fIdDokter = fIdPegawai
    Else
        fIdDokter = ""
    End If
    If fKdLaboratory = "" Then
        Set fRS = Nothing
        fQuery = "select KdPelayananRS from ConvertPelayananToJasaDokter where KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdPelayananRS='" & fKdPelayananRS & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "'"
        Else
            If (fIdDokter = "") Then
                fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen not in ('02','04','14')"
            End If
            If (fIdPegawai2 = "") And (fIdPegawai3 = "") And (fIdDokter <> "") Then
                fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen not in ('04','14')"
            End If
            If (fIdPegawai2 <> "") And (fIdPegawai3 = "") And (fIdDokter <> "") Then
                fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen<>'14'"
            End If
            If (fIdPegawai2 <> "") And (fIdPegawai3 <> "") And (fIdDokter <> "") Then
                fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "'"
            End If
        End If
        Set fRS = Nothing
        Call msubRecFO(fRS, fQuery)
        While fRS.EOF = False
            fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
            Set fRS2 = Nothing
            fQuery2 = "select dbo.FB_NewTakeTarifBPTMK('" & fNoPendaftaran & "', '" & fKdPelayananRS & "', '" & fKdKelas & "', '" & fKdJenisTarif & "', '" & fKdKomponen & "') as Harga"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = False Then fHarga = fRS2("Harga").value Else fHarga = 0
            fJmlPembebasanPerKomp = 0
            If fTarifKelasPenjaminDB = 0 Then
                fJmlHutangPerKomp = 0
                fJmlTanggunganPerKomp = 0
            Else
                fJmlHutangPerKomp = (CDec(fHarga) / CDec(fTotalTarifPenjamin)) * CDec(fJmlHutangPenjaminDB)
                fJmlTanggunganPerKomp = (CDec(fHarga) / CDec(fTotalTarifPenjamin)) * CDec(fJmlTanggunganRSDB)
            End If
            Set fRS2 = Nothing
            fQuery2 = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                If fKdKomponen <> "04" And fKdKomponen <> "14" Then
                    fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "'," & fHarga & "," & fJmlPelayanan & ", null," & fIdPegawai1 & "," & fJmlHutangPerKomp & "," & fJmlTanggunganPerKomp & "," & fJmlPembebasanPerKomp & ",null)"
                End If
                If fKdKomponen = "04" Then
                    fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "'," & fHarga & "," & fJmlPelayanan & ", null," & fIdPegawai2 & "," & fJmlHutangPerKomp & "," & fJmlTanggunganPerKomp & "," & fJmlPembebasanPerKomp & ",null)"
                End If
                If fKdKomponen = "14" Then
                    fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "'," & fHarga & "," & fJmlPelayanan & ", null," & fIdPegawai3 & "," & fJmlHutangPerKomp & "," & fJmlTanggunganPerKomp & "," & fJmlPembebasanPerKomp & ",null)"
                End If
            Else
                If fKdKomponen <> "04" And fKdKomponen <> "14" Then
                    fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fHarga & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
                End If
                If fKdKomponen = "04" Then
                    fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fHarga & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai=" & fIdPegawai2 & ",JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
                End If
                If fKdKomponen = "14" Then
                    fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fHarga & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai=" & fIdPegawai3 & ",JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
                End If
            End If
            Set fRS2 = Nothing
            Call msubRecFO(fRS2, fQuery2)
            Call f_AMDataPelayananTMPasienPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, "A")
            If fKdJenisPegawai1 = "001" And fKdKomponen <> "04" And fKdKomponen <> "14" And fKdKomponen <> "01" Then
                Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, fIdPegawai1, "A")
            End If
            If fKdJenisPegawai2 = "001" And fKdKomponen = "04" Then
                Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, fIdPegawai2, "A")
            End If
            If fKdJenisPegawai3 = "001" And fKdKomponen = "14" Then
                Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, fIdPegawai3, "A")
            End If
            fRS.MoveNext
        Wend
    Else
        Set fRS = Nothing
        fQuery = "select dbo.FB_NewTakeTarifBPTM('" & fNoPendaftaran & "', '" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdJenisTarif & "','" & fStatusCito & "'," & fIdPegawai1 & "," & fIdPegawai2 & "," & fIdPegawai3 & ", 'T') as Harga"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fTarifTotal = fRS("Harga").value Else fTarifTotal = 0
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & fTglPelayanan & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='16' and NoStruk is null"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            fQuery = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','16','" & fKdJenisTarif & "'," & fTarifTotal & "," & fJmlPelayanan & ", null," & fIdPegawai1 & "," & fJmlHutangPenjaminDB & "," & fJmlTanggunganRSDB & "," & fJmlPembebasanDB & ",null)"
        Else
            fQuery = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fTarifTotal & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin=" & fJmlHutangPenjaminDB & ",JmlTanggunganRS=" & fJmlTanggunganRSDB & ",JmlPembebasan=" & fJmlPembebasanDB & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='16' and NoStruk is null"
        End If
        Set fRS = Nothing
        Call msubRecFO(fRS, fQuery)
        Call f_AMDataPelayananTMPasienPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, "16", fHarga, fJmlHutangPenjaminDB, fJmlTanggunganRSDB, fJmlPembebasanDB, fKdKelas, "A")
        If fKdJenisPegawai1 = "001" Then
            Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, "16", fHarga, fJmlHutangPenjaminDB, fJmlTanggunganRSDB, fJmlPembebasanDB, fKdKelas, fIdPegawai1, "A")
        End If
    End If
    '--begin Tarif Total
    Set fRS = Nothing
    fQuery = "select KdKomponenTarifTotalTM from MasterDataPendukung"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then fKdKomponenTarifTotal = "12" Else fKdKomponenTarifTotal = fRS("KdKomponenTarifTotalTM").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTM('" & fNoPendaftaran & "', '" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdJenisTarif & "','" & fStatusCito & "'," & fIdPegawai1 & "," & fIdPegawai2 & "," & fIdPegawai3 & ", 'T') as Harga"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTarifTotal = fRS("Harga").value Else fTarifTotal = 0
    Set fRS = Nothing
    fQuery = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & fTglPelayanan & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifTotal & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponenTarifTotal & "','" & fKdJenisTarif & "'," & fTarifTotal & "," & fJmlPelayanan & ", null," & fIdPegawai1 & "," & fJmlHutangPenjaminDB & "," & fJmlTanggunganRSDB & "," & fJmlPembebasanDB & ",null)"
    Else
        fQuery = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fTarifTotal & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin=" & fJmlHutangPenjaminDB & ",JmlTanggunganRS=" & fJmlTanggunganRSDB & ",JmlPembebasan=" & fJmlPembebasanDB & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifTotal & "' and NoStruk is null"
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
    Call f_AMDataPelayananTMPasienPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponenTarifTotal, fTarifTotal, fJmlHutangPenjaminDB, fJmlTanggunganRSDB, fJmlPembebasanDB, fKdKelas, "A")
    'end Tarif Total

    'begin Tarif Cito
    If fStatusCito = "1" Then
        Set fRS = Nothing
        fQuery = "select KdKomponenTarifCito from MasterDataPendukung"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then fKdKomponenTarifCito = "07" Else fKdKomponenTarifCito = fRS("KdKomponenTarifCito").value
        fJmlPembebasanPerKomp = 0
        If fTarifKelasPenjaminDB = 0 Then
            fJmlHutangPerKomp = 0
            fJmlTanggunganPerKomp = 0
        Else
            fJmlHutangPerKomp = (CDec(fTarifCito) / CDec(fTotalTarifPenjamin)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKomp = (CDec(fTarifCito) / CDec(fTotalTarifPenjamin)) * CDec(fJmlTanggunganRSDB)
        End If
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifCito & "' and NoStruk is null"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            fQuery = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponenTarifCito & "','" & fKdJenisTarif & "'," & fTarifCito & "," & fJmlPelayanan & ", null," & fIdPegawai1 & "," & fJmlHutangPerKomp & "," & fJmlTanggunganPerKomp & "," & fJmlPembebasanPerKomp & ",null)"
        Else
            fQuery = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fTarifCito & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifCito & "' and NoStruk is null"
        End If
        Set fRS = Nothing
        Call msubRecFO(fRS, fQuery)
        Call f_AMDataPelayananTMPasienPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponenTarifCito, fTarifCito, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, "A")
        If fKdJenisPegawai1 = "001" Then
            Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponenTarifCito, fTarifCito, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fKdKelas, fIdPegawai1, "A")
        End If
    End If
    'end Tarif Cito
End Function

'Konversi dari SP: AM_DataPelayananApotikPH
Public Function f_AMDataPelayananApotikPH(fNoStruk As String, fTglStruk As Date, fKdRuangan As String, fKdRuanganAsal As String, fKdBarang As String, fKdAsal As String, fSatuanJml As String, fKdKomponen As String, fHarga As Currency, fJmlService As Integer, fJmlBarang As Double, fStatus As String)
    'fStatus=A=Add; M=Min
    Dim fTotalHarga As Currency
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fKdPelayananRS As String
    Dim fJmlHutangPenjaminTotal As Currency
    Dim fJmlTanggunganRSTotal As Currency
    Dim fTotalBiaya As Currency
    Dim fTarifService As Currency
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fDiscount As Currency
    Dim fHargaAkhir As Currency
    Dim fHargaSatuan As Currency

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien from V_StrukPelayananApotik where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value)
        fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
    End If
    Set fRS = Nothing
    fQuery = "select TarifService,JmlHutangPenjamin,JmlTanggunganRS,Discount,HargaSatuan from ApotikJual where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and SatuanJml='" & fSatuanJml & "' and KdAsal='" & fKdAsal & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fTarifService = IIf(IsNull(fRS("TarifService").value), 0, fRS("TarifService").value)
        fJmlHutangPenjamin = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRS = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fDiscount = IIf(IsNull(fRS("Discount").value), 0, fRS("Discount").value)
        fHargaSatuan = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
    End If
    fHargaAkhir = fHargaSatuan - fDiscount
    fTotalHarga = (fTarifService + fHargaAkhir)
    If fKdKomponen = "10" Then
        fTotalBiaya = (fHarga * fJmlService)
        fJmlHutangPenjaminTotal = CDec(fJmlService) * ((CDec(fHarga) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjamin))
        fJmlTanggunganRSTotal = CDec(fJmlService) * ((CDec(fHarga) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRS))
    Else
        fTotalBiaya = (fJmlBarang * fHarga)
        fJmlHutangPenjaminTotal = CDec(fJmlBarang) * ((CDec(fHarga) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjamin))
        fJmlTanggunganRSTotal = CDec(fJmlBarang) * ((CDec(fHarga) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRS))
    End If
    Set fRS = Nothing
    fQuery = "select KdRuangan from DataPelayananApotikPH where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdBarang='" & fKdBarang & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "') and (datepart(hh, TglStruk)=datepart(hh, '" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and day(TglStruk)=day('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and month(TglStruk)=month('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and year(TglStruk)=year('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into DataPelayananApotikPH values('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdBarang & "','" & fKdAsal & "','" & fKdKomponen & "'," & fJmlBarang & "," & fTotalBiaya & "," & fJmlHutangPenjaminTotal & "," & fJmlTanggunganRSTotal & ")"
    Else
        If UCase(fStatus) = "A" Then
            fQuery = "update DataPelayananApotikPH set JmlBarang=JmlBarang+" & fJmlBarang & ",TotalBiaya=TotalBiaya+" & fTotalBiaya & ",TotalHutangPenjamin=TotalHutangPenjamin+" & fJmlHutangPenjaminTotal & ",TotalTanggunganRS=TotalTanggunganRS+" & fJmlTanggunganRSTotal & " where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdBarang='" & fKdBarang & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "') and (datepart(hh, TglStruk)=datepart(hh, '" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and day(TglStruk)=day('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and month(TglStruk)=month('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and year(TglStruk)=year('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery = "update DataPelayananApotikPH set JmlBarang=JmlBarang-" & fJmlBarang & ",TotalBiaya=TotalBiaya-" & fTotalBiaya & ",TotalHutangPenjamin=TotalHutangPenjamin-" & fJmlHutangPenjaminTotal & ",TotalTanggunganRS=TotalTanggunganRS-" & fJmlTanggunganRSTotal & " where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdBarang='" & fKdBarang & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "') and (datepart(hh, TglStruk)=datepart(hh, '" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and day(TglStruk)=day('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and month(TglStruk)=month('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and year(TglStruk)=year('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
        'tambah onede
    End If
    '
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
End Function

'Konversi dari SP: AM_DataPelayananOAPasienPH
Public Function f_AMDataPelayananOAPasienPH(fNoPendaftaran As String, fTglPelayanan As Date, fKdRuangan As String, fKdRuanganAsal As String, fKdBarang As String, fKdAsal As String, fSatuanJml As String, fKdKomponen As String, fHarga As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fJmlService As Integer, fJmlBarang As Double, fStatus As String)
    'fStatus = A:Add; M:Min
    Dim fTotalBiaya As Currency
    Dim fTotalHutangPenjamin As Currency
    Dim fTotalTanggunganRS As Currency
    Dim fTotalPembebasan As Currency
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fKdSubInstalasi As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fKdJenisKelamin As String
    Dim fKdKelas As String
    Dim fKdPelayananRS As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdPelayananRSOA from MasterDataPendukung"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then fKdPelayananRS = "000001" Else fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRSOA").value), "000001", fRS("KdPelayananRSOA").value)
    Set fRS = Nothing
    fQuery = "select KdJenisKelamin from V_JenisKelaminPasienTerdaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisKelamin = IIf(IsNull(fRS("KdJenisKelamin").value), "", fRS("KdJenisKelamin").value) Else fKdJenisKelamin = ""
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien,KdDetailJenisJasaPelayanan from V_JenisPasienNPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value)
        fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
        fKdDetailJenisJasaPelayanan = IIf(IsNull(fRS("KdDetailJenisJasaPelayanan").value), "", fRS("KdDetailJenisJasaPelayanan").value)
    End If
    Set fRS = Nothing
    fQuery = "select KdSubInstalasi,KdKelas from DetailPemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuanJml & "' and KdAsal='" & fKdAsal & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
    End If
    If fKdKomponen = "10" Then
        fTotalBiaya = fJmlService * fHarga
        fTotalHutangPenjamin = fJmlService * fJmlHutangPenjamin
        fTotalTanggunganRS = fJmlService * fJmlTanggunganRS
        fTotalPembebasan = fJmlService * fJmlPembebasan
    Else
        fTotalBiaya = fJmlBarang * fHarga
        fTotalHutangPenjamin = fJmlBarang * fJmlHutangPenjamin
        fTotalTanggunganRS = fJmlBarang * fJmlTanggunganRS
        fTotalPembebasan = fJmlBarang * fJmlPembebasan
    End If
    Set fRS = Nothing
    fQuery = "select KdRuangan from DataPelayananOAPasienPH where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdBarang='" & fKdBarang & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into DataPelayananOAPasienPH values('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdDetailJenisJasaPelayanan & "','" & fKdKelas & "','" & fKdAsal & "','" & fKdBarang & "','" & fKdKomponen & "','" & fKdJenisKelamin & "'," & fJmlBarang & "," & fTotalBiaya & "," & fTotalHutangPenjamin & "," & fTotalTanggunganRS & "," & fTotalPembebasan & ",'" & fKdPelayananRS & "')"
    Else
        If UCase(fStatus) = "A" Then
            fQuery = "update DataPelayananOAPasienPH set JmlBarang=JmlBarang+" & fJmlBarang & ",TotalBiaya=TotalBiaya+" & fTotalBiaya & ",TotalHutangPenjamin=TotalHutangPenjamin+" & fTotalHutangPenjamin & ",TotalTanggunganRS=TotalTanggunganRS+" & fTotalTanggunganRS & ",TotalPembebasan=TotalPembebasan+" & fTotalPembebasan & "" _
            & "where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdBarang='" & fKdBarang & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery = "update DataPelayananOAPasienPH set JmlBarang=JmlBarang-" & fJmlBarang & ",TotalBiaya=TotalBiaya-" & fTotalBiaya & ",TotalHutangPenjamin=TotalHutangPenjamin-" & fTotalHutangPenjamin & ",TotalTanggunganRS=TotalTanggunganRS-" & fTotalTanggunganRS & ",TotalPembebasan=TotalPembebasan-" & fTotalPembebasan & " " _
            & "where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdBarang='" & fKdBarang & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
End Function

'Konversi dari SP: AM_DataPelayananTMPasienPH
Public Function f_AMDataPelayananTMPasienPH(fNoPendaftaran As String, fKdPelayananRS As String, fTglPelayanan As Date, fKdRuangan As String, fKdRuanganAsal As String, fKdKomponen As String, fHarga As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fKdKelas As String, fStatus As String)
    'fStatus= A:Add; M:Min
    Dim fTotalBiaya As Currency
    Dim fTotalHutangPenjamin As Currency
    Dim fTotalTanggunganRS As Currency
    Dim fTotalPembebasan As Currency
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fJmlPelayanan As Integer
    Dim fKdAsal As String
    Dim fKdSubInstalasi As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fKdJenisKelamin As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdJenisKelamin from V_JenisKelaminPasienTerdaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisKelamin = IIf(IsNull(fRS("KdJenisKelamin").value), "", fRS("KdJenisKelamin").value) Else fKdJenisKelamin = ""
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien,KdDetailJenisJasaPelayanan from V_JenisPasienNPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value)
        fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
        fKdDetailJenisJasaPelayanan = IIf(IsNull(fRS("KdDetailJenisJasaPelayanan").value), "", fRS("KdDetailJenisJasaPelayanan").value)
    End If
    Set fRS = Nothing
    fQuery = "select KdSubInstalasi,StatusAPBD,JmlPelayanan from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
        fKdAsal = IIf(IsNull(fRS("StatusAPBD").value), "", fRS("StatusAPBD").value)
        fJmlPelayanan = IIf(IsNull(fRS("JmlPelayanan").value), 0, fRS("JmlPelayanan").value)
    End If
    fTotalBiaya = fJmlPelayanan * fHarga
    fTotalHutangPenjamin = fJmlPelayanan * fJmlHutangPenjamin
    fTotalTanggunganRS = fJmlPelayanan * fJmlTanggunganRS
    fTotalPembebasan = fJmlPelayanan * fJmlPembebasan
    Set fRS = Nothing
    fQuery = "select KdRuangan from DataPelayananTMPasienPH where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into DataPelayananTMPasienPH values('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdDetailJenisJasaPelayanan & "','" & fKdKelas & "','" & fKdAsal & "','" & fKdPelayananRS & "','" & fKdKomponen & "','" & fKdJenisKelamin & "'," & fJmlPelayanan & "," & msubKonversiKomaTitik(CStr(fTotalBiaya)) & "," & msubKonversiKomaTitik(CStr(fTotalHutangPenjamin)) & "," & msubKonversiKomaTitik(CStr(fTotalTanggunganRS)) & "," & msubKonversiKomaTitik(CStr(fTotalPembebasan)) & ")"
    Else
        If UCase(fStatus) = "A" Then
            fQuery = "update DataPelayananTMPasienPH set JmlPelayanan=JmlPelayanan+" & fJmlPelayanan & ",TotalBiaya=TotalBiaya+" & msubKonversiKomaTitik(CStr(fTotalBiaya)) & ",TotalHutangPenjamin=TotalHutangPenjamin+" & msubKonversiKomaTitik(CStr(fTotalHutangPenjamin)) & ",TotalTanggunganRS=TotalTanggunganRS+" & msubKonversiKomaTitik(CStr(fTotalTanggunganRS)) & ",TotalPembebasan=TotalPembebasan+" & msubKonversiKomaTitik(CStr(fTotalPembebasan)) & "" _
            & " where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery = "update DataPelayananTMPasienPH set JmlPelayanan=JmlPelayanan-" & fJmlPelayanan & ",TotalBiaya=TotalBiaya-" & msubKonversiKomaTitik(CStr(fTotalBiaya)) & ",TotalHutangPenjamin=TotalHutangPenjamin-" & msubKonversiKomaTitik(CStr(fTotalHutangPenjamin)) & ",TotalTanggunganRS=TotalTanggunganRS-" & msubKonversiKomaTitik(CStr(fTotalTanggunganRS)) & ",TotalPembebasan=TotalPembebasan-" & msubKonversiKomaTitik(CStr(fTotalPembebasan)) & "" _
            & " where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
End Function

'Konversi dari SP: AM_DataPelayananTMPasienDokterPH
Public Function f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran As String, fKdPelayananRS As String, fTglPelayanan As Date, fKdRuangan As String, fKdRuanganAsal As String, fKdKomponen As String, fHarga As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fKdKelas As String, fIdPegawai As Variant, fStatus As String)
    'fStatus= A:Add; M:Min
    Dim fTotalBiaya As Currency
    Dim fTotalHutangPenjamin As Currency
    Dim fTotalTanggunganRS As Currency
    Dim fTotalPembebasan As Currency
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fJmlPelayanan As Integer
    Dim fKdAsal As String
    Dim fKdSubInstalasi As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fKdJenisKelamin As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdJenisKelamin from V_JenisKelaminPasienTerdaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisKelamin = IIf(IsNull(fRS("KdJenisKelamin").value), "", fRS("KdJenisKelamin").value)
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien,KdDetailJenisJasaPelayanan from V_JenisPasienNPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value)
        fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
        fKdDetailJenisJasaPelayanan = IIf(IsNull(fRS("KdDetailJenisJasaPelayanan").value), "", fRS("KdDetailJenisJasaPelayanan").value)
    End If
    Set fRS = Nothing
    fQuery = "select KdSubInstalasi,StatusAPBD,JmlPelayanan from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
        fKdAsal = IIf(IsNull(fRS("StatusAPBD").value), "", fRS("StatusAPBD").value)
        fJmlPelayanan = IIf(IsNull(fRS("JmlPelayanan").value), 0, fRS("JmlPelayanan").value)
    End If
    fTotalBiaya = fJmlPelayanan * fHarga
    fTotalHutangPenjamin = fJmlPelayanan * fJmlHutangPenjamin
    fTotalTanggunganRS = fJmlPelayanan * fJmlTanggunganRS
    fTotalPembebasan = fJmlPelayanan * fJmlPembebasan
    Set fRS = Nothing
    fQuery = "select KdRuangan from DataPelayananTMPasienDokterPH where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and IdPegawai='" & fIdPegawai & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into DataPelayananTMPasienDokterPH values('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdDetailJenisJasaPelayanan & "','" & fKdKelas & "','" & fKdAsal & "','" & fKdPelayananRS & "','" & fKdKomponen & "'," & IIf(IsNull(fIdPegawai), "null", "'" & fIdPegawai & "'") & ",'" & fKdJenisKelamin & "'," & fJmlPelayanan & "," & msubKonversiKomaTitik(CStr(fTotalBiaya)) & "," & msubKonversiKomaTitik(CStr(fTotalHutangPenjamin)) & "," & msubKonversiKomaTitik(CStr(fTotalTanggunganRS)) & "," & msubKonversiKomaTitik(CStr(fTotalPembebasan)) & ")"
    Else
        If UCase(fStatus) = "A" Then
            'IdPegawai dibuat jadi '" & IdPegawai & "'
            fQuery = "update DataPelayananTMPasienDokterPH set JmlPelayanan=JmlPelayanan+" & fJmlPelayanan & ",TotalBiaya=TotalBiaya+" & msubKonversiKomaTitik(CStr(fTotalBiaya)) & ",TotalHutangPenjamin=TotalHutangPenjamin+" & msubKonversiKomaTitik(CStr(fTotalHutangPenjamin)) & ",TotalTanggunganRS=TotalTanggunganRS+" & msubKonversiKomaTitik(CStr(fTotalTanggunganRS)) & ",TotalPembebasan=TotalPembebasan+" & msubKonversiKomaTitik(CStr(fTotalPembebasan)) & "" _
            & "where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and IdPegawai='" & fIdPegawai & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery = "update DataPelayananTMPasienDokterPH set JmlPelayanan=JmlPelayanan-" & fJmlPelayanan & ",TotalBiaya=TotalBiaya-" & msubKonversiKomaTitik(CStr(fTotalBiaya)) & ",TotalHutangPenjamin=TotalHutangPenjamin-" & msubKonversiKomaTitik(CStr(fTotalHutangPenjamin)) & ",TotalTanggunganRS=TotalTanggunganRS-" & msubKonversiKomaTitik(CStr(fTotalTanggunganRS)) & ",TotalPembebasan=TotalPembebasan-" & msubKonversiKomaTitik(CStr(fTotalPembebasan)) & "" _
            & "where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and IdPegawai='" & fIdPegawai & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
End Function

'Konversi dari SP: Add_TempHargaKomponenOAResep
Public Function f_AddTempHargaKomponenOAResep(fNoPendaftaran As String, fKdRuangan As String, fTglPelayanan As Date, fKdBarang As String, fKdAsal As String, fSatuanJml As String, fHargaSatuan As Currency, fHargaBeli As Currency, fJmlBarang As Double, fKdJenisObat As Variant, fJmlService As Integer, fTarifService As Currency, fNoResep As Variant, fBiayaAdministrasi As Currency, fKdRuanganAsal As String)
    Dim fKdKomponenProfit As String
    Dim fKdKomponenTotal As String
    Dim fKdKomponenHargaNetto As String
    Dim fHargaBersih As Currency
    Dim fKdKomponenTarifService As String
    Dim fKdKomponenAdm As String
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fTarifServiceResep As Currency
    Dim fJasaRS As Currency
    Dim fJmlPembebasanPerKompP As Currency
    Dim fJmlHutangPerKompP As Currency
    Dim fJmlTanggunganPerKompP As Currency
    Dim fJmlPembebasanPerKompHN As Currency
    Dim fJmlHutangPerKompHN As Currency
    Dim fJmlTanggunganPerKompHN As Currency
    Dim fJmlPembebasanPerKompTotal As Currency
    Dim fJmlHutangPerKompTotal As Currency
    Dim fJmlTanggunganPerKompTotal As Currency
    Dim fJmlPembebasanPerKompAdm As Currency
    Dim fJmlHutangPerKompAdm As Currency
    Dim fJmlTanggunganPerKompAdm As Currency
    Dim fJmlPembebasanPerKompService As Currency
    Dim fJmlHutangPerKompService As Currency
    Dim fJmlTanggunganPerKompService As Currency
    Dim fJmlPembebasanPerKompRS As Currency
    Dim fJmlHutangPerKompRS As Currency
    Dim fJmlTanggunganPerKompRS As Currency
    Dim fJmlHutangPenjaminDB As Currency
    Dim fJmlTanggunganRSDB As Currency
    Dim fJmlPembebasanDB As Currency
    Dim fTotalHarga As Currency

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from DetailPemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fJmlHutangPenjaminDB = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRSDB = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanDB = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
    End If
    Set fRS = Nothing
    fQuery = "select KdKelompokPasien,IdPenjamin from V_KelasTanggunganPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value)
        fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
    End If
    fHargaBersih = fHargaSatuan - fHargaBeli
    fTotalHarga = fHargaSatuan + fTarifService + fBiayaAdministrasi
    Set fRS = Nothing
    fQuery = "select KdKomponenTarifTotalOA,KdKomponenProfit,KdKomponenHargaNetto,KdKomponenTarifServisResep,KdKomponenAdm from MasterDataPendukung"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fKdKomponenTotal = IIf(IsNull(fRS("KdKomponenTarifTotalOA").value), "", fRS("KdKomponenTarifTotalOA").value)
        fKdKomponenProfit = IIf(IsNull(fRS("KdKomponenProfit").value), "", fRS("KdKomponenProfit").value)
        fKdKomponenHargaNetto = IIf(IsNull(fRS("KdKomponenHargaNetto").value), "", fRS("KdKomponenHargaNetto").value)
        fKdKomponenTarifService = IIf(IsNull(fRS("KdKomponenTarifServisResep").value), "", fRS("KdKomponenTarifServisResep").value)
        fKdKomponenAdm = IIf(IsNull(fRS("KdKomponenAdm").value), "", fRS("KdKomponenAdm").value)
    End If
    If fKdKomponenProfit = "" Then fKdKomponenProfit = "13"
    If fKdKomponenHargaNetto = "" Then fKdKomponenHargaNetto = "09"
    If fKdKomponenTotal = "" Then fKdKomponenTotal = "06"
    If fKdKomponenTarifService = "" Then fKdKomponenTarifService = "10"
    If fKdKomponenAdm = "" Then fKdKomponenAdm = "22"
    'begin Tarif Total
    Set fRS = Nothing
    fQuery = "select NoPendaftaran from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenTotal & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        If fJmlPembebasanDB <= fHargaSatuan Then
            fJmlPembebasanPerKompTotal = (CDec(fHargaSatuan) / CDec(fTotalHarga)) * CDec(fJmlPembebasanDB)
        Else
            fJmlPembebasanPerKompTotal = (CDec(fHargaSatuan) / CDec(fTotalHarga)) * CDec(fHargaSatuan)
        End If
        fJmlHutangPerKompTotal = (CDec(fHargaSatuan) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
        fJmlTanggunganPerKompTotal = (CDec(fHargaSatuan) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
        fQuery2 = " insert into TempHargaKomponenObatAlkes values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenTotal & "'," & msubKonversiKomaTitik(CStr(fHargaSatuan)) & "," & fJmlBarang & ",null," & fKdJenisObat & "," & fNoResep & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKompTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKompTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKompTotal)) & ",null)"
    Else
        If fJmlPembebasanDB <= fHargaSatuan Then
            fJmlPembebasanPerKompTotal = (CDec(fHargaSatuan) / CDec(fTotalHarga)) * CDec(fJmlPembebasanDB)
        Else
            fJmlPembebasanPerKompTotal = (CDec(fHargaSatuan) / CDec(fTotalHarga)) * CDec(fHargaSatuan)
        End If
        fJmlHutangPerKompTotal = (CDec(fHargaSatuan) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
        fJmlTanggunganPerKompTotal = (CDec(fHargaSatuan) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
        fQuery2 = "update TempHargaKomponenObatAlkes set JmlHutangPenjamin=" & fJmlHutangPerKompTotal & ",JmlTanggunganRS=" & fJmlTanggunganPerKompTotal & ",JmlPembebasan=" & fJmlPembebasanPerKompTotal & ",HargaSatuan=" & fHargaSatuan & ",JmlBarang=" & fJmlBarang & ",KdJenisObat=" & fKdJenisObat & ",NoResep='" & fNoResep & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:dd") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenTotal & "' and NoStruk is null"
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
    'end Tarif Total
    'begin Harga Netto
    Set fRS = Nothing
    fQuery = "select NoPendaftaran from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenHargaNetto & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        If fJmlPembebasanDB <= fHargaBeli Then
            fJmlPembebasanPerKompHN = (CDec(fHargaBeli) / CDec(fTotalHarga)) * CDec(fJmlPembebasanDB)
        Else
            fJmlPembebasanPerKompHN = (CDec(fHargaBeli) / CDec(fTotalHarga)) * CDec(fHargaBeli)
        End If
        fJmlHutangPerKompHN = (CDec(fHargaBeli) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
        fJmlTanggunganPerKompHN = (CDec(fHargaBeli) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
        fQuery2 = " insert into TempHargaKomponenObatAlkes values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenHargaNetto & "'," & msubKonversiKomaTitik(CStr(fHargaBeli)) & "," & fJmlBarang & ",null," & fKdJenisObat & "," & fNoResep & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKompHN)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKompHN)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKompHN)) & ",null)"
    Else
        If fJmlPembebasanDB <= fHargaSatuan Then
            fJmlPembebasanPerKompHN = (CDec(fHargaBeli) / CDec(fTotalHarga)) * CDec(fJmlPembebasanDB)
        Else
            fJmlPembebasanPerKompHN = (CDec(fHargaBeli) / CDec(fTotalHarga)) * CDec(fHargaBeli)
        End If
        fJmlHutangPerKompHN = (CDec(fHargaBeli) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
        fJmlTanggunganPerKompHN = (CDec(fHargaBeli) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
        fQuery2 = "update TempHargaKomponenObatAlkes set JmlHutangPenjamin=" & msubKonversiKomaTitik(CStr(fJmlHutangPerKompHN)) & ",JmlTanggunganRS=" & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKompHN)) & ",JmlPembebasan=" & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKompHN)) & ",HargaSatuan=" & msubKonversiKomaTitik(CStr(fHargaBeli)) & ",JmlBarang=" & fJmlBarang & ",KdJenisObat=" & fKdJenisObat & ",NoResep='" & fNoResep & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:dd") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenHargaNetto & "' and NoStruk is null"
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
    'end Harga Netto
    'begin Profit atau Keuntungan
    If fHargaBersih <> 0 Then
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenProfit & "' and NoStruk is null"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            If fJmlPembebasanDB > fHargaBeli Then
                fJmlPembebasanPerKompP = (CDec(fHargaBersih) / CDec(fTotalHarga)) * (CDec(fJmlPembebasanDB) - CDec(fHargaBeli))
            Else
                fJmlPembebasanPerKompP = 0
            End If
            fJmlHutangPerKompP = (CDec(fHargaBersih) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKompP = (CDec(fHargaBersih) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
            fQuery2 = " insert into TempHargaKomponenObatAlkes values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenProfit & "'," & msubKonversiKomaTitik(CStr(fHargaBersih)) & "," & fJmlBarang & ",null," & fKdJenisObat & "," & fNoResep & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKompP)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKompP)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKompP)) & ",null)"
        Else
            If fJmlPembebasanDB > fHargaBeli Then
                fJmlPembebasanPerKompP = (CDec(fHargaBersih) / CDec(fTotalHarga)) * (CDec(fJmlPembebasanDB) - CDec(fHargaBeli))
            Else
                fJmlPembebasanPerKompP = 0
            End If
            fJmlHutangPerKompP = (CDec(fHargaBersih) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKompP = (CDec(fHargaBersih) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
            fQuery2 = "update TempHargaKomponenObatAlkes set JmlHutangPenjamin=" & msubKonversiKomaTitik(CStr(fJmlHutangPerKompP)) & ",JmlTanggunganRS=" & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKompP)) & ",JmlPembebasan=" & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKompP)) & ",HargaSatuan=" & msubKonversiKomaTitik(CStr(fHargaBersih)) & ",JmlBarang=" & fJmlBarang & ",KdJenisObat=" & fKdJenisObat & ",NoResep='" & fNoResep & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:dd") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenProfit & "' and NoStruk is null"
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End If
    'end Profit atau Keuntungan
    'begin Tarif Service Resep
    Set fRS = Nothing
    fQuery = "select TarifService from DetailTarifJenisObat where KdJenisObat=" & fKdJenisObat & " and KdKomponen='" & fKdKomponenTarifService & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then fTarifServiceResep = 0 Else fTarifServiceResep = IIf(IsNull(fRS("TarifService").value), 0, fRS("TarifService").value)
    Set fRS = Nothing
    fQuery = "select TarifService from DetailTarifJenisObat where KdJenisObat=" & fKdJenisObat & " and KdKomponen='01' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then fJasaRS = 0 Else fJasaRS = IIf(IsNull(fRS("TarifService").value), 0, fRS("TarifService").value)
    If (fTarifServiceResep = 0 And fJasaRS = 0) And fTarifService <> 0 Then
        fTarifServiceResep = fTarifService
    End If
    If fTarifServiceResep <> 0 Then
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenTarifService & "' and NoStruk is null"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            If fJmlPembebasanDB > fHargaSatuan Then
                fJmlPembebasanPerKompService = (CDec(fTarifServiceResep) / CDec(fTotalHarga)) * (CDec(fJmlPembebasanDB) - CDec(fHargaSatuan))
            Else
                fJmlPembebasanPerKompService = 0
            End If
            fJmlHutangPerKompService = (CDec(fTarifServiceResep) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKompService = (CDec(fTarifServiceResep) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
            fQuery2 = "insert into TempHargaKomponenObatAlkes values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenTarifService & "'," & fTarifServiceResep & "," & fJmlService & ",null," & fKdJenisObat & "," & fNoResep & "," & fJmlHutangPerKompService & "," & fJmlTanggunganPerKompService & "," & fJmlPembebasanPerKompService & ",null)"
        Else
            If fJmlPembebasanDB > fHargaSatuan Then
                fJmlPembebasanPerKompService = (CDec(fTarifServiceResep) / CDec(fTotalHarga)) * (CDec(fJmlPembebasanDB) - CDec(fHargaSatuan))
            Else
                fJmlPembebasanPerKompService = 0
            End If
            fJmlHutangPerKompService = (CDec(fTarifServiceResep) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKompService = (CDec(fTarifServiceResep) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
            fQuery2 = "update TempHargaKomponenObatAlkes set JmlHutangPenjamin=" & fJmlHutangPerKompService & ",JmlTanggunganRS=" & fJmlTanggunganPerKompService & ",JmlPembebasan=" & fJmlPembebasanPerKompService & ",HargaSatuan=" & fTarifServiceResep & ",JmlBarang=" & fJmlService & ",KdJenisObat=" & fKdJenisObat & ",NoResep='" & fNoResep & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:dd") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenTarifService & "' and NoStruk is null"
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End If
    If fJasaRS <> 0 Then
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='01' and NoStruk is null"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            If fJmlPembebasanDB > (fHargaSatuan + fTarifServiceResep) Then
                fJmlPembebasanPerKompRS = (CDec(fJasaRS) / CDec(fTotalHarga)) * (CDec(fJmlPembebasanDB) - CDec(fHargaSatuan) - CDec(fTarifServiceResep))
            Else
                fJmlPembebasanPerKompRS = 0
            End If
            fJmlHutangPerKompRS = (CDec(fJasaRS) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKompRS = (CDec(fJasaRS) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
            fQuery2 = "insert into TempHargaKomponenObatAlkes values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','01'," & fJasaRS & "," & fJmlService & ",null," & fKdJenisObat & "," & fNoResep & "," & fJmlHutangPerKompRS & "," & fJmlTanggunganPerKompRS & "," & fJmlPembebasanPerKompRS & ",null)"
        Else
            If fJmlPembebasanDB > (fHargaSatuan + fTarifServiceResep) Then
                fJmlPembebasanPerKompRS = (CDec(fJasaRS) / CDec(fTotalHarga)) * (CDec(fJmlPembebasanDB) - CDec(fHargaSatuan) - CDec(fTarifServiceResep))
            Else
                fJmlPembebasanPerKompRS = 0
            End If
            fJmlHutangPerKompRS = (CDec(fJasaRS) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKompRS = (CDec(fJasaRS) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
            fQuery2 = "update TempHargaKomponenObatAlkes set JmlHutangPenjamin=" & fJmlHutangPerKompRS & ",JmlTanggunganRS=" & fJmlTanggunganPerKompRS & ",JmlPembebasan=" & fJmlPembebasanPerKompRS & ",HargaSatuan=" & fJasaRS & ",JmlBarang=" & fJmlService & ",KdJenisObat=" & fKdJenisObat & ",NoResep='" & fNoResep & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:dd") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='01' and NoStruk is null"
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End If
    'end Tarif Service Resep
    'begin Biaya Administrasi
    If fBiayaAdministrasi <> 0 Then
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenAdm & "' and NoStruk is null"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            If fJmlPembebasanDB > (fHargaSatuan + fTarifServiceResep + fJasaRS) Then
                fJmlPembebasanPerKompAdm = (CDec(fBiayaAdministrasi) / CDec(fTotalHarga)) * (CDec(fJmlPembebasanDB) - CDec(fHargaSatuan) - CDec(fTarifServiceResep) - CDec(fJasaRS))
            Else
                fJmlPembebasanPerKompAdm = 0
            End If
            fJmlHutangPerKompAdm = (CDec(fBiayaAdministrasi) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKompAdm = (CDec(fBiayaAdministrasi) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
            fQuery2 = "insert into TempHargaKomponenObatAlkes values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenAdm & "'," & fBiayaAdministrasi & ",1,null," & fKdJenisObat & "," & fNoResep & "," & fJmlHutangPerKompAdm & "," & fJmlTanggunganPerKompAdm & "," & fJmlPembebasanPerKompAdm & ",null)"
        Else
            If fJmlPembebasanDB > (fHargaSatuan + fTarifServiceResep + fJasaRS) Then
                fJmlPembebasanPerKompAdm = (CDec(fBiayaAdministrasi) / CDec(fTotalHarga)) * (CDec(fJmlPembebasanDB) - CDec(fHargaSatuan) - CDec(fTarifServiceResep) - CDec(fJasaRS))
            Else
                fJmlPembebasanPerKompAdm = 0
            End If
            fJmlHutangPerKompAdm = (CDec(fBiayaAdministrasi) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKompAdm = (CDec(fBiayaAdministrasi) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
            fQuery2 = "update TempHargaKomponenObatAlkes set JmlHutangPenjamin=" & fJmlHutangPerKompAdm & ",JmlTanggunganRS=" & fJmlTanggunganPerKompAdm & ",JmlPembebasan=" & fJmlPembebasanPerKompAdm & ",HargaSatuan=" & fBiayaAdministrasi & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:dd") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenAdm & "' and NoStruk is null"
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End If
    'end Biaya Administrasi
End Function

'Konversi dari SP: Add_TempHargaKomponenApotik
Public Function f_AddTempHargaKomponenApotik(fNoStruk As String, fKdRuangan As String, fKdBarang As String, fKdAsal As String, fSatuanJml As String, fHargaSatuan As Currency, fHargaBeli As Currency, fJmlBarang As Double, fKdJenisObat As Variant, fJmlService As Integer, fTarifService As Currency, fBiayaAdministrasi As Currency)
    Dim fKdKomponenProfit As String
    Dim fKdKomponenTotal As String
    Dim fKdKomponenHargaNetto As String
    Dim fHargaBersih As Currency
    Dim fKdKomponenTarifService As String
    Dim fKdRuanganAsal As String
    Dim fTglStruk As Date
    Dim fKdKomponenAdm As String
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fTarifServiceResep As Currency
    Dim fJasaRS As Currency
    Dim fDiscount As Currency
    Dim fJmlPembebasanPerKompP As Currency
    Dim fJmlHutangPerKompP As Currency
    Dim fJmlTanggunganPerKompP As Currency
    Dim fJmlPembebasanPerKompHN As Currency
    Dim fJmlHutangPerKompHN As Currency
    Dim fJmlTanggunganPerKompHN As Currency
    Dim fJmlPembebasanPerKompTotal As Currency
    Dim fJmlHutangPerKompTotal As Currency
    Dim fJmlTanggunganPerKompTotal As Currency
    Dim fJmlPembebasanPerKompAdm As Currency
    Dim fJmlHutangPerKompAdm As Currency
    Dim fJmlTanggunganPerKompAdm As Currency
    Dim fJmlPembebasanPerKompService As Currency
    Dim fJmlHutangPerKompService As Currency
    Dim fJmlTanggunganPerKompService As Currency
    Dim fJmlPembebasanPerKompRS As Currency
    Dim fJmlHutangPerKompRS As Currency
    Dim fJmlTanggunganPerKompRS As Currency
    Dim fJmlHutangPenjaminDB As Currency
    Dim fJmlTanggunganRSDB As Currency
    Dim fJmlPembebasanDB As Currency
    Dim fTotalPembebasan As Currency
    Dim fTotalHarga As Currency

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select TglStruk,KdRuanganAsal,KdKelompokPasien,IdPenjamin from V_StrukPelayananApotik where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
        fTglStruk = IIf(IsNull(fRS("TglStruk").value), "", fRS("TglStruk").value)
        fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
        fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value)
    End If
    If fKdRuanganAsal = "" Then fKdRuanganAsal = fKdRuangan
    Set fRS = Nothing
    fQuery = "select Discount,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from ApotikJual where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fDiscount = IIf(IsNull(fRS("Discount").value), 0, fRS("Discount").value)
        fJmlHutangPenjaminDB = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRSDB = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanDB = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
    End If
    fHargaBersih = fHargaSatuan - fHargaBeli
    fTotalPembebasan = fJmlPembebasanDB + fDiscount
    fTotalHarga = fHargaSatuan + fTarifService + fBiayaAdministrasi
    Set fRS = Nothing
    fQuery = "select KdKomponenTarifTotalOA,KdKomponenProfit,KdKomponenHargaNetto,KdKomponenTarifServisResep,KdKomponenAdm from MasterDataPendukung"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fKdKomponenTotal = IIf(IsNull(fRS("KdKomponenTarifTotalOA").value), "", fRS("KdKomponenTarifTotalOA").value)
        fKdKomponenProfit = IIf(IsNull(fRS("KdKomponenProfit").value), "", fRS("KdKomponenProfit").value)
        fKdKomponenHargaNetto = IIf(IsNull(fRS("KdKomponenHargaNetto").value), "", fRS("KdKomponenHargaNetto").value)
        fKdKomponenTarifService = IIf(IsNull(fRS("KdKomponenTarifServisResep").value), "", fRS("KdKomponenTarifServisResep").value)
        fKdKomponenAdm = IIf(IsNull(fRS("KdKomponenAdm").value), "", fRS("KdKomponenAdm").value)
    End If
    If fKdKomponenProfit = "" Then fKdKomponenProfit = "13"
    If fKdKomponenHargaNetto = "" Then fKdKomponenHargaNetto = "09"
    If fKdKomponenTotal = "" Then fKdKomponenTotal = "06"
    If fKdKomponenTarifService = "" Then fKdKomponenTarifService = "10"
    If fKdKomponenAdm = "" Then fKdKomponenAdm = "22"
    'begin Tarif Total
    Set fRS = Nothing
    fQuery = "select NoStruk from TempHargaKomponenApotik where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenTotal & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        If fTotalPembebasan <= fHargaSatuan Then
            fJmlPembebasanPerKompTotal = (CDec(fHargaSatuan) / CDec(fTotalHarga)) * CDec(fTotalPembebasan)
        Else
            fJmlPembebasanPerKompTotal = (CDec(fHargaSatuan) / CDec(fTotalHarga)) * CDec(fHargaSatuan)
        End If
        fJmlHutangPerKompTotal = (CDec(fHargaSatuan) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
        fJmlTanggunganPerKompTotal = (CDec(fHargaSatuan) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
        fQuery2 = " insert into TempHargaKomponenApotik values('" & fNoStruk & "','" & fKdRuangan & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenTotal & "'," & msubKonversiKomaTitik(CStr(fJmlBarang)) & "," & msubKonversiKomaTitik(CStr(fHargaSatuan)) & "," & fKdJenisObat & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKompTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKompTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKompTotal)) & ",null)"
    Else
        If fTotalPembebasan <= fHargaSatuan Then
            fJmlPembebasanPerKompTotal = (CDec(fHargaSatuan) / CDec(fTotalHarga)) * CDec(fTotalPembebasan)
        Else
            fJmlPembebasanPerKompTotal = (CDec(fHargaSatuan) / CDec(fTotalHarga)) * CDec(fHargaSatuan)
        End If
        fJmlHutangPerKompTotal = (CDec(fHargaSatuan) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
        fJmlTanggunganPerKompTotal = (CDec(fHargaSatuan) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
        fQuery2 = "update TempHargaKomponenApotik set JmlHutangPenjamin=" & fJmlHutangPerKompTotal & ",JmlTanggunganRS=" & fJmlTanggunganPerKompTotal & ",JmlPembebasan=" & fJmlPembebasanPerKompTotal & ",HargaSatuan=" & fHargaSatuan & ",JmlBarang=" & fJmlBarang & ",KdJenisObat=" & fKdJenisObat & " where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenTotal & "'"
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
    'end Tarif Total
    'begin Harga Netto
    Set fRS = Nothing
    fQuery = "select NoStruk from TempHargaKomponenApotik where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenHargaNetto & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        If fTotalPembebasan <= fHargaBeli Then
            fJmlPembebasanPerKompHN = (CDec(fHargaBeli) / CDec(fTotalHarga)) * CDec(fTotalPembebasan)
        Else
            fJmlPembebasanPerKompHN = (CDec(fHargaBeli) / CDec(fTotalHarga)) * CDec(fHargaBeli)
        End If
        fJmlHutangPerKompHN = (CDec(fHargaBeli) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
        fJmlTanggunganPerKompHN = (CDec(fHargaBeli) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
        fQuery2 = " insert into TempHargaKomponenApotik values('" & fNoStruk & "','" & fKdRuangan & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenHargaNetto & "'," & msubKonversiKomaTitik(CStr(fJmlBarang)) & "," & msubKonversiKomaTitik(CStr(fHargaBeli)) & "," & fKdJenisObat & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKompHN)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKompHN)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKompHN)) & ",null)"
    Else
        If fTotalPembebasan <= fHargaSatuan Then
            fJmlPembebasanPerKompHN = (CDec(fHargaBeli) / CDec(fTotalHarga)) * CDec(fTotalPembebasan)
        Else
            fJmlPembebasanPerKompHN = (CDec(fHargaBeli) / CDec(fTotalHarga)) * CDec(fHargaBeli)
        End If
        fJmlHutangPerKompHN = (CDec(fHargaBeli) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
        fJmlTanggunganPerKompHN = (CDec(fHargaBeli) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
        fQuery2 = "update TempHargaKomponenApotik set JmlHutangPenjamin=" & fJmlHutangPerKompHN & ",JmlTanggunganRS=" & fJmlTanggunganPerKompHN & ",JmlPembebasan=" & fJmlPembebasanPerKompHN & ",HargaSatuan=" & fHargaBeli & ",JmlBarang=" & fJmlBarang & ",KdJenisObat=" & fKdJenisObat & " where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenHargaNetto & "'"
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
    'end Harga Netto
    'begin Profit atau Keuntungan
    If fHargaBersih <> 0 Then
        Set fRS = Nothing
        fQuery = "select NoStruk from TempHargaKomponenApotik where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenProfit & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            If fTotalPembebasan > fHargaBeli Then
                fJmlPembebasanPerKompP = (CDec(fHargaBersih) / CDec(fTotalHarga)) * (CDec(fTotalPembebasan) - CDec(fHargaBeli))
            Else
                fJmlPembebasanPerKompP = 0
            End If
            fJmlHutangPerKompP = (CDec(fHargaBersih) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKompP = (CDec(fHargaBersih) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
            fQuery2 = " insert into TempHargaKomponenApotik values('" & fNoStruk & "','" & fKdRuangan & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenProfit & "'," & msubKonversiKomaTitik(CStr(fJmlBarang)) & "," & msubKonversiKomaTitik(CStr(fHargaBersih)) & "," & fKdJenisObat & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKompP)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKompP)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKompP)) & ",null)"
        Else
            If fTotalPembebasan > fHargaBeli Then
                fJmlPembebasanPerKompP = (CDec(fHargaBersih) / CDec(fTotalHarga)) * (CDec(fTotalPembebasan) - CDec(fHargaBeli))
            Else
                fJmlPembebasanPerKompP = 0
            End If
            fJmlHutangPerKompP = (CDec(fHargaBersih) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKompP = (CDec(fHargaBersih) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
            fQuery2 = "update TempHargaKomponenApotik set JmlHutangPenjamin=" & fJmlHutangPerKompP & ",JmlTanggunganRS=" & fJmlTanggunganPerKompP & ",JmlPembebasan=" & fJmlPembebasanPerKompP & ",HargaSatuan=" & fHargaBersih & ",JmlBarang=" & fJmlBarang & ",KdJenisObat=" & fKdJenisObat & " where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenProfit & "'"
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End If
    'end Profit atau Keuntungan
    'begin Tarif Service Resep
    Set fRS = Nothing
    fQuery = "select TarifService from DetailTarifJenisObat where KdJenisObat=" & fKdJenisObat & " and KdKomponen='" & fKdKomponenTarifService & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then fTarifServiceResep = 0 Else fTarifServiceResep = IIf(IsNull(fRS("TarifService").value), 0, fRS("TarifService").value)
    Set fRS = Nothing
    fQuery = "select TarifService from DetailTarifJenisObat where KdJenisObat=" & fKdJenisObat & " and KdKomponen='01' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then fJasaRS = 0 Else fJasaRS = IIf(IsNull(fRS("TarifService").value), 0, fRS("TarifService").value)
    If (fTarifServiceResep = 0 And fJasaRS = 0) And fTarifService <> 0 Then fTarifServiceResep = fTarifService
    If fTarifServiceResep <> 0 Then
        Set fRS = Nothing
        fQuery = "select NoStruk from TempHargaKomponenApotik where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenTarifService & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            If fTotalPembebasan > fHargaSatuan Then
                fJmlPembebasanPerKompService = (CDec(fTarifServiceResep) / CDec(fTotalHarga)) * (CDec(fTotalPembebasan) - CDec(fHargaSatuan))
            Else
                fJmlPembebasanPerKompService = 0
            End If
            fJmlHutangPerKompService = (CDec(fTarifServiceResep) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKompService = (CDec(fTarifServiceResep) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
            fQuery2 = "insert into TempHargaKomponenApotik values('" & fNoStruk & "','" & fKdRuangan & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenTarifService & "'," & msubKonversiKomaTitik(CStr(fJmlService)) & "," & msubKonversiKomaTitik(CStr(fTarifServiceResep)) & "," & fKdJenisObat & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKompService)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKompService)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKompService)) & ",null)"
        Else
            If fTotalPembebasan > fHargaSatuan Then
                fJmlPembebasanPerKompService = (CDec(fTarifServiceResep) / CDec(fTotalHarga)) * (CDec(fTotalPembebasan) - CDec(fHargaSatuan))
            Else
                fJmlPembebasanPerKompService = 0
            End If
            fJmlHutangPerKompService = (CDec(fTarifServiceResep) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKompService = (CDec(fTarifServiceResep) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
            fQuery2 = "update TempHargaKomponenApotik set JmlHutangPenjamin=" & fJmlHutangPerKompService & ",JmlTanggunganRS=" & fJmlTanggunganPerKompService & ",JmlPembebasan=" & fJmlPembebasanPerKompService & ",HargaSatuan=" & fTarifServiceResep & ",JmlBarang=" & fJmlService & ",KdJenisObat=" & fKdJenisObat & " where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenTarifService & "'"
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End If
    If fJasaRS <> 0 Then
        Set fRS = Nothing
        fQuery = "select NoStruk from TempHargaKomponenApotik where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='01'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            If fTotalPembebasan > (fHargaSatuan + fTarifServiceResep) Then
                fJmlPembebasanPerKompRS = (CDec(fJasaRS) / CDec(fTotalHarga)) * (CDec(fTotalPembebasan) - CDec(fHargaSatuan) - CDec(fTarifServiceResep))
            Else
                fJmlPembebasanPerKompRS = 0
            End If
            fJmlHutangPerKompRS = (CDec(fJasaRS) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKompRS = (CDec(fJasaRS) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
            fQuery2 = "insert into TempHargaKomponenApotik values('" & fNoStruk & "','" & fKdRuangan & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','01'," & fJmlService & "," & fJasaRS & "," & fKdJenisObat & "," & fJmlHutangPerKompRS & "," & fJmlTanggunganPerKompRS & "," & fJmlPembebasanPerKompRS & ",null)"
        Else
            If fTotalPembebasan > (fHargaSatuan + fTarifServiceResep) Then
                fJmlPembebasanPerKompRS = (CDec(fJasaRS) / CDec(fTotalHarga)) * (CDec(fTotalPembebasan) - CDec(fHargaSatuan) - CDec(fTarifServiceResep))
            Else
                fJmlPembebasanPerKompRS = 0
            End If
            fJmlHutangPerKompRS = (CDec(fJasaRS) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKompRS = (CDec(fJasaRS) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
            fQuery2 = "update TempHargaKomponenApotik set JmlHutangPenjamin=" & fJmlHutangPerKompRS & ",JmlTanggunganRS=" & fJmlTanggunganPerKompRS & ",JmlPembebasan=" & fJmlPembebasanPerKompRS & ",HargaSatuan=" & fJasaRS & ",JmlBarang=" & fJmlService & ",KdJenisObat=" & fKdJenisObat & " where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='01'"
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End If
    'end Tarif Service Resep
    'begin Biaya Administrasi
    If fBiayaAdministrasi <> 0 Then
        Set fRS = Nothing
        fQuery = "select NoStruk from TempHargaKomponenApotik where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenAdm & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            If fTotalPembebasan > (fHargaSatuan + fTarifServiceResep + fJasaRS) Then
                fJmlPembebasanPerKompAdm = (CDec(fBiayaAdministrasi) / CDec(fTotalHarga)) * (CDec(fTotalPembebasan) - CDec(fHargaSatuan) - CDec(fTarifServiceResep) - CDec(fJasaRS))
            Else
                fJmlPembebasanPerKompAdm = 0
            End If
            fJmlHutangPerKompAdm = (CDec(fBiayaAdministrasi) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKompAdm = (CDec(fBiayaAdministrasi) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
            fQuery2 = "insert into TempHargaKomponenApotik values('" & fNoStruk & "','" & fKdRuangan & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenAdm & "',1," & msubKonversiKomaTitik(CStr(fBiayaAdministrasi)) & "," & fKdJenisObat & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKompAdm)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKompAdm)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKompAdm)) & ",null)"
        Else
            If fTotalPembebasan > (fHargaSatuan + fTarifServiceResep + fJasaRS) Then
                fJmlPembebasanPerKompAdm = (CDec(fBiayaAdministrasi) / CDec(fTotalHarga)) * (CDec(fTotalPembebasan) - CDec(fHargaSatuan) - CDec(fTarifServiceResep) - CDec(fJasaRS))
            Else
                fJmlPembebasanPerKompAdm = 0
            End If
            fJmlHutangPerKompAdm = (CDec(fBiayaAdministrasi) / CDec(fTotalHarga)) * CDec(fJmlHutangPenjaminDB)
            fJmlTanggunganPerKompAdm = (CDec(fBiayaAdministrasi) / CDec(fTotalHarga)) * CDec(fJmlTanggunganRSDB)
            fQuery2 = "update TempHargaKomponenApotik set JmlHutangPenjamin=" & fJmlHutangPerKompAdm & ",JmlTanggunganRS=" & fJmlTanggunganPerKompAdm & ",JmlPembebasan=" & fJmlPembebasanPerKompAdm & ",HargaSatuan=" & fBiayaAdministrasi & " where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenAdm & "'"
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End If
    'end Biaya Administrasi
End Function

'Konversi dari SP: Add_TempHargaKomponenIBS
Public Function f_AddTempHargaKomponenIBS(fNoPendaftaran As String, fKdRuangan As String, fTglPelayanan As Date, fKdPelayananRS As String, fKdKelas As String, fKdJenisTarif As String, fTarifCito As Integer, fJmlPelayanan As Integer, fStatusCito As String, fIdPegawai As Variant, fIdPegawaiAnastesi As String, fIdPegawai2 As Variant)
    'fIdPegawai= IdDokter; fIdPegawaiAnastesi= IdDokterAnastesi; fIdPegawai2= IdDokterPendamping/Pembantu
    Dim fKdKomponen As String
    Dim fHarga As Currency
    Dim fTotalTarif As Currency
    Dim fKdKomponenTarifTotal As String
    Dim fKdKomponenTarifCito As String
    Dim fTarifTotal As Currency
    Dim fKdJenisPegawai As String
    Dim fIdDokter As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fKdJenisPelayanan As String
    Dim fJasaDokterPendamping As Currency
    Dim fJmlDokter As Integer
    Dim fHargaJPO As Currency
    Dim fHargaJPA As Currency
    Dim fHargaJPP As Currency
    Dim fHargaJPOAkhir As Currency
    Dim fKdPelayananRSL As String
    Dim fHargaJS As Currency
    Dim fHargaJPOTemp As Currency
    Dim fTotalTarifCito As Currency
    Dim fKdJenisPegawai2 As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai1 & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisPegawai1 = fRS("KdJenisPegawai").value Else fKdJenisPegawai1 = ""
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai=" & fIdPegawai2 & ""
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisPegawai2 = fRS("KdJenisPegawai").value Else fKdJenisPegawai2 = ""
    Set fRS = Nothing
    fQuery = "select KdJnsPelayanan from ListPelayananRS where KdPelayananRS='" & fKdPelayananRS & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisPelayanan = fRS("KdJnsPelayanan").value Else fKdJenisPelayanan = ""
    Set fRS = Nothing
    fQuery = "select KdDetailJenisJasaPelayanan from PasienDaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdDetailJenisJasaPelayanan = fRS("KdDetailJenisJasaPelayanan").value Else fKdDetailJenisJasaPelayanan = ""
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTMK('" & fNoPendaftaran & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdJenisTarif & "','02') as Harga"
    Call msubRecFO(fRS, fQuery)
    fHarga = IIf(IsNull(fRS("Harga").value), 0, fRS("Harga").value)
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTMK('" & fNoPendaftaran & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdJenisTarif & "','01') as HargaJS"
    Call msubRecFO(fRS, fQuery)
    fHargaJS = IIf(IsNull(fRS("HargaJS").value), 0, fRS("HargaJS").value)
    Set fRS = Nothing
    fQuery = "select count(IdDokter) as JmlDokter from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdKomponen='02' and IdDokter=" & fIdPegawai & ""
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fJmlDokter = IIf(IsNull(fRS("JmlDokter").value), 0, fRS("JmlDokter").value)
    If fJmlDokter = 0 Then
        fHargaJPOAkhir = fHarga
        fHargaJPA = (40 * fHargaJPOAkhir) / 100
        fHargaJPP = (14 * fHargaJPOAkhir) / 100
        fJasaDokterPendamping = (20 * fHargaJPOAkhir) / 100
        If fKdJenisPegawai = "001" Then
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','01'," & fIdPegawai & "," & msubKonversiKomaTitik(CStr(fHargaJS)) & ")"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','02'," & fIdPegawai & "," & msubKonversiKomaTitik(CStr(fHargaJPOAkhir)) & ")"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','05'," & fIdPegawai & "," & msubKonversiKomaTitik(CStr(fHargaJPP)) & ")"
            Call msubRecFO(fRS2, fQuery2)
            If (fKdJenisPelayanan = "001" Or fKdJenisPelayanan = "002" Or fKdJenisPelayanan = "003" Or fKdJenisPelayanan = "004" Or fKdJenisPelayanan = "005" Or fKdJenisPelayanan = "006" Or fKdJenisPelayanan = "007") And fIdPegawaiAnastesi <> "" Then
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','04','" & fIdPegawaiAnastesi & "'," & msubKonversiKomaTitik(CStr(fHargaJPA)) & ")"
                Call msubRecFO(fRS2, fQuery2)
            End If
        Else
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','01'," & fIdPegawai & "," & msubKonversiKomaTitik(CStr(fHargaJS)) & ")"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','05'," & fIdPegawai & "," & msubKonversiKomaTitik(CStr(fHargaJPP)) & ")"
            Call msubRecFO(fRS2, fQuery2)
        End If
        If fKdDetailJenisJasaPelayanan = "02" And fKdJenisPegawai2 = "001" Then
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14'," & fIdPegawai & "," & msubKonversiKomaTitik(CStr(fJasaDokterPendamping)) & ")"
            Call msubRecFO(fRS2, fQuery2)
        End If
    End If
    If fJmlDokter = 1 Then
        Set fRS2 = Nothing
        fQuery2 = "select max(Harga) as HargaJPO from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdKomponen='02' and IdDokter=" & fIdPegawai & ""
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fHargaJPO = IIf(IsNull(fRS2("HargaJPO").value), 0, fRS2("HargaJPO").value)
        Set fRS2 = Nothing
        fQuery2 = "select KdPelayananRS from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdKomponen='02' and IdDokter=" & fIdPegawai & " and Harga=" & fHargaJPO & ""
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fKdPelayananRSL = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        If fHarga >= fHargaJPO Then
            fHargaJPOAkhir = fHarga * 1.5
            fHargaJPA = (40 * fHargaJPOAkhir) / 100
            fHargaJPP = (14 * fHargaJPOAkhir) / 100
            fJasaDokterPendamping = (20 * fHargaJPOAkhir) / 100
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenIBS set Harga=0 where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen in('02','04','05','14')"
            Call msubRecFO(fRS2, fQuery2)
            If fKdJenisPegawai = "001" Then
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','01'," & fIdPegawai & "," & fHargaJS & ")"
                Call msubRecFO(fRS2, fQuery2)
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','02'," & fIdPegawai & "," & fHargaJPOAkhir & ")"
                Call msubRecFO(fRS2, fQuery2)
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','05'," & fIdPegawai & "," & fHargaJPP & ")"
                Call msubRecFO(fRS2, fQuery2)
                If (fKdJenisPelayanan = "001" Or fKdJenisPelayanan = "002" Or fKdJenisPelayanan = "003" Or fKdJenisPelayanan = "004" Or fKdJenisPelayanan = "005" Or fKdJenisPelayanan = "006" Or fKdJenisPelayanan = "007") And fIdPegawaiAnastesi <> "" Then
                    Set fRS2 = Nothing
                    fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','04','" & fIdPegawaiAnastesi & "'," & fHargaJPA & ")"
                    Call msubRecFO(fRS2, fQuery2)
                End If
            Else
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','01'," & fIdPegawai & "," & fHargaJS & ")"
                Call msubRecFO(fRS2, fQuery2)
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','05'," & fIdPegawai & "," & fHargaJPP & ")"
                Call msubRecFO(fRS2, fQuery2)
            End If
            If fKdDetailJenisJasaPelayanan = "02" And fKdJenisPegawai2 = "001" Then
                Set fRS2 = Nothing
                fQuery2 = "select NoPendaftaran from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='14'"
                Call msubRecFO(fRS2, fQuery2)
                If fRS2.EOF = True Then
                    Set fRS2 = Nothing
                    fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14'," & fIdPegawai & "," & fJasaDokterPendamping & ")"
                    Call msubRecFO(fRS2, fQuery2)
                Else
                    Set fRS2 = Nothing
                    fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14'," & fIdPegawai & ",0)"
                    Call msubRecFO(fRS2, fQuery2)
                    Set fRS2 = Nothing
                    fQuery2 = "update TempHargaKomponenIBS set Harga=" & fJasaDokterPendamping & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='14'"
                    Call msubRecFO(fRS2, fQuery2)
                End If
            End If
        Else
            fHargaJPOAkhir = fHargaJPO * 1.5
            fHargaJPA = (40 * fHargaJPOAkhir) / 100
            fHargaJPP = (14 * fHargaJPOAkhir) / 100
            If fKdKomponen = "02" And fKdJenisPegawai2 = "001" Then
                fJasaDokterPendamping = (20 * fHargaJPOAkhir) / 100
            End If
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenIBS set Harga=" & fHargaJPOAkhir & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='02'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenIBS set Harga=" & fHargaJPA & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='04'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenIBS set Harga=" & fHargaJPP & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='05'"
            Call msubRecFO(fRS2, fQuery2)
            If fKdJenisPegawai = "001" Then
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','01'," & fIdPegawai & "," & fHargaJS & ")"
                Call msubRecFO(fRS2, fQuery2)
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','02'," & fIdPegawai & ",0)"
                Call msubRecFO(fRS2, fQuery2)
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','05'," & fIdPegawai & ",0)"
                Call msubRecFO(fRS2, fQuery2)
                If (fKdJenisPelayanan = "001" Or fKdJenisPelayanan = "002" Or fKdJenisPelayanan = "003" Or fKdJenisPelayanan = "004" Or fKdJenisPelayanan = "005" Or fKdJenisPelayanan = "006" Or fKdJenisPelayanan = "007") And fIdPegawaiAnastesi <> "" Then
                    Set fRS2 = Nothing
                    fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','04','" & fIdPegawaiAnastesi & "',0)"
                    Call msubRecFO(fRS2, fQuery2)
                End If
            Else
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','01'," & fIdPegawai & "," & fHargaJS & ")"
                Call msubRecFO(fRS2, fQuery2)
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','05'," & fIdPegawai & ",0)"
                Call msubRecFO(fRS2, fQuery2)
            End If
            If fKdDetailJenisJasaPelayanan = "02" And fKdJenisPegawai2 = "001" Then
                Set fRS = Nothing
                fQuery = "select NoPendaftaran from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='14'"
                Call msubRecFO(fRS, fQuery)
                If fRS.EOF = True Then
                    Set fRS2 = Nothing
                    fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14'," & fIdPegawai & ",0)"
                    Call msubRecFO(fRS2, fQuery2)
                End If
            Else
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14'," & fIdPegawai & ",0)"
                Call msubRecFO(fRS2, fQuery2)
                Set fRS2 = Nothing
                fQuery2 = "update TempHargaKomponenIBS set Harga=" & fJasaDokterPendamping & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='14'"
                Call msubRecFO(fRS2, fQuery2)
            End If
        End If
    End If
    If fJmlDokter > 1 Then
        Set fRS2 = Nothing
        fQuery2 = "select max(Harga) as HargaJPOTemp from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdKomponen='02' and IdDokter=" & fIdPegawai & ""
        Call msubRecFO(fRS2, fQuery2)
        If fRS.EOF = False Then fHargaJPOTemp = IIf(IsNull(fRS2("HargaJPOTemp").value), 0, fRS2("HargaJPOTemp").value)
        Set fRS2 = Nothing
        fQuery2 = "select KdPelayananRS from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdKomponen='02' and IdDokter=" & fIdPegawai & " and Harga=" & fHargaJPOTemp & ""
        Call msubRecFO(fRS2, fQuery2)
        If fRS.EOF = False Then fKdPelayananRSL = IIf(IsNull(fRS2("KdPelayananRS").value), "", fRS2("KdPelayananRS").value)
        Set fRS2 = Nothing
        fQuery2 = "select dbo.FB_NewTakeTarifBPTMK('" & fNoPendaftaran & "','" & fKdPelayananRSL & "','" & fKdKelas & "','" & fKdJenisTarif & "','02') as HargaJPO"
        Call msubRecFO(fRS2, fQuery2)
        fHargaJPO = IIf(IsNull(fRS2("HargaJPO").value), 0, fRS2("HargaJPO").value)
        If fHarga >= fHargaJPO Then
            fHargaJPOAkhir = fHarga * 2
            fHargaJPA = (40 * fHargaJPOAkhir) / 100
            fHargaJPP = (14 * fHargaJPOAkhir) / 100
            fJasaDokterPendamping = (20 * fHargaJPOAkhir) / 100
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenIBS set Harga=0 where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdKomponen in('02','04','05','14')"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','01'," & fIdPegawai & "," & fHargaJS & ")"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','02'," & fIdPegawai & "," & fHargaJPOAkhir & ")"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','05'," & fIdPegawai & "," & fHargaJPP & ")"
            Call msubRecFO(fRS2, fQuery2)
            If (fKdJenisPelayanan = "001" Or fKdJenisPelayanan = "002" Or fKdJenisPelayanan = "003" Or fKdJenisPelayanan = "004" Or fKdJenisPelayanan = "005" Or fKdJenisPelayanan = "006" Or fKdJenisPelayanan = "007") And fIdPegawaiAnastesi <> "" Then
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','04','" & fIdPegawaiAnastesi & "'," & fHargaJPA & ")"
                Call msubRecFO(fRS2, fQuery2)
            End If
            If fKdDetailJenisJasaPelayanan = "02" And fKdJenisPegawai2 = "001" Then
                Set fRS2 = Nothing
                fQuery2 = "select NoPendaftaran from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='14'"
                Call msubRecFO(fRS2, fQuery2)
                If fRS2.EOF = True Then
                    Set fRS = Nothing
                    fQuery = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14'," & fIdPegawai & "," & fJasaDokterPendamping & ")"
                    Call msubRecFO(fRS, fQuery)
                Else
                    Set fRS = Nothing
                    fQuery = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14'," & fIdPegawai & ",0)"
                    Call msubRecFO(fRS, fQuery)
                    Set fRS = Nothing
                    fQuery = "update TempHargaKomponenIBS set Harga=" & fJasaDokterPendamping & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='14'"
                    Call msubRecFO(fRS, fQuery)
                End If
            End If
        Else
            fHargaJPOAkhir = fHargaJPO * 2
            fHargaJPA = (40 * fHargaJPOAkhir) / 100
            fHargaJPP = (14 * fHargaJPOAkhir) / 100
            fJasaDokterPendamping = (20 * fHargaJPOAkhir) / 100
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenIBS set Harga=" & fHargaJPOAkhir & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='02'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenIBS set Harga=" & fHargaJPA & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='04'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenIBS set Harga=" & fHargaJPP & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='05'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','01'," & fIdPegawai & "," & fHargaJS & ")"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','02'," & fIdPegawai & ",0)"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','05'," & fIdPegawai & ",0)"
            Call msubRecFO(fRS2, fQuery2)
            If (fKdJenisPelayanan = "001" Or fKdJenisPelayanan = "002" Or fKdJenisPelayanan = "003" Or fKdJenisPelayanan = "004" Or fKdJenisPelayanan = "005" Or fKdJenisPelayanan = "006" Or fKdJenisPelayanan = "007") And fIdPegawaiAnastesi <> "" Then
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','04','" & fIdPegawaiAnastesi & "',0)"
                Call msubRecFO(fRS2, fQuery2)
            End If
            If fKdDetailJenisJasaPelayanan = "02" And fKdJenisPegawai2 = "001" Then
                Set fRS2 = Nothing
                fQuery2 = "select NoPendaftaran from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='14'"
                Call msubRecFO(fRS2, fQuery2)
                If fRS2.EOF = True Then
                    Set fRS = Nothing
                    fQuery = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14'," & fIdPegawai & ",0)"
                    Call msubRecFO(fRS, fQuery)
                Else
                    Set fRS = Nothing
                    fQuery = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14'," & fIdPegawai & ",0)"
                    Call msubRecFO(fRS, fQuery)
                    Set fRS = Nothing
                    fQuery = "update TempHargaKomponenIBS set Harga=" & fJasaDokterPendamping & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='14'"
                    Call msubRecFO(fRS, fQuery)
                End If
            End If
        End If
    End If
    '--Tarif Cito
    If fStatusCito = "1" Then
        If fKdDetailJenisJasaPelayanan = "02" Then
            fTotalTarifCito = (6 * fHargaJPOAkhir) / 100
        Else
            fTotalTarifCito = 25 * (fHargaJPA + fHargaJPOAkhir) / 100
        End If
        Set fRS2 = Nothing
        fQuery2 = "select KdKomponenTarifCito from MasterDataPendukung"
        Call msubRecFO(fRS2, fQuery2)
        If fRS.EOF = False Then fKdKomponenTarifCito = IIf(IsNull(fRS2("KdKomponenTarifCito").value), "", fRS2("KdKomponenTarifCito").value)
        If fKdKomponenTarifCito = "" Then fKdKomponenTarifCito = "07"
        Set fRS2 = Nothing
        fQuery2 = "select NoPendaftaran from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifCito & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then
            Set fRS = Nothing
            fQuery = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKomponenTarifCito & "'," & fIdPegawai & "," & fTotalTarifCito & ")"
            Call msubRecFO(fRS, fQuery)
            If fKdDetailJenisJasaPelayanan = "01" Then
                Set fRS = Nothing
                fQuery = "update TempHargaKomponenIBS set Harga=0 where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter=" & fIdPegawai & " and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='05'"
                Call msubRecFO(fRS, fQuery)
            End If
        Else
            Set fRS = Nothing
            fQuery = "update TempHargaKomponenIBS set Harga=" & fTotalTarifCito & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifCito & "'"
            Call msubRecFO(fRS, fQuery)
            If fKdDetailJenisJasaPelayanan = "01" Then
                Set fRS = Nothing
                fQuery = "update TempHargaKomponenIBS set Harga=0 where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='05' and IdDokter=" & fIdPegawai & ""
                Call msubRecFO(fRS, fQuery)
            End If
        End If
    End If
End Function

'Konversi dari SP: Delete_TempHargaKomponen
Public Function f_DeleteTempHargaKomponen(fNoPendaftaran As String, fKdPelayananRS As String, fTglPelayanan As Date, fKdRuangan As String)
    Dim fKdKomponen As String
    Dim fKdKelas As String
    Dim fIdPegawai As Variant
    Dim fKdJenisPegawai As String
    Dim fHarga As Currency
    Dim fKdRuanganAsal As String
    Dim fKdInstalasi As String
    Dim fNoLab_Rad As Variant
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fJmlPembebasan As Currency

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select NoLab_Rad from BiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdRuangan='" & fKdRuangan & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fNoLab_Rad = fRS("NoLab_Rad").value
    fNoLab_Rad = IIf(IsNull(fRS("NoLab_Rad").value), "null", "'" & fRS("NoLab_Rad").value & "'")
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "'," & fNoLab_Rad & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
    Set fRS = Nothing
    fQuery = "select KdKelas,Harga,KdKomponen,IdPegawai,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdRuangan='" & fKdRuangan & "' and NoStruk is null and NoClosing is not null"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
        fHarga = IIf(IsNull(fRS("Harga").value), 0, fRS("Harga").value)
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        fIdPegawai = IIf(IsNull(fRS("IdPegawai").value), "null", "'" & fRS("IdPegawai").value & "'")
        Set fRS2 = Nothing
        fQuery2 = "select KdJenisPegawai from DataPegawai where IdPegawai=" & fIdPegawai & ""
        Call msubRecFO(fRS2, fQuery2)
        If fRS.EOF = False Then fKdJenisPegawai = IIf(IsNull(fRS2("KdJenisPegawai").value), "", fRS2("KdJenisPegawai").value)
        Call f_AMDataPelayananTMPasienPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, fKdKelas, "M")
        If fKdJenisPegawai = "001" Then
            Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, fKdKelas, fIdPegawai, "M")
        End If
        fRS.MoveNext
    Wend
    Set fRS = Nothing
End Function

'Konversi dari SP: Delete_TempHargaKomponenApotik
Public Function f_DeleteTempHargaKomponenApotik(fNoStruk As String, fTglStruk As Date, fKdRuangan As String, fKdBarang As String, fKdAsal As String, fSatuanJml As String)
    Dim fKdKomponen As String
    Dim fHarga As Currency
    Dim fKdRuanganAsal As String
    Dim fJmlBarang As Double
    Dim fJmlService As Integer
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdRuanganAsal from V_StrukPelayananApotik where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
    If fKdRuanganAsal = "" Then fKdRuanganAsal = fKdRuangan
    Set fRS = Nothing
    fQuery = "select KdKomponen,HargaSatuan,JmlBarang from TempHargaKomponenApotik where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and NoClosing is not null"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        fHargaSatuan = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
        fJmlBarang = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang").value)
        Call f_AMDataPelayananApotikPH(fNoStruk, fTglStruk, fKdRuangan, fKdRuanganAsal, fKdBarang, fKdAsal, fSatuanJml, fKdKomponen, fHarga, fJmlService, fJmlBarang, "M")
        fRS.MoveNext
    Wend
    Set fRS = Nothing
End Function

'Konversi dari SP: Delete_TempHargaKomponenObatAlkes
Public Function f_DeleteTempHargaKomponenObatAlkes(fNoPendaftaran As String, fKdBarang As String, fTglPelayanan As Date, fKdRuangan As String, fKdAsal As String, fSatuanJml As String)
    Dim fKdKomponen As String
    Dim fKdKelas As String
    Dim fJmlBarang As Double
    Dim fHarga As Currency
    Dim fKdRuanganAsal As String
    Dim fKdInstalasi As String
    Dim fNoLab_Rad As Variant
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fJmlPembebasan As Currency

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "',null,'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','OA') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
    Set fRS = Nothing
    fQuery = "select JmlBarang,HargaSatuan,KdKomponen,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdBarang='" & fKdBarang & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdRuangan='" & fKdRuangan & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and NoStruk is null and NoClosing is not null"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fJmlBarang = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang").value)
        fHarga = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        fJmlHutangPenjamin = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRS = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasan = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        Call f_AMDataPelayananOAPasienPH(fNoPendaftaran, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdBarang, fKdAsal, fSatuanJml, fKdKomponen, fHarga, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, CInt(fJmlBarang), fJmlBarang, "M")
        fRS.MoveNext
    Wend
    Set fRS = Nothing
End Function

'Konversi dari SP: AM_RekapitulasiJasaBPApotik
Public Function f_AMRekapitulasiJasaBPApotik(fNoStruk As String, fNoBKM As String, fKdRuangan As String, fKdBarang As String, fKdAsal As String, fSatuanJml As String, fKdKomponen As String, fJmlBrg As Double, fTarif As Currency, fJmlBayar As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency, fStatus As String)
    'fStatus : A=Tambah; M=Minus
    Dim fTglBKM As Date
    Dim fTotalTarif As Currency
    Dim fJmlBayarTotal As Currency
    Dim fJmlHutangPenjaminTotal As Currency
    Dim fJmlTanggunganRSTotal As Currency
    Dim fJmlPembebasanTotal As Currency
    Dim fSisaTagihanTotal As Currency
    Dim fKdRuanganKasir As String
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fKdRuanganAsal As String
    Dim fKdPelayananRS As String
    Dim fKdDetailJenisJasaPelayanan As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    fKdPelayananRS = "000001"
    fKdDetailJenisJasaPelayanan = "03"
    Set fRS = Nothing
    fQuery = "select KdRuanganAsal from V_StrukPelayananApotik where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
    If fKdRuanganAsal = "" Then fKdRuanganAsal = fKdRuangan
    Set fRS = Nothing
    fQuery = "select TglBKM,KdRuangan from StrukBuktiKasMasuk where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fTglBKM = IIf(IsNull(fRS("TglBKM").value), "", fRS("TglBKM").value)
        fKdRuanganKasir = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
    End If
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien from StrukPelayananPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value)
        fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
    End If
    fTotalTarif = fJmlBrg * fTarif
    fJmlBayarTotal = fJmlBrg * fJmlBayar
    fJmlHutangPenjaminTotal = fJmlBrg * fJmlHutangPenjamin
    fJmlTanggunganRSTotal = fJmlBrg * fJmlTanggunganRS
    fJmlPembebasanTotal = fJmlBrg * fJmlPembebasan
    fSisaTagihanTotal = fJmlBrg * fSisaTagihan
    Set fRS = Nothing
    fQuery = "select KdRuangan from RekapitulasiJasaBPApotik where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdKomponen='" & fKdKomponen & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery2 = "insert into RekapitulasiJasaBPApotik values('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuanganKasir & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdBarang & "','" & fKdAsal & "','" & fKdKomponen & "'," & fJmlBrg & "," & msubKonversiKomaTitik(CStr(fTotalTarif)) & "," & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & ",'" & fKdPelayananRS & "','" & fKdDetailJenisJasaPelayanan & "')"
    Else
        If UCase(fStatus) = "A" Then
            fQuery2 = "update RekapitulasiJasaBPApotik set JmlBarang=JmlBarang+" & fJmlBrg & ", TotalBiaya=TotalBiaya+" & msubKonversiKomaTitik(CStr(fTotalTarif)) & ", TotalBayar=TotalBayar+" & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & ", TotalHutangPenjamin=TotalHutangPenjamin+" & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & ", TotalTanggunganRS=TotalTanggunganRS+" & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & ", TotalPembebasan=TotalPembebasan+" & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & ", TotalSisaTagihan=TotalSisaTagihan+" & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & " " _
            & " where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdKomponen='" & fKdKomponen & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery2 = "update RekapitulasiJasaBPApotik set JmlBarang=JmlBarang-" & fJmlBrg & ", TotalBiaya=TotalBiaya-" & msubKonversiKomaTitik(CStr(fTotalTarif)) & ", TotalBayar=TotalBayar-" & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & ", TotalHutangPenjamin=TotalHutangPenjamin-" & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & ", TotalTanggunganRS=TotalTanggunganRS-" & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & ", TotalPembebasan=TotalPembebasan-" & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & ", TotalSisaTagihan=TotalSisaTagihan-" & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & " " _
            & " where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdKomponen='" & fKdKomponen & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
End Function

'Konversi dari SP: AM_RekapitulasiJasaBPOAForRemunerasiFV
Public Function f_AMRekapitulasiJasaBPOAForRemunerasiFV(fNoStruk As String, fNoBKM As String, fNoPendaftaran As String, fKdRuangan As String, fKdBarang As String, fKdAsal As String, fTglPelayanan As Date, fSatuanJml As String, fKdKomponen As String, fJmlBrg As Double, fTarif As Currency, fJmlBayar As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency, fKdDetailJenisJasaPelayanan As String, fKdKelas As String, fNoLab_Rad As Variant, fStatus As String)
    'fStatus: A=Tambah; M=Minus
    Dim fTglBKM As Date
    Dim fTotalTarif As Currency
    Dim fJmlBayarTotal As Currency
    Dim fJmlHutangPenjaminTotal As Currency
    Dim fJmlTanggunganRSTotal As Currency
    Dim fJmlPembebasanTotal As Currency
    Dim fSisaTagihanTotal As Currency
    Dim fKdRuanganKasir As String
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fKdSubInstalasi As String
    Dim fKdRuanganAsal As String
    Dim fKdInstalasi As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "'," & fNoLab_Rad & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','OA') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
    Set fRS = Nothing
    fQuery = "select TglBKM,KdRuangan from StrukBuktiKasMasuk where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fTglBKM = IIf(IsNull(fRS("TglBKM").value), "", fRS("TglBKM").value)
        fKdRuanganKasir = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
    End If
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien from StrukPelayananPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value)
        fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
    End If
    Set fRS = Nothing
    fQuery = "select KdSubInstalasi from PemakaianAlkes where NoStruk='" & fNoStruk & "' and NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
    fTotalTarif = fJmlBrg * fTarif
    fJmlBayarTotal = fJmlBrg * fJmlBayar
    fJmlHutangPenjaminTotal = fJmlBrg * fJmlHutangPenjamin
    fJmlTanggunganRSTotal = fJmlBrg * fJmlTanggunganRS
    fJmlPembebasanTotal = fJmlBrg * fJmlPembebasan
    fSisaTagihanTotal = fJmlBrg * fSisaTagihan
    Set fRS = Nothing
    fQuery = "select KdRuangan from RekapitulasiJasaBPOA4Remunerasi where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdKomponen='" & fKdKomponen & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery2 = "insert into RekapitulasiJasaBPOA4Remunerasi values('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuanganKasir & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdSubInstalasi & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdDetailJenisJasaPelayanan & "','" & fKdKelas & "','" & fKdBarang & "','" & fKdAsal & "','" & fKdKomponen & "','000001'," & msubKonversiKomaTitik(CStr(fJmlBrg)) & "," & msubKonversiKomaTitik(CStr(fTotalTarif)) & "," & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & ")"
    Else
        If UCase(fStatus) = "A" Then
            fQuery2 = "update RekapitulasiJasaBPOA4Remunerasi set JmlBarang=JmlBarang+" & msubKonversiKomaTitik(CStr(fJmlBrg)) & ", TotalBiaya=TotalBiaya+" & msubKonversiKomaTitik(CStr(fTotalTarif)) & ", JmlBayar=JmlBayar+" & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & ", JmlHutangPenjamin=JmlHutangPenjamin+" & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & ", JmlTanggunganRS=JmlTanggunganRS+" & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & ", JmlPembebasan=JmlPembebasan+" & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & ", SisaTagihan=SisaTagihan+" & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & " " _
            & "where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdKomponen='" & fKdKomponen & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery2 = "update RekapitulasiJasaBPOA4Remunerasi set JmlBarang=JmlBarang-" & fJmlBrg & ", TotalBiaya=TotalBiaya-" & msubKonversiKomaTitik(CStr(fTotalTarif)) & ", JmlBayar=JmlBayar-" & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & ", JmlHutangPenjamin=JmlHutangPenjamin-" & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & ", JmlTanggunganRS=JmlTanggunganRS-" & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & ", JmlPembebasan=JmlPembebasan-" & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & ", SisaTagihan=SisaTagihan-" & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & " " _
            & "where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdKomponen='" & fKdKomponen & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
End Function

'Konversi dari SP: AM_RekapitulasiJasaBPTMForRemunerasiFV
Public Function f_AMRekapitulasiJasaBPTMForRemunerasiFV(fNoBKM As String, fNoStruk As String, fNoPendaftaran As String, fKdRuangan As String, fKdPelayananRS As String, fKdKomponen As String, fTglPelayanan As Date, fJmlPelayanan As Integer, fTarif As Currency, fJmlBayar As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency, fKdDetailJenisJasaPelayanan As String, fKdKelas As String, fNoLab_Rad As Variant, fStatus As String)
    'fStatus : A=Tambah; M=Minus
    Dim fTglBKM As Date
    Dim fTotalTarif As Currency
    Dim fJmlBayarTotal As Currency
    Dim fJmlHutangPenjaminTotal As Currency
    Dim fJmlTanggunganRSTotal As Currency
    Dim fJmlPembebasanTotal As Currency
    Dim fSisaTagihanTotal As Currency
    Dim fKdRuanganKasir As String
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fKdAsal As String
    Dim fKdSubInstalasi As String
    Dim fKdRuanganAsal As String
    Dim fKdInstalasi As String
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "'," & fNoLab_Rad & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
    Set fRS = Nothing
    fQuery = "select TglBKM,KdRuangan from StrukBuktiKasMasuk where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fTglBKM = IIf(IsNull(fRS("TglBKM").value), "", fRS("TglBKM").value)
        fKdRuanganKasir = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
    End If
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien from StrukPelayananPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value)
        fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
    End If
    Set fRS = Nothing
    fQuery = "select StatusAPBD,KdSubInstalasi from BiayaPelayanan where NoStruk='" & fNoStruk & "' and NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
        fKdAsal = IIf(IsNull(fRS("StatusAPBD").value), "", fRS("StatusAPBD").value)
    End If
    fTotalTarif = fJmlPelayanan * fTarif
    fJmlBayarTotal = fJmlPelayanan * fJmlBayar
    fJmlHutangPenjaminTotal = fJmlPelayanan * fJmlHutangPenjamin
    fJmlTanggunganRSTotal = fJmlPelayanan * fJmlTanggunganRS
    fJmlPembebasanTotal = fJmlPelayanan * fJmlPembebasan
    fSisaTagihanTotal = fJmlPelayanan * fSisaTagihan
    Set fRS = Nothing
    fQuery = "select KdRuangan from RekapitulasiJasaBPTM4Remunerasi where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery2 = "insert into RekapitulasiJasaBPTM4Remunerasi values('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuanganKasir & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdSubInstalasi & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdDetailJenisJasaPelayanan & "','" & fKdKelas & "','" & fKdPelayananRS & "','" & fKdKomponen & "','" & fKdAsal & "'," & fJmlPelayanan & "," & msubKonversiKomaTitik(CStr(fTotalTarif)) & "," & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & ")"
    Else
        If UCase(fStatus) = "A" Then
            fQuery2 = "update RekapitulasiJasaBPTM4Remunerasi set JmlPelayanan=JmlPelayanan+" & fJmlPelayanan & ",TotalBiaya=TotalBiaya+" & msubKonversiKomaTitik(CStr(fTotalTarif)) & ", JmlBayar=JmlBayar+" & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & ", JmlHutangPenjamin=JmlHutangPenjamin+" & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & ", JmlTanggunganRS=JmlTanggunganRS+" & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & ", JmlPembebasan=JmlPembebasan+" & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & ", SisaTagihan=SisaTagihan+" & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & "" _
            & "where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery2 = "update RekapitulasiJasaBPTM4Remunerasi set JmlPelayanan=JmlPelayanan-" & fJmlPelayanan & ",TotalBiaya=TotalBiaya-" & msubKonversiKomaTitik(CStr(fTotalTarif)) & ", JmlBayar=JmlBayar-" & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & ", JmlHutangPenjamin=JmlHutangPenjamin-" & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & ", JmlTanggunganRS=JmlTanggunganRS-" & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & ", JmlPembebasan=JmlPembebasan-" & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & ", SisaTagihan=SisaTagihan-" & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & "" _
            & "where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
End Function

'Konversi dari SP: AM_RekapitulasiJasaBPDokterForRemunerasiFV
Public Function f_AMRekapitulasiJasaBPDokterForRemunerasiFV(fNoBKM As String, fNoStruk As String, fNoPendaftaran As String, fKdRuangan As String, fKdPelayananRS As String, fKdKomponen As String, fTglPelayanan As Date, fJmlPelayanan As Integer, fTarif As Currency, fJmlBayar As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency, fKdDetailJenisJasaPelayanan As String, fKdKelas As String, fNoLab_Rad As Variant, fIdPegawai As Variant, fStatus As String)
    'fStatus : A=Tambah; M=Minus
    Dim fTglBKM As Date
    Dim fTotalTarif As Currency
    Dim fJmlBayarTotal As Currency
    Dim fJmlHutangPenjaminTotal As Currency
    Dim fJmlTanggunganRSTotal As Currency
    Dim fJmlPembebasanTotal As Currency
    Dim fSisaTagihanTotal As Currency
    Dim fKdRuanganKasir As String
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fKdAsal As String
    Dim fKdSubInstalasi As String
    Dim fKdRuanganAsal As String
    Dim fKdJenisPegawai As String
    Dim fKdInstalasi As String
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai=" & fIdPegawai & ""
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisPegawai = IIf(IsNull(fRS("KdJenisPegawai").value), "", fRS("KdJenisPegawai").value)
    If fKdJenisPegawai = "001" Then
        Set fRS = Nothing
        fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "'," & fNoLab_Rad & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
        Call msubRecFO(fRS, fQuery)
        fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
        Set fRS = Nothing
        fQuery = "select TglBKM,KdRuangan from StrukBuktiKasMasuk where NoBKM='" & fNoBKM & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then
            fTglBKM = IIf(IsNull(fRS("TglBKM").value), "", fRS("TglBKM").value)
            fKdRuanganKasir = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        End If
        Set fRS = Nothing
        fQuery = "select IdPenjamin,KdKelompokPasien from StrukPelayananPasien where NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then
            fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value)
            fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
        End If
        Set fRS = Nothing
        fQuery = "select StatusAPBD,KdSubInstalasi from BiayaPelayanan where NoStruk='" & fNoStruk & "' and NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then
            fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
            fKdAsal = IIf(IsNull(fRS("StatusAPBD").value), "", fRS("StatusAPBD").value)
        End If
        fTotalTarif = fJmlPelayanan * fTarif
        fJmlBayarTotal = fJmlPelayanan * fJmlBayar
        fJmlHutangPenjaminTotal = fJmlPelayanan * fJmlHutangPenjamin
        fJmlTanggunganRSTotal = fJmlPelayanan * fJmlTanggunganRS
        fJmlPembebasanTotal = fJmlPelayanan * fJmlPembebasan
        fSisaTagihanTotal = fJmlPelayanan * fSisaTagihan
        Set fRS = Nothing
        fQuery = "select KdRuangan from RekapitulasiJasaBPDokter4Remunerasi where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and IdPegawai=" & fIdPegawai & ") and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            fQuery2 = "insert into RekapitulasiJasaBPDokter4Remunerasi values('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuanganKasir & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdSubInstalasi & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdDetailJenisJasaPelayanan & "','" & fKdKelas & "','" & fKdPelayananRS & "','" & fKdKomponen & "','" & fKdAsal & "'," & fIdPegawai & "," & fJmlPelayanan & "," & msubKonversiKomaTitik(CStr(fTotalTarif)) & "," & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & ")"
        Else
            If UCase(fStatus) = "A" Then
                fQuery2 = "update RekapitulasiJasaBPDokter4Remunerasi set JmlPelayanan=JmlPelayanan+" & msubKonversiKomaTitik(CStr(fJmlPelayanan)) & ",TotalBiaya=TotalBiaya+" & msubKonversiKomaTitik(CStr(fTotalTarif)) & ", JmlBayar=JmlBayar+" & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & ", JmlHutangPenjamin=JmlHutangPenjamin+" & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & ", JmlTanggunganRS=JmlTanggunganRS+" & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & ", JmlPembebasan=JmlPembebasan+" & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & ", SisaTagihan=SisaTagihan+" & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & "" _
                & "where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and IdPegawai=" & fIdPegawai & ") and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
            Else
                fQuery2 = "update RekapitulasiJasaBPDokter4Remunerasi set JmlPelayanan=JmlPelayanan-" & fJmlPelayanan & ",TotalBiaya=TotalBiaya-" & msubKonversiKomaTitik(CStr(fTotalTarif)) & ", JmlBayar=JmlBayar-" & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & ", JmlHutangPenjamin=JmlHutangPenjamin-" & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & ", JmlTanggunganRS=JmlTanggunganRS-" & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & ", JmlPembebasan=JmlPembebasan-" & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & ", SisaTagihan=SisaTagihan-" & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & "" _
                & "where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and IdPegawai=" & fIdPegawai & ") and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
            End If
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End If
End Function

'Konversi dari SP: AM_RekapitulasiKomponenRemunerasiApotik
Public Function f_AMRekapitulasiKomponenRemunerasiApotik(fNoStruk As String, fNoBKM As String, fKdRuangan As String, fKdBarang As String, fKdAsal As String, fSatuanJml As String, fKdPelayananRS As String, fKdKomponenR As String, fKdDetailKomponenR As String, fJmlBrg As Double, fTarif As Currency, fJmlBayar As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency, fStatus As String)
    'fStatus : A=Tambah; M=Minus
    Dim fTglBKM As Date
    Dim fTotalTarif As Currency
    Dim fJmlBayarTotal As Currency
    Dim fJmlHutangPenjaminTotal As Currency
    Dim fJmlTanggunganRSTotal As Currency
    Dim fJmlPembebasanTotal As Currency
    Dim fSisaTagihanTotal As Currency
    Dim fKdRuanganKasir As String
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fKdRuanganAsal As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select fKdRuanganAsal=KdRuanganAsal from V_StrukPelayananApotik where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdRuanganAsal = IIf(IsNull(fRS("fKdRuanganAsal").value), "", fRS("fKdRuanganAsal").value)
    If fKdRuanganAsal = "" Then fKdRuanganAsal = fKdRuangan
    Set fRS = Nothing
    fQuery = "select TglBKM,KdRuangan from StrukBuktiKasMasuk where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fTglBKM = IIf(IsNull(fRS("TglBKM").value), "", fRS("TglBKM").value)
        fKdRuanganKasir = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
    End If
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien from StrukPelayananPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value)
        fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
    End If
    fTotalTarif = fJmlBrg * fTarif
    fJmlBayarTotal = fJmlBrg * fJmlBayar
    fJmlHutangPenjaminTotal = fJmlBrg * fJmlHutangPenjamin
    fJmlTanggunganRSTotal = fJmlBrg * fJmlTanggunganRS
    fJmlPembebasanTotal = fJmlBrg * fJmlPembebasan
    fSisaTagihanTotal = fJmlBrg * fSisaTagihan
    Set fRS = Nothing
    fQuery = "select KdRuangan from RekapitulasiKomponenRemunerasiApotik where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdKomponenR='" & fKdKomponenR & "' and KdDetailKomponenR='" & fKdDetailKomponenR & "' and KdPelayananRS='" & fKdPelayananRS & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery2 = "insert into RekapitulasiKomponenRemunerasiApotik values('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuanganKasir & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdPelayananRS & "','" & fKdBarang & "','" & fKdAsal & "','" & fKdKomponenR & "','" & fKdDetailKomponenR & "'," & fJmlBrg & "," & msubKonversiKomaTitik(CStr(fTotalTarif)) & "," & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & ",null)"
    Else
        If UCase(fStatus) = "A" Then
            fQuery2 = "update RekapitulasiKomponenRemunerasiApotik set JmlBarang=JmlBarang+" & fJmlBrg & ", TotalBiaya=TotalBiaya+" & msubKonversiKomaTitik(CStr(fTotalTarif)) & ", TotalBayar=TotalBayar+" & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & ", TotalHutangPenjamin=TotalHutangPenjamin+" & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & ", TotalTanggunganRS=TotalTanggunganRS+" & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & ", TotalPembebasan=TotalPembebasan+" & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & ", TotalSisaTagihan=TotalSisaTagihan+" & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & " " _
            & "where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdKomponenR='" & fKdKomponenR & "' and KdDetailKomponenR='" & fKdDetailKomponenR & "' and KdPelayananRS='" & fKdPelayananRS & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery2 = "update RekapitulasiKomponenRemunerasiApotik set JmlBarang=JmlBarang-" & fJmlBrg & ", TotalBiaya=TotalBiaya-" & msubKonversiKomaTitik(CStr(fTotalTarif)) & ", TotalBayar=TotalBayar-" & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & ", TotalHutangPenjamin=TotalHutangPenjamin-" & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & ", TotalTanggunganRS=TotalTanggunganRS-" & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & ", TotalPembebasan=TotalPembebasan-" & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & ", TotalSisaTagihan=TotalSisaTagihan-" & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & " " _
            & "where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdKomponenR='" & fKdKomponenR & "' and KdDetailKomponenR='" & fKdDetailKomponenR & "' and KdPelayananRS='" & fKdPelayananRS & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
End Function

'Konversi dari SP: AM_RekapitulasiKomponenRemunerasiDokter
Public Function f_AMRekapitulasiKomponenRemunerasiDokter(fNoBKM As String, fNoStruk As String, fNoPendaftaran As String, fKdRuangan As String, fKdPelayananRS As String, fKdKomponenR As String, fKdDetailKomponenR As String, fTglPelayanan As Date, fIdPegawai As Variant, fJmlPelayanan As Integer, fTarif As Currency, fJmlBayar As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency, fKdDetailJenisJasaPelayanan As String, fKdKelas As String, fNoLab_Rad As Variant, fKdAsal As String, fKdSubInstalasi As String, fStatus As String)
    'fStatus : A=Tambah; M=Minus
    Dim fTglBKM As Date
    Dim fTotalTarif As Currency
    Dim fJmlBayarTotal As Currency
    Dim fJmlHutangPenjaminTotal As Currency
    Dim fJmlTanggunganRSTotal As Currency
    Dim fJmlPembebasanTotal As Currency
    Dim fSisaTagihanTotal As Currency
    Dim fKdRuanganKasir As String
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fKdRuanganAsal As String
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "'," & fNoLab_Rad & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
    Set fRS = Nothing
    fQuery = "select TglBKM,KdRuangan from StrukBuktiKasMasuk where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fTglBKM = IIf(IsNull(fRS("TglBKM").value), "", fRS("TglBKM").value)
        fKdRuanganKasir = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
    End If
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien from StrukPelayananPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value)
        fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
    End If
    fTotalTarif = fJmlPelayanan * fTarif
    fJmlBayarTotal = fJmlPelayanan * fJmlBayar
    fJmlHutangPenjaminTotal = fJmlPelayanan * fJmlHutangPenjamin
    fJmlTanggunganRSTotal = fJmlPelayanan * fJmlTanggunganRS
    fJmlPembebasanTotal = fJmlPelayanan * fJmlPembebasan
    fSisaTagihanTotal = fJmlPelayanan * fSisaTagihan
    Set fRS = Nothing
    fQuery = "select KdRuangan from RekapitulasiKomponenRemunerasiDokter where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponenR='" & fKdKomponenR & "' and KdDetailKomponenR='" & fKdDetailKomponenR & "' and KdAsal='" & fKdAsal & "' and IdPegawai=" & fIdPegawai & ") and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery2 = "insert into RekapitulasiKomponenRemunerasiDokter values('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuanganKasir & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdSubInstalasi & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdDetailJenisJasaPelayanan & "','" & fKdKelas & "','" & fKdPelayananRS & "','" & fKdKomponenR & "','" & fKdDetailKomponenR & "','" & fKdAsal & "'," & fIdPegawai & "," & fJmlPelayanan & "," & msubKonversiKomaTitik(CStr(fTotalTarif)) & "," & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & ",null)"
    Else
        If UCase(fStatus) = "A" Then
            fQuery2 = "update RekapitulasiKomponenRemunerasiDokter set JmlPelayanan=JmlPelayanan+" & fJmlPelayanan & ",TotalBiaya=TotalBiaya+" & msubKonversiKomaTitik(CStr(fTotalTarif)) & ", TotalBayar=TotalBayar+" & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & ", TotalHutangPenjamin=TotalHutangPenjamin+" & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & ", TotalTanggunganRS=TotalTanggunganRS+" & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & ", TotalPembebasan=TotalPembebasan+" & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & ", TotalSisaTagihan=TotalSisaTagihan+" & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & " " _
            & "where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponenR='" & fKdKomponenR & "' and KdDetailKomponenR='" & fKdDetailKomponenR & "' and KdAsal='" & fKdAsal & "' and IdPegawai=" & fIdPegawai & ") and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery2 = "update RekapitulasiKomponenRemunerasiDokter set JmlPelayanan=JmlPelayanan-" & fJmlPelayanan & ",TotalBiaya=TotalBiaya-" & msubKonversiKomaTitik(CStr(fTotalTarif)) & ", TotalBayar=TotalBayar-" & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & ", TotalHutangPenjamin=TotalHutangPenjamin-" & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & ", TotalTanggunganRS=TotalTanggunganRS-" & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & ", TotalPembebasan=TotalPembebasan-" & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & ", TotalSisaTagihan=TotalSisaTagihan-" & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & " " _
            & "where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponenR='" & fKdKomponenR & "' and KdDetailKomponenR='" & fKdDetailKomponenR & "' and KdAsal='" & fKdAsal & "' and IdPegawai=" & fIdPegawai & ") and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
End Function

'Konversi dari SP: AM_RekapitulasiKomponenRemunerasiOATM
Public Function f_AMRekapitulasiKomponenRemunerasiOATM(fNoBKM As String, fNoStruk As String, fNoPendaftaran As String, fKdRuangan As String, fKdPelayananRS As String, fKdKomponenR As String, fKdDetailKomponenR As String, fTglPelayanan As Date, fJmlPelayanan As Integer, fTarif As Currency, fJmlBayar As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency, fKdDetailJenisJasaPelayanan As String, fKdKelas As String, fNoLab_Rad As Variant, fKdAsal As String, fKdSubInstalasi As String, fJenisOATM As String, fStatus As String)
    'fStatus : A=Tambah; M=Minus
    'fJenisOATM : OA=Obat & Alkes; TM=Tindakan Medis
    Dim fTglBKM As Date
    Dim fKdRuanganTemp As String
    Dim fTotalTarif As Currency
    Dim fJmlBayarTotal As Currency
    Dim fJmlHutangPenjaminTotal As Currency
    Dim fJmlTanggunganRSTotal As Currency
    Dim fJmlPembebasanTotal As Currency
    Dim fSisaTagihanTotal As Currency
    Dim fKdRuanganKasir As String
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fKdRuanganAsal As String
    Dim fKdInstalasi As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "'," & fNoLab_Rad & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fJenisOATM & "') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
    Set fRS = Nothing
    fQuery = "select TglBKM,KdRuangan from StrukBuktiKasMasuk where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fTglBKM = IIf(IsNull(fRS("TglBKM").value), "", fRS("TglBKM").value)
        fKdRuanganKasir = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
    End If
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien from StrukPelayananPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value)
        fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
    End If
    fTotalTarif = fJmlPelayanan * fTarif
    fJmlBayarTotal = fJmlPelayanan * fJmlBayar
    fJmlHutangPenjaminTotal = fJmlPelayanan * fJmlHutangPenjamin
    fJmlTanggunganRSTotal = fJmlPelayanan * fJmlTanggunganRS
    fJmlPembebasanTotal = fJmlPelayanan * fJmlPembebasan
    fSisaTagihanTotal = fJmlPelayanan * fSisaTagihan
    Set fRS = Nothing
    fQuery = "select KdRuangan from RekapitulasiKomponenRemunerasiOATM where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponenR='" & fKdKomponenR & "' and KdDetailKomponenR='" & fKdDetailKomponenR & "' and KdAsal='" & fKdAsal & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery2 = "insert into RekapitulasiKomponenRemunerasiOATM values('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuanganKasir & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdSubInstalasi & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdDetailJenisJasaPelayanan & "','" & fKdKelas & "','" & fKdPelayananRS & "','" & fKdKomponenR & "','" & fKdDetailKomponenR & "','" & fKdAsal & "'," & fJmlPelayanan & "," & msubKonversiKomaTitik(CStr(fTotalTarif)) & "," & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & ",null)"
    Else
        If UCase(fStatus) = "A" Then
            fQuery2 = "update RekapitulasiKomponenRemunerasiOATM set JmlPelayanan=JmlPelayanan+" & fJmlPelayanan & ",TotalBiaya=TotalBiaya+" & msubKonversiKomaTitik(CStr(fTotalTarif)) & ", TotalBayar=TotalBayar+" & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & ", TotalHutangPenjamin=TotalHutangPenjamin+" & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & ", TotalTanggunganRS=TotalTanggunganRS+" & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & ", TotalPembebasan=TotalPembebasan+" & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & ", TotalSisaTagihan=TotalSisaTagihan+" & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & " " _
            & "where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponenR='" & fKdKomponenR & "' and KdDetailKomponenR='" & fKdDetailKomponenR & "' and KdAsal='" & fKdAsal & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery2 = "update RekapitulasiKomponenRemunerasiOATM set JmlPelayanan=JmlPelayanan-" & fJmlPelayanan & ",TotalBiaya=TotalBiaya-" & msubKonversiKomaTitik(CStr(fTotalTarif)) & ", TotalBayar=TotalBayar-" & msubKonversiKomaTitik(CStr(fJmlBayarTotal)) & ", TotalHutangPenjamin=TotalHutangPenjamin-" & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTotal)) & ", TotalTanggunganRS=TotalTanggunganRS-" & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTotal)) & ", TotalPembebasan=TotalPembebasan-" & msubKonversiKomaTitik(CStr(fJmlPembebasanTotal)) & ", TotalSisaTagihan=TotalSisaTagihan-" & msubKonversiKomaTitik(CStr(fSisaTagihanTotal)) & " " _
            & "where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponenR='" & fKdKomponenR & "' and KdDetailKomponenR='" & fKdDetailKomponenR & "' and KdAsal='" & fKdAsal & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
End Function

Public Sub Add_HistoryLoginActivity(strNamaObjekDB)
    On Error GoTo hell_
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdAplikasi", adChar, adParamInput, 3, strKdAplikasi)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("TglActivity", adDate, adParamInput, , Format(Now, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("HostName", adVarChar, adParamInput, 50, strNamaHostLocal)
        .Parameters.Append .CreateParameter("NamaObjekDB", adVarChar, adParamInput, 200, strNamaObjekDB)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_HistoryLoginActivity"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("RETURN_VALUE").value <> 0 Then
            MsgBox "Ada Kesalahan dalam Hapus Rekap Komponen Biaya Pelayanan", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Sub
hell_:
    Call msubPesanError("-Add_HistoryLoginActivity")
End Sub

Public Sub subSp_HistoryLoginAplikasi(strStatus)
    On Error GoTo hell_
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdAplikasi", adChar, adParamInput, 3, strKdAplikasi)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("NamaHostAplikasi", adVarChar, adParamInput, 50, strNamaHostLocal)
        .Parameters.Append .CreateParameter("TglLogin", adDate, adParamInput, , Format(dTglLogin, "yyyy/MM/dd HH:mm:ss"))

        If strStatus = "A" Then
            .Parameters.Append .CreateParameter("TglLogout", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglLogout", adDate, adParamInput, , Format(dTglLogout, "yyyy/MM/dd HH:mm:ss"))
        End If
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, strStatus)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AU_HistoryLoginAplikasi"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("RETURN_VALUE").value <> 0 Then
            MsgBox "Ada Kesalahan dalam Hapus Rekap Komponen Biaya Pelayanan", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Sub
hell_:
    Call msubPesanError("-AU_HistoryLoginAplikasi")
End Sub

'Konversi dari SP: Add_RekapKomponenBPRemunerasiApotik
Public Function f_AddRekapKomponenBPRemunerasiApotik(fNoBKM As String, fKdBarang As String, fKdAsal As String, fKdRuangan As String, fSatuanJml As String, fKdKomponen As String, fJmlBarang As Double, fHargaSatuan As Currency, fNoStruk As String, fJmlBayarPerKomp As Currency, fJmlHutangPerKomp As Currency, fJmlTanggunganPerKomp As Currency, fJmlPembebasanPerKomp As Currency, fSisaTagihanPerKomp As Currency, fKdPelayananRS As String)
    Dim fKdKomponenR As String
    Dim fKdDetailKomponenR As String
    Dim fJmlBayarPerKompR As Currency
    Dim fJmlHutangPerKompR As Currency
    Dim fJmlTanggunganPerKompR As Currency
    Dim fJmlPembebasanPerKompR As Currency
    Dim fSisaTagihanPerKompR As Currency
    Dim fRumusPersentase As String
    Dim fKdJnsPelayanan As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fHasilRumus As Double

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    fKdDetailJenisJasaPelayanan = "01"
    Set fRS = Nothing
    fQuery = "select KdJnsPelayanan from ListPelayananRS where KdPelayananRS='" & fKdPelayananRS & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJnsPelayanan = IIf(IsNull(fRS("KdJnsPelayanan").value), "", fRS("KdJnsPelayanan").value)
    Set fRS = Nothing
    fQuery = "select KdKomponenR,KdDetailKomponenR,RumusPersentase from V_PersentaseDataRemunerasi where KdKomponen='" & fKdKomponen & "' and KdJnsPelayanan='" & fKdJnsPelayanan & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdKomponenR = IIf(IsNull(fRS("KdKomponenR").value), "", fRS("KdKomponenR").value)
        fKdDetailKomponenR = IIf(IsNull(fRS("KdDetailKomponenR").value), "", fRS("KdDetailKomponenR").value)
        fRumusPersentase = IIf(IsNull(fRS("RumusPersentase").value), "", fRS("RumusPersentase").value)
        If fRumusPersentase <> "" Then
            Set fRS2 = Nothing
            fQuery2 = "select dbo.FB_TakeRumusRemunerasi('" & fRumusPersentase & "') as HasilRumus"
            Call msubRecFO(fRS2, fQuery2)
            fHasilRumus = IIf(IsNull(fRS2("HasilRumus").value), 0, fRS2("HasilRumus").value)
            fJmlBayarPerKompR = fHasilRumus * fJmlBayarPerKomp
            fJmlHutangPerKompR = fHasilRumus * fJmlHutangPerKomp
            fJmlTanggunganPerKompR = fHasilRumus * fJmlTanggunganPerKomp
            fJmlPembebasanPerKompR = fHasilRumus * fJmlPembebasanPerKomp
            fSisaTagihanPerKompR = fHasilRumus * fSisaTagihanPerKomp
            Set fRS2 = Nothing
            fQuery2 = "select NoBKM from RekapKomponenBPRemunerasiApotik where NoBKM='" & fNoBKM & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "'  and KdKomponen='" & fKdKomponen & "' and KdKomponenR='" & fKdKomponenR & "' and KdDetailKomponenR='" & fKdDetailKomponenR & "' and NoStruk='" & fNoStruk & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                Set fRS2 = Nothing
                fQuery2 = "insert into RekapKomponenBPRemunerasiApotik values('" & fNoBKM & "','" & fKdBarang & "','" & fKdAsal & "','" & fKdRuangan & "','" & fSatuanJml & "','" & fKdKomponen & "','" & fKdKomponenR & "','" & fKdDetailKomponenR & "'," & fJmlBarang & "," & msubKonversiKomaTitik(CStr(fHargaSatuan)) & ",'" & fNoStruk & "'," & msubKonversiKomaTitik(CStr(fJmlBayarPerKompR)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKompR)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKompR)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKompR)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanPerKompR)) & ",'" & fKdPelayananRS & "',null)"
                Call msubRecFO(fRS2, fQuery2)
                Call f_AMRekapitulasiKomponenRemunerasiApotik(fNoStruk, fNoBKM, fKdRuangan, fKdBarang, fKdAsal, fSatuanJml, fKdPelayananRS, fKdKomponenR, fKdDetailKomponenR, fJmlBarang, fHargaSatuan, fJmlBayarPerKompR, fJmlHutangPerKompR, fJmlTanggunganPerKompR, fJmlPembebasanPerKompR, fSisaTagihanPerKompR, "A")
            End If
        End If
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Add_RekapKomponenBPRemunerasiOA
Public Function f_AddRekapKomponenBPRemunerasiOA(fNoBKM As String, fNoPendaftaran As String, fKdRuangan As String, fKdKelas As String, fKdKomponen As String, fKdBarang As String, fKdAsal As String, fJmlBarang As Double, fHargaSatuan As Currency, fTglPelayanan As Date, fNoStruk As String, fNoLab_Rad As Variant, fIdPegawai As Variant, fSatuanJml As String, fJmlBayarPerKomp As Currency, fJmlHutangPerKomp As Currency, fJmlTanggunganPerKomp As Currency, fJmlPembebasanPerKomp As Currency, fSisaTagihanPerKomp As Currency, fKdDetailJenisJasaPelayanan As String, fKdPaket As Variant, fNoResep As Variant, fKdPelayananRS As String, fKdSubInstalasi As String)
    Dim fKdKomponenR As String
    Dim fKdDetailKomponenR As String
    Dim fJmlBayarPerKompR As Currency
    Dim fJmlHutangPerKompR As Currency
    Dim fJmlTanggunganPerKompR As Currency
    Dim fJmlPembebasanPerKompR As Currency
    Dim fSisaTagihanPerKompR As Currency
    Dim fRumusPersentase As String
    Dim fKdJnsPelayanan As String
    Dim fHasilRumus As Double

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdJnsPelayanan from ListPelayananRS where KdPelayananRS='" & fKdPelayananRS & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJnsPelayanan = IIf(IsNull(fRS("KdJnsPelayanan").value), "", fRS("KdJnsPelayanan").value)
    Set fRS = Nothing
    fQuery = "select KdKomponenR,KdDetailKomponenR,RumusPersentase from V_PersentaseDataRemunerasi where KdKomponen='" & fKdKomponen & "' and KdJnsPelayanan='" & fKdJnsPelayanan & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdKomponenR = IIf(IsNull(fRS("KdKomponenR").value), "", fRS("KdKomponenR").value)
        fKdDetailKomponenR = IIf(IsNull(fRS("KdDetailKomponenR").value), "", fRS("KdDetailKomponenR").value)
        fRumusPersentase = IIf(IsNull(fRS("RumusPersentase").value), "", fRS("RumusPersentase").value)
        If fRumusPersentase <> "" Then
            Set fRS2 = Nothing
            fQuery2 = "select dbo.FB_TakeRumusRemunerasi('" & fRumusPersentase & "') as HasilRumus"
            Call msubRecFO(fRS2, fQuery2)
            fHasilRumus = IIf(IsNull(fRS2("HasilRumus").value), 0, fRS2("HasilRumus").value)
            fJmlBayarPerKompR = fHasilRumus * fJmlBayarPerKomp
            fJmlHutangPerKompR = fHasilRumus * fJmlHutangPerKomp
            fJmlTanggunganPerKompR = fHasilRumus * fJmlTanggunganPerKomp
            fJmlPembebasanPerKompR = fHasilRumus * fJmlPembebasanPerKomp
            fSisaTagihanPerKompR = fHasilRumus * fSisaTagihanPerKomp
            Set fRS2 = Nothing
            fQuery2 = "select NoBKM from RekapKomponenBPRemunerasiOA where NoBKM='" & fNoBKM & "' and NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "'  and KdKomponen='" & fKdKomponen & "' and KdKomponenR='" & fKdKomponenR & "' and KdDetailKomponenR='" & fKdDetailKomponenR & "' and NoStruk='" & fNoStruk & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                Set fRS2 = Nothing
                fQuery2 = "insert into RekapKomponenBPRemunerasiOA values('" & fNoBKM & "','" & fNoPendaftaran & "','" & fKdRuangan & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdKomponenR & "','" & fKdDetailKomponenR & "','" & fKdBarang & "','" & fKdAsal & "'," & msubKonversiKomaTitik(CStr(fJmlBarang)) & "," & msubKonversiKomaTitik(CStr(fHargaSatuan)) & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fNoStruk & "'," & fNoLab_Rad & "," & fIdPegawai & ",'" & fSatuanJml & "'," & msubKonversiKomaTitik(CStr(fJmlBayarPerKompR)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKompR)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKompR)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKompR)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanPerKompR)) & ",'" & fKdDetailJenisJasaPelayanan & "'," & fKdPaket & "," & fNoResep & ",'" & fKdPelayananRS & "','" & fKdSubInstalasi & "',null)"
                Call msubRecFO(fRS2, fQuery2)
                Call f_AMRekapitulasiKomponenRemunerasiOATM(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponenR, fKdDetailKomponenR, fTglPelayanan, CDec(fJmlBarang), fHargaSatuan, fJmlBayarPerKompR, fJmlHutangPerKompR, fJmlTanggunganPerKompR, fJmlPembebasanPerKompR, fSisaTagihanPerKompR, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, fKdAsal, fKdSubInstalasi, "OA", "A")
            End If
        End If
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Add_RekapKomponenBPRemunerasiTM
Public Function f_AddRekapKomponenBPRemunerasiTM(fNoBKM As String, fNoPendaftaran As String, fKdRuangan As String, fKdPelayananRS As String, fKdKomponen As String, fKdKelas As String, fJmlPelayanan As Integer, fTglPelayanan As Date, fTarif As Currency, fNoLab_Rad As Variant, fIdPegawai As Variant, fNoStruk As String, fKdDetailJenisJasaPelayanan As String, fJmlBayarPerKomp As Currency, fJmlHutangPerKomp As Currency, fJmlTanggunganPerKomp As Currency, fJmlPembebasanPerKomp As Currency, fSisaTagihanPerKomp As Currency, fKdPaket As Variant, fKdSubInstalasi As String)
    Dim fKdKomponenR As String
    Dim fKdDetailKomponenR As String
    Dim fJmlBayarPerKompR As Currency
    Dim fJmlHutangPerKompR As Currency
    Dim fJmlTanggunganPerKompR As Currency
    Dim fJmlPembebasanPerKompR As Currency
    Dim fSisaTagihanPerKompR As Currency
    Dim fRumusPersentase As String
    Dim fKdJnsPelayanan As String
    Dim fKdAsal As String
    Dim fHasilRumus As Double

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    fKdAsal = "01"
    Set fRS = Nothing
    fQuery = "select KdJnsPelayanan from ListPelayananRS where KdPelayananRS='" & fKdPelayananRS & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJnsPelayanan = IIf(IsNull(fRS("KdJnsPelayanan").value), "", fRS("KdJnsPelayanan").value)
    Set fRS = Nothing
    fQuery = "select KdKomponenR,KdDetailKomponenR,RumusPersentase from V_PersentaseDataRemunerasi where KdKomponen='" & fKdKomponen & "' and KdJnsPelayanan='" & fKdJnsPelayanan & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdKomponenR = IIf(IsNull(fRS("KdKomponenR").value), "", fRS("KdKomponenR").value)
        fKdDetailKomponenR = IIf(IsNull(fRS("KdDetailKomponenR").value), "", fRS("KdDetailKomponenR").value)
        fRumusPersentase = IIf(IsNull(fRS("RumusPersentase").value), "", fRS("RumusPersentase").value)
        If fRumusPersentase <> "" Then
            Set fRS2 = Nothing
            fQuery2 = "select dbo.FB_TakeRumusRemunerasi('" & fRumusPersentase & "') as HasilRumus"
            Call msubRecFO(fRS2, fQuery2)
            fHasilRumus = IIf(IsNull(fRS2("HasilRumus").value), 0, fRS2("HasilRumus").value)
            fJmlBayarPerKompR = fHasilRumus * fJmlBayarPerKomp
            fJmlHutangPerKompR = fHasilRumus * fJmlHutangPerKomp
            fJmlTanggunganPerKompR = fHasilRumus * fJmlTanggunganPerKomp
            fJmlPembebasanPerKompR = fHasilRumus * fJmlPembebasanPerKomp
            fSisaTagihanPerKompR = fHasilRumus * fSisaTagihanPerKomp
            Set fRS2 = Nothing
            fQuery2 = "select NoBKM from RekapKomponenBPRemunerasiTM where NoBKM='" & fNoBKM & "' and NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdKomponenR='" & fKdKomponenR & "' and KdDetailKomponenR='" & fKdDetailKomponenR & "' and NoStruk='" & fNoStruk & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                Set fRS2 = Nothing

                fQuery2 = "insert into RekapKomponenBPRemunerasiTM values('" & fNoBKM & "','" & fNoPendaftaran & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKomponen & "','" & fKdKomponenR & "','" & fKdDetailKomponenR & "','" & fKdKelas & "'," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & msubKonversiKomaTitik(CStr(fTarif)) & "," & fNoLab_Rad & "," & fIdPegawai & ",'" & fNoStruk & "','" & fKdDetailJenisJasaPelayanan & "'," & msubKonversiKomaTitik(CStr(fJmlBayarPerKompR)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKompR)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKompR)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKompR)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanPerKompR)) & "," & fKdPaket & "," & fKdSubInstalasi & ",null)"
                Call msubRecFO(fRS2, fQuery2)
                Call f_AMRekapitulasiKomponenRemunerasiOATM(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponenR, fKdDetailKomponenR, fTglPelayanan, fJmlPelayanan, fTarif, fJmlBayarPerKompR, fJmlHutangPerKompR, fJmlTanggunganPerKompR, fJmlPembebasanPerKompR, fSisaTagihanPerKompR, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, fKdAsal, fKdSubInstalasi, "TM", "A")
                Call f_AMRekapitulasiKomponenRemunerasiDokter(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponenR, fKdDetailKomponenR, fTglPelayanan, fIdPegawai, fJmlPelayanan, fTarif, fJmlBayarPerKompR, fJmlHutangPerKompR, fJmlTanggunganPerKompR, fJmlPembebasanPerKompR, fSisaTagihanPerKompR, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, fKdAsal, fKdSubInstalasi, "A")
            End If
        End If
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Add_RekapKomponenBiayaPelayananApotik
Public Function f_AddRekapKomponenBiayaPelayananApotik(fNoBKM As String, fNoStruk As String, fTotalBiayaHrsDibayar As Currency, fJmlBayar As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency)
    Dim fKdRuangan As String
    Dim fKdBarang As String
    Dim fKdAsal As String
    Dim fJmlBarang As Double
    Dim fHargaSatuan As Currency
    Dim fJmlHutangPenjaminDP As Currency
    Dim fJmlTanggunganRSDP As Currency
    Dim fJmlPembebasanDP As Currency
    Dim fJmlHrsDibayar As Currency
    Dim fHargaBeli As Currency
    Dim fHargaPerKomponen As Currency
    Dim fJmlPembebasanPerKomp As Currency
    Dim fJmlBayarPerKomp As Currency
    Dim fJmlHutangPerKomp As Currency
    Dim fJmlTanggunganPerKomp As Currency
    Dim fSisaTagihanPerKomp As Currency
    Dim fNoResep As Variant
    Dim fPPn As Currency
    Dim fDiscount As Currency
    Dim fJmlService As Integer
    Dim fTarifService As Currency
    Dim fKdKomponen As String
    Dim fTotalTarif As Currency
    Dim fSatuanJml As String
    Dim fTempJmlBayar As Currency
    Dim fBiayaAdministrasi As Currency
    Dim fJmlItem As Double
    Dim fHargaSatuanKomp As Currency
    Dim fX1 As Currency 'hutang barang
    Dim fX2 As Currency 'hutang service
    Dim fX3 As Currency 'hutang admin
    Dim fY1 As Currency 'tanggungan barang
    Dim fY2 As Currency 'tanggungan service
    Dim fY3 As Currency 'tanggungan admin
    Dim fTotalBiaya As Currency
    Dim fTempPembebasan As Currency
    Dim fJmlHutangPenjaminDB As Currency
    Dim fJmlTanggunganRSDB As Currency
    Dim fJmlPembebasanDB As Currency
    Dim fKdPelayananRS As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    fKdPelayananRS = "000002"
    Set fRS = Nothing
    fQuery = "select KdRuangan,KdBarang,KdAsal,SatuanJml,JmlBarang,HargaSatuan,Ppn,Discount,HargaBeli,JmlService,TarifService,JmlHutangPenjamin,JmlTanggunganRS,BiayaAdministrasi,JmlPembebasan from ApotikJual where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdBarang = IIf(IsNull(fRS("KdBarang").value), "", fRS("KdBarang").value)
        fKdAsal = IIf(IsNull(fRS("KdAsal").value), "", fRS("KdAsal").value)
        fSatuanJml = IIf(IsNull(fRS("SatuanJml").value), "", fRS("SatuanJml").value)
        fJmlBarang = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang").value)
        fHargaSatuan = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
        fPPn = IIf(IsNull(fRS("Ppn").value), 0, fRS("Ppn").value)
        fDiscount = IIf(IsNull(fRS("Discount").value), 0, fRS("Discount").value)
        fHargaBeli = IIf(IsNull(fRS("HargaBeli").value), 0, fRS("HargaBeli").value)
        fJmlService = IIf(IsNull(fRS("JmlService").value), 0, fRS("JmlService").value)
        fTarifService = IIf(IsNull(fRS("TarifService").value), 0, fRS("TarifService").value)
        fJmlHutangPenjaminDP = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRSDP = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fBiayaAdministrasi = IIf(IsNull(fRS("BiayaAdministrasi").value), 0, fRS("BiayaAdministrasi").value)
        fJmlPembebasanDP = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)

        fTotalTarif = (((fHargaSatuan * fJmlBarang) + (fPPn * fJmlBarang) + (fTarifService * fJmlService)) - (fDiscount * fJmlBarang)) + fBiayaAdministrasi
        fTotalBiaya = fHargaSatuan + fPPn + fTarifService + fBiayaAdministrasi - fDiscount
        'Hitung Hutang Penjamin Per Komponen
        fX1 = (CDec((fHargaSatuan + fPPn - fDiscount)) / CDec(fTotalBiaya)) * CDec(fJmlHutangPenjaminDP)
        fX2 = (CDec(fTarifService) / CDec(fTotalBiaya)) * CDec(fJmlHutangPenjaminDP)
        fX3 = (CDec(fBiayaAdministrasi) / CDec(fTotalBiaya)) * CDec(fJmlHutangPenjaminDP)
        'Hitung Tanggungan RS Per Komponen
        fY1 = (CDec((fHargaSatuan + fPPn - fDiscount)) / CDec(fTotalBiaya)) * CDec(fJmlTanggunganRSDP)
        fY2 = (CDec(fTarifService) / CDec(fTotalBiaya)) * CDec(fJmlTanggunganRSDP)
        fY3 = (CDec(fBiayaAdministrasi) / CDec(fTotalBiaya)) * CDec(fJmlTanggunganRSDP)
        fJmlHrsDibayar = fTotalTarif - (fX1 * fJmlBarang) - (fX2 * fJmlService) - fX3 - (fY1 * fJmlBarang) - (fY2 * fJmlService) - fY3
        Set fRS2 = Nothing
        fQuery2 = "select KdKomponen,HargaSatuan,JmlBarang,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from TempHargaKomponenApotik where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "'"
        Call msubRecFO(fRS2, fQuery2)
        While fRS2.EOF = False
            fKdKomponen = IIf(IsNull(fRS2("KdKomponen").value), "", fRS2("KdKomponen").value)
            fHargaSatuanKomp = IIf(IsNull(fRS2("HargaSatuan").value), 0, fRS2("HargaSatuan").value)
            fJmlItem = IIf(IsNull(fRS2("JmlBarang").value), 0, fRS2("JmlBarang").value)

            fJmlHutangPenjaminDB = IIf(IsNull(fRS2("JmlHutangPenjamin").value), 0, fRS2("JmlHutangPenjamin").value)
            fJmlTanggunganRSDB = IIf(IsNull(fRS2("JmlTanggunganRS").value), 0, fRS2("JmlTanggunganRS").value)
            fJmlPembebasanDB = IIf(IsNull(fRS2("JmlPembebasan").value), 0, fRS2("JmlPembebasan").value)

            If fJmlHrsDibayar = 0 Then
                fSisaTagihanPerKomp = 0
                fJmlBayarPerKomp = 0
                If fJmlPembebasanDB = 0 And fJmlPembebasanDP <> 0 Then
                    fJmlPembebasanPerKomp = (CDec(fHargaSatuanKomp) / CDec(fTotalBiaya)) * CDec(fJmlPembebasanDP)
                Else
                    fJmlPembebasanPerKomp = 0
                End If
                If fJmlHutangPenjaminDB = 0 And fJmlHutangPenjaminDP <> 0 Then
                    fJmlHutangPerKomp = (CDec(fHargaSatuanKomp) / CDec(fTotalBiaya)) * CDec(fJmlHutangPenjaminDP)
                Else
                    fJmlHutangPerKomp = fJmlHutangPenjaminDB
                End If
                If fJmlTanggunganRSDB = 0 And fJmlTanggunganRSDP <> 0 Then
                    fJmlTanggunganPerKomp = (CDec(fHargaSatuanKomp) / CDec(fTotalBiaya)) * CDec(fJmlTanggunganRSDP)
                Else
                    fJmlTanggunganPerKomp = fJmlTanggunganRSDB
                End If
            Else
                If fTotalBiayaHrsDibayar = 0 Then
                    fTempJmlBayar = 0
                    fTempPembebasan = 0
                Else
                    fTempJmlBayar = (CDec(fJmlHrsDibayar) / CDec(fTotalBiayaHrsDibayar)) * CDec(fJmlBayar) 'hitung jumlah bayar per barang
                    fTempPembebasan = (CDec(fJmlHrsDibayar) / CDec(fTotalBiayaHrsDibayar)) * CDec(fJmlPembebasan)
                End If
                fJmlPembebasanPerKomp = (CDec(fHargaSatuanKomp) / CDec(fTotalTarif)) * CDec(fTempPembebasan)
                fSisaTagihanPerKomp = (CDec(fHargaSatuanKomp) / CDec(fTotalTarif)) * CDec(fSisaTagihan)
                fJmlBayarPerKomp = (CDec(fHargaSatuanKomp) / CDec(fTotalTarif)) * CDec(fTempJmlBayar)
                If fJmlHutangPenjaminDB = 0 And fJmlHutangPenjaminDP <> 0 Then
                    fJmlHutangPerKomp = (CDec(fHargaSatuanKomp) / CDec(fTotalBiaya)) * CDec(fJmlHutangPenjaminDP)
                Else
                    fJmlHutangPerKomp = fJmlHutangPenjaminDB
                End If
                If fJmlTanggunganRSDB = 0 And fJmlTanggunganRSDP <> 0 Then
                    fJmlTanggunganPerKomp = (CDec(fHargaSatuanKomp) / CDec(fTotalBiaya)) * CDec(fJmlTanggunganRSDP)
                Else
                    fJmlTanggunganPerKomp = fJmlTanggunganRSDB
                End If
            End If
            Set fRS3 = Nothing
            fQuery3 = "insert into RekapKomponenBiayaPelayananApotik values('" & fNoBKM & "','" & fKdBarang & "','" & fKdAsal & "','" & fKdRuangan & "','" & fSatuanJml & "','" & fKdKomponen & "'," & fJmlItem & "," & msubKonversiKomaTitik(CStr(fHargaSatuanKomp)) & ",'" & fNoStruk & "'," & msubKonversiKomaTitik(CStr(fJmlBayarPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKomp)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanPerKomp)) & ",'" & fKdPelayananRS & "')"
            Call msubRecFO(fRS3, fQuery3)
            Call f_AMRekapitulasiJasaBPApotik(fNoStruk, fNoBKM, fKdRuangan, fKdBarang, fKdAsal, fSatuanJml, fKdKomponen, fJmlItem, fHargaSatuanKomp, fJmlBayarPerKomp, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, "A")
            If fKdKomponen = "13" Then
                Call f_AddRekapKomponenBPRemunerasiApotik(fNoBKM, fKdBarang, fKdAsal, fKdRuangan, fSatuanJml, fKdKomponen, fJmlItem, fHargaSatuanKomp, fNoStruk, fJmlBayarPerKomp, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, fKdPelayananRS)
            End If
            fRS2.MoveNext
        Wend
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Add_RekapKomponenBiayaPelayananApotikClaim
Public Function f_AddRekapKomponenBiayaPelayananApotikClaim(fNoBKM As String, fNoBKMSebelumnya As String, fNoStruk As String, fJmlBayar As Currency)
    Dim fKdRuangan As String
    Dim fKdBarang As String
    Dim fKdAsal As String
    Dim fJmlBarang As Double
    Dim fHargaSatuan As Currency
    Dim fKdKomponen As String
    Dim fJmlHutangPenjaminL As Currency
    Dim fJmlTanggunganRSL As Currency
    Dim fJmlBayarL As Currency
    Dim fJmlPembebasanL As Currency
    Dim fSisaTagihanL As Currency
    Dim fJmlPembebasanPerKomp As Currency
    Dim fJmlBayarPerKomp As Currency
    Dim fJmlHutangPerKomp As Currency
    Dim fTotalHutangPenjamin As Currency
    Dim fSatuanJml As String
    Dim fKdPelayananRS As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    fKdPelayananRS = "000002"
    Set fRS = Nothing
    fQuery = "select KdRuangan,KdBarang,KdAsal,KdKomponen,SatuanJml,JmlBarang,HargaSatuan,JmlBayar,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,SisaTagihan from RekapKomponenBiayaPelayananApotik where NoBKM='" & fNoBKMSebelumnya & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdBarang = IIf(IsNull(fRS("KdBarang").value), "", fRS("KdBarang").value)
        fKdAsal = IIf(IsNull(fRS("KdAsal").value), "", fRS("KdAsal").value)
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        fSatuanJml = IIf(IsNull(fRS("SatuanJml").value), "", fRS("SatuanJml").value)
        fJmlBarang = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang").value)
        fHargaSatuan = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
        fJmlBayarL = IIf(IsNull(fRS("JmlBayar").value), 0, fRS("JmlBayar").value)
        fJmlHutangPenjaminL = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRSL = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanL = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fSisaTagihanL = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value)
        If fJmlHutangPenjaminL <> 0 Then
            Set fRS2 = Nothing
            fQuery2 = "select sum(JmlHutangPenjamin) as JmlHutangPenjaminSum from RekapKomponenBiayaPelayananApotik where  NoBKM='" & fNoBKMSebelumnya & "' and NoStruk='" & fNoStruk & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = False Then fTotalHutangPenjamin = IIf(IsNull(fRS2("JmlHutangPenjaminSum").value), 0, fRS2("JmlHutangPenjaminSum").value)
            fJmlBayarPerKomp = (CDec(fJmlHutangPenjaminL) / CDec(fTotalHutangPenjamin)) * CDec(fJmlBayar)
            fJmlHutangPerKomp = CDec(fJmlHutangPenjaminL) - fJmlBayarPerKomp
            Set fRS2 = Nothing
            fQuery2 = "insert into RekapKomponenBiayaPelayananApotik values('" & fNoBKM & "','" & fKdBarang & "','" & fKdAsal & "','" & fKdRuangan & "','" & fSatuanJml & "','" & fKdKomponen & "'," & msubKonversiKomaTitik(CStr(fJmlBarang)) & "," & msubKonversiKomaTitik(CStr(fHargaSatuan)) & ",'" & fNoStruk & "'," & msubKonversiKomaTitik(CStr(fJmlBayarPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSL)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanL)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanL)) & ",'" & fKdPelayananRS & "')"
            Call msubRecFO(fRS2, fQuery2)
            Call f_AMRekapitulasiJasaBPApotik(fNoStruk, fNoBKM, fKdRuangan, fKdBarang, fKdAsal, fSatuanJml, fKdKomponen, fJmlBarang, fHargaSatuan, fJmlBayarPerKomp, fJmlHutangPerKomp, fJmlTanggunganRSL, fJmlPembebasanL, fSisaTagihanL, "A")
            If fKdKomponen = "13" Then
                Call f_AddRekapKomponenBPRemunerasiApotik(fNoBKM, fKdBarang, fKdAsal, fKdRuangan, fSatuanJml, fKdKomponen, fJmlBarang, fHargaSatuan, fNoStruk, fJmlBayarPerKomp, fJmlHutangPerKomp, fJmlTanggunganRSL, fJmlPembebasanL, fSisaTagihanL, fKdPelayananRS)
            End If
        End If
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Add_RekapKomponenBiayaPelayananApotikKredit
Public Function f_AddRekapKomponenBiayaPelayananApotikKredit(fNoBKM As String, fNoBKMSebelumnya As String, fNoStruk As String, fJmlBayar As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency)
    Dim fKdRuangan As String
    Dim fKdBarang As String
    Dim fKdAsal As String
    Dim fJmlBarang As Double
    Dim fHargaSatuan As Currency
    Dim fKdKomponen As String
    Dim fJmlHutangPenjaminL As Currency
    Dim fJmlTanggunganRSL As Currency
    Dim fJmlBayarL As Currency
    Dim fJmlPembebasanL As Currency
    Dim fSisaTagihanL As Currency
    Dim fJmlPembebasanPerKomp As Currency
    Dim fJmlBayarPerKomp As Currency
    Dim fSisaTagihanPerKomp As Currency
    Dim fTotalSisaTagihan As Currency
    Dim fSatuanJml As String
    Dim fKdPelayananRS As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    fKdPelayananRS = "000002"
    Set fRS = Nothing
    fQuery = "select KdRuangan,KdBarang,KdAsal,KdKomponen,SatuanJml,JmlBarang,HargaSatuan,JmlBayar,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,SisaTagihan from RekapKomponenBiayaPelayananApotik where NoBKM='" & fNoBKMSebelumnya & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdBarang = IIf(IsNull(fRS("KdBarang").value), "", fRS("KdBarang").value)
        fKdAsal = IIf(IsNull(fRS("KdAsal").value), "", fRS("KdAsal").value)
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        fSatuanJml = IIf(IsNull(fRS("SatuanJml").value), "", fRS("SatuanJml").value)
        fJmlBarang = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang").value)
        fHargaSatuan = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
        fJmlBayarL = IIf(IsNull(fRS("JmlBayar").value), 0, fRS("JmlBayar").value)
        fJmlHutangPenjaminL = IIf(IsNull(fRS("JmlHutangPenjaminL").value), 0, fRS("JmlHutangPenjaminL").value)
        fJmlTanggunganRSL = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanL = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fSisaTagihanL = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value)
        If fSisaTagihanL <> 0 Then
            Set fRS2 = Nothing
            fQuery2 = "select sum(SisaTagihan) as SisaTagihanSum from RekapKomponenBiayaPelayananApotik where  NoBKM='" & fNoBKMSebelumnya & "' and NoStruk='" & fNoStruk & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = False Then fTotalSisaTagihan = IIf(IsNull(fRS2("SisaTagihanSum").value), 0, fRS2("SisaTagihanSum").value)
            fJmlPembebasanPerKomp = (CDec(fSisaTagihanL) / CDec(fTotalSisaTagihan)) * CDec(fJmlPembebasan)
            fSisaTagihanPerKomp = (CDec(fSisaTagihanL) / CDec(fTotalSisaTagihan)) * CDec(fSisaTagihan)
            fJmlBayarPerKomp = (CDec(fSisaTagihanL) / CDec(fTotalSisaTagihan)) * CDec(fJmlBayar)
            Set fRS2 = Nothing
            fQuery2 = "insert into RekapKomponenBiayaPelayananApotik values('" & fNoBKM & "','" & fKdBarang & "','" & fKdAsal & "','" & fKdRuangan & "','" & fSatuanJml & "','" & fKdKomponen & "'," & msubKonversiKomaTitik(CStr(fJmlBarang)) & "," & msubKonversiKomaTitik(CStr(fHargaSatuan)) & ",'" & fNoStruk & "'," & msubKonversiKomaTitik(CStr(fJmlBayarPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminL)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSL)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKomp)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanPerKomp)) & ",'" & fKdPelayananRS & "')"
            Call msubRecFO(fRS2, fQuery2)
            Call f_AMRekapitulasiJasaBPApotik(fNoStruk, fNoBKM, fKdRuangan, fKdBarang, fKdAsal, fSatuanJml, fKdKomponen, fJmlBarang, fHargaSatuan, fJmlBayarPerKomp, fJmlHutangPenjaminL, fJmlTanggunganRSL, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, "A")
            If fKdKomponen = "13" Then
                Call f_AddRekapKomponenBPRemunerasiApotik(fNoBKM, fKdBarang, fKdAsal, fKdRuangan, fSatuanJml, fKdKomponen, fJmlBarang, fHargaSatuan, fNoStruk, fJmlBayarPerKomp, fJmlHutangPenjaminL, fJmlTanggunganRSL, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, fKdPelayananRS)
            End If
        End If
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Add_RekapKomponenBiayaPelayananOA
Public Function f_AddRekapKomponenBiayaPelayananOA(fNoBKM As String, fNoStruk As String, fTotalBiayaHrsDibayar As Currency, fJmlBayar As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency, fJmlDiscount As Currency)
    Dim fNoPendaftaran As String
    Dim fKdRuangan As String
    Dim fKdKelas As String
    Dim fKdBarang As String
    Dim fKdAsal As String
    Dim fJmlBarang As Double
    Dim fHargaSatuan As Currency
    Dim fTglPelayanan As Date
    Dim fNoLab_Rad As Variant
    Dim fIdPegawai As Variant
    Dim fSatuanJml As String
    Dim fJmlPembebasanPerBrg As Currency
    Dim fJmlBayarPerBrg As Currency
    Dim fJmlHutangPerBrg As Currency
    Dim fJmlTanggunganPerBrg As Currency
    Dim fSisaTagihanPerBrg As Currency
    Dim fKdKomponen As String
    Dim fKdKelasPenjaminDP As String
    Dim fTarifKelasPenjaminDP As Currency
    Dim fJmlHutangPenjaminDP As Currency
    Dim fJmlTanggunganRSDP As Currency
    Dim fJmlPembebasanDP As Currency
    Dim fSelisihTarifKelasPenjamin As Currency
    Dim fSelisihTarifDgnTanggungan As Currency
    Dim fJmlHrsDibayar As Currency
    Dim fTempJmlBayar As Currency
    Dim fTempSisaTagihan As Currency
    Dim fTempPembebasan As Currency
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fHargaBeli As Currency
    Dim fIdPegawai2 As Variant
    Dim fIdUser As String
    Dim fHargaPerKomponen As Currency
    Dim fJmlPembebasanPerKomp As Currency
    Dim fJmlBayarPerKomp As Currency
    Dim fJmlHutangPerKomp As Currency
    Dim fJmlTanggunganPerKomp As Currency
    Dim fSisaTagihanPerKomp As Currency
    Dim fKdKomponenTemp As String
    Dim fNoResep As Variant
    Dim fTotalHargaSatuan As Currency
    Dim fTarifService As Currency
    Dim fJmlService As Integer
    Dim fKdPaket As Variant
    Dim fBiayaAdministrasi As Currency
    Dim fJmlItem As Double
    Dim fHargaSatuanKomp As Currency
    Dim fX1 As Currency 'hutang barang
    Dim fX2 As Currency 'hutang service
    Dim fX3 As Currency 'hutang admin
    Dim fY1 As Currency 'tanggungan barang
    Dim fY2 As Currency 'tanggungan service
    Dim fY3 As Currency 'tanggungan admin
    Dim fZ1 As Currency 'pembebasan barang
    Dim fZ2 As Currency 'pembebasan service
    Dim fZ3 As Currency 'pembebasan admin
    Dim fTotalBiaya As Currency
    Dim fKdSubInstalasi As String
    Dim fJmlHutangPenjaminDB As Currency
    Dim fJmlTanggunganRSDB As Currency
    Dim fJmlPembebasanDB As Currency
    Dim fKdPelayananRS As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    fKdPelayananRS = "000001"
    Set fRS = Nothing
    fQuery = "select KdRuangan,KdBarang,KdAsal,TglPelayanan,SatuanJml,NoPendaftaran,KdKelas,JmlBarang,HargaSatuan,NoLab_Rad,IdPegawai,KdKelasPenjamin,TarifKelasPenjamin,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,IdPegawai2,NoResep,TarifService,JmlService,KdPaket,BiayaAdministrasi,KdSubInstalasi from DetailPemakaianAlkes where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdBarang = IIf(IsNull(fRS("KdBarang").value), "", fRS("KdBarang").value)
        fKdAsal = IIf(IsNull(fRS("KdAsal").value), "", fRS("KdAsal").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fSatuanJml = IIf(IsNull(fRS("SatuanJml").value), "", fRS("SatuanJml").value)
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
        fJmlBarang = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang").value)
        fHargaSatuan = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
        fNoLab_Rad = IIf(IsNull(fRS("NoLab_Rad").value), "null", "'" & fRS("NoLab_Rad").value & "'")
        fIdPegawai = IIf(IsNull(fRS("IdPegawai").value), "null", "'" & fRS("IdPegawai").value & "'")
        fKdKelasPenjaminDP = IIf(IsNull(fRS("KdKelasPenjamin").value), "", fRS("KdKelasPenjamin").value)
        fTarifKelasPenjaminDP = IIf(IsNull(fRS("TarifKelasPenjamin").value), 0, fRS("TarifKelasPenjamin").value)
        fJmlHutangPenjaminDP = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRSDP = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanDP = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fIdPegawai2 = IIf(IsNull(fRS("IdPegawai2").value), "", fRS("IdPegawai2").value)
        fNoResep = IIf(IsNull(fRS("NoResep").value), "null", "'" & fRS("NoResep").value & "'")
        fTarifService = IIf(IsNull(fRS("TarifService").value), 0, fRS("TarifService").value)
        fJmlService = IIf(IsNull(fRS("JmlService").value), 0, fRS("JmlService").value)
        fKdPaket = IIf(IsNull(fRS("KdPaket").value), "null", "'" & fRS("KdPaket").value & "'")
        fBiayaAdministrasi = IIf(IsNull(fRS("BiayaAdministrasi").value), 0, fRS("BiayaAdministrasi").value)
        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
        fTotalHargaSatuan = (fHargaSatuan * fJmlBarang) + (fTarifService * fJmlService) + fBiayaAdministrasi
        fTotalBiaya = fHargaSatuan + fTarifService + fBiayaAdministrasi
        'Hitung Hutang Penjamin Per Komponen
        If fTotalBiaya = 0 Then
            fX1 = 0
            fX2 = 0
            fX3 = 0
            fY1 = 0
            fY2 = 0
            fY3 = 0
            fZ1 = 0
            fZ2 = 0
            fZ3 = 0
        Else
            'Hitung Hutang Penjamin Per Komponen
            fX1 = (CDec((fHargaSatuan)) / CDec(fTotalBiaya)) * CDec(fJmlHutangPenjaminDP)
            fX2 = (CDec(fTarifService) / CDec(fTotalBiaya)) * CDec(fJmlHutangPenjaminDP)
            fX3 = (CDec(fBiayaAdministrasi) / CDec(fTotalBiaya)) * CDec(fJmlHutangPenjaminDP)
            'Hitung Tanggungan RS Per Komponen
            fY1 = (CDec((fHargaSatuan)) / CDec(fTotalBiaya)) * CDec(fJmlTanggunganRSDP)
            fY2 = (CDec(fTarifService) / CDec(fTotalBiaya)) * CDec(fJmlTanggunganRSDP)
            fY3 = (CDec(fBiayaAdministrasi) / CDec(fTotalBiaya)) * CDec(fJmlTanggunganRSDP)
            'Hitung Pembebasan Per Komponen
            fZ1 = (CDec((fHargaSatuan)) / CDec(fTotalBiaya)) * CDec(fJmlPembebasanDP)
            fZ2 = (CDec(fTarifService) / CDec(fTotalBiaya)) * CDec(fJmlPembebasanDP)
            fZ3 = (CDec(fBiayaAdministrasi) / CDec(fTotalBiaya)) * CDec(fJmlPembebasanDP)
        End If
        fJmlHrsDibayar = fTotalHargaSatuan - (fX1 * fJmlBarang) - (fX2 * fJmlService) - fX3 - (fY1 * fJmlBarang) - (fY2 * fJmlService) - fY3 - (fZ1 * fJmlBarang) - (fZ2 * fJmlService) - fZ3
        Set fRS2 = Nothing
        fQuery2 = "select KdDetailJenisJasaPelayanan from DetailKelasPelayanan where KdKelas='" & fKdKelas & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fKdDetailJenisJasaPelayanan = IIf(IsNull(fRS2("KdDetailJenisJasaPelayanan").value), "", fRS2("KdDetailJenisJasaPelayanan").value)
        Set fRS2 = Nothing
        fQuery2 = "select KdKomponen,HargaSatuan,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuanJml & "'"
        Call msubRecFO(fRS2, fQuery2)
        While fRS2.EOF = False
            fKdKomponen = IIf(IsNull(fRS2("KdKomponen").value), "", fRS2("KdKomponen").value)
            fHargaPerKomponen = IIf(IsNull(fRS2("HargaSatuan").value), 0, fRS2("HargaSatuan").value)
            fJmlHutangPenjaminDB = IIf(IsNull(fRS2("JmlHutangPenjamin").value), 0, fRS2("JmlHutangPenjamin").value)
            fJmlTanggunganRSDB = IIf(IsNull(fRS2("JmlTanggunganRS").value), 0, fRS2("JmlTanggunganRS").value)
            fJmlPembebasanDB = IIf(IsNull(fRS2("JmlPembebasan").value), 0, fRS2("JmlPembebasan").value)
            If fJmlHrsDibayar = 0 Then
                If fJmlPembebasanDB = 0 And fJmlPembebasanDP <> 0 Then
                    fJmlPembebasanPerKomp = (CDec(fHargaPerKomponen) / CDec(fTotalBiaya)) * CDec(fJmlPembebasanDP)
                Else
                    fJmlPembebasanPerKomp = fJmlPembebasanDB
                End If
                fSisaTagihanPerKomp = 0
                If fTotalBiaya = 0 Then
                    fJmlHutangPerKomp = 0
                    fJmlTanggunganPerKomp = 0
                Else
                    If fJmlHutangPenjaminDB = 0 And fJmlHutangPenjaminDP <> 0 Then
                        fJmlHutangPerKomp = (CDec(fHargaPerKomponen) / CDec(fTotalBiaya)) * CDec(fJmlHutangPenjaminDP)
                    Else
                        fJmlHutangPerKomp = fJmlHutangPenjaminDB
                    End If
                    If fJmlTanggunganRSDB = 0 And fJmlTanggunganRSDP <> 0 Then
                        fJmlTanggunganPerKomp = (CDec(fHargaPerKomponen) / CDec(fTotalBiaya)) * CDec(fJmlTanggunganRSDP)
                    Else
                        fJmlTanggunganPerKomp = fJmlTanggunganRSDB
                    End If
                End If
                fJmlBayarPerKomp = 0
            Else
                If fTotalBiayaHrsDibayar = 0 Then
                    fTempJmlBayar = 0
                    fTempPembebasan = 0
                Else
                    fTempJmlBayar = (CDec(fJmlHrsDibayar) / CDec(fTotalBiayaHrsDibayar)) * CDec(fJmlBayar)
                    fTempPembebasan = (CDec(fJmlHrsDibayar) / CDec(fTotalBiayaHrsDibayar)) * CDec(fJmlDiscount)
                End If
                If fTotalHargaSatuan = 0 Then
                    fJmlPembebasanPerKomp = 0
                    fJmlBayarPerKomp = 0
                Else
                    If fJmlPembebasanDP = 0 Then
                        fJmlPembebasanPerKomp = (CDec(fHargaPerKomponen) / CDec(fTotalHargaSatuan)) * CDec(fTempPembebasan)
                    Else
                        If fJmlPembebasanDB = 0 And fJmlPembebasanDP <> 0 Then
                            fJmlPembebasanPerKomp = (CDec(fHargaPerKomponen) / CDec(fTotalBiaya)) * CDec(fJmlPembebasanDP)
                        Else
                            fJmlPembebasanPerKomp = fJmlPembebasanDB
                        End If
                    End If
                    fJmlBayarPerKomp = (CDec(fHargaPerKomponen) / CDec(fTotalHargaSatuan)) * CDec(fTempJmlBayar)
                End If

                If fSisaTagihan <> 0 Then
                    fSisaTagihanPerKomp = CDec(fHargaPerKomponen) - fJmlBayarPerKomp - fJmlPembebasanPerKomp
                Else
                    fSisaTagihanPerKomp = 0
                End If
                If fTotalBiaya = 0 Then
                    fJmlHutangPerKomp = 0
                    fJmlTanggunganPerKomp = 0
                Else
                    If fJmlHutangPenjaminDB = 0 And fJmlHutangPenjaminDP <> 0 Then
                        fJmlHutangPerKomp = (CDec(fHargaPerKomponen) / CDec(fTotalBiaya)) * CDec(fJmlHutangPenjaminDP)
                    Else
                        fJmlHutangPerKomp = fJmlHutangPenjaminDB
                    End If
                    If fJmlTanggunganRSDB = 0 And fJmlTanggunganRSDP <> 0 Then
                        fJmlTanggunganPerKomp = (CDec(fHargaPerKomponen) / CDec(fTotalBiaya)) * CDec(fJmlTanggunganRSDP)
                    Else
                        fJmlTanggunganPerKomp = fJmlTanggunganRSDB
                    End If
                End If
            End If
            Set fRS3 = Nothing
            fQuery3 = "insert into RekapKomponenBiayaPelayananOA values" & _
            "('" & fNoBKM & "','" & fNoPendaftaran & "','" & fKdRuangan & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdBarang & "'" & _
            ",'" & fKdAsal & "'," & msubKonversiKomaTitik(CStr(fJmlBarang)) & "," & msubKonversiKomaTitik(CStr(fHargaPerKomponen)) & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fNoStruk & "'," & fNoLab_Rad & "," & fIdPegawai & ",'" & fSatuanJml & "'," & msubKonversiKomaTitik(CStr(fJmlBayarPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKomp)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanPerKomp)) & ", '" & fKdDetailJenisJasaPelayanan & "'," & fKdPaket & "," & fNoResep & ",'" & fKdPelayananRS & "','" & fKdSubInstalasi & "')"
            Call msubRecFO(fRS3, fQuery3)
            Call f_AMRekapitulasiJasaBPOAForRemunerasiFV(fNoStruk, fNoBKM, fNoPendaftaran, fKdRuangan, fKdBarang, fKdAsal, fTglPelayanan, fSatuanJml, fKdKomponen, fJmlBarang, fHargaPerKomponen, fJmlBayarPerKomp, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, "A")
            If fKdKomponen = "13" Then
                Call f_AddRekapKomponenBPRemunerasiOA(fNoBKM, fNoPendaftaran, fKdRuangan, fKdKelas, fKdKomponen, fKdBarang, fKdAsal, fJmlBarang, fHargaPerKomponen, fTglPelayanan, fNoStruk, fNoLab_Rad, fIdPegawai, fSatuanJml, fJmlBayarPerKomp, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, fKdDetailJenisJasaPelayanan, fKdPaket, fNoResep, fKdPelayananRS, fKdSubInstalasi)
            End If
            fRS2.MoveNext
        Wend
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Add_RekapKomponenBiayaPelayananOAClaim
Public Function f_AddRekapKomponenBiayaPelayananOAClaim(fNoBKM As String, fNoBKMSebelumnya As String, fNoStruk As String, fJmlBayar As Currency)
    Dim fNoPendaftaran As String
    Dim fKdRuangan As String
    Dim fKdKelas As String
    Dim fKdBarang As String
    Dim fKdAsal As String
    Dim fJmlBarang As Double
    Dim fHargaSatuan As Currency
    Dim fTglPelayanan As Date
    Dim fNoLab_Rad As Variant
    Dim fIdPegawai As Variant
    Dim fSatuanJml As String
    Dim fJmlPembebasanPerBrg As Currency
    Dim fJmlBayarPerBrg As Currency
    Dim fJmlHutangPerBrg As Currency
    Dim fJmlTanggunganPerBrg As Currency
    Dim fSisaTagihanPerBrg As Currency
    Dim fKdKomponen As String
    Dim fJmlBayarL As Currency
    Dim fJmlHutangPenjaminL As Currency
    Dim fJmlTanggunganRSL As Currency
    Dim fSisaTagihanL As Currency
    Dim fJmlPembebasanL As Currency
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fIdPegawai2 As Variant
    Dim fJmlBayarPerKomp As Currency
    Dim fJmlHutangPerKomp As Currency
    Dim fNoResep As Variant
    Dim fTotalHutangPenjamin As Currency
    Dim fKdPaket As Variant
    Dim fKdSubInstalasi As String
    Dim fKdPelayananRS As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    fKdPelayananRS = "000001"
    Set fRS = Nothing
    fQuery = "select KdRuangan,KdKomponen,KdBarang,KdAsal,TglPelayanan,NoStruk,SatuanJml,NoPendaftaran,KdKelas,JmlBarang,HargaSatuan,NoLab_Rad,IdPegawai,JmlBayar,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,SisaTagihan,KdDetailJenisJasaPelayanan,KdPaket,NoResep,KdSubInstalasi from RekapKomponenBiayaPelayananOA where NoBKM='" & fNoBKMSebelumnya & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        fKdBarang = IIf(IsNull(fRS("KdBarang").value), "", fRS("KdBarang").value)
        fKdAsal = IIf(IsNull(fRS("KdAsal").value), "", fRS("KdAsal").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fNoStruk = IIf(IsNull(fRS("NoStruk").value), "", fRS("NoStruk").value)
        fSatuanJml = IIf(IsNull(fRS("SatuanJml").value), "", fRS("SatuanJml").value)
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
        fJmlBarang = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang").value)
        fHargaSatuan = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
        fNoLab_Rad = IIf(IsNull(fRS("NoLab_Rad").value), "null", "'" & fRS("NoLab_Rad").value & "'")
        fIdPegawai = IIf(IsNull(fRS("IdPegawai").value), "null", "'" & fRS("IdPegawai").value & "'")
        fJmlBayarL = IIf(IsNull(fRS("JmlBayar").value), 0, fRS("JmlBayar").value)
        fJmlHutangPenjaminL = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRSL = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanL = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fSisaTagihanL = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value)
        fKdDetailJenisJasaPelayanan = IIf(IsNull(fRS("KdDetailJenisJasaPelayanan").value), "", fRS("KdDetailJenisJasaPelayanan").value)
        fKdPaket = IIf(IsNull(fRS("KdPaket").value), "null", "'" & fRS("KdPaket").value & "'")
        fNoResep = IIf(IsNull(fRS("NoResep").value), "null", "'" & fRS("NoResep").value & "'")
        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
        If fJmlHutangPenjaminL <> 0 Then
            Set fRS2 = Nothing
            fQuery2 = "select sum(JmlHutangPenjamin) as JmlHutangPenjaminSum from RekapKomponenBiayaPelayananOA where NoStruk='" & fNoStruk & "' and NoBKM='" & fNoBKMSebelumnya & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = False Then fTotalHutangPenjamin = IIf(IsNull(fRS2("JmlHutangPenjaminSum").value), 0, fRS2("JmlHutangPenjaminSum").value)
            fJmlBayarPerKomp = (CDec(fJmlHutangPenjaminL) / CDec(fTotalHutangPenjamin)) * CDec(fJmlBayar)
            fJmlHutangPerKomp = CDec(fJmlHutangPenjaminL) - fJmlBayarPerKomp
            Set fRS2 = Nothing
            fQuery2 = "insert into RekapKomponenBiayaPelayananOA values('" & fNoBKM & "','" & fNoPendaftaran & "','" & fKdRuangan & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdBarang & "','" & fKdAsal & "'," & msubKonversiKomaTitik(CStr(fJmlBarang)) & "," & msubKonversiKomaTitik(CStr(fHargaSatuan)) & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fNoStruk & "'," & fNoLab_Rad & "," & fIdPegawai & ",'" & fSatuanJml & "'," & msubKonversiKomaTitik(CStr(fJmlBayarPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSL)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanL)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanL)) & ",'" & fKdDetailJenisJasaPelayanan & "'," & fKdPaket & "," & fNoResep & ",'" & fKdPelayananRS & "','" & fKdSubInstalasi & "')"
            Call msubRecFO(fRS2, fQuery2)
            Call f_AMRekapitulasiJasaBPOAForRemunerasiFV(fNoStruk, fNoBKM, fNoPendaftaran, fKdRuangan, fKdBarang, fKdAsal, fTglPelayanan, fSatuanJml, fKdKomponen, fJmlBarang, fHargaSatuan, fJmlBayarPerKomp, fJmlHutangPerKomp, fJmlTanggunganRSL, fJmlPembebasanL, fSisaTagihanL, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, "A")
            If fKdKomponen = "13" Then
                Call f_AddRekapKomponenBPRemunerasiOA(fNoBKM, fNoPendaftaran, fKdRuangan, fKdKelas, fKdKomponen, fKdBarang, fKdAsal, fJmlBarang, fHargaSatuan, fTglPelayanan, fNoStruk, fNoLab_Rad, fIdPegawai, fSatuanJml, fJmlBayarPerKomp, fJmlHutangPerKomp, fJmlTanggunganRSL, fJmlPembebasanL, fSisaTagihanL, fKdDetailJenisJasaPelayanan, fKdPaket, fNoResep, fKdPelayananRS, fKdSubInstalasi)
            End If
        End If
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Add_RekapKomponenBiayaPelayananOAKredit
Public Function f_AddRekapKomponenBiayaPelayananOAKredit(fNoBKM As String, fNoBKMSebelumnya As String, fNoStruk As String, fJmlBayar As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency)
    Dim fNoPendaftaran As String
    Dim fKdRuangan As String
    Dim fKdKelas As String
    Dim fKdBarang As String
    Dim fKdAsal As String
    Dim fJmlBarang As Double
    Dim fHargaSatuan As Currency
    Dim fTglPelayanan As Date
    Dim fNoLab_Rad As Variant
    Dim fIdPegawai As Variant
    Dim fSatuanJml As String
    Dim fJmlPembebasanPerBrg As Currency
    Dim fJmlBayarPerBrg As Currency
    Dim fJmlHutangPerBrg As Currency
    Dim fJmlTanggunganPerBrg As Currency
    Dim fSisaTagihanPerBrg As Currency
    Dim fKdKomponen As String
    Dim fJmlBayarL As Currency
    Dim fJmlHutangPenjaminL As Currency
    Dim fJmlTanggunganRSL As Currency
    Dim fSisaTagihanL As Currency
    Dim fJmlPembebasanL As Currency
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fIdPegawai2 As Variant
    Dim fJmlPembebasanPerKomp As Currency
    Dim fJmlBayarPerKomp As Currency
    Dim fSisaTagihanPerKomp As Currency
    Dim fNoResep As Variant
    Dim fTotalSisaTagihan As Currency
    Dim fKdPaket As Variant
    Dim fKdSubInstalasi As String
    Dim fKdPelayananRS As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    fKdPelayananRS = "000001"
    Set fRS = Nothing
    fQuery = "select KdRuangan,KdKomponen,KdBarang,KdAsal,TglPelayanan,NoStruk,SatuanJml,NoPendaftaran,KdKelas,JmlBarang,HargaSatuan,NoLab_Rad,IdPegawai,JmlBayar,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,SisaTagihan,KdDetailJenisJasaPelayanan,KdPaket,NoResep,KdSubInstalasi from RekapKomponenBiayaPelayananOA where NoBKM='" & fNoBKMSebelumnya & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        fKdBarang = IIf(IsNull(fRS("KdBarang").value), "", fRS("KdBarang").value)
        fKdAsal = IIf(IsNull(fRS("KdAsal").value), "", fRS("KdAsal").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fNoStruk = IIf(IsNull(fRS("NoStruk").value), "", fRS("NoStruk").value)
        fSatuanJml = IIf(IsNull(fRS("SatuanJml").value), "", fRS("SatuanJml").value)
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
        fJmlBarang = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang ").value)
        fHargaSatuan = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
        fNoLab_Rad = IIf(IsNull(fRS("NoLab_Rad").value), "null", "'" & fRS("NoLab_Rad").value & "'")
        fIdPegawai1 = IIf(IsNull(fRS("IdPegawai").value), "null", "'" & fRS("IdPegawai").value & "'")

        fJmlBayarL = IIf(IsNull(fRS("JmlBayar").value), 0, fRS("JmlBayar").value)
        fJmlHutangPenjaminL = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRSL = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanL = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fSisaTagihanL = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value)
        fKdDetailJenisJasaPelayanan = IIf(IsNull(fRS("KdDetailJenisJasaPelayanan").value), "", fRS("KdDetailJenisJasaPelayanan").value)
        fKdPaket = IIf(IsNull(fRS("KdPaket").value), "null", "'" & fRS("KdPaket").value & "'")
        fNoResep = IIf(IsNull(fRS("NoResep").value), "null", "'" & fRS("NoResep").value & "'")
        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi ").value), "", fRS("KdSubInstalasi ").value)
        If fSisaTagihanL <> 0 Then
            Set fRS2 = Nothing
            fQuery2 = "select sum(SisaTagihan) as SisaTagihanSum from RekapKomponenBiayaPelayananOA where NoStruk='" & fNoStruk & "' and NoBKM='" & fNoBKMSebelumnya & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = False Then fTotalSisaTagihan = IIf(IsNull(fRS2("SisaTagihanSum").value), 0, fRS2("SisaTagihanSum").value)
            fJmlPembebasanPerKomp = (CDec(fSisaTagihanL) / CDec(fTotalSisaTagihan)) * CDec(fJmlPembebasan)
            fSisaTagihanPerKomp = (CDec(fSisaTagihanL) / CDec(fTotalSisaTagihan)) * CDec(fSisaTagihan)
            fJmlBayarPerKomp = (CDec(fSisaTagihanL) / CDec(fTotalSisaTagihan)) * CDec(fJmlBayar)
            Set fRS2 = Nothing
            fQuery2 = "insert into RekapKomponenBiayaPelayananOA values('" & fNoBKM & "','" & fNoPendaftaran & "','" & fKdRuangan & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdBarang & "','" & fKdAsal & "'," & msubKonversiKomaTitik(CStr(fJmlBarang)) & "," & msubKonversiKomaTitik(CStr(fHargaSatuan)) & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fNoStruk & "'," & fNoLab_Rad & "," & fIdPegawai & ",'" & fSatuanJml & "'," & msubKonversiKomaTitik(CStr(fJmlBayarPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminL)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSL)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKomp)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanPerKomp)) & ",'" & fKdDetailJenisJasaPelayanan & "'," & fKdPaket & "," & fNoResep & ",'" & fKdPelayananRS & "','" & fKdSubInstalasi & "')"
            Call msubRecFO(fRS2, fQuery2)
            Call f_AMRekapitulasiJasaBPOAForRemunerasiFV(fNoStruk, fNoBKM, fNoPendaftaran, fKdRuangan, fKdBarang, fKdAsal, fTglPelayanan, fSatuanJml, fKdKomponen, fJmlBarang, fHargaSatuan, fJmlBayarPerKomp, fJmlHutangPenjaminL, fJmlTanggunganRSL, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, "A")
            If fKdKomponen = "13" Then
                Call f_AddRekapKomponenBPRemunerasiOA(fNoBKM, fNoPendaftaran, fKdRuangan, fKdKelas, fKdKomponen, fKdBarang, fKdAsal, fJmlBarang, fHargaSatuan, fTglPelayanan, fNoStruk, fNoLab_Rad, fIdPegawai, fSatuanJml, fJmlBayarPerKomp, fJmlHutangPenjaminL, fJmlTanggunganRSL, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, fKdDetailJenisJasaPelayanan, fKdPaket, fNoResep, fKdPelayananRS, fKdSubInstalasi)
            End If
        End If
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Add_RekapKomponenBiayaPelayananTM
Public Function f_AddRekapKomponenBiayaPelayananTM(fNoBKM As String, fNoStruk As String, fTotalBiayaHrsDibayar As Currency, fJmlBayar As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency, fJmlDiscount As Currency)
    Dim fNoPendaftaran As String
    Dim fKdRuangan As String
    Dim fKdPelayananRS As String
    Dim fKdKelas As String
    Dim fJmlPelayanan As Integer
    Dim fTglPelayanan As Date
    Dim fNoLab_Rad As Variant
    Dim fIdPegawai As Variant
    Dim fKdKomponen As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fTarif As Currency
    Dim fJmlPembebasanPerKomp As Currency
    Dim fJmlBayarPerKomp As Currency
    Dim fJmlHutangPerKomp As Currency
    Dim fJmlTanggunganPerKomp As Currency
    Dim fSisaTagihanPerKomp As Currency
    Dim fTotal As Currency
    Dim fTarifCito As Currency
    Dim fTotalTarif As Currency
    Dim fKdKelasPenjaminDB As String
    Dim fTarifKelasPenjaminDB As Currency
    Dim fJmlHutangPenjaminDB As Currency
    Dim fJmlTanggunganRSDB As Currency
    Dim fJmlPembebasanDB As Currency
    Dim fSelisihTarifKelasPenjamin As Currency
    Dim fSelisihTarifDgnTanggungan As Currency
    Dim fJmlHrsDibayar As Currency
    Dim fTempJmlBayar As Currency
    Dim fTempSisaTagihan As Currency
    Dim fTempPembebasan As Currency
    Dim fIdPegawai2 As Variant
    Dim fTarifPenjamin As Currency
    Dim fKdJenisTarif As String
    Dim fKdPaket As Variant
    Dim fIdPenjamin As String
    Dim fIdPegawai3 As Variant
    Dim fKdSubInstalasi As String
    Dim fJmlHutangPenjaminKomp As Currency
    Dim fJmlTanggunganRSKomp As Currency
    Dim fJmlPembebasanKomp As Currency
    Dim fIdPemeriksa As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select IdPenjamin from StrukPelayananPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "", fRS("IdPenjamin").value)
    Set fRS = Nothing
    fQuery = "select KdRuangan,KdPelayananRS,TglPelayanan,NoPendaftaran,KdKelas,JmlPelayanan,NoLab_Rad,IdPegawai,Tarif,TarifCito,KdKelasPenjamin,TarifKelasPenjamin,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,IdPegawai2,KdJenisTarif,KdPaket,IdPegawai3,KdSubInstalasi from DetailBiayaPelayanan where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
        fJmlPelayanan = IIf(IsNull(fRS("JmlPelayanan").value), 0, fRS("JmlPelayanan").value)
        fNoLab_Rad = IIf(IsNull(fRS("NoLab_Rad").value), "null", "'" & fRS("NoLab_Rad").value & "'")
        fIdPegawai = IIf(IsNull(fRS("IdPegawai").value), "null", "'" & fRS("IdPegawai").value & "'")
        fTotal = IIf(IsNull(fRS("Tarif").value), 0, fRS("Tarif").value)
        fTarifCito = IIf(IsNull(fRS("TarifCito").value), 0, fRS("TarifCito").value)
        fKdKelasPenjaminDB = IIf(IsNull(fRS("KdKelasPenjamin").value), "", fRS("KdKelasPenjamin").value)
        fTarifKelasPenjaminDB = IIf(IsNull(fRS("TarifKelasPenjamin").value), 0, fRS("TarifKelasPenjamin").value)
        fJmlHutangPenjaminDB = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRSDB = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanDB = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fIdPegawai2 = IIf(IsNull(fRS("IdPegawai2").value), "null", "'" & fRS("IdPegawai2").value & "'")
        fKdJenisTarif = IIf(IsNull(fRS("KdJenisTarif").value), "", fRS("KdJenisTarif").value)
        fKdPaket = IIf(IsNull(fRS("KdPaket").value), "null", "'" & fRS("KdPaket").value & "'")
        fIdPegawai3 = IIf(IsNull(fRS("IdPegawai3").value), "null", "'" & fRS("IdPegawai3").value & "'")
        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
        fTotalTarif = (fTotal * fJmlPelayanan) + (fTarifCito * fJmlPelayanan)
        fSelisihTarifKelasPenjamin = fTotalTarif - (fTarifKelasPenjaminDB * fJmlPelayanan)
        If fSelisihTarifKelasPenjamin < 0 Then fSelisihTarifKelasPenjamin = 0
        fSelisihTarifDgnTanggungan = (fTarifKelasPenjaminDB * fJmlPelayanan) - (fJmlHutangPenjaminDB * fJmlPelayanan) - (fJmlTanggunganRSDB * fJmlPelayanan) - (fJmlPembebasanDB * fJmlPelayanan)
        If fSelisihTarifDgnTanggungan < 0 Then fSelisihTarifDgnTanggungan = 0
        fJmlHrsDibayar = fSelisihTarifKelasPenjamin + fSelisihTarifDgnTanggungan
        If fJmlHrsDibayar < 0 Then fJmlHrsDibayar = 0
        Set fRS2 = Nothing
        fQuery2 = "select KdDetailJenisJasaPelayanan from DetailKelasPelayanan where KdKelas='" & fKdKelas & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fKdDetailJenisJasaPelayanan = IIf(IsNull(fRS2("KdDetailJenisJasaPelayanan").value), "", fRS2("KdDetailJenisJasaPelayanan").value)
        Set fRS2 = Nothing
        fQuery2 = "select KdKomponen,Harga,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
        Call msubRecFO(fRS2, fQuery2)
        While fRS2.EOF = False
            fKdKomponen = IIf(IsNull(fRS2("KdKomponen").value), "", fRS2("KdKomponen").value)
            fTarif = IIf(IsNull(fRS2("Harga").value), 0, fRS2("Harga").value)
            fJmlHutangPenjaminKomp = IIf(IsNull(fRS2("JmlHutangPenjamin").value), 0, fRS2("JmlHutangPenjamin").value)
            fJmlTanggunganRSKomp = IIf(IsNull(fRS2("JmlTanggunganRS").value), 0, fRS2("JmlTanggunganRS").value)
            fJmlPembebasanKomp = IIf(IsNull(fRS2("JmlPembebasan").value), 0, fRS2("JmlPembebasan").value)
            If fJmlHrsDibayar = 0 Then
                If fJmlPembebasanKomp = 0 Then
                    fJmlPembebasanPerKomp = (CDec(fTarif) / CDec(fTarifKelasPenjaminDB)) * CDec(fJmlPembebasanDB)
                Else
                    fJmlPembebasanPerKomp = fJmlPembebasanKomp
                End If
                fSisaTagihanPerKomp = 0
                If fTarifKelasPenjaminDB = 0 Then
                    fJmlHutangPerKomp = 0
                    fJmlTanggunganPerKomp = 0
                Else
                    If fJmlHutangPenjaminKomp = 0 Then
                        fJmlHutangPerKomp = (CDec(fTarif) / CDec(fTarifKelasPenjaminDB)) * CDec(fJmlHutangPenjaminDB)
                    Else
                        fJmlHutangPerKomp = fJmlHutangPenjaminKomp
                    End If
                    If fJmlTanggunganRSKomp = 0 Then
                        fJmlTanggunganPerKomp = (CDec(fTarif) / CDec(fTarifKelasPenjaminDB)) * CDec(fJmlTanggunganRSDB)
                    Else
                        fJmlTanggunganPerKomp = fJmlTanggunganRSKomp
                    End If
                End If
                fJmlBayarPerKomp = 0
            Else
                If fTotalBiayaHrsDibayar = 0 Then
                    fTempJmlBayar = 0
                    fTempPembebasan = 0
                Else
                    fTempJmlBayar = (CDec(fJmlHrsDibayar) / CDec(fTotalBiayaHrsDibayar)) * CDec(fJmlBayar)
                    fTempPembebasan = (CDec(fJmlHrsDibayar) / CDec(fTotalBiayaHrsDibayar)) * CDec(fJmlDiscount)
                End If
                If fTotalTarif = 0 Then
                    fJmlPembebasanPerKomp = 0
                    fJmlBayarPerKomp = 0
                Else
                    If fJmlPembebasanKomp = 0 Then
                        If fJmlPembebasanDB = 0 Then
                            fJmlPembebasanPerKomp = (CDec(fTarif) / CDec(fTotalTarif)) * CDec(fTempPembebasan)
                        Else
                            fJmlPembebasanPerKomp = (CDec(fTarif) / CDec(fTarifKelasPenjaminDB)) * CDec(fJmlPembebasanDB)
                        End If
                    Else
                        fJmlPembebasanPerKomp = fJmlPembebasanKomp
                    End If
                    fJmlBayarPerKomp = (CDec((fTarif)) / CDec(fTotalTarif)) * CDec(fTempJmlBayar)
                End If
                If fSisaTagihan <> 0 Then
                    fSisaTagihanPerKomp = CDec(fTarif) - fJmlBayarPerKomp - fJmlPembebasanPerKomp
                Else
                    fSisaTagihanPerKomp = 0
                End If
                If fTarifKelasPenjaminDB = 0 Then
                    fJmlHutangPerKomp = 0
                    fJmlTanggunganPerKomp = 0
                Else
                    If fJmlHutangPenjaminKomp = 0 Then
                        fJmlHutangPerKomp = (CDec(fTarif) / CDec(fTarifKelasPenjaminDB)) * CDec(fJmlHutangPenjaminDB)
                    Else
                        fJmlHutangPerKomp = fJmlHutangPenjaminKomp
                    End If
                    If fJmlTanggunganRSKomp = 0 Then
                        fJmlTanggunganPerKomp = (CDec(fTarif) / CDec(fTarifKelasPenjaminDB)) * CDec(fJmlTanggunganRSDB)
                    Else
                        fJmlTanggunganPerKomp = fJmlTanggunganRSKomp
                    End If
                End If
            End If
            If fJmlPembebasanPerKomp = 0 Then fJmlPembebasanPerKomp = 0
            If fKdKomponen <> "04" And fKdKomponen <> "14" Then
                fIdPemeriksa = fIdPegawai
                Set fRS3 = Nothing
                fQuery3 = "insert into RekapKomponenBiayaPelayananTM values('" & fNoBKM & "','" & fNoPendaftaran & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKomponen & "','" & fKdKelas & "'," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & msubKonversiKomaTitik(CStr(fTarif)) & "," & fNoLab_Rad & "," & fIdPegawai & ",'" & fNoStruk & "','" & fKdDetailJenisJasaPelayanan & "'," & msubKonversiKomaTitik(CStr(fJmlBayarPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKomp)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanPerKomp)) & "," & fKdPaket & "," & fKdSubInstalasi & ")"
                Call msubRecFO(fRS3, fQuery3)
                Call f_AMRekapitulasiJasaBPTMForRemunerasiFV(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponen, fTglPelayanan, fJmlPelayanan, fTarif, fJmlBayarPerKomp, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, "A")
                Call f_AMRekapitulasiJasaBPDokterForRemunerasiFV(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponen, fTglPelayanan, fJmlPelayanan, fTarif, fJmlBayarPerKomp, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, fIdPegawai, "A")
            End If
            If fKdKomponen = "04" Then
                fIdPemeriksa = fIdPegawai2
                Set fRS3 = Nothing
                fQuery3 = "insert into RekapKomponenBiayaPelayananTM values('" & fNoBKM & "','" & fNoPendaftaran & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKomponen & "','" & fKdKelas & "'," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & msubKonversiKomaTitik(CStr(fTarif)) & "," & fNoLab_Rad & "," & fIdPegawai & ",'" & fNoStruk & "','" & fKdDetailJenisJasaPelayanan & "'," & msubKonversiKomaTitik(CStr(fJmlBayarPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKomp)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanPerKomp)) & "," & fKdPaket & "," & fKdSubInstalasi & ")"
                Call msubRecFO(fRS3, fQuery3)
                Call f_AMRekapitulasiJasaBPTMForRemunerasiFV(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponen, fTglPelayanan, fJmlPelayanan, fTarif, fJmlBayarPerKomp, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, "A")
                Call f_AMRekapitulasiJasaBPDokterForRemunerasiFV(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponen, fTglPelayanan, fJmlPelayanan, fTarif, fJmlBayarPerKomp, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, fIdPegawai2, "A")
            End If
            If fKdKomponen = "14" Then
                fIdPemeriksa = fIdPegawai3
                Set fRS3 = Nothing
                fQuery3 = "insert into RekapKomponenBiayaPelayananTM values('" & fNoBKM & "','" & fNoPendaftaran & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKomponen & "','" & fKdKelas & "'," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & msubKonversiKomaTitik(CStr(fTarif)) & "," & fNoLab_Rad & "," & fIdPegawai2 & ",'" & fNoStruk & "','" & fKdDetailJenisJasaPelayanan & "'," & msubKonversiKomaTitik(CStr(fJmlBayarPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKomp)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanPerKomp)) & "," & fKdPaket & "," & fKdSubInstalasi & ")"
                Call msubRecFO(fRS3, fQuery3)
                Call f_AMRekapitulasiJasaBPTMForRemunerasiFV(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponen, fTglPelayanan, fJmlPelayanan, fTarif, fJmlBayarPerKomp, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, "A")
                Call f_AMRekapitulasiJasaBPDokterForRemunerasiFV(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponen, fTglPelayanan, fJmlPelayanan, fTarif, fJmlBayarPerKomp, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, fIdPegawai3, "A")
            End If
            If fKdKomponen <> "01" And fKdKomponen <> "12" Then
                Call f_AddRekapKomponenBPRemunerasiTM(fNoBKM, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponen, fKdKelas, fJmlPelayanan, fTglPelayanan, fTarif, fNoLab_Rad, fIdPemeriksa, fNoStruk, fKdDetailJenisJasaPelayanan, fJmlBayarPerKomp, fJmlHutangPerKomp, fJmlTanggunganPerKomp, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, fKdPaket, fKdSubInstalasi)
            End If
            fRS2.MoveNext
        Wend
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Add_RekapKomponenBiayaPelayananTMClaim
Public Function f_AddRekapKomponenBiayaPelayananTMClaim(fNoBKM As String, fNoBKMSebelumnya As String, fNoStruk As String, fJmlBayar As Currency)
    Dim fNoPendaftaran As String
    Dim fKdRuangan As String
    Dim fKdPelayananRS As String
    Dim fKdKelas As String
    Dim fJmlPelayanan As Integer
    Dim fTglPelayanan As Date
    Dim fNoLab_Rad As Variant
    Dim fIdPegawai As Variant
    Dim fKdKomponen As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fTarif As Currency
    Dim fJmlBayarPerKomp As Currency
    Dim fJmlHutangPenjaminPerKomp As Currency
    Dim fJmlBayarL As Currency
    Dim fSisaTagihanL As Currency
    Dim fJmlPembebasanL As Currency
    Dim fIdPegawai2 As Variant
    Dim fTotalHutangPenjamin As Currency
    Dim fJmlHutangPenjaminL As Currency
    Dim fJmlTanggunganRSL As Currency
    Dim fKdPaket As Variant
    Dim fKdSubInstalasi As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdRuangan,KdPelayananRS,KdKomponen,TglPelayanan,NoStruk,NoPendaftaran,KdKelas,JmlPelayanan,Tarif,NoLab_Rad,IdPegawai,KdDetailJenisJasaPelayanan,JmlBayar,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,SisaTagihan,KdPaket,KdSubInstalasi from RekapKomponenBiayaPelayananTM where NoBKM='" & fNoBKMSebelumnya & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fNoStruk = IIf(IsNull(fRS("NoStruk").value), "", fRS("NoStruk").value)
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
        fJmlPelayanan = IIf(IsNull(fRS("JmlPelayanan").value), 0, fRS("JmlPelayanan").value)
        fTarif = IIf(IsNull(fRS("Tarif").value), 0, fRS("Tarif").value)
        fNoLab_Rad = IIf(IsNull(fRS("NoLab_Rad").value), "null", "'" & fRS("NoLab_Rad").value & "'")
        fIdPegawai = IIf(IsNull(fRS("IdPegawai").value), "null", "'" & fRS("IdPegawai").value & "'")

        fKdDetailJenisJasaPelayanan = IIf(IsNull(fRS("KdDetailJenisJasaPelayanan").value), "", fRS("KdDetailJenisJasaPelayanan").value)
        fJmlBayarL = IIf(IsNull(fRS("JmlBayar").value), 0, fRS("JmlBayar").value)
        fJmlHutangPenjaminL = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRSL = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanL = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fSisaTagihanL = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value)
        fKdPaket = IIf(IsNull(fRS("KdPaket").value), "null", "'" & fRS("KdPaket").value & "'")
        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
        If fJmlHutangPenjaminL <> 0 Then
            Set fRS2 = Nothing
            fQuery2 = "select sum(JmlHutangPenjamin) as JmlHutangPenjaminSum from RekapKomponenBiayaPelayananTM where NoStruk='" & fNoStruk & "' and NoBKM='" & fNoBKMSebelumnya & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = False Then fTotalHutangPenjamin = IIf(IsNull(fRS2("JmlHutangPenjaminSum").value), "", fRS2("JmlHutangPenjaminSum").value)
            fJmlBayarPerKomp = (CDec(fJmlHutangPenjaminL) / CDec(fTotalHutangPenjamin)) * CDec(fJmlBayar)
            fJmlHutangPenjaminPerKomp = CDec(fJmlHutangPenjaminL) - fJmlBayarPerKomp
            Set fRS2 = Nothing
            fQuery2 = "insert into RekapKomponenBiayaPelayananTM values('" & fNoBKM & "','" & fNoPendaftaran & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKomponen & "','" & fKdKelas & "'," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & msubKonversiKomaTitik(CStr(fTarif)) & "," & fNoLab_Rad & "," & fIdPegawai & ",'" & fNoStruk & "','" & fKdDetailJenisJasaPelayanan & "'," & msubKonversiKomaTitik(CStr(fJmlBayarPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSL)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanL)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanL)) & "," & fKdPaket & ",'" & fKdSubInstalasi & "')"
            Call msubRecFO(fRS2, fQuery2)
            Call f_AMRekapitulasiJasaBPTMForRemunerasiFV(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponen, fTglPelayanan, fJmlPelayanan, fTarif, fJmlBayarPerKomp, fJmlHutangPenjaminPerKomp, fJmlTanggunganRSL, fJmlPembebasanL, fSisaTagihanL, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, "A")
            Call f_AMRekapitulasiJasaBPDokterForRemunerasiFV(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponen, fTglPelayanan, fJmlPelayanan, fTarif, fJmlBayarPerKomp, fJmlHutangPenjaminPerKomp, fJmlTanggunganRSL, fJmlPembebasanL, fSisaTagihanL, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, fIdPegawai, "A")
            If fKdKomponen <> "01" And fKdKomponen <> "12" Then
                Call f_AddRekapKomponenBPRemunerasiTM(fNoBKM, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponen, fKdKelas, fJmlPelayanan, fTglPelayanan, fTarif, fNoLab_Rad, fIdPegawai, fNoStruk, fKdDetailJenisJasaPelayanan, fJmlBayarPerKomp, fJmlHutangPenjaminPerKomp, fJmlTanggunganRSL, fJmlPembebasanL, fSisaTagihanL, fKdPaket, fKdSubInstalasi)
            End If
        End If
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Add_RekapKomponenBiayaPelayananTMKredit
Public Function f_AddRekapKomponenBiayaPelayananTMKredit(fNoBKM As String, fNoBKMSebelumnya As String, fNoStruk As String, fJmlBayar As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency)
    Dim fNoPendaftaran As String
    Dim fKdRuangan As String
    Dim fKdPelayananRS As String
    Dim fKdKelas As String
    Dim fJmlPelayanan As Integer
    Dim fTglPelayanan As Date
    Dim fNoLab_Rad As Variant
    Dim fIdPegawai As Variant
    Dim fKdKomponen As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fTarif As Currency
    Dim fJmlPembebasanPerKomp As Currency
    Dim fJmlBayarPerKomp As Currency
    Dim fSisaTagihanPerKomp As Currency
    Dim fJmlBayarL As Currency
    Dim fJmlHutangPenjaminL As Currency
    Dim fJmlTanggunganRSL As Currency
    Dim fJmlPembebasanL As Currency
    Dim fSisaTagihanL As Currency
    Dim fIdPegawai2 As Variant
    Dim fTotalSisaTagihan As Currency
    Dim fKdPaket As Variant
    Dim fKdSubInstalasi As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdRuangan,KdPelayananRS,KdKomponen,TglPelayanan,NoStruk,NoPendaftaran,KdKelas,JmlPelayanan,Tarif,NoLab_Rad,IdPegawai,KdDetailJenisJasaPelayanan,JmlBayar,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,SisaTagihan,KdPaket,KdSubInstalasi from RekapKomponenBiayaPelayananTM where NoBKM='" & fNoBKMSebelumnya & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fNoStruk = IIf(IsNull(fRS("NoStruk").value), "", fRS("NoStruk").value)
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
        fJmlPelayanan = IIf(IsNull(fRS("JmlPelayanan").value), 0, fRS("JmlPelayanan").value)
        fTarif = IIf(IsNull(fRS("Tarif").value), 0, fRS("Tarif").value)
        fNoLab_Rad = IIf(IsNull(fRS("NoLab_Rad").value), "null", "'" & fRS("NoLab_Rad").value & "'")
        fIdPegawai = IIf(IsNull(fRS("IdPegawai").value), "null", "'" & fRS("IdPegawai").value & "'")
        fKdDetailJenisJasaPelayanan = IIf(IsNull(fRS("KdDetailJenisJasaPelayanan").value), "", fRS("KdDetailJenisJasaPelayanan").value)
        fJmlBayarL = IIf(IsNull(fRS("JmlBayar").value), 0, fRS("JmlBayar").value)
        fJmlHutangPenjaminL = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRSL = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanL = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fSisaTagihanL = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value)
        fKdPaket = IIf(IsNull(fRS("KdPaket").value), "null", "'" & fRS("KdPaket").value & "'")
        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
        If fSisaTagihanL <> 0 Then
            Set fRS2 = Nothing
            fQuery2 = "select sum(SisaTagihan) as SisaTagihanSum from RekapKomponenBiayaPelayananTM where NoStruk='" & fNoStruk & "' and NoBKM='" & fNoBKMSebelumnya & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS.EOF = False Then fTotalSisaTagihan = IIf(IsNull(fRS2("SisaTagihanSum").value), 0, fRS2("SisaTagihanSum").value)
            fJmlPembebasanPerKomp = (CDec(fSisaTagihanL) / CDec(fTotalSisaTagihan)) * CDec(fJmlPembebasan)
            fSisaTagihanPerKomp = (CDec(fSisaTagihanL) / CDec(fTotalSisaTagihan)) * CDec(fSisaTagihan)
            fJmlBayarPerKomp = (CDec(fSisaTagihanL) / CDec(fTotalSisaTagihan)) * CDec(fJmlBayar)
            Set fRS2 = Nothing
            fQuery2 = "insert into RekapKomponenBiayaPelayananTM values('" & fNoBKM & "','" & fNoPendaftaran & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKomponen & "','" & fKdKelas & "'," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & msubKonversiKomaTitik(CStr(fTarif)) & "," & fNoLab_Rad & "," & fIdPegawai & ",'" & fNoStruk & "','" & fKdDetailJenisJasaPelayanan & "'," & msubKonversiKomaTitik(CStr(fJmlBayarPerKomp)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminL)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSL)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanPerKomp)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanPerKomp)) & "," & fKdPaket & ",'" & fKdSubInstalasi & "')"
            Call msubRecFO(fRS2, fQuery2)
            Call f_AMRekapitulasiJasaBPTMForRemunerasiFV(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponen, fTglPelayanan, fJmlPelayanan, fTarif, fJmlBayarPerKomp, fJmlHutangPenjaminL, fJmlTanggunganRSL, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, "A")
            Call f_AMRekapitulasiJasaBPDokterForRemunerasiFV(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponen, fTglPelayanan, fJmlPelayanan, fTarif, fJmlBayarPerKomp, fJmlHutangPenjaminL, fJmlTanggunganRSL, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, fIdPegawai, "A")
            If fKdKomponen <> "01" And fKdKomponen <> "12" Then
                Call f_AddRekapKomponenBPRemunerasiTM(fNoBKM, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponen, fKdKelas, fJmlPelayanan, fTglPelayanan, fTarif, fNoLab_Rad, fIdPegawai, fNoStruk, fKdDetailJenisJasaPelayanan, fJmlBayarPerKomp, fJmlHutangPenjaminL, fJmlTanggunganRSL, fJmlPembebasanPerKomp, fSisaTagihanPerKomp, fKdPaket, fKdSubInstalasi)
            End If
        End If
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Delete_RekapKomponenBiayaPelayananApotik
Public Function f_DeleteRekapKomponenBiayaPelayananApotik(fNoBKM As String, fNoStruk As String, fStatus As String)
    'fStatus: M=Minus
    Dim fKdRuangan As String
    Dim fKdKomponen As String
    Dim fKdAsal As String
    Dim fJmlBrg As Double
    Dim fTarif As Currency
    Dim fJmlBayar As Currency
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fJmlPembebasan As Currency
    Dim fSisaTagihan As Currency
    Dim fKdBarang As String
    Dim fSatuanJml As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdRuangan,KdBarang,KdAsal,SatuanJml,KdKomponen,JmlBarang,HargaSatuan,JmlBayar,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,SisaTagihan from RekapKomponenBiayaPelayananApotik where NoBKM='" & fNoBKM & "' and NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdBarang = IIf(IsNull(fRS("KdBarang").value), "", fRS("KdBarang").value)
        fKdAsal = IIf(IsNull(fRS("KdAsal").value), "", fRS("KdAsal").value)
        fSatuanJml = IIf(IsNull(fRS("SatuanJml").value), "", fRS("SatuanJml").value)
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        fJmlBrg = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang").value)
        fTarif = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
        fJmlBayar = IIf(IsNull(fRS("JmlBayar").value), 0, fRS("JmlBayar").value)
        fJmlHutangPenjamin = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRS = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasan = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fSisaTagihan = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value)
        Call f_AMRekapitulasiJasaBPApotik(fNoStruk, fNoBKM, fKdRuangan, fKdBarang, fKdAsal, fSatuanJml, fKdKomponen, fJmlBrg, fTarif, fJmlBayar, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, fSisaTagihan, "M")
        fRS.MoveNext
    Wend
    Set fRS = Nothing
    fQuery = "delete from RekapKomponenBiayaPelayananApotik where NoBKM='" & fNoBKM & "' and NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
End Function

'Konversi dari SP: Delete_RekapKomponenBiayaPelayananOA
Public Function f_DeleteRekapKomponenBiayaPelayananOA(fNoBKM As String, fNoStruk As String, fStatus As String)
    'fStatus: M=Minus
    Dim fNoPendaftaran As String
    Dim fKdRuangan As String
    Dim fKdKomponen As String
    Dim fKdAsal As String
    Dim fJmlBrg As Double
    Dim fTarif As Currency
    Dim fJmlBayar As Currency
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fJmlPembebasan As Currency
    Dim fSisaTagihan As Currency
    Dim fKdBarang As String
    Dim fTglPelayanan As Date
    Dim fSatuanJml As String
    Dim fNoLab_Rad As Variant
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fKdKelas As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select NoPendaftaran,KdRuangan,KdBarang,KdAsal,TglPelayanan,SatuanJml,KdKomponen,JmlBarang,HargaSatuan,JmlBayar,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,SisaTagihan,KdDetailJenisJasaPelayanan,KdKelas,NoLab_Rad from RekapKomponenBiayaPelayananOA where NoBKM='" & fNoBKM & "' and NoStruk='" & fNoStruk & "' and NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdBarang = IIf(IsNull(fRS("KdBarang").value), "", fRS("KdBarang").value)
        fKdAsal = IIf(IsNull(fRS("KdAsal").value), "", fRS("KdAsal").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fSatuanJml = IIf(IsNull(fRS("SatuanJml").value), "", fRS("SatuanJml").value)
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        fJmlBrg = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang").value)
        fTarif = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
        fJmlBayar = IIf(IsNull(fRS("JmlBayar").value), 0, fRS("JmlBayar").value)
        fJmlHutangPenjamin = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRS = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasan = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fSisaTagihan = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value)
        fKdDetailJenisJasaPelayanan = IIf(IsNull(fRS("KdDetailJenisJasaPelayanan").value), "", fRS("KdDetailJenisJasaPelayanan").value)
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
        fNoLab_Rad = fRS("NoLab_Rad").value
        Call f_AMRekapitulasiJasaBPOAForRemunerasiFV(fNoStruk, fNoBKM, fNoPendaftaran, fKdRuangan, fKdBarang, fKdAsal, fTglPelayanan, fSatuanJml, fKdKomponen, fJmlBrg, fTarif, fJmlBayar, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, fSisaTagihan, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, fStatus)
        fRS.MoveNext
    Wend
    Set fRS = Nothing
    fQuery = "delete from RekapKomponenBiayaPelayananOA where NoBKM='" & fNoBKM & "' and NoStruk='" & fNoStruk & "' and NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
End Function

'Konversi dari SP: Delete_RekapKomponenBiayaPelayananTM
Public Function f_DeleteRekapKomponenBiayaPelayananTM(fNoBKM As String, fNoStruk As String, fStatus As String)
    'fStatus : M=Minus
    Dim fNoPendaftaran As String
    Dim fTarif As Currency
    Dim fKdRuangan As String
    Dim fKdKomponen As String
    Dim fKdAsal As String
    Dim fJmlBayar As Currency
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fJmlPembebasan As Currency
    Dim fSisaTagihan As Currency
    Dim fKdPelayananRS As String
    Dim fTglPelayanan As Date
    Dim fNoLab_Rad As Variant
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fKdKelas As String
    Dim fIdPegawai As Variant
    Dim fJmlPelayanan As Integer

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select NoPendaftaran,KdRuangan,KdPelayananRS,KdKomponen,TglPelayanan,JmlPelayanan,Tarif,JmlBayar,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,SisaTagihan,KdDetailJenisJasaPelayanan,KdKelas,NoLab_Rad,IdPegawai from RekapKomponenBiayaPelayananTM where NoBKM='" & fNoBKM & "' and NoStruk='" & fNoStruk & "' and NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fJmlPelayanan = IIf(IsNull(fRS("JmlPelayanan").value), 0, fRS("JmlPelayanan").value)
        fTarif = IIf(IsNull(fRS("Tarif").value), 0, fRS("Tarif").value)
        fJmlBayar = IIf(IsNull(fRS("JmlBayar").value), 0, fRS("JmlBayar").value)
        fJmlHutangPenjamin = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRS = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasan = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fSisaTagihan = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value)
        fKdDetailJenisJasaPelayanan = IIf(IsNull(fRS("KdDetailJenisJasaPelayanan").value), "", fRS("KdDetailJenisJasaPelayanan").value)
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
        fNoLab_Rad = IIf(IsNull(fRS("NoLab_Rad").value), "null", "'" & fRS("NoLab_Rad").value & "'")
        fIdPegawai1 = IIf(IsNull(fRS("IdPegawai").value), "null", "'" & fRS("IdPegawai").value & "'")

        Call f_AMRekapitulasiJasaBPDokterForRemunerasiFV(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponen, fTglPelayanan, fJmlPelayanan, fTarif, fJmlBayar, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, fSisaTagihan, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, fIdPegawai, fStatus)
        Call f_AMRekapitulasiJasaBPTMForRemunerasiFV(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponen, fTglPelayanan, fJmlPelayanan, fTarif, fJmlBayar, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, fSisaTagihan, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, fStatus)
        fRS.MoveNext
    Wend
    Set fRS = Nothing
    fQuery = "delete from RekapKomponenBiayaPelayananTM where NoBKM='" & fNoBKM & "' and NoStruk='" & fNoStruk & "' and NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
End Function

'Konversi dari SP: Delete_RekapKomponenBPRemunerasiApotik
Public Function f_DeleteRekapKomponenBPRemunerasiApotik(fNoBKM As String, fNoStruk As String, fStatus As String)
    'fStatus: M=Minus
    Dim fTarif As Currency
    Dim fKdRuangan As String
    Dim fKdKomponenR As String
    Dim fKdDetailKomponenR As String
    Dim fKdAsal As String
    Dim fJmlBrg As Double
    Dim fJmlBayarPerKompR As Currency
    Dim fJmlHutangPerKompR As Currency
    Dim fJmlTanggunganPerKompR As Currency
    Dim fJmlPembebasanPerKompR As Currency
    Dim fSisaTagihanPerKompR As Currency
    Dim fKdBarang As String
    Dim fSatuanJml As String
    Dim fJmlPelayanan As Integer
    Dim fKdPelayananRS As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdRuangan,KdBarang,KdAsal,SatuanJml,KdKomponenR,KdDetailKomponenR,JmlBarang,HargaSatuan,JmlBayar,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,SisaTagihan,KdPelayananRS from RekapKomponenBPRemunerasiApotik where NoBKM='" & fNoBKM & "' and NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdBarang = IIf(IsNull(fRS("KdBarang").value), "", fRS("KdBarang").value)
        fKdAsal = IIf(IsNull(fRS("KdAsal").value), "", fRS("KdAsal").value)
        fSatuanJml = IIf(IsNull(fRS("SatuanJml").value), "", fRS("SatuanJml").value)
        fKdKomponenR = IIf(IsNull(fRS("KdKomponenR").value), "", fRS("KdKomponenR").value)
        fKdDetailKomponenR = IIf(IsNull(fRS("KdDetailKomponenR").value), "", fRS("KdDetailKomponenR").value)
        fJmlBrg = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang").value)
        fTarif = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
        fJmlBayarPerKompR = IIf(IsNull(fRS("JmlBayar").value), 0, fRS("JmlBayar").value)
        fJmlHutangPerKompR = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganPerKompR = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanPerKompR = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fSisaTagihanPerKompR = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value)
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        Call f_AMRekapitulasiKomponenRemunerasiApotik(fNoStruk, fNoBKM, fKdRuangan, fKdBarang, fKdAsal, fSatuanJml, fKdPelayananRS, fKdKomponenR, fKdDetailKomponenR, fJmlBrg, fTarif, fJmlBayarPerKompR, fJmlHutangPerKompR, fJmlTanggunganPerKompR, fJmlPembebasanPerKompR, fSisaTagihanPerKompR, fStatus)
        fRS.MoveNext
    Wend
    Set fRS = Nothing
    fQuery = "delete from RekapKomponenBPRemunerasiApotik where NoBKM='" & fNoBKM & "' and NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
End Function

'Konversi dari SP: Delete_RekapKomponenBPRemunerasiOA
Public Function f_DeleteRekapKomponenBPRemunerasiOA(fNoBKM As String, fNoStruk As String, fStatus As String)
    'fStatus: M=Minus
    Dim fNoPendaftaran As String
    Dim fTarif As Currency
    Dim fKdRuangan As String
    Dim fKdKomponen As String
    Dim fKdKomponenR As String
    Dim fKdDetailKomponenR As String
    Dim fKdAsal As String
    Dim fJmlBrg As Double
    Dim fJmlBayarPerKompR As Currency
    Dim fJmlHutangPerKompR As Currency
    Dim fJmlTanggunganPerKompR As Currency
    Dim fJmlPembebasanPerKompR As Currency
    Dim fSisaTagihanPerKompR As Currency
    Dim fKdBarang As String
    Dim fTglPelayanan As Date
    Dim fSatuanJml As String
    Dim fNoLab_Rad As Variant
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fKdKelas As String
    Dim fIdPegawai As Variant
    Dim fJmlPelayanan As Integer
    Dim fKdSubInstalasi As String
    Dim fKdPelayananRS As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select NoPendaftaran,KdRuangan,KdBarang,KdAsal,TglPelayanan,SatuanJml,KdKomponenR,KdDetailKomponenR,JmlBarang,HargaSatuan,JmlBayar,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,SisaTagihan,KdDetailJenisJasaPelayanan,KdKelas,NoLab_Rad,KdSubInstalasi,KdPelayananRS from RekapKomponenBPRemunerasiOA where NoBKM='" & fNoBKM & "' and NoStruk='" & fNoStruk & "' and NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdBarang = IIf(IsNull(fRS("KdBarang").value), "", fRS("KdBarang").value)
        fKdAsal = IIf(IsNull(fRS("KdAsal").value), "", fRS("KdAsal").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fSatuanJml = IIf(IsNull(fRS("SatuanJml").value), "", fRS("SatuanJml").value)
        fKdKomponenR = IIf(IsNull(fRS("KdKomponenR").value), "", fRS("KdKomponenR").value)
        fKdDetailKomponenR = IIf(IsNull(fRS("KdDetailKomponenR").value), "", fRS("KdDetailKomponenR").value)
        fJmlBrg = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang").value)
        fTarif = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
        fJmlBayarPerKompR = IIf(IsNull(fRS("JmlBayar").value), 0, fRS("JmlBayar").value)
        fJmlHutangPerKompR = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganPerKompR = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanPerKompR = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fSisaTagihanPerKompR = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value)
        fKdDetailJenisJasaPelayanan = IIf(IsNull(fRS("KdDetailJenisJasaPelayanan").value), "", fRS("KdDetailJenisJasaPelayanan").value)
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
        fNoLab_Rad = fRS("NoLab_Rad").value
        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        Call f_AMRekapitulasiKomponenRemunerasiOATM(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponenR, fKdDetailKomponenR, fTglPelayanan, fJmlPelayanan, fTarif, fJmlBayarPerKompR, fJmlHutangPerKompR, fJmlTanggunganPerKompR, fJmlPembebasanPerKompR, fSisaTagihanPerKompR, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, fKdAsal, fKdSubInstalasi, "OA", fStatus)
        fRS.MoveNext
    Wend
    Set fRS = Nothing
    fQuery = "delete from RekapKomponenBPRemunerasiOA where NoBKM='" & fNoBKM & "' and NoStruk='" & fNoStruk & "' and NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
End Function

'Konversi dari SP: Delete_RekapKomponenBPRemunerasiTM
Public Function f_DeleteRekapKomponenBPRemunerasiTM(fNoBKM As String, fNoStruk As String, fStatus As String)
    'fStatus: M=Minus
    Dim fNoPendaftaran As String
    Dim fTarif As Currency
    Dim fKdRuangan As String
    Dim fKdKomponen As String
    Dim fKdKomponenR As String
    Dim fKdDetailKomponenR As String
    Dim fKdAsal As String
    Dim fJmlBayarPerKompR As Currency
    Dim fJmlHutangPerKompR As Currency
    Dim fJmlTanggunganPerKompR As Currency
    Dim fJmlPembebasanPerKompR As Currency
    Dim fSisaTagihanPerKompR As Currency
    Dim fKdPelayananRS As String
    Dim fTglPelayanan As Date
    Dim fNoLab_Rad As Variant
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fKdKelas As String
    Dim fIdPegawai As Variant
    Dim fJmlPelayanan As Integer
    Dim fKdSubInstalasi As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select NoPendaftaran,KdRuangan,KdPelayananRS,KdKomponen,TglPelayanan,JmlPelayanan,Tarif,JmlBayar,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,SisaTagihan,KdDetailJenisJasaPelayanan,KdKelas,NoLab_Rad,IdPegawai,KdSubInstalasi from RekapKomponenBPRemunerasiTM where NoBKM='" & fNoBKM & "' and NoStruk='" & fNoStruk & "' and NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fJmlPelayanan = IIf(IsNull(fRS("JmlPelayanan").value), 0, fRS("JmlPelayanan").value)
        fTarif = IIf(IsNull(fRS("Tarif").value), 0, fRS("Tarif").value)
        fJmlBayarPerKompR = IIf(IsNull(fRS("JmlBayar").value), 0, fRS("JmlBayar").value)
        fJmlHutangPerKompR = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganPerKompR = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanPerKompR = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fSisaTagihanPerKompR = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value)
        fKdDetailJenisJasaPelayanan = IIf(IsNull(fRS("KdDetailJenisJasaPelayanan").value), "", fRS("KdDetailJenisJasaPelayanan").value)
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
        fNoLab_Rad = IIf(IsNull(fRS("NoLab_Rad").value), "null", "'" & fRS("NoLab_Rad").value & "'")
        fIdPegawai = IIf(IsNull(fRS("IdPegawai").value), "null", "'" & fRS("IdPegawai").value & "'")

        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
        Call f_AMRekapitulasiKomponenRemunerasiOATM(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponenR, fKdDetailKomponenR, fTglPelayanan, fJmlPelayanan, fTarif, fJmlBayarPerKompR, fJmlHutangPerKompR, fJmlTanggunganPerKompR, fJmlPembebasanPerKompR, fSisaTagihanPerKompR, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, fKdAsal, fKdSubInstalasi, "TM", fStatus)
        Call f_AMRekapitulasiKomponenRemunerasiDokter(fNoBKM, fNoStruk, fNoPendaftaran, fKdRuangan, fKdPelayananRS, fKdKomponenR, fKdDetailKomponenR, fTglPelayanan, fIdPegawai, fJmlPelayanan, fTarif, fJmlBayarPerKompR, fJmlHutangPerKompR, fJmlTanggunganPerKompR, fJmlPembebasanPerKompR, fSisaTagihanPerKompR, fKdDetailJenisJasaPelayanan, fKdKelas, fNoLab_Rad, fKdAsal, fKdSubInstalasi, fStatus)
        fRS.MoveNext
    Wend
    Set fRS = Nothing
    fQuery = "delete from RekapKomponenBPRemunerasiTM where NoBKM='" & fNoBKM & "' and NoStruk='" & fNoStruk & "' and NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
End Function

'Konversi dari SP: LoopUpdateTarifPelayananOnBatalKeluarKamar
Public Function f_LoopUpdateTarifPelayananOnBatalKeluarKamar(fNoPendaftaran As String, fKdRuanganAkhir As String, fLamaJamDirawat As Double)
    'fLamaJamDirawat: Jumlah Jam Di Bawah 24
    Dim fKdPelayananRS As String
    Dim fTglPelayanan As Date
    Dim fKdRuangan As String
    Dim fTarif As Currency
    Dim fTarifCito As Currency
    Dim fRangeJamMin As Double
    Dim fRangeJamMax As Double
    Dim fPersentaseDiscount As Double
    Dim fKdKomponen As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdPelayananRS,KdRuangan,TglPelayanan,Tarif,TarifCito,RangeJamMin,RangeJamMax,PersentaseDiscount,KdKomponen from V_DaftarDiscountTarifOnKeluarKamar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fTarif = IIf(IsNull(fRS("Tarif").value), 0, fRS("Tarif").value)
        fTarifCito = IIf(IsNull(fRS("TarifCito").value), 0, fRS("TarifCito").value)
        fRangeJamMin = IIf(IsNull(fRS("RangeJamMin").value), 0, fRS("RangeJamMin").value)
        fRangeJamMax = IIf(IsNull(fRS("RangeJamMax").value), 0, fRS("RangeJamMax").value)
        fPersentaseDiscount = IIf(IsNull(fRS("PersentaseDiscount").value), 0, fRS("PersentaseDiscount").value)
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        If fLamaJamDirawat <= fRangeJamMax And fLamaJamDirawat >= fRangeJamMin Then
            Set fRS2 = Nothing
            fQuery2 = "update BiayaPelayanan set Tarif=cast(((Tarif*100)/fPersentaseDiscount) as decimal),TarifCito=cast(((TarifCito*100)/fPersentaseDiscount) as decimal) where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganAkhir & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update DetailBiayaPelayanan set Tarif=cast(((Tarif*100)/fPersentaseDiscount) as decimal),TarifCito=cast(((TarifCito*100)/fPersentaseDiscount) as decimal) where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganAkhir & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponen set Harga=cast(((Harga*100)/fPersentaseDiscount) as decimal) where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganAkhir & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponen set Harga=cast(((Harga*100)/fPersentaseDiscount) as decimal) where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganAkhir & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='12'"
            Call msubRecFO(fRS2, fQuery2)
        End If
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: LoopUpdateTarifPelayananOnKeluarKamar
Public Function f_LoopUpdateTarifPelayananOnKeluarKamar(fNoPendaftaran As String, fKdRuanganAkhir As String, fLamaJamDirawat As Double)
    'fLamaJamDirawat: Jumlah Jam Di Bawah 24
    Dim fKdPelayananRS As String
    Dim fTglPelayanan As Date
    Dim fKdRuangan As String
    Dim fTarif As Currency
    Dim fTarifCito As Currency
    Dim fRangeJamMin As Double
    Dim fRangeJamMax As Double
    Dim fPersentaseDiscount As Double
    Dim fKdKomponen As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdPelayananRS,KdRuangan,TglPelayanan,Tarif,TarifCito,RangeJamMin,RangeJamMax,PersentaseDiscount,KdKomponen from V_DaftarDiscountTarifOnKeluarKamar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fTarif = IIf(IsNull(fRS("Tarif").value), 0, fRS("Tarif").value)
        fTarifCito = IIf(IsNull(fRS("TarifCito").value), 0, fRS("TarifCito").value)
        fRangeJamMin = IIf(IsNull(fRS("RangeJamMin").value), 0, fRS("RangeJamMin").value)
        fRangeJamMax = IIf(IsNull(fRS("RangeJamMax").value), 0, fRS("RangeJamMax").value)
        fPersentaseDiscount = IIf(IsNull(fRS("PersentaseDiscount").value), 0, fRS("PersentaseDiscount").value)
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        If fLamaJamDirawat <= fRangeJamMax And fLamaJamDirawat >= fRangeJamMin Then
            Set fRS2 = Nothing
            fQuery2 = "update BiayaPelayanan set Tarif=cast((Tarif*fPersentaseDiscount/100) as decimal),TarifCito=cast((TarifCito*fPersentaseDiscount/100) as decimal) where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganAkhir & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update DetailBiayaPelayanan set Tarif=cast((Tarif*fPersentaseDiscount/100) as decimal),TarifCito=cast((TarifCito*fPersentaseDiscount/100) as decimal) where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganAkhir & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponen set Harga=cast((Harga*fPersentaseDiscount/100) as decimal) where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganAkhir & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponen set Harga=cast((Harga*fPersentaseDiscount/100) as decimal) where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganAkhir & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='12'"
            Call msubRecFO(fRS2, fQuery2)
        End If
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Loop_AddTransferBPOAToHutangPenjamin
Public Function f_LoopAddTransferBPOAToHutangPenjamin(fNoPendaftaran As String, fTglTransfer As Date, fKdRuangan As String, fIdUser As String)
    'fKdRuangan=KdRuangan Login; fIdUser=IdUser Login
    Dim fKdRuanganPelayanan As String
    Dim fTglPelayanan As Date
    Dim fKdBarang As String
    Dim fKdAsal As String
    Dim fSatuanJml As String
    Dim fTotalTarif As Currency
    Dim fJmlHutangPenjamin As Currency
    Dim fHargaSatuan As Currency
    Dim fTarifService As Currency
    Dim fSelisihTarif As Currency

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdRuangan,TglPelayanan,KdBarang,KdAsal,SatuanJml,HargaSatuan,TarifService from PemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdRuanganPelayanan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fKdBarang = IIf(IsNull(fRS("KdBarang").value), "", fRS("KdBarang").value)
        fKdAsal = IIf(IsNull(fRS("KdAsal").value), "", fRS("KdAsal").value)
        fSatuanJml = IIf(IsNull(fRS("SatuanJml").value), "", fRS("SatuanJml").value)
        fHargaSatuan = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
        fTarifService = IIf(IsNull(fRS("TarifService").value), 0, fRS("TarifService").value)
        fTotalTarif = fHargaSatuan + fTarifService
        Set fRS2 = Nothing
        fQuery2 = "select JmlHutangPenjamin from DetailPemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganPelayanan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and NoStruk is null"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fJmlHutangPenjamin = IIf(IsNull(fRS2("JmlHutangPenjamin").value), 0, fRS2("JmlHutangPenjamin").value)
        fSelisihTarif = fTotalTarif - fJmlHutangPenjamin
        If fJmlHutangPenjamin <> 0 Then
            Set fRS2 = Nothing
            fQuery2 = "update DetailPemakaianAlkes set JmlTanggunganRS=" & fSelisihTarif & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganPelayanan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenObatAlkes set JmlTanggunganRS=HargaSatuan - JmlHutangPenjamin where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganPelayanan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
        Else
            Set fRS2 = Nothing
            fQuery2 = "update DetailPemakaianAlkes set JmlHutangPenjamin=HargaSatuan + TarifService where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganPelayanan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenObatAlkes set JmlHutangPenjamin=HargaSatuan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganPelayanan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
        End If
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Loop_AddTransferBPTMToHutangPenjamin
Public Function f_LoopAddTransferBPTMToHutangPenjamin(fNoPendaftaran As String, fTglTransfer As Date, fKdRuangan As String, fIdUser As String)
    'fKdRuangan=KdRuangan Login; fIdUser=IdUser Login
    Dim fKdRuanganPelayanan As String
    Dim fTglPelayanan As Date
    Dim fKdPelayananRS As String
    Dim fTotalTarif As Currency
    Dim fJmlHutangPenjamin As Currency
    Dim fTarif As Currency
    Dim fTarifCito As Currency
    Dim fSelisihTarif As Currency

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdRuangan,TglPelayanan,KdPelayananRS,Tarif,TarifCito from BiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdRuanganPelayanan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        fTarif = IIf(IsNull(fRS("Tarif").value), 0, fRS("Tarif").value)
        fTarifCito = IIf(IsNull(fRS("TarifCito").value), 0, fRS("TarifCito").value)
        fTotalTarif = fTarif + fTarifCito
        Set fRS2 = Nothing
        fQuery2 = "select JmlHutangPenjamin from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganPelayanan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and NoStruk is null"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fJmlHutangPenjamin = IIf(IsNull(fRS2("JmlHutangPenjamin").value), 0, fRS2("JmlHutangPenjamin").value)
        fSelisihTarif = fTotalTarif - fJmlHutangPenjamin
        If fJmlHutangPenjamin <> 0 Then
            Set fRS2 = Nothing
            fQuery2 = "update DetailBiayaPelayanan set JmlTanggunganRS=" & msubKonversiKomaTitik(CStr(fSelisihTarif)) & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganPelayanan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponen set JmlTanggunganRS=Harga - JmlHutangPenjamin where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganPelayanan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
        Else
            Set fRS2 = Nothing
            fQuery2 = "update DetailBiayaPelayanan set JmlHutangPenjamin=Tarif + TarifCito where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganPelayanan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponen set JmlHutangPenjamin=Harga where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuanganPelayanan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
        End If
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Loop_AddDataPelayananPasienTMOAApotikPH
Public Function f_LoopAddDataPelayananPasienTMOAApotikPH(fNoClosing As String, fNoPendaftaran As String, fKdRuangan As String, fTglPelayanan As Date, fKdItem As String, fKdAsal As String, fSatuanJml As String, fNoLab_Rad As Variant, fJenis As String)
    'fNoLab_Rad: Allow Null; fJenis: TM=Tindakan Medis,OA=Obat Alkes, AP:Apotik
    Dim fKdKomponen As String
    Dim fKdKelas As String
    Dim fHarga As Currency
    Dim fIdPegawai As Variant
    Dim fKdJenisPegawai As String
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fJmlPembebasan As Currency
    Dim fKdRuanganAsal As String
    Dim fJmlPelayanan As Double
    Dim fNoResep As Variant

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    If UCase(fJenis) = "TM" Then
        Set fRS = Nothing
        fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "'," & fNoLab_Rad & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
        Call msubRecFO(fRS, fQuery)
        fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
        Set fRS = Nothing
        fQuery = "select KdKomponen,KdKelas,Harga,JmlPelayanan,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,IdPegawai from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdItem & "' and NoClosing is null"
        Call msubRecFO(fRS, fQuery)
        While fRS.EOF = False
            fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
            fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
            fHarga = IIf(IsNull(fRS("Harga").value), 0, fRS("Harga").value)
            fJmlPelayanan = IIf(IsNull(fRS("JmlPelayanan").value), 0, fRS("JmlPelayanan").value)
            fJmlHutangPenjamin = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
            fJmlTanggunganRS = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
            fJmlPembebasan = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
            fIdPegawai = IIf(IsNull(fRS("IdPegawai").value), "null", "'" & fRS("IdPegawai").value & "'")

            Call f_AMDataPelayananTMPasienPH(fNoPendaftaran, fKdItem, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, fKdKelas, "A")
            Set fRS2 = Nothing
            fQuery2 = "select KdJenisPegawai from DataPegawai where IdPegawai=" & fIdPegawai & ""
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = False Then fKdJenisPegawai = IIf(IsNull(fRS2("KdJenisPegawai").value), "", fRS2("KdJenisPegawai").value) Else fKdJenisPegawai = ""
            If fKdJenisPegawai = "001" Then
                Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdItem, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, fKdKelas, fIdPegawai, "A")
            End If
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponen set NoClosing='" & fNoClosing & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdItem & "' and KdKomponen='" & fKdKomponen & "' and NoClosing is null"
            Call msubRecFO(fRS2, fQuery2)
            fRS.MoveNext
        Wend
    End If
    If UCase(fJenis) = "OA" Then
        Set fRS = Nothing
        fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "'," & fNoLab_Rad & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','OA') as KdRuanganAsal"
        Call msubRecFO(fRS, fQuery)
        fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
        Set fRS = Nothing
        fQuery = "select KdKomponen,HargaSatuan,JmlBarang,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdItem & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and NoClosing is null"
        Call msubRecFO(fRS, fQuery)
        While fRS.EOF = False
            fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
            fHarga = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
            fJmlPelayanan = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang").value)
            fJmlHutangPenjamin = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
            fJmlTanggunganRS = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
            fJmlPembebasan = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
            Call f_AMDataPelayananOAPasienPH(fNoPendaftaran, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdItem, fKdAsal, fSatuanJml, fKdKomponen, fHarga, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, CInt(fJmlPelayanan), fJmlPelayanan, "A")
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenObatAlkes set NoClosing=fNoClosing where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdItem & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponen & "' and NoClosing is null"
            Call msubRecFO(fRS2, fQuery2)
            fRS.MoveNext
        Wend
    End If
    If UCase(fJenis) = "AP" Then
        Set fRS = Nothing
        fQuery = "select NoResep from StrukPelayananPasien where NoStruk='" & fNoPendaftaran & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fNoResep = fRS("NoResep").value
        If fNoResep <> "" Then
            Set fRS = Nothing
            fQuery = "select KdRuanganAsal from ResepObat where NoResep='" & fNoResep & "'"
            Call msubRecFO(fRS, fQuery)
            If fRS.EOF = False Then fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
        End If
        If fKdRuanganAsal = "" Then fKdRuanganAsal = fKdRuangan
        Set fRS = Nothing
        fQuery = "select KdKomponen,HargaSatuan,JmlBarang,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from TempHargaKomponenApotik where NoStruk='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdItem & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and NoClosing is null"
        Call msubRecFO(fRS, fQuery)
        While fRS.EOF = False
            fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
            fHarga = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
            fJmlPelayanan = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang").value)
            fJmlHutangPenjamin = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
            fJmlTanggunganRS = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
            fJmlPembebasan = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
            Call f_AMDataPelayananApotikPH(fNoPendaftaran, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdItem, fKdAsal, fSatuanJml, fKdKomponen, fHarga, CInt(fJmlPelayanan), fJmlPelayanan, "A")
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenApotik set NoClosing='" & fNoClosing & "' where NoStruk='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdItem & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponen & "' and NoClosing is null"
            Call msubRecFO(fRS2, fQuery2)
            fRS.MoveNext
        Wend
    End If
End Function

'Konversi dari SP: Update_BiayaPelayananOnUbahJenisPasien
Public Function f_UpdateBiayaPelayananOnUbahJenisPasien(fNoPendaftaran As String)
    Dim fKdRuangan As String
    Dim fKdPelayananRS As String
    Dim fKdKelas As String
    Dim fStatusCito As String
    Dim fTglPelayanan As Date
    Dim fKdJenisTarif As String
    Dim fTarifBaru As Currency
    Dim fTarifCitoBaru As Currency
    Dim fIdPegawai As Variant
    Dim fIdPegawai2 As Variant
    Dim fIdPegawai3 As Variant

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdJenisTarif from v_JenisTarifPasien where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisTarif = IIf(IsNull(fRS("KdJenisTarif").value), "", fRS("KdJenisTarif").value)
    Set fRS = Nothing
    fQuery = "select NoPendaftaran,KdRuangan,KdPelayananRS,TglPelayanan,KdKelas,StatusCITO,IdPegawai,IdPegawai2,IdPegawai3 from BiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
        fStatusCito = IIf(IsNull(fRS("StatusCITO").value), "", fRS("StatusCITO").value)
        fIdPegawai = IIf(IsNull(fRS("IdPegawai").value), "null", "'" & fRS("IdPegawai").value & "'")
        fIdPegawai2 = IIf(IsNull(fRS("IdPegawai2").value), "null", "'" & fRS("IdPegawai2").value & "'")
        fIdPegawai3 = IIf(IsNull(fRS("IdPegawai3").value), "null", "'" & fRS("IdPegawai3").value & "'")

        Set fRS2 = Nothing
        fQuery2 = "select dbo.FB_NewTakeTarifBPTM ('" & fNoPendaftaran & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdJenisTarif & "','" & fStatusCito & "', " & fIdPegawai & "," & fIdPegawai2 & "," & fIdPegawai3 & ", 'C') as TarifCito"
        Call msubRecFO(fRS2, fQuery2)
        fTarifCitoBaru = IIf(IsNull(fRS2("TarifCito").value), 0, fRS2("TarifCito").value)
        Set fRS2 = Nothing
        fQuery2 = "select dbo.FB_NewTakeTarifBPTM ('" & fNoPendaftaran & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdJenisTarif & "','" & fStatusCito & "', " & fIdPegawai & "," & fIdPegawai2 & "," & fIdPegawai3 & ", 'T') as Tarif"
        Call msubRecFO(fRS2, fQuery2)
        fTarifBaru = IIf(IsNull(fRS2("Tarif").value), 0, fRS2("Tarif").value)
        Set fRS2 = Nothing
        fQuery2 = "update BiayaPelayanan set TarifCito=" & fTarifCitoBaru & ",Tarif=" & fTarifBaru & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and NoStruk is null"
        Call msubRecFO(fRS2, fQuery2)
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Update_DataMorbiditasPasienRI
Public Function f_UpdateDataMorbiditasPasienRI(fNoCM As String, fKdKondisiPulang As String, fNoPendaftaran As String)
    Dim fTglPeriksa As Date
    Dim fKdRuangan As String
    Dim fKdSubInstalasi As String
    Dim fJmlOutPriaTemp As Integer
    Dim fJmlOutPria As Integer
    Dim fJmlOutWanitaTemp As Integer
    Dim fJmlOutWanita As Integer
    Dim fJmlOutHidupTemp As Integer
    Dim fJmlOutHidup As Integer
    Dim fJmlOutMatiTemp As Integer
    Dim fJmlOutMati As Integer
    Dim fJK As String
    Dim fKdDiagnosa As String
    Dim fKdKelompokPasien As String
    Dim fStatusKasus As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select JenisKelamin from Pasien where NoCM='" & fNoCM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fJK = IIf(IsNull(fRS("JenisKelamin").value), "", fRS("JenisKelamin").value)
    Set fRS = Nothing
    fQuery = "select KdKelompokPasien from PasienDaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
    Set fRS = Nothing
    fQuery = "select TglPeriksa,KdRuangan,KdSubInstalasi,KdDiagnosa,StatusKasus from PeriksaDiagnosa where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fTglPeriksa = IIf(IsNull(fRS("TglPeriksa").value), "", fRS("TglPeriksa").value)
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
        fKdDiagnosa = IIf(IsNull(fRS("KdDiagnosa").value), "", fRS("KdDiagnosa").value)
        fStatusKasus = IIf(IsNull(fRS("StatusKasus").value), "", fRS("StatusKasus").value)
        Set fRS2 = Nothing
        fQuery2 = "select KdDiagnosa from DataMorbiditasPasienRI where (KdSubInstalasi='" & fKdSubInstalasi & "' and KdRuangan='" & fKdRuangan & "' and KdDiagnosa='" & fKdDiagnosa & "' and StatusKasus='" & fStatusKasus & "' and KdKelompokPasien='" & fKdKelompokPasien & "') and (day(TglPeriksa)=day('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPeriksa)=month('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPeriksa)=year('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'))"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then
            Set fRS3 = Nothing
            fQuery3 = "select JmlPasienOutPria,JmlPasienOutWanita,JmlPasienOutHidup,JmlPasienOutMati from DataMorbiditasPasienRI where (KdSubInstalasi='" & fKdSubInstalasi & "' and KdRuangan='" & fKdRuangan & "' and KdDiagnosa='" & fKdDiagnosa & "' and StatusKasus='" & fStatusKasus & "' and KdKelompokPasien='" & fKdKelompokPasien & "') and (day(TglPeriksa)=day('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPeriksa)=month('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPeriksa)=year('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'))"
            Call msubRecFO(fRS3, fQuery3)
            If fRS3.EOF = False Then
                fJmlOutPriaTemp = IIf(IsNull(fRS3("JmlPasienOutPria").value), 0, fRS3("JmlPasienOutPria").value)
                fJmlOutWanitaTemp = IIf(IsNull(fRS3("JmlPasienOutWanita").value), 0, fRS3("JmlPasienOutWanita").value)
                fJmlOutHidupTemp = IIf(IsNull(fRS3("JmlPasienOutHidup").value), 0, fRS3("JmlPasienOutHidup").value)
                fJmlOutMatiTemp = IIf(IsNull(fRS3("JmlPasienOutMati").value), 0, fRS3("JmlPasienOutMati").value)
            End If
            If fJK = "L" Then
                fJmlOutPria = fJmlOutPriaTemp + 1
                fJmlOutWanita = fJmlOutWanitaTemp
            Else
                fJmlOutWanita = fJmlOutWanitaTemp + 1
                fJmlOutPria = fJmlOutPriaTemp
            End If
            If fKdKondisiPulang = "04" Or fKdKondisiPulang = "05" Then
                fJmlOutMati = fJmlOutMatiTemp + 1
                fJmlOutHidup = fJmlOutHidupTemp
            Else
                fJmlOutHidup = fJmlOutHidupTemp + 1
                fJmlOutMati = fJmlOutMatiTemp
            End If
            Set fRS3 = Nothing
            fQuery3 = "update DataMorbiditasPasienRI set JmlPasienOutPria=" & fJmlOutPria & ",JmlPasienOutWanita=" & fJmlOutWanita & ",JmlPasienOutHidup=" & fJmlOutHidup & ",JmlPasienOutMati=" & fJmlOutMati & " where (KdSubInstalasi='" & fKdSubInstalasi & "' and KdRuangan='" & fKdRuangan & "' and KdDiagnosa='" & fKdDiagnosa & "' and StatusKasus='" & fStatusKasus & "' and KdKelompokPasien='" & fKdKelompokPasien & "') and (day(TglPeriksa)=day('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPeriksa)=month('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPeriksa)=year('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'))"
            Call msubRecFO(fRS3, fQuery3)
        End If
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Update_DetailBiayaPelayanan4PasienNU
Public Function f_UpdateDetailBiayaPelayanan4PasienNU(fNoPendaftaran As String, fKdRuangan As String, fKode_Item As String, fTglPelayanan As Date, fKdAsal As String, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fJenis As String, fIdUser As String, fSatuan As String, fStatus As String)
    ' fKode_Item: KdPelayananRS atau KdBarang; fJenis: TM=untuk Pelayanan Tindakan Medis; OA=untuk Pelayanan Obat&Alkes; fStatus : T=Update Total aja; K=Update per Komponen
    Dim fJmlPembebasanPerKomp As Currency
    Dim fJmlHutangPerKomp As Currency
    Dim fJmlTanggunganPerKomp As Currency
    Dim fKdKomponen As String
    Dim fHarga As Currency

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    If UCase(fJenis) = "TM" Then
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKode_Item & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then
            Set fRS2 = Nothing
            fQuery2 = "update DetailBiayaPelayanan set JmlHutangPenjamin=" & fJmlHutangPenjamin & ",JmlTanggunganRS=" & fJmlTanggunganRS & ", JmlPembebasan=" & fJmlPembebasan & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKode_Item & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
            Call msubRecFO(fRS2, fQuery2)
        End If
        If UCase(fStatus) = "T" Then
            Set fRS = Nothing
            fQuery = "select KdKomponen,Harga from TempHargaKomponen where  NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKode_Item & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
            Call msubRecFO(fRS, fQuery)
            While fRS.EOF = False
                fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
                fHarga = IIf(IsNull(fRS("Harga").value), 0, fRS("Harga").value)
                If fJmlHutangPenjamin = 0 Then
                    fJmlHutangPerKomp = 0
                Else
                    fJmlHutangPerKomp = (CDec(fHarga) / CDec(fJmlHutangPenjamin)) * CDec(fJmlHutangPenjamin)
                End If
                If fJmlTanggunganRS = 0 Then
                    fJmlTanggunganPerKomp = 0
                Else
                    fJmlTanggunganPerKomp = (CDec(fHarga) / CDec(fJmlTanggunganRS)) * CDec(fJmlTanggunganRS)
                End If
                If fJmlPembebasan = 0 Then
                    fJmlPembebasanPerKomp = 0
                Else
                    fJmlPembebasanPerKomp = (CDec(fHarga) / CDec(fJmlPembebasan)) * CDec(fJmlPembebasan)
                End If
                Set fRS2 = Nothing
                fQuery2 = "update TempHargaKomponen set JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKode_Item & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
                Call msubRecFO(fRS2, fQuery2)
                fRS.MoveNext
            Wend
        End If
    Else
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from DetailPemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKode_Item & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdAsal='" & fKdAsal & "' and SatuanJml = '" & fSatuan & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then
            Set fRS2 = Nothing
            fQuery2 = "update DetailPemakaianAlkes set JmlHutangPenjamin=" & fJmlHutangPenjamin & ",JmlTanggunganRS=" & fJmlTanggunganRS & ", JmlPembebasan=" & fJmlPembebasan & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKode_Item & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdAsal='" & fKdAsal & "' and SatuanJml = '" & fSatuan & "'"
            Call msubRecFO(fRS2, fQuery2)
        End If
        Set fRS = Nothing
        fQuery = "select KdKomponen,HargaSatuan from TempHargaKomponenObatAlkes where  NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKode_Item & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdAsal='" & fKdAsal & "' and SatuanJml = '" & fSatuan & "'"
        Call msubRecFO(fRS, fQuery)
        While fRS.EOF = False
            fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
            fHarga = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
            If fJmlHutangPenjamin = 0 Then
                fJmlHutangPerKomp = 0
            Else
                fJmlHutangPerKomp = (CDec(fHarga) / CDec(fJmlHutangPenjamin)) * CDec(fJmlHutangPenjamin)
            End If
            If fJmlTanggunganRS = 0 Then
                fJmlTanggunganPerKomp = 0
            Else
                fJmlTanggunganPerKomp = (CDec(fHarga) / CDec(fJmlTanggunganRS)) * CDec(fJmlTanggunganRS)
            End If
            If fJmlPembebasan = 0 Then
                fJmlPembebasanPerKomp = 0
            Else
                fJmlPembebasanPerKomp = (CDec(fHarga) / CDec(fJmlPembebasan)) * CDec(fJmlPembebasan)
            End If
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenObatAlkes set JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKode_Item & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and SatuanJml = '" & fSatuan & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
            fRS.MoveNext
        Wend
    End If
End Function

'Konversi dari SP: AM_DataPelayananTMPasienOnUbahDokterPH
Public Function f_AMDataPelayananTMPasienOnUbahDokterPH(fNoPendaftaran As String, fKdRuangan As String, fIdDokterBaru As String, fTglMasuk As Date)
    Dim fKdPelayananRS As String
    Dim fKdKomponen As String
    Dim fKdKelas As String
    Dim fIdPegawai As Variant
    Dim fHarga As Currency
    Dim fKdJenisPegawai As String
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fJmlPembebasan As Currency

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdPelayananRS from BiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglMasuk, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        Set fRS2 = Nothing
        fQuery2 = "select KdKomponen,Harga,KdKelas,IdPegawai,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglMasuk, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and NoClosing is not null and NoStruk is null"
        Call msubRecFO(fRS2, fQuery2)
        While fRS2.EOF = False
            fKdKomponen = IIf(IsNull(fRS2("KdKomponen").value), "", fRS2("KdKomponen").value)
            fHarga = IIf(IsNull(fRS2("Harga").value), 0, fRS2("Harga").value)
            fKdKelas = IIf(IsNull(fRS2("KdKelas").value), "", fRS2("KdKelas").value)
            fIdPegawai = IIf(IsNull(fRS2("IdPegawai").value), "null", "'" & fRS2("IdPegawai").value & "'")
            fJmlHutangPenjamin = IIf(IsNull(fRS2("JmlHutangPenjamin").value), 0, fRS2("JmlHutangPenjamin").value)
            fJmlTanggunganRS = IIf(IsNull(fRS2("fJmlTanggunganRS").value), 0, fRS2("fJmlTanggunganRS").value)
            fJmlPembebasan = IIf(IsNull(fRS2("JmlPembebasan").value), 0, fRS2("JmlPembebasan").value)
            Set fRS3 = Nothing
            fQuery3 = "select KdJenisPegawai from DataPegawai where IdPegawai=" & fIdPegawai & ""
            Call msubRecFO(fRS3, fQuery3)
            If fRS3.EOF = False Then fKdJenisPegawai = IIf(IsNull(fRS3("KdJenisPegawai").value), "", fRS3("KdJenisPegawai").value)
            If fKdJenisPegawai = "001" Then
                Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglMasuk, fKdRuangan, fKdRuangan, fKdKomponen, fHarga, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, fKdKelas, fIdPegawai, "M")
                Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglMasuk, fKdRuangan, fKdRuangan, fKdKomponen, fHarga, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, fKdKelas, fIdPegawai, "A")
            End If
            fRS2.MoveNext
        Wend
    End If
End Function

'Konversi dari SP: AM_DataTransaksiBarangMedis
Public Function f_AMDataTransaksiBarangMedis(fTglTransaksi As Date, fKdRuangan As String, fKdBarang As String, fKdAsal As String, fJmlBarang As Double, fHargaNetto As Currency, fHargaJual As Currency, fDiscount As Currency, fStatusTransaksi As String, fStatus As String)
    'fTglTransaksi: TglTerima/TglPelayanan/TglKeluar; fStatusTransaksi: i=Barang Masuk; o=Barang Keluar; fStatus: A=Add & Ubah; M=Minus
    'dim fStokAwal as double

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdBarang from DataTransaksiBarangMedis where KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and (day(TglTransaksi)=day('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and month(TglTransaksi)=month('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and year(TglTransaksi)=year('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        Set fRS = Nothing
        fQuery = "select JmlStok from StokRuangan where KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fStokAwal = IIf(IsNull(fRS("JmlStok").value), 0, fRS("JmlStok").value)
        If fRS.EOF = True Then fStokAwal = 0
    Else
        Set fRS = Nothing
        fQuery = "select distinct JmlStokAwal from DataTransaksiBarangMedis where KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and (day(TglTransaksi)=day('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and month(TglTransaksi)=month('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and year(TglTransaksi)=year('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "'))"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fStokAwal = IIf(IsNull(fRS("JmlStokAwal").value), 0, fRS("JmlStokAwal").value)
        If fRS.EOF = True Then fStokAwal = 0
    End If

    Set fRS = Nothing
    fQuery = "select KdRuangan from DataTransaksiBarangMedis where KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and (datepart(hh, TglTransaksi)=datepart(hh, '" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and day(TglTransaksi)=day('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and month(TglTransaksi)=month('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and year(TglTransaksi)=year('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        If LCase(fStatusTransaksi) = "i" Then
            fQuery = "insert into DataTransaksiBarangMedis values('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuangan & "','" & fKdBarang & "','" & fKdAsal & "'," & fStokAwal & "," & msubKonversiKomaTitik(CStr(fJmlBarang)) & ",0," & msubKonversiKomaTitik(CStr(fJmlBarang)) & " * " & msubKonversiKomaTitik(CStr(fHargaNetto)) & ",0," & msubKonversiKomaTitik(CStr(fJmlBarang)) & " * " & msubKonversiKomaTitik(CStr(fDiscount)) & ",0,0,0,null)"
        Else
            fQuery = "insert into DataTransaksiBarangMedis values('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuangan & "','" & fKdBarang & "','" & fKdAsal & "'," & fStokAwal & ",0," & msubKonversiKomaTitik(CStr(fJmlBarang)) & "," & msubKonversiKomaTitik(CStr(fJmlBarang)) & " * " & msubKonversiKomaTitik(CStr(fHargaNetto)) & ",0,0," & msubKonversiKomaTitik(CStr(fJmlBarang)) & " * " & msubKonversiKomaTitik(CStr(fHargaNetto)) & "," & msubKonversiKomaTitik(CStr(fJmlBarang)) & " * " & msubKonversiKomaTitik(CStr(fHargaJual)) & "," & msubKonversiKomaTitik(CStr(fJmlBarang)) & " * " & msubKonversiKomaTitik(CStr(fDiscount)) & ",null)"
        End If
    Else
        If UCase(fStatus) = "A" Then
            If LCase(fStatusTransaksi) = "i" Then
                fQuery = "update DataTransaksiBarangMedis set JmlTerima=JmlTerima + " & fJmlBarang & ",TotalNettoi=TotalNettoi + (" & fJmlBarang & " * " & fHargaNetto & "),TotalDiscounti=TotalDiscounti + (" & fJmlBarang & " * " & fDiscount & ") where KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and (datepart(hh, TglTransaksi)=datepart(hh, '" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and day(TglTransaksi)=day('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and month(TglTransaksi)=month('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and year(TglTransaksi)=year('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "'))"
            Else
                fQuery = "update DataTransaksiBarangMedis set JmlKeluar=JmlKeluar + " & fJmlBarang & ",TotalNettoo=TotalNettoo + (" & fJmlBarang & " * " & fHargaNetto & "),TotalJualo=TotalJualo + (" & fJmlBarang & " * " & fHargaJual & "),TotalDiscounto=TotalDiscounto + (" & fJmlBarang & " * " & fDiscount & ") where KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and (datepart(hh, TglTransaksi)=datepart(hh, '" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and day(TglTransaksi)=day('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and month(TglTransaksi)=month('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and year(TglTransaksi)=year('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "'))"
            End If
        Else
            If LCase(fStatusTransaksi) = "i" Then
                fQuery = "update DataTransaksiBarangMedis set JmlTerima=JmlTerima - " & fJmlBarang & ",TotalNettoi=TotalNettoi - (" & fJmlBarang & " * " & fHargaNetto & "),TotalDiscounti=TotalDiscounti - (" & fJmlBarang & " * " & fDiscount & ") where KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and (datepart(hh, TglTransaksi)=datepart(hh, '" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and day(TglTransaksi)=day('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and month(TglTransaksi)=month('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and year(TglTransaksi)=year('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "'))"
            Else
                fQuery = "update DataTransaksiBarangMedis set JmlKeluar=JmlKeluar - " & fJmlBarang & ",TotalNettoo=TotalNettoo - (" & fJmlBarang & " * " & fHargaNetto & "),TotalJualo=TotalJualo - (" & fJmlBarang & " * " & fHargaJual & "),TotalDiscounto=TotalDiscounto - (" & fJmlBarang & " * " & fDiscount & ") where KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and (datepart(hh, TglTransaksi)=datepart(hh, '" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and day(TglTransaksi)=day('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and month(TglTransaksi)=month('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and year(TglTransaksi)=year('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "'))"
            End If
        End If
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
End Function

'Konversi dari SP: Add_ApotikJual
Public Function f_AddApotikJual(fKdBarang As String, fKdAsal As String, fKdRuangan As String, fSatuan As String, fJmlBrg As Double, fNoStruk As String, fHargaSatuan As Currency, fPPn As Currency, fDisc As Currency, fHargaBeli As Currency, fKdJenisObat As Variant, fJmlService As Integer, fTarifService As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fRke As Integer, fBiayaAdministrasi As Currency)
    'fSatuan: S (Standar), K (Kecil)
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String
    Dim fJmlBrgTemp As Double
    Dim fJmlJualTerkecil As Double
    Dim fJmlTerkecil As Double
    Dim fJmlStokRu As Double
    Dim fJmlBrgTempRu As Double
    Dim fJmlStokTerkecilRu As Double
    Dim fJmlModBrgTemp As Double
    Dim fJmlDivBrgTemp As Double
    Dim fJmlStokRuNow As Double
    Dim fJmlStokBrgTempNow As Double
    Dim fKdBrgTemp As String
    Dim ftempKdBarang As String
    Dim fTempJmlBrg As Double
    Dim fTotalJmlBrg As Double
    Dim fTglTransaksi As Date
    Dim fKdInstalasi As String
    Dim fNoStrukTemp As String
    Dim fNoResep As Variant

    Set fRS = Nothing
    fQuery = "select TglStruk from StrukPelayananPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTglTransaksi = IIf(IsNull(fRS("TglStruk").value), "", fRS("TglStruk").value)
    If fSatuan = "S" Then
        Set fRS = Nothing
        fQuery = "select JmlStok from StokRuangan where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fJmlStokRu = IIf(IsNull(fRS("JmlStok").value), 0, fRS("JmlStok").value)
        fJmlBrgTemp = fJmlStokRu - fJmlBrg
        GoTo SimpanS
    Else
        Set fRS = Nothing
        fQuery = "select JmlTerkecil,JmlJualTerkecil from MasterBarang where KdBarang='" & fKdBarang & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fJmlTerkecil = IIf(IsNull(fRS("JmlTerkecil").value), 0, fRS("JmlTerkecil").value) Else fJmlTerkecil = 0
        If fRS.EOF = False Then fJmlJualTerkecil = IIf(IsNull(fRS("JmlJualTerkecil").value), 0, fRS("JmlJualTerkecil").value) Else fJmlJualTerkecil = 0
        Set fRS = Nothing
        fQuery = "select JmlBarangTemp from JmlBarangTemp where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fJmlBrgTempRu = IIf(IsNull(fRS("JmlBarangTemp").value), 0, fRS("JmlBarangTemp").value) Else fJmlBrgTempRu = 0
        Set fRS = Nothing
        fQuery = "select JmlStok from StokRuangan where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fJmlStokRu = IIf(IsNull(fRS("JmlStok").value), 0, fRS("JmlStok").value) Else fJmlStokRu = 0
        fJmlBrgTemp = (fJmlBrg * fJmlJualTerkecil) + fJmlBrgTempRu
        fJmlStokTerkecilRu = fJmlStokRu * fJmlTerkecil
        If CInt(fJmlTerkecil) = 0 Then
            fJmlModBrgTemp = 0
        Else
            fJmlModBrgTemp = CInt(fJmlBrgTemp) Mod CInt(fJmlTerkecil)
        End If
        fJmlDivBrgTemp = fJmlBrgTemp / fJmlTerkecil
        fJmlStokRuNow = fJmlStokRu - fJmlDivBrgTemp
        fJmlStokBrgTempNow = fJmlModBrgTemp
        GoTo SimpanK
    End If
SimpanS:
    Set fRS = Nothing
    fQuery = "select KdBarang from ApotikJual where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "' and NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into ApotikJual values('" & fKdBarang & "','" & fKdAsal & "','" & fKdRuangan & "','" & fSatuan & "'," & msubKonversiKomaTitik(CStr(fJmlBrg)) & "," & msubKonversiKomaTitik(CStr(fHargaSatuan)) & "," & msubKonversiKomaTitik(CStr(fPPn)) & "," & msubKonversiKomaTitik(CStr(fDisc)) & ",'" & fNoStruk & "'," & msubKonversiKomaTitik(CStr(fHargaBeli)) & "," & IIf(fKdJenisObat = "", "null", "'" & fKdJenisObat & "'") & "," & msubKonversiKomaTitik(CStr(fJmlService)) & "," & msubKonversiKomaTitik(CStr(fTarifService)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjamin)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRS)) & "," & msubKonversiKomaTitik(CStr(fBiayaAdministrasi)) & ",0)"
    Else
        Set fRS2 = Nothing
        fQuery2 = "select JmlBarang from ApotikJual where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "' and NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS.EOF = False Then fTempJmlBrg = IIf(IsNull(fRS2("JmlBarang").value), 0, fRS2("JmlBarang").value)
        fTotalJmlBrg = fTempJmlBrg + fJmlBrg
        fQuery = "update ApotikJual set JmlBarang=" & msubKonversiKomaTitik(CStr(fTotalJmlBrg)) & " where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "' and NoStruk='" & fNoStruk & "'"
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
    Set fRS = Nothing
    fQuery = "update StokRuangan set JmlStok= JmlStok - " & msubKonversiKomaTitik(CStr(fJmlBrg)) & " where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "'"
    Call msubRecFO(fRS, fQuery)
    GoTo Selesai
SimpanK:
    Set fRS = Nothing
    fQuery = "insert into ApotikJual values('" & fKdBarang & "','" & fKdAsal & "','" & fKdRuangan & "','" & fSatuan & "'," & msubKonversiKomaTitik(CStr(fJmlBrg)) & "," & msubKonversiKomaTitik(CStr(fHargaSatuan)) & "," & msubKonversiKomaTitik(CStr(fPPn)) & "," & msubKonversiKomaTitik(CStr(fDisc)) & ",'" & fNoStruk & "'," & msubKonversiKomaTitik(CStr(fHargaBeli)) & "," & fKdJenisObat & "," & msubKonversiKomaTitik(CStr(fJmlService)) & "," & msubKonversiKomaTitik(CStr(fTarifService)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjamin)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRS)) & "," & msubKonversiKomaTitik(CStr(fBiayaAdministrasi)) & ",0)"
    Call msubRecFO(fRS, fQuery)
    Set fRS = Nothing
    fQuery = "update StokRuangan set JmlStok=" & msubKonversiKomaTitik(CStr(fJmlStokRuNow)) & " where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "'"
    Call msubRecFO(fRS, fQuery)
    Set fRS = Nothing
    fQuery = "select KdBarang from JmlBarangTemp where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into JmlBarangTemp values('" & fKdBarang & "','" & fKdAsal & "','" & fKdRuangan & "'," & msubKonversiKomaTitik(CStr(fJmlStokBrgTempNow)) & ")"
    Else
        fQuery = "update JmlBarangTemp set JmlBarangTemp=" & msubKonversiKomaTitik(CStr(fJmlStokBrgTempNow)) & " where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "'"
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
Selesai:
    Call f_AMDataTransaksiBarangMedis(fTglTransaksi, fKdRuangan, fKdBarang, fKdAsal, fJmlBrg, fHargaBeli, fHargaSatuan, fDisc, "o", "A")
    Call f_AddTempHargaKomponenApotik(fNoStruk, fKdRuangan, fKdBarang, fKdAsal, fSatuan, fHargaSatuan, fHargaBeli, fJmlBrg, fKdJenisObat, fJmlService, fTarifService, fBiayaAdministrasi)
    Set fRS = Nothing
    fQuery = "select NoResep from StrukPelayananPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fNoResep = fRS("NoResep").value Else fNoResep = ""
    If fNoResep <> "" Then
        Set fRS = Nothing
        fQuery = "SELECT NoResep FROM DetailResepObat WHERE (NoResep = " & fNoResep & ") AND (KdRuangan = '" & fKdRuangan & "') AND (KdBarang = '" & fKdBarang & "') AND (KdAsal = '" & fKdAsal & "') AND (SatuanJml = '" & fSatuan & "') AND (TglPelayanan = '" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "')"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            Set fRS2 = Nothing
            fQuery2 = "insert into DetailResepObat values('" & fNoResep & "','" & fKdRuangan & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuan & "','" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "'," & fRke & ")"
            Call msubRecFO(fRS2, fQuery2)
        End If
    End If
End Function

''Konversi dari SP: Add_BiayaPelayanan
Public Function f_AddBiayaPelayanan(fNoPendaftaran As String, fKdSubInstalasi As String, fKdRuangan As String, fKdPelayananRS As String, fKdKelas As String, fStatusCito As String, fTarif As Double, fJmlPelayanan As Integer, fTglPelayanan As Date, fNoLab_Rad As Variant, fIdPegawai As Variant, fStatusAPBD As String, fKdJenisTarif As String, fTarifCito As Integer, fIdUser As String, fIdPegawai2 As Variant)
    Dim fIdPenjamin As String
    Dim fKdKelasPenjamin As String
    Dim fKdKelompokPasien As String
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fJmlTanggunganRSL As Currency
    Dim fPersenTanggungan As Double
    Dim fPersenTanggunganRS As Double
    Dim fTotalTarif As Currency
    Dim fTarifKelasPenjamin As Currency
    Dim fTarifCitoKelasPenjamin As Currency
    Dim fPersenTarifCito As Double
    Dim fTarifCitoPenjamin As Currency
    Dim fTotalTarifPenjamin As Currency
    Dim fKdPaket As Variant
    Dim fTotalBiayaPaket As Currency
    Dim fTotalTanggunganPaket As Currency
    Dim fKdPaketL As String
    Dim fTarifKelasPenjaminL As Currency
    Dim fJmlHutangPenjaminL As Currency
    Dim fKdPelayananRSL As String
    Dim fTglPelayananL As Date
    Dim fKdInstalasi As String
    Dim fTglPelayananAdm As Date
    Dim fKdPelayananRSAdmin As String
    Dim fJmlHutangPenjaminPPT As Currency
    Dim fJmlPelayananTemp As Integer
    Dim fKdPaketTM As String
    Dim fKdPaketPaket As String
    Dim fSisaTagihanPasien As Currency
    Dim fSisaTagihanPasienL As Currency
    Dim fTarifAdmin As Currency
    Dim fKdRuanganAsal As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdPelayananRSAdmin from MasterDataPendukung"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdPelayananRSAdmin = IIf(IsNull(fRS("KdPelayananRSAdmin").value), "001001", fRS("KdPelayananRSAdmin").value) Else fKdPelayananRSAdmin = "001001"
    'Begin Setting Jumlah Biaya Administrasi Per Hari
    Set fRS = Nothing
    fQuery = "select sum(JmlPelayanan) as JmlPelayananSum from BiayaPelayanan where KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null and (day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') ) and KdPelayananRS<>'" & fKdPelayananRSAdmin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fJmlPelayananTemp = IIf(IsNull(fRS("JmlPelayananSum").value), 0, fRS("JmlPelayananSum").value) Else fJmlPelayananTemp = 0
    If fJmlPelayananTemp <= 5 Or fJmlPelayananTemp = 0 Then
        Set fRS = Nothing
        fQuery = "select min(TglPelayanan) as TglPelayananMin from BiayaPelayanan where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fTglPelayananAdm = IIf(IsNull(fRS("TglPelayananMin").value), "", fRS("TglPelayananMin").value) Else fTglPelayananAdm = ""
        If fTglPelayananAdm <> "" Then
            Set fRS2 = Nothing
            fQuery2 = "update BiayaPelayanan set JmlPelayanan=JmlPelayanan+" & fJmlPelayanan & " where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null and TglPelayanan='" & Format(fTglPelayananAdm, "yyyy/MM/dd HH:mm:ss") & "'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update DetailBiayaPelayanan set JmlPelayanan=JmlPelayanan+" & fJmlPelayanan & " where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null and TglPelayanan='" & Format(fTglPelayananAdm, "yyyy/MM/dd HH:mm:ss") & "'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponen set JmlPelayanan=JmlPelayanan+" & fJmlPelayanan & " where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null and TglPelayanan='" & Format(fTglPelayananAdm, "yyyy/MM/dd HH:mm:ss") & "'"
            Call msubRecFO(fRS2, fQuery2)
            Call f_AddTempHargaKomponen(fNoPendaftaran, fKdRuangan, fTglPelayananAdm, fKdPelayananRSAdmin, fKdKelas, fKdJenisTarif, CDbl(fTarifCito), fJmlPelayanan, fStatusCito, CStr(fIdPegawai))
        End If
    Else
        Set fRS2 = Nothing
        fQuery2 = "select dbo.FB_NewTakeTarifBPTM ('" & fNoPendaftaran & "', '" & fKdPelayananRSAdmin & "', '" & fKdKelas & "', '" & fKdJenisTarif & "', '0', " & fIdPegawai & ", null, null, 'T') as TarifAdmin"
        Call msubRecFO(fRS2, fQuery2)
        If fRS.EOF = False Then fTarifAdmin = IIf(IsNull(fRS("TarifAdmin").value), 0, fRS("TarifAdmin").value) Else fTarifAdmin = 0
        Call f_AddBiayaPelayananAdmin(fNoPendaftaran, fKdSubInstalasi, fKdRuangan, fKdPelayananRSAdmin, fKdKelas, "0", CDbl(fTarifAdmin), 1, fTglPelayanan, fNoLab_Rad, fIdPegawai, "01", fKdJenisTarif, 0, CStr(fIdPegawai), Null)
    End If
    'End Setting Jumlah Biaya Administrasi Per Hari
    Set fRS = Nothing
    fQuery = "insert into BiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & msubKonversiKomaTitik(CStr(fTarif)) & "," & msubKonversiKomaTitik(CStr(fTarifCito)) & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "'," & fIdPegawai2 & ",'" & fIdUser & "',null)"
    Call msubRecFO(fRS, fQuery)
    'aktifkan kode berikut jika Paket Pelayanan TM sudah dioperasikan
    'select fKdPaketTM=KdPaket from PasienDaftar where NoPendaftaran=fNoPendaftaran
    'if(fKdPaketTM is not null) and (fKdPaketTM<>'')
    '    insert into BiayaPelayananPaketTM values(fNoPendaftaran,fKdRuangan,fKdPelayananRS,fTglPelayanan,fKdPaketTM,fTarif,fTarifCito,fJmlPelayanan)
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelas,KdKelompokPasien from V_KelasTanggunganPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value) Else fIdPenjamin = "2222222222"
    If fRS.EOF = False Then fKdKelasPenjamin = IIf(IsNull(fRS("KdKelasPenjamin").value), fKdKelas, fRS("KdKelasPenjamin").value) Else fKdKelasPenjamin = fKdKelas
    If fRS.EOF = False Then fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value) Else fKdKelompokPasien = "01"
    Set fRS = Nothing
    fQuery = "select KdPaket from V_PaketNonPaketTanggunganPenjamin where KdPelayananRS='" & fKdPelayananRS & "' and IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdKelas='" & fKdKelasPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdPaket = IIf(IsNull(fRS("KdPaket").value), "030", fRS("KdPaket").value) Else fKdPaket = "030"
    Set fRS = Nothing
    fQuery = "select KdPaket from V_PaketPenjamin where KdPelayananRS='" & fKdPelayananRS & "' and IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdKelas='" & fKdKelasPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdPaketPaket = fRS("KdPaket").value Else fKdPaketPaket = ""
    fTotalTarif = fTarif + fTarifCito
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTM(fNoPendaftaran,fKdPelayananRS,fKdKelasPenjamin,fKdJenisTarif,fStatusCITO,fIdPegawai,fIdPegawai2,null,'C') as TarifCitoPenjamin"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTarifCitoPenjamin = IIf(IsNull(fRS("TarifCitoPenjamin").value), 0, fRS("TarifCitoPenjamin").value) Else fTarifCitoPenjamin = 0
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTM(fNoPendaftaran,fKdPelayananRS,fKdKelasPenjamin,fKdJenisTarif,fStatusCITO,fIdPegawai,fIdPegawai2,null,'T') as TarifKelasPenjamin"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTarifKelasPenjamin = IIf(IsNull(fRS("TarifKelasPenjamin").value), 0, fRS("TarifKelasPenjamin").value) Else fTarifKelasPenjamin = 0
    If fTarifKelasPenjamin = 0 Then fTarifKelasPenjamin = fTarif
    fTotalTarifPenjamin = fTarifCitoPenjamin + fTarifKelasPenjamin
    Set fRS = Nothing
    fQuery = "select PersenTanggunganTM,PersenTanggunganRSTM from PersentaseTPBPTM where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fPersenTanggungan = IIf(IsNull(fRS("PersenTanggunganTM").value), 0, fRS("PersenTanggunganTM").value) Else fPersenTanggungan = 0
    If fRS.EOF = False Then fPersenTanggunganRS = IIf(IsNull(fRS("PersenTanggunganRSTM").value), 0, fRS("PersenTanggunganRSTM").value) Else fPersenTanggunganRS = 0
    'Cek Apakah Ada Penjamin di Paket & Non Paket Asuransi
    Set fRS = Nothing
    fQuery = "select distinct IdPenjamin from V_PaketNonPaketTanggunganPenjamin where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        'Tidak Ada di Paket & Non Paket Asuransi
        Set fRS2 = Nothing
        fQuery2 = "select KdPelayananRS  from DaftarTMNonTanggungan where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS.EOF = True Then
            fJmlHutangPenjamin = (fTotalTarifPenjamin * fPersenTanggungan) / 100
            fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
            If fSisaTagihanPasien > 0 Then
                fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
            Else
                fJmlTanggunganRS = 0
            End If
        Else
            fJmlHutangPenjamin = 0
            fJmlTanggunganRS = 0
        End If
        Set fRS3 = Nothing
        fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & fTarifKelasPenjamin & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "',null)"
        Call msubRecFO(fRS3, fQuery3)
    Else
        'Ada Penjamin di Paket & Non Paket Asuransi
        'Cek Apakah Ada Layanan di Paket & Non Paket Asuransi
        Set fRS2 = Nothing
        fQuery2 = "select IdPenjamin from V_PaketNonPaketTanggunganPenjamin where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelasPenjamin & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then
            'Tidak Ada Layanan di Paket & Non Paket Asuransi
            Set fRS2 = Nothing
            fQuery2 = "select KdPelayananRS  from DaftarTMNonTanggungan where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "'  and KdPelayananRS='" & fKdPelayananRS & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                fJmlHutangPenjamin = (fTotalTarifPenjamin * fPersenTanggungan) / 100
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
            Else
                fJmlHutangPenjamin = (fTotalTarifPenjamin * fPersenTanggungan) / 100
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
            End If
            Set fRS3 = Nothing
            fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & fTarifKelasPenjamin & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & ",0," & fKdPaket & "," & fIdPegawai2 & ",''" & fIdUser & "'',null)"
            Call msubRecFO(fRS3, fQuery3)
        Else
            'Cek Apakah Ada di Paket Asuransi
            Set fRS2 = Nothing
            fQuery2 = "select IdPenjamin from V_PaketPenjamin where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelasPenjamin & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                'Ada di Non Paket Asuransi
                Set fRS3 = Nothing
                fQuery3 = "select JmlTanggungan from TanggunganAsuransiNonPaket where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelasPenjamin & "'"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fJmlHutangPenjamin = IIf(IsNull(fRS3("JmlTanggungan").value), 0, fRS3("JmlTanggungan").value) Else fJmlHutangPenjamin = 0
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
                Set fRS3 = Nothing
                fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & fTarifKelasPenjamin & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & ",0," & fKdPaket & "," & fIdPegawai2 & ",''" & fIdUser & "'',null)"
                Call msubRecFO(fRS3, fQuery3)
            Else
                'Ada di Paket Asuransi
                Set fRS3 = Nothing
                fQuery3 = "select sum(Tarif) as TarifSum from V_ListBiayaPelayananPasienStrukNullPaket where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketPaket & "' and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "')"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fTotalBiayaPaket = IIf(IsNull(fRS3("TarifSum").value), 0, fRS3("TarifSum").value) Else fTotalBiayaPaket = 0
                Set fRS3 = Nothing
                fQuery3 = "select JmlTanggungan from V_JmlTanggunganPaketPenjamin where KdPaket='" & fKdPaketPaket & "' and IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdKelas='" & fKdKelasPenjamin & "'"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fTotalTanggunganPaket = IIf(IsNull(fRS3("JmlTanggungan").value), 0, fRS3("JmlTanggungan").value) Else fTotalTanggunganPaket = 0
                If fTotalBiayaPaket = 0 Then
                    fJmlHutangPenjamin = 0
                Else
                    fJmlHutangPenjamin = (CDec(fTotalTarifPenjamin) / CDec(fTotalBiayaPaket)) * CDec(fTotalTanggunganPaket)
                End If
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
                Set fRS3 = Nothing
                fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & fTarifKelasPenjamin & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "',null)"
                Call msubRecFO(fRS3, fQuery3)
                'begin of update Tanggungan yg termasuk Paket
                Set fRS = Nothing
                fQuery = "select KdPaket,TglPelayanan from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketPaket & "' and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "')"
                Call msubRecFO(fRS, fQuery)
                While fRS.EOF = False
                    fKdPaketL = fRS("KdPaket").value
                    fTglPelayananL = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
                    Set fRS2 = Nothing
                    fQuery2 = "select KdPelayananRS,TarifKelasPenjamin from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketPaket & "' and day(TglPelayanan)=day('" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "')"
                    Call msubRecFO(fRS2, fQuery2)
                    While fRS2.EOF = False
                        fKdPelayananRSL = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
                        fTarifKelasPenjaminL = IIf(IsNull(fRS("TarifKelasPenjamin").value), 0, fRS("TarifKelasPenjamin").value)
                        fJmlHutangPenjaminL = (CDec(fTarifKelasPenjaminL) / CDec(fTotalBiayaPaket)) * CDec(fTotalTanggunganPaket)
                        Set fRS3 = Nothing
                        fQuery3 = "SELECT  JmlTanggungan FROM TanggunganPaketAsuransiPerTindakan WHERE KdPaket = '" & fKdPaketPaket & "' AND IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' AND KdPelayananRS = '" & fKdPelayananRS & "' AND KdKelas = '" & fKdKelasPenjamin & "'"
                        Call msubRecFO(fRS3, fQuery3)
                        If fRS3.EOF = False Then fJmlHutangPenjaminPPT = IIf(IsNull(fRS3("JmlTanggungan").value), 0, fRS("JmlTanggungan").value) Else fJmlHutangPenjaminPPT = 0
                        If fRS3.EOF = False Then fJmlHutangPenjaminL = fJmlHutangPenjaminPPT
                        fSisaTagihanPasienL = fTotalTarifPenjamin - fJmlHutangPenjaminL
                        If fSisaTagihanPasienL > 0 Then
                            fJmlTanggunganRSL = (fSisaTagihanPasienL * fPersenTanggunganRS) / 100
                        Else
                            fJmlTanggunganRSL = 0
                        End If
                        Set fRS3 = Nothing
                        fQuery3 = "update DetailBiayaPelayanan set JmlHutangPenjamin=" & fJmlHutangPenjaminL & ",JmlTanggunganRS=" & fJmlTanggunganRSL & " where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketL & "' and TglPelayanan='" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRSL & "'"
                        Call msubRecFO(fRS3, fQuery3)
                        fRS2.MoveNext
                    Wend
                    fRS.MoveNext
                Wend
                'end of update Tanggungan yg termasuk Paket
            End If
        End If
    End If
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "'," & fNoLab_Rad & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
    Call f_AddTempHargaKomponen(fNoPendaftaran, fKdRuangan, fTglPelayanan, fKdPelayananRS, fKdKelas, fKdJenisTarif, CDec(fTarifCito), fJmlPelayanan, fStatusCito, CStr(fIdPegawai))
    Call f_AMDataKunjunganPelayananTMPasienPH(fNoPendaftaran, fKdRuangan, fKdRuanganAsal, fTglPelayanan, fKdPelayananRS, fIdPenjamin, fKdKelompokPasien, fJmlPelayanan, fNoLab_Rad, "A")
End Function

'Konversi dari SP: Add_BiayaPelayananAdmin
Public Function f_AddBiayaPelayananAdmin(fNoPendaftaran As String, fKdSubInstalasi As String, fKdRuangan As String, fKdPelayananRS As String, fKdKelas As String, fStatusCito As String, fTarif As Double, fJmlPelayanan As Integer, fTglPelayanan As Date, fNoLab_Rad As Variant, fIdPegawai As Variant, fStatusAPBD As String, fKdJenisTarif As String, fTarifCito As Integer, fIdUser As String, fIdPegawai2 As Variant)
    Dim fIdPenjamin As String
    Dim fKdKelasPenjamin As String
    Dim fKdKelompokPasien As String
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fJmlTanggunganRSL As Currency
    Dim fPersenTanggungan As Double
    Dim fPersenTanggunganRS As Double
    Dim fTotalTarif As Currency
    Dim fTarifKelasPenjamin As Currency
    Dim fTarifCitoKelasPenjamin As Currency
    Dim fPersenTarifCito As Double
    Dim fTarifCitoPenjamin As Currency
    Dim fTotalTarifPenjamin As Currency
    Dim fKdPaket As Variant
    Dim fTotalBiayaPaket As Currency
    Dim fTotalTanggunganPaket As Currency
    Dim fKdPaketL As String
    Dim fTarifKelasPenjaminL As Currency
    Dim fJmlHutangPenjaminL As Currency
    Dim fKdPelayananRSL As String
    Dim fTglPelayananL As Date
    Dim fKdInstalasi As String
    Dim fJmlHutangPenjaminPPT As Currency
    Dim fKdPaketTM As String
    Dim fKdPaketPaket As String
    Dim fSisaTagihanPasien As Currency
    Dim fSisaTagihanPasienL As Currency
    Dim fKdRuanganAsal As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "insert into BiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "'," & fIdPegawai2 & ",'" & fIdUser & "',null)"
    Call msubRecFO(fRS, fQuery)
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelas,KdKelompokPasien from V_KelasTanggunganPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value) Else fIdPenjamin = "2222222222"
    If fRS.EOF = False Then fKdKelasPenjamin = IIf(IsNull(fRS("KdKelasPenjamin").value), fKdKelas, fRS("KdKelasPenjamin").value) Else fKdKelasPenjamin = fKdKelas
    If fRS.EOF = False Then fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value) Else fKdKelompokPasien = "01"
    Set fRS = Nothing
    fQuery = "select KdPaket from V_PaketNonPaketTanggunganPenjamin where KdPelayananRS='" & fKdPelayananRS & "' and IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdKelas='" & fKdKelasPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdPaket = IIf(IsNull(fRS("KdPaket").value), "030", fRS("KdPaket").value) Else fKdPaket = "030"
    Set fRS = Nothing
    fQuery = "select KdPaket from V_PaketPenjamin where KdPelayananRS='" & fKdPelayananRS & "' and IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdKelas='" & fKdKelasPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdPaketPaket = fRS("KdPaket").value Else fKdPaketPaket = ""
    fTotalTarif = fTarif + fTarifCito
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTM(fNoPendaftaran,fKdPelayananRS,fKdKelasPenjamin,fKdJenisTarif,fStatusCITO,fIdPegawai,fIdPegawai2,null,'C') as TarifCitoPenjamin"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTarifCitoPenjamin = IIf(IsNull(fRS("TarifCitoPenjamin").value), 0, fRS("TarifCitoPenjamin").value) Else fTarifCitoPenjamin = 0
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTM(fNoPendaftaran,fKdPelayananRS,fKdKelasPenjamin,fKdJenisTarif,fStatusCITO,fIdPegawai,fIdPegawai2,null,'T') as TarifKelasPenjamin"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTarifKelasPenjamin = IIf(IsNull(fRS("TarifKelasPenjamin").value), 0, fRS("TarifKelasPenjamin").value) Else fTarifKelasPenjamin = 0
    If fTarifKelasPenjamin = 0 Then fTarifKelasPenjamin = fTarif
    fTotalTarifPenjamin = fTarifCitoPenjamin + fTarifKelasPenjamin
    Set fRS = Nothing
    fQuery = "select PersenTanggunganTM,PersenTanggunganRSTM from PersentaseTPBPTM where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fPersenTanggungan = IIf(IsNull(fRS("PersenTanggunganTM").value), 0, fRS("PersenTanggunganTM").value) Else fPersenTanggungan = 0
    If fRS.EOF = False Then fPersenTanggunganRS = IIf(IsNull(fRS("PersenTanggunganRSTM").value), 0, fRS("PersenTanggunganRSTM").value) Else fPersenTanggunganRS = 0
    'Cek Apakah Ada Penjamin di Paket & Non Paket Asuransi
    Set fRS = Nothing
    fQuery = "select distinct IdPenjamin from V_PaketNonPaketTanggunganPenjamin where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        'Tidak Ada di Paket & Non Paket Asuransi
        Set fRS2 = Nothing
        fQuery2 = "select KdPelayananRS  from DaftarTMNonTanggungan where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS.EOF = True Then
            fJmlHutangPenjamin = (fTotalTarifPenjamin * fPersenTanggungan) / 100
            fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
            If fSisaTagihanPasien > 0 Then
                fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
            Else
                fJmlTanggunganRS = 0
            End If
        Else
            fJmlHutangPenjamin = 0
            fJmlTanggunganRS = 0
        End If
        Set fRS3 = Nothing
        fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & fTarifKelasPenjamin & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "',null)"
        Call msubRecFO(fRS3, fQuery3)
    Else
        'Ada Penjamin di Paket & Non Paket Asuransi
        'Cek Apakah Ada Layanan di Paket & Non Paket Asuransi
        Set fRS2 = Nothing
        fQuery2 = "select IdPenjamin from V_PaketNonPaketTanggunganPenjamin where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelasPenjamin & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then
            'Tidak Ada Layanan di Paket & Non Paket Asuransi
            Set fRS2 = Nothing
            fQuery2 = "select KdPelayananRS  from DaftarTMNonTanggungan where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "'  and KdPelayananRS='" & fKdPelayananRS & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                fJmlHutangPenjamin = (fTotalTarifPenjamin * fPersenTanggungan) / 100
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
            Else
                fJmlHutangPenjamin = (fTotalTarifPenjamin * fPersenTanggungan) / 100
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
            End If
            Set fRS3 = Nothing
            fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & fTarifKelasPenjamin & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "',null)"
            Call msubRecFO(fRS3, fQuery3)
        Else
            'Cek Apakah Ada di Paket Asuransi
            Set fRS2 = Nothing
            fQuery2 = "select IdPenjamin from V_PaketPenjamin where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelasPenjamin & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                'Ada di Non Paket Asuransi
                Set fRS3 = Nothing
                fQuery3 = "select JmlTanggungan from TanggunganAsuransiNonPaket where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelasPenjamin & "'"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fJmlHutangPenjamin = IIf(IsNull(fRS3("JmlTanggungan").value), 0, fRS3("JmlTanggungan").value) Else fJmlHutangPenjamin = 0
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
                Set fRS3 = Nothing
                fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & fTarifKelasPenjamin & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "',null)"
                Call msubRecFO(fRS3, fQuery3)
            Else
                'Ada di Paket Asuransi
                Set fRS3 = Nothing
                fQuery3 = "select sum(Tarif) as TarifSum from V_ListBiayaPelayananPasienStrukNullPaket where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketPaket & "' and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "')"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fTotalBiayaPaket = IIf(IsNull(fRS3("TarifSum").value), 0, fRS3("TarifSum").value) Else fTotalBiayaPaket = 0
                Set fRS3 = Nothing
                fQuery3 = "select JmlTanggungan from V_JmlTanggunganPaketPenjamin where KdPaket='" & fKdPaketPaket & "' and IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdKelas='" & fKdKelasPenjamin & "'"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fTotalTanggunganPaket = IIf(IsNull(fRS3("JmlTanggungan").value), 0, fRS3("JmlTanggungan").value) Else fTotalTanggunganPaket = 0
                If fTotalBiayaPaket = 0 Then
                    fJmlHutangPenjamin = 0
                Else
                    fJmlHutangPenjamin = (CDec(fTotalTarifPenjamin) / CDec(fTotalBiayaPaket)) * CDec(fTotalTanggunganPaket)
                End If
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
                Set fRS3 = Nothing
                fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & fTarifKelasPenjamin & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "',null)"
                Call msubRecFO(fRS3, fQuery3)
                'begin of update Tanggungan yg termasuk Paket
                Set fRS = Nothing
                fQuery = "select KdPaket,TglPelayanan from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketPaket & "' and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "')"
                Call msubRecFO(fRS, fQuery)
                While fRS.EOF = False
                    fKdPaketL = fRS("KdPaket").value
                    fTglPelayananL = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
                    Set fRS2 = Nothing
                    fQuery2 = "select KdPelayananRS,TarifKelasPenjamin from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketPaket & "' and day(TglPelayanan)=day('" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "')"
                    Call msubRecFO(fRS2, fQuery2)
                    While fRS2.EOF = False
                        fKdPelayananRSL = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
                        fTarifKelasPenjaminL = IIf(IsNull(fRS("TarifKelasPenjamin").value), 0, fRS("TarifKelasPenjamin").value)
                        fJmlHutangPenjaminL = (CDec(fTarifKelasPenjaminL) / CDec(fTotalBiayaPaket)) * CDec(fTotalTanggunganPaket)
                        Set fRS3 = Nothing
                        fQuery3 = "SELECT  JmlTanggungan FROM TanggunganPaketAsuransiPerTindakan WHERE KdPaket = '" & fKdPaketPaket & "' AND IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' AND KdPelayananRS = '" & fKdPelayananRS & "' AND KdKelas = '" & fKdKelasPenjamin & "'"
                        Call msubRecFO(fRS3, fQuery3)
                        If fRS3.EOF = False Then fJmlHutangPenjaminPPT = IIf(IsNull(fRS3("JmlTanggungan").value), 0, fRS("JmlTanggungan").value) Else fJmlHutangPenjaminPPT = 0
                        If fRS3.EOF = False Then fJmlHutangPenjaminL = fJmlHutangPenjaminPPT
                        fSisaTagihanPasienL = fTotalTarifPenjamin - fJmlHutangPenjaminL
                        If fSisaTagihanPasienL > 0 Then
                            fJmlTanggunganRSL = (fSisaTagihanPasienL * fPersenTanggunganRS) / 100
                        Else
                            fJmlTanggunganRSL = 0
                        End If
                        Set fRS3 = Nothing
                        fQuery3 = "update DetailBiayaPelayanan set JmlHutangPenjamin=" & fJmlHutangPenjaminL & ",JmlTanggunganRS=" & fJmlTanggunganRSL & " where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketL & "' and TglPelayanan='" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRSL & "'"
                        Call msubRecFO(fRS3, fQuery3)
                        fRS2.MoveNext
                    Wend
                    fRS.MoveNext
                Wend
                'end of update Tanggungan yg termasuk Paket
            End If
        End If
    End If
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "'," & fNoLab_Rad & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
    Call f_AddTempHargaKomponen(fNoPendaftaran, fKdRuangan, fTglPelayanan, fKdPelayananRS, fKdKelas, fKdJenisTarif, CDec(fTarifCito), fJmlPelayanan, fStatusCito, CStr(fIdPegawai))
    Call f_AMDataKunjunganPelayananTMPasienPH(fNoPendaftaran, fKdRuangan, fKdRuanganAsal, fTglPelayanan, fKdPelayananRS, fIdPenjamin, fKdKelompokPasien, fJmlPelayanan, fNoLab_Rad, "A")
End Function

''Konversi dari SP: Add_BiayaPelayananIBS
Public Function f_AddBiayaPelayananIBS(fNoPendaftaran As String, fKdSubInstalasi As String, fKdRuangan As String, fKdPelayananRS As String, fKdKelas As String, fStatusCito As String, fTarif As Currency, fJmlPelayanan As Integer, fTglPelayanan As Date, fNoLab_Rad As Variant, fIdPegawai As Variant, fStatusAPBD As String, fKdJenisTarif As String, fTarifCito As Currency, fIdUser As String, fIdPegawai2 As Variant, fIdPegawai3 As Variant)
    Dim fIdPenjamin As String
    Dim fKdKelasPenjamin As String
    Dim fKdKelompokPasien As String
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fPersenTanggungan As Double
    Dim fPersenTanggunganRS As Double
    Dim fTotalTarif As Currency
    Dim fTarifKelasPenjamin As Currency
    Dim fTarifCitoKelasPenjamin As Currency
    Dim fPersenTarifCito As Double
    Dim fTarifCitoPenjamin As Currency
    Dim fTotalTarifPenjamin As Currency
    Dim fKdPaket As Variant
    Dim fTotalBiayaPaket As Currency
    Dim fTotalTanggunganPaket As Currency
    Dim fKdPaketL As String
    Dim fTarifKelasPenjaminL As Currency
    Dim fJmlHutangPenjaminL As Currency
    Dim fKdPelayananRSL As String
    Dim fTglPelayananL As Date
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fKdRuanganAsal As String
    Dim fTglPelayananAdm As Date
    Dim fKdPelayananRSAdmin As String
    Dim fJmlHutangPenjaminPPT As Currency
    Dim fJmlPelayananTemp As Integer
    Dim fKdPaketTM As String
    Dim fKdPaketPaket As String
    Dim fSisaTagihanPasien As Currency
    Dim fTarifAdmin As Currency
    Dim fJmlTanggunganRSL As Currency
    Dim fSisaTagihanPasienL As Currency

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdPelayananRSAdmin from MasterDataPendukung"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdPelayananRSAdmin = IIf(IsNull(fRS("KdPelayananRSAdmin").value), "001001", fRS("KdPelayananRSAdmin").value) Else fKdPelayananRSAdmin = "001001"
    'Begin Setting Jumlah Biaya Administrasi Per Hari
    Set fRS = Nothing
    fQuery = "select sum(JmlPelayanan) as JmlPelayananSum from BiayaPelayanan where KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null and (day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') ) and KdPelayananRS<>'" & fKdPelayananRSAdmin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fJmlPelayananTemp = IIf(IsNull(fRS("JmlPelayananSum").value), 0, fRS("JmlPelayananSum").value) Else fJmlPelayananTemp = 0
    If fJmlPelayananTemp <= 5 Or fJmlPelayananTemp = 0 Then
        Set fRS = Nothing
        fQuery = "select min(TglPelayanan) as TglPelayananMin from BiayaPelayanan where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fTglPelayananAdm = IIf(IsNull(fRS("TglPelayananMin").value), "", fRS("TglPelayananMin").value) Else fTglPelayananAdm = ""
        If fTglPelayananAdm <> "" Then
            Set fRS2 = Nothing
            fQuery2 = "update BiayaPelayanan set JmlPelayanan=JmlPelayanan+" & fJmlPelayanan & " where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null and TglPelayanan='" & Format(fTglPelayananAdm, "yyyy/MM/dd HH:mm:ss") & "'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update DetailBiayaPelayanan set JmlPelayanan=JmlPelayanan+" & fJmlPelayanan & " where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null and TglPelayanan='" & Format(fTglPelayananAdm, "yyyy/MM/dd HH:mm:ss") & "'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponen set JmlPelayanan=JmlPelayanan+" & fJmlPelayanan & " where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null and TglPelayanan='" & Format(fTglPelayananAdm, "yyyy/MM/dd HH:mm:ss") & "'"
            Call msubRecFO(fRS2, fQuery2)
            'Call f_AddTempHargaKomponen(fNoPendaftaran, fKdRuangan, fTglPelayananAdm, fKdPelayananRSAdmin, fKdKelas, fKdJenisTarif, fTarifCito, fJmlPelayanan, fStatusCito, fIdPegawai)
            Call f_AddTempHargaKomponen(fNoPendaftaran, fKdRuangan, fTglPelayananAdm, fKdPelayananRSAdmin, fKdKelas, fKdJenisTarif, CDbl(fTarifCito), fJmlPelayanan, fStatusCito, CStr(fIdPegawai))
            'tambah onede
        End If
        '
    Else
        Set fRS2 = Nothing
        fQuery2 = "select dbo.FB_NewTakeTarifBPTM ('" & fNoPendaftaran & "', '" & fKdPelayananRSAdmin & "', '" & fKdKelas & "', '" & fKdJenisTarif & "', '0', " & fIdPegawai & ", null, null, 'T') as TarifAdmin"
        Call msubRecFO(fRS2, fQuery2)
        If fRS.EOF = False Then fTarifAdmin = IIf(IsNull(fRS("TarifAdmin").value), 0, fRS("TarifAdmin").value) Else fTarifAdmin = 0
        Call f_AddBiayaPelayananAdmin(fNoPendaftaran, fKdSubInstalasi, fKdRuangan, fKdPelayananRSAdmin, fKdKelas, "0", CDbl(fTarifAdmin), 1, fTglPelayanan, fNoLab_Rad, fIdPegawai, "01", fKdJenisTarif, 0, CStr(fIdPegawai), Null)
    End If
    'End Setting Jumlah Biaya Administrasi Per Hari
    Set fRS = Nothing
    fQuery = "insert into BiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "'," & fIdPegawai2 & ",'" & fIdUser & "'," & fIdPegawai3 & ")"
    Call msubRecFO(fRS, fQuery)
    'aktifkan kode berikut jika Paket Pelayanan TM sudah dioperasikan
    'select fKdPaketTM=KdPaket from PasienDaftar where NoPendaftaran=fNoPendaftaran
    'if(fKdPaketTM is not null) and (fKdPaketTM<>'')
    '    insert into BiayaPelayananPaketTM values(fNoPendaftaran,fKdRuangan,fKdPelayananRS,fTglPelayanan,fKdPaketTM,fTarif,fTarifCito,fJmlPelayanan)
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelas,KdKelompokPasien from V_KelasTanggunganPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value) Else fIdPenjamin = "2222222222"
    If fRS.EOF = False Then fKdKelasPenjamin = IIf(IsNull(fRS("KdKelasPenjamin").value), fKdKelas, fRS("KdKelasPenjamin").value) Else fKdKelasPenjamin = fKdKelas
    If fRS.EOF = False Then fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value) Else fKdKelompokPasien = "01"
    Set fRS = Nothing
    fQuery = "select KdPaket from V_PaketNonPaketTanggunganPenjamin where KdPelayananRS='" & fKdPelayananRS & "' and IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdKelas='" & fKdKelasPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdPaket = IIf(IsNull(fRS("KdPaket").value), "030", fRS("KdPaket").value) Else fKdPaket = "030"
    Set fRS = Nothing
    fQuery = "select KdPaket from V_PaketPenjamin where KdPelayananRS='" & fKdPelayananRS & "' and IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdKelas='" & fKdKelasPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdPaketPaket = fRS("KdPaket").value Else fKdPaketPaket = ""
    fTotalTarif = fTarif + fTarifCito
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTM(fNoPendaftaran,fKdPelayananRS,fKdKelasPenjamin,fKdJenisTarif,fStatusCITO,fIdPegawai,fIdPegawai2,null,'C') as TarifCitoPenjamin"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTarifCitoPenjamin = IIf(IsNull(fRS("TarifCitoPenjamin").value), 0, fRS("TarifCitoPenjamin").value) Else fTarifCitoPenjamin = 0
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTM(fNoPendaftaran,fKdPelayananRS,fKdKelasPenjamin,fKdJenisTarif,fStatusCITO,fIdPegawai,fIdPegawai2,null,'T') as TarifKelasPenjamin"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTarifKelasPenjamin = IIf(IsNull(fRS("TarifKelasPenjamin").value), 0, fRS("TarifKelasPenjamin").value) Else fTarifKelasPenjamin = 0
    If fTarifKelasPenjamin = 0 Then fTarifKelasPenjamin = fTarif
    fTotalTarifPenjamin = fTarifCitoPenjamin + fTarifKelasPenjamin
    Set fRS = Nothing
    fQuery = "select PersenTanggunganTM,PersenTanggunganRSTM from PersentaseTPBPTM where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fPersenTanggungan = IIf(IsNull(fRS("PersenTanggunganTM").value), 0, fRS("PersenTanggunganTM").value) Else fPersenTanggungan = 0
    If fRS.EOF = False Then fPersenTanggunganRS = IIf(IsNull(fRS("PersenTanggunganRSTM").value), 0, fRS("PersenTanggunganRSTM").value) Else fPersenTanggunganRS = 0
    'Cek Apakah Ada Penjamin di Paket & Non Paket Asuransi
    Set fRS = Nothing
    fQuery = "select distinct IdPenjamin from V_PaketNonPaketTanggunganPenjamin where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        'Tidak Ada di Paket & Non Paket Asuransi
        Set fRS2 = Nothing
        fQuery2 = "select KdPelayananRS  from DaftarTMNonTanggungan where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS.EOF = True Then
            fJmlHutangPenjamin = (fTotalTarifPenjamin * fPersenTanggungan) / 100
            fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
            If fSisaTagihanPasien > 0 Then
                fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
            Else
                fJmlTanggunganRS = 0
            End If
        Else
            fJmlHutangPenjamin = 0
            fJmlTanggunganRS = 0
        End If
        Set fRS3 = Nothing
        fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & fTarifKelasPenjamin & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "'," & fIdPegawai3 & ")"
        Call msubRecFO(fRS3, fQuery3)
    Else
        'Ada Penjamin di Paket & Non Paket Asuransi
        'Cek Apakah Ada Layanan di Paket & Non Paket Asuransi
        Set fRS2 = Nothing
        fQuery2 = "select IdPenjamin from V_PaketNonPaketTanggunganPenjamin where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelasPenjamin & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then
            'Tidak Ada Layanan di Paket & Non Paket Asuransi
            Set fRS2 = Nothing
            fQuery2 = "select KdPelayananRS  from DaftarTMNonTanggungan where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "'  and KdPelayananRS='" & fKdPelayananRS & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                fJmlHutangPenjamin = (fTotalTarifPenjamin * fPersenTanggungan) / 100
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
            Else
                fJmlHutangPenjamin = (fTotalTarifPenjamin * fPersenTanggungan) / 100
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
            End If
            Set fRS3 = Nothing
            fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & fTarifKelasPenjamin & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "'," & fIdPegawai3 & ")"
            Call msubRecFO(fRS3, fQuery3)
        Else
            'Cek Apakah Ada di Paket Asuransi
            Set fRS2 = Nothing
            fQuery2 = "select IdPenjamin from V_PaketPenjamin where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelasPenjamin & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                'Ada di Non Paket Asuransi
                Set fRS3 = Nothing
                fQuery3 = "select JmlTanggungan from TanggunganAsuransiNonPaket where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelasPenjamin & "'"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fJmlHutangPenjamin = IIf(IsNull(fRS3("JmlTanggungan").value), 0, fRS3("JmlTanggungan").value) Else fJmlHutangPenjamin = 0
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
                Set fRS3 = Nothing
                fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & fTarifKelasPenjamin & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "'," & fIdPegawai3 & ")"
                Call msubRecFO(fRS3, fQuery3)
            Else
                'Ada di Paket Asuransi
                Set fRS3 = Nothing
                fQuery3 = "select sum(Tarif) as TarifSum from V_ListBiayaPelayananPasienStrukNullPaket where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketPaket & "' and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "')"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fTotalBiayaPaket = IIf(IsNull(fRS3("TarifSum").value), 0, fRS3("TarifSum").value) Else fTotalBiayaPaket = 0
                Set fRS3 = Nothing
                fQuery3 = "select JmlTanggungan from V_JmlTanggunganPaketPenjamin where KdPaket='" & fKdPaketPaket & "' and IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdKelas='" & fKdKelasPenjamin & "'"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fTotalTanggunganPaket = IIf(IsNull(fRS3("JmlTanggungan").value), 0, fRS3("JmlTanggungan").value) Else fTotalTanggunganPaket = 0
                If fTotalBiayaPaket = 0 Then
                    fJmlHutangPenjamin = 0
                Else
                    fJmlHutangPenjamin = (CDec(fTotalTarifPenjamin) / CDec(fTotalBiayaPaket)) * CDec(fTotalTanggunganPaket)
                End If
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
                Set fRS3 = Nothing
                fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & fTarifKelasPenjamin & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "'," & fIdPegawai3 & ")"
                Call msubRecFO(fRS3, fQuery3)
                'begin of update Tanggungan yg termasuk Paket
                Set fRS = Nothing
                fQuery = "select KdPaket,TglPelayanan from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketPaket & "' and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "')"
                Call msubRecFO(fRS, fQuery)
                While fRS.EOF = False
                    fKdPaketL = fRS("KdPaket").value
                    fTglPelayananL = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
                    Set fRS2 = Nothing
                    fQuery2 = "select KdPelayananRS,TarifKelasPenjamin from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketPaket & "' and day(TglPelayanan)=day('" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "')"
                    Call msubRecFO(fRS2, fQuery2)
                    While fRS2.EOF = False
                        fKdPelayananRSL = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
                        fTarifKelasPenjaminL = IIf(IsNull(fRS("TarifKelasPenjamin").value), 0, fRS("TarifKelasPenjamin").value)
                        fJmlHutangPenjaminL = (CDec(fTarifKelasPenjaminL) / CDec(fTotalBiayaPaket)) * CDec(fTotalTanggunganPaket)
                        Set fRS3 = Nothing
                        fQuery3 = "SELECT  JmlTanggungan FROM TanggunganPaketAsuransiPerTindakan WHERE KdPaket = '" & fKdPaketPaket & "' AND IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' AND KdPelayananRS = '" & fKdPelayananRS & "' AND KdKelas = '" & fKdKelasPenjamin & "'"
                        Call msubRecFO(fRS3, fQuery3)
                        If fRS3.EOF = False Then fJmlHutangPenjaminPPT = IIf(IsNull(fRS3("JmlTanggungan").value), 0, fRS("JmlTanggungan").value) Else fJmlHutangPenjaminPPT = 0
                        If fRS3.EOF = False Then fJmlHutangPenjaminL = fJmlHutangPenjaminPPT
                        fSisaTagihanPasienL = fTotalTarifPenjamin - fJmlHutangPenjaminL
                        If fSisaTagihanPasienL > 0 Then
                            fJmlTanggunganRSL = (fSisaTagihanPasienL * fPersenTanggunganRS) / 100
                        Else
                            fJmlTanggunganRSL = 0
                        End If
                        Set fRS3 = Nothing
                        fQuery3 = "update DetailBiayaPelayanan set JmlHutangPenjamin=" & fJmlHutangPenjaminL & ",JmlTanggunganRS=" & fJmlTanggunganRSL & " where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketL & "' and TglPelayanan='" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRSL & "'"
                        Call msubRecFO(fRS3, fQuery3)
                        fRS2.MoveNext
                    Wend
                    fRS.MoveNext
                Wend
                'end of update Tanggungan yg termasuk Paket
            End If
        End If
    End If
    Call f_AddTempHargaKomponenForIBS(fNoPendaftaran, fKdRuangan, fTglPelayanan, fKdPelayananRS, fKdKelas, fKdJenisTarif, fJmlPelayanan)
    Call f_AMDataKunjunganPelayananTMPasienPH(fNoPendaftaran, fKdRuangan, fKdRuanganAsal, fTglPelayanan, fKdPelayananRS, fIdPenjamin, fKdKelompokPasien, fJmlPelayanan, fNoLab_Rad, "A")
    Set fRS = Nothing
    fQuery = "delete from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & KdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "'"
    Call msubRecFO(fRS, fQuery)
End Function

'Konversi dari SP: Add_BiayaPelayananOtomatis
Public Function f_AddBiayaPelayananOtomatis(fNoPendaftaran As String, fKdSubInstalasi As String, fKdRuangan As String, fTglMasuk As Date, fKdKelas As String, fKdKelasPel As String, fNoLab_Rad As Variant, fIdPegawai As Variant, fStatus As String)
    'fKdKelas: Kelas Kamar & allow null; fNoLab_Rad: NoIBS,NoLaboratorium,NoRadiology, NoPakai -->Allow null
    'fStatus: AL=RJ,IGD; MA=Rawat Mandiri;RG=Rawat Gabung-->allow null
    Dim fKdPelayananRS As String
    Dim fKdJenisTarif As String
    Dim fTarif As Currency
    Dim fKdRuanganAsal As String
    Dim fKdPelayananRS2 As String
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    If fKdKelas = "04" Then
        Set fRS = Nothing
        fQuery = "select KdPelayananRS from PelayananRuangan where KdRuangan='" & fKdRuangan & "' and Status='CU'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            Set fRS2 = Nothing
            fQuery2 = "select KdPelayananRS from PelayananRuangan where KdRuangan='" & fKdRuangan & "' and Status='" & fStatus & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = False Then
                fKdPelayananRS = IIf(IsNull(fRS2("KdPelayananRS").value), "", fRS2("KdPelayananRS").value)
                Set fRS3 = Nothing
                fQuery3 = "select KdJenisTarif from v_JenisTarifPasien where NoPendaftaran='" & fNoPendaftaran & "'"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fKdJenisTarif = IIf(IsNull(fRS3("KdJenisTarif").value), "01", fRS3("KdJenisTarif").value) Else fKdJenisTarif = "01"
                Set fRS3 = Nothing
                fQuery3 = "select dbo.FB_NewTakeTarifBPTM ('" & fNoPendaftaran & "', '" & fKdPelayananRS & "', '" & fKdKelasPel & "', '" & fKdJenisTarif & "', '0', " & fIdPegawai & ", null, null, 'T') as Tarif"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fTarif = IIf(IsNull(fRS3("Tarif").value), 0, fRS3("Tarif").value) Else fTarif = 0
                If fKdPelayananRS <> "" Then
                    Call f_AddBiayaPelayanan(fNoPendaftaran, fKdSubInstalasi, fKdRuangan, fKdPelayananRS, fKdKelasPel, "0", CDbl(fTarif), 1, fTglMasuk, fNoLab_Rad, fIdPegawai, "01", fKdJenisTarif, 0, CStr(fIdPegawai), Null)
                End If
            End If
        Else
            Set fRS2 = Nothing
            fQuery2 = "select KdPelayananRS from PelayananRuangan where KdRuangan='" & fKdRuangan & "' and Status='CU'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = False Then
                fKdPelayananRS = IIf(IsNull(fRS2("KdPelayananRS").value), "", fRS2("KdPelayananRS").value)
                If fKdPelayananRS <> "" Then
                    Set fRS3 = Nothing
                    fQuery3 = "select KdJenisTarif from v_JenisTarifPasien where NoPendaftaran='" & fNoPendaftaran & "'"
                    Call msubRecFO(fRS3, fQuery3)
                    If fRS3.EOF = False Then fKdJenisTarif = IIf(IsNull(fRS3("KdJenisTarif").value), "01", fRS3("KdJenisTarif").value) Else fKdJenisTarif = "01"
                    Set fRS3 = Nothing
                    fQuery3 = "select dbo.FB_NewTakeTarifBPTM ('" & fNoPendaftaran & "', '" & fKdPelayananRS & "', '" & fKdKelasPel & "', '" & fKdJenisTarif & "', '0', " & fIdPegawai & ", null, null, 'T') as Tarif"
                    Call msubRecFO(fRS3, fQuery3)
                    If fRS3.EOF = False Then fTarif = IIf(IsNull(fRS3("Tarif").value), 0, fRS3("Tarif").value) Else fTarif = 0
                    Call f_AddBiayaPelayanan(fNoPendaftaran, fKdSubInstalasi, fKdRuangan, fKdPelayananRS, fKdKelasPel, "0", CDbl(fTarif), 1, fTglMasuk, fNoLab_Rad, fIdPegawai, "01", fKdJenisTarif, 0, CStr(fIdPegawai), Null)
                End If
            End If
        End If
    Else
        Set fRS2 = Nothing
        fQuery2 = "select KdPelayananRS from PelayananRuangan where KdRuangan='" & fKdRuangan & "' and Status='" & fStatus & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fKdPelayananRS = IIf(IsNull(fRS2("KdPelayananRS").value), "", fRS2("KdPelayananRS").value)
        If fKdPelayananRS <> "" Then
            Set fRS3 = Nothing
            fQuery3 = "select KdJenisTarif from v_JenisTarifPasien where NoPendaftaran='" & fNoPendaftaran & "'"
            Call msubRecFO(fRS3, fQuery3)
            If fRS3.EOF = False Then fKdJenisTarif = IIf(IsNull(fRS3("KdJenisTarif").value), "01", fRS3("KdJenisTarif").value) Else fKdJenisTarif = "01"
            Set fRS3 = Nothing
            fQuery3 = "select dbo.FB_NewTakeTarifBPTM ('" & fNoPendaftaran & "', '" & fKdPelayananRS & "', '" & fKdKelasPel & "', '" & fKdJenisTarif & "', '0', " & fIdPegawai & ", null, null, 'T') as Tarif"
            Call msubRecFO(fRS3, fQuery3)
            If fRS3.EOF = False Then fTarif = IIf(IsNull(fRS3("Tarif").value), 0, fRS3("Tarif").value) Else fTarif = 0
            Call f_AddBiayaPelayanan(fNoPendaftaran, fKdSubInstalasi, fKdRuangan, fKdPelayananRS, fKdKelasPel, "0", CDbl(fTarif), 1, fTglMasuk, fNoLab_Rad, fIdPegawai, "01", fKdJenisTarif, 0, CStr(fIdPegawai), Null)
        End If
    End If
    Set fRS2 = Nothing
    fQuery2 = "select KdPelayananRS from PelayananRuangan where KdRuangan='" & fKdRuangan & "' and Status='AD'"
    Call msubRecFO(fRS2, fQuery2)
    If fRS2.EOF = False Then
        fKdPelayananRS2 = IIf(IsNull(fRS2("KdPelayananRS").value), "", fRS2("KdPelayananRS").value)
        If fKdPelayananRS2 <> "" Then
            Set fRS3 = Nothing
            fQuery3 = "select KdJenisTarif from v_JenisTarifPasien where NoPendaftaran='" & fNoPendaftaran & "'"
            Call msubRecFO(fRS3, fQuery3)
            If fRS3.EOF = False Then fKdJenisTarif = IIf(IsNull(fRS3("KdJenisTarif").value), "01", fRS3("KdJenisTarif").value) Else fKdJenisTarif = "01"
            Set fRS3 = Nothing
            fQuery3 = "select dbo.FB_NewTakeTarifBPTM ('" & fNoPendaftaran & "', '" & fKdPelayananRS2 & "', '" & fKdKelasPel & "', '" & fKdJenisTarif & "', '0', " & fIdPegawai & ", null, null, 'T') as Tarif"
            Call msubRecFO(fRS3, fQuery3)
            If fRS3.EOF = False Then fTarif = IIf(IsNull(fRS3("Tarif").value), 0, fRS3("Tarif").value) Else fTarif = 0
            Call f_AddBiayaPelayanan(fNoPendaftaran, fKdSubInstalasi, fKdRuangan, fKdPelayananRS2, fKdKelasPel, "0", CDbl(fTarif), 1, fTglMasuk, fNoLab_Rad, fIdPegawai, "01", fKdJenisTarif, 0, CStr(fIdPegawai), Null)
        End If
    End If
End Function

'Konversi dari SP: Add_BiayaPelayananPenunjangM
Public Function f_AddBiayaPelayananPenunjangM(fNoPendaftaran As String, fKdSubInstalasi As String, fKdRuangan As String, fKdPelayananRS As String, fKdKelas As String, fStatusCito As String, fTarif As Double, fJmlPelayanan As Integer, fTglPelayanan As Date, fNoLab_Rad As Variant, fIdPegawai As Variant, fStatusAPBD As String, fKdJenisTarif As String, fTarifCito As Integer, fIdUser As String, fIdPegawai2 As Variant, fKdLaboratory As String)
    Dim fIdPenjamin As String
    Dim fKdKelasPenjamin As String
    Dim fKdKelompokPasien As String
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fJmlTanggunganRSL As Currency
    Dim fPersenTanggungan As Double
    Dim fPersenTanggunganRS As Double
    Dim fTotalTarif As Currency
    Dim fTarifKelasPenjamin As Currency
    Dim fTarifCitoKelasPenjamin As Currency
    Dim fPersenTarifCito As Double
    Dim fTarifCitoPenjamin As Currency
    Dim fTotalTarifPenjamin As Currency
    Dim fKdPaket As Variant
    Dim fTotalBiayaPaket As Currency
    Dim fTotalTanggunganPaket As Currency
    Dim fKdPaketL As String
    Dim fTarifKelasPenjaminL As Currency
    Dim fJmlHutangPenjaminL As Currency
    Dim fKdPelayananRSL As String
    Dim fTglPelayananL As Date
    Dim fKdInstalasi As String
    Dim fTglPelayananAdm As Date
    Dim fKdPelayananRSAdmin As String
    Dim fJmlHutangPenjaminPPT As Currency
    Dim fJmlPelayananTemp As Integer
    Dim fKdPaketTM As String
    Dim fKdPaketPaket As String
    Dim fSisaTagihanPasien As Currency
    Dim fSisaTagihanPasienL As Currency
    Dim fTarifAdmin As Currency
    Dim fKdRuanganAsal As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "KdInstalasi from Ruangan where KdRuangan='" & fKdRuangan & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdInstalasi = IIf(IsNull(fRS("KdInstalasi").value), "", fRS("KdInstalasi").value) Else fKdInstalasi = ""
    If fKdInstalasi <> "10" And fKdInstalasi <> "09" And fKdInstalasi <> "16" Then
        Set fRS = Nothing
        fQuery = "select KdPelayananRSAdmin from MasterDataPendukung"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fKdPelayananRSAdmin = IIf(IsNull(fRS("KdPelayananRSAdmin").value), "001001", fRS("KdPelayananRSAdmin").value) Else fKdPelayananRSAdmin = "001001"
        'Begin Setting Jumlah Biaya Administrasi Per Hari
        Set fRS = Nothing
        fQuery = "select sum(JmlPelayanan) as JmlPelayananSum from BiayaPelayanan where KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null and (day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') ) and KdPelayananRS<>'" & fKdPelayananRSAdmin & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fJmlPelayananTemp = IIf(IsNull(fRS("JmlPelayananSum").value), 0, fRS("JmlPelayananSum").value) Else fJmlPelayananTemp = 0
        If fJmlPelayananTemp <= 5 Or fJmlPelayananTemp = 0 Then
            Set fRS = Nothing
            fQuery = "select min(TglPelayanan) as TglPelayananMin from BiayaPelayanan where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
            Call msubRecFO(fRS, fQuery)
            If fRS.EOF = False Then fTglPelayananAdm = IIf(IsNull(fRS("TglPelayananMin").value), "", fRS("TglPelayananMin").value) Else fTglPelayananAdm = ""
            If fTglPelayananAdm <> "" Then
                Set fRS2 = Nothing
                fQuery2 = "update BiayaPelayanan set JmlPelayanan=JmlPelayanan+" & fJmlPelayanan & " where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null and TglPelayanan='" & Format(fTglPelayananAdm, "yyyy/MM/dd HH:mm:ss") & "'"
                Call msubRecFO(fRS2, fQuery2)
                Set fRS2 = Nothing
                fQuery2 = "update DetailBiayaPelayanan set JmlPelayanan=JmlPelayanan+" & fJmlPelayanan & " where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null and TglPelayanan='" & Format(fTglPelayananAdm, "yyyy/MM/dd HH:mm:ss") & "'"
                Call msubRecFO(fRS2, fQuery2)
                Set fRS2 = Nothing
                fQuery2 = "update TempHargaKomponen set JmlPelayanan=JmlPelayanan+" & fJmlPelayanan & " where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null and TglPelayanan='" & Format(fTglPelayananAdm, "yyyy/MM/dd HH:mm:ss") & "'"
                Call msubRecFO(fRS2, fQuery2)
                Call f_AddTempHargaKomponen(fNoPendaftaran, fKdRuangan, fTglPelayananAdm, fKdPelayananRSAdmin, fKdKelas, fKdJenisTarif, CDbl(fTarifCito), fJmlPelayanan, fStatusCito, CStr(fIdPegawai))
            End If
        Else
            Set fRS2 = Nothing
            fQuery2 = "select dbo.FB_NewTakeTarifBPTM ('" & fNoPendaftaran & "', '" & fKdPelayananRSAdmin & "', '" & fKdKelas & "', '" & fKdJenisTarif & "', '0', " & fIdPegawai & ", null, null, 'T') as TarifAdmin"
            Call msubRecFO(fRS2, fQuery2)
            If fRS.EOF = False Then fTarifAdmin = IIf(IsNull(fRS("TarifAdmin").value), 0, fRS("TarifAdmin").value) Else fTarifAdmin = 0
            Call f_AddBiayaPelayananAdmin(fNoPendaftaran, fKdSubInstalasi, fKdRuangan, fKdPelayananRSAdmin, fKdKelas, "0", CDbl(fTarifAdmin), 1, fTglPelayanan, fNoLab_Rad, fIdPegawai, "01", fKdJenisTarif, 0, CStr(fIdPegawai), Null)
        End If
    End If
    'End Setting Jumlah Biaya Administrasi Per Hari
    Set fRS = Nothing
    fQuery = "insert into BiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "'," & fIdPegawai2 & ",'" & fIdUser & "',null)"
    Call msubRecFO(fRS, fQuery)
    'aktifkan kode berikut jika Paket Pelayanan TM sudah dioperasikan
    'select fKdPaketTM=KdPaket from PasienDaftar where NoPendaftaran=fNoPendaftaran
    'if(fKdPaketTM is not null) and (fKdPaketTM<>'')
    '    insert into BiayaPelayananPaketTM values(fNoPendaftaran,fKdRuangan,fKdPelayananRS,fTglPelayanan,fKdPaketTM,fTarif,fTarifCito,fJmlPelayanan)
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelas,KdKelompokPasien from V_KelasTanggunganPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value) Else fIdPenjamin = "2222222222"
    If fRS.EOF = False Then fKdKelasPenjamin = IIf(IsNull(fRS("KdKelasPenjamin").value), fKdKelas, fRS("KdKelasPenjamin").value) Else fKdKelasPenjamin = fKdKelas
    If fRS.EOF = False Then fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value) Else fKdKelompokPasien = "01"
    Set fRS = Nothing
    fQuery = "select KdPaket from V_PaketNonPaketTanggunganPenjamin where KdPelayananRS='" & fKdPelayananRS & "' and IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdKelas='" & fKdKelasPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdPaket = IIf(IsNull(fRS("KdPaket").value), "030", fRS("KdPaket").value) Else fKdPaket = "030"
    Set fRS = Nothing
    fQuery = "select KdPaket from V_PaketPenjamin where KdPelayananRS='" & fKdPelayananRS & "' and IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdKelas='" & fKdKelasPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdPaketPaket = fRS("KdPaket").value Else fKdPaketPaket = ""
    fTotalTarif = fTarif + fTarifCito
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTM(fNoPendaftaran,fKdPelayananRS,fKdKelasPenjamin,fKdJenisTarif,fStatusCITO,fIdPegawai,fIdPegawai2,null,'C') as TarifCitoPenjamin"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTarifCitoPenjamin = IIf(IsNull(fRS("TarifCitoPenjamin").value), 0, fRS("TarifCitoPenjamin").value) Else fTarifCitoPenjamin = 0
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTM(fNoPendaftaran,fKdPelayananRS,fKdKelasPenjamin,fKdJenisTarif,fStatusCITO,fIdPegawai,fIdPegawai2,null,'T') as TarifKelasPenjamin"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTarifKelasPenjamin = IIf(IsNull(fRS("TarifKelasPenjamin").value), 0, fRS("TarifKelasPenjamin").value) Else fTarifKelasPenjamin = 0
    If fTarifKelasPenjamin = 0 Then fTarifKelasPenjamin = fTarif
    fTotalTarifPenjamin = fTarifCitoPenjamin + fTarifKelasPenjamin
    Set fRS = Nothing
    fQuery = "select PersenTanggunganTM,PersenTanggunganRSTM from PersentaseTPBPTM where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fPersenTanggungan = IIf(IsNull(fRS("PersenTanggunganTM").value), 0, fRS("PersenTanggunganTM").value) Else fPersenTanggungan = 0
    If fRS.EOF = False Then fPersenTanggunganRS = IIf(IsNull(fRS("PersenTanggunganRSTM").value), 0, fRS("PersenTanggunganRSTM").value) Else fPersenTanggunganRS = 0
    'Cek Apakah Ada Penjamin di Paket & Non Paket Asuransi
    Set fRS = Nothing
    fQuery = "select distinct IdPenjamin from V_PaketNonPaketTanggunganPenjamin where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        'Tidak Ada di Paket & Non Paket Asuransi
        Set fRS2 = Nothing
        fQuery2 = "select KdPelayananRS  from DaftarTMNonTanggungan where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS.EOF = True Then
            fJmlHutangPenjamin = (fTotalTarifPenjamin * fPersenTanggungan) / 100
            fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
            If fSisaTagihanPasien > 0 Then
                fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
            Else
                fJmlTanggunganRS = 0
            End If
        Else
            fJmlHutangPenjamin = 0
            fJmlTanggunganRS = 0
        End If
        Set fRS3 = Nothing
        fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & fTarifKelasPenjamin & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "',null)"
        Call msubRecFO(fRS3, fQuery3)
    Else
        'Ada Penjamin di Paket & Non Paket Asuransi
        'Cek Apakah Ada Layanan di Paket & Non Paket Asuransi
        Set fRS2 = Nothing
        fQuery2 = "select IdPenjamin from V_PaketNonPaketTanggunganPenjamin where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelasPenjamin & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then
            'Tidak Ada Layanan di Paket & Non Paket Asuransi
            Set fRS2 = Nothing
            fQuery2 = "select KdPelayananRS  from DaftarTMNonTanggungan where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "'  and KdPelayananRS='" & fKdPelayananRS & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                fJmlHutangPenjamin = (fTotalTarifPenjamin * fPersenTanggungan) / 100
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
            Else
                fJmlHutangPenjamin = (fTotalTarifPenjamin * fPersenTanggungan) / 100
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
            End If
            Set fRS3 = Nothing
            fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & fTarifKelasPenjamin & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "',null)"
            Call msubRecFO(fRS3, fQuery3)
        Else
            'Cek Apakah Ada di Paket Asuransi
            Set fRS2 = Nothing
            fQuery2 = "select IdPenjamin from V_PaketPenjamin where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelasPenjamin & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                'Ada di Non Paket Asuransi
                Set fRS3 = Nothing
                fQuery3 = "select JmlTanggungan from TanggunganAsuransiNonPaket where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelasPenjamin & "'"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fJmlHutangPenjamin = IIf(IsNull(fRS3("JmlTanggungan").value), 0, fRS3("JmlTanggungan").value) Else fJmlHutangPenjamin = 0
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
                Set fRS3 = Nothing
                fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & fTarifKelasPenjamin & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "',null)"
                Call msubRecFO(fRS3, fQuery3)
            Else
                'Ada di Paket Asuransi
                Set fRS3 = Nothing
                fQuery3 = "select sum(Tarif) as TarifSum from V_ListBiayaPelayananPasienStrukNullPaket where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketPaket & "' and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "')"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fTotalBiayaPaket = IIf(IsNull(fRS3("TarifSum").value), 0, fRS3("TarifSum").value) Else fTotalBiayaPaket = 0
                Set fRS3 = Nothing
                fQuery3 = "select JmlTanggungan from V_JmlTanggunganPaketPenjamin where KdPaket='" & fKdPaketPaket & "' and IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdKelas='" & fKdKelasPenjamin & "'"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fTotalTanggunganPaket = IIf(IsNull(fRS3("JmlTanggungan").value), 0, fRS3("JmlTanggungan").value) Else fTotalTanggunganPaket = 0
                If fTotalBiayaPaket = 0 Then
                    fJmlHutangPenjamin = 0
                Else
                    fJmlHutangPenjamin = (CDec(fTotalTarifPenjamin) / CDec(fTotalBiayaPaket)) * CDec(fTotalTanggunganPaket)
                End If
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
                Set fRS3 = Nothing
                fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & fTarif & "," & fTarifCito & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & fTarifKelasPenjamin & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "',null)"
                Call msubRecFO(fRS3, fQuery3)
                'begin of update Tanggungan yg termasuk Paket
                Set fRS = Nothing
                fQuery = "select KdPaket,TglPelayanan from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketPaket & "' and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "')"
                Call msubRecFO(fRS, fQuery)
                While fRS.EOF = False
                    fKdPaketL = fRS("KdPaket").value
                    fTglPelayananL = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
                    Set fRS2 = Nothing
                    fQuery2 = "select KdPelayananRS,TarifKelasPenjamin from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketPaket & "' and day(TglPelayanan)=day('" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "')"
                    Call msubRecFO(fRS2, fQuery2)
                    While fRS2.EOF = False
                        fKdPelayananRSL = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
                        fTarifKelasPenjaminL = IIf(IsNull(fRS("TarifKelasPenjamin").value), 0, fRS("TarifKelasPenjamin").value)
                        fJmlHutangPenjaminL = (CDec(fTarifKelasPenjaminL) / CDec(fTotalBiayaPaket)) * CDec(fTotalTanggunganPaket)
                        Set fRS3 = Nothing
                        fQuery3 = "SELECT  JmlTanggungan FROM TanggunganPaketAsuransiPerTindakan WHERE KdPaket = '" & fKdPaketPaket & "' AND IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' AND KdPelayananRS = '" & fKdPelayananRS & "' AND KdKelas = '" & fKdKelasPenjamin & "'"
                        Call msubRecFO(fRS3, fQuery3)
                        If fRS3.EOF = False Then fJmlHutangPenjaminPPT = IIf(IsNull(fRS3("JmlTanggungan").value), 0, fRS("JmlTanggungan").value) Else fJmlHutangPenjaminPPT = 0
                        If fRS3.EOF = False Then fJmlHutangPenjaminL = fJmlHutangPenjaminPPT
                        fSisaTagihanPasienL = fTotalTarifPenjamin - fJmlHutangPenjaminL
                        If fSisaTagihanPasienL > 0 Then
                            fJmlTanggunganRSL = (fSisaTagihanPasienL * fPersenTanggunganRS) / 100
                        Else
                            fJmlTanggunganRSL = 0
                        End If
                        Set fRS3 = Nothing
                        fQuery3 = "update DetailBiayaPelayanan set JmlHutangPenjamin=" & fJmlHutangPenjaminL & ",JmlTanggunganRS=" & fJmlTanggunganRSL & " where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketL & "' and TglPelayanan='" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRSL & "'"
                        Call msubRecFO(fRS3, fQuery3)
                        fRS2.MoveNext
                    Wend
                    fRS.MoveNext
                Wend
                'end of update Tanggungan yg termasuk Paket
            End If
        End If
    End If
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "'," & fNoLab_Rad & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
    Call f_AddTempHargaKomponenForPenunjangM(fNoPendaftaran, fKdRuangan, fTglPelayanan, fKdPelayananRS, fKdKelas, fKdJenisTarif, CDec(fTarifCito), fJmlPelayanan, fStatusCito, fKdLaboratory)
    Call f_AMDataKunjunganPelayananTMPasienPH(fNoPendaftaran, fKdRuangan, fKdRuanganAsal, fTglPelayanan, fKdPelayananRS, fIdPenjamin, fKdKelompokPasien, fJmlPelayanan, fNoLab_Rad, "A")
    If fKdLaboratory <> "" Then
        Set fRS = Nothing
        fQuery = "insert into DetailPelayananLaboratororium values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & ",'" & fKdLaboratory & "')"
        Call msubRecFO(fRS, fQuery)
    End If
End Function

'Konversi dari SP: Add_DetailBiayaPelayanan
Public Function f_AddDetailBiayaPelayanan(fNoPendaftaran As String, fKdSubInstalasi As String, fKdRuangan As String, fKdPelayananRS As String, fKdKelas As String, fStatusCito As String, fTarif As Double, fJmlPelayanan As Integer, fTglPelayanan As Date, fNoLab_Rad As Variant, fIdPegawai As Variant, fStatusAPBD As String, fKdJenisTarif As String, fTarifCito As Integer, fIdUser As String, fIdPegawai2 As Variant, fIdPegawai3 As Variant)
    Dim fIdPenjamin As String
    Dim fKdKelasPenjamin As String
    Dim fKdKelompokPasien As String
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fJmlTanggunganRSL As Currency
    Dim fPersenTanggungan As Double
    Dim fPersenTanggunganRS As Double
    Dim fTotalTarif As Currency
    Dim fTarifKelasPenjamin As Currency
    Dim fTarifCitoKelasPenjamin As Currency
    Dim fPersenTarifCito As Double
    Dim fTarifCitoPenjamin As Currency
    Dim fTotalTarifPenjamin As Currency
    Dim fKdPaket As Variant
    Dim fTotalBiayaPaket As Currency
    Dim fTotalTanggunganPaket As Currency
    Dim fKdPaketL As String
    Dim fTarifKelasPenjaminL As Currency
    Dim fJmlHutangPenjaminL As Currency
    Dim fKdPelayananRSL As String
    Dim fTglPelayananL As Date
    Dim fJmlHutangPenjaminPPT As Currency
    Dim fKdPaketTM As String
    Dim fKdPaketPaket As String
    Dim fSisaTagihanPasien As Currency
    Dim fSisaTagihanPasienL As Currency

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelas,KdKelompokPasien from V_KelasTanggunganPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value) Else fIdPenjamin = "2222222222"
    If fRS.EOF = False Then fKdKelasPenjamin = IIf(IsNull(fRS("KdKelas").value), fKdKelas, fRS("KdKelas").value) Else fKdKelasPenjamin = fKdKelas
    If fRS.EOF = False Then fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value) Else fKdKelompokPasien = "01"
    Set fRS = Nothing
    fQuery = "select KdPaket from V_PaketNonPaketTanggunganPenjamin where KdPelayananRS='" & fKdPelayananRS & "' and IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdKelas='" & fKdKelasPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdPaket = IIf(IsNull(fRS("KdPaket").value), "030", fRS("KdPaket").value) Else fKdPaket = "030"
    Set fRS = Nothing
    fQuery = "select KdPaket from V_PaketPenjamin where KdPelayananRS='" & fKdPelayananRS & "' and IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdKelas='" & fKdKelasPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdPaketPaket = fRS("KdPaket").value Else fKdPaketPaket = ""
    fTotalTarif = fTarif + fTarifCito
    Set fRS = Nothing

    fQuery = "select dbo.FB_NewTakeTarifBPTM('" & fNoPendaftaran & "','" & fKdPelayananRS & "','" & fKdKelasPenjamin & "','" & fKdJenisTarif & "','" & fStatusCito & "'," & fIdPegawai & "," & fIdPegawai2 & ",null,'C') as TarifCitoPenjamin"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTarifCitoPenjamin = IIf(IsNull(fRS("TarifCitoPenjamin").value), 0, fRS("TarifCitoPenjamin").value) Else fTarifCitoPenjamin = 0
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTM('" & fNoPendaftaran & "' ,'" & fKdPelayananRS & "' ,'" & fKdKelasPenjamin & "','" & fKdJenisTarif & "','" & fStatusCito & "'," & fIdPegawai & "," & fIdPegawai2 & ",null,'T') as TarifKelasPenjamin"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTarifKelasPenjamin = IIf(IsNull(fRS("TarifKelasPenjamin").value), 0, fRS("TarifKelasPenjamin").value) Else fTarifKelasPenjamin = 0
    If fTarifKelasPenjamin = 0 Then fTarifKelasPenjamin = fTarif
    fTotalTarifPenjamin = fTarifCitoPenjamin + fTarifKelasPenjamin
    Set fRS = Nothing
    fQuery = "select PersenTanggunganTM,PersenTanggunganRSTM from PersentaseTPBPTM where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fPersenTanggungan = IIf(IsNull(fRS("PersenTanggunganTM").value), 0, fRS("PersenTanggunganTM").value) Else fPersenTanggungan = 0
    If fRS.EOF = False Then fPersenTanggunganRS = IIf(IsNull(fRS("PersenTanggunganRSTM").value), 0, fRS("PersenTanggunganRSTM").value) Else fPersenTanggunganRS = 0
    'Cek Apakah Ada Penjamin di Paket & Non Paket Asuransi
    Set fRS = Nothing
    fQuery = "select distinct IdPenjamin from V_PaketNonPaketTanggunganPenjamin where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        'Tidak Ada di Paket & Non Paket Asuransi
        Set fRS2 = Nothing
        fQuery2 = "select KdPelayananRS  from DaftarTMNonTanggungan where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS.EOF = True Then
            fJmlHutangPenjamin = (fTotalTarifPenjamin * fPersenTanggungan) / 100
            fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
            If fSisaTagihanPasien > 0 Then
                fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
            Else
                fJmlTanggunganRS = 0
            End If
        Else
            fJmlHutangPenjamin = 0
            fJmlTanggunganRS = 0
        End If
        Set fRS3 = Nothing
        fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & msubKonversiKomaTitik(CStr(fTarif)) & "," & msubKonversiKomaTitik(CStr(fTarifCito)) & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & msubKonversiKomaTitik(CStr(fTarifKelasPenjamin)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjamin)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRS)) & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "'," & fIdPegawai3 & ")"

        Call msubRecFO(fRS3, fQuery3)
    Else
        'Ada Penjamin di Paket & Non Paket Asuransi
        'Cek Apakah Ada Layanan di Paket & Non Paket Asuransi
        Set fRS2 = Nothing
        fQuery2 = "select IdPenjamin from V_PaketNonPaketTanggunganPenjamin where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelasPenjamin & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then
            'Tidak Ada Layanan di Paket & Non Paket Asuransi
            Set fRS2 = Nothing
            fQuery2 = "select KdPelayananRS  from DaftarTMNonTanggungan where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "'  and KdPelayananRS='" & fKdPelayananRS & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                fJmlHutangPenjamin = (fTotalTarifPenjamin * fPersenTanggungan) / 100
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
            Else
                fJmlHutangPenjamin = (fTotalTarifPenjamin * fPersenTanggungan) / 100
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
            End If
            Set fRS3 = Nothing
            fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & msubKonversiKomaTitik(CStr(fTarif)) & "," & msubKonversiKomaTitik(CStr(fTarifCito)) & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & msubKonversiKomaTitik(CStr(fTarifKelasPenjamin)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjamin)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRS)) & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "'," & fIdPegawai3 & ")"
            Call msubRecFO(fRS3, fQuery3)
        Else
            'Cek Apakah Ada di Paket Asuransi
            Set fRS2 = Nothing
            fQuery2 = "select IdPenjamin from V_PaketPenjamin where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelasPenjamin & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                'Ada di Non Paket Asuransi
                Set fRS3 = Nothing
                fQuery3 = "select JmlTanggungan from TanggunganAsuransiNonPaket where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelasPenjamin & "'"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fJmlHutangPenjamin = IIf(IsNull(fRS3("JmlTanggungan").value), 0, fRS3("JmlTanggungan").value) Else fJmlHutangPenjamin = 0
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
                Set fRS3 = Nothing
                fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & msubKonversiKomaTitik(CStr(fTarif)) & "," & msubKonversiKomaTitik(CStr(fTarifCito)) & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & msubKonversiKomaTitik(CStr(fTarifKelasPenjamin)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjamin)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRS)) & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "'," & fIdPegawai3 & ")"
                Call msubRecFO(fRS3, fQuery3)
            Else
                'Ada di Paket Asuransi
                Set fRS3 = Nothing
                fQuery3 = "select sum(Tarif) as TarifSum from V_ListBiayaPelayananPasienStrukNullPaket where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketPaket & "' and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "')"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fTotalBiayaPaket = IIf(IsNull(fRS3("TarifSum").value), 0, fRS3("TarifSum").value) Else fTotalBiayaPaket = 0
                Set fRS3 = Nothing
                fQuery3 = "select JmlTanggungan from V_JmlTanggunganPaketPenjamin where KdPaket='" & fKdPaketPaket & "' and IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdKelas='" & fKdKelasPenjamin & "'"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = False Then fTotalTanggunganPaket = IIf(IsNull(fRS3("JmlTanggungan").value), 0, fRS3("JmlTanggungan").value) Else fTotalTanggunganPaket = 0
                If fTotalBiayaPaket = 0 Then
                    fJmlHutangPenjamin = 0
                Else
                    fJmlHutangPenjamin = (CDec(fTotalTarifPenjamin) / CDec(fTotalBiayaPaket)) * CDec(fTotalTanggunganPaket)
                End If
                fSisaTagihanPasien = fTotalTarifPenjamin - fJmlHutangPenjamin
                If fSisaTagihanPasien > 0 Then
                    fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
                Else
                    fJmlTanggunganRS = 0
                End If
                Set fRS3 = Nothing
                fQuery3 = "insert into DetailBiayaPelayanan values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fStatusCito & "'," & msubKonversiKomaTitik(CStr(fTarif)) & "," & msubKonversiKomaTitik(CStr(fTarifCito)) & "," & fJmlPelayanan & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLab_Rad & "," & fIdPegawai & ",null,'" & fStatusAPBD & "','" & fKdJenisTarif & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & msubKonversiKomaTitik(CStr(fTarifKelasPenjamin)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjamin)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRS)) & ",0," & fKdPaket & "," & fIdPegawai2 & ",'" & fIdUser & "'," & fIdPegawai3 & ")"
                Call msubRecFO(fRS3, fQuery3)
                'begin of update Tanggungan yg termasuk Paket
                Set fRS = Nothing
                fQuery = "select KdPaket,TglPelayanan from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketPaket & "' and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "')"
                Call msubRecFO(fRS, fQuery)
                While fRS.EOF = False
                    fKdPaketL = fRS("KdPaket").value
                    fTglPelayananL = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
                    Set fRS2 = Nothing
                    fQuery2 = "select KdPelayananRS,TarifKelasPenjamin from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketPaket & "' and day(TglPelayanan)=day('" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "')"
                    Call msubRecFO(fRS2, fQuery2)
                    While fRS2.EOF = False
                        fKdPelayananRSL = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
                        fTarifKelasPenjaminL = IIf(IsNull(fRS("TarifKelasPenjamin").value), 0, fRS("TarifKelasPenjamin").value)
                        fJmlHutangPenjaminL = (CDec(fTarifKelasPenjaminL) / CDec(fTotalBiayaPaket)) * CDec(fTotalTanggunganPaket)
                        Set fRS3 = Nothing
                        fQuery3 = "SELECT  JmlTanggungan FROM TanggunganPaketAsuransiPerTindakan WHERE KdPaket = '" & fKdPaketPaket & "' AND IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' AND KdPelayananRS = '" & fKdPelayananRS & "' AND KdKelas = '" & fKdKelasPenjamin & "'"
                        Call msubRecFO(fRS3, fQuery3)
                        If fRS3.EOF = False Then fJmlHutangPenjaminPPT = IIf(IsNull(fRS3("JmlTanggungan").value), 0, fRS("JmlTanggungan").value) Else fJmlHutangPenjaminPPT = 0
                        If fRS3.EOF = False Then fJmlHutangPenjaminL = fJmlHutangPenjaminPPT
                        fSisaTagihanPasienL = fTotalTarifPenjamin - fJmlHutangPenjaminL
                        If fSisaTagihanPasienL > 0 Then
                            fJmlTanggunganRSL = (fSisaTagihanPasienL * fPersenTanggunganRS) / 100
                        Else
                            fJmlTanggunganRSL = 0
                        End If
                        Set fRS3 = Nothing
                        fQuery3 = "update DetailBiayaPelayanan set JmlHutangPenjamin=" & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminL)) & ",JmlTanggunganRS=" & msubKonversiKomaTitik(CStr(fJmlTanggunganRSL)) & " where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket='" & fKdPaketL & "' and TglPelayanan='" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRSL & "'"
                        Call msubRecFO(fRS3, fQuery3)
                        fRS2.MoveNext
                    Wend
                    fRS.MoveNext
                Wend
                'end of update Tanggungan yg termasuk Paket
            End If
        End If
    End If
    Call f_AddTempHargaKomponen(fNoPendaftaran, fKdRuangan, fTglPelayanan, fKdPelayananRS, fKdKelas, fKdJenisTarif, CDec(fTarifCito), fJmlPelayanan, fStatusCito, CStr(fIdPegawai))
End Function

'Konversi dari SP: Add_DetailBiayaPelayananOnUbahJenisPasien
Public Function f_AddDetailBiayaPelayananOnUbahJenisPasien(fNoPendaftaran As String)
    Dim fKdSubInstalasi As String
    Dim fKdRuangan As String
    Dim fKdPelayananRS As String
    Dim fKdKelas As String
    Dim fStatusCito As String
    Dim fTarif As Double
    Dim fJmlPelayanan As Integer
    Dim fTglPelayanan As Date
    Dim fNoLab_Rad As Variant
    Dim fIdPegawai As Variant
    Dim fStatusAPBD As String
    Dim fKdJenisTarif As String
    Dim fTarifCito As Integer
    Dim fIdPegawai2 As Variant
    Dim fIdUser As String
    Dim fIdPegawai3 As Variant
    Dim fxCounter As Integer

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "delete from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    Set fRS = Nothing
    fQuery = "select NoPendaftaran,KdSubInstalasi,KdRuangan,KdPelayananRS,KdKelas,StatusCITO,Tarif,TarifCito,JmlPelayanan,TglPelayanan,NoLab_Rad,IdPegawai,StatusAPBD,KdJenisTarif,IdPegawai2,IdUser,IdPegawai3 from BiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    fxCounter = 0
    While fRS.EOF = False
        fxCounter = fxCounter + 1
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
        fStatusCito = IIf(IsNull(fRS("StatusCITO").value), "", fRS("StatusCITO").value)
        fTarif = IIf(IsNull(fRS("Tarif").value), 0, fRS("Tarif").value)
        fTarifCito = IIf(IsNull(fRS("TarifCito").value), 0, fRS("TarifCito").value)
        fJmlPelayanan = IIf(IsNull(fRS("JmlPelayanan").value), 0, fRS("JmlPelayanan").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fNoLab_Rad = IIf(IsNull(fRS("NoLab_Rad").value), "null", "'" & fRS("NoLab_Rad").value & "'")
        fIdPegawai = IIf(IsNull(fRS("IdPegawai").value), "null", "'" & fRS("IdPegawai").value & "'")
        fStatusAPBD = IIf(IsNull(fRS("StatusAPBD").value), "", fRS("StatusAPBD").value)
        fKdJenisTarif = IIf(IsNull(fRS("KdJenisTarif").value), "", fRS("KdJenisTarif").value)
        fIdPegawai2 = IIf(IsNull(fRS("IdPegawai2").value), "null", "'" & fRS("IdPegawai2").value & "'")
        fIdUser = IIf(IsNull(fRS("IdUser").value), "", fRS("IdUser").value)
        fIdPegawai3 = IIf(IsNull(fRS("IdPegawai3").value), "null", "'" & fRS("IdPegawai3").value & "'")
        If fxCounter < 2 Then
            Call f_DeleteTempHargaKomponen(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan)
            Set fRS2 = Nothing
            fQuery2 = "delete from TempHargaKomponen  where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
        End If
        Call f_AddDetailBiayaPelayanan(fNoPendaftaran, fKdSubInstalasi, fKdRuangan, fKdPelayananRS, fKdKelas, fStatusCito, fTarif, fJmlPelayanan, fTglPelayanan, fNoLab_Rad, fIdPegawai, fStatusAPBD, fKdJenisTarif, fTarifCito, fIdUser, fIdPegawai2, fIdPegawai3)
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Add_DetailPemakaianObatAlkesOnUbahJenisPasien
Public Function f_AddDetailPemakaianObatAlkesOnUbahJenisPasien(fNoPendaftaran As String)
    Dim fKdBarang As String
    Dim fKdAsal As String
    Dim fKdRuangan As String
    Dim fSatuan As String
    Dim fJmlBrg As Double
    Dim fKdSubInstalasi As String
    Dim fKdKelas As String
    Dim fTglPelayanan As Date
    Dim fNoLabRad As Variant
    Dim fIdDokter As String
    Dim fIdPenjamin As String
    Dim fKdKelasPenjamin As String
    Dim fKdKelompokPasien As String
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fPersenTanggungan As Double
    Dim fPersenTanggunganRS As Double
    Dim fIdPenjaminTemp As String
    Dim fTarifKelasPenjamin As Currency
    Dim fIdPegawai2 As Variant
    Dim fIdUser As String
    Dim fHargaSatuan As Currency
    Dim fHargaBeli As Currency
    Dim fJmlService As Integer
    Dim fTarifService As Currency
    Dim fNoResep As Variant
    Dim fHargaTotal As Currency
    Dim fKdJenisObat As Variant
    Dim fBiayaAdministrasi As Currency
    Dim fStatusStok As String
    Dim fKdRuanganAsal As String
    Dim fSisaTagihanPasien As Currency
    Dim fx As Integer

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "delete from DetailPemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    fx = 0
    Set fRS = Nothing
    fQuery = "select NoPendaftaran,KdSubInstalasi,KdRuangan,KdKelas,KdBarang,KdAsal,JmlBarang,TglPelayanan,NoLab_Rad,IdPegawai,SatuanJml,IdPegawai2,IdUser,JmlService,TarifService,NoResep,HargaSatuan,HargaBeli,KdJenisObat,BiayaAdministrasi,StatusStok,KdRuanganAsal from PemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fx = fx + 1
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
        fKdBarang = IIf(IsNull(fRS("KdBarang").value), "", fRS("KdBarang").value)
        fKdAsal = IIf(IsNull(fRS("KdAsal").value), "", fRS("KdAsal").value)
        fJmlBrg = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fNoLabRad = IIf(IsNull(fRS("NoLab_Rad").value), "null", "'" & fRS("NoLab_Rad").value & "'")
        fIdDokter = IIf(IsNull(fRS("IdPegawai").value), "", fRS("IdPegawai").value)
        fSatuan = IIf(IsNull(fRS("SatuanJml").value), "", fRS("SatuanJml").value)
        fIdPegawai2 = IIf(IsNull(fRS("IdPegawai2").value), "null", "'" & fRS("IdPegawai2").value & "'")
        fIdUser = IIf(IsNull(fRS("IdUser").value), "", fRS("IdUser").value)
        fJmlService = IIf(IsNull(fRS("JmlService").value), 0, fRS("JmlService").value)
        fTarifService = IIf(IsNull(fRS("TarifService").value), 0, fRS("TarifService").value)
        fNoResep = IIf(IsNull(fRS("NoResep").value), "null", "'" & fRS("NoResep").value & "'")
        fHargaSatuan = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
        fHargaBeli = IIf(IsNull(fRS("HargaBeli").value), 0, fRS("HargaBeli").value)
        fKdJenisObat = IIf(IsNull(fRS("KdJenisObat").value), "null", "'" & fRS("KdJenisObat").value & "'")
        fBiayaAdministrasi = IIf(IsNull(fRS("BiayaAdministrasi").value), 0, fRS("BiayaAdministrasi").value)
        fStatusStok = IIf(IsNull(fRS("StatusStok").value), "", fRS("StatusStok").value)
        fKdRuanganAsal = IIf(IsNull(fRS("KdRuanganAsal").value), "", fRS("KdRuanganAsal").value)
        Set fRS2 = Nothing
        fQuery2 = "select IdPenjamin,KdKelas,KdKelompokPasien from V_KelasTanggunganPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fIdPenjamin = IIf(IsNull(fRS2("IdPenjamin").value), "2222222222", fRS2("IdPenjamin").value) Else fIdPenjamin = "2222222222"
        If fRS2.EOF = False Then fKdKelasPenjamin = IIf(IsNull(fRS2("KdKelas").value), fKdKelas, fRS2("KdKelas").value) Else fKdKelasPenjamin = fKdKelas
        If fRS2.EOF = False Then fKdKelompokPasien = IIf(IsNull(fRS2("KdKelompokPasien").value), "01", fRS2("KdKelompokPasien").value) Else fKdKelompokPasien = "01"
        fTarifKelasPenjamin = fHargaSatuan + fTarifService + fBiayaAdministrasi
        fHargaTotal = fHargaSatuan + fTarifService + fBiayaAdministrasi
        Set fRS2 = Nothing
        fQuery2 = "select PersenTanggunganOA,PersenTanggunganRSOA from PersentaseTPBPOA where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdAsal='" & fKdAsal & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fPersenTanggungan = IIf(IsNull(fRS2("PersenTanggunganOA").value), 0, fRS2("PersenTanggunganOA").value) Else fPersenTanggungan = 0
        If fRS2.EOF = False Then fPersenTanggunganRS = IIf(IsNull(fRS2("PersenTanggunganRSOA").value), 0, fRS2("PersenTanggunganRSOA").value) Else fPersenTanggunganRS = 0
        Set fRS2 = Nothing
        fQuery2 = "select IdPenjamin from DaftarOANonTanggungan where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then
            fJmlHutangPenjamin = (fHargaTotal * fPersenTanggungan) / 100
            fSisaTagihanPasien = fHargaTotal - fJmlHutangPenjamin
            If fSisaTagihanPasien > 0 Then
                fJmlTanggunganRS = (fSisaTagihanPasien * fPersenTanggunganRS) / 100
            Else
                fJmlTanggunganRS = 0
            End If
        Else
            fJmlHutangPenjamin = 0
            fJmlTanggunganRS = 0
        End If
        Set fRS2 = Nothing
        fQuery2 = "insert into DetailPemakaianAlkes values('" & fNoPendaftaran & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdKelas & "','" & fKdBarang & "','" & fKdAsal & "'," & fJmlBrg & "," & msubKonversiKomaTitik(CStr(fHargaSatuan)) & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'," & fNoLabRad & ",null,'" & fIdDokter & "','" & fSatuan & "','" & fIdPenjamin & "','" & fKdKelasPenjamin & "'," & msubKonversiKomaTitik(CStr(fTarifKelasPenjamin)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjamin)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRS)) & ",0," & msubKonversiKomaTitik(CStr(fHargaBeli)) & "," & fIdPegawai2 & ",'" & fIdUser & "',null,null," & fJmlService & "," & msubKonversiKomaTitik(CStr(fTarifService)) & "," & fNoResep & "," & msubKonversiKomaTitik(CStr(fBiayaAdministrasi)) & ",'" & fStatusStok & "','" & fKdRuanganAsal & "')"
        Call msubRecFO(fRS2, fQuery2)
        If fx < 2 Then
            Call f_DeleteTempHargaKomponenObatAlkes(fNoPendaftaran, fKdBarang, fTglPelayanan, fKdRuangan, fKdAsal, fSatuan)
            Set fRS2 = Nothing
            fQuery2 = "delete from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
        End If
        Call f_AddTempHargaKomponenOAResep(fNoPendaftaran, fKdRuangan, fTglPelayanan, fKdBarang, fKdAsal, fSatuan, fHargaSatuan, fHargaBeli, fJmlBrg, fKdJenisObat, fJmlService, fTarifService, fNoResep, fBiayaAdministrasi, fKdRuanganAsal)
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Add_PembatalanStrukPelayananKasir
Public Function f_AddPembatalanStrukPelayananKasir(fNoBKM As String, fNoStruk As String, fPembayaranKe As Integer, fKdRuangan As String, fIdUser As String)
    Dim fNoBKMTempOA As String
    Dim fNoPendaftaran As String
    Dim fNoCM As String
    Dim fMaxPembayaranKe As Integer
    Dim fSisaTagihanTM As Currency
    Dim fSisaTagihanOA As Currency
    Dim fStatusPiutang As String
    Dim fJmlBayar As Currency
    Dim fSisaTagihan As Currency
    Dim fBackSisaTagihanTM As Currency
    Dim fBackSisaTagihanOA As Currency
    Dim fNoRiwayat As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeNoRiwayat('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "')"
    Call msubRecFO(fRS, fQuery)
    fNoRiwayat = fRS.Fields(0)

    Set fRS = Nothing
    fQuery = "select max(PembayaranKe) as PembayaranKeMax from PembayaranTagihanPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fMaxPembayaranKe = IIf(IsNull(fRS("PembayaranKeMax").value), 1, fRS("PembayaranKeMax").value) Else fMaxPembayaranKe = 1
    If fPembayaranKe = 1 And fMaxPembayaranKe = 1 Then
        Set fRS = Nothing
        fQuery = "select NoPendaftaran,NoCM from StrukPelayananPasien where NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value) Else fNoPendaftaran = ""
        If fRS.EOF = False Then fNoCM = IIf(IsNull(fRS("NoCM").value), "", fRS("NoCM").value) Else fNoCM = ""
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from PasienBelumBayar where NoPendaftaran='" & fNoPendaftaran & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            Set fRS = Nothing
            fQuery = "insert into PasienBelumBayar values('" & fNoPendaftaran & "','" & fNoCM & "')"
            Call msubRecFO(fRS, fQuery)
        End If
        Set fRS = Nothing
        fQuery = "delete from PasienSudahBayar where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "update BiayaPelayanan set NoStruk=null where NoStruk='" & fNoStruk & "' and NoPendaftaran = '" & fNoPendaftaran & "'"
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "update DetailBiayaPelayanan set NoStruk=null where NoStruk='" & fNoStruk & "' and NoPendaftaran = '" & fNoPendaftaran & "'"
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "update TempHargaKomponen set NoStruk=null where NoStruk='" & fNoStruk & "' and NoPendaftaran = '" & fNoPendaftaran & "'"
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "update PemakaianAlkes set NoStruk=null where NoStruk='" & fNoStruk & "' and NoPendaftaran = '" & fNoPendaftaran & "'"
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "update DetailPemakaianAlkes set NoStruk=null where NoStruk='" & fNoStruk & "' and NoPendaftaran = '" & fNoPendaftaran & "'"
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "update TempHargaKomponenObatAlkes set NoStruk=null where NoStruk='" & fNoStruk & "' and NoPendaftaran = '" & fNoPendaftaran & "'"
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "Delete ConvertSBKMToNoPendaftaran WHERE NoPendaftaran='" & fNoPendaftaran & "' AND NoBKM='" & fNoBKM & "'"
        Call msubRecFO(fRS, fQuery)

        Set fRS = Nothing
        fQuery = "Update StrukPelayananPasien set NoRiwayat = '" & fNoRiwayat & "' where NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS, fQuery)

        Set fRS = Nothing
        fQuery = "Update StrukBuktiKasMasuk set NoRiwayat = '" & fNoRiwayat & "' where NoBKM='" & fNoBKM & "'"
        Call msubRecFO(fRS, fQuery)
    Else
        Set fRS = Nothing
        fQuery = "select JmlBayar from StrukBuktiKasMasuk where NoBKM='" & fNoBKM & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fJmlBayar = IIf(IsNull(fRS("JmlBayar").value), 0, fRS("JmlBayar").value) Else fJmlBayar = 0
        Set fRS = Nothing
        fQuery = "select StatusPiutang,SisaTagihan from PembayaranTagihanPasien where NoBKM='" & fNoBKM & "' and PembayaranKe=" & fPembayaranKe & "-1"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fStatusPiutang = IIf(IsNull(fRS("StatusPiutang").value), 0, fRS("StatusPiutang").value) Else fStatusPiutang = 0
        If fRS.EOF = False Then fSisaTagihan = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value) Else fSisaTagihan = 0
        If fSisaTagihan <> 0 Then
            If fStatusPiutang = "OA" Then
                Set fRS = Nothing
                fQuery = "update TotalBiayaPelayananOA set SisaTagihan=SisaTagihan + " & fJmlBayar & " where NoStruk='" & fNoStruk & "'"
                Call msubRecFO(fRS, fQuery)
            End If
            If fStatusPiutang = "TM" Then
                Set fRS = Nothing
                fQuery = "update TotalBiayaPelayananTM set SisaTagihan=SisaTagihan + " & fJmlBayar & " where NoStruk='" & fNoStruk & "'"
                Call msubRecFO(fRS, fQuery)
            End If
            If fStatusPiutang = "SA" Then
                Set fRS = Nothing
                fQuery = "select SisaTagihan from TotalBiayaPelayananTM where NoStruk='" & fNoStruk & "'"
                Call msubRecFO(fRS, fQuery)
                If fRS.EOF = False Then fSisaTagihanTM = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value) Else fSisaTagihanTM = 0
                Set fRS = Nothing
                fQuery = "select SisaTagihan from TotalBiayaPelayananOA where NoStruk='" & fNoStruk & "'"
                Call msubRecFO(fRS, fQuery)
                If fRS.EOF = False Then fSisaTagihanOA = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value) Else fSisaTagihanOA = 0
                fBackSisaTagihanTM = (CDec(fSisaTagihanTM) / CDec(fSisaTagihan)) * CDec(fJmlBayar)
                fBackSisaTagihanOA = (CDec(fSisaTagihanOA) / CDec(fSisaTagihan)) * CDec(fJmlBayar)
                Set fRS = Nothing
                fQuery = "update TotalBiayaPelayananTM set SisaTagihan=SisaTagihan + " & fBackSisaTagihanTM & " where NoStruk='" & fNoStruk & "'"
                Call msubRecFO(fRS, fQuery)
                Set fRS = Nothing
                fQuery = "update TotalBiayaPelayananOA set SisaTagihan=SisaTagihan + " & fBackSisaTagihanOA & " where NoStruk='" & fNoStruk & "'"
                Call msubRecFO(fRS, fQuery)
            End If
        End If
        Set fRS = Nothing
        fQuery = "delete from  PembayaranTagihanPasien where NoBKM='" & fNoBKM & "' and PembayaranKe=" & fPembayaranKe & ""
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "delete from StrukBuktiKasMasuk where NoBKM='" & fNoBKM & "'"
        Call msubRecFO(fRS, fQuery)
    End If
    Set fRS = Nothing
    fQuery = "select NoBKM from RekapKomponenBiayaPelayananTM where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        Call f_DeleteRekapKomponenBiayaPelayananTM(fNoBKM, fNoStruk, "M")
        Call f_DeleteRekapKomponenBPRemunerasiTM(fNoBKM, fNoStruk, "M")
    End If
    Set fRS = Nothing
    fQuery = "select NoBKM from RekapKomponenBiayaPelayananOA where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        Call f_DeleteRekapKomponenBiayaPelayananOA(fNoBKM, fNoStruk, "M")
        Call f_DeleteRekapKomponenBPRemunerasiOA(fNoBKM, fNoStruk, "M")
    End If
End Function

'Konversi dari SP: Add_PembatalanStrukPelayananKasirApotik
Public Function f_AddPembatalanStrukPelayananKasirApotik(fNoStruk As String, fNoBKM As String, fPembayaranKe As Integer, fKdRuangan As String, fIdUser As String)

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select NoBKM from RekapKomponenBiayaPelayananApotik where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        Call f_DeleteRekapKomponenBiayaPelayananApotik(fNoBKM, fNoStruk, "M")
        Call f_DeleteRekapKomponenBPRemunerasiApotik(fNoBKM, fNoStruk, "M")
    End If
    Set fRS = Nothing
    fQuery = "delete from PembayaranTagihanPasien where NoBKM='" & fNoBKM & "' and PembayaranKe=" & fPembayaranKe & ""
    Call msubRecFO(fRS, fQuery)
    Set fRS = Nothing
    fQuery = "delete from StrukBuktiKasMasuk where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
End Function

'Konversi dari SP: Add_StrukPelayananPasien
Public Function f_AddStrukPelayananPasienDetail(fNoBKM As String, fNoStruk As String, fNoPendaftaran As String, fNoCM As String, fJmlBayar As Currency, fJmlDiscount As Currency, fSisaTagihan As Currency, fStatusBayar As String, fTotalBiayaOA As Currency, fJmlBayarOA As Currency, fJmlHutangPenjaminOA As Currency, fJmlTanggunganRSOA As Currency, fJmlPembebasanOA As Currency, fJmlHrsDibayarOA As Currency, fJmlDiscountOA As Currency, fSisaTagihanOA As Currency, fTotalBiayaTM As Currency, fJmlBayarTM As Currency, fJmlHutangPenjaminTM As Currency, fJmlTanggunganRSTM As Currency, fJmlPembebasanTM As Currency, fJmlHrsDibayarTM As Currency, fJmlDiscountTM As Currency, fSisaTagihanTM As Currency)
    'fStatusPiutang: TM=Tindakan Medis, OA=Obat & Alkes, SA=Semua; fStatusBayar: 1=Dibayar Semua; 0=Ada Yang Tidak Dibayar
    'fStatusBayarSemua: Y=Tindakan dibayar semua, T=Tindakan ada yang belum dibayar
    Dim fTglMasuk As Date
    Dim fTglKeluar As Date
    Dim fKdRuanganAkhir As String
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    If fStatusBayar = "1" Then
        If UCase(fStatusPiutang) = "TM" Then
            Set fRS = Nothing
            fQuery = "update BiayaPelayanan set NoStruk='" & fNoStruk & "' where NoPendaftaran='" & fNoPendaftaran & "' AND NoStruk IS NULL"
            Call msubRecFO(fRS, fQuery)
            Set fRS = Nothing
            fQuery = "update DetailBiayaPelayanan set NoStruk='" & fNoStruk & "' where NoPendaftaran='" & fNoPendaftaran & "' AND NoStruk IS NULL"
            Call msubRecFO(fRS, fQuery)
            Set fRS = Nothing
            fQuery = "update TempHargaKomponen set NoStruk='" & fNoStruk & "' where NoPendaftaran='" & fNoPendaftaran & "' AND NoStruk IS NULL"
            Call msubRecFO(fRS, fQuery)
        End If
        If UCase(fStatusPiutang) = "OA" Then
            Set fRS = Nothing
            fQuery = "update PemakaianAlkes set NoStruk='" & fNoStruk & "' where NoPendaftaran='" & fNoPendaftaran & "' AND NoStruk IS NULL"
            Call msubRecFO(fRS, fQuery)
            Set fRS = Nothing
            fQuery = "update TempHargaKomponenObatAlkes set NoStruk='" & fNoStruk & "' where NoPendaftaran='" & fNoPendaftaran & "' AND NoStruk IS NULL"
            Call msubRecFO(fRS, fQuery)
            Set fRS = Nothing
            fQuery = "update DetailPemakaianAlkes set NoStruk='" & fNoStruk & "' where NoPendaftaran='" & fNoPendaftaran & "' AND NoStruk IS NULL"
            Call msubRecFO(fRS, fQuery)
        End If
        If UCase(fStatusPiutang) = "SA" Then
            Set fRS = Nothing
            fQuery = "update BiayaPelayanan set NoStruk='" & fNoStruk & "' where NoPendaftaran='" & fNoPendaftaran & "' AND NoStruk IS NULL"
            Call msubRecFO(fRS, fQuery)
            Set fRS = Nothing
            fQuery = "update DetailBiayaPelayanan set NoStruk='" & fNoStruk & "' where NoPendaftaran='" & fNoPendaftaran & "' AND NoStruk IS NULL"
            Call msubRecFO(fRS, fQuery)
            Set fRS = Nothing
            fQuery = "update PemakaianAlkes set NoStruk='" & fNoStruk & "' where NoPendaftaran='" & fNoPendaftaran & "' AND NoStruk IS NULL"
            Call msubRecFO(fRS, fQuery)
            Set fRS = Nothing
            fQuery = "update TempHargaKomponenObatAlkes set NoStruk='" & fNoStruk & "' where NoPendaftaran='" & fNoPendaftaran & "' AND NoStruk IS NULL"
            Call msubRecFO(fRS, fQuery)
            Set fRS = Nothing
            fQuery = "update DetailPemakaianAlkes set NoStruk='" & fNoStruk & "' where NoPendaftaran='" & fNoPendaftaran & "' AND NoStruk IS NULL"
            Call msubRecFO(fRS, fQuery)
            Set fRS = Nothing
            fQuery = "update TempHargaKomponen set NoStruk='" & fNoStruk & "' where NoPendaftaran='" & fNoPendaftaran & "' AND NoStruk IS NULL"
            Call msubRecFO(fRS, fQuery)
        End If
    End If
    If UCase(fStatusPiutang) = "TM" Then
        Set fRS = Nothing

        fQuery = "insert into TotalBiayaPelayananTM values('" & fNoStruk & "'," & msubKonversiKomaTitik(CStr(fTotalBiayaTM)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTM)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTM)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanTM)) & "," & msubKonversiKomaTitik(CStr(fJmlHrsDibayarTM)) & "," & msubKonversiKomaTitik(CStr(fJmlDiscountTM)) & "," & msubKonversiKomaTitik(CStr(fJmlBayarTM)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanTM)) & ")"
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "insert into PembayaranTagihanPasien values('" & fNoBKM & "','" & fNoStruk & "',0," & msubKonversiKomaTitik(CStr(fJmlBayarTM)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanTM)) & "+" & msubKonversiKomaTitik(CStr(fJmlDiscountTM)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanTM)) & ",1,'TM')"
        Call msubRecFO(fRS, fQuery)
    End If
    If UCase(fStatusPiutang) = "OA" Then
        Set fRS = Nothing
        fQuery = "insert into TotalBiayaPelayananOA values('" & fNoStruk & "'," & msubKonversiKomaTitik(CStr(fTotalBiayaOA)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminOA)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSOA)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanOA)) & "," & msubKonversiKomaTitik(CStr(fJmlHrsDibayarOA)) & "," & msubKonversiKomaTitik(CStr(fJmlDiscountOA)) & "," & msubKonversiKomaTitik(CStr(fJmlBayarOA)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanOA)) & ")"
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "insert into PembayaranTagihanPasien values('" & fNoBKM & "','" & fNoStruk & "',0," & msubKonversiKomaTitik(CStr(fJmlBayarOA)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanOA)) & "+" & msubKonversiKomaTitik(CStr(fJmlDiscountOA)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanOA)) & ",1,'OA')"
        Call msubRecFO(fRS, fQuery)
    End If
    If UCase(fStatusPiutang) = "SA" Then
        Set fRS = Nothing
        fQuery = "insert into TotalBiayaPelayananOA values('" & fNoStruk & "'," & msubKonversiKomaTitik(CStr(fTotalBiayaOA)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminOA)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSOA)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanOA)) & "," & msubKonversiKomaTitik(CStr(fJmlHrsDibayarOA)) & "," & msubKonversiKomaTitik(CStr(fJmlDiscountOA)) & "," & msubKonversiKomaTitik(CStr(fJmlBayarOA)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanOA)) & ")"
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "insert into TotalBiayaPelayananTM values('" & fNoStruk & "'," & msubKonversiKomaTitik(CStr(fTotalBiayaTM)) & "," & msubKonversiKomaTitik(CStr(fJmlHutangPenjaminTM)) & "," & msubKonversiKomaTitik(CStr(fJmlTanggunganRSTM)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasanTM)) & "," & msubKonversiKomaTitik(CStr(fJmlHrsDibayarTM)) & "," & msubKonversiKomaTitik(CStr(fJmlDiscountTM)) & "," & msubKonversiKomaTitik(CStr(fJmlBayarTM)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihanTM)) & ")"
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "insert into PembayaranTagihanPasien values('" & fNoBKM & "','" & fNoStruk & "',0," & msubKonversiKomaTitik(CStr(fJmlBayar)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasan)) & "+" & msubKonversiKomaTitik(CStr(fJmlDiscount)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihan)) & ",1,'SA')"
        Call msubRecFO(fRS, fQuery)
    End If
    If UCase(fStatusBayarSemua) = "Y" Then
        Set fRS2 = Nothing
        fQuery2 = "SELECT NoPendaftaran FROM PasienSudahBayar WHERE NoPendaftaran = '" & fNoPendaftaran & "' AND NoStruk = '" & fNoStruk & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then
            Set fRS2 = Nothing
            fQuery2 = "INSERT INTO PasienSudahBayar VALUES ('" & fNoPendaftaran & "','" & fNoCM & "','" & fNoStruk & "')"
            Call msubRecFO(fRS2, fQuery2)
        End If
        Set fRS2 = Nothing
        fQuery2 = "SELECT NoPendaftaran FROM PasienBelumBayar WHERE NoPendaftaran = '" & fNoPendaftaran & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then
            Set fRS2 = Nothing
            fQuery2 = "DELETE FROM PasienBelumBayar WHERE NoPendaftaran = '" & fNoPendaftaran & "'"
            Call msubRecFO(fRS2, fQuery2)
        End If
    End If
    Set fRS2 = Nothing
    fQuery2 = "select TglPendaftaran,KdRuanganAkhir,TglPulang from PasienDaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS2, fQuery2)
    If fRS2.EOF = False Then fTglMasuk = IIf(IsNull(fRS2("TglPendaftaran").value), "", fRS2("TglPendaftaran").value) Else fTglMasuk = ""
    If fRS2.EOF = False Then fKdRuanganAkhir = IIf(IsNull(fRS2("KdRuanganAkhir").value), "", fRS2("KdRuanganAkhir").value) Else fKdRuanganAkhir = ""
    If fRS2.EOF = False Then fTglKeluar = IIf(IsNull(fRS2("TglPulang").value), "", fRS2("TglPulang").value) Else fTglKeluar = ""
    If Len(Trim(fKdRuanganAkhir)) <> "" And Len(Trim(fTglKeluar)) <> "" And Len(Trim(fTglMasuk)) <> "" Then
        Set fRS2 = Nothing
        fQuery2 = "insert into ConvertStrukPelayananToPasienPulang values('" & fNoStruk & "','" & fKdRuanganAkhir & "','" & Format(fTglMasuk, "yyyy/MM/dd HH:mm:ss") & "','" & Format(fTglKeluar, "yyyy/MM/dd HH:mm:ss") & "')"
        Call msubRecFO(fRS2, fQuery2)
    End If
End Function

'Konversi dari SP: Update_JenisPasien
Public Function f_UpdateJenisPasienDetail(fNoPendaftaran As String)
    'execute reports Jenis Pasien Lama
    Call f_UpdateReportsOAOnUbahJenisPasienLama(fNoPendaftaran)
    Call f_UpdateReportsTMOnUbahJenisPasienLama(fNoPendaftaran)
    'execute reports Jenis Pasien Baru
    Call f_UpdateBiayaPelayananOnUbahJenisPasien(fNoPendaftaran)
    Call f_AddDetailBiayaPelayananOnUbahJenisPasien(fNoPendaftaran)
    Call f_UpdatePemakaianAlkesOnUbahJenisPasien(fNoPendaftaran)
    Call f_AddDetailPemakaianObatAlkesOnUbahJenisPasien(fNoPendaftaran)
End Function

'Konversi dari SP: Update_PemakaianAlkesOnUbahJenisPasien
Public Function f_UpdatePemakaianAlkesOnUbahJenisPasien(fNoPendaftaran As String)
    Dim fKdRuangan As String
    Dim fKdBarang As String
    Dim fKdAsal As String
    Dim fSatuanJml As String
    Dim fStatusCito As String
    Dim fTglPelayanan As Date
    Dim fHargaSatuanBaru As Currency
    Dim fHargaBeliBaru As Currency
    Dim fIdPenjamin As String
    Dim fKdKelompokPasien As String
    Dim fJenisHargaNetto As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien from V_KelasTanggunganPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value)
        fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
    End If
    Set fRS = Nothing
    fQuery = "select NoPendaftaran,KdRuangan,KdBarang,KdAsal,TglPelayanan,SatuanJml from PemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdBarang = IIf(IsNull(fRS("KdBarang").value), "", fRS("KdBarang").value)
        fKdAsal = IIf(IsNull(fRS("KdAsal").value), "", fRS("KdAsal").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fSatuanJml = IIf(IsNull(fRS("SatuanJml").value), "", fRS("SatuanJml").value)
        Set fRS2 = Nothing
        fQuery2 = "select distinct JenisHargaNetto from PersentaseUpTarifOA where IdPenjamin='" & fIdPenjamin & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdAsal='" & fKdAsal & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then fJenisHargaNetto = IIf(IsNull(fRS2("JenisHargaNetto").value), "1", fRS2("JenisHargaNetto").value) Else fJenisHargaNetto = "1"
        If fJenisHargaNetto = "1" Then
            Set fRS2 = Nothing
            fQuery2 = "select HargaNetto1 from DetailHargaNettoBarang where  KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and Satuan='" & fSatuanJml & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = False Then fHargaBeliBaru = IIf(IsNull(fRS2("HargaNetto1").value), 0, fRS2("HargaNetto1").value) Else fHargaBeliBaru = 0
        Else
            Set fRS2 = Nothing
            fQuery2 = "select HargaNetto2 from DetailHargaNettoBarang where  KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and Satuan='" & fSatuanJml & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = False Then fHargaBeliBaru = IIf(IsNull(fRS2("HargaNetto2").value), 0, fRS2("HargaNetto2").value) Else fHargaBeliBaru = 0
        End If
        Set fRS2 = Nothing
        fQuery2 = "select dbo.FB_TakeTarifOA('" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdAsal & "'," & msubKonversiKomaTitik(CStr(fHargaBeliBaru)) & ") as HargaSatuanBaru"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fHargaSatuanBaru = IIf(IsNull(fRS2("HargaSatuanBaru").value), 0, fRS2("HargaSatuanBaru").value) Else fHargaSatuanBaru = 0
        Set fRS2 = Nothing
        fQuery2 = "update PemakaianAlkes set HargaSatuan=" & msubKonversiKomaTitik(CStr(fHargaSatuanBaru)) & ",HargaBeli=" & msubKonversiKomaTitik(CStr(fHargaBeliBaru)) & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and NoStruk is null"
        Call msubRecFO(fRS2, fQuery2)
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Update_ReportsOAOnUbahJenisPasienLama
Public Function f_UpdateReportsOAOnUbahJenisPasienLama(fNoPendaftaran As String)
    Dim fKdBarang As String
    Dim fKdAsal As String
    Dim fSatuanJml As String
    Dim fTglPelayanan As Date
    Dim fKdRuangan As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select NoPendaftaran,KdRuangan,KdBarang,KdAsal,TglPelayanan,SatuanJml from PemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdBarang = IIf(IsNull(fRS("KdBarang").value), "", fRS("KdBarang").value)
        fKdAsal = IIf(IsNull(fRS("KdAsal").value), "", fRS("KdAsal").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fSatuanJml = IIf(IsNull(fRS("SatuanJml").value), "", fRS("SatuanJml").value)
        Call f_DeleteTempHargaKomponenObatAlkes(fNoPendaftaran, fKdBarang, fTglPelayanan, fKdRuangan, fKdAsal, fSatuanJml)
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Update_ReportsTMOnUbahJenisPasienLama
Public Function f_UpdateReportsTMOnUbahJenisPasienLama(fNoPendaftaran As String)
    Dim fKdRuangan As String
    Dim fKdPelayananRS As String
    Dim fTglPelayanan As Date

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select NoPendaftaran,KdRuangan,KdPelayananRS,TglPelayanan from BiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        Call f_DeleteTempHargaKomponen(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan)
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Update_JenisPasienUmum
Public Function f_UpdateJenisPasienUmum(fKdKelompokPasien As String, fNoPendaftaran As String)
    Dim fMaksTglSJP As Date
    Dim fNoSJP As Variant

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select max(TglSJP) as MaksSJP from PemakaianAsuransi where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fMaksTglSJP = IIf(IsNull(fRS("MaksSJP").value), "", fRS("MaksSJP").value) Else fMaksTglSJP = ""
    Set fRS = Nothing
    fQuery = "select NoSJP from PemakaianAsuransi where NoPendaftaran= '" & fNoPendaftaran & "' and TglSJP='" & Format(fMaksTglSJP, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    fNoSJP = fRS("NoSJP").value
    'execute reports Jenis Pasien Lama
    Call f_UpdateReportsOAOnUbahJenisPasienLama(fNoPendaftaran)
    Call f_UpdateReportsTMOnUbahJenisPasienLama(fNoPendaftaran)
    Set fRS = Nothing
    fQuery = "update PasienDaftar set KdKelompokPasien = '" & fKdKelompokPasien & "' where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    Set fRS = Nothing
    fQuery = "delete from PemakaianAsuransi where NoPendaftaran= '" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    Set fRS = Nothing
    fQuery = "update BiayaPelayanan set KdJenisTarif = '01' where (NoPendaftaran='" & fNoPendaftaran & "' ) and (NoStruk is null)"
    Call msubRecFO(fRS, fQuery)
    'execute reports Jenis Pasien Baru
    Call f_UpdateBiayaPelayananOnUbahJenisPasien(fNoPendaftaran)
    Call f_AddDetailBiayaPelayananOnUbahJenisPasien(fNoPendaftaran)
    Call f_UpdatePemakaianAlkesOnUbahJenisPasien(fNoPendaftaran)
    Call f_AddDetailPemakaianObatAlkesOnUbahJenisPasien(fNoPendaftaran)
End Function

'Konversi dari SP: Update_KelasNBiayaPelayananPasien
Public Function f_UpdateKelasNBiayaPelayananPasien(fNoPendaftaran As String, fKdRuanganLogin As String, fKdKelasBaru As String, fNoPakai As String)
    Dim fKdPelayananRS As String
    Dim fStatusCito As String
    Dim fTglPelayanan As Date
    Dim fKdJenisTarif As String
    Dim fTarifBaru As Currency
    Dim fTarifCitoBaru As Currency
    Dim fKdInstalasi As String
    Dim fKdRuangan As String
    Dim fKdSubInstalasi As String
    Dim fJmlPelayanan As Integer
    Dim fNoLab_Rad As Variant
    Dim fIdPegawai As Variant
    Dim fStatusAPBD As String
    Dim fIdUser As String
    Dim fIdPegawai2 As Variant
    Dim fTglMasuk As Date
    Dim fIdPegawai3 As Variant

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select TglMasuk from PemakaianKamar where NoPakai='" & fNoPakai & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTglMasuk = IIf(IsNull(fRS("TglMasuk").value), "", fRS("TglMasuk").value) Else fTglMasuk = ""
    Set fRS = Nothing
    fQuery = "select NoPendaftaran,KdRuangan,KdPelayananRS,TglPelayanan,StatusCITO,KdJenisTarif,KdSubInstalasi,JmlPelayanan,NoLab_Rad,IdPegawai,StatusAPBD,IdUser,IdPegawai2,IdPegawai3 from BiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        fStatusCito = IIf(IsNull(fRS("StatusCITO").value), "", fRS("StatusCITO").value)
        fKdJenisTarif = IIf(IsNull(fRS("KdJenisTarif").value), "", fRS("KdJenisTarif").value)
        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
        fJmlPelayanan = IIf(IsNull(fRS("JmlPelayanan").value), 0, fRS("JmlPelayanan").value)
        fNoLab_Rad = fRS("NoLab_Rad").value
        fIdPegawai = fRS("IdPegawai").value
        fStatusAPBD = IIf(IsNull(fRS("StatusAPBD").value), "01", fRS("StatusAPBD").value)
        fIdUser = IIf(IsNull(fRS("IdUser").value), "", fRS("IdUser").value)
        fIdPegawai2 = fRS("NoPendaftaran").value
        fIdPegawai3 = fRS("NoPendaftaran").value
        Set fRS2 = Nothing
        fQuery2 = "select dbo.FB_NewTakeTarifBPTM ('" & fNoPendaftaran & "','" & fKdPelayananRS & "','" & fKdKelasBaru & "', '" & fKdJenisTarif & "', '" & fStatusCito & "', " & fIdPegawai & "," & fIdPegawai2 & ", " & fIdPegawai3 & ", 'T') as TarifBaru"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fTarifBaru = IIf(IsNull(fRS2("TarifBaru").value), 0, fRS2("TarifBaru").value) Else fTarifBaru = 0
        Set fRS2 = Nothing
        fQuery2 = "select dbo.FB_NewTakeTarifBPTM ('" & fNoPendaftaran & "','" & fKdPelayananRS & "','" & fKdKelasBaru & "', '" & fKdJenisTarif & "', '" & fStatusCito & "', " & fIdPegawai & "," & fIdPegawai2 & ", " & fIdPegawai3 & ", 'C') as TarifCitoBaru"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fTarifCitoBaru = IIf(IsNull(fRS2("TarifCitoBaru").value), 0, fRS2("TarifCitoBaru").value) Else fTarifCitoBaru = 0
        Set fRS2 = Nothing
        fQuery2 = "select KdInstalasi from Ruangan where KdRuangan=fKdRuangan"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fKdInstalasi = IIf(IsNull(fRS2("KdInstalasi").value), "", fRS2("KdInstalasi").value) Else fKdInstalasi = ""
        If (((fKdInstalasi <> "03" And fKdInstalasi <> "08") And fKdRuangan <> fKdRuanganLogin) Or (fKdRuangan = fKdRuanganLogin)) And fTglPelayanan >= fTglMasuk Then
            Set fRS2 = Nothing
            fQuery2 = "update BiayaPelayanan set KdKelas='" & fKdKelasBaru & "',TarifCito=" & fTarifCitoBaru & ",Tarif=" & fTarifBaru & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update PemakaianAlkes set KdKelas='" & fKdKelasBaru & "' where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update DetailPemakaianAlkes set KdKelas='" & fKdKelasBaru & "' where NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "delete from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
            Call f_AddDetailBiayaPelayanan(fNoPendaftaran, fKdSubInstalasi, fKdRuangan, fKdPelayananRS, fKdKelasBaru, fStatusCito, CDbl(fTarifBaru), fJmlPelayanan, fTglPelayanan, fNoLab_Rad, fIdPegawai, fStatusAPBD, fKdJenisTarif, CInt(fTarifCitoBaru), fIdUser, fIdPegawai2, fIdPegawai3)
            Call f_UpdateBiayaPelayananFromBackupBiayaPelayanan(fNoPendaftaran, fKdRuangan, fKdPelayananRS, fTglPelayanan, CStr(fIdPegawai))
        End If
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Update_BiayaPelayananFromBackupBiayaPelayanan
Public Function f_UpdateBiayaPelayananFromBackupBiayaPelayanan(fNoPendaftaran As String, fKdRuangan As String, fKdPelayananRS As String, fTglPelayanan As Date, fIdDokter As String)
    Dim fTempHarga As Currency
    Dim fTotalHarga As Currency
    Dim fTempTarifCito As Currency
    Dim fTotalTarifCito As Currency
    Dim fTempTarifBP As Currency
    Dim fTotalTarifBP As Currency
    Dim fKdKelas As String
    Dim fKdJenisTarif As String
    Dim fKdKomponen As String
    Dim fJmlDiscount As Currency
    Dim fJmlCharge As Currency
    Dim fTarif As Currency
    Dim fKdRuanganAsal As String
    Dim fNoLab_Rad As Variant
    Dim fKdInstalasi As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select NoPendaftaran from DetailBackupUpdatingBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        Set fRS = Nothing
        fQuery = "select TarifCito,Tarif,KdKelas,KdJenisTarif,NoLab_Rad from BiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdPegawai='" & fIdDokter & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fTempTarifCito = IIf(IsNull(fRS("TarifCito").value), 0, fRS("TarifCito").value) Else fTempTarifCito = 0
        If fRS.EOF = False Then fTempTarifBP = IIf(IsNull(fRS("Tarif").value), 0, fRS("Tarif").value) Else fTempTarifBP = 0
        If fRS.EOF = False Then fKdJenisTarif = IIf(IsNull(fRS("KdJenisTarif").value), "01", fRS("KdJenisTarif").value) Else fKdJenisTarif = "01"
        If fRS.EOF = False Then fKdKelas = IIf(IsNull(fRS("KdKelas").value), "01", fRS("KdKelas").value) Else fKdKelas = "01"
        fNoLab_Rad = fRS("NoLab_Rad").value
        Set fRS = Nothing
        fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "'," & fNoLab_Rad & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
        Call msubRecFO(fRS, fQuery)
        fKdRuanganAsal = fRS("KdRuanganAsal").value
        Set fRS = Nothing
        fQuery = "select KdKomponen,JmlDiscount,JmlCharge from DetailBackupUpdatingBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
        Call msubRecFO(fRS, fQuery)
        While fRS.EOF = False
            fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
            fJmlDiscount = IIf(IsNull(fRS("JmlDiscount").value), 0, fRS("JmlDiscount").value)
            fJmlCharge = IIf(IsNull(fRS("JmlCharge").value), 0, fRS("JmlCharge").value)
            Set fRS2 = Nothing
            fQuery2 = "select NoPendaftaran from BiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdPegawai='" & fIdDokter & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = False Then
                Set fRS2 = Nothing
                fQuery2 = "select KdKomponen from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
                Call msubRecFO(fRS2, fQuery2)
                If fRS2.EOF = False Then
                    Set fRS2 = Nothing
                    fQuery2 = "select Harga from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdKomponen='" & fKdKomponen & "'"
                    Call msubRecFO(fRS2, fQuery2)
                    If fRS2.EOF = False Then fTempHarga = IIf(IsNull(fRS2("Harga").value), 0, fRS2("Harga").value) Else fTempHarga = 0
                    If fKdKomponen = "07" Then
                        If fJmlCharge = 0 Then
                            fTotalTarifCito = fTempTarifCito - fJmlDiscount
                            fTotalHarga = fTempHarga - fJmlDiscount
                        Else
                            fTotalTarifCito = fTempTarifCito + fJmlCharge
                            fTotalHarga = fTempHarga + fJmlCharge
                        End If
                        Set fRS3 = Nothing
                        fQuery3 = "update BiayaPelayanan set TarifCito=" & fTotalTarifCito & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdPegawai='" & fIdDokter & "' and NoStruk is null"
                        Call msubRecFO(fRS3, fQuery3)
                        Set fRS3 = Nothing
                        fQuery3 = "update DetailBiayaPelayanan set TarifCito=" & fTotalTarifCito & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdPegawai='" & fIdDokter & "' and NoStruk is null"
                        Call msubRecFO(fRS3, fQuery3)
                        Set fRS3 = Nothing
                        fQuery3 = "update TempHargaKomponen set Harga=" & fTotalHarga & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
                        Call msubRecFO(fRS3, fQuery3)
                    Else
                        If fJmlCharge = 0 Then
                            fTotalTarifBP = fTempTarifBP - fJmlDiscount
                            fTotalHarga = fTempHarga - fJmlDiscount
                        Else
                            fTotalTarifBP = fTempTarifBP + fJmlCharge
                            fTotalHarga = fTempHarga + fJmlCharge
                        End If
                        Set fRS3 = Nothing
                        fQuery3 = "update BiayaPelayanan set TarifCito=" & fTotalTarifBP & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdPegawai='" & fIdDokter & "' and NoStruk is null"
                        Call msubRecFO(fRS3, fQuery3)
                        Set fRS3 = Nothing
                        fQuery3 = "update DetailBiayaPelayanan set TarifCito=" & fTotalTarifBP & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdPegawai='" & fIdDokter & "' and NoStruk is null"
                        Call msubRecFO(fRS3, fQuery3)
                        Set fRS3 = Nothing
                        fQuery3 = "update TempHargaKomponen set Harga=" & fTotalHarga & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
                        Call msubRecFO(fRS3, fQuery3)
                    End If
                    Set fRS3 = Nothing
                    fQuery3 = "select NoClosing from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdKomponen='" & fKdKomponen & "' and NoClosing is not null "
                    Call msubRecFO(fRS3, fQuery3)
                    If fRS3.EOF = False Then
                        Call f_AMDataPelayananTMPasienPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fTempHarga, 0, 0, 0, fKdKelas, "M")
                        Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fTempHarga, 0, 0, 0, fKdKelas, fIdDokter, "M")
                        Call f_AMDataPelayananTMPasienPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fTotalHarga, 0, 0, 0, fKdKelas, "A")
                        Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fTotalHarga, 0, 0, 0, fKdKelas, fIdDokter, "A")
                    End If
                Else
                    Set fRS2 = Nothing
                    fQuery2 = "select Harga from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen='" & fKdKomponen & "'"
                    Call msubRecFO(fRS2, fQuery2)
                    If fRS2.EOF = False Then fTarif = IIf(IsNull(fRS2("Harga").value), 0, fRS2("Harga").value) Else fTarif = 0
                    If fKdKomponen = "07" Then
                        fTotalTarifCito = fTempTarifCito + fTarif
                        Set fRS3 = Nothing
                        fQuery3 = "update BiayaPelayanan set TarifCito=" & fTotalTarifCito & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdPegawai='" & fIdDokter & "' and NoStruk is null"
                        Call msubRecFO(fRS3, fQuery3)
                        Set fRS3 = Nothing
                        fQuery3 = "update DetailBiayaPelayanan set TarifCito=" & fTotalTarifCito & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdPegawai='" & fIdDokter & "' and NoStruk is null"
                        Call msubRecFO(fRS3, fQuery3)
                    Else
                        fTotalTarifBP = fTempTarifBP + fTarif
                        Set fRS3 = Nothing
                        fQuery3 = "update BiayaPelayanan set Tarif=" & fTotalTarifBP & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdPegawai='" & fIdDokter & "' and NoStruk is null"
                        Call msubRecFO(fRS3, fQuery3)
                        Set fRS3 = Nothing
                        fQuery3 = "update DetailBiayaPelayanan set Tarif=" & fTotalTarifBP & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdPegawai='" & fIdDokter & "' and NoStruk is null"
                        Call msubRecFO(fRS3, fQuery3)
                    End If
                    Set fRS3 = Nothing
                    fQuery3 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "'," & fTarif & ",1,null,'" & fIdDokter & "',0,0,0,null)"
                    Call msubRecFO(fRS3, fQuery3)
                    'aktifkan ini jika ingin di rekap otomatis
                    'Call f_AMDataPelayananTMPasienPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fTarif, 0, 0, 0, fKdKelas, "A")
                    'Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fTarif, 0, 0, 0, fKdKelas, fIdDokter, "A")
                End If
            End If
            fRS.MoveNext
        Wend
    End If
End Function

'Konversi dari SP: AM_DataKunjunganPelayananTMPasienPH
Public Function f_AMDataKunjunganPelayananTMPasienPH(fNoPendaftaran As String, fKdRuangan As String, fKdRuanganAsal As String, fTglPelayanan As Date, fKdPelayananRS As String, fIdPenjamin As String, fKdKelompokPasien As String, fJmlPelayanan As Integer, fNoLab_Rad As Variant, fStatus As String)
    'fStatus: A=Add, M=Min
    Dim fKdJenisKelamin As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fStatusPasien As String
    Dim fKdRujukanAsal As String
    Dim fKdSubInstalasi As String
    Dim fKdKelas As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "'," & fNoLab_Rad & ",'" & fKdRuangan & "','1') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fStatusPasien = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "'," & fNoLab_Rad & ",'" & fKdRuangan & "','2') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdRujukanAsal = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "'," & fNoLab_Rad & ",'" & fKdRuangan & "','3') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdSubInstalasi = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "'," & fNoLab_Rad & ",'" & fKdRuangan & "','4') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdKelas = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select KdJenisKelamin from V_JenisKelaminPasienTerdaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisKelamin = IIf(IsNull(fRS("KdJenisKelamin").value), "", fRS("KdJenisKelamin").value) Else fKdJenisKelamin = ""
    Set fRS = Nothing
    fQuery = "select KdDetailJenisJasaPelayanan from V_JenisPasienNPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdDetailJenisJasaPelayanan = IIf(IsNull(fRS("KdDetailJenisJasaPelayanan").value), "", fRS("KdDetailJenisJasaPelayanan").value) Else fKdDetailJenisJasaPelayanan = ""
    Set fRS = Nothing
    fQuery = "select KdRuangan from DataKunjunganPelayananTMPasienPH where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and KdPelayananRS='" & fKdPelayananRS & "') and (day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        Set fRS = Nothing
        fQuery = "insert into DataKunjunganPelayananTMPasienPH values('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdDetailJenisJasaPelayanan & "','" & fKdRujukanAsal & "','" & fKdKelas & "','" & fStatusPasien & "','" & fKdPelayananRS & "','" & fKdJenisKelamin & "'," & fJmlPelayanan & ")"
        Call msubRecFO(fRS, fQuery)
    Else
        If UCase(fStatus) = "A" Then
            fQuery = "update DataKunjunganPelayananTMPasienPH set JmlPasien=JmlPasien + " & fJmlPelayanan & " where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and KdPelayananRS='" & fKdPelayananRS & "') and (day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery = "update DataKunjunganPelayananTMPasienPH set JmlPasien=JmlPasien - " & fJmlPelayanan & " where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and KdPelayananRS='" & fKdPelayananRS & "') and (day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
        Set fRS = Nothing
        Call msubRecFO(fRS, fQuery)
    End If
End Function

'Konversi dari SP: Add_RekapitulasiKamarRawatInapMasuk
Public Function f_AddRekapitulasiKamarRawatInapMasuk(fTglHitung As Date)
    'fTglHitung= TglMasuk
    Dim fKdRuangan As String
    Dim fKdKelas As String
    Dim fJmlBedTerisiTemp As Integer
    Dim fJmlBedTerisi As Integer
    Dim fJmlBedKosongTemp As Integer
    Dim fJmlBedKosong As Integer
    Dim fKdRuanganTemp As String
    Dim fKdKamar As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdRuangan,KdKelas,KdKamar from V_InformasiKamarRawatInap"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
        fKdKamar = IIf(IsNull(fRS("KdKamar").value), "", fRS("KdKamar").value)
        Set fRS2 = Nothing
        fQuery2 = "select count(NoBed) as JmlBedIsi from V_InformasiKamarRawatInap where KdKamar='" & fKdKamar & "' and KdKelas='" & fKdKelas & "' and KdRuangan='" & fKdRuangan & "' and Status='Isi'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fJmlBedTerisiTemp = IIf(IsNull(fRS2("JmlBedIsi").value), 0, fRS2("JmlBedIsi").value) Else fJmlBedTerisiTemp = 0
        Set fRS2 = Nothing
        fQuery2 = "select count(NoBed) as JmlBedKosong from V_InformasiKamarRawatInap where KdKamar='" & fKdKamar & "' and KdKelas='" & fKdKelas & "' and KdRuangan='" & fKdRuangan & "' and Status='Kosong'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fJmlBedKosongTemp = IIf(IsNull(fRS2("JmlBedKosong").value), 0, fRS2("JmlBedKosong").value) Else fJmlBedKosongTemp = 0
        fJmlBedTerisi = fJmlBedTerisiTemp
        fJmlBedKosong = fJmlBedKosongTemp
        Set fRS2 = Nothing
        fQuery2 = "select KdRuangan from RekapitulasiKamarRawatInap where (KdRuangan='" & fKdRuangan & "' and KdKamar='" & fKdKamar & "' and KdKelas='" & fKdKelas & "') and (day(TglHitung)=day('" & Format(fTglHitung, "yyyy/MM/dd HH:mm:ss") & "') and month(TglHitung)=month('" & Format(fTglHitung, "yyyy/MM/dd HH:mm:ss") & "') and year(TglHitung)=year('" & Format(fTglHitung, "yyyy/MM/dd HH:mm:ss") & "'))"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then
            fQuery3 = "insert into RekapitulasiKamarRawatInap values('" & Format(fTglHitung, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuangan & "','" & fKdKelas & "','" & fKdKamar & "'," & fJmlBedTerisi & "," & fJmlBedKosong & ")"
        Else
            fQuery3 = "update RekapitulasiKamarRawatInap set JmlBedTerisi=" & fJmlBedTerisi & ",JmlBedKosong=" & fJmlBedKosong & " where (KdRuangan='" & fKdRuangan & "' and KdKamar='" & fKdKamar & "' and KdKelas=fKdKelas) and (day(TglHitung)=day('" & Format(fTglHitung, "yyyy/MM/dd HH:mm:ss") & "') and month(TglHitung)=month('" & Format(fTglHitung, "yyyy/MM/dd HH:mm:ss") & "') and year(TglHitung)=year('" & Format(fTglHitung, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
        Set fRS3 = Nothing
        Call msubRecFO(fRS3, fQuery3)
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: AM_DataKunjunganPasienMasukPH
Public Function f_AMDataKunjunganPasienMasukPH(fNoPendaftaran As String, fNoLab_Rad_IBS As Variant, fNoCM As String, fKdRuangan As String, fKdRuanganAsal As String, fKdKelompokPasien As String, fTglPendaftaran As Date, fStatus As String)
    'fStatus: A=Add, M=Min
    Dim fKdJenisKelamin As String
    Dim fKecamatan As String
    Dim fStatusPasien As String
    Dim fKdRujukanAsal As String
    Dim fKdSubInstalasi As String
    Dim fKdKelas As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "'," & fNoLab_Rad_IBS & ",'" & fKdRuangan & "','1') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fStatusPasien = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "'," & fNoLab_Rad_IBS & ",'" & fKdRuangan & "','2') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdRujukanAsal = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "'," & fNoLab_Rad_IBS & ",'" & fKdRuangan & "','3') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdSubInstalasi = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "'," & fNoLab_Rad_IBS & ",'" & fKdRuangan & "','4') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdKelas = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select distinct KdJenisKelamin,Kecamatan from V_JenisKelaminPasienTerdaftar where NoCM='" & fNoCM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisKelamin = IIf(IsNull(fRS("KdJenisKelamin").value), "01", fRS("KdJenisKelamin").value) Else fKdJenisKelamin = "01"
    If fRS.EOF = False Then fKecamatan = IIf(IsNull(fRS("Kecamatan").value), "Lain - Lain", fRS("Kecamatan").value) Else fKecamatan = "Lain - Lain"
    Set fRS = Nothing
    fQuery = "select KdRuangan from DataKunjunganPasienMasukPH where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and Kecamatan='" & fKecamatan & "') and (day(TglPendaftaran)=day('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPendaftaran)=month('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPendaftaran)=year('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery2 = "insert into DataKunjunganPasienMasukPH values('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdKelompokPasien & "','" & fKdRujukanAsal & "','" & fKdKelas & "','" & fStatusPasien & "','" & fKecamatan & "','" & fKdJenisKelamin & "',1)"
    Else
        If UCase(fStatus) = "A" Then
            fQuery2 = "update DataKunjunganPasienMasukPH set JmlPasien=JmlPasien + 1 where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and Kecamatan='" & fKecamatan & "') and (day(TglPendaftaran)=day('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPendaftaran)=month('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPendaftaran)=year('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery2 = "update DataKunjunganPasienMasukPH set JmlPasien=JmlPasien - 1 where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and Kecamatan='" & fKecamatan & "') and (day(TglPendaftaran)=day('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPendaftaran)=month('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPendaftaran)=year('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery)
End Function

'Konversi dari SP: AM_DataDiagnosaPasienPH
Public Function f_AMDataDiagnosaPasienPH(fNoCM As String, fKdRuangan As String, fKdKelompokPasien As String, fTglPeriksa As Date, fKdJenisDiagnosa As String, fKdDiagnosa As String, fStatusKasus As String, fStatus As String)
    'fStatus: A=Add, M=Min
    Dim fKdJenisKelamin As String
    Dim fKecamatan As String
    Dim fStatusPasien As String
    Dim fKdRujukanAsal As String
    Dim fKdSubInstalasi As String
    Dim fKdKelas As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "',null,'" & fKdRuangan & "','1') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fStatusPasien = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "',null,'" & fKdRuangan & "','2') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdRujukanAsal = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "',null,'" & fKdRuangan & "','3') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdSubInstalasi = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "',null,'" & fKdRuangan & "','4') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdKelas = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select distinct KdJenisKelamin,Kecamatan from V_JenisKelaminPasienTerdaftar where NoCM='" & fNoCM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisKelamin = IIf(IsNull(fRS("KdJenisKelamin").value), "01", fRS("KdJenisKelamin").value) Else fKdJenisKelamin = "01"
    If fRS.EOF = False Then fKecamatan = IIf(IsNull(fRS("Kecamatan").value), "Lain - Lain", fRS("Kecamatan").value) Else fKecamatan = "Lain - Lain"
    Set fRS = Nothing
    fQuery = "select KdRuangan from DataDiagnosaPasienPH where (KdRuangan='" & fKdRuangan & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdJenisDiagnosa='" & fKdJenisDiagnosa & "' and KdDiagnosa='" & fKdDiagnosa & "' and StatusKasus='" & fStatusKasus & "' and Kecamatan='" & fKecamatan & "' and KdJenisKelamin='" & fKdJenisKelamin & "') and (day(TglPeriksa)=day('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPeriksa)=month('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPeriksa)=year('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery2 = "insert into DataDiagnosaPasienPH values('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdKelompokPasien & "','" & fKdRujukanAsal & "','" & fKdKelas & "','" & fStatusPasien & "','" & fKdJenisDiagnosa & "','" & fKdDiagnosa & "','" & fStatusKasus & "','" & fKecamatan & "','" & fKdJenisKelamin & "',1)"
    Else
        If UCase(fStatus) = "A" Then
            fQuery2 = "update DataDiagnosaPasienPH set JmlPasien=JmlPasien + 1 where (KdRuangan='" & fKdRuangan & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdJenisDiagnosa='" & fKdJenisDiagnosa & "' and KdDiagnosa='" & fKdDiagnosa & "' and StatusKasus='" & fStatusKasus & "' and Kecamatan='" & fKecamatan & "' and KdJenisKelamin='" & fKdJenisKelamin & "') and (day(TglPeriksa)=day('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPeriksa)=month('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPeriksa)=year('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery2 = "update DataDiagnosaPasienPH set JmlPasien=JmlPasien - 1 where (KdRuangan='" & fKdRuangan & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdJenisDiagnosa='" & fKdJenisDiagnosa & "' and KdDiagnosa='" & fKdDiagnosa & "' and StatusKasus='" & fStatusKasus & "' and Kecamatan='" & fKecamatan & "' and KdJenisKelamin='" & fKdJenisKelamin & "') and (day(TglPeriksa)=day('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPeriksa)=month('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPeriksa)=year('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
End Function

'Konversi dari SP: AM_DataKunjunganPasienKeluarIBSPH
Public Function f_AMDataKunjunganPasienKeluarIBSPH(fNoCM As String, fKdRuangan As String, fKdRuanganAsal As String, fNoIBS As String, fKdKelompokPasien As String, fTglOperasi As Date, fKdTindakanOperasi As String, fStatus As String)
    'fStatus: A=Add, M=Min
    Dim fKdJenisKelamin As String
    Dim fKecamatan As String
    Dim fStatusPasien As String
    Dim fKdRujukanAsal As String
    Dim fKdSubInstalasi As String
    Dim fKdKelas As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "','" & fNoIBS & "','" & fKdRuangan & "','1') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fStatusPasien = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "','" & fNoIBS & "','" & fKdRuangan & "','2') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdRujukanAsal = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "','" & fNoIBS & "','" & fKdRuangan & "','3') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdSubInstalasi = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "','" & fNoIBS & "','" & fKdRuangan & "','4') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdKelas = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select distinct KdJenisKelamin,Kecamatan from V_JenisKelaminPasienTerdaftar where NoCM='" & fNoCM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisKelamin = IIf(IsNull(fRS("KdJenisKelamin").value), "01", fRS("KdJenisKelamin").value) Else fKdJenisKelamin = "01"
    If fRS.EOF = False Then fKecamatan = IIf(IsNull(fRS("Kecamatan").value), "Lain - Lain", fRS("Kecamatan").value) Else fKecamatan = "Lain - Lain"
    Set fRS = Nothing
    fQuery = "select KdRuangan from DataKunjunganPasienKeluarIBSPH where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and KdTindakanOperasi='" & fKdTindakanOperasi & "' and Kecamatan='" & fKecamatan & "') and (day(TglOperasi)=day('" & Format(fTglOperasi, "yyyy/MM/dd HH:mm:ss") & "') and month(TglOperasi)=month('" & Format(fTglOperasi, "yyyy/MM/dd HH:mm:ss") & "') and year(TglOperasi)=year('" & Format(fTglOperasi, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery2 = "insert into DataKunjunganPasienKeluarIBSPH values('" & Format(fTglOperasi, "yyyy/MM/dd HH:mm:ss") & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdKelompokPasien & "','" & fKdRujukanAsal & "','" & fKdKelas & "','" & fStatusPasien & "','" & fKdTindakanOperasi & "','" & fKecamatan & "','" & fKdJenisKelamin & "',1)"
    Else
        If UCase(fStatus) = "A" Then
            fQuery2 = "update DataKunjunganPasienKeluarIBSPH set JmlPasien=JmlPasien + 1 where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and KdTindakanOperasi='" & fKdTindakanOperasi & "' and Kecamatan='" & fKecamatan & "') and (day(TglOperasi)=day('" & Format(fTglOperasi, "yyyy/MM/dd HH:mm:ss") & "') and month(TglOperasi)=month('" & Format(fTglOperasi, "yyyy/MM/dd HH:mm:ss") & "') and year(TglOperasi)=year('" & Format(fTglOperasi, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery2 = "update DataKunjunganPasienKeluarIBSPH set JmlPasien=JmlPasien - 1 where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and KdTindakanOperasi='" & fKdTindakanOperasi & "' and Kecamatan='" & fKecamatan & "') and (day(TglOperasi)=day('" & Format(fTglOperasi, "yyyy/MM/dd HH:mm:ss") & "') and month(TglOperasi)=month('" & Format(fTglOperasi, "yyyy/MM/dd HH:mm:ss") & "') and year(TglOperasi)=year('" & Format(fTglOperasi, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
End Function

'Konversi dari SP: AM_DataKunjunganPasienKeluarPH
Public Function f_AMDataKunjunganPasienKeluarPH(fNoCM As String, fKdRuangan As String, fKdKelompokPasien As String, fKdStatusKeluar As String, fKdKondisiPulang As String, fTglKeluar As Date, fStatus As String)
    'fStatus: A=add, M=Min
    Dim fKdJenisKelamin As String
    Dim fKecamatan As String
    Dim fStatusPasien As String
    Dim fKdRujukanAsal As String
    Dim fKdSubInstalasi As String
    Dim fKdKelas As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "',null,'" & fKdRuangan & "','1') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fStatusPasien = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "',null,'" & fKdRuangan & "','2') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdRujukanAsal = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "',null,'" & fKdRuangan & "','3') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdSubInstalasi = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "',null,'" & fKdRuangan & "','4') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdKelas = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select distinct KdJenisKelamin,Kecamatan from V_JenisKelaminPasienTerdaftar where NoCM='" & fNoCM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisKelamin = IIf(IsNull(fRS("KdJenisKelamin").value), "01", fRS("KdJenisKelamin").value) Else fKdJenisKelamin = "01"
    If fRS.EOF = False Then fKecamatan = IIf(IsNull(fRS("Kecamatan").value), "Lain - Lain", fRS("Kecamatan").value) Else fKecamatan = "Lain - Lain"
    Set fRS = Nothing
    fQuery = "select KdRuangan from DataKunjunganPasienKeluarPH where (KdRuangan='" & fKdRuangan & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and StatusPasien='" & fStatusPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and KdStatusKeluar='" & fKdStatusKeluar & "' and KdKondisiPulang='" & fKdKondisiPulang & "' and KdKelas='" & fKdKelas & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and Kecamatan='" & fKecamatan & "') and (day(TglKeluar)=day('" & Format(fTglKeluar, "yyyy/MM/dd HH:mm:ss") & "') and month(TglKeluar)=month('" & Format(fTglKeluar, "yyyy/MM/dd HH:mm:ss") & "') and year(TglKeluar)=year('" & Format(fTglKeluar, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If rs.EOF = True Then
        fQuery2 = "insert into DataKunjunganPasienKeluarPH values('" & Format(fTglKeluar, "yyyy/MM/dd HH:mm:ss") & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdKelompokPasien & "','" & fKdRujukanAsal & "','" & fKdKelas & "','" & fStatusPasien & "','" & fKdStatusKeluar & "','" & fKdKondisiPulang & "','" & fKecamatan & "','" & fKdJenisKelamin & "',1)"
    Else
        If UCase(fStatus) = "A" Then
            fQuery2 = "update DataKunjunganPasienKeluarPH set JmlPasien=JmlPasien + 1 where (KdRuangan='" & fKdRuangan & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and StatusPasien='" & fStatusPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and KdStatusKeluar='" & fKdStatusKeluar & "' and KdKondisiPulang='" & fKdKondisiPulang & "' and KdKelas='" & fKdKelas & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and Kecamatan='" & fKecamatan & "') and (day(TglKeluar)=day('" & Format(fTglKeluar, "yyyy/MM/dd HH:mm:ss") & "') and month(TglKeluar)=month('" & Format(fTglKeluar, "yyyy/MM/dd HH:mm:ss") & "') and year(TglKeluar)=year('" & Format(fTglKeluar, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery2 = "update DataKunjunganPasienKeluarPH set JmlPasien=JmlPasien - 1 where (KdRuangan='" & fKdRuangan & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and StatusPasien='" & fStatusPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and KdStatusKeluar='" & fKdStatusKeluar & "' and KdKondisiPulang='" & fKdKondisiPulang & "' and KdKelas='" & fKdKelas & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and Kecamatan='" & fKecamatan & "') and (day(TglKeluar)=day('" & Format(fTglKeluar, "yyyy/MM/dd HH:mm:ss") & "') and month(TglKeluar)=month('" & Format(fTglKeluar, "yyyy/MM/dd HH:mm:ss") & "') and year(TglKeluar)=year('" & Format(fTglKeluar, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
End Function

'Konversi dari SP: AM_DataKunjunganPasienMasukIBSPH
Public Function f_AMDataKunjunganPasienMasukIBSPH(fNoCM As String, fKdRuangan As String, fKdRuanganAsal As String, fNoIBS As String, fKdKelompokPasien As String, fTglPendaftaran As Date, fKdJenisOperasi As String, fStatus As String)
    'fStatus: A=Add, M=Min
    Dim fKdJenisKelamin As String
    Dim fKecamatan As String
    Dim fStatusPasien As String
    Dim fKdRujukanAsal As String
    Dim fKdSubInstalasi As String
    Dim fKdKelas As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "',null,'" & fKdRuangan & "','1') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fStatusPasien = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "',null,'" & fKdRuangan & "','2') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdRujukanAsal = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "',null,'" & fKdRuangan & "','3') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdSubInstalasi = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "',null,'" & fKdRuangan & "','4') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdKelas = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select distinct KdJenisKelamin,Kecamatan from V_JenisKelaminPasienTerdaftar where NoCM='" & fNoCM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisKelamin = IIf(IsNull(fRS("KdJenisKelamin").value), "01", fRS("KdJenisKelamin").value) Else fKdJenisKelamin = "01"
    If fRS.EOF = False Then fKecamatan = IIf(IsNull(fRS("Kecamatan").value), "Lain - Lain", fRS("Kecamatan").value) Else fKecamatan = "Lain - Lain"
    Set fRS = Nothing
    fQuery = "select KdRuangan from DataKunjunganPasienMasukIBSPH where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and KdJenisOperasi='" & fKdJenisOperasi & "' and Kecamatan='" & fKecamatan & "') and (day(TglPendaftaran)=day('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPendaftaran)=month('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPendaftaran)=year('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery2 = "insert into DataKunjunganPasienMasukIBSPH values(fTglPendaftaran,'" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdKelompokPasien & "','" & fKdRujukanAsal & "','" & fKdKelas & "','" & fStatusPasien & "','" & fKdJenisOperasi & "','" & fKecamatan & "','" & fKdJenisKelamin & "',1)"
    Else
        If UCase(fStatus) = "A" Then
            fQuery2 = "update DataKunjunganPasienMasukIBSPH set JmlPasien=JmlPasien + 1 where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and KdJenisOperasi='" & fKdJenisOperasi & "' and Kecamatan='" & fKecamatan & "') and (day(TglPendaftaran)=day('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPendaftaran)=month('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPendaftaran)=year('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery2 = "update DataKunjunganPasienMasukIBSPH set JmlPasien=JmlPasien - 1 where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and KdJenisOperasi='" & fKdJenisOperasi & "' and Kecamatan='" & fKecamatan & "') and (day(TglPendaftaran)=day('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPendaftaran)=month('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPendaftaran)=year('" & Format(fTglPendaftaran, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
End Function

'Konversi dari SP: AM_DataKunjunganPasienMasukTriasePH
Public Function f_AMDataKunjunganPasienMasukTriasePH(fNoCM As String, fKdRuangan As String, fKdRuanganAsal As String, fKdKelompokPasien As String, fTglPeriksa As Date, fKdTriase As String, fStatus As String)
    'fStatus: A=Add, M=Min
    Dim fKdJenisKelamin As String
    Dim fKecamatan As String
    Dim fStatusPasien As String
    Dim fKdRujukanAsal As String
    Dim fKdSubInstalasi As String
    Dim fKdKelas As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "',null,'" & fKdRuangan & "','1') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fStatusPasien = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "',null,'" & fKdRuangan & "','2') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdRujukanAsal = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "',null,'" & fKdRuangan & "','3') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdSubInstalasi = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "',null,'" & fKdRuangan & "','4') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdKelas = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select distinct KdJenisKelamin,Kecamatan from V_JenisKelaminPasienTerdaftar where NoCM='" & fNoCM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisKelamin = IIf(IsNull(fRS("KdJenisKelamin").value), "01", fRS("KdJenisKelamin").value) Else fKdJenisKelamin = "01"
    If fRS.EOF = False Then fKecamatan = IIf(IsNull(fRS("Kecamatan").value), "Lain - Lain", fRS("Kecamatan").value) Else fKecamatan = "Lain - Lain"
    Set fRS = Nothing
    fQuery = "select KdRuangan from DataKunjunganPasienMasukTriasePH where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and KdTriase='" & fKdTriase & "' and Kecamatan='" & fKecamatan & "') and (day(TglPeriksa)=day('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPeriksa)=month('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPeriksa)=year('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery2 = "insert into DataKunjunganPasienMasukTriasePH values('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdKelompokPasien & "','" & fKdRujukanAsal & "','" & fKdKelas & "','" & fStatusPasien & "','" & fKdTriase & "','" & fKecamatan & "','" & fKdJenisKelamin & "',1)"
    Else
        If UCase(fStatus) = "A" Then
            fQuery2 = "update DataKunjunganPasienMasukTriasePH set JmlPasien=JmlPasien + 1 where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and KdTriase='" & fKdTriase & "' and Kecamatan='" & fKecamatan & "') and (day(TglPeriksa)=day('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPeriksa)=month('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPeriksa)=year('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery2 = "update DataKunjunganPasienMasukTriasePH set JmlPasien=JmlPasien - 1 where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and KdRujukanAsal='" & fKdRujukanAsal & "' and StatusPasien='" & fStatusPasien & "' and KdKelas='" & fKdKelas & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and KdTriase='" & fKdTriase & "' and Kecamatan='" & fKecamatan & "') and (day(TglPeriksa)=day('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPeriksa)=month('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPeriksa)=year('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
End Function

'Konversi dari SP: AM_DataMorbiditasPasien
Public Function f_AMDataMorbiditasPasien(fNoPendaftaran As String, fNoCM As String, fKdRuangan As String, fKdSubInstalasi As String, fTglPeriksa As Date, fKdDiagnosa As String, fStatusKasus As String, fStatus As String)
    'fStatus: A=Add, M=Min
    Dim fTglLahir As Date
    Dim fJmlPriaTemp As Integer
    Dim fJmlPria As Integer
    Dim fJmlWanitaTemp As Integer
    Dim fJmlWanita As Integer
    Dim fJK As String
    Dim fKelUmur1 As Integer
    Dim fKelUmur2 As Integer
    Dim fKelUmur3 As Integer
    Dim fKelUmur4 As Integer
    Dim fKelUmur5 As Integer
    Dim fKelUmur6 As Integer
    Dim fKelUmur7 As Integer
    Dim fKelUmur8 As Integer
    Dim fKelUmur1Temp As Integer
    Dim fKelUmur2Temp As Integer
    Dim fKelUmur3Temp As Integer
    Dim fKelUmur4Temp As Integer
    Dim fKelUmur5Temp As Integer
    Dim fKelUmur6Temp As Integer
    Dim fKelUmur7Temp As Integer
    Dim fKelUmur8Temp As Integer
    Dim fJmlKunjungan As Integer
    Dim fJmlKunjunganTemp As Integer
    Dim fJmlUmurDlmHari As Integer
    Dim fJmlUmurDlmThn As Integer
    Dim fThnKabisat As Integer
    Dim fJmlHariThnKabisat As Integer
    Dim fKdInstalasi As String
    Dim fJmlPasienOutPria As Integer
    Dim fJmlPasienOutWanita As Integer
    Dim fJmlPasienOutHidup As Integer
    Dim fJmlPasienOutMati As Integer
    Dim fKdKelompokPasien As String

    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdKelompokPasien from PasienDaftar where NoPendaftaran=fNoPendaftaran"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value) Else fKdKelompokPasien = "01"
    Set fRS = Nothing
    fQuery = "select KdInstalasi from Ruangan where KdRuangan=fKdRuangan"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdInstalasi = IIf(IsNull(fRS("KdInstalasi").value), "", fRS("KdInstalasi").value) Else fKdInstalasi = ""
    Set fRS = Nothing
    fQuery = "select JenisKelamin,TglLahir from Pasien where NoCM=fNoCM"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fJK = IIf(IsNull(fRS("JenisKelamin").value), "B", fRS("JenisKelamin").value) Else fJK = "B"
    If fRS.EOF = False Then fTglLahir = IIf(IsNull(fRS("TglLahir").value), "", fRS("TglLahir").value) Else fTglLahir = ""
    fJmlUmurDlmHari = CInt(DateDiff(dd, fTglLahir, fTglPeriksa, vbSunday, vbUseSystem))
    fJmlUmurDlmThn = CInt(DateDiff(yyyy, fTglLahir, fTglPeriksa, vbSunday, vbUseSystem))
    fThnKabisat = CInt(Year(fTglPeriksa) Mod 4)
    If fThnKabisat = 0 Then
        fJmlHariThnKabisat = 366
    Else
        fJmlHariThnKabisat = 365
    End If
    If fKdInstalasi = "01" Or fKdInstalasi = "02" Or fKdInstalasi = "06" Or fKdInstalasi = "11" Then
        Set fRS = Nothing
        fQuery = "select KdDiagnosa from DataMorbiditasPasien where (KdSubInstalasi='" & fKdSubInstalasi & "' and KdRuangan='" & fKdRuangan & "' and KdDiagnosa='" & fKdDiagnosa & "' and StatusKasus='" & fStatusKasus & "' and KdKelompokPasien='" & fKdKelompokPasien & "') and (day(TglPeriksa)=day('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPeriksa)=month('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPeriksa)=year('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'))"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            fKelUmur1Temp = 0
            fKelUmur2Temp = 0
            fKelUmur3Temp = 0
            fKelUmur4Temp = 0
            fKelUmur5Temp = 0
            fKelUmur6Temp = 0
            fKelUmur7Temp = 0
            fKelUmur8Temp = 0
            fJmlPriaTemp = 0
            fJmlWanitaTemp = 0
            fJmlKunjunganTemp = 0
            fJmlKunjungan = fJmlKunjunganTemp + 1
            If fJmlUmurDlmHari < 28 Then
                fKelUmur1 = fKelUmur1Temp + 1
            Else
                fKelUmur1 = fKelUmur1Temp
            End If
            If fJmlUmurDlmHari >= 28 And fJmlUmurDlmHari < fJmlHariThnKabisat Then
                fKelUmur2 = fKelUmur2Temp + 1
            Else
                fKelUmur2 = fKelUmur2Temp
            End If
            If fJmlUmurDlmThn >= 1 And fJmlUmurDlmThn <= 4 Then
                fKelUmur3 = fKelUmur3Temp + 1
            Else
                fKelUmur3 = fKelUmur3Temp
            End If
            If fJmlUmurDlmThn >= 5 And fJmlUmurDlmThn <= 14 Then
                fKelUmur4 = fKelUmur4Temp + 1
            Else
                fKelUmur4 = fKelUmur4Temp
            End If
            If fJmlUmurDlmThn >= 15 And fJmlUmurDlmThn <= 24 Then
                fKelUmur5 = fKelUmur5Temp + 1
            Else
                fKelUmur5 = fKelUmur5Temp
            End If
            If fJmlUmurDlmThn >= 25 And fJmlUmurDlmThn <= 44 Then
                fKelUmur6 = fKelUmur6Temp + 1
            Else
                fKelUmur6 = fKelUmur6Temp
            End If
            If fJmlUmurDlmThn >= 45 And fJmlUmurDlmThn <= 64 Then
                fKelUmur7 = fKelUmur7Temp + 1
            Else
                fKelUmur7 = fKelUmur7Temp
            End If
            If fJmlUmurDlmThn >= 65 Then
                fKelUmur8 = fKelUmur8Temp + 1
            Else
                fKelUmur8 = fKelUmur8Temp
            End If
            If fJK = "L" Then
                fJmlPria = fJmlPriaTemp + 1
                fJmlWanita = fJmlWanitaTemp
            End If
            If fJK = "P" Then
                fJmlPria = fJmlPriaTemp
                fJmlWanita = fJmlWanitaTemp + 1
            End If
            Set fRS2 = Nothing
            fQuery2 = "insert DataMorbiditasPasien values ('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "','" & fKdDiagnosa & "','" & fKdRuangan & "','" & fKdSubInstalasi & "','" & fStatusKasus & "','" & fKdKelompokPasien & "'," & fKelUmur1 & "," & fKelUmur2 & "," & fKelUmur3 & "," & fKelUmur4 & "," & fKelUmur5 & "," & fKelUmur6 & "," & fKelUmur7 & "," & fKelUmur8 & "," & fJmlPria & "," & fJmlWanita & "," & fJmlKunjungan & ")"
            Call msubRecFO(fRS2, fQuery2)
        Else
            Set fRS2 = Nothing
            fQuery2 = "select JmlPasienKel1,JmlPasienKel2,JmlPasienKel3,JmlPasienKel4,JmlPasienKel5,JmlPasienKel6,JmlPasienKel7,JmlPasienKel8,JmlPasienPria,JmlPasienWanita,JmlKunjungan from DataMorbiditasPasien where (KdSubInstalasi='" & fKdSubInstalasi & "' and KdRuangan='" & fKdRuangan & "' and KdDiagnosa='" & fKdDiagnosa & "' and StatusKasus='" & fStatusKasus & "' and KdKelompokPasien='" & fKdKelompokPasien & "') and (day(TglPeriksa)=day('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPeriksa)=month('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPeriksa)=year('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'))"
            Call msubRecFO(fRS2, fQuery2)
            If fRS.EOF = False Then fKelUmur1Temp = IIf(IsNull(fRS("JmlPasienKel1").value), 0, fRS("JmlPasienKel1").value) Else fKelUmur1Temp = 0
            If fRS.EOF = False Then fKelUmur2Temp = IIf(IsNull(fRS("JmlPasienKel2").value), 0, fRS("JmlPasienKel2").value) Else fKelUmur2Temp = 0
            If fRS.EOF = False Then fKelUmur3Temp = IIf(IsNull(fRS("JmlPasienKel3").value), 0, fRS("JmlPasienKel3").value) Else fKelUmur3Temp = 0
            If fRS.EOF = False Then fKelUmur4Temp = IIf(IsNull(fRS("JmlPasienKel4").value), 0, fRS("JmlPasienKel4").value) Else fKelUmur4Temp = 0
            If fRS.EOF = False Then fKelUmur5Temp = IIf(IsNull(fRS("JmlPasienKel5").value), 0, fRS("JmlPasienKel5").value) Else fKelUmur5Temp = 0
            If fRS.EOF = False Then fKelUmur6Temp = IIf(IsNull(fRS("JmlPasienKel6").value), 0, fRS("JmlPasienKel6").value) Else fKelUmur6Temp = 0
            If fRS.EOF = False Then fKelUmur7Temp = IIf(IsNull(fRS("JmlPasienKel7").value), 0, fRS("JmlPasienKel7").value) Else fKelUmur7Temp = 0
            If fRS.EOF = False Then fKelUmur8Temp = IIf(IsNull(fRS("JmlPasienKel8").value), 0, fRS("JmlPasienKel8").value) Else fKelUmur8Temp = 0
            If fRS.EOF = False Then fJmlPriaTemp = IIf(IsNull(fRS("JmlPasienPria").value), 0, fRS("JmlPasienPria").value) Else fJmlPriaTemp = 0
            If fRS.EOF = False Then fJmlWanitaTemp = IIf(IsNull(fRS("JmlPasienWanita").value), 0, fRS("JmlPasienWanita").value) Else fJmlWanitaTemp = 0
            If fRS.EOF = False Then fJmlKunjunganTemp = IIf(IsNull(fRS("JmlKunjungan").value), 0, fRS("JmlKunjungan").value) Else fJmlKunjunganTemp = 0
            If fStatus = "A" Then
                fJmlKunjungan = fJmlKunjunganTemp + 1
                If fJmlUmurDlmHari < 28 Then
                    fKelUmur1 = fKelUmur1Temp + 1
                Else
                    fKelUmur1 = fKelUmur1Temp
                End If
                If fJmlUmurDlmHari >= 28 And fJmlUmurDlmHari < fJmlHariThnKabisat Then
                    fKelUmur2 = fKelUmur2Temp + 1
                Else
                    fKelUmur2 = fKelUmur2Temp
                End If
                If fJmlUmurDlmThn >= 1 And fJmlUmurDlmThn <= 4 Then
                    fKelUmur3 = fKelUmur3Temp + 1
                Else
                    fKelUmur3 = fKelUmur3Temp
                End If
                If fJmlUmurDlmThn >= 5 And fJmlUmurDlmThn <= 14 Then
                    fKelUmur4 = fKelUmur4Temp + 1
                Else
                    fKelUmur4 = fKelUmur4Temp
                End If
                If fJmlUmurDlmThn >= 15 And fJmlUmurDlmThn <= 24 Then
                    fKelUmur5 = fKelUmur5Temp + 1
                Else
                    fKelUmur5 = fKelUmur5Temp
                End If
                If fJmlUmurDlmThn >= 25 And fJmlUmurDlmThn <= 44 Then
                    fKelUmur6 = fKelUmur6Temp + 1
                Else
                    fKelUmur6 = fKelUmur6Temp
                End If
                If fJmlUmurDlmThn >= 45 And fJmlUmurDlmThn <= 64 Then
                    fKelUmur7 = fKelUmur7Temp + 1
                Else
                    fKelUmur7 = fKelUmur7Temp
                End If
                If fJmlUmurDlmThn >= 65 Then
                    fKelUmur8 = fKelUmur8Temp + 1
                Else
                    fKelUmur8 = fKelUmur8Temp
                End If
                If fJK = "L" Then
                    fJmlPria = fJmlPriaTemp + 1
                    fJmlWanita = fJmlWanitaTemp
                End If
                If fJK = "P" Then
                    fJmlPria = fJmlPriaTemp
                    fJmlWanita = fJmlWanitaTemp + 1
                End If
            Else
                fJmlKunjungan = fJmlKunjunganTemp - 1
                If fJmlUmurDlmHari < 28 Then
                    fKelUmur1 = fKelUmur1Temp - 1
                Else
                    fKelUmur1 = fKelUmur1Temp
                End If
                If fJmlUmurDlmHari >= 28 And fJmlUmurDlmHari < fJmlHariThnKabisat Then
                    fKelUmur2 = fKelUmur2Temp - 1
                Else
                    fKelUmur2 = fKelUmur2Temp
                End If
                If fJmlUmurDlmThn >= 1 And fJmlUmurDlmThn <= 4 Then
                    fKelUmur3 = fKelUmur3Temp - 1
                Else
                    fKelUmur3 = fKelUmur3Temp
                End If
                If fJmlUmurDlmThn >= 5 And fJmlUmurDlmThn <= 14 Then
                    fKelUmur4 = fKelUmur4Temp - 1
                Else
                    fKelUmur4 = fKelUmur4Temp
                End If
                If fJmlUmurDlmThn >= 15 And fJmlUmurDlmThn <= 24 Then
                    fKelUmur5 = fKelUmur5Temp - 1
                Else
                    fKelUmur5 = fKelUmur5Temp
                End If
                If fJmlUmurDlmThn >= 25 And fJmlUmurDlmThn <= 44 Then
                    fKelUmur6 = fKelUmur6Temp - 1
                Else
                    fKelUmur6 = fKelUmur6Temp
                End If
                If fJmlUmurDlmThn >= 45 And fJmlUmurDlmThn <= 64 Then
                    fKelUmur7 = fKelUmur7Temp - 1
                Else
                    fKelUmur7 = fKelUmur7Temp
                End If
                If fJmlUmurDlmThn >= 65 Then
                    fKelUmur8 = fKelUmur8Temp - 1
                Else
                    fKelUmur8 = fKelUmur8Temp
                End If
                If fJK = "L" Then
                    fJmlPria = fJmlPriaTemp - 1
                    fJmlWanita = fJmlWanitaTemp
                End If
                If fJK = "P" Then
                    fJmlPria = fJmlPriaTemp
                    fJmlWanita = fJmlWanitaTemp - 1
                End If
            End If
            Set fRS2 = Nothing
            fQuery2 = "update DataMorbiditasPasien set JmlPasienKel1=" & fKelUmur1 & ",JmlPasienKel2=" & fKelUmur2 & ",JmlPasienKel3=" & fKelUmur3 & ",JmlPasienKel4=" & fKelUmur4 & ",JmlPasienKel5=" & fKelUmur5 & ",JmlPasienKel6=" & fKelUmur6 & ",JmlPasienKel7=" & fKelUmur7 & ",JmlPasienKel8=" & fKelUmur8 & ",JmlPasienPria=" & fJmlPria & ",JmlPasienWanita=" & fJmlWanita & ",JmlKunjungan=" & fJmlKunjungan & " where " _
            & "where (KdSubInstalasi='" & fKdSubInstalasi & "' and KdRuangan='" & fKdRuangan & "' and KdDiagnosa='" & fKdDiagnosa & "' and StatusKasus='" & fStatusKasus & "' and KdKelompokPasien='" & fKdKelompokPasien & "') and (day(TglPeriksa)=day('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPeriksa)=month('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPeriksa)=year('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'))"
            Call msubRecFO(fRS2, fQuery2)
        End If
    End If
    If fKdInstalasi = "03" Then
        Set fRS = Nothing
        fQuery = "select KdDiagnosa from DataMorbiditasPasienRI where (KdSubInstalasi='" & fKdSubInstalasi & "' and KdRuangan='" & fKdRuangan & "' and KdDiagnosa='" & fKdDiagnosa & "' and StatusKasus='" & fStatusKasus & "' and KdKelompokPasien='" & fKdKelompokPasien & "') and (day(TglPeriksa)=day('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPeriksa)=month('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPeriksa)=year('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'))"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            fKelUmur1Temp = 0
            fKelUmur2Temp = 0
            fKelUmur3Temp = 0
            fKelUmur4Temp = 0
            fKelUmur5Temp = 0
            fKelUmur6Temp = 0
            fKelUmur7Temp = 0
            fKelUmur8Temp = 0
            fJmlPasienOutPria = 0
            fJmlPasienOutWanita = 0
            fJmlPasienOutHidup = 0
            fJmlPasienOutMati = 0
            If fJmlUmurDlmHari < 28 Then
                fKelUmur1 = fKelUmur1Temp + 1
            Else
                fKelUmur1 = fKelUmur1Temp
            End If
            If fJmlUmurDlmHari >= 28 And fJmlUmurDlmHari < fJmlHariThnKabisat Then
                fKelUmur2 = fKelUmur2Temp + 1
            Else
                fKelUmur2 = fKelUmur2Temp
            End If
            If fJmlUmurDlmThn >= 1 And fJmlUmurDlmThn <= 4 Then
                fKelUmur3 = fKelUmur3Temp + 1
            Else
                fKelUmur3 = fKelUmur3Temp
            End If
            If fJmlUmurDlmThn >= 5 And fJmlUmurDlmThn <= 14 Then
                fKelUmur4 = fKelUmur4Temp + 1
            Else
                fKelUmur4 = fKelUmur4Temp
            End If
            If fJmlUmurDlmThn >= 15 And fJmlUmurDlmThn <= 24 Then
                fKelUmur5 = fKelUmur5Temp + 1
            Else
                fKelUmur5 = fKelUmur5Temp
            End If
            If fJmlUmurDlmThn >= 25 And fJmlUmurDlmThn <= 44 Then
                fKelUmur6 = fKelUmur6Temp + 1
            Else
                fKelUmur6 = fKelUmur6Temp
            End If
            If fJmlUmurDlmThn >= 45 And fJmlUmurDlmThn <= 64 Then
                fKelUmur7 = fKelUmur7Temp + 1
            Else
                fKelUmur7 = fKelUmur7Temp
            End If
            If fJmlUmurDlmThn >= 65 Then
                fKelUmur8 = fKelUmur8Temp + 1
            Else
                fKelUmur8 = fKelUmur8Temp
            End If
            Set fRS2 = Nothing
            fQuery2 = "insert DataMorbiditasPasienRI values ('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "','" & fKdDiagnosa & "','" & fKdRuangan & "','" & fKdSubInstalasi & "','" & fStatusKasus & "','" & fKdKelompokPasien & "'," & fKelUmur1 & "," & fKelUmur2 & "," & fKelUmur3 & "," & fKelUmur4 & "," & fKelUmur5 & "," & fKelUmur6 & "," & fKelUmur7 & "," & fKelUmur8 & "," & fJmlPasienOutPria & "," & fJmlPasienOutWanita & "," & fJmlPasienOutHidup & "," & fJmlPasienOutMati & ")"
            Call msubRecFO(fRS2, fQuery2)
        Else
            Set fRS = Nothing
            fQuery = "select fKelUmur1Temp=JmlPasienKel1,fKelUmur2Temp=JmlPasienKel2,fKelUmur3Temp=JmlPasienKel3,fKelUmur4Temp=JmlPasienKel4,fKelUmur5Temp=JmlPasienKel5,fKelUmur6Temp=JmlPasienKel6,fKelUmur7Temp=JmlPasienKel7,fKelUmur8Temp=JmlPasienKel8 from DataMorbiditasPasienRI where (KdSubInstalasi=fKdSubInstalasi and KdRuangan=fKdRuangan and KdDiagnosa=fKdDiagnosa and StatusKasus=fStatusKasus and KdKelompokPasien=fKdKelompokPasien) and (day(TglPeriksa)=day(fTglPeriksa) and month(TglPeriksa)=month(fTglPeriksa) and year(TglPeriksa)=year(fTglPeriksa))"
            Call msubRecFO(fRS, fQuery)
            If fStatus = "A" Then
                If fJmlUmurDlmHari < 28 Then
                    fKelUmur1 = fKelUmur1Temp + 1
                Else
                    fKelUmur1 = fKelUmur1Temp
                End If
                If fJmlUmurDlmHari >= 28 And fJmlUmurDlmHari < fJmlHariThnKabisat Then
                    fKelUmur2 = fKelUmur2Temp + 1
                Else
                    fKelUmur2 = fKelUmur2Temp
                End If
                If fJmlUmurDlmThn >= 1 And fJmlUmurDlmThn <= 4 Then
                    fKelUmur3 = fKelUmur3Temp + 1
                Else
                    fKelUmur3 = fKelUmur3Temp
                End If
                If fJmlUmurDlmThn >= 5 And fJmlUmurDlmThn <= 14 Then
                    fKelUmur4 = fKelUmur4Temp + 1
                Else
                    fKelUmur4 = fKelUmur4Temp
                End If
                If fJmlUmurDlmThn >= 15 And fJmlUmurDlmThn <= 24 Then
                    fKelUmur5 = fKelUmur5Temp + 1
                Else
                    fKelUmur5 = fKelUmur5Temp
                End If
                If fJmlUmurDlmThn >= 25 And fJmlUmurDlmThn <= 44 Then
                    fKelUmur6 = fKelUmur6Temp + 1
                Else
                    fKelUmur6 = fKelUmur6Temp
                End If
                If fJmlUmurDlmThn >= 45 And fJmlUmurDlmThn <= 64 Then
                    fKelUmur7 = fKelUmur7Temp + 1
                Else
                    fKelUmur7 = fKelUmur7Temp
                End If
                If fJmlUmurDlmThn >= 65 Then
                    fKelUmur8 = fKelUmur8Temp + 1
                Else
                    fKelUmur8 = fKelUmur8Temp
                End If
            Else
                If fJmlUmurDlmHari < 28 Then
                    fKelUmur1 = fKelUmur1Temp - 1
                Else
                    fKelUmur1 = fKelUmur1Temp
                End If
                If fJmlUmurDlmHari >= 28 And fJmlUmurDlmHari < fJmlHariThnKabisat Then
                    fKelUmur2 = fKelUmur2Temp - 1
                Else
                    fKelUmur2 = fKelUmur2Temp
                End If
                If fJmlUmurDlmThn >= 1 And fJmlUmurDlmThn <= 4 Then
                    fKelUmur3 = fKelUmur3Temp - 1
                Else
                    fKelUmur3 = fKelUmur3Temp
                End If
                If fJmlUmurDlmThn >= 5 And fJmlUmurDlmThn <= 14 Then
                    fKelUmur4 = fKelUmur4Temp - 1
                Else
                    fKelUmur4 = fKelUmur4Temp
                End If
                If fJmlUmurDlmThn >= 15 And fJmlUmurDlmThn <= 24 Then
                    fKelUmur5 = fKelUmur5Temp - 1
                Else
                    fKelUmur5 = fKelUmur5Temp
                End If
                If fJmlUmurDlmThn >= 25 And fJmlUmurDlmThn <= 44 Then
                    fKelUmur6 = fKelUmur6Temp - 1
                Else
                    fKelUmur6 = fKelUmur6Temp
                End If
                If fJmlUmurDlmThn >= 45 And fJmlUmurDlmThn <= 64 Then
                    fKelUmur7 = fKelUmur7Temp - 1
                Else
                    fKelUmur7 = fKelUmur7Temp
                End If
                If fJmlUmurDlmThn >= 65 Then
                    fKelUmur8 = fKelUmur8Temp - 1
                Else
                    fKelUmur8 = fKelUmur8Temp
                End If
            End If
            Set fRS2 = Nothing
            fQuery2 = "update DataMorbiditasPasienRI set JmlPasienKel1=" & fKelUmur1 & ",JmlPasienKel2=" & fKelUmur2 & ",JmlPasienKel3=" & fKelUmur3 & ",JmlPasienKel4=" & fKelUmur4 & ",JmlPasienKel5=" & fKelUmur5 & ",JmlPasienKel6=" & fKelUmur6 & ",JmlPasienKel7=" & fKelUmur7 & ",JmlPasienKel8=" & fKelUmur8 & " where " _
            & "(KdSubInstalasi='" & fKdSubInstalasi & "' and KdRuangan='" & fKdRuangan & "' and KdDiagnosa='" & fKdDiagnosa & "' and StatusKasus='" & fStatusKasus & "' and KdKelompokPasien='" & fKdKelompokPasien & "') and (day(TglPeriksa)=day('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPeriksa)=month('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPeriksa)=year('" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'))"
            Call msubRecFO(fRS2, fQuery2)
        End If
    End If
End Function

'Konversi dari SP: AM_RekapitulasiDistribusiBarangNonMedis
Public Function f_AMRekapitulasiDistribusiBarangNonMedis(fTglTransaksi As Date, fKdRuangan As String, fKdRuanganPenerima As String, fKdBarang As String, fKdAsal As String, fKdMerk As String, fKdType As String, fKdBahanBarang As String, fJmlBarang As Double, fHargaNetto As Currency, fHargaJual As Currency, fDiscount As Currency, fStatus As String)
    'fStatus : A=Add & Ubah; M=Minus
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String
    Set fRS = Nothing
    fQuery = "select KdRuangan from RekapitulasiDistribusiBarangNonMedis where (KdRuangan='" & fKdRuangan & "' and KdRuanganPenerima='" & fKdRuanganPenerima & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdMerk='" & fKdMerk & "' and KdType='" & fKdType & "' and KdBahanBarang='" & fKdBahanBarang & "') and (datepart(hh, TglKirim)=datepart(hh, '" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and day(TglKirim)=day('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and month(TglKirim)=month('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and year(TglKirim)=year('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into RekapitulasiDistribusiBarangNonMedis values('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuangan & "','" & fKdRuanganPenerima & "','" & fKdBarang & "','" & fKdAsal & "','" & fKdMerk & "','" & fKdType & "','" & fKdBahanBarang & "'," & fJmlBarang & "," & fJmlBarang & " * " & fHargaNetto & "," & fJmlBarang & " * " & fHargaJual & "," & fJmlBarang & " * " & fDiscount & ",null)"
    Else
        If UCase(fStatus) = "A" Then
            fQuery = "update RekapitulasiDistribusiBarangNonMedis set JmlKirim=JmlKirim + " & fJmlBarang & ",TotalNetto=TotalNetto + (" & fJmlBarang & " * " & fHargaNetto & "),TotalJual=TotalJual + (" & fJmlBarang & " * " & fHargaJual & "),TotalDiscount=TotalDiscount + (" & fJmlBarang & " * " & fDiscount & ") where (KdRuangan='" & fKdRuangan & "' and KdRuanganPenerima='" & fKdRuanganPenerima & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdMerk='" & fKdMerk & "' and KdType='" & fKdType & "' and KdBahanBarang='" & fKdBahanBarang & "') and (datepart(hh, TglKirim)=datepart(hh, '" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and day(TglKirim)=day('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and month(TglKirim)=month('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and year(TglKirim)=year('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery = "update RekapitulasiDistribusiBarangNonMedis set JmlKirim=JmlKirim - " & fJmlBarang & ",TotalNetto=TotalNetto - (" & fJmlBarang & " * " & fHargaNetto & "),TotalJual=TotalJual - (" & fJmlBarang & " * " & fHargaJual & "),TotalDiscount=TotalDiscount - (" & fJmlBarang & " * " & fDiscount & ") where (KdRuangan='" & fKdRuangan & "' and KdRuanganPenerima='" & fKdRuanganPenerima & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdMerk='" & fKdMerk & "' and KdType='" & fKdType & "' and KdBahanBarang='" & fKdBahanBarang & "') and (datepart(hh, TglKirim)=datepart(hh, '" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and day(TglKirim)=day('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and month(TglKirim)=month('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and year(TglKirim)=year('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
End Function

'Konversi dari SP: AM_RekapitulasiDistribusiBarangMedis
Public Function f_AMRekapitulasiDistribusiBarangMedis(fTglTransaksi As Date, fKdRuanganPenerima As String, fKdRuangan As String, fKdBarang As String, fKdAsal As String, fJmlBarang As Double, fHargaNetto As Currency, fHargaJual As Currency, fDiscount As Currency, fStatus As String)
    'fStatus: A=Add & Ubah; M=Minus
    Dim fStokAwal As Double
    Dim fKdBarangTemp As String
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select KdRuangan from RekapitulasiDistribusiBarangMedis where KdRuangan='" & fKdRuangan & "' and KdRuanganPenerima='" & fKdRuanganPenerima & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and (datepart(hh, TglKirim)=datepart(hh,'" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and day(TglKirim)=day('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and month(TglKirim)=month('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and year(TglKirim)=year('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into RekapitulasiDistribusiBarangMedis values('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuangan & "','" & fKdRuanganPenerima & "','" & fKdBarang & "','" & fKdAsal & "'," & fJmlBarang & "," & fJmlBarang & " * " & fHargaNetto & "," & fJmlBarang & " * " & fHargaJual & "," & fJmlBarang & " * " & fDiscount & ",null)"
    Else
        If UCase(fStatus) = "A" Then
            fQuery = "update RekapitulasiDistribusiBarangMedis set JmlKirim=JmlKirim + " & fJmlBarang & ",TotalNetto=TotalNetto + (" & fJmlBarang & " * " & fHargaNetto & "),TotalJual=TotalJual + (" & fJmlBarang & " * " & fHargaJual & "),TotalDiscount=TotalDiscount + (" & fJmlBarang & " * " & fDiscount & ") where KdRuangan='" & fKdRuangan & "' and KdRuanganPenerima='" & fKdRuanganPenerima & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and (datepart(hh, TglKirim)=datepart(hh,'" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and day(TglKirim)=day('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and month(TglKirim)=month('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and year(TglKirim)=year('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery = "update RekapitulasiDistribusiBarangMedis set JmlKirim=JmlKirim - " & fJmlBarang & ",TotalNetto=TotalNetto - (" & fJmlBarang & " * " & fHargaNetto & "),TotalJual=TotalJual - (" & fJmlBarang & " * " & fHargaJual & "),TotalDiscount=TotalDiscount - (" & fJmlBarang & " * " & fDiscount & ") where KdRuangan='" & fKdRuangan & "' and KdRuanganPenerima='" & fKdRuanganPenerima & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and (datepart(hh, TglKirim)=datepart(hh,'" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and day(TglKirim)=day('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and month(TglKirim)=month('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "') and year(TglKirim)=year('" & Format(fTglTransaksi, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
End Function

'Konversi dari SP: AM_RekapitulasiJasaBPDokterForRemunerasiOnUpdateDokter
Public Function f_AMRekapitulasiJasaBPDokterForRemunerasiOnUpdateDokter(fNoBKM As String, fNoStruk As String, fKdRuanganPelayanan As String, fTglMasuk As Date, fStatus As String)
    'fStatus: A=Add; M=Min
    Dim fTglBKM As Date
    Dim fTotalTarif As Currency
    Dim fJmlBayarTotal As Currency
    Dim fJmlHutangPenjaminTotal As Currency
    Dim fJmlTanggunganRSTotal As Currency
    Dim fJmlPembebasanTotal As Currency
    Dim fSisaTagihanTotal As Currency
    Dim fKdRuanganKasir As String
    Dim fKdKelompokPasien As String
    Dim fNoPendaftaran As String
    Dim fIdPenjamin As String
    Dim fJmlPelayanan As Integer
    Dim fTarif As Currency
    Dim fKdRuangan As String
    Dim fKdKomponen As String
    Dim fKdAsal As String
    Dim fJmlBayar As Currency
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fJmlPembebasan As Currency
    Dim fSisaTagihan As Currency
    Dim fKdPelayananRS As String
    Dim fTglPelayanan As Date
    Dim fKdSubInstalasi As String
    Dim fKdRuanganAsal As String
    Dim fNoLab_Rad As Variant
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fKdKelas As String
    Dim fIdPegawai As Variant
    Dim fKdJenisPegawai As String
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select TglBKM,KdRuangan from StrukBuktiKasMasuk where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fTglBKM = IIf(IsNull(fRS("TglBKM").value), "", fRS("TglBKM").value)
        fKdRuanganKasir = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
    End If
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien from StrukPelayananPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value)
        fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
    End If
    Set fRS = Nothing
    fQuery = "select NoPendaftaran,KdRuangan,KdPelayananRS,KdKomponen,TglPelayanan from RekapKomponenBiayaPelayananTM where NoBKM='" & fNoBKM & "' and NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuanganPelayanan & "' and TglPelayanan='" & Format(fTglMasuk, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fNoPendaftaran = IIf(IsNull(fRS("NoPendaftaran").value), "", fRS("NoPendaftaran").value)
        fKdRuangan = IIf(IsNull(fRS("KdRuangan").value), "", fRS("KdRuangan").value)
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        fTglPelayanan = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
        Set fRS2 = Nothing
        fQuery2 = "select StatusAPBD,KdSubInstalasi,NoLab_Rad from BiayaPelayanan where NoStruk='" & fNoStruk & "' and NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then
            fKdSubInstalasi = IIf(IsNull(fRS2("KdSubInstalasi").value), "", fRS2("KdSubInstalasi").value)
            fKdAsal = IIf(IsNull(fRS2("StatusAPBD").value), "", fRS2("StatusAPBD").value)
            fNoLab_Rad = fRS2("NoLab_Rad").value
        End If
        Set fRS2 = Nothing
        fQuery2 = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "'," & fNoLab_Rad & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fKdRuanganAsal = fRS2("KdRuanganAsal").value Else fKdRuanganAsal = ""
        Set fRS2 = Nothing
        fQuery2 = "select JmlPelayanan,Tarif,JmlBayar,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,SisaTagihan,KdDetailJenisJasaPelayanan,KdKelas,IdPegawai from RekapKomponenBiayaPelayananTM where NoBKM='" & fNoBKM & "' and NoStruk='" & fNoStruk & "' and NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
        Call msubRecFO(fRS2, fQuery2)
        While fRS2.EOF = False
            fJmlPelayanan = IIf(IsNull(fRS2("JmlPelayanan").value), 0, fRS2("JmlPelayanan").value)
            fTarif = IIf(IsNull(fRS2("Tarif").value), 0, fRS2("Tarif").value)
            fJmlBayar = IIf(IsNull(fRS2("JmlBayar").value), 0, fRS2("JmlBayar").value)
            fJmlHutangPenjamin = IIf(IsNull(fRS2("JmlHutangPenjamin").value), 0, fRS2("JmlHutangPenjamin").value)
            fJmlTanggunganRS = IIf(IsNull(fRS2("JmlTanggunganRS").value), 0, fRS2("JmlTanggunganRS").value)
            fJmlPembebasan = IIf(IsNull(fRS2("JmlPembebasan").value), 0, fRS2("JmlPembebasan").value)
            fSisaTagihan = IIf(IsNull(fRS2("SisaTagihan").value), 0, fRS2("SisaTagihan").value)
            fKdDetailJenisJasaPelayanan = IIf(IsNull(fRS2("KdDetailJenisJasaPelayanan").value), "01", fRS2("KdDetailJenisJasaPelayanan").value)
            fKdKelas = IIf(IsNull(fRS2("KdKelas").value), "01", fRS2("KdKelas").value)
            fIdPegawai = fRS2("IdPegawai").value
            Set fRS3 = Nothing
            fQuery3 = "KdJenisPegawai from DataPegawai where IdPegawai=" & fIdPegawai & ""
            Call msubRecFO(fRS3, fQuery3)
            If fRS3.EOF = False Then fKdJenisPegawai = IIf(IsNull(fRS3("KdJenisPegawai").value), "", fRS3("KdJenisPegawai").value) Else fKdJenisPegawai = ""
            If fKdJenisPegawai = "001" Then
                fTotalTarif = fJmlPelayanan * fTarif
                fJmlBayarTotal = fJmlPelayanan * fJmlBayar
                fJmlHutangPenjaminTotal = fJmlPelayanan * fJmlHutangPenjamin
                fJmlTanggunganRSTotal = fJmlPelayanan * fJmlTanggunganRS
                fJmlPembebasanTotal = fJmlPelayanan * fJmlPembebasan
                fSisaTagihanTotal = fJmlPelayanan * fSisaTagihan
                Set fRS3 = Nothing
                fQuery3 = "select KdRuangan from RekapitulasiJasaBPDokter4Remunerasi where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and IdPegawai=" & fIdPegawai & ") and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
                Call msubRecFO(fRS3, fQuery3)
                If fRS3.EOF = True Then
                    fQuery3 = "insert into RekapitulasiJasaBPDokter4Remunerasi values('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuanganKasir & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdSubInstalasi & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdDetailJenisJasaPelayanan & "','" & fKdKelas & "','" & fKdPelayananRS & "','" & fKdKomponen & "','" & fKdAsal & "'," & fIdPegawai & "," & fJmlPelayanan & "," & fTotalTarif & "," & fJmlBayarTotal & "," & fJmlHutangPenjaminTotal & "," & fJmlTanggunganRSTotal & "," & fJmlPembebasanTotal & "," & fSisaTagihanTotal & ")"
                Else
                    If fStatus = "A" Then
                        fQuery3 = "update RekapitulasiJasaBPDokter4Remunerasi set JmlPelayanan=JmlPelayanan+" & fJmlPelayanan & ",TotalBiaya=TotalBiaya+" & fTotalTarif & ", JmlBayar=JmlBayar+" & fJmlBayarTotal & ", JmlHutangPenjamin=JmlHutangPenjamin+" & fJmlHutangPenjaminTotal & ", JmlTanggunganRS=JmlTanggunganRS+" & fJmlTanggunganRSTotal & ", JmlPembebasan=JmlPembebasan+" & fJmlPembebasanTotal & ", SisaTagihan=SisaTagihan+" & fSisaTagihanTotal & " where " _
                        & "(KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and IdPegawai=" & fIdPegawai & ") and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
                    Else
                        fQuery3 = "update RekapitulasiJasaBPDokter4Remunerasi set JmlPelayanan=JmlPelayanan+" & fJmlPelayanan & ",TotalBiaya=TotalBiaya+" & fTotalTarif & ", JmlBayar=JmlBayar+" & fJmlBayarTotal & ", JmlHutangPenjamin=JmlHutangPenjamin+" & fJmlHutangPenjaminTotal & ", JmlTanggunganRS=JmlTanggunganRS+" & fJmlTanggunganRSTotal & ", JmlPembebasan=JmlPembebasan+" & fJmlPembebasanTotal & ", SisaTagihan=SisaTagihan+" & fSisaTagihanTotal & " where " _
                        & "(KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and IdPegawai=" & fIdPegawai & ") and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
                    End If
                End If
                Set fRS3 = Nothing
                Call msubRecFO(fRS3, fQuery3)
            End If
            fRS2.MoveNext
        Wend
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Add_PostingDataKasirPenerimaan
Public Function f_AddPostingDataKasirPenerimaan(fNoPosting As String, fNoBKM As String, fNoStruk As String, fJmlHrsDibayar As Currency, fJmlPembebasan As Currency, fJmlDiscount As Currency, fJmlBayar As Currency, fSisaTagihan As Currency, fJenisTransaksi As String)
    'fJenisTransaksi: TM=Tindakan Medis, OA=Obat & Alkes, AP=Apotik
    Dim fTotalPembebasanApotik As Currency
    Dim fNoBKMBefore As String
    Dim fPembayaranKe As Integer
    Dim fPembayaranKeBefore As Integer
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select PembayaranKe from PembayaranTagihanPasien where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fPembayaranKe = IIf(IsNull(fRS("PembayaranKe").value), 0, fRS("PembayaranKe").value) Else fPembayaranKe = 0
    If fPembayaranKe = 1 Then
        Set fRS = Nothing
        fQuery = "select NoPosting from PostingDataKasirPenerimaan where NoPosting='" & fNoPosting & "' and NoBKM='" & fNoBKM & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            Set fRS2 = Nothing
            fQuery2 = "insert into PostingDataKasirPenerimaan values('" & fNoPosting & "','" & fNoBKM & "','" & fNoStruk & "')"
            Call msubRecFO(fRS2, fQuery2)
        End If
        If UCase(fJenisTransaksi) = "TM" Then
            'add sementara
            strSQL = "SELECT DISTINCT NoBKM  FROM   RekapKomponenBiayaPelayananTM where NoBKM = '" & fNoBKM & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = True Then
                Call f_AddRekapKomponenBiayaPelayananTM(fNoBKM, fNoStruk, fJmlHrsDibayar, fJmlBayar, fJmlPembebasan, fSisaTagihan, fJmlDiscount)
            End If
        End If
        If UCase(fJenisTransaksi) = "OA" Then
            'add sementara
            strSQL = "SELECT DISTINCT NoBKM  FROM   RekapKomponenBiayaPelayananOA where NoBKM = '" & fNoBKM & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = True Then
                Call f_AddRekapKomponenBiayaPelayananOA(fNoBKM, fNoStruk, fJmlHrsDibayar, fJmlBayar, fJmlPembebasan, fSisaTagihan, fJmlDiscount)
            End If
        End If
        If UCase(fJenisTransaksi) = "AP" Then
            'add sementara
            strSQL = "SELECT DISTINCT NoBKM  FROM   RekapKomponenBiayaPelayananApotik where NoBKM = '" & fNoBKM & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = True Then
                fTotalPembebasanApotik = fJmlPembebasan + fJmlDiscount
                Call f_AddRekapKomponenBiayaPelayananApotik(fNoBKM, fNoStruk, fJmlHrsDibayar, fJmlBayar, fTotalPembebasanApotik, fSisaTagihan)
            End If
        End If
    Else
        fPembayaranKeBefore = fPembayaranKe - 1
        Set fRS = Nothing
        fQuery = "select NoBKM from PembayaranTagihanPasien where NoStruk= '" & fNoStruk & "' and PembayaranKe=" & fPembayaranKeBefore & " "
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then
            fNoBKMBefore = IIf(IsNull(fRS("NoBKM").value), "", fRS("NoBKM").value)
            Set fRS = Nothing
            fQuery = "select NoPosting from PostingDataKasirPenerimaan where NoPosting='" & fNoPosting & "' and NoBKM='" & fNoBKM & "'"
            Call msubRecFO(fRS, fQuery)
            If fRS.EOF = True Then
                Set fRS2 = Nothing
                fQuery2 = "insert into PostingDataKasirPenerimaan values('" & fNoPosting & "','" & fNoBKM & "','" & fNoStruk & "')"
                Call msubRecFO(fRS2, fQuery2)
            End If
            If UCase(fJenisTransaksi) = "TM" Then
                Call f_AddRekapKomponenBiayaPelayananTMKredit(fNoBKM, fNoBKMBefore, fNoStruk, fJmlBayar, fJmlPembebasan, fSisaTagihan)
            End If
            If UCase(fJenisTransaksi) = "OA" Then
                Call f_AddRekapKomponenBiayaPelayananOAKredit(fNoBKM, fNoBKMBefore, fNoStruk, fJmlBayar, fJmlPembebasan, fSisaTagihan)
            End If
            If UCase(fJenisTransaksi) = "AP" Then
                Call f_AddRekapKomponenBiayaPelayananApotikKredit(fNoBKM, fNoBKMBefore, fNoStruk, fJmlBayar, fJmlPembebasan, fSisaTagihan)
            End If
        End If
    End If
End Function

'Konversi dari SP: Add_PembayaranReturStrukPelayananPasien
Public Function f_AddPembayaranReturStrukPelayananPasien(fNoRetur As String, fNoBKM As String, fNoStruk As String, fTotalBiaya As Currency, fTotalPpn As Currency, fTotalDiscount As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlHarusDiretur As Currency, fNoBKK As String)
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String
    Set fRS = Nothing
    fQuery = "insert into PembayaranReturStrukPelayananPasien values('" & fNoRetur & "','" & fNoBKM & "','" & fNoStruk & "'," & fTotalBiaya & "," & fTotalPpn & "," & fTotalDiscount & "," & fJmlHutangPenjamin & "," & fJmlTanggunganRS & "," & fJmlHarusDiretur & ")"
    Call msubRecFO(fRS, fQuery)
    Set fRS = Nothing
    fQuery = "UPDATE Retur SET NoBKK = '" & fNoBKK & "' WHERE (NoRetur = '" & fNoRetur & "')"
    Call msubRecFO(fRS, fQuery)
    Set fRS = Nothing
    fQuery = "select NoBKM from RekapKomponenBiayaPelayananTM where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        Call f_DeleteRekapKomponenBiayaPelayananTM(fNoBKM, fNoStruk, "M")
    End If
    Set fRS = Nothing
    fQuery = "select NoBKM from RekapKomponenBiayaPelayananOA where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        Call f_DeleteRekapKomponenBiayaPelayananOA(fNoBKM, fNoStruk, "M")
    End If
End Function

'Konversi dari SP: Add_PembayaranTagihanHutangPenjaminClaim
Public Function f_AddPembayaranTagihanHutangPenjaminClaim(fNoBKM As String, fNoBKMSebelumnya As String, fNoStruk As String, fJmlBayar As Currency, fJmlPembebasan As Currency, fSisaHutangPenjamin As Currency, fStatusPiutang As String)
    'fStatusPiutang: TM=Tindakan Medis; OA=Obat Alkes; AP=Penjualan Apotik; SA=TM & OA
    Dim fMinPembayaranKe As Integer
    Dim fTempPembayaranKe As Integer
    Dim fPembayaranKe As Integer
    Dim fMaksPembayaranKe As Integer
    Dim fJmlBayarSebelumnya As Currency
    Dim fNoBKMClaimSebelumnya As String
    Dim fJmlPembebasanSebelumnya As Currency
    Dim fJmlSudahDibayar As Currency
    Dim fJmlBayarTM As Currency
    Dim fJmlPembebasanTM As Currency
    Dim fSisaHutangPenjaminTM As Currency
    Dim fJmlBayarOA As Currency
    Dim fJmlPembebasanOA As Currency
    Dim fSisaHutangPenjaminOA As Currency
    Dim fJmlBayarApotik As Currency
    Dim fJmlPembebasanApotik As Currency
    Dim fSisaHutangPenjaminApotik As Currency
    Dim fTotalJmlSudahDibayarTM As Currency
    Dim fTotalJmlPembebasanTM As Currency
    Dim fTotalSisaHutangPenjaminTM As Currency
    Dim fTotalJmlSudahDibayarOA As Currency
    Dim fTotalJmlPembebasanOA As Currency
    Dim fTotalSisaHutangPenjaminOA As Currency
    Dim fTotalJmlSudahDibayarApotik As Currency
    Dim fTotalJmlPembebasanApotik As Currency
    Dim fTotalSisaHutangPenjaminApotik As Currency
    Dim fSisaHutangPenjaminSebelumnya As Currency
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select NoStruk from PembayaranClaimPenjaminPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        Set fRS2 = Nothing
        fQuery2 = "select JmlHutangPenjamin from TotalBiayaPelayananTM where NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fTotalSisaHutangPenjaminTM = IIf(IsNull(fRS2("JmlHutangPenjamin").value), 0, fRS2("JmlHutangPenjamin").value) Else fTotalSisaHutangPenjaminTM = 0
        Set fRS2 = Nothing
        fQuery2 = "select JmlHutangPenjamin from TotalBiayaPelayananOA where NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fTotalSisaHutangPenjaminOA = IIf(IsNull(fRS2("JmlHutangPenjamin").value), 0, fRS2("JmlHutangPenjamin").value) Else fTotalSisaHutangPenjaminOA = 0
        Set fRS2 = Nothing
        fQuery2 = "select JmlHutangPenjamin from StrukPelayananPasien where NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fTotalSisaHutangPenjaminApotik = IIf(IsNull(fRS2("JmlHutangPenjamin").value), 0, fRS2("JmlHutangPenjamin").value) Else fTotalSisaHutangPenjaminApotik = 0
        fSisaHutangPenjaminSebelumnya = fTotalSisaHutangPenjaminTM + fTotalSisaHutangPenjaminOA
        Set fRS2 = Nothing
        fQuery2 = "insert into PembayaranClaimPenjaminPasien values('" & fNoBKM & "','" & fNoStruk & "',0," & fJmlBayar & "," & fJmlPembebasan & "," & fSisaHutangPenjamin & ",1,'" & fStatusPiutang & "')"
        Call msubRecFO(fRS2, fQuery2)
        If UCase(fStatusPiutang) = "SA" Then
            If fTotalSisaHutangPenjaminTM <> 0 Then
                fJmlBayarTM = (CDec(fTotalSisaHutangPenjaminTM) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlBayar)
                fJmlPembebasanTM = (CDec(fTotalSisaHutangPenjaminTM) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlPembebasan)
                fSisaHutangPenjaminTM = (CDec(fTotalSisaHutangPenjaminTM) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fSisaHutangPenjamin)
                Call f_AddRekapKomponenBiayaPelayananTMClaim(fNoBKM, fNoBKMSebelumnya, fNoStruk, fJmlBayarTM)
            End If
            If fTotalSisaHutangPenjaminOA <> 0 Then
                fJmlBayarOA = (CDec(fTotalSisaHutangPenjaminOA) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlBayar)
                fJmlPembebasanOA = (CDec(fTotalSisaHutangPenjaminOA) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlPembebasan)
                fSisaHutangPenjaminOA = (CDec(fTotalSisaHutangPenjaminOA) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fSisaHutangPenjamin)
                Call f_AddRekapKomponenBiayaPelayananOAClaim(fNoBKM, fNoBKMSebelumnya, fNoStruk, fJmlBayarOA)
            End If
        End If
        If fTotalSisaHutangPenjaminTM <> 0 And UCase(fStatusPiutang) = "TM" Then
            fJmlBayarTM = (CDec(fTotalSisaHutangPenjaminTM) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlBayar)
            fJmlPembebasanTM = (CDec(fTotalSisaHutangPenjaminTM) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlPembebasan)
            fSisaHutangPenjaminTM = (CDec(fTotalSisaHutangPenjaminTM) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fSisaHutangPenjamin)
            Call f_AddRekapKomponenBiayaPelayananTMClaim(fNoBKM, fNoBKMSebelumnya, fNoStruk, fJmlBayarTM)
        End If
        If fTotalSisaHutangPenjaminOA <> 0 And UCase(fStatusPiutang) = "OA" Then
            fJmlBayarOA = (CDec(fTotalSisaHutangPenjaminOA) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlBayar)
            fJmlPembebasanOA = (CDec(fTotalSisaHutangPenjaminOA) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlPembebasan)
            fSisaHutangPenjaminOA = (CDec(fTotalSisaHutangPenjaminOA) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fSisaHutangPenjamin)
            Call f_AddRekapKomponenBiayaPelayananOAClaim(fNoBKM, fNoBKMSebelumnya, fNoStruk, fJmlBayarOA)
        End If
        If fTotalSisaHutangPenjaminApotik <> 0 And UCase(fStatusPiutang) = "AP" Then
            fJmlBayarApotik = fJmlBayar
            fJmlPembebasanApotik = fJmlPembebasan
            fSisaHutangPenjaminApotik = fSisaHutangPenjamin
            Call f_AddRekapKomponenBiayaPelayananApotikClaim(fNoBKM, fNoBKMSebelumnya, fNoStruk, fJmlBayarApotik)
        End If
    Else
        Set fRS2 = Nothing
        fQuery2 = "select max(PembayaranKe) as PembayaranKeMax from PembayaranClaimPenjaminPasien where NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fTempPembayaranKe = IIf(IsNull(fRS2("PembayaranKeMax").value), 0, fRS2("PembayaranKeMax").value) Else fTempPembayaranKe = 0
        If fTempPembayaranKe = 0 Then
            fPembayaranKe = 1
        Else
            fPembayaranKe = fTempPembayaranKe + 1
            fMaksPembayaranKe = fTempPembayaranKe
        End If
        Set fRS2 = Nothing
        fQuery2 = "select NoBKM,JmlSudahDibayar,SisaHutangPenjamin,JmlPembebasan from PembayaranClaimPenjaminPasien where NoStruk='" & fNoStruk & "' and PembayaranKe=" & fMaksPembayaranKe & ""
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fNoBKMClaimSebelumnya = IIf(IsNull(fRS2("NoBKM").value), "", fRS2("NoBKM").value) Else fNoBKMClaimSebelumnya = ""
        If fRS2.EOF = False Then fJmlSudahDibayar = IIf(IsNull(fRS2("JmlSudahDibayar").value), 0, fRS2("JmlSudahDibayar").value) Else fJmlSudahDibayar = 0
        If fRS2.EOF = False Then fSisaHutangPenjaminSebelumnya = IIf(IsNull(fRS2("SisaHutangPenjamin").value), 0, fRS2("SisaHutangPenjamin").value) Else fSisaHutangPenjaminSebelumnya = 0
        If fRS2.EOF = False Then fJmlPembebasanSebelumnya = IIf(IsNull(fRS2("JmlPembebasan").value), 0, fRS2("JmlPembebasan").value) Else fJmlPembebasanSebelumnya = 0
        Set fRS2 = Nothing
        fQuery2 = "select JmlBayar from StrukBuktiKasMasuk where NoBKM='" & fNoBKMClaimSebelumnya & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fJmlBayarSebelumnya = IIf(IsNull(fRS2("JmlBayar").value), 0, fRS2("JmlBayar").value) Else fJmlBayarSebelumnya = 0
        Set fRS2 = Nothing
        fQuery2 = "select sum(JmlHutangPenjamin) as JmlHutangPenjaminSum from RekapKomponenBiayaPelayananTM where NoBKM='" & fNoBKMClaimSebelumnya & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fTotalSisaHutangPenjaminTM = IIf(IsNull(fRS2("JmlHutangPenjaminSum").value), 0, fRS2("JmlHutangPenjaminSum").value) Else fTotalSisaHutangPenjaminTM = 0
        Set fRS2 = Nothing
        fQuery2 = "select sum(JmlHutangPenjamin) as JmlHutangPenjaminSum from RekapKomponenBiayaPelayananOA where NoBKM='" & fNoBKMClaimSebelumnya & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fTotalSisaHutangPenjaminOA = IIf(IsNull(fRS2("JmlHutangPenjaminSum").value), 0, fRS2("JmlHutangPenjaminSum").value) Else fTotalSisaHutangPenjaminOA = 0
        Set fRS2 = Nothing
        fQuery2 = "select sum(JmlHutangPenjamin) as JmlHutangPenjaminSum from RekapKomponenBiayaPelayananApotik where NoBKM='" & fNoBKMClaimSebelumnya & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fTotalSisaHutangPenjaminApotik = IIf(IsNull(fRS2("JmlHutangPenjaminSum").value), 0, fRS2("JmlHutangPenjaminSum").value) Else fTotalSisaHutangPenjaminApotik = 0
        Set fRS2 = Nothing
        fQuery2 = "insert into PembayaranClaimPenjaminPasien values('" & fNoBKM & "','" & fNoStruk & "'," & fJmlBayarSebelumnya & "," & fJmlSudahDibayar & " + " & fJmlBayar & "," & fJmlPembebasanSebelumnya & " + " & fJmlPembebasan & "," & fSisaHutangPenjamin & "," & fPembayaranKe & ",'" & fStatusPiutang & "')"
        Call msubRecFO(fRS2, fQuery2)
        If UCase(fStatusPiutang) = "SA" Then
            If fTotalSisaHutangPenjaminTM <> 0 Then
                fJmlBayarTM = (CDec(fTotalSisaHutangPenjaminTM) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlBayar)
                fJmlPembebasanTM = (CDec(fTotalSisaHutangPenjaminTM) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlPembebasan)
                fSisaHutangPenjaminTM = (CDec(fTotalSisaHutangPenjaminTM) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fSisaHutangPenjamin)
            End If
            If fTotalSisaHutangPenjaminOA <> 0 Then
                fJmlBayarOA = (CDec(fTotalSisaHutangPenjaminOA) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlBayar)
                fJmlPembebasanOA = (CDec(fTotalSisaHutangPenjaminOA) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlPembebasan)
                fSisaHutangPenjaminOA = (CDec(fTotalSisaHutangPenjaminOA) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fSisaHutangPenjamin)
            End If
        End If
        If fTotalSisaHutangPenjaminTM <> 0 And UCase(fStatusPiutang) = "TM" Then
            fJmlBayarTM = (CDec(fTotalSisaHutangPenjaminTM) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlBayar)
            fJmlPembebasanTM = (CDec(fTotalSisaHutangPenjaminTM) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlPembebasan)
            fSisaHutangPenjaminTM = (CDec(fTotalSisaHutangPenjaminTM) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fSisaHutangPenjamin)
        End If
        If fTotalSisaHutangPenjaminOA <> 0 And UCase(fStatusPiutang) = "OA" Then
            fJmlBayarOA = (CDec(fTotalSisaHutangPenjaminOA) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlBayar)
            fJmlPembebasanOA = (CDec(fTotalSisaHutangPenjaminOA) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlPembebasan)
            fSisaHutangPenjaminOA = (CDec(fTotalSisaHutangPenjaminOA) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fSisaHutangPenjamin)
        End If
        If fTotalSisaHutangPenjaminApotik <> 0 And UCase(fStatusPiutang) = "AP" Then
            fJmlBayarApotik = (CDec(fTotalSisaHutangPenjaminApotik) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlBayar)
            fJmlPembebasanApotik = (CDec(fTotalSisaHutangPenjaminApotik) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fJmlPembebasan)
            fSisaHutangPenjaminApotik = (CDec(fTotalSisaHutangPenjaminApotik) / CDec(fSisaHutangPenjaminSebelumnya)) * CDec(fSisaHutangPenjamin)
        End If
    End If
End Function

'Konversi dari SP: Add_PembayaranTagihanPasienApotikKredit
Public Function f_AddPembayaranTagihanPasienApotikKredit(fNoBKM As String, fNoStruk As String, fJmlBayar As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency)
    Dim fStatusPiutang As String
    Dim fTempPembayaranKe As Integer
    Dim fPembayaranKe As Integer
    Dim fMaksPembayaranKe As Integer
    Dim fJmlBayarSebelumnya As Currency
    Dim fJmlSudahDibayar As Currency
    Dim fNoBKMSebelumnya As String
    Dim fJmlBayarTM As Currency
    Dim fJmlPembebasanTM As Currency
    Dim fSisaTagihanTM As Currency
    Dim fJmlBayarOA As Currency
    Dim fJmlPembebasanOA As Currency
    Dim fSisaTagihanOA As Currency
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select max(PembayaranKe) as PembayaranKeMax from PembayaranTagihanPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTempPembayaranKe = IIf(IsNull(fRS("PembayaranKeMax").value), 0, fRS("PembayaranKeMax").value) Else fTempPembayaranKe = 0
    If fTempPembayaranKe = 0 Then
        fPembayaranKe = 1
    Else
        fPembayaranKe = fTempPembayaranKe + 1
        fMaksPembayaranKe = fTempPembayaranKe
    End If
    Set fRS = Nothing
    fQuery = "select NoBKM,JmlSudahDibayar from PembayaranTagihanPasien where NoStruk='" & fNoStruk & "' and PembayaranKe=" & fMaksPembayaranKe & ""
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fNoBKMSebelumnya = IIf(IsNull(fRS("NoBKM").value), "", fRS("NoBKM").value) Else fNoBKMSebelumnya = ""
    If fRS.EOF = False Then fJmlSudahDibayar = IIf(IsNull(fRS("JmlSudahDibayar").value), 0, fRS("JmlSudahDibayar").value) Else fJmlSudahDibayar = 0
    Set fRS = Nothing
    fQuery = "select JmlBayar from StrukBuktiKasMasuk where NoBKM='" & fNoBKMSebelumnya & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fJmlBayarSebelumnya = IIf(IsNull(fRS("JmlBayar").value), 0, fRS("JmlBayar").value) Else fJmlBayarSebelumnya = 0
    Set fRS = Nothing
    fQuery = "insert into PembayaranTagihanPasien values('" & fNoBKM & "','" & fNoStruk & "'," & fJmlBayarSebelumnya & "," & fJmlSudahDibayar & " + " & fJmlBayar & "," & fJmlPembebasan & "," & fSisaTagihan & "," & fPembayaranKe & ",'" & fStatusPiutang & "')"
    Call msubRecFO(fRS, fQuery)
    Call f_AddRekapKomponenBiayaPelayananApotikKredit(fNoBKM, fNoBKMSebelumnya, fNoStruk, fJmlBayar, fJmlPembebasan, fSisaTagihan)
End Function

'Konversi dari SP: Add_PembayaranTagihanPasienKredit
Public Function f_AddPembayaranTagihanPasienKredit(fNoBKM As String, fNoStruk As String, fJmlBayar As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency)
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String
    Dim fStatusPiutang As String
    Dim fMinPembayaranKe As Integer
    Dim fTempPembayaranKe As Integer
    Dim fPembayaranKe As Integer
    Dim fMaksPembayaranKe As Integer
    Dim fJmlBayarSebelumnya As Currency
    Dim fJmlSudahDibayar As Currency
    Dim fNoBKMSebelumnya As String
    Dim fJmlBayarTM As Currency
    Dim fJmlPembebasanTM As Currency
    Dim fSisaTagihanTM As Currency
    Dim fJmlBayarOA As Currency
    Dim fJmlPembebasanOA As Currency
    Dim fSisaTagihanOA As Currency
    Dim fTotalJmlSudahDibayarTM As Currency
    Dim fTotalJmlPembebasanTM As Currency
    Dim fTotalSisaTagihanTM As Currency
    Dim fTotalJmlSudahDibayarOA As Currency
    Dim fTotalJmlPembebasanOA As Currency
    Dim fTotalSisaTagihanOA As Currency
    Dim fSisaTagihanSebelumnya As Currency
    Dim fTglBKM As Date

    Set fRS = Nothing
    fQuery = "select TglBKM from StrukBuktiKasMasuk where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTglBKM = IIf(IsNull(fRS("TglBKM").value), "", fRS("TglBKM").value) Else fTglBKM = ""
    Set fRS = Nothing
    fQuery = "select JmlSudahDibayar,JmlPembebasan,SisaTagihan from TotalBiayaPelayananTM where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTotalJmlSudahDibayarTM = IIf(IsNull(fRS("JmlSudahDibayar").value), 0, fRS("JmlSudahDibayar").value) Else fTotalJmlSudahDibayarTM = 0
    If fRS.EOF = False Then fTotalJmlPembebasanTM = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value) Else fTotalJmlPembebasanTM = 0
    If fRS.EOF = False Then fTotalSisaTagihanTM = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value) Else fTotalSisaTagihanTM = 0
    Set fRS = Nothing
    fQuery = "select JmlSudahDibayar,JmlPembebasan,SisaTagihan from TotalBiayaPelayananOA where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTotalJmlSudahDibayarOA = IIf(IsNull(fRS("JmlSudahDibayar").value), 0, fRS("JmlSudahDibayar").value) Else fTotalJmlSudahDibayarOA = 0
    If fRS.EOF = False Then fTotalJmlPembebasanOA = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value) Else fTotalJmlPembebasanOA = 0
    If fRS.EOF = False Then fTotalSisaTagihanOA = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value) Else fTotalSisaTagihanOA = 0
    Set fRS = Nothing
    fQuery = "select min(PembayaranKe) as PembayaranKeMin from PembayaranTagihanPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fMinPembayaranKe = IIf(IsNull(fRS("PembayaranKeMin").value), 0, fRS("PembayaranKeMin").value) Else fMinPembayaranKe = 0
    Set fRS = Nothing
    fQuery = "select StatusPiutang from PembayaranTagihanPasien where NoStruk='" & fNoStruk & "' and PembayaranKe=" & fMinPembayaranKe & ""
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fStatusPiutang = IIf(IsNull(fRS("StatusPiutang").value), "SA", fRS("StatusPiutang").value) Else fStatusPiutang = "SA"
    Set fRS = Nothing
    fQuery = "select max(PembayaranKe) as PembayaranKeMax from PembayaranTagihanPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTempPembayaranKe = IIf(IsNull(fRS("PembayaranKeMax").value), 0, fRS("PembayaranKeMax").value) Else fTempPembayaranKe = 0
    If fTempPembayaranKe = 0 Then
        fPembayaranKe = 1
    Else
        fPembayaranKe = fTempPembayaranKe + 1
        fMaksPembayaranKe = fTempPembayaranKe
    End If
    Set fRS = Nothing
    fQuery = "select NoBKM,JmlSudahDibayar,SisaTagihan from PembayaranTagihanPasien where NoStruk='" & fNoStruk & "' and PembayaranKe=" & fMaksPembayaranKe & ""
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fNoBKMSebelumnya = IIf(IsNull(fRS("NoBKM").value), "", fRS("NoBKM").value) Else fNoBKMSebelumnya = ""
    If fRS.EOF = False Then fJmlSudahDibayar = IIf(IsNull(fRS("JmlSudahDibayar").value), 0, fRS("JmlSudahDibayar").value) Else fJmlSudahDibayar = 0
    If fRS.EOF = False Then fSisaTagihanSebelumnya = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value) Else fSisaTagihanSebelumnya = 0
    Set fRS = Nothing
    fQuery = "select JmlBayar from StrukBuktiKasMasuk where NoBKM='" & fNoBKMSebelumnya & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fJmlBayarSebelumnya = IIf(IsNull(fRS("JmlBayar").value), 0, fRS("JmlBayar").value) Else fJmlBayarSebelumnya = 0
    If fStatusPiutang = "SA" Then
        fJmlBayarTM = (CDec(fTotalSisaTagihanTM) / CDec(fSisaTagihanSebelumnya)) * CDec(fJmlBayar)
        fJmlPembebasanTM = (CDec(fTotalSisaTagihanTM) / CDec(fSisaTagihanSebelumnya)) * CDec(fJmlPembebasan)
        fSisaTagihanTM = (CDec(fTotalSisaTagihanTM) / CDec(fSisaTagihanSebelumnya)) * CDec(fSisaTagihan)
        fJmlBayarOA = (CDec(fTotalSisaTagihanOA) / CDec(fSisaTagihanSebelumnya)) * CDec(fJmlBayar)
        fJmlPembebasanOA = (CDec(fTotalSisaTagihanOA) / CDec(fSisaTagihanSebelumnya)) * CDec(fJmlPembebasan)
        fSisaTagihanOA = (CDec(fTotalSisaTagihanOA) / CDec(fSisaTagihanSebelumnya)) * CDec(fSisaTagihan)
        Set fRS = Nothing
        fQuery = "update TotalBiayaPelayananTM set JmlSudahDibayar=" & msubKonversiKomaTitik(CStr(fJmlBayarTM)) & " + " & msubKonversiKomaTitik(CStr(fTotalJmlSudahDibayarTM)) & ",JmlPembebasan=" & msubKonversiKomaTitik(CStr(fJmlPembebasanTM)) & " + " & msubKonversiKomaTitik(CStr(fTotalJmlPembebasanTM)) & ",SisaTagihan=" & msubKonversiKomaTitik(CStr(fSisaTagihanTM)) & " where NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "update TotalBiayaPelayananOA set JmlSudahDibayar=" & msubKonversiKomaTitik(CStr(fJmlBayarOA)) & " + " & msubKonversiKomaTitik(CStr(fTotalJmlSudahDibayarOA)) & ",JmlPembebasan=" & msubKonversiKomaTitik(CStr(fJmlPembebasanOA)) & " + " & msubKonversiKomaTitik(CStr(fTotalJmlPembebasanOA)) & ",SisaTagihan=" & msubKonversiKomaTitik(CStr(fSisaTagihanOA)) & " where NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "insert into PembayaranTagihanPasien values('" & fNoBKM & "','" & fNoStruk & "'," & msubKonversiKomaTitik(CStr(fJmlBayarSebelumnya)) & "," & msubKonversiKomaTitik(CStr(fJmlSudahDibayar)) & " + " & msubKonversiKomaTitik(CStr(fJmlBayar)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasan)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihan)) & "," & fPembayaranKe & ",'" & fStatusPiutang & "')"
        Call msubRecFO(fRS, fQuery)
        'Insert To table Histroy(TM,OA)
        fQuery = " update TotalBiayaPelayananTMHistory set JmlSudahDibayar=" & msubKonversiKomaTitik(CStr(fJmlBayarTM)) & " + " & msubKonversiKomaTitik(CStr(fTotalJmlSudahDibayarTM)) & ",JmlPembebasan=" & msubKonversiKomaTitik(CStr(fJmlPembebasanTM)) & " + " & msubKonversiKomaTitik(CStr(fTotalJmlPembebasanTM)) & ",SisaTagihan=" & msubKonversiKomaTitik(CStr(fSisaTagihanTM)) & " where NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS, fQuery)
        fQuery = "update TotalBiayaPelayananOAHistory set JmlSudahDibayar=" & msubKonversiKomaTitik(CStr(fJmlBayarOA)) & " + " & msubKonversiKomaTitik(CStr(fTotalJmlSudahDibayarOA)) & ",JmlPembebasan=" & msubKonversiKomaTitik(CStr(fJmlPembebasanOA)) & " + " & msubKonversiKomaTitik(CStr(fTotalJmlPembebasanOA)) & ",SisaTagihan=" & msubKonversiKomaTitik(CStr(fSisaTagihanOA)) & " where NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS, fQuery)
        'end add
    End If
    If fStatusPiutang = "TM" Then
        fJmlBayarTM = fJmlBayar
        fJmlPembebasanTM = fJmlPembebasan
        fSisaTagihanTM = fSisaTagihan
        Set fRS = Nothing
        fQuery = "update TotalBiayaPelayananTM set JmlSudahDibayar=" & msubKonversiKomaTitik(CStr(fJmlBayarTM)) & " + " & msubKonversiKomaTitik(CStr(fTotalJmlSudahDibayarTM)) & ",JmlPembebasan=" & msubKonversiKomaTitik(CStr(fJmlPembebasanTM)) & " + " & msubKonversiKomaTitik(CStr(fTotalJmlPembebasanTM)) & ",SisaTagihan=" & msubKonversiKomaTitik(CStr(fSisaTagihanTM)) & " where NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "insert into PembayaranTagihanPasien values('" & fNoBKM & "','" & fNoStruk & "'," & msubKonversiKomaTitik(CStr(fJmlBayarSebelumnya)) & "," & msubKonversiKomaTitik(CStr(fJmlSudahDibayar)) & " + " & msubKonversiKomaTitik(CStr(fJmlBayar)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasan)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihan)) & "," & fPembayaranKe & ",'" & fStatusPiutang & "')"
        Call msubRecFO(fRS, fQuery)
        fQuery = " update TotalBiayaPelayananTMHistory set JmlSudahDibayar=" & msubKonversiKomaTitik(CStr(fJmlBayarTM)) & " + " & msubKonversiKomaTitik(CStr(fTotalJmlSudahDibayarTM)) & ",JmlPembebasan=" & msubKonversiKomaTitik(CStr(fJmlPembebasanTM)) & " + " & msubKonversiKomaTitik(CStr(fTotalJmlPembebasanTM)) & ",SisaTagihan=" & msubKonversiKomaTitik(CStr(fSisaTagihanTM)) & " where NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS, fQuery)
    End If
    If fStatusPiutang = "OA" Then
        fJmlBayarOA = fJmlBayar
        fJmlPembebasanOA = fJmlPembebasan
        fSisaTagihanOA = fSisaTagihan
        Set fRS = Nothing
        fQuery = "update TotalBiayaPelayananOA set JmlSudahDibayar=" & msubKonversiKomaTitik(CStr(fJmlBayarOA)) & " + " & msubKonversiKomaTitik(CStr(fTotalJmlSudahDibayarOA)) & ",JmlPembebasan=" & msubKonversiKomaTitik(CStr(fJmlPembebasanOA)) & " + " & msubKonversiKomaTitik(CStr(fTotalJmlPembebasanOA)) & ",SisaTagihan=" & msubKonversiKomaTitik(CStr(fSisaTagihanOA)) & " where NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS, fQuery)
        Set fRS = Nothing
        fQuery = "insert into PembayaranTagihanPasien values('" & fNoBKM & "','" & fNoStruk & "'," & msubKonversiKomaTitik(CStr(fJmlBayarSebelumnya)) & "," & msubKonversiKomaTitik(CStr(fJmlSudahDibayar)) & " + " & msubKonversiKomaTitik(CStr(fJmlBayar)) & "," & msubKonversiKomaTitik(CStr(fJmlPembebasan)) & "," & msubKonversiKomaTitik(CStr(fSisaTagihan)) & "," & fPembayaranKe & ",'" & fStatusPiutang & "')"
        Call msubRecFO(fRS, fQuery)
        fQuery = "update TotalBiayaPelayananOAHistory set JmlSudahDibayar=" & msubKonversiKomaTitik(CStr(fJmlBayarOA)) & " + " & msubKonversiKomaTitik(CStr(fTotalJmlSudahDibayarOA)) & ",JmlPembebasan=" & msubKonversiKomaTitik(CStr(fJmlPembebasanOA)) & " + " & msubKonversiKomaTitik(CStr(fTotalJmlPembebasanOA)) & ",SisaTagihan=" & msubKonversiKomaTitik(CStr(fSisaTagihanOA)) & " where NoStruk='" & fNoStruk & "'"
        Call msubRecFO(fRS, fQuery)
    End If
End Function

'Konversi dari SP: Add_PostingJurnalTransaksiPelayananPasien
Public Function f_AddPostingJurnalTransaksiPelayananPasien(fNoPosting As String, fNoBuktiTransaksi As String, fTglBuktiTransaksi As Date, fKdJenisJurnal As String, fKdRekeningImpact As String, fIdPenjamin As String, fJenisTransaksi As String)
    'fNoBuktiTransaksi: NoBKM,NoBKK,NoStruk atau NoKuitansi/Bukti Lainnya
    'fKdRekeningImpact: Apakah pendapatan/pengeluaran berakibat ke rekening KAS atau KAS BANK atau...
    'fJenisTransaksi: TM=Tindakan Medis; OA=Obat Alkes; AP=Apotik
    Dim fKdRekening As String
    Dim fSaldoNormal As String
    Dim fSaldoNormalImpact As String
    Dim fJmlBayarPerKomp As Currency
    Dim fJmlHutangPenjaminPerKomp As Currency
    Dim fJmlTanggunganRSPerKomp As Currency
    Dim fJmlPembebasanPerKomp As Currency
    Dim fJmlSisaTagihanPerKomp As Currency
    Dim fKdPelayananRS As String
    Dim fKdKomponen As String
    Dim fJmlPelayanan As Integer
    Dim fNamaPelayanan As String
    Dim fTotalBayarPerKomp As Currency
    Dim fTotalHutangPenjaminPerKomp As Currency
    Dim fTotalTanggunganRSPerKomp As Currency
    Dim fTotalPembebasanPerKomp As Currency
    Dim fTotalSisaTagihanPerKomp As Currency
    Dim fKdRekeningPenjamin As String
    Dim fKdRekeningTanggunganRS As String
    Dim fKdRekeningPembebasan As String
    Dim fKdRekeningSisaTagihan As String
    Dim fSaldoNormalPenjamin As String
    Dim fSaldoNormalTanggunganRS As String
    Dim fSaldoNormalPembebasan As String
    Dim fSaldoNormalSisaTagihan As String
    Dim fTempNoPosting As String
    Dim fNoStruk As String
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    'Transaksi Pelayanan Tindakan Medis
    If UCase(fJenisTransaksi) = "TM" Then
        fQuery = "select NoStruk,KdPelayananRS,KdKomponen,JmlPelayanan as JmlBarang,JmlBayar,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,SisaTagihan from RekapKomponenBiayaPelayananTM where NoBKM='" & fNoBuktiTransaksi & "' and KdKomponen<>'12'"
    End If
    'Transaksi Pelayanan Obat & Alkes Ruangan
    If UCase(fJenisTransaksi) = "OA" Then
        fQuery = "select NoStruk,KdPelayananRS,KdKomponen,JmlBarang,JmlBayar,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,SisaTagihan from RekapKomponenBiayaPelayananOA where NoBKM='" & fNoBuktiTransaksi & "' and KdKomponen='06'"
    End If
    'Transaksi Pelayanan Obat & Alkes Apotik
    If UCase(fJenisTransaksi) = "AP" Then
        fQuery = "select NoStruk,KdPelayananRS,KdKomponen,JmlBarang,JmlBayar,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,SisaTagihan from RekapKomponenBiayaPelayananApotik where NoBKM='" & fNoBuktiTransaksi & "' and KdKomponen='06'"
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fNoStruk = IIf(IsNull(fRS("NoStruk").value), "", fRS("NoStruk").value)
        fKdPelayananRS = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
        fJmlPelayanan = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang").value)
        fJmlBayarPerKomp = IIf(IsNull(fRS("JmlBayar").value), 0, fRS("JmlBayar").value)
        fJmlHutangPenjaminPerKomp = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
        fJmlTanggunganRSPerKomp = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
        fJmlPembebasanPerKomp = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
        fJmlSisaTagihanPerKomp = IIf(IsNull(fRS("SisaTagihan").value), 0, fRS("SisaTagihan").value)
        Set fRS2 = Nothing
        fQuery2 = "select KdRekening,SaldoNormal,NamaPelayanan from V_ConvertRekeningToPelayananRS where KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fKdRekening = IIf(IsNull(fRS2("KdRekening").value), "", fRS2("KdRekening").value) Else fKdRekening = ""
        If fRS2.EOF = False Then fSaldoNormal = IIf(IsNull(fRS2("SaldoNormal").value), "", fRS2("SaldoNormal").value) Else fSaldoNormal = ""
        If fRS2.EOF = False Then fNamaPelayanan = IIf(IsNull(fRS2("NamaPelayanan").value), "", fRS2("NamaPelayanan").value) Else fNamaPelayanan = ""
        Set fRS2 = Nothing
        fQuery2 = "select SaldoNormal from DaftarRekening where KdRekening='" & fKdRekeningImpact & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then fSaldoNormalImpact = IIf(IsNull(fRS2("SaldoNormal").value), "", fRS2("SaldoNormal").value) Else fSaldoNormalImpact = ""
        Set fRS2 = Nothing
        fQuery2 = "select NoPosting from JurnalTransaksi where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then
            Set fRS3 = Nothing
            fQuery3 = "insert into JurnalTransaksi values('" & fNoPosting & "','" & fNoBuktiTransaksi & "','" & Format(fTglBuktiTransaksi, "yyyy/MM/dd HH:mm:ss") & "','" & fKdJenisJurnal & "','" & fNamaPelayanan & "',null)"
            Call msubRecFO(fRS3, fQuery3)
        End If
        Set fRS2 = Nothing
        fQuery2 = "select NoPosting from PostingDataKasirPenerimaan where NoPosting='" & fNoPosting & "' and NoBKM='" & fNoBuktiTransaksi & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then
            Set fRS3 = Nothing
            fQuery3 = "insert into PostingDataKasirPenerimaan values('" & fNoPosting & "','" & fNoBuktiTransaksi & "','" & fNoStruk & "')"
            Call msubRecFO(fRS3, fQuery3)
        End If
        fTotalBayarPerKomp = fJmlPelayanan * fJmlBayarPerKomp
        fTotalHutangPenjaminPerKomp = fJmlPelayanan * fJmlHutangPenjaminPerKomp
        fTotalTanggunganRSPerKomp = fJmlPelayanan * fJmlTanggunganRSPerKomp
        fTotalPembebasanPerKomp = fJmlPelayanan * fJmlPembebasanPerKomp
        fTotalSisaTagihanPerKomp = fJmlPelayanan * fJmlSisaTagihanPerKomp
        If fTotalBayarPerKomp <> 0 Then
            Set fRS2 = Nothing
            fQuery2 = "select NoPosting from DetailJurnalTransaksi where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekening & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                If UCase(fSaldoNormal) = "D" Then
                    fQuery3 = "insert into DetailJurnalTransaksi values('" & fNoPosting & "','" & fNoBuktiTransaksi & "','" & fKdRekening & "'," & fTotalBayarPerKomp & ",0)"
                Else
                    fQuery3 = "insert into DetailJurnalTransaksi values('" & fNoPosting & "','" & fNoBuktiTransaksi & "','" & fKdRekening & "',0," & fTotalBayarPerKomp & ")"
                End If
            Else
                If UCase(fSaldoNormal) = "D" Then
                    fQuery3 = "update DetailJurnalTransaksi set JmlDebet=JmlDebet + " & fTotalBayarPerKomp & " where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekening & "'"
                Else
                    fQuery3 = "update DetailJurnalTransaksi set JmlKredit=JmlKredit + " & fTotalBayarPerKomp & " where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekening & "'"
                End If
            End If
            Set fRS3 = Nothing
            Call msubRecFO(fRS3, fQuery3)
            'impact rekening
            Set fRS2 = Nothing
            fQuery2 = "select NoPosting from DetailJurnalTransaksi where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekeningImpact & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                If UCase(fSaldoNormalImpact) = "D" Then
                    fQuery3 = "insert into DetailJurnalTransaksi values('" & fNoPosting & "','" & fNoBuktiTransaksi & "','" & fKdRekeningImpact & "'," & fTotalBayarPerKomp & ",0)"
                Else
                    fQuery3 = "insert into DetailJurnalTransaksi values('" & fNoPosting & "','" & fNoBuktiTransaksi & "','" & fKdRekeningImpact & "',0," & fTotalBayarPerKomp & ")"
                End If
            Else
                If UCase(fSaldoNormalImpact) = "D" Then
                    fQuery3 = "update DetailJurnalTransaksi set JmlDebet=JmlDebet + " & fTotalBayarPerKomp & " where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekeningImpact & "'"
                Else
                    fQuery3 = "update DetailJurnalTransaksi set JmlKredit=JmlKredit + " & fTotalBayarPerKomp & " where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekeningImpact & "'"
                End If
            End If
            Set fRS3 = Nothing
            Call msubRecFO(fRS3, fQuery3)
        End If
        If fTotalHutangPenjaminPerKomp <> 0 Then
            Set fRS2 = Nothing
            fQuery2 = "select KdRekening,SaldoNormal from V_ConvertPenjaminToRekening where IdPenjamin='" & fIdPenjamin & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = False Then fKdRekeningPenjamin = IIf(IsNull(fRS2("KdRekening").value), "", fRS2("KdRekening").value) Else fKdRekeningPenjamin = ""
            If fRS2.EOF = False Then fSaldoNormalPenjamin = IIf(IsNull(fRS2("SaldoNormal").value), "", fRS2("SaldoNormal").value) Else fSaldoNormalPenjamin = ""
            Set fRS2 = Nothing
            fQuery2 = "select NoPosting from DetailJurnalTransaksi where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekeningPenjamin & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                If UCase(fSaldoNormalPenjamin) = "D" Then
                    fQuery3 = "insert into DetailJurnalTransaksi values('" & fNoPosting & "','" & fNoBuktiTransaksi & "','" & fKdRekeningPenjamin & "'," & fTotalHutangPenjaminPerKomp & ",0)"
                Else
                    fQuery3 = "insert into DetailJurnalTransaksi values('" & fNoPosting & "','" & fNoBuktiTransaksi & "','" & fKdRekeningPenjamin & "',0," & fTotalHutangPenjaminPerKomp & ")"
                End If
            Else
                If UCase(fSaldoNormalPenjamin) = "D" Then
                    fQuery3 = "update DetailJurnalTransaksi set JmlDebet=JmlDebet + " & fTotalHutangPenjaminPerKomp & " where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekeningPenjamin & "'"
                Else
                    fQuery3 = "update DetailJurnalTransaksi set JmlKredit=JmlKredit + " & fTotalHutangPenjaminPerKomp & " where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekeningPenjamin & "'"
                End If
            End If
            Set fRS3 = Nothing
            Call msubRecFO(fRS3, fQuery3)
        End If
        If fTotalTanggunganRSPerKomp <> 0 Then
            Set fRS2 = Nothing
            fQuery2 = "select KdRekeningTanggunganRS,SaldoNormalTanggunganRS from V_SettingRekeningStandar"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = False Then fKdRekeningTanggunganRS = IIf(IsNull(fRS2("KdRekeningTanggunganRS").value), "", fRS2("KdRekeningTanggunganRS").value) Else fKdRekeningTanggunganRS = ""
            If fRS2.EOF = False Then fSaldoNormalTanggunganRS = IIf(IsNull(fRS2("SaldoNormalTanggunganRS").value), "", fRS2("SaldoNormalTanggunganRS").value) Else fSaldoNormalTanggunganRS = ""
            Set fRS2 = Nothing
            fQuery2 = "select NoPosting from DetailJurnalTransaksi where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekeningTanggunganRS & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                If UCase(fSaldoNormalTanggunganRS) = "D" Then
                    fQuery3 = "insert into DetailJurnalTransaksi values('" & fNoPosting & "','" & fNoBuktiTransaksi & "','" & fKdRekeningTanggunganRS & "'," & fTotalTanggunganRSPerKomp & ",0)"
                Else
                    fQuery3 = "insert into DetailJurnalTransaksi values('" & fNoPosting & "','" & fNoBuktiTransaksi & "','" & fKdRekeningTanggunganRS & "',0," & fTotalTanggunganRSPerKomp & ")"
                End If
            Else
                If UCase(fSaldoNormalTanggunganRS) = "D" Then
                    fQuery3 = "update DetailJurnalTransaksi set JmlDebet=JmlDebet + " & fTotalTanggunganRSPerKomp & " where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekeningTanggunganRS & "'"
                Else
                    fQuery3 = "update DetailJurnalTransaksi set JmlKredit=JmlKredit + " & fTotalTanggunganRSPerKomp & " where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekeningTanggunganRS & "'"
                End If
            End If
            Set fRS3 = Nothing
            Call msubRecFO(fRS3, fQuery3)
        End If
        If fTotalPembebasanPerKomp <> 0 Then
            Set fRS2 = Nothing
            fQuery2 = "select KdRekeningPembebasan,SaldoNormalPembebasan from V_SettingRekeningStandar"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = False Then fKdRekeningPembebasan = IIf(IsNull(fRS2("KdRekeningPembebasan").value), "", fRS2("KdRekeningPembebasan").value) Else fKdRekeningPembebasan = ""
            If fRS2.EOF = False Then fSaldoNormalPembebasan = IIf(IsNull(fRS2("SaldoNormalPembebasan").value), "", fRS2("SaldoNormalPembebasan").value) Else fSaldoNormalPembebasan = ""
            Set fRS2 = Nothing
            fQuery2 = "select NoPosting from DetailJurnalTransaksi where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekeningPembebasan & "'"
            Call msubRecFO(fRS2, fQuery2)
            If ffrowcount = 0 Then
                If UCase(fSaldoNormalPembebasan) = "D" Then
                    fQuery3 = "insert into DetailJurnalTransaksi values('" & fNoPosting & "','" & fNoBuktiTransaksi & "','" & fKdRekeningPembebasan & "'," & fTotalPembebasanPerKomp & ",0)"
                Else
                    fQuery3 = "insert into DetailJurnalTransaksi values('" & fNoPosting & "','" & fNoBuktiTransaksi & "','" & fKdRekeningPembebasan & "',0," & fTotalPembebasanPerKomp & ")"
                End If
            Else
                If UCase(fSaldoNormalPembebasan) = "D" Then
                    fQuery3 = "update DetailJurnalTransaksi set JmlDebet=JmlDebet + " & fTotalPembebasanPerKomp & " where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekeningPembebasan & "'"
                Else
                    fQuery3 = "update DetailJurnalTransaksi set JmlKredit=JmlKredit + " & fTotalPembebasanPerKomp & " where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekeningPembebasan & "'"
                End If
            End If
            Set fRS3 = Nothing
            Call msubRecFO(fRS3, fQuery3)
        End If
        If fTotalSisaTagihanPerKomp <> 0 Then
            Set fRS2 = Nothing
            fQuery2 = "select KdRekening,SaldoNormal from V_ConvertPenjaminToRekening where IdPenjamin='2222222222'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = False Then fKdRekeningSisaTagihan = IIf(IsNull(fRS2("KdRekening").value), "", fRS2("KdRekening").value) Else fKdRekeningSisaTagihan = ""
            If fRS2.EOF = False Then fSaldoNormalSisaTagihan = IIf(IsNull(fRS2("SaldoNormal").value), "", fRS2("SaldoNormal").value) Else fSaldoNormalSisaTagihan = ""
            Set fRS2 = Nothing
            fQuery2 = "select NoPosting from DetailJurnalTransaksi where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekeningSisaTagihan & "'"
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                If UCase(fSaldoNormalSisaTagihan) = "D" Then
                    fQuery3 = "insert into DetailJurnalTransaksi values('" & fNoPosting & "','" & fNoBuktiTransaksi & "','" & fKdRekeningSisaTagihan & "'," & fTotalSisaTagihanPerKomp & ",0)"
                Else
                    fQuery3 = "insert into DetailJurnalTransaksi values('" & fNoPosting & "','" & fNoBuktiTransaksi & "','" & fKdRekeningSisaTagihan & "',0," & fTotalSisaTagihanPerKomp & ")"
                End If
            Else
                If UCase(fSaldoNormalSisaTagihan) = "D" Then
                    fQuery3 = "update DetailJurnalTransaksi set JmlDebet=JmlDebet + " & fTotalSisaTagihanPerKomp & " where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekeningSisaTagihan & "'"
                Else
                    fQuery3 = "update DetailJurnalTransaksi set JmlKredit=JmlKredit + " & fTotalSisaTagihanPerKomp & " where NoPosting='" & fNoPosting & "' and NoBuktiTransaksi='" & fNoBuktiTransaksi & "' and KdRekening='" & fKdRekeningSisaTagihan & "'"
                End If
            End If
            Set fRS3 = Nothing
            Call msubRecFO(fRS3, fQuery3)
        End If
        fRS.MoveNext
    Wend
End Function

'Konversi dari SP: Add_PeriksaDiagnosa
Public Function f_AddPeriksaDiagnosa(fNoPendaftaran As String, fNoCM As String, fKdSubInstalasi As String, fKdRuangan As String, fIdDokter As String, fTglPeriksa As Date, fKdDiagnosa As String, fKdJenisDiagnosa As String, fIdUser As String)
    Dim fStatusKasus As String
    Dim fJmlKdDiagnosa As Integer
    Dim fKdKelompokPasien As String
    Dim fKdRujukanAsal As String
    Dim fStatusPasien As String
    Dim fKdKelas As String
    Dim fKdKondisiPulang As String
    Dim fNoPakai As Variant
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select NoPakai from PemakaianKamar where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglMasuk<'" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fNoPakai = fRS("NoPakai").value Else fNoPakai = Null
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "'," & fNoPakai & ",'" & fKdRuangan & "','1') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fStatusPasien = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "'," & fNoPakai & ",'" & fKdRuangan & "','2') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdRujukanAsal = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "'," & fNoPakai & ",'" & fKdRuangan & "','4') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdKelas = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select KdKelompokPasien from PasienDaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value) Else fKdKelompokPasien = "01"
    Set fRS = Nothing
    fQuery = "select count(*) as JmlDiagnosa from PeriksaDiagnosa where KdRuangan='" & fKdRuangan & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdDiagnosa='" & fKdDiagnosa & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fJmlKdDiagnosa = IIf(IsNull(fRS("JmlDiagnosa").value), 0, fRS("JmlDiagnosa").value) Else fJmlKdDiagnosa = 0
    If fJmlKdDiagnosa = 0 Then
        fStatusKasus = "Baru"
    Else
        fStatusKasus = "Lama"
    End If
    Set fRS = Nothing
    fQuery = "insert into PeriksaDiagnosa values('" & fNoPendaftaran & "','" & fNoCM & "','" & fKdDiagnosa & "','" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "','" & fKdJenisDiagnosa & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fIdDokter & "','" & fStatusKasus & "')"
    Call msubRecFO(fRS, fQuery)
    Call f_AMDataMorbiditasPasien(fNoPendaftaran, fNoCM, fKdRuangan, fKdSubInstalasi, fTglPeriksa, fKdDiagnosa, fStatusKasus, "A")
    Call f_AMDataDiagnosaPasienPH(fNoCM, fKdRuangan, fKdKelompokPasien, fTglPeriksa, fKdJenisDiagnosa, fKdDiagnosa, fStatusKasus, "A")
    Set fRS = Nothing
    fQuery = "select KdKondisiPulang from PasienPulang where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        Call f_UpdateDataMorbiditasPasienRI(fNoCM, fKdKondisiPulang, fNoPendaftaran)
    End If
End Function

'Konversi dari SP: Delete_Diagnosa
Public Function f_DeleteDiagnosa(fNoPendaftaran As String, fKdRuangan As String, fKdDiagnosa As String, fTglPeriksa As Date, fKdSubInstalasi As String, fStatusKasus As String, fNoCM As String, fIdUser As String)
    Dim fKdJenisDiagnosa As String
    Dim fKdKelompokPasien As String
    Dim fKdRujukanAsal As String
    Dim fStatusPasien As String
    Dim fKdKelas As String
    Dim fNoPakai As Variant
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select NoPakai from PemakaianKamar where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglMasuk<'" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fNoPakai = fRS("NoPakai").value Else fNoPakai = Null
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "'," & fNoPakai & ",'" & fKdRuangan & "','1') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fStatusPasien = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "'," & fNoPakai & ",'" & fKdRuangan & "','2') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdRujukanAsal = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeStatusRujukanAsalSubInstalasiKelasPasien('" & fNoPendaftaran & "'," & fNoPakai & ",'" & fKdRuangan & "','4') as DataReturnValue"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdKelas = fRS("DataReturnValue").value
    Set fRS = Nothing
    fQuery = "select KdKelompokPasien from PasienDaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value) Else fKdKelompokPasien = "01"
    Set fRS = Nothing
    fQuery = "select KdJenisDiagnosa from PeriksaDiagnosa where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdDiagnosa='" & fKdDiagnosa & "' and TglPeriksa='" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisDiagnosa = IIf(IsNull(fRS("KdJenisDiagnosa").value), "", fRS("KdJenisDiagnosa").value) Else fKdJenisDiagnosa = ""
    Set fRS = Nothing
    fQuery = "delete from PeriksaDiagnosa where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdDiagnosa='" & fKdDiagnosa & "' and TglPeriksa='" & Format(fTglPeriksa, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    Call f_AMDataMorbiditasPasien(fNoPendaftaran, fNoCM, fKdRuangan, fKdSubInstalasi, fTglPeriksa, fKdDiagnosa, fStatusKasus, "M")
    Call f_AMDataDiagnosaPasienPH(fNoCM, fKdRuangan, fKdKelompokPasien, fTglPeriksa, fKdJenisDiagnosa, fKdDiagnosa, fStatusKasus, "M")
End Function

'Konversi dari SP: Delete_BiayaPelayanan
Public Function f_DeleteBiayaPelayanan(fNoPendaftaran As String, fKdRuangan As String, fKdPelayananRS As String, fTglPelayanan As Date, fIdUser As String)
    Dim fIdPenjamin As String
    Dim fKdPaket As Variant
    Dim fTotalBiayaPaket As Currency
    Dim fTotalTanggunganPaket As Currency
    Dim fKdPaketL As Variant
    Dim fTarifKelasPenjaminL As Currency
    Dim fJmlHutangPenjaminL As Currency
    Dim fKdPelayananRSL As String
    Dim fTglPelayananL As Date
    Dim fNoCM As String
    Dim fKdSubInstalasi As String
    Dim fJmlPelayanan As Integer
    Dim fIdPegawai As Variant
    Dim fIdPegawai2 As Variant
    Dim fKdPelayananRSAdmin As String
    Dim fKdInstalasi As String
    Dim fTglPelayananAdm As Date
    Dim fKdKelas As String
    Dim fStatusCito As String
    Dim fKdJenisTarif As String
    Dim fTarifCito As Currency
    Dim fKdRuanganAsal As String
    Dim fNoLab_Rad As Variant
    Dim fJmlPelayananTemp As Integer
    Dim fKdKelasPenjamin As String
    Dim fKdKelompokPasien As String
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Call f_DeleteTempHargaKomponen(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan)
    Set fRS = Nothing
    fQuery = "select StatusAPBD,KdSubInstalasi,JmlPelayanan,IdPegawai,IdPegawai2,KdKelas,NoLab_Rad,TarifCito,StatusCito from BiayaPelayanan where NoStruk is null and NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fKdSubInstalasi = IIf(IsNull(fRS("KdSubInstalasi").value), "", fRS("KdSubInstalasi").value)
        fKdAsal = IIf(IsNull(fRS("StatusAPBD").value), "", fRS("StatusAPBD").value)
        fJmlPelayanan = IIf(IsNull(fRS("JmlPelayanan").value), 0, fRS("JmlPelayanan").value)
        fIdPegawai = fRS("IdPegawai").value
        fIdPegawai2 = fRS("IdPegawai2").value
        fKdKelas = IIf(IsNull(fRS("KdKelas").value), "", fRS("KdKelas").value)
        fNoLab_Rad = fRS("NoLab_Rad").value
        fTarifCito = IIf(IsNull(fRS("TarifCito").value), 0, fRS("TarifCito").value)
        fStatusCito = IIf(IsNull(fRS("StatusCito").value), 0, fRS("StatusCito").value)
    End If
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "'," & fNoLab_Rad & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = fRS("KdRuanganAsal").value
    Set fRS = Nothing
    fQuery = "select KdPelayananRSAdmin from MasterDataPendukung"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdPelayananRSAdmin = IIf(IsNull(fRS("KdPelayananRSAdmin").value), "001001", fRS("KdPelayananRSAdmin").value) Else fKdPelayananRSAdmin = "001001"
    Set fRS = Nothing
    fQuery = "select KdJenisTarif from v_JenisTarifPasien where NoPendaftaran=fNoPendaftaran"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisTarif = IIf(IsNull(fRS("KdJenisTarif").value), "01", fRS("KdJenisTarif").value) Else fKdJenisTarif = "01"
    Set fRS = Nothing
    fQuery = "select min(TglPelayanan) as TglPelayananAdmMin from BiayaPelayanan where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fTglPelayananAdm = IIf(IsNull(fRS("TglPelayananAdmMin").value), "", fRS("TglPelayananAdmMin").value) Else fTglPelayananAdm = ""
    If fTglPelayananAdm <> "" Then
        Set fRS = Nothing
        fQuery = "select JmlPelayanan from BiayaPelayanan where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null and TglPelayanan='" & Format(fTglPelayananAdm, "yyyy/MM/dd HH:mm:ss") & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fJmlPelayananTemp = IIf(IsNull(fRS("JmlPelayanan").value), 0, fRS("JmlPelayanan").value) Else fJmlPelayananTemp = 0
        If fJmlPelayananTemp <> 0 Then
            Set fRS2 = Nothing
            fQuery2 = "update BiayaPelayanan set JmlPelayanan=JmlPelayanan-" & fJmlPelayanan & " where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null and TglPelayanan='" & Format(fTglPelayananAdm, "yyyy/MM/dd HH:mm:ss") & "'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update DetailBiayaPelayanan set JmlPelayanan=JmlPelayanan-" & fJmlPelayanan & " where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null and TglPelayanan='" & Format(fTglPelayananAdm, "yyyy/MM/dd HH:mm:ss") & "'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponen set JmlPelayanan=JmlPelayanan-" & fJmlPelayanan & " where KdPelayananRS='" & fKdPelayananRSAdmin & "' and KdRuangan='" & fKdRuangan & "' and NoPendaftaran='" & fNoPendaftaran & "' and NoStruk is null and TglPelayanan='" & Format(fTglPelayananAdm, "yyyy/MM/dd HH:mm:ss") & "'"
            Call msubRecFO(fRS2, fQuery2)
            Call f_AddTempHargaKomponen(fNoPendaftaran, fKdRuangan, fTglPelayananAdm, fKdPelayananRSAdmin, fKdKelas, fKdJenisTarif, CDbl(fTarifCito), fJmlPelayanan, fStatusCito, CStr(fIdPegawai))
        End If
    End If
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelas,KdKelompokPasien from V_KelasTanggunganPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPenjamin = IIf(IsNull(fRS("IdPenjamin").value), "2222222222", fRS("IdPenjamin").value)
        fKdKelasPenjamin = IIf(IsNull(fRS("KdKelas").value), fKdKelas, fRS("KdKelas").value)
        fKdKelompokPasien = IIf(IsNull(fRS("KdKelompokPasien").value), "01", fRS("KdKelompokPasien").value)
    End If
    Set fRS = Nothing
    fQuery = "select distinct KdPaket from V_PaketPenjamin where KdPelayananRS='" & fKdPelayananRS & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdKelas='" & fKdKelasPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdPaket = IIf(IsNull(fRS("KdPaket").value), "", fRS("KdPaket").value)
    If fRS.EOF = True Then
        Set fRS2 = Nothing
        fQuery2 = "delete from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and NoStruk is null"
        Call msubRecFO(fRS2, fQuery2)
    Else
        Set fRS2 = Nothing
        fQuery2 = "delete from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and NoStruk is null"
        Call msubRecFO(fRS2, fQuery2)
        Set fRS2 = Nothing
        fQuery2 = "select sum(TarifKelasPenjamin) as TarifKelasPenjaminSum from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket=" & fKdPaket & " and NoStruk is null and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "')"
        Call msubRecFO(fRS2, fQuery2)
        If fRS.EOF = False Then fTotalBiayaPaket = IIf(IsNull(fRS("TarifKelasPenjaminSum").value), 0, fRS("TarifKelasPenjaminSum").value)
        Set fRS2 = Nothing
        fQuery2 = "select sum(JmlHutangPenjamin) as JmlHutangPenjaminSum from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket=" & fKdPaket & " and NoStruk is null and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "')"
        Call msubRecFO(fRS2, fQuery2)
        If fRS.EOF = False Then fTotalTanggunganPaket = IIf(IsNull(fRS("JmlHutangPenjaminSum").value), 0, fRS("JmlHutangPenjaminSum").value)
        Set fRS = Nothing
        fQuery = "select KdPaket,TglPelayanan from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket=" & fKdPaket & " and NoStruk is null and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "')"
        Call msubRecFO(fRS, fQuery)
        While fRS.EOF = False
            fKdPaketL = fRS("KdPaket").value
            fTglPelayananL = IIf(IsNull(fRS("TglPelayanan").value), "", fRS("TglPelayanan").value)
            Set fRS2 = Nothing
            fQuery2 = "select KdPelayananRS,TarifKelasPenjamin from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket=" & fKdPaket & " and TglPelayanan='" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "' and NoStruk is null"
            Call msubRecFO(fRS2, fQuery2)
            While fRS2.EOF = False
                fKdPelayananRSL = IIf(IsNull(fRS("KdPelayananRS").value), "", fRS("KdPelayananRS").value)
                fTarifKelasPenjaminL = IIf(IsNull(fRS("TarifKelasPenjamin").value), 0, fRS("TarifKelasPenjamin").value)
                If fTotalBiayaPaket = 0 Then
                    fJmlHutangPenjaminL = 0
                Else
                    fJmlHutangPenjaminL = (CDec(fTarifKelasPenjaminL) / CDec(fTotalBiayaPaket)) * CDec(fTotalTanggunganPaket)
                End If
                Set fRS3 = Nothing
                fQuery3 = "update DetailBiayaPelayanan set JmlHutangPenjamin=" & fJmlHutangPenjaminL & " where NoPendaftaran='" & fNoPendaftaran & "' and KdPaket=" & fKdPaket & " and TglPelayanan='" & Format(fTglPelayananL, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRSL & "' and NoStruk is null"
                Call msubRecFO(fRS3, fQuery3)
                fRS2.MoveNext
            Wend
            fRS.MoveNext
        Wend
    End If
    Set fRS3 = Nothing
    fQuery3 = "delete from PetugasPemeriksaPasienBP where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS3, fQuery3)
    Set fRS3 = Nothing
    fQuery3 = "delete from DetailBackupUpdatingBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS3, fQuery3)
    Set fRS3 = Nothing
    fQuery3 = "delete from BackupUpdatingBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS3, fQuery3)
    Set fRS3 = Nothing
    fQuery3 = "delete from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and NoStruk is null"
    Call msubRecFO(fRS3, fQuery3)
    Set fRS3 = Nothing
    fQuery3 = "delete from BiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and NoStruk is null"
    Call msubRecFO(fRS3, fQuery3)
End Function

'Konversi dari SP: Delete_PemakaianObatAlkes
Public Function f_DeletePemakaianObatAlkes(fKdBarang As String, fKdAsal As String, fKdRuangan As String, fSatuan As String, fJmlBrg As Double, fNoPendaftaran As String, fTglPelayanan As Date, fIdUser As String)
    'fSatuan: S (Standar), K (Kecil)
    Dim fJmlBrgTemp As Double
    Dim fJmlJualTerkecil As Double
    Dim fJmlTerkecil As Double
    Dim fJmlStokRu As Double
    Dim fJmlBrgTempRu As Double
    Dim fJmlStokTerkecilRu As Double
    Dim fJmlModBrgTemp As Double
    Dim fJmlDivBrgTemp As Double
    Dim fJmlStokRuNow As Double
    Dim fJmlStokBrgTempNow As Double
    Dim fNoResep As Variant
    Dim fTempNoResep As String
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    Set fRS = Nothing
    fQuery = "select NoResep from PemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fNoResep = IIf(IsNull(fRS("NoResep").value), Null, fRS("NoResep").value) Else fNoResep = Null
    If UCase(fSatuan) = "S" Then
        Set fRS = Nothing
        fQuery = "select JmlStok from StokRuangan where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fJmlStokRu = IIf(IsNull(fRS("JmlStok").value), 0, fRS("JmlStok").value) Else fJmlStokRu = 0
        fJmlBrgTemp = fJmlStokRu + fJmlBrg
        GoTo SimpanS
    Else
        Set fRS = Nothing
        fQuery = "select JmlTerkecil from MasterBarang where KdBarang='" & fKdBarang & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fJmlTerkecil = IIf(IsNull(fRS("JmlTerkecil").value), 0, fRS("JmlTerkecil").value) Else fJmlTerkecil = 0
        Set fRS = Nothing
        fQuery = "select JmlJualTerkecil from MasterBarang where KdBarang='" & fKdBarang & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fJmlJualTerkecil = IIf(IsNull(fRS("JmlJualTerkecil").value), 0, fRS("JmlJualTerkecil").value) Else fJmlJualTerkecil = 0
        Set fRS = Nothing
        fQuery = "select JmlBarangTemp from JmlBarangTemp where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fJmlBrgTempRu = IIf(IsNull(fRS("JmlBarangTemp").value), 0, fRS("JmlBarangTemp").value) Else fJmlBrgTempRu = 0
        Set fRS = Nothing
        fQuery = "select JmlStok from StokRuangan where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fJmlStokRu = IIf(IsNull(fRS("JmlStok").value), 0, fRS("JmlStok").value) Else fJmlStokRu = 0
        fJmlBrgTemp = (fJmlBrg * fJmlJualTerkecil) + fJmlBrgTempRu
        fJmlModBrgTemp = CInt(fJmlBrgTemp) Mod CInt(fJmlTerkecil)
        fJmlDivBrgTemp = fJmlBrgTemp / fJmlTerkecil
        fJmlStokRuNow = fJmlStokRu + fJmlDivBrgTemp
        fJmlStokBrgTempNow = fJmlModBrgTemp
        GoTo SimpanK
    End If
SimpanS:
    Call f_DeleteTempHargaKomponenObatAlkes(fNoPendaftaran, fKdBarang, fTglPelayanan, fKdRuangan, fKdAsal, fSatuan)
    Set fRS2 = Nothing
    fQuery2 = "delete from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "'"
    Call msubRecFO(fRS2, fQuery2)
    Set fRS2 = Nothing
    fQuery2 = "delete from PetugasPemeriksaPasienOA where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "'"
    Call msubRecFO(fRS2, fQuery2)
    Set fRS2 = Nothing
    fQuery2 = "delete from DetailPemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "'"
    Call msubRecFO(fRS2, fQuery2)
    Set fRS2 = Nothing
    fQuery2 = "delete from PemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "'"
    Call msubRecFO(fRS2, fQuery2)
    Set fRS2 = Nothing
    fQuery2 = "update StokRuangan set JmlStok=" & fJmlBrgTemp & " where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "'"
    Call msubRecFO(fRS2, fQuery2)
    GoTo Selesai
SimpanK:
    Call f_DeleteTempHargaKomponenObatAlkes(fNoPendaftaran, fKdBarang, fTglPelayanan, fKdRuangan, fKdAsal, fSatuan)
    Set fRS2 = Nothing
    fQuery2 = "delete from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "'"
    Call msubRecFO(fRS2, fQuery2)
    Set fRS2 = Nothing
    fQuery2 = "delete from PetugasPemeriksaPasienOA where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "'"
    Call msubRecFO(fRS2, fQuery2)
    Set fRS2 = Nothing
    fQuery2 = "delete from DetailPemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "'"
    Call msubRecFO(fRS2, fQuery2)
    Set fRS2 = Nothing
    fQuery2 = "delete from PemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "'"
    Call msubRecFO(fRS2, fQuery2)
    Set fRS2 = Nothing
    fQuery2 = "update StokRuangan set JmlStok=" & fJmlStokRuNow & " where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "'"
    Call msubRecFO(fRS2, fQuery2)
    Set fRS2 = Nothing
    fQuery2 = "update JmlBarangTemp set JmlBarangTemp=" & fJmlStokBrgTempNow & " where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "'"
    Call msubRecFO(fRS2, fQuery2)
    GoTo Selesai
Selesai:
    If fNoResep <> Null Then
        Set fRS = Nothing
        fQuery = "select NoResep from PemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and NoResep=" & fNoResep & ""
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            Set fRS2 = Nothing
            fQuery2 = "delete from ResepObat where NoResep=" & fNoResep & ""
            Call msubRecFO(fRS2, fQuery2)
        End If
    End If
End Function

'Konversi dari SP: Delete_ReturnPemakaianObatAlkes
Public Function f_DeleteReturnPemakaianObatAlkes(fNoRetur As String, fNoPendaftaran As String, fKdRuangan As String, fKdBarang As String, fKdAsal As String, fTglPelayanan As Date, fSatuan As String, fJmlRetur As Double, fIdUser As String)
    'fSatuan: S (Standar), K (Kecil)
    Dim fJmlBrgTemp As Double
    Dim fJmlStokRu As Double
    Dim fKdBarangTemp As String
    Dim fJmlBrgPA As Double
    Dim fJmlBrgNow As Double
    Dim ftempJmlService As Integer
    Dim ftempKdJenisObat As Variant
    Dim fKdRuanganAsal As String
    Dim fNoLab_Rad As Variant
    Dim fKdKomponen As String
    Dim fHarga As Currency
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fJmlPembebasan As Currency
    Dim fRS As New ADODB.recordset
    Dim fRS2 As New ADODB.recordset
    Dim fRS3 As New ADODB.recordset
    Dim fQuery As String
    Dim fQuery2 As String
    Dim fQuery3 As String

    If UCase(fSatuan) = "S" Then
        Set fRS = Nothing
        fQuery = "select JmlStok from StokRuangan where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = False Then fJmlStokRu = IIf(IsNull(fRS("JmlStok").value), 0, fRS("JmlStok").value) Else fJmlStokRu = 0
        fJmlBrgTemp = fJmlStokRu - fJmlRetur
        If (fJmlStokRu >= fJmlRetur) Then
            Set fRS = Nothing
            fQuery = "select NoLab_Rad from PemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
            Call msubRecFO(fRS, fQuery)
            If fRS.EOF = False Then fNoLab_Rad = IIf(IsNull(fRS("NoLab_Rad").value), Null, fRS("NoLab_Rad").value) Else fNoLab_Rad = Null
            Set fRS = Nothing
            fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "'," & fNoLab_Rad & ",'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','OA') as KdRuanganAsal"
            Call msubRecFO(fRS, fQuery)
            fKdRuanganAsal = fRS("KdRuanganAsal").value
            Set fRS = Nothing
            fQuery = "select KdJenisObat, JmlService from PemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "'"
            Call msubRecFO(fRS, fQuery)
            If fRS.EOF = False Then
                ftempKdJenisObat = IIf(IsNull(fRS("KdJenisObat").value), Null, fRS("KdJenisObat").value)
                ftempJmlService = IIf(IsNull(fRS("JmlService").value), 0, fRS("JmlService").value)
                Set fRS = Nothing
                fQuery = "delete from ReturnPemakaianAlkes where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "' and NoRetur='" & fNoRetur & "' and NoPendaftaran='" & fNoPendaftaran & "' and SatuanJml='" & fSatuan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
                Call msubRecFO(fRS, fQuery)
                Set fRS = Nothing
                fQuery = "update StokRuangan set JmlStok=" & fJmlBrgTemp & " where KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdRuangan='" & fKdRuangan & "'"
                Call msubRecFO(fRS, fQuery)
                Set fRS = Nothing
                fQuery = "select JmlBarang from PemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "'"
                Call msubRecFO(fRS, fQuery)
                If fRS.EOF = False Then fJmlBrgPA = IIf(IsNull(fRS("JmlBarang").value), 0, fRS("JmlBarang").value) Else fJmlBrgPA = 0
                fJmlBrgNow = fJmlBrgPA + fJmlRetur
                If ftempKdJenisObat = "01" Then
                    ftempJmlService = 1
                    Set fRS = Nothing
                    fQuery = "update PemakaianAlkes set JmlBarang=" & fJmlBrgNow & ", JmlService = " & ftempJmlService & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "'"
                    Call msubRecFO(fRS, fQuery)
                    Set fRS = Nothing
                    fQuery = "update DetailPemakaianAlkes set JmlBarang=" & fJmlBrgNow & ", JmlService = " & ftempJmlService & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "'"
                    Call msubRecFO(fRS, fQuery)
                    Set fRS = Nothing
                    fQuery = "update TempHargaKomponenObatAlkes set JmlBarang=" & fJmlBrgNow & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "'"
                    Call msubRecFO(fRS, fQuery)
                    Set fRS = Nothing
                    fQuery = "select KdKomponen,HargaSatuan,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdBarang='" & fKdBarang & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdRuangan='" & fKdRuangan & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuan & "'"
                    Call msubRecFO(fRS, fQuery)
                    While fRS.EOF = False
                        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
                        fHarga = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
                        fJmlHutangPenjamin = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
                        fJmlTanggunganRS = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
                        fJmlPembebasan = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
                        Call f_AMDataPelayananOAPasienPH(fNoPendaftaran, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdBarang, fKdAsal, fSatuan, fKdKomponen, fHarga, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, ftempJmlService, fJmlRetur, "A")
                        fRS.MoveNext
                    Wend
                Else
                    Set fRS = Nothing
                    fQuery = "update PemakaianAlkes set JmlBarang=" & fJmlBrgNow & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "'"
                    Call msubRecFO(fRS, fQuery)
                    Set fRS = Nothing
                    fQuery = "update DetailPemakaianAlkes set JmlBarang=" & fJmlBrgNow & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "'"
                    Call msubRecFO(fRS, fQuery)
                    Set fRS = Nothing
                    fQuery = "update TempHargaKomponenObatAlkes set JmlBarang=" & fJmlBrgNow & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuan & "' and KdKomponen not in('10','18')"
                    Call msubRecFO(fRS, fQuery)
                    Set fRS = Nothing
                    fQuery = "select KdKomponen,HargaSatuan,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdBarang='" & fKdBarang & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdRuangan='" & fKdRuangan & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuan & "'"
                    Call msubRecFO(fRS, fQuery)
                    While fRS.EOF = False
                        fKdKomponen = IIf(IsNull(fRS("KdKomponen").value), "", fRS("KdKomponen").value)
                        fHarga = IIf(IsNull(fRS("HargaSatuan").value), 0, fRS("HargaSatuan").value)
                        fJmlHutangPenjamin = IIf(IsNull(fRS("JmlHutangPenjamin").value), 0, fRS("JmlHutangPenjamin").value)
                        fJmlTanggunganRS = IIf(IsNull(fRS("JmlTanggunganRS").value), 0, fRS("JmlTanggunganRS").value)
                        fJmlPembebasan = IIf(IsNull(fRS("JmlPembebasan").value), 0, fRS("JmlPembebasan").value)
                        Call f_AMDataPelayananOAPasienPH(fNoPendaftaran, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdBarang, fKdAsal, fSatuan, fKdKomponen, fHarga, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, ftempJmlService, fJmlRetur, "A")
                        fRS.MoveNext
                    Wend
                End If
            End If
        End If
    End If
End Function

'untuk updata antrian pasien
Public Function Update_AntrianPasienRegistrasi(iKdAntrian As Integer, sNoCM As String, sKdRuangan As String, sKdKelas As String, sKdKelompokPasien As String, sNoPendaftaran As String, sStatus As String) As Boolean
    On Error GoTo errSimpan
    Set adoCommand = New ADODB.Command
    Update_AntrianPasienRegistrasi = False
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdAntiran", adInteger, adParamInput, , iKdAntrian)
        .Parameters.Append .CreateParameter("TglAntrian", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("NoAntrian", adChar, adParamInput, 6, 0)
        .Parameters.Append .CreateParameter("StatusPasien", adInteger, adParamInput, , 0)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, sNoCM)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, sKdRuangan)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, sKdKelas)
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, sKdKelompokPasien)
        .Parameters.Append .CreateParameter("KdDokterOrder", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("NoAntrianOrder", adSmallInt, adParamInput, , Null)
        .Parameters.Append .CreateParameter("KeteranganOrder", adVarChar, adParamInput, 100, Null)
        .Parameters.Append .CreateParameter("NoReturOrder", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , Null)

        .Parameters.Append .CreateParameter("NoLoketCounter", adTinyInt, adParamInput, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, sNoPendaftaran)

        .Parameters.Append .CreateParameter("Status", adVarChar, adParamInput, 50, sStatus)

        .ActiveConnection = dbConn
        .CommandText = "Update_AntrianPasienRegistrasi"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data", vbCritical, "Validasi"
        Else
            Update_AntrianPasienRegistrasi = True
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set adoCommand = Nothing
    End With
    Exit Function
errSimpan:
    Call deleteADOCommandParameters(dbcmd)
    Set adoCommand = Nothing
    Call msubPesanError
End Function
Public Function sp_CetakLaporan(f_NamaCR As String, f_Nostruk As String, f_KdRuangan As String, F_IdUser As String) As Boolean
    On Error GoTo errLoad

    sp_CetakLaporan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoCetak", adVarChar, adParamInput, 15, Null)
        .Parameters.Append .CreateParameter("TglCetak", adDate, adParamInput, , Format(Now, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("ObjectCetak", adVarChar, adParamInput, 150, f_NamaCR)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_Nostruk)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, F_IdUser)
        .Parameters.Append .CreateParameter("OutputNoCetak", adVarChar, adParamOutput, 15, Null)
        .Parameters.Append .CreateParameter("OutputCetakKe", adSmallInt, adParamOutput, , Null)
        
        .ActiveConnection = dbConn
        .CommandText = "Add_CetakLaporan"
        .CommandType = adCmdStoredProc
        .Execute
        If .Parameters("RETURN_VALUE").value <> 0 Then
            MsgBox "Ada Kesalahan dalam proses", vbCritical, "Validasi"
            sp_CetakLaporan = False
        Else
            If Not IsNull(.Parameters("OutputNoCetak").value) Then strCetakKendaliLaporan = .Parameters("OutputNoCetak").value
            If Not IsNull(.Parameters("OutputCetakke").value) Then intCetakKe = .Parameters("OutputCetakke").value
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    sp_CetakLaporan = False
    Call msubPesanError("-sp_CetakLaporan")
End Function

