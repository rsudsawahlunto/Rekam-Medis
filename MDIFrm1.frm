VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIUtama 
   BackColor       =   &H8000000C&
   Caption         =   "Medifirst2000 - Rekam Medis (Medical Record)"
   ClientHeight    =   8565
   ClientLeft      =   2145
   ClientTop       =   2010
   ClientWidth     =   11400
   Icon            =   "MDIFrm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIFrm1.frx":0CCA
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrTempDW 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1200
      Top             =   120
   End
   Begin VB.Timer tmrTempDWKemarin 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1680
      Top             =   120
   End
   Begin MSComDlg.CommonDialog CDPrinter 
      Left            =   720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8310
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
            MinWidth        =   7832
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   9596
            MinWidth        =   9596
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "18/09/2019"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:23"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnberkas 
      Caption         =   "&Berkas"
      Begin VB.Menu mnuData 
         Caption         =   "Data"
         Begin VB.Menu MDaftarPasienRJRIdanIGD 
            Caption         =   "Daftar Pasien"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnucdp 
            Caption         =   "Cari Data Pasien"
            Shortcut        =   {F3}
         End
         Begin VB.Menu mndaftarreservasi 
            Caption         =   "Daftar Pasien Reservasi"
         End
         Begin VB.Menu mnasuransi 
            Caption         =   "Daftar Pasien Asuransi"
         End
         Begin VB.Menu mnusepdpl 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDaftarPasienLama 
            Caption         =   "Daftar Pasien Lama (RI)"
            Shortcut        =   +{F2}
         End
         Begin VB.Menu MDaftarDokumenRekamMedis 
            Caption         =   "Daftar Dokumen Rekam Medis"
            Shortcut        =   {F12}
         End
         Begin VB.Menu mnusepmdd2 
            Caption         =   "-"
         End
         Begin VB.Menu MInformasiKamar 
            Caption         =   "Informasi Kamar"
            Shortcut        =   {F4}
         End
         Begin VB.Menu LDaftarAntrianPasien 
            Caption         =   "-"
         End
         Begin VB.Menu MDaftarAntrianRegistrasi 
            Caption         =   "Daftar Antrian Registrasi"
         End
         Begin VB.Menu MDaftarAntrianPasien 
            Caption         =   "Daftar Antrian Pasien"
            Shortcut        =   {F6}
         End
         Begin VB.Menu MDaftarPasienKonsul 
            Caption         =   "Daftar Pasien Konsul"
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnusepcdp 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMDD 
            Caption         =   "Master Diagnosa"
         End
         Begin VB.Menu mnusepmdkp 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDiagnosaKeperawatan 
            Caption         =   "Master Diagnosa Keperawatan"
         End
         Begin VB.Menu mnuDetailDiagnosaKeperawatan 
            Caption         =   "Detail Diagnosa Keperawatan"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuTujuanNRencanaTindakan 
            Caption         =   "Tujuan && RencanaTindakan"
            Visible         =   0   'False
         End
         Begin VB.Menu mnusepdd 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMDU 
            Caption         =   "Master Sosial"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuMDKP 
            Caption         =   "Master Daftar Kontrol Pasien"
            Visible         =   0   'False
         End
         Begin VB.Menu MImunisasiJenisKontrasepsi 
            Caption         =   "Master Imunisasi && Jenis Kontrasepsi"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuMDP 
            Caption         =   "Master Pendukung"
            Visible         =   0   'False
         End
         Begin VB.Menu mnusepdp 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu MTindakanOperasi 
            Caption         =   "Master Tindakan Operasi"
            Visible         =   0   'False
         End
         Begin VB.Menu MRangeHasilKomponenScore 
            Caption         =   "Range Hasil Komponen Score"
            Visible         =   0   'False
         End
         Begin VB.Menu mnDokter 
            Caption         =   "Jadwal Praktek Dokter"
         End
         Begin VB.Menu mnHadir 
            Caption         =   "Master Kehadiran Dokter"
         End
         Begin VB.Menu MInformasiTarifPelayanan 
            Caption         =   "Informasi Tarif Pelayanan"
         End
         Begin VB.Menu mnuVerifikasi 
            Caption         =   "Verifikasi Data"
            Visible         =   0   'False
            Begin VB.Menu mnuVerDaftar 
               Caption         =   "Verifikasi Pendaftaran"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuVerJenisPasien 
               Caption         =   "Verifikasi Jenis Pasien"
            End
         End
         Begin VB.Menu mnuMonitoringData 
            Caption         =   "Monitoring Data"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnudiag 
         Caption         =   "-"
      End
      Begin VB.Menu mnureg 
         Caption         =   "Registrasi"
         Begin VB.Menu mnupb2 
            Caption         =   "Pasien Baru"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuPl 
            Caption         =   "Pasien Lama"
            Shortcut        =   ^L
         End
         Begin VB.Menu mnusepdw 
            Caption         =   "-"
         End
         Begin VB.Menu mnRegPenunjang 
            Caption         =   "Registrasi RJ && Penunjang"
         End
         Begin VB.Menu mnReservasi 
            Caption         =   "Reservasi Pendaftaran"
         End
         Begin VB.Menu mnDaftarAsuransi 
            Caption         =   "Pendaftaran Asuransi"
         End
      End
      Begin VB.Menu mnusepsp 
         Caption         =   "-"
      End
      Begin VB.Menu mSettingPrinter 
         Caption         =   "&Setting Printer"
         Shortcut        =   ^P
      End
      Begin VB.Menu MSettingPrinterBarcode 
         Caption         =   "Setting Printer &Barcode"
      End
      Begin VB.Menu mGantiKataKunci 
         Caption         =   "Ganti Kata Kunci"
         Shortcut        =   ^G
      End
      Begin VB.Menu mspace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnlogout 
         Caption         =   "Log Off"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnSelesai 
         Caption         =   "Keluar"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnINfo 
      Caption         =   "&Informasi"
      Begin VB.Menu mnINfoJadwal 
         Caption         =   "Informasi Jadwal Praktek Dokter"
      End
      Begin VB.Menu mnuCekKepesertaanBPJS 
         Caption         =   "Cek Kepesertaan BPJS"
      End
   End
   Begin VB.Menu mnuinventory 
      Caption         =   "In&ventory "
      Begin VB.Menu mnuPB 
         Caption         =   "Pemesanan Barang"
      End
      Begin VB.Menu mnuPemakaianBahandanAlat 
         Caption         =   "Pemakaian Bahan dan Alat"
         Visible         =   0   'False
      End
      Begin VB.Menu batasinv 
         Caption         =   "-"
      End
      Begin VB.Menu MBarangMedis 
         Caption         =   "Barang Medis"
         Visible         =   0   'False
         Begin VB.Menu mnusb 
            Caption         =   "Stok Barang"
         End
         Begin VB.Menu MClosingStok 
            Caption         =   "Closing Stok"
            Begin VB.Menu MCetakFormulirStok 
               Caption         =   "Cetak Lembar Input"
            End
            Begin VB.Menu MStokOpname 
               Caption         =   "Input Stok Opname"
            End
         End
         Begin VB.Menu mnn 
            Caption         =   "-"
         End
         Begin VB.Menu MInformasiPemesananPenerimaanBarang 
            Caption         =   "Informasi Pemesanan && Penerimaan Barang"
         End
         Begin VB.Menu MInformasiPemakaianBarang 
            Caption         =   "Informasi Pemakaian Barang"
         End
         Begin VB.Menu MLaporanSaldoBarang 
            Caption         =   "Laporan Saldo Barang"
         End
      End
      Begin VB.Menu MBarangNonMedis 
         Caption         =   "Barang Non Medis"
         Begin VB.Menu MStokBarangNonMedis 
            Caption         =   "Stok Barang"
         End
         Begin VB.Menu MKondisiBarangNonMedis 
            Caption         =   "Kondisi Barang"
         End
         Begin VB.Menu mMutasiBarangNM 
            Caption         =   "Mutasi Barang"
         End
         Begin VB.Menu MClosingStokNonMedis 
            Caption         =   "Closing Stok"
            Begin VB.Menu MCetakFormulirStokNonMedis 
               Caption         =   "Cetak Lembar Input"
            End
            Begin VB.Menu MStokOpnameNonMedis 
               Caption         =   "Input Stok Opname"
            End
            Begin VB.Menu MNilaiPersediaanNonMedis 
               Caption         =   "Nilai Persediaan"
            End
         End
         Begin VB.Menu mnlln 
            Caption         =   "-"
         End
         Begin VB.Menu MInformasiPemesananPenerimaanBarangNonMedis 
            Caption         =   "Informasi Pemesanan && Penerimaan Barang"
         End
      End
   End
   Begin VB.Menu mnulaporan 
      Caption         =   "&Laporan"
      Begin VB.Menu MnBkRegister 
         Caption         =   "Buku Register"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBukuRegisterPasien 
         Caption         =   "Buku Register Masuk"
      End
      Begin VB.Menu mnuBukuRegisterPelayanan 
         Caption         =   "Buku Register Pelayanan"
      End
      Begin VB.Menu MnSensusH 
         Caption         =   "Sensus Harian"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBWilayahDJenis 
         Caption         =   "Data Kunjungan Rumah Sakit"
      End
      Begin VB.Menu mnuRekapKP 
         Caption         =   "Kunjungan Berdasarkan Status && Jenis Pasien"
      End
      Begin VB.Menu mnRkstatusKasusPenyakit 
         Caption         =   "Kunjungan Berdasarkan Status && Kasus Penyakit Pasien"
      End
      Begin VB.Menu MnStausBRujukan 
         Caption         =   "Kunjungan Berdasarkan Status && Rujukan Pasien"
      End
      Begin VB.Menu MndKodisiPulang_Status 
         Caption         =   "Kunjungan Berdasarkan Status && Kondisi Pulang Pasien"
      End
      Begin VB.Menu MsLaporanKelasdanStatus 
         Caption         =   "Kunjungan Berdasarkan Status && Kelas"
      End
      Begin VB.Menu MnJenisOperasi_Status 
         Caption         =   "Kunjungan Berdasarkan Status && Jenis Operasi Pasien"
      End
      Begin VB.Menu mnusepdmpw 
         Caption         =   "-"
      End
      Begin VB.Menu RPBD 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Diagnosa"
      End
      Begin VB.Menu RPBW 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Wilayah"
      End
      Begin VB.Menu RPBWD2 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Wilayah && Diagnosa"
      End
      Begin VB.Menu mnuBkecDstatus 
         Caption         =   "Rekapitulasi Berdasarkan Kode Wilayah"
      End
      Begin VB.Menu qqq 
         Caption         =   "-"
      End
      Begin VB.Menu JHRP 
         Caption         =   "Jumlah Hari Rawatan Pasien"
      End
      Begin VB.Menu JPBWJ 
         Caption         =   "Jumlah Pasien RI Berdasarkan Wilayah && Jenis"
      End
      Begin VB.Menu JPBWJK 
         Caption         =   "Jumlah Pasien RI Berdasarkan Wilayah && Jenis Kelamin"
      End
      Begin VB.Menu qqqqqqqqqq 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDaftarPasienRJ 
         Caption         =   "Daftar Pasien Rawat Jalan"
         Visible         =   0   'False
      End
      Begin VB.Menu MnDaftarPasienMeninggal 
         Caption         =   "Daftar Pasien &Meninggal"
         Visible         =   0   'False
      End
      Begin VB.Menu MIndexDiagnosaPasien 
         Caption         =   "Index Diagnosa Pasien"
         Visible         =   0   'False
      End
      Begin VB.Menu mnujpp 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnur10p 
         Caption         =   "Rekapitulasi 10 Besar Penyakit"
      End
      Begin VB.Menu mRK 
         Caption         =   "Rekapitulasi 10 Besar Kematian"
         Visible         =   0   'False
      End
      Begin VB.Menu mnump 
         Caption         =   "Data Surveilens Morbiditas Pasien"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusepr10bp 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRekapPasienRIperSMF 
         Caption         =   "Rekapitulasi Pasien Rawat Inap Per SMF"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuiprs 
         Caption         =   "Indikator Pelayanan Rumah Sakit"
         Visible         =   0   'False
      End
      Begin VB.Menu LRekapitulasiKamarRawatInap 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusepdmp 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu CMD_INADRG 
         Caption         =   "INADRG"
         Visible         =   0   'False
         Begin VB.Menu cmdKonversiKetxt 
            Caption         =   "Konversi Ke txt"
         End
         Begin VB.Menu cmdKonversidaritxt 
            Caption         =   "Konversi dari txt"
         End
         Begin VB.Menu cmdKonversiINADRG 
            Caption         =   "-"
         End
         Begin VB.Menu cmdLaporanDetailPasienINADRG 
            Caption         =   "Laporan Detail Pasien INADRG"
         End
         Begin VB.Menu cmdLaporanRekapPasienINADRG 
            Caption         =   "Laporan Rekap Pasien INADRG"
         End
         Begin VB.Menu cmdKonversiINADRG1 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu cmdLapRealINADRG 
            Caption         =   "Realisasi INADRG"
         End
         Begin VB.Menu cmdLapSuratPengesahanaINADRG 
            Caption         =   "Surat Pengesahan"
         End
      End
      Begin VB.Menu mnuka 
         Caption         =   "Kesimpulan Akhir Pelayanan"
         Visible         =   0   'False
      End
      Begin VB.Menu mnulaporanrl 
         Caption         =   "Rekapitulasi Laporan (RL) versi 4"
         Visible         =   0   'False
         Begin VB.Menu mnuRL1 
            Caption         =   "RL 1"
            Begin VB.Menu mnupelayananrawatinap 
               Caption         =   "RL 1 Hal.1"
            End
            Begin VB.Menu mnuPengunjungRumahSakit 
               Caption         =   "RL 1 Hal.2"
            End
            Begin VB.Menu mnuKunjunganRawatJalan 
               Caption         =   "RL 1.3. Kunjungan Rawat Jalan"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuKegiatanKebidananPerinatologi 
               Caption         =   "RL 1.4. Kegiatan Kebidanan Perinatologi"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuKegiatanPembedahan 
               Caption         =   "RL 1.5. Kegiatan Pembedahan"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuKesehatanJiwa 
               Caption         =   "RL 1.6. Kesehatan Jiwa"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuPelayananRawatDarurat 
               Caption         =   "RL 1.7. Pelayanan Rawat Darurat"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuKunjunganRumah 
               Caption         =   "RL 1.8. Kunjungan Rumah"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuKegiatanRadiologi 
               Caption         =   "RL 1.9. Kegiatan Radiologi"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuradiodiagnotik 
               Caption         =   "            A. RADIODIAGNOTIK"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuradiotherapi 
               Caption         =   "            B. RADIOTHERAPI"
               Visible         =   0   'False
            End
            Begin VB.Menu mnukedokterannuklir 
               Caption         =   "            C. KEDOKTERAN NUKLIR"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuimaging 
               Caption         =   "            D. IMAGING / PENCITRAAN"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuKegiatanPelayananKhusus 
               Caption         =   "RL 1.10. Kegiatan Pelayanan Khusus"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuPemeriksaanLaboratorium 
               Caption         =   "RL 1.11. Pemeriksaan Laboratorium"
               Visible         =   0   'False
            End
            Begin VB.Menu mnupatologiklinik 
               Caption         =   "            A. PATOLOGI KLINIK"
               Visible         =   0   'False
            End
            Begin VB.Menu mnupatologianatomi 
               Caption         =   "            B. PATOLOGI ANATOMI"
               Visible         =   0   'False
            End
            Begin VB.Menu mnutoksilogi 
               Caption         =   "            C. TOKSILOGI"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuKegiatanRumahSakit 
               Caption         =   "RL 1.12. Kegiatan Rumah Sakit"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuPengadaanObat 
               Caption         =   "RL 1 Hal.3"
            End
            Begin VB.Menu mnuPenulisandanPelayananResep 
               Caption         =   "            B. Penulisan dan Pelayanan Resep"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuPelayananRehabilitasiMedik 
               Caption         =   "RL 1.13. Pelayanan Rehabilitasi Medik"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuKegiatanKeluargaBerencana 
               Caption         =   "RL 1 Hal.4"
            End
            Begin VB.Menu mnuKegiatanPenyuluhanKesehatan 
               Caption         =   "RL 1.15. Kegiatan Penyuluhan Kesehatan"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuKegiatanKesehatanGigi_Mulut 
               Caption         =   "RL 1.16. Kegiatan Kesehatan Gigi & Mulut"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuPemantauaanDokter_TenagaKesehatanAsingLainnya 
               Caption         =   "RL 1.17. Pemantauaan Dokter & Tenaga Kesehatan Asing Lainnya"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuTransfusiDarah 
               Caption         =   "RL 1.18. Transfusi Darah"
               Visible         =   0   'False
            End
            Begin VB.Menu mnulatihan 
               Caption         =   "RL 1 Hal.5"
            End
            Begin VB.Menu mnuPembedahanMata 
               Caption         =   "RL 1.20. Pembedahan Mata"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuPenangananPenyalahgunaanNapza 
               Caption         =   "RL 1.21. Penanganan Penyalahgunaan Napza"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuKegiatanBayiTabung 
               Caption         =   "RL 1.22. Kegiatan Bayi Tabung"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuCaraPembayaran 
               Caption         =   "RL 1 Hal.6"
            End
            Begin VB.Menu mnuKegiatanRujukan 
               Caption         =   "RL 1.24. Kegiatan Rujukan"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuRL2 
            Caption         =   "RL 2"
            Begin VB.Menu mnuDataKeadaanMorbiditasPasienRawatInapRumahSakit 
               Caption         =   "2.a Data Keadaan Morbiditas Pasien Rawat Inap Rumah Sakit"
            End
            Begin VB.Menu mnuDataKeadaanMorbiditasRawatInapSurveilensTerpaduRumahSakit 
               Caption         =   "      2.a.1. Data Keadaan Morbiditas Rawat Inap Surveilens Terpadu Rumah Sakit"
            End
            Begin VB.Menu mnuDataKeadaanMorbiditasPasienRawatJalanRumahSakit 
               Caption         =   "2.b Data Keadaan Morbiditas Pasien Rawat Jalan Rumah Sakit"
            End
            Begin VB.Menu mnuDataKeadaanMorbiditasRawatJalanSurveilensTerpaduRumahSakit 
               Caption         =   "      2.b.1. Data Keadaan Morbiditas Rawat Jalan Surveilens Terpadu Rumah Sakit"
            End
            Begin VB.Menu mnuDataStatusImunisasi 
               Caption         =   "2.c Data Status Imunisasi"
            End
            Begin VB.Menu mnuDataIndividualMorbiditasPasienRawatInap1 
               Caption         =   "2,1 Data Individual Morbiditas Pasien Rawat Inap (Pasien Umum)"
            End
            Begin VB.Menu mnuDataIndividualMorbiditasPasienRawatInap2 
               Caption         =   "2,2 Data Individual Morbiditas Pasien Rawat Inap (PasienObstetri)"
            End
            Begin VB.Menu mnumnuDataIndividualMorbiditasPasienRawatInap3 
               Caption         =   "2,3 Data Individual Morbiditas Pasien Rawat Inap (Perinatal)"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuRL3 
            Caption         =   "RL 3"
            Begin VB.Menu mnuDataDasarRumahSakit 
               Caption         =   "3 Data Dasar Rumah Sakit"
            End
         End
         Begin VB.Menu mnuRL4 
            Caption         =   "RL 4 "
            Begin VB.Menu mnuJumlahKetenagaanKesehatanMenurutJenis 
               Caption         =   "4.a Jumlah Ketenagaan Kesehatan Menurut Jenis"
            End
            Begin VB.Menu mnuTenagaMedis 
               Caption         =   "      4.a.1. Tenaga Medis"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuTenagaKeperawatan 
               Caption         =   "      4.a.2. Tenaga Keperawatan"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuTenagaKefarmasian 
               Caption         =   "      4.a.3. Tenaga Kefarmasian"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuTenagaKesehatanMasyarakat 
               Caption         =   "      4.a.4. Tenaga Kesehatan Masyarakat"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuTenagaGizi 
               Caption         =   "      4.a.5. Tenaga Gizi"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuTenagaKeterapianFisik 
               Caption         =   "      4.a.6. Tenaga Keterapian Fisik"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuTenagaKeteknisianMedis 
               Caption         =   "      4.a.7. Tenaga Keteknisian Medis"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuRL5 
            Caption         =   "RL 5"
            Begin VB.Menu mnuDataPeralatanMedikRumahSakit 
               Caption         =   "5,1 Data Peralatan Medik Rumah Sakit"
            End
            Begin VB.Menu mnuDataKePenerimaan 
               Caption         =   "5,2 Data Ke Penerimaan"
            End
            Begin VB.Menu mnu59 
               Caption         =   "5,9 Data Kegiatan Kesehatan Lingkungan"
            End
         End
         Begin VB.Menu mnurl6 
            Caption         =   "RL 6"
            Begin VB.Menu mnuFormulirPelaporanInfeksiNosokmial 
               Caption         =   "Formulir Pelaporan Infeksi Nosokmial"
            End
         End
      End
      Begin VB.Menu mnulaporanrlnew 
         Caption         =   "Rekapitulasi Laporan (RL) Baru versi 5"
         Begin VB.Menu mnuRL1New 
            Caption         =   "RL 1"
            Begin VB.Menu mnurl1_1 
               Caption         =   "RL 1.1 Data Dasar Rumah Sakit"
            End
            Begin VB.Menu mnurl1_2 
               Caption         =   "RL 1.2 Indikator Pelayanan Rumah Sakit"
            End
            Begin VB.Menu mnurl1_3 
               Caption         =   "RL 1.3 Fasilitas Tempat Tidur"
            End
         End
         Begin VB.Menu mnuRL2New2 
            Caption         =   "RL 2 Ketenagaan"
         End
         Begin VB.Menu mnuRL3New2 
            Caption         =   "RL 3 Data Kegiatan Pelayanan RS"
            Begin VB.Menu mnurl3_1 
               Caption         =   "RL 3.1 Kegiatan Pelayanan RI"
            End
            Begin VB.Menu mnurl3_2 
               Caption         =   "RL 3.2 Kegiatan Pelayanan Rawat Darurat"
            End
            Begin VB.Menu mnurl3_3 
               Caption         =   "RL 3.3 Kegiatan Kesehatan Gigi Dan Mulut"
            End
            Begin VB.Menu mnurl3_4 
               Caption         =   "RL 3.4 Kegiatan Kebidanan"
            End
            Begin VB.Menu mnurl3_5 
               Caption         =   "RL 3.5 Kegiatan Perinatologi"
            End
            Begin VB.Menu mnurl3_6 
               Caption         =   "RL 3.6 Kegiatan Pembedahan"
            End
            Begin VB.Menu mnurl3_7 
               Caption         =   "RL 3.7 Kegiatan Radiologi"
            End
            Begin VB.Menu mnurl3_8 
               Caption         =   "RL 3.8 Pemeriksaan Laboratorium"
            End
            Begin VB.Menu mnurl3_9 
               Caption         =   "RL 3.9 Pelayanan Rehabilitasi Medik"
            End
            Begin VB.Menu mnuRL3_10 
               Caption         =   "RL 3.10 Kegiatan Pelayanan Khusus"
            End
            Begin VB.Menu mnurl3_11 
               Caption         =   "RL 3.11 Kegiatan Kesehatan Jiwa"
            End
            Begin VB.Menu mnuRL3_12 
               Caption         =   "RL 3.12 Kegiatan Keluarga Berencana"
            End
            Begin VB.Menu mnurl3_13 
               Caption         =   "RL 3.13 Pengadaan Obat, Penulisan & Pelayanan Resep"
            End
            Begin VB.Menu mnuRL3_14 
               Caption         =   "RL 3.14 Kegiatan Rujukan"
            End
            Begin VB.Menu mnurl3_15 
               Caption         =   "RL 3.15 Cara Bayar"
            End
         End
         Begin VB.Menu mnuRL4New 
            Caption         =   "RL 4 Morbiditas"
            Begin VB.Menu mnuRL4A 
               Caption         =   "RL 4A Data Keadaan Morbiditas Pasien RI"
            End
            Begin VB.Menu mnurl4apk 
               Caption         =   "RL 4A Penyebab Kecelakaan"
            End
            Begin VB.Menu mnuRL4B 
               Caption         =   "RL 4B Data Keadaan Morbiditas Pasien RJ"
            End
            Begin VB.Menu mnurl4bpk 
               Caption         =   "RL 4B Penyebab Kecelakaan"
            End
         End
         Begin VB.Menu mnuRL5New2 
            Caption         =   "RL 5 Data Bulanan"
            Begin VB.Menu mnuRL5_1 
               Caption         =   "RL 5.1 Pungunjung Rumah Sakit"
            End
            Begin VB.Menu mnurl5_2 
               Caption         =   "RL 5.2 Kunjungan Rawat Jalan"
            End
            Begin VB.Menu mnurl5_3 
               Caption         =   "RL 5.3 Daftar 10 Besar Penyakit RI"
            End
            Begin VB.Menu mnurl5_4 
               Caption         =   "RL 5.4 Daftar 10 Besar Penyakit RJ"
            End
         End
      End
      Begin VB.Menu mnuLapRL6 
         Caption         =   "Rekapitulasi Laporan (RL) Baru versi 6"
         Begin VB.Menu mnuRL1New2 
            Caption         =   "RL 1"
            Begin VB.Menu mnuRL1_2New 
               Caption         =   "RL 1.2 Indikator Pelayanan Rumah Sakit"
            End
            Begin VB.Menu mnuRL1_3New 
               Caption         =   "RL 1.3 Fasilitas Tempat Tidur"
            End
         End
         Begin VB.Menu mnuRL2New 
            Caption         =   "RL 2 Ketenagaan"
         End
         Begin VB.Menu mnuRL3New 
            Caption         =   "RL 3 Data Kegiatan Pelayanan RS"
            Begin VB.Menu mnuRL3_1New 
               Caption         =   "RL 3.1 Kegiatan Pelayanan RI"
            End
            Begin VB.Menu mnuRL3_2New 
               Caption         =   "RL 3.2 Kegiatan Pelayanan Rawat Darurat"
            End
            Begin VB.Menu mnuRL3_3New 
               Caption         =   "RL 3.3 Kegiatan Kesehatan Gigi Dan Mulut"
            End
            Begin VB.Menu mnuRL3_4New 
               Caption         =   "RL 3.4 Kegiatan Kebidanan"
            End
            Begin VB.Menu mnuRL3_5New 
               Caption         =   "RL 3.5 Kegiatan Perinatologi"
            End
            Begin VB.Menu mnuRL3_6New 
               Caption         =   "RL 3.6 Kegiatan Pembedahan"
            End
            Begin VB.Menu mnuRL3_7New 
               Caption         =   "RL 3.7 Kegiatan Radiologi"
            End
            Begin VB.Menu mnuRL3_8New 
               Caption         =   "RL 3.8 Pemeriksaan Laboratorium"
            End
            Begin VB.Menu mnuRL3_9New 
               Caption         =   "RL 3.9 Pelayanan Rehabilitasi Medik"
            End
            Begin VB.Menu mnuRL3_10New 
               Caption         =   "RL 3.10 Kegiatan Pelayanan Khusus"
            End
            Begin VB.Menu mnuRL3_11New 
               Caption         =   "RL 3.11 Kegiatan Kesehatan Jiwa"
            End
            Begin VB.Menu mnuRL3_12New 
               Caption         =   "RL 3.12 Kegiatan Keluarga Berencana"
            End
            Begin VB.Menu mnuRL3_13New 
               Caption         =   "RL 3.13 Pengadaan Obat, Penulisan & Pelayanan Resep"
            End
            Begin VB.Menu mnuRL3_14New 
               Caption         =   "RL 3.14 Kegiatan Rujukan"
            End
            Begin VB.Menu mnuRL3_15New 
               Caption         =   "RL 3.15 Cara Bayar"
            End
         End
         Begin VB.Menu mnuRL4New2 
            Caption         =   "RL 4 Morbiditas"
            Begin VB.Menu mnuRL4_ANew 
               Caption         =   "RL 4A Data Keadaan Morbiditas Pasien RI"
            End
            Begin VB.Menu mnuRL4APenyebabKecelakaan 
               Caption         =   "RL 4A Penyebab Kecelakaan"
            End
            Begin VB.Menu mnuRL4BNew2 
               Caption         =   "RL 4B Data Keadaan Morbiditas Pasien RJ"
            End
            Begin VB.Menu mnuRL4BNew 
               Caption         =   "RL 4B Penyebab Kecelakaan"
            End
         End
         Begin VB.Menu mnuRL5New 
            Caption         =   "RL 5 Data Bulanan"
            Begin VB.Menu mnuRL5_1New 
               Caption         =   "RL 5.1 Pungunjung Rumah Sakit"
            End
            Begin VB.Menu mnuRL5_2New 
               Caption         =   "RL 5.2 Kunjungan Rawat Jalan"
            End
            Begin VB.Menu mnuRL5_3New 
               Caption         =   "RL 5.3 Daftar 10 Besar Penyakit RI"
            End
            Begin VB.Menu mnuRL5_4New 
               Caption         =   "RL 5.4 Daftar 10 Besar Penyakit RJ"
            End
         End
      End
   End
   Begin VB.Menu mWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu MCascade 
         Caption         =   "&Cascade"
      End
   End
   Begin VB.Menu mbantuan 
      Caption         =   "Ban&tuan"
      Begin VB.Menu mTentang2 
         Caption         =   "Tentang Medifirst2000"
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "MDIUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sepuh As Boolean

Private Sub cmdKonversidaritxt_Click()
    frmLoadINADRG.Show
End Sub

Private Sub cmdKonversiKetxt_Click()
    frmConvertINADRG.Show
End Sub

Private Sub cmdLaporanDetailPasienINADRG_Click()
    frmLap20091201003.Show
End Sub

Private Sub cmdLaporanRekapPasienINADRG_Click()
    frmLap20091201004.Show
End Sub

Private Sub cmdLapRealINADRG_Click()
    frmINADRGBillingRealisasi.Show
End Sub

Private Sub cmdLapSuratPengesahanaINADRG_Click()
    frmINADRGLoadSuratPengesahan.Show
End Sub

Private Sub JHRP_Click()
'    strCetak = "LapJmlHariRawat"
'    With frmLapRKP_KPSK
'        .Caption = "Medifirst2000 - Jumlah Hari Rawatan Pasien"
'        .Show
'    End With
    
     strCetak = "LapJmlHariRawat"
    With frmLaporanTahunan
        .Caption = "Medifirst2000 - Jumlah Hari Rawatan Pasien"
        .Show
    End With
End Sub

Private Sub JPBWJ_Click()
    strCetak = "LapJmlWilayahJenis"
    With frmLapRKP_KPSK
        .Caption = "Medifirst2000 - Jumlah Pasien Berdasarkan Wilayah dan Jenis Pasien"
        .Show
    End With
End Sub

Private Sub JPBWJK_Click()
    strCetak = "LapJmlWilayahJenisKelamin"
    With frmLapRKP_KPSK
        .Caption = "Medifirst2000 - Jumlah Pasien Berdasarkan Wilayah dan Jenis Kelamin"
        .Show
    End With
End Sub

Private Sub MCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub MCetakFormulirStok_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmDaftarCetakInputStokOpname.Show
End Sub

Private Sub MCetakFormulirStokNonMedis_Click()
    mstrKdKelompokBarang = "01" 'non medis
    frmDaftarCetakInputStokOpnameNM.Show
End Sub

Private Sub MDaftarAntrianPasien_Click()
    frmDaftarAntrianPasien.Show
End Sub

Private Sub MDaftarAntrianRegistrasi_Click()
    frmDaftarAntrianRegistrasi.Show
End Sub

Private Sub MDaftarDokumenRekamMedis_Click()
    frmDaftarDokumenRekamMedisPasien.Show
End Sub

Private Sub MDaftarPasienKonsul_Click()
    frmDaftarPasienKonsul.Show
End Sub

Private Sub MDaftarPasienRJRIdanIGD_Click()
    frmDaftarPasienRJRIIGD.Show
End Sub

Private Sub MDIForm_Load()
    editpoli = False
    strSQL = "SELECT * FROM DataPegawai WHERE IdPegawai = '" & strIDPegawaiAktif & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    strNmPegawai = rs.Fields("NamaLengkap").value
    Set rs = Nothing
    StatusBar1.Panels(1).Text = "Nama User : " & strNmPegawai
    StatusBar1.Panels(2).Text = "Nama Ruangan : " & mstrNamaRuangan
    StatusBar1.Panels(5).Text = "Nama Komputer : " & strNamaHostLocal
    StatusBar1.Panels(6).Text = "Server : " & strServerName & " (" & strDatabaseName & ")"
    mnlogout.Caption = "Log Off..." & strNmPegawai

    strSQL = "select StatusAntrian from SettingDataUmum"
    Call msubRecFO(rs, strSQL)
    Dim coba As Long
    If Not rs.EOF Then
        If rs(0).value = "1" Then
             MDaftarAntrianRegistrasi.Visible = True
          Else
             MDaftarAntrianRegistrasi.Visible = False
        End If
    End If

    'add rhmt 2009/04/28'
    Dim strFileDAT As String
    Dim strFileDATKemarin As String
    Dim strTglKemarin As String

    strTglKemarin = Format(DateAdd("d", -1, Now), "yyyyMMdd")
    strFileDATKemarin = "tempDW" & strTglKemarin & ".dat"

    strFileDAT = "tempDW" & Format(Now, "yyyyMMdd") & ".dat"
    strFolderDAT = funcGetSpecialFolder(Me.hWnd, CSIDL_DOCUMENTS) & "\RekamMedikTempDW"
    If Not fso.FolderExists(strFolderDAT) Then
        fso.CreateFolder (strFolderDAT)
    End If
    strFolderDAT = strFolderDAT & "\" & mstrNamaRuangan
    If Not fso.FolderExists(strFolderDAT) Then
        fso.CreateFolder (strFolderDAT)
    End If
    strLokasiFileDATKemarin = strFolderDAT & "\" & strFileDATKemarin
    strLokasiFileDAT = strFolderDAT & "\" & strFileDAT
    If fso.FileExists(strLokasiFileDATKemarin) Then
        Call subOpenReadFile(strLokasiFileDATKemarin)
        Me.tmrTempDWKemarin.Enabled = True
    Else
        Call subOpenReadFile(strLokasiFileDAT)
        Me.tmrTempDW.Enabled = True
    End If
    
    ' untuk mendapatkan kode ruangan Rekam Medis dan setting global di sesuaikan dengan kode ruangan RekamMedis
    strSQL = "Select Value from SettingGlobal where prefix='KdRuanganRekamMedis'"
    Call msubRecFO(rs1, strSQL)
    If rs1.EOF = False Then
        strKdRuanganRekamMedis = rs1.Fields("Value").value
    End If

    ' untuk mendapatkan kode ruangan Registrasi RJ dan setting global di sesuaikan dengan kode ruangan Registrasi RJ
    strSQL = "Select Value from SettingGlobal where prefix='KdRuanganRegistrasiRJ'"
    Call msubRecFO(rs2, strSQL)
    If rs1.EOF = False Then
        strKdRuanganRegistrasiRJ = rs2.Fields("Value").value
    End If

    ' untuk mendapatkan kode ruangan Registrasi RI dan setting global di sesuaikan dengan kode ruangan Registrasi RI
    strSQL = "Select Value from SettingGlobal where prefix='KdRuanganRegistrasiRI'"
    Call msubRecFO(rs3, strSQL)
    If rs1.EOF = False Then
        strKdRuanganRegistrasiRI = rs3.Fields("Value").value
    End If

    
    
End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Exit Sub
    PopupMenu mnuData
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim q As String
    If sepuh = True Then
        q = MsgBox("Log Off user " & strNmPegawai & " ", vbQuestion + vbOKCancel, "Konfirmasi")
        If q = 2 Then
            Unload frmLogin
            Cancel = 1
        Else

        Close intFreeFile
        Cancel = 0
        frmLogin.Show
    End If
    sepuh = False
Else
    q = MsgBox("Tutup aplikasi ", vbQuestion + vbOKCancel, "Konfirmasi")
    If q = 2 Then
        Unload frmLogin
        Cancel = 1
    Else
        dTglLogout = Now
        Call subSp_HistoryLoginAplikasi("U")

    Close intFreeFile

    Cancel = 0
End If
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error GoTo errLoad

    Dim adoCommand As New ADODB.Command
    openConnection
    strQuery = "UPDATE Login SET Status = '0' " & _
    "WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
    adoCommand.ActiveConnection = dbConn
    adoCommand.CommandText = strQuery
    adoCommand.CommandType = adCmdText
    adoCommand.Execute

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub mGantiKataKunci_Click()
    frmLoginEditAccount.Show
End Sub

Private Sub MImunisasiJenisKontrasepsi_Click()
    frmImunisasiJenisKontrasepsi.Show
End Sub

Private Sub MIndexDiagnosaPasien_Click()
    frmPeriodeLaporanIndexDiagnosaPasien.Show
End Sub

Private Sub MInformasiKamar_Click()
    frminfokamar.Show
End Sub

Private Sub MInformasiPemakaianBarang_Click()
    frmDaftarPakaiAlkesKaryawan.Show
End Sub

Private Sub MInformasiPemesananPenerimaanBarang_Click()
    mstrKdKelompokBarang = "02"  'medis
    frmInfoPesanBarang.Show
End Sub

Private Sub MInformasiPemesananPenerimaanBarangNonMedis_Click()
    mstrKdKelompokBarang = "01"  'non medis
    frmInfoPesanBarangNM.Show
End Sub

Private Sub MInformasiTarifPelayanan_Click()
    frmInformasiTarifPelayanan.Show
End Sub

Private Sub MKomponenKlinis_Click()
    FrmKomponenKlinis.Show
End Sub

Private Sub MKomponenScore_Click()
    FrmKomponenScore.Show
End Sub

Private Sub MKondisiBarangNonMedis_Click()
    frmKondisiBarangNM.Show
End Sub

Private Sub MLaporanSaldoBarang_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmLaporanSaldoBarangMedis_v3.Show
End Sub

Private Sub MLaporanSaldoBarangNonMedis_Click()
    mstrKdKelompokBarang = "01"     'non medis
    frmLaporanSaldoBarangNM_v3.Show
End Sub

Private Sub MMutasiBarangNonMedis_Click()
    frmMutasiBarangNM.Show
End Sub

Private Sub mMutasiBarangNM_Click()
    frmMutasiBarangNM.Show
End Sub

Private Sub mnAsuransi_Click()
    frmDaftarPasienAsuransi.Show
End Sub

Private Sub MnBkRegister_Click()
    FrmBukuRegister.Show
End Sub

Private Sub mnDaftarAsuransi_Click()
    frmAsuransi.Show
End Sub

Private Sub MnDaftarPasienMeninggal_Click()
    frmDaftarPasienMeninggal.Show
End Sub

Private Sub mndaftarreservasi_Click()
    frmDaftarReservasiPasien.Show
End Sub

Private Sub MndKodisiPulang_Status_Click()
    strCetak = "LapKunjunganKonPulang_Status"
    With frmLapRKP_KPSK
        .Caption = "Medifirst2000 - Kunjungan Pasien Berdasarkan Status & Kondisi pulang pasien"
        .Show
    End With
End Sub

Private Sub mnDokter_Click()
    frmJadwalPraktekDokter.Show
End Sub

Private Sub mnHadir_Click()
    frmStatusHadirDokter.Show
End Sub

Private Sub MNilaiPersediaan_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmNilaiPersediaan.Show
End Sub

Private Sub MNilaiPersediaanNonMedis_Click()
    mstrKdKelompokBarang = "01"
    frmNilaiPersediaanNM.Show
End Sub


Private Sub mnINfoJadwal_Click()
    frmInformasiJadwalPraktek.Show
End Sub

Private Sub MnJenisOperasi_Status_Click()
    strCetak = "LapKunjunganJenisOperasi_Status"
    With frmLapRKP_KPSK
        .Caption = "Medifirst2000 - Kunjungan Pasien Berdasarkan Status & Jenis Operasi"
        .Show
    End With
End Sub

Private Sub mnlogout_Click()
    Dim adoCommand As New ADODB.Command
    openConnection
    sepuh = True
    strQuery = "UPDATE Login SET Status = '0' " & _
    "WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
    adoCommand.ActiveConnection = dbConn
    adoCommand.CommandText = strQuery
    adoCommand.CommandType = adCmdText
    
    adoCommand.Execute

    dTglLogout = Now
    Call subSp_HistoryLoginAplikasi("U")
    Unload Me
    Call openConnection
End Sub

Private Sub mnRegPenunjang_Click()
    frmRegistrasiRJPenunjang.Show
End Sub

Private Sub mnReservasi_Click()
    frmReservasi.Show
End Sub

Private Sub mnRkstatusKasusPenyakit_Click()
    strCetak = "LapKunjunganSt_PnyktPsn"
    With frmLapRKP_KPSK
        .Caption = "Medifirst2000 - Kunjungan Pasien Berdasarkan Status & Kasus Penyakit Pasien"
        .Show
    End With
End Sub

Private Sub mnSelesai_Click()
    Dim pesan As VbMsgBoxResult
    Dim adoCommand As New ADODB.Command
    pesan = MsgBox("Tutup aplikasi ", vbQuestion + vbYesNo, "Konfirmasi")
    If pesan = vbYes Then

        openConnection
        strQuery = "UPDATE Login SET Status = '0' " & _
        "WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
        adoCommand.ActiveConnection = dbConn
        adoCommand.CommandText = strQuery
        adoCommand.CommandType = adCmdText
        adoCommand.Execute

        dTglLogout = Now
        Call subSp_HistoryLoginAplikasi("U")
        End
    Else
    End If
End Sub

Private Sub mnuibs_Click()
    KdInstalasi = "04"
End Sub

Private Sub MnSensusH_Click()
    frmLapSensusHarian.Show
End Sub

Private Sub MnStausBRujukan_Click()
    strCetak = "LapKunjunganRujukanBStatus"
    With frmLapRKP_KPSK
        .Caption = "Medifirst2000 - Kunjungan Pasien Berdasarkan Status & Rujukan Pasien"
        .Show
    End With
End Sub

Private Sub mnu59_Click()
    frmDataKegiatanKesehatanLingkunganRL5.Show
End Sub

Private Sub mnuBkecDstatus_Click()
strCetak = "LapKunjunganWilayahKotaKecJenisStatus"
    frmLapRKP_KPSK.Show
    frmLapRKP_KPSK.cmdgrafik.Visible = False
End Sub

Private Sub mnuBukuRegisterPasien_Click()
FrmBukuRegisterPasien.Show
'FrmBukuRegister.Show

End Sub

Private Sub mnuBukuRegisterPelayanan_Click()
  FrmBukuRegisterPelayanan.Show
End Sub

Private Sub mnuBWilayahDJenis_Click()
    strCetak = "LapKunjunganWilayahJenisStatus"
'    With frmLapRKP_KPSK
'        .Caption = "Medifirst2000 - Kunjungan Pasien Berdasarkan Wilayah & Jenis Pasien"
'        .Show
'    End With
    
    With frmDataPengunjung
        .Caption = "Medifirst2000 - Kunjungan Pasien Rumah Sakit"
        .Show
    End With
End Sub

Private Sub mnuCaraPembayaran_Click()
    frmRL1HAL6.Show
End Sub

Private Sub mnucdp_Click()
    frmCariPasien.Show
End Sub

Private Sub mnuCekKepesertaanBPJS_Click()
    frmCekKepesertaanBPJSVclaim.Show
End Sub

Private Sub mnuDaftarPasienLama_Click()
    frmDaftarPasienLama.Show
End Sub

Private Sub mnuDaftarPasienRJ_Click()
    frmDaftarPasienRawatJalan.Show
End Sub

Private Sub mnuDataDasarRumahSakit_Click()
    frmRL3.Show
End Sub

Private Sub mnuDataIndividualMorbiditasPasienRawatInap1_Click()
    frmDaftarPasienRL21.Show
End Sub

Private Sub mnuDataIndividualMorbiditasPasienRawatInap2_Click()
    frmDaftarPasienRL22.Show
End Sub

Private Sub mnuDataKeadaanMorbiditasPasienRawatInapRumahSakit_Click()
    frmRL2aHAL1.Show
End Sub

Private Sub mnuDataKeadaanMorbiditasPasienRawatJalanRumahSakit_Click()
    frmRL2bHAL1.Show
End Sub

Private Sub mnuDataKeadaanMorbiditasRawatInapSurveilensTerpaduRumahSakit_Click()
    frmRL2asurveilans.Show
End Sub

Private Sub mnuDataKeadaanMorbiditasRawatJalanSurveilensTerpaduRumahSakit_Click()
    frmRL2bsurveilans.Show
End Sub

Private Sub mnuDataKePenerimaan_Click()
    frmRL52.Show
End Sub

Private Sub mnuDataPeralatanMedikRumahSakit_Click()
    frmRL51.Show
End Sub

Private Sub mnuDataStatusImunisasi_Click()
    frmRL2C.Show
End Sub

Private Sub mnuDetailDiagnosaKeperawatan_Click()
    frmDetailDiagnosaKeperawatan.Show
End Sub

Private Sub mnuDiagnosaKeperawatan_Click()
    frmMasterDiagnosaKeperawatan.Show
End Sub

Private Sub mnuipb_Click()
    frmInfoPesanBarang.Show
End Sub

Private Sub mnuDoktoral_Click()
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmdoktoral.Show
hell:
End Sub

Private Sub mnuFormulirdataIndividualKepegawaiaanDirektoratJendralPelayananMedik_Click()
    frmDaftarPegawai.Show
End Sub

Private Sub mnuFormulirPelaporanInfeksiNosokmial_Click()
    frmRL6.Show
End Sub

Private Sub mnuiprs_Click()
    frmLapIndPlynRS.Show
End Sub

Private Sub mnuKegiatan_Click()
    frmKegiatanFarmasi.Show
End Sub

Private Sub mnuJumlahKetenagaanKesehatanMenurutJenis_Click()
    frmRL4a.Show
End Sub

Private Sub mnuJumlahKetenagaanNonKesehatanMenurutJenis_Click()
    frmRL4b.Show
End Sub

Private Sub mnuKegiatanKeluargaBerencana_Click()
    frmRL1HAL4.Show
End Sub

Private Sub mnuKeluargaPasien_Click()
    frmKeluargaPegawai.Show
End Sub

Private Sub mnuKunjunganRawatJalan_Click()
    frmKunjunganRawatJalan.Show
End Sub

Private Sub mnulatihan_Click()
    frmRL1HAL5.Show
End Sub

Private Sub mnuMDD_Click()
    frmMasterDiagnosa.Show
End Sub

Private Sub mnuMDKP_Click()
    frmMasterDaftarKontrolPasien.Show
End Sub

Private Sub mnuMDP_Click()
    frmMasterPelayanan.Show
End Sub

Private Sub mnuMDU_Click()
    frmMasterUmum.Show
End Sub

Private Sub mnuMDW_Click()
    frmMasterWilayah.Show
End Sub

Private Sub mnuMonitoringData_Click()
    frmTempDW.Show
End Sub

Private Sub mnump_Click()
    frmLapMorbiditas.Show
End Sub

Private Sub mnuPascaSarjana_Click()
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmpascasarjana.Show
hell:
End Sub

Private Sub mnupb_Click()
    frmPemesananBarang.Show
End Sub

Private Sub mnupb2_Click()
    strPasien = "Baru"
    boltampil = True
    frmPasienBaru.Show
End Sub

Private Sub mnupelayananrawatinap_Click()
    frmRL1HAL1.Show
End Sub

Private Sub mnuPemakaianBahandanAlat_Click()
    frmPemakaianBahanAlat.Show
End Sub

Private Sub mnuPengadaanObat_Click()
    On Error GoTo hell
    frmRL1HAL3.Show
hell:
End Sub

Private Sub mnuPengunjungRumahSakit_Click()
    frmRL1HAL2.Show
End Sub

Private Sub mnupl_Click()
'    frmRegistrasiAll.Show
    frmCariPasien.Show
End Sub

Private Sub mnur10p_Click()
    FrmPeriodeLaporanTopTen1.Show
    FrmPeriodeLaporanTopTen1.Caption = "Medifirst2000 - Rekapitulasi 10 Besar Penyakit"
End Sub

Private Sub mnuRekapKP_Click()
    strCetak = "LapKunjunganJenisStatus"
    With frmLapRKP_KPSK
        .Caption = "Medifirst2000 - Kunjungan Pasien Berdasarkan Status & Jenis Pasien"
        .Show
    End With
End Sub

Private Sub mnuRekapPasienRIperSMF_Click()
    FrmLaporanPasienRIPerSMF.Show
End Sub

Private Sub mnurl1_1_Click()
    frm1sub1New.Show
End Sub

Private Sub mnurl1_2_Click()
    frmRLNew1Sub2.Show
End Sub

Private Sub mnuRL1_2New_Click()
    frmRLNew1Sub2New2.Show
End Sub

Private Sub mnurl1_3_Click()
    frm1sub3New.Show
End Sub

Private Sub mnuRL1_3New_Click()
    frm1sub3New2.Show
End Sub

Private Sub mnurl2new_Click()
    frmRL2New2.Show
End Sub

Private Sub mnuRL2New2_Click()
    frmRL2New.Show
End Sub

Private Sub mnurl3_1_Click()
    frm3sub01New.Show
End Sub

Private Sub mnurl3_10_Click()
    frmRL3Sub3_10New.Show
End Sub

Private Sub mnuRL3_10New_Click()
    frmRL3Sub3_10New2.Show
End Sub

Private Sub mnurl3_11_Click()
    frm3sub11New.Show
End Sub

Private Sub mnuRL3_11New_Click()
    frm3sub11New2.Show
End Sub

Private Sub mnurl3_12_Click()
    frmRL3Sub3_12New.Show
End Sub

Private Sub mnuRL3_12New_Click()
    frmRL3Sub3_12New2.Show
End Sub

Private Sub mnurl3_13_Click()
    frm3sub13New.Show
End Sub

Private Sub mnuRL3_13New_Click()
    frm3sub13New2.Show
End Sub

Private Sub mnurl3_14_Click()
    frm3sub14New.Show
End Sub

Private Sub mnuRL3_14New_Click()
    frm3sub14New2.Show
End Sub

Private Sub mnurl3_15_Click()
    frm3sub15New.Show
End Sub

Private Sub mnuRL3_15New_Click()
    frm3sub15New2.Show
End Sub

Private Sub mnuRL3_1New_Click()
    frm3sub01New2.Show
End Sub

Private Sub mnurl3_2_Click()
    frmRL3Sub3_2New.Show
End Sub

Private Sub mnuRL3_2New_Click()
    frmRL3Sub3_2New2.Show
End Sub

Private Sub mnurl3_3_Click()
    frm3sub03New.Show
End Sub

Private Sub mnuRL3_3New_Click()
    frm3sub03New2.Show
End Sub

Private Sub mnurl3_4_Click()
    frmRL3Sub3_4New.Show
End Sub

Private Sub mnuRL3_4New_Click()
    frmRL3Sub3_4New2.Show
End Sub

Private Sub mnurl3_5_Click()
    frm3sub05New.Show
End Sub

Private Sub mnuRL3_5New_Click()
    frm3sub05New2.Show
End Sub

Private Sub mnurl3_6_Click()
    frmRL3Sub3_6New.Show
End Sub

Private Sub mnuRL3_6New_Click()
    frmRL3Sub3_6New2.Show
End Sub

Private Sub mnurl3_7_Click()
    frm3sub07New.Show
End Sub

Private Sub mnuRL3_7New_Click()
    frm3sub07New2.Show
End Sub

Private Sub mnurl3_8_Click()
    frmRL3Sub3_8New.Show
End Sub

Private Sub mnuRL3_8New_Click()
    frmRL3Sub3_8New2.Show
End Sub

Private Sub mnurl3_9_Click()
    frm3sub09New.Show
End Sub

Private Sub mnuRL3_9New_Click()
    frm3sub09New2.Show
End Sub

Private Sub mnuRL4_ANew_Click()
    frmRL4Sub4_aNew2.Show
End Sub

Private Sub mnurl4a_Click()
    frmRL4Sub4_aNew.Show
End Sub

Private Sub mnuRL4APenyebabKecelakaan_Click()
    frmRL4Sub4_PenyebabKecelakaanRINew2.Show
End Sub

Private Sub mnurl4apk_Click()
    frmRL4Sub4_PenyebabKecelakaanRINew.Show
End Sub

Private Sub mnurl4b_Click()
    frm4sub02New.Show
End Sub

Private Sub mnuRL4BNew_Click()
    frmRL4Sub4_PenyebabKecelakaanRJNew2.Show
End Sub

Private Sub mnuRL4BNew2_Click()
    frm4sub02New2.Show
End Sub

Private Sub mnurl4bpk_Click()
    frmRL4Sub4_PenyebabKecelakaanRJNew.Show
End Sub

Private Sub mnurl5_1_Click()
    frmRL5Sub5_1New.Show
End Sub

Private Sub mnuRL5_1New_Click()
    frmRL5Sub5_1New2.Show
End Sub

Private Sub mnurl5_2_Click()
    frm5sub02New.Show
End Sub

Private Sub mnuRL5_2New_Click()
    frm5sub02New2.Show
End Sub

Private Sub mnurl5_3_Click()
    frmRL5Sub5_3New.Show
End Sub

Private Sub mnuRL5_3New_Click()
    frmRL5Sub5_3New2.Show
End Sub

Private Sub mnurl5_4_Click()
    frm5sub04New.Show
End Sub

Private Sub mnuRL5_4New_Click()
    frm5sub04New2.Show
End Sub

Private Sub mnuSarjana_Click()
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmsarjana.Show
hell:
End Sub

Private Sub mnuSarjanaMuda_Click()
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmsarjanamuda.Show
hell:
End Sub

Private Sub mnusb_Click()
    frmStokBrg.Show
End Sub

Private Sub mnuTerimaBarangLangsung_Click()
    frmTerimaBarangLangsung.Show
End Sub

Private Sub mnuSekolahMenengahTingkatAtas_Click()
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmsmta.Show
hell:
End Sub

Private Sub mnuSMTP_Click()
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmsmtp.Show
hell:
End Sub

Private Sub mnuTenagaGizi_Click()
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmTenagaGizi.Show
hell:

End Sub

Private Sub mnuTenagaKefarmasian_Click()
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmTenagaKefarmasian.Show
hell:

End Sub

Private Sub mnuTenagaKeperawatan_Click()
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmTenagakeperawatan.Show
hell:
End Sub

Private Sub mnuTenagaKesehatanMasyarakat_Click()
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmTenagaKesehatanMasyarakat.Show
hell:
End Sub

Private Sub mnuTenagaKeteknisianMedis_Click()
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmTenagaKeteknisanMedis.Show
hell:
End Sub

Private Sub mnuTenagaKeterapianFisik_Click()
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmTenagaKeterapianFisik.Show
hell:
End Sub

Private Sub mnuTenagaMedis_Click()
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmTenagaMedis.Show
hell:
End Sub

Private Sub mnuTujuanNRencanaTindakan_Click()
    On Error GoTo hell
    frmTujuanNRencanaTindakan.Show
hell:
End Sub

Private Sub mnuVerDaftar_Click()
    frmValidasiDataPendaftaran.Show
End Sub

Private Sub mnuVerJenisPasien_Click()
    frmVerifikasiDataAsuransiPasien.Show
End Sub

Private Sub MPemakaianBarangKaryawan_Click()
    frmPOAKaryawan.Show
End Sub

Private Sub MRangeHasilKomponenScore_Click()
    FrmPointRangeHasilKomponenScore.Show
End Sub

Private Sub mRK_Click()
    FrmPeriodeLaporanKematianNew.Show
End Sub

Private Sub mSettingPrinter_Click()
'    frmSetupPrinter.Show
    frmSetupPrinter2.Show
End Sub

Private Sub MSettingPrinterBarcode_Click()
    frmSetPrinter.Show
End Sub

Private Sub MsLaporanKelasdanStatus_Click()
    strCetak = "LapKunjunganKelasStatus"
    With frmLapRKP_KPSK
        .Caption = "Medifirst2000 - Kunjungan Pasien Berdasarkan Kelas & Satus Pasien"
        .Show
    End With
End Sub

Private Sub MStokBarangNonMedis_Click()
    frmStokBarangNonMedis.Show
End Sub

Private Sub MStokOpname_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmStokOpname.Show
End Sub

Private Sub MStokOpnameNonMedis_Click()
    mstrKdKelompokBarang = "01"
    frmStokOpnameNM.Show
End Sub

Private Sub mTentang_Click()
    frmAbout.Show
End Sub

Private Sub mTentang2_Click()
    frmAbout.Show
End Sub

Private Sub MTindakanOperasi_Click()
    frmDataTindakanOperasi.Show
End Sub

Private Sub RPBD_Click()
    strCetak = "LapKunjunganBDiagnosa"
    frmLapRKP_KPSK.Show
    frmLapRKP_KPSK.cmdgrafik.Visible = False
End Sub

Private Sub RPBW_Click()
    strCetak = "LapKunjunganBwilayah"
    frmLapRKP_KPSK.Show
    frmLapRKP_KPSK.cmdgrafik.Visible = False

End Sub

Private Sub RPBWD2_Click()
    strCetak = "LapKunjunganPasienBDiagnosaWilayah"
    frmLapRKP_KPSK.Show
    frmLapRKP_KPSK.cmdgrafik.Visible = False
End Sub
