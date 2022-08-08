VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakLembarMasukDanKeluarV2 
   Caption         =   "Cetak Lembur Masuk Dan Keluar"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13050
   Icon            =   "frmCetakLembarMasukDanKeluarV2.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   13050
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
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakLembarMasukDanKeluarV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Report As New crLembarMasukDanKeluar
'
'Private Sub Form_Load()
'    On Error GoTo hell
'    Call openConnection
'
'    strSQL = "SELECT * FROM V_LembarMasukDanKeluarRI WHERE NoPendaftaran = '" & mstrNoPen & "'"
'
'    Set rs = Nothing
'    Call msubRecFO(rs, strSQL)
'    With Report
'        .TxtRS.SetText ""
'        .TxtPuskesmas.SetText ""
'        .txtDokter.SetText ""
'        .TxtTPolis.SetText ""
'        .TxtYpolis.SetText " "
'        .TxtYPHB.SetText ""
'        .TxtTPHB.SetText ""
'        .TxtSendiri.SetText " "
'        .TxtDarurat.SetText " "
'        .TxtBiasa.SetText " "
'        .Text123.SetText " "
'
'        .txtNoCM.SetText rs("NoCM")
'        .txtNoPendaftaran.SetText rs("Nopendaftaran")
'        .TxtAyah.SetText IIf(IsNull(rs("NamaAyah")), "-", rs("NamaAyah"))
'        .TxtIbu.SetText IIf(IsNull(rs("NamaIbu")), "-", rs("NamaIbu"))
'        .TxtIstri.SetText IIf(IsNull(rs("NamaSuamiIstri")), "-", rs("NamaSuamiIstri"))
'        .TxtRperawatan.SetText IIf(IsNull(rs("RuanganPerawatan")), "", rs("RuanganPerawatan"))
'        .TxtSt.SetText IIf(IsNull(rs("StatusNikah")), "", rs("StatusNikah"))
'        .txtPendidikan.SetText IIf(IsNull(rs("Pendidikan")), "-", rs("Pendidikan"))
'        .TxtPkerjaan.SetText IIf(IsNull(rs("Pekerjaan")), "-", rs("Pekerjaan"))
'        .txtAlamat.SetText IIf(IsNull(rs("AlamatLengkapPasien")), "-", rs("AlamatLengkapPasien"))
'        .TxtKampung.SetText IIf(IsNull(rs("Kelurahan")), "-", rs("Kelurahan"))
'        .TxtKec.SetText IIf(IsNull(rs("Kecamatan")), "-", rs("Kecamatan"))
'        .txtKota.SetText IIf(IsNull(rs("Kota")), "-", rs("Kota"))
'
'        .TxtPJB.SetText IIf(IsNull(rs("NamaPJ")), "-", rs("NamaPJ"))
'        .TxtALamatJB.SetText IIf(IsNull(rs("AlamatPJ")), "-", rs("AlamatPJ"))
'        .TxtPerkerjaanJB.SetText IIf(IsNull(rs("PekerjaanPJ")), "-", rs("PekerjaanPJ"))
'        .TxtKotaJB.SetText IIf(IsNull(rs("Kotapj")), "-", rs("KotaPJ"))
'        .TxtKecJB.SetText IIf(IsNull(rs("KecamatanPJ")), "-", rs("KecamatanPJ"))
'        .TxtDesaJB.SetText IIf(IsNull(rs("KelurahanPJ")), "-", rs("KelurahanPJ"))
'
'        .txtTglMasuk.SetText IIf(IsNull(rs("TglMasuk")), "-", Format(rs("TglMasuk"), "dd/mm/YYYY HH:MM:SS"))
'        .txtnmpasien.SetText IIf(IsNull(rs("NamaPasien")), "-", rs("NamaPasien"))
'        .txtKelas.SetText IIf(IsNull(rs("KelasPelayanan")), "-", rs("KelasPelayanan"))
'        .txtBln.SetText IIf(IsNull(rs("Bln")), "-", rs("bln"))
'        .txtThn.SetText IIf(IsNull(rs("Thn")), "-", rs("thn"))
'        .txtHari.SetText IIf(IsNull(rs("Hari")), "-", rs("Hari"))
'        .TxtSMF.SetText IIf(IsNull(rs("SMF")), "-", rs("SMF"))
'        .txtAgama.SetText IIf(IsNull(rs("Agama")), "-", rs("Agama"))
'        .txtPendidikan.SetText IIf(IsNull(rs("Pendidikan")), "-", rs("Pendidikan"))
'        .TxtYPHB.SetText IIf(rs("JenisPasien") <> "UMUM", "X", " ")
'        .TxtTPHB.SetText IIf(rs("JenisPasien") = "UMUM", "X", " ")
'        .TxtYpolis.SetText IIf(Not IsNull(rs("KdRujukanAsal")), IIf(rs("KdRujukanAsal") = "12", "X", ""), "")
'        .TxtTPolis.SetText IIf(Not IsNull(rs("KdRujukanAsal")), IIf(rs("KdRujukanAsal") <> "12", "X", ""), "")
'        If rs("CaraMasuk") <> "Unit Gawat Darurat" Then
'            .TxtBiasa.SetText IIf(Not IsNull(rs("CaraMasuk")), IIf(rs("CaraMasuk") <> "02", "X", ""), "")
'        Else
'            .TxtDarurat.SetText IIf(Not IsNull(rs("CaraMasuk")), IIf(rs("CaraMasuk") <> "02", "X", ""), "")
'        End If
'        .TxtSt.SetText IIf(IsNull(rs("StatusNikah")), "-", rs("StatusNikah"))
'        .TxtRt.SetText IIf(IsNull(rs("RTRW")), "-", rs("RTRW"))
'        .txtJK.SetText IIf(IsNull(rs("JK")), "", IIf(rs("Jk") = "L", "Laki-Laki", "Perempuan"))
'        .TxtDokterPenanggungjawab.SetText IIf(IsNull(rs("DokterPemeriksa")), "-", rs("DokterPemeriksa"))
'
'        'Rujukan
'        If Not IsNull(rs("KdRujukanAsal")) Then
'            If rs("KdRujukanAsal") = "04" Or rs("KdRujukanAsal") = "03" Or rs("KdRujukanAsal") = "08" Or rs("KdRujukanAsal") = "09" Then
'                .TxtRS.SetText "X"
'            ElseIf rs("KdRujukanAsal") = "02" Then
'                .TxtPuskesmas.SetText "X"
'            ElseIf rs("KdRujukanAsal") = "04" Then
'                .txtDokter.SetText "X"
'            Else
'                .TxtSendiri.SetText "X"
'            End If
'        End If
'        .txtNamaRS.SetText strNNamaRS
'        .txtAlamatRS.SetText strNAlamatRS & " Kodepos : " & strNKodepos
'        .Text123.SetText strNAlamatRS & " Kodepos : " & strNKodepos
'        .web.SetText strWebsite & " " & "Email : " & strEmail
'    End With
'    Screen.MousePointer = vbHourglass
'    With CRViewer1
'        .ReportSource = Report
'        .ViewReport
'        .EnableGroupTree = False
'        .Zoom 1
'    End With
'    Screen.MousePointer = vbDefault
'    Exit Sub
'hell:
'    Call msubPesanError
'End Sub
'
'Private Sub Form_Resize()
'    CRViewer1.Top = 0
'    CRViewer1.Left = 0
'    CRViewer1.Height = ScaleHeight
'    CRViewer1.Width = ScaleWidth
'End Sub
