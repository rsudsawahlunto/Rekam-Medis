VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakLembarMasukDanKeluar 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
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
Attribute VB_Name = "frmCetakLembarMasukDanKeluar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crLembarMasukDanKeluar

Private Sub Form_Load()
    Call openConnection
    
    strSQL = "SELECT * FROM V_LembarMasukDanKeluarRI WHERE NoPendaftaran = '" & mstrNoPen & "'"
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    
    With Report
        If rs.EOF = True Then Exit Sub
        .txtNoCM.SetText rs("NoCM")
        .txtRPerawatan.SetText rs("RuanganPerawatan")
        .txtPendidikan.SetText IIf(IsNull(rs("Pendidikan")), "-", rs("Pendidikan"))
        .txtPekerjaan.SetText IIf(IsNull(rs("Pekerjaan")), "-", rs("Pekerjaan"))
        .txtAlamat.SetText IIf(IsNull(rs("AlamatLengkapPasien")), "-", rs("AlamatLengkapPasien"))
        .txtKelurahan.SetText IIf(IsNull(rs("Kelurahan")), "-", rs("Kelurahan"))
        .txtKecamatan.SetText IIf(IsNull(rs("Kecamatan")), "-", rs("Kecamatan"))
        .txtKota.SetText IIf(IsNull(rs("Kota")), "-", rs("Kota"))
        .txtNamaPJawab.SetText IIf(IsNull(rs("NamaPJ")), "-", rs("NamaPJ"))
        .txtAlamatPJawab.SetText IIf(IsNull(rs("AlamatPJ")), "-", rs("AlamatPJ"))
        .txtTlpPJawab.SetText IIf(IsNull(rs("TeleponPJ")), "-", rs("TeleponPJ"))
        .txtHubPJawab.SetText IIf(IsNull(rs("Hubungan")), "-", rs("Hubungan"))

        .txtNmPasien.SetText IIf(IsNull(rs("NamaPasien")), "-", rs("NamaPasien")) & " ( " & IIf(IsNull(rs("JK")), "-", rs("JK")) & " ) "
        .txtKelasPelayanan.SetText IIf(IsNull(rs("KelasPelayanan")), "-", rs("KelasPelayanan"))
        .txtTglLahir.SetText IIf(IsNull(rs("TglLahir")), "-", rs("TglLahir")) & " - " & IIf(IsNull(rs("Umur")), "-", rs("Umur"))
        .txtSMF.SetText IIf(IsNull(rs("SMF")), "-", rs("SMF"))
        
        .txtAgama.SetText IIf(IsNull(rs("Agama")), "-", rs("Agama"))
        .txtNoID.SetText IIf(IsNull(rs("NoIdentitas")), "-", rs("NoIdentitas"))
        .txtBangsa.SetText IIf(IsNull(rs("Warganegara")), "-", rs("Warganegara"))
        .txtSuku.SetText IIf(IsNull(rs("Suku")), "-", rs("Suku"))
        .txtStatusKawin.SetText IIf(IsNull(rs("StatusNikah")), "-", rs("StatusNikah"))
        .txtJenisPasien.SetText IIf(IsNull(rs("JenisPasien")), "-", rs("JenisPasien"))
        .txtCaraMasuk.SetText IIf(IsNull(rs("CaraMasuk")), "-", rs("CaraMasuk"))
        .txtTglMasuk.SetText IIf(IsNull(rs("TglMasuk")), "-", rs("TglMasuk"))
        .txtNoKamar.SetText IIf(IsNull(rs("NoKamar")), "-", rs("NoKamar"))
        
        .txtNamaRS.SetText strNNamaRS
        .txtAlamatRS.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtWebsiteRS.SetText "Website : " & strWebSite & "    Email : " & strEmail
 '       .PrintOut False
    End With
    

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .EnableGroupTree = False
        .Zoom 1
    End With

'    Unload Me

End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub
