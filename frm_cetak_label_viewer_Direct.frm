VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_cetak_label_viewer_Direct 
   Caption         =   "Cetal Label Registrasi"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_cetak_label_viewer_Direct.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   6765
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
Attribute VB_Name = "frm_cetak_label_viewer_Direct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Report As New Cr_cetakLabel_bayar
Dim Report As New Cr_cetakLabel_Direct

Private Sub Form_Load()
On Error GoTo errLoad
    Dim adocomd As New ADODB.Command

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

'    adocomd.ActiveConnection = dbConn
'    If (mstrKdInstalasi = "02") Or (mstrKdInstalasi = "11") Or (mstrKdInstalasi = "06") Then
'    adocomd.CommandText = "SELECT * from V_CetakLabelRegistrasiPasienMRS " _
'                         & " WHERE (NoPendaftaran) =('" & mstrNoPen & "')"
'    Else
'    adocomd.CommandText = "SELECT * from V_CetakLabelRegistrasiPasienMRS " _
'                         & " WHERE (NoPendaftaran) =('" & mstrNoPen & "')"
'    End If
'    adocomd.CommandType = adCmdText
    
    
    
    Dim tanggal As String
    tanggal = Format(TglPeriodeAwal, "MMMM yyyy") '& " S/d " & Format(frmregister.DTPickerAkhir, "dd MMMM yyyy")
    
'report untuk yg sudah bayar
If (mstrKdInstalasi = "02") Or (mstrKdInstalasi = "11") Or (mstrKdInstalasi = "06") Then
    
    With Report
   ' .Database.AddADOCommand dbConn, adocomd
        If mstrNoStruk = "" Then
            .txtStatusBayar.SetText "BELUM Bayar"
        Else
            .txtStatusBayar.SetText "SUDAH Bayar"
            .txtNoBKM.SetText mstrNoBKM
            .txtJmlBayar.SetText mcurAll_HrsDibyr
        End If
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS '& " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text3.SetText strEmail 'strWebSite & ", " & strEmail
        
        .txtTglMasuk.SetText dTglDaftar
        '.udtgl.SetUnboundFieldSource ("{ado.tglmasuk}")
        .txtNoPendaftaran.SetText mstrNoPen
        '.usnodft.SetUnboundFieldSource ("{ado.nopendaftaran}")
        .txtNoCM.SetText mstrNoCM
        '.usnocm.SetUnboundFieldSource ("{Ado.nocm}")
        .txtRuangPeriksa.SetText sRuangPeriksa
        '.usruangperiksa.SetUnboundFieldSource ("{Ado.Ruang Tujuan}")
        
        .txtNamaPasien.SetText sNamaPasien
        '.usnmpasien.SetUnboundFieldSource ("{Ado.nama pasien}")
        
        '.usJK.SetUnboundFieldSource ("{ado.jk}")
        .txtJK.SetText sJK
        '.usUmur.SetUnboundFieldSource ("{ado.Umur}")
        .txtUmur.SetText sUmur
        '.usAlamat.SetUnboundFieldSource ("{Ado.Alamat}")
        .txtAlamat.SetText sAlamat
        
        '.usPenjamin.SetUnboundFieldSource ("{Ado.JenisPasien}")
        .txtPenjamin.SetText sPenjamin
        '.USJenisKelas.SetUnboundFieldSource ("{ado.DetailJenisJasaPelayanan}")
        .txtKelas.SetText sKelas
        '.usNoBed.SetUnboundFieldSource ("{Ado.NoBed}")
        .txtNoBed.SetText sNoBed
        
        If mstrKdInstalasi = "03" Then
            .txtNoRuangan.SetText "No. Kamar"
            '.usNoRuangan.SetUnboundFieldSource ("{Ado.NamaKamar}")
            .txtNoRuang.SetText sNoBed
            .txtNoBed.Suppress = False
            .txtBed.Suppress = False
        Else
            .txtNoRuangan.SetText "No. Ruangan"
            '.usNoRuangan.SetUnboundFieldSource ("{Ado.No. Ruangan}")
            .txtBed.SetText sNoBed
            .txtNoBed.Suppress = True
            .txtBed.Suppress = True
        End If
        
        '.usnoantri.SetUnboundFieldSource ("{Ado.NoAntrian}")
        .txtNoAntrian.SetText iNoAntrian
        .ususer.SetText (strNmPegawai)
        '.txtMessage.SetText (rs(0).Value)
        .PrintOut False
    End With

Else
    With ReportNonBayar
    .Database.AddADOCommand dbConn, adocomd
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS '& " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text3.SetText strEmail 'strWebSite & ", " & strEmail
        .udtgl.SetUnboundFieldSource ("{ado.tglmasuk}")
        .usnodft.SetUnboundFieldSource ("{ado.nopendaftaran}")
        .usnocm.SetUnboundFieldSource ("{Ado.nocm}")
        .usruangperiksa.SetUnboundFieldSource ("{Ado.Ruang Tujuan}")
        
        .usnmpasien.SetUnboundFieldSource ("{Ado.nama pasien}")
        .usJK.SetUnboundFieldSource ("{ado.jk}")
        .usUmur.SetUnboundFieldSource ("{ado.Umur}")
        .usAlamat.SetUnboundFieldSource ("{Ado.Alamat}")
'        .usAlias.SetUnboundFieldSource ("{Ado.Alias}")
        
        .usPenjamin.SetUnboundFieldSource ("{Ado.JenisPasien}")
        .USJenisKelas.SetUnboundFieldSource ("{ado.DetailJenisJasaPelayanan}")
        .usNoBed.SetUnboundFieldSource ("{Ado.NoBed}")
        
        If mstrKdInstalasi = "03" Then
            .txtNoRuangan.SetText "No. Kamar"
            .usNoRuangan.SetUnboundFieldSource ("{Ado.NamaKamar}")
            .txtNoBed.Suppress = False
            .usNoBed.Suppress = False
        Else
            .txtNoRuangan.SetText "No. Ruangan"
            .usNoRuangan.SetUnboundFieldSource ("{Ado.No. Ruangan}")
            .txtNoBed.Suppress = True
            .usNoBed.Suppress = True
        End If
        
        .usnoantri.SetUnboundFieldSource ("{Ado.NoAntrian}")
        .ususer.SetText (strNmPegawai)
        '.txtMessage.SetText (rs(0).Value)
        .PrintOut False
    End With
End If
Screen.MousePointer = vbHourglass
'With CRViewer1
'        .ReportSource = Report
'        .ViewReport
'        .Zoom 1
'End With
Screen.MousePointer = vbDefault
    Unload Me

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm_cetak_label_viewer_Direct = Nothing
End Sub
