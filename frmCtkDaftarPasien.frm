VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCtkDaftarPasien 
   Caption         =   "Cetak Dokumen Rekam Medis Pasien"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmCtkDaftarPasien.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   11400
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13815
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
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
Attribute VB_Name = "frmCtkDaftarPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crDaftarPasien

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Set frmCtkDaftarPasien = Nothing
    Dim adocomd As New ADODB.Command
    Report.txtNamaRS.SetText strNNamaRS
    Report.txtAlamatRS.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    Report.txtWebsiteRS.SetText strWebsite & ", " & strEmail
    Report.txtTanggal.SetText ("Periode  : " & Format(frmDaftarPasienRJRIIGD.dtpAwal.value, "dd MMMM yyyy HH:mm") & " s/d " & Format(frmDaftarPasienRJRIIGD.dtpAkhir, "dd MMMM yyyy HH:mm"))
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL & "ORDER BY NamaPasien"
    adocomd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, adocomd

    With Report
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")

        If frmDaftarPasienRJRIIGD.dcAsalPasien = "" Then
            .txtKecamatan.Suppress = False
            .txtKecamatan.SetText "Semua Kecamatan"
            .usKecamatan.SetUnboundFieldSource ("{ado.Kecamatan}")
        Else
            .txtKecamatan.Suppress = True
            .usKecamatan.SetUnboundFieldSource ("{ado.Kecamatan}")
        End If

        If frmDaftarPasienRJRIIGD.dcInstalasi = "" Then
            .txtInstalasi.Suppress = False
            .txtInstalasi.SetText "Semua Instalasi"
            .usInstalasi.Suppress = True
        Else
            .txtInstalasi.Suppress = True
            .usInstalasi.SetUnboundFieldSource ("{ado.NamaInstalasi}")

        End If

        If frmDaftarPasienRJRIIGD.dcRuangan = "" Then
            .txtRuangan.Suppress = False
            .txtRuangan.SetText "Semua Ruangan"
            .usRuangan.Suppress = True
        Else
            .txtRuangan.Suppress = True
            .usRuangan.SetUnboundFieldSource ("{ado.RuanganPerawatan}")
        End If

        .usKelurahan.SetUnboundFieldSource ("{ado.Kelurahan}")
        .usAlamat.SetUnboundFieldSource ("{ado.Alamat}")
        .usNamaRuangan.SetUnboundFieldSource ("{ado.RuanganPerawatan}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .usUmur.SetUnboundFieldSource ("{ado.Umur}")
        If frmDaftarPasienRJRIIGD.dcJenisPasien = "" Then
            .usJenisPasien.Suppress = True
        Else
            .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
        End If

        If frmDaftarPasienRJRIIGD.dcPenjamin = "" Then
            .usPenjamin.Suppress = True
        Else
            .usPenjamin.SetUnboundFieldSource ("{ado.NamaPenjamin}")
        End If
        .usKelas.SetUnboundFieldSource ("{ado.kdKelas}")
        .udTglMasuk.SetUnboundFieldSource ("{ado.TglMasuk}")
        .udTglKeluar.SetUnboundFieldSource ("{ado.TglKeluar}")
        .usStatusPulang.SetUnboundFieldSource ("{ado.StatusKeluar}")
        .usKondisiPulang.SetUnboundFieldSource ("{ado.KondisiPulang}")
        .usKasusPenyakit.SetUnboundFieldSource ("{ado.KasusPenyakit}")
        .usNoKamar.SetUnboundFieldSource ("{ado.NoKamar}")
        .usNoBed.SetUnboundFieldSource ("{ado.NoBed}")
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 1
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
    Set frmCtkDaftarPasien = Nothing
End Sub
