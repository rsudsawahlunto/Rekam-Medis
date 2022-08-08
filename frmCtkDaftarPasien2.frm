VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCtkDaftarPasien2 
   Caption         =   "Cetak Dokumen Rekam Medis Pasien"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13755
   Icon            =   "frmCtkDaftarPasien2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   13755
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13335
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
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCtkDaftarPasien2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim report2 As New crDaftarPasien2

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Dim adocomd As New ADODB.Command
    report2.txtNamaRS.SetText strNNamaRS
    report2.txtAlamatRS.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    report2.txtWebsiteRS.SetText strWebsite & ", " & strEmail
    report2.txtTanggal.SetText ("Periode  : " & Format(frmDaftarPasienRJRIIGD.dtpAwal.value, "dd MMMM yyyy HH:mm") & " s/d " & Format(frmDaftarPasienRJRIIGD.dtpAkhir, "dd MMMM yyyy HH:mm"))
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText
    report2.Database.AddADOCommand dbConn, adocomd

    With report2
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
        .usInstalasi.SetUnboundFieldSource ("{ado.NamaInstalasi}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPerawatan}")
        .usKecamatan.SetUnboundFieldSource ("{ado.Kecamatan}")
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
        .usKelas.SetUnboundFieldSource ("{ado.Kelas}")
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
        .ReportSource = report2
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
    Set frmCtkDaftarPasien2 = Nothing
End Sub
