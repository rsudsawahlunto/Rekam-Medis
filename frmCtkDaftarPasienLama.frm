VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCtkDaftarPasienLama 
   Caption         =   "Cetak Dokumen Rekam Medis Pasien"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11100
   Icon            =   "frmCtkDaftarPasienLama.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   19609.86
   ScaleMode       =   0  'User
   ScaleWidth      =   9450
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5805
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
Attribute VB_Name = "frmCtkDaftarPasienLama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crDaftarPasienLama
Private Sub Form_Load()

    Me.WindowState = 2
    Dim adocomd As New ADODB.Command
    Call openConnection
    Set frmCtkDaftarPasienLama = Nothing

    Report.txtNamaRS.SetText strNNamaRS
    Report.txtAlamatRS.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    Report.txtWebsiteRS.SetText strWebsite & ", " & strEmail
    Report.txtTanggal.SetText Format(mdTglAwal, "dd MMMM yyyy HH:mm") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm")
    Report.txtTanggal.SetText "Periode : " & Format(mdTglAwal, "dd MMMM yyyy HH:mm") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm")
    Report.txtTanggal.SetText ("Periode  : " & Format(frmDaftarPasienLama.dtpAwal.value, "dd MMMM yyyy HH:mm") & " s/d " & Format(frmDaftarPasienLama.dtpAkhir, "dd MMMM yyyy HH:mm"))
    Report.txtTanggal.SetText ("Periode  : " & Format(frmDaftarPasienLama.dtpAwal.value, "dd MMMM yyyy HH:mm") & " s/d " & Format(frmDaftarPasienLama.dtpAkhir, "dd MMMM yyyy HH:mm"))
    strSQL = "select Ruangan,NoPendaftaran,NoCM,[Nama Pasien],JK,Alamat,TglMasuk,TglKeluar,[Cara Keluar],RuanganTujuan,TglPulang,[Cara Pulang],[Kondisi Pulang],JenisPasien,Kelas " & _
    " from V_DaftarPasienLamaRI where (NoCM like '%" & frmDaftarPasienLama.txtParameter.Text & "%' OR [Nama Pasien] like '%" & frmDaftarPasienLama.txtParameter.Text & "%' OR Alamat like '%" & frmDaftarPasienLama.txtParameter.Text & "%') " & _
    " AND (TglKeluar between '" & Format(frmDaftarPasienLama.dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(frmDaftarPasienLama.dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') AND Ruangan like '%" & frmDaftarPasienLama.dcRuangan.Text & "%' order by Ruangan"

    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, adocomd
    With Report
        .usNoRegistrasi.SetUnboundFieldSource ("{ado.NoPendaftaran}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.Nama Pasien}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .usAlamat.SetUnboundFieldSource ("{ado.Alamat}")
        .udTglMasuk.SetUnboundFieldSource ("{ado.TglMasuk}")
        .udTglKeluar.SetUnboundFieldSource ("{ado.TglKeluar}")
        .usCaraKeluar.SetUnboundFieldSource ("{ado.Cara Keluar}")
        .usRuanganTujuan.SetUnboundFieldSource ("{ado.RuanganTujuan}")
        .udTglPulang.SetUnboundFieldSource ("{ado.TglPulang}")
        .usKondisiPulang.SetUnboundFieldSource ("{ado.Kondisi Pulang}")
        .usCaraPulang.SetUnboundFieldSource ("{ado.Cara Pulang}")
        .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
        .usKelas.SetUnboundFieldSource ("{ado.Kelas}")
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
    Set frmCtkDaftarPasienLama = Nothing
End Sub

