VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakDaftarAntrianPasien 
   Caption         =   "Medifisrt2000 - Daftar Antrian Pasien"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   Icon            =   "frmCetakDaftarAntrianPasien.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   5850
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
Attribute VB_Name = "frmCetakDaftarAntrianPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crDaftarAntrianPasien

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    Call openConnection
    Set frmCetakDaftarAntrianPasien = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = "select * " & _
    " from V_DaftarAntrianPasienMRS_IRM " & _
    " where ([Nama Pasien] like '%" & frmDaftarAntrianPasien.txtParameter.Text & "%' OR NoCM like '%" & frmDaftarAntrianPasien.txtParameter.Text & "%' OR Ruangan like '%" & frmDaftarAntrianPasien.txtParameter.Text & "%') and TglMasuk between '" & Format(frmDaftarAntrianPasien.dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(frmDaftarAntrianPasien.dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "' and  [Status Periksa] = '" & frmDaftarAntrianPasien.dcStatusPeriksa.Text & "'" & _
    " order by TglMasuk DESC, [No. Urut] ASC, NoCM ASC"

    adocomd.CommandType = adCmdText

    With Report
        .Database.AddADOCommand dbConn, adocomd

        .txtNamaRS.SetText strNNamaRS
        .txtAlamatRS.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .txtWebsiteRS.SetText strWebsite & ", " & strEmail
        .txtPeriode.SetText "" & " Periode : " & Format(frmDaftarAntrianPasien.dtpAwal.value, "dd MMMM yyyy") & " s/d " & Format(frmDaftarAntrianPasien.dtpAkhir.value, "dd MMMM yyyy")

        .usRuangan.SetUnboundFieldSource ("{Ado.Ruangan}")
        .usStatusPeriksa.SetUnboundFieldSource ("{Ado.Status Periksa}")
        .udTglMasuk.SetUnboundFieldSource ("{Ado.TglMasuk}")
        .usNoUrut.SetUnboundFieldSource ("{Ado.No. Urut}")
        .usNoRegistrasi.SetUnboundFieldSource ("{Ado.NoPendaftaran}")
        .usNoCM.SetUnboundFieldSource ("{Ado.NoCM}")
        .usNamaPasien.SetUnboundFieldSource ("{Ado.Nama Pasien}")
        .usJK.SetUnboundFieldSource ("{Ado.JK}")
        .usUmur.SetUnboundFieldSource ("{Ado.Umur}")
        .usJenisPasien.SetUnboundFieldSource ("{Ado.Jenis Pasien}")
        .usKelas.SetUnboundFieldSource ("{Ado.Kelas}")
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .EnableGroupTree = False
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
    Set frmCetakDaftarAntrianPasien = Nothing
End Sub

