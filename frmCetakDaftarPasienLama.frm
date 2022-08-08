VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakDaftarPasienLama 
   Caption         =   "Medifirst2000 - Data Daftar Pasien Lama"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   Icon            =   "frmCetakDaftarPasienLama.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9450
   ScaleMode       =   0  'User
   ScaleWidth      =   9450
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakDaftarPasienLama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crDaftarPasienLama

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    Call openConnection
    Set frmCetakDaftarPasienLama = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    
    If frmDaftarPasienLama.dcRuangan.BoundText = "001" Then
        If frmDaftarPasienLama.optPindahan.value = True Then
'            If frmDaftarPasienLama.dtpAwal.Day <> frmDaftarPasienLama.dtpAkhir.Day Or frmDaftarPasienLama.dtpAwal.Month <> frmDaftarPasienLama.dtpAkhir.Month Or frmDaftarPasienLama.dtpAwal.Year <> frmDaftarPasienLama.dtpAkhir.Year Then
'                adocomd.CommandText = "select top 100 * from V_DaftarPasienLamaRI where tglPulang is null and RuanganTujuan is not null AND (TglKeluar between '" & Format(frmDaftarPasienLama.dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(frmDaftarPasienLama.dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & frmDaftarPasienLama.dcRuangan.Text & "%' ORDER BY Ruangan,TglKeluar DESC, [Cara Pulang] DESC "
'            Else
                adocomd.CommandText = "select * from V_DaftarPasienLamaIGD where tglPulang is null and RuanganTujuan is not null AND (TglKeluar between '" & Format(frmDaftarPasienLama.dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(frmDaftarPasienLama.dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "')and ruangan like'" & frmDaftarPasienLama.dcRuangan.Text & "%' ORDER BY Ruangan,TglKeluar DESC, [Cara Pulang] DESC "
'            End If
        End If
    
        If frmDaftarPasienLama.optPulang.value = True Then
            If frmDaftarPasienLama.dtpAwal.Day <> frmDaftarPasienLama.dtpAkhir.Day Or frmDaftarPasienLama.dtpAwal.Month <> frmDaftarPasienLama.dtpAkhir.Month Or frmDaftarPasienLama.dtpAwal.Year <> frmDaftarPasienLama.dtpAkhir.Year Then
                adocomd.CommandText = "select NoPendaftaran, NoCM, [Nama Pasien], JK, Alamat, TglMasuk, TglKeluar, RuanganTujuan, TglPulang, [Kondisi Pulang], [Cara Pulang], JenisPasien, Kelas, Ruangan, KdRuangan, KdSubInstalasi, KdJenisTarif, KdKelas, Thn, Bln, Hr, TglPendaftaran, [Cara Keluar], NoPakai, dbo.Ambil_KdDiagnosaUtama(NoPendaftaran) AS DiagnosaUtama from V_DaftarPasienLamaIGD where tglPulang is not null and RuanganTujuan is null AND (TglKeluar between '" & Format(frmDaftarPasienLama.dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(frmDaftarPasienLama.dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & frmDaftarPasienLama.dcRuangan.Text & "%' ORDER BY Ruangan,TglKeluar DESC, [Cara Pulang] DESC "
            Else
                adocomd.CommandText = "select * from V_DaftarPasienLamaIGD where tglPulang is not null and RuanganTujuan is null AND (TglKeluar between '" & Format(frmDaftarPasienLama.dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(frmDaftarPasienLama.dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & frmDaftarPasienLama.dcRuangan.Text & "%'ORDER BY Ruangan,TglKeluar DESC, [Cara Pulang] DESC "
            End If
        End If
    
        If frmDaftarPasienLama.OptSemua.value = True Then
                adocomd.CommandText = "SELECT DISTINCT NoPendaftaran, NoCM, [Nama Pasien], JK, Alamat, TglMasuk, TglKeluar, RuanganTujuan, TglPulang, [Kondisi Pulang], [Cara Pulang], JenisPasien, Kelas, Ruangan, KdRuangan, KdSubInstalasi, KdJenisTarif, KdKelas, Thn, Bln, Hr, TglPendaftaran, [Cara Keluar], NoPakai, dbo.Ambil_KdDiagnosaUtama(NoPendaftaran) AS DiagnosaUtama FROM V_DaftarPasienLamaIGD " & _
                "where (TglKeluar between '" & Format(frmDaftarPasienLama.dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(frmDaftarPasienLama.dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & frmDaftarPasienLama.dcRuangan.Text & "%' and [Cara Keluar] like '%" & frmDaftarPasienLama.dcStatusKeluar.Text & "%' ORDER BY Ruangan,TglKeluar DESC, [Cara Pulang] DESC "
        End If
    Else
        If frmDaftarPasienLama.optPindahan.value = True Then
            If frmDaftarPasienLama.dtpAwal.Day <> frmDaftarPasienLama.dtpAkhir.Day Or frmDaftarPasienLama.dtpAwal.Month <> frmDaftarPasienLama.dtpAkhir.Month Or frmDaftarPasienLama.dtpAwal.Year <> frmDaftarPasienLama.dtpAkhir.Year Then
                adocomd.CommandText = "select top 100 * from V_DaftarPasienLamaRI where tglPulang is null and RuanganTujuan is not null AND (TglKeluar between '" & Format(frmDaftarPasienLama.dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(frmDaftarPasienLama.dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & frmDaftarPasienLama.dcRuangan.Text & "%' ORDER BY Ruangan,TglKeluar DESC, [Cara Pulang] DESC "
            Else
                adocomd.CommandText = "select * from V_DaftarPasienLamaRI where tglPulang is null and RuanganTujuan is not null AND (TglKeluar between '" & Format(frmDaftarPasienLama.dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(frmDaftarPasienLama.dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "')and ruangan like'" & frmDaftarPasienLama.dcRuangan.Text & "%' ORDER BY Ruangan,TglKeluar DESC, [Cara Pulang] DESC "
            End If
        End If
    
        If frmDaftarPasienLama.optPulang.value = True Then
            If frmDaftarPasienLama.dtpAwal.Day <> frmDaftarPasienLama.dtpAkhir.Day Or frmDaftarPasienLama.dtpAwal.Month <> frmDaftarPasienLama.dtpAkhir.Month Or frmDaftarPasienLama.dtpAwal.Year <> frmDaftarPasienLama.dtpAkhir.Year Then
    '            adocomd.CommandText = "select top 100 * from V_DaftarPasienLamaRI where tglPulang is not null and RuanganTujuan is null AND (TglKeluar between '" & Format(frmDaftarPasienLama.dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(frmDaftarPasienLama.dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & frmDaftarPasienLama.dcRuangan.Text & "%' ORDER BY Ruangan,TglKeluar DESC, [Cara Pulang] DESC "
                adocomd.CommandText = "select NoPendaftaran, NoCM, [Nama Pasien], JK, Alamat, TglMasuk, TglKeluar, RuanganTujuan, TglPulang, [Kondisi Pulang], [Cara Pulang], JenisPasien, Kelas, Ruangan, KdRuangan, KdSubInstalasi, KdJenisTarif, KdKelas, Thn, Bln, Hr, TglPendaftaran, [Cara Keluar], NoPakai, dbo.Ambil_KdDiagnosaUtama(NoPendaftaran) AS DiagnosaUtama from V_DaftarPasienLamaRI where tglPulang is not null and RuanganTujuan is null AND (TglKeluar between '" & Format(frmDaftarPasienLama.dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(frmDaftarPasienLama.dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & frmDaftarPasienLama.dcRuangan.Text & "%' ORDER BY Ruangan,TglKeluar DESC, [Cara Pulang] DESC "
            Else
                adocomd.CommandText = "select * from V_DaftarPasienLamaRI where tglPulang is not null and RuanganTujuan is null AND (TglKeluar between '" & Format(frmDaftarPasienLama.dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(frmDaftarPasienLama.dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & frmDaftarPasienLama.dcRuangan.Text & "%'ORDER BY Ruangan,TglKeluar DESC, [Cara Pulang] DESC "
            End If
        End If
    
        If frmDaftarPasienLama.OptSemua.value = True Then
    '        If frmDaftarPasienLama.dtpAwal.Day <> frmDaftarPasienLama.dtpAkhir.Day Or frmDaftarPasienLama.dtpAwal.Month <> frmDaftarPasienLama.dtpAkhir.Month Or frmDaftarPasienLama.dtpAwal.Year <> frmDaftarPasienLama.dtpAkhir.Year Then
    '            adocomd.CommandText = "select top 100 * from V_DaftarPasienLamaRI where (TglKeluar between '" & Format(frmDaftarPasienLama.dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(frmDaftarPasienLama.dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & frmDaftarPasienLama.dcRuangan.Text & "%'ORDER BY Ruangan,TglKeluar DESC, [Cara Pulang] DESC "
                adocomd.CommandText = "SELECT DISTINCT NoPendaftaran, NoCM, [Nama Pasien], JK, Thn, Bln, Hr, Alamat, TglMasuk, TglKeluar, RuanganTujuan, TglPulang, [Kondisi Pulang], [Cara Pulang], JenisPasien, Kelas, Ruangan, KdRuangan, KdSubInstalasi, KdJenisTarif, KdKelas, Thn, Bln, Hr, TglPendaftaran, [Cara Keluar], NoPakai, dbo.Ambil_KdDiagnosaUtama(NoPendaftaran) AS DiagnosaUtama FROM V_DaftarPasienLamaRI " & _
                "where (TglKeluar between '" & Format(frmDaftarPasienLama.dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(frmDaftarPasienLama.dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & frmDaftarPasienLama.dcRuangan.Text & "%' and [Cara Keluar] like '%" & frmDaftarPasienLama.dcStatusKeluar.Text & "%' ORDER BY Ruangan,TglKeluar DESC, [Cara Pulang] DESC "
    '        Else
    '            adocomd.CommandText = "select * from V_DaftarPasienLamaRI where (TglKeluar between '" & Format(frmDaftarPasienLama.dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(frmDaftarPasienLama.dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') and ruangan like'" & frmDaftarPasienLama.dcRuangan.Text & "%'ORDER BY Ruangan,TglKeluar DESC, [Cara Pulang] DESC "
    '        End If
        End If
    End If
    adocomd.CommandType = adCmdText
    With Report
        .txtNamaRS.SetText strNNamaRS
        .txtAlamatRS.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .txtWebsiteRS.SetText strWebsite & ", " & strEmail
        .txtTgl.SetText Format(Now, "dd mmmm yyyy")
        .txtUser.SetText strNmPegawai
    

        .Database.AddADOCommand dbConn, adocomd
        .txtPeriode.SetText "Periode :" & Format(mdTglAwal, "dd MMM yyyy HH:mm") & " - " & Format(mdTglAwal, "dd MMM yyyy HH:mm")
        .usNoRegistrasi.SetUnboundFieldSource ("{Ado.NoPendaftaran}")
        .usNoCM.SetUnboundFieldSource ("{Ado.NoCM}")
        .usNamaPasien.SetUnboundFieldSource ("{Ado.Nama Pasien}")
        .usJK.SetUnboundFieldSource ("{Ado.JK}")
        .usAlamat.SetUnboundFieldSource ("{Ado.Alamat}")
        .udTglMasuk.SetUnboundFieldSource ("{Ado.TglMasuk}")
        .udTglKeluar.SetUnboundFieldSource ("{Ado.TglKeluar}")
        .usCaraKeluar.SetUnboundFieldSource ("{Ado.Cara Keluar}")
        .usRuanganTujuan.SetUnboundFieldSource ("{Ado.RuanganTujuan}")
        .udTglPulang.SetUnboundFieldSource ("{Ado.TglPulang}")
        .usKondisiPulang.SetUnboundFieldSource ("{Ado.Kondisi Pulang}")
        .usCaraPulang.SetUnboundFieldSource ("{Ado.Cara Pulang}")
        .usJenisPasien.SetUnboundFieldSource ("{Ado.JenisPasien}")
        .usKelas.SetUnboundFieldSource ("{Ado.Kelas}")
        .usRuangan.SetUnboundFieldSource ("{ado.Ruangan}")
        .usDiagnosa.SetUnboundFieldSource ("{ado.DiagnosaUtama}")
        
        .usThn.SetUnboundFieldSource ("{ado.Thn}")
        .usBln.SetUnboundFieldSource ("{ado.Bln}")
        .usHr.SetUnboundFieldSource ("{ado.Hr}")
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .EnableGroupTree = True
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
    Set frmCetakDaftarPasienLama = Nothing
End Sub

