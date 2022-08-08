VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_cetak_RincianBiayaJasaDokter 
   Caption         =   "Cetak Rincian Biaya Pelayanan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frm_cetak_RincianBiayaJasaDokter.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
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
Attribute VB_Name = "frm_cetak_RincianBiayaJasaDokter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New cr_RincianBiayaJasaDokter

Private Sub Form_Load()
On Error GoTo errLoad
Dim adocomd As New ADODB.Command

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    
If frmLaporanJasaDokter.optTotalBiaya.Value = True Then 'non Penjamin

    If frmLaporanJasaDokter.optBelumBayar.Value = True Then
    
        strSQLx = "INSERT INTO RekapKomponenBPRemunerasiTM_JasaDokter " & _
                    " SELECT *, '" & strNamaHostLocal & "', '" & Format(TglCetak, "yyyy/MM/dd HH:mm:ss") & "' " & _
                    " FROM V_RekapKomponenBPRemunerasiTM_JasaDokter " & _
                    " WHERE TotalBayar >0 and Tunai = 'Y' and NoBKK is null and IdPegawai like '%" & frmLaporanJasaDokter.dcNamaDokter.BoundText & "%' AND year(TglBKM) between '" & Year(frmLaporanJasaDokter.dtpawal.Value) & "' AND '" & Year(frmLaporanJasaDokter.dtpakhir.Value) & "' and month(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "MM") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "MM") & "' and day(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "dd") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "dd") & "' AND KdDetailKomponenR = '01' AND KdRuanganKasir <> '178'" & _
                    "" & sKdJenisPasien & "" & _
                    "" & sIdPenjamin & ""

'        dbConn.CommandTimeout = 0
        dbConn.Execute strSQLx
    
        strSQL = "SELECT  DokterPemeriksa, JenisPasien, PenjaminPasien, sum(TotalBayar) as TotalBiaya " & _
        " FROM  RekapKomponenBPRemunerasiTM_JasaDokter " & _
        " WHERE TotalBayar >0 and Tunai = 'Y' and NoBKK is null and IdPegawai like '%" & frmLaporanJasaDokter.dcNamaDokter.BoundText & "%' AND year(TglBKM) between '" & Year(frmLaporanJasaDokter.dtpawal.Value) & "' AND '" & Year(frmLaporanJasaDokter.dtpakhir.Value) & "' and month(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "MM") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "MM") & "' and day(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "dd") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "dd") & "' AND KdDetailKomponenR = '01' AND KdRuanganKasir <> '178' AND NamaKomputer ='" & strNamaHostLocal & "' AND TglCetak='" & Format(TglCetak, "yyyy/MM/dd HH:mm:ss") & "'" & _
        "" & sKdJenisPasien & "" & _
        "" & sIdPenjamin & "" & _
        " GROUP BY DokterPemeriksa, JenisPasien, PenjaminPasien"
    End If
        
    If frmLaporanJasaDokter.optSudahBayar.Value = True Then
    
        strSQLx = "INSERT INTO RekapKomponenBPRemunerasiTM_JasaDokter " & _
                    " SELECT *, '" & strNamaHostLocal & "', '" & Format(TglCetak, "yyyy/MM/dd HH:mm:ss") & "' " & _
                    " FROM V_RekapKomponenBPRemunerasiTM_JasaDokter " & _
                    " WHERE TotalBayar >0 and Tunai = 'Y' and NoBKK is null and IdPegawai like '%" & frmLaporanJasaDokter.dcNamaDokter.BoundText & "%' AND year(TglBKM) between '" & Year(frmLaporanJasaDokter.dtpawal.Value) & "' AND '" & Year(frmLaporanJasaDokter.dtpakhir.Value) & "' and month(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "MM") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "MM") & "' and day(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "dd") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "dd") & "' AND KdDetailKomponenR = '01' AND KdRuanganKasir <> '178'" & _
                    "" & sKdJenisPasien & "" & _
                    "" & sIdPenjamin & ""

'        dbConn.CommandTimeout = 0
        dbConn.Execute strSQLx
    
        strSQL = "SELECT DokterPemeriksa, JenisPasien, PenjaminPasien, sum(TotalBayar) as TotalBiaya " & _
                 "FROM  RekapKomponenBPRemunerasiTM_JasaDokter " & _
        " WHERE TotalBayar >0 and Tunai = 'Y' and NoBKK is not null and IdPegawai like '%" & frmLaporanJasaDokter.dcNamaDokter.BoundText & "%' AND year(TglBKM) between '" & Year(frmLaporanJasaDokter.dtpawal.Value) & "' AND '" & Year(frmLaporanJasaDokter.dtpakhir.Value) & "' and month(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "MM") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "MM") & "' and day(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "dd") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "dd") & "' AND KdDetailKomponenR = '01' AND KdRuanganKasir <> '178' AND NamaKomputer ='" & strNamaHostLocal & "' AND TglCetak='" & Format(TglCetak, "yyyy/MM/dd HH:mm:ss") & "'" & _
        "" & sKdJenisPasien & "" & _
        "" & sIdPenjamin & "" & _
        " GROUP BY DokterPemeriksa, JenisPasien, PenjaminPasien"
    End If
End If

If frmLaporanJasaDokter.optHutangPenjamin.Value = True Then 'penjamin yang sudah cair
    If frmLaporanJasaDokter.optBelumBayar.Value = True Then
    
        strSQLx = "INSERT INTO RekapKomponenBPRemunerasiTM_JasaDokter " & _
                    " SELECT *, '" & strNamaHostLocal & "', '" & Format(TglCetak, "yyyy/MM/dd HH:mm:ss") & "' " & _
                    " FROM V_RekapKomponenBPRemunerasiTM_JasaDokter " & _
                    " WHERE TotalHutangPenjamin >0 and Tunai = 'X' and NoBKK is null and IdPegawai like '%" & frmLaporanJasaDokter.dcNamaDokter.BoundText & "%' AND year(TglBKM) between '" & Year(frmLaporanJasaDokter.dtpawal.Value) & "' AND '" & Year(frmLaporanJasaDokter.dtpakhir.Value) & "' and month(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "MM") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "MM") & "' and day(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "dd") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "dd") & "' AND KdDetailKomponenR = '01' AND KdRuanganKasir <> '178'" & _
                    "" & sKdJenisPasien & "" & _
                    "" & sIdPenjamin & ""

'        dbConn.CommandTimeout = 0
        dbConn.Execute strSQLx
        
        strSQL = "SELECT  DokterPemeriksa, JenisPasien, PenjaminPasien, sum(TotalHutangPenjamin) as TotalBiaya " & _
        " FROM  RekapKomponenBPRemunerasiTM_JasaDokter " & _
        " WHERE TotalHutangPenjamin > 0 and Tunai = 'X' and NoBKK is null and IdPegawai like '%" & frmLaporanJasaDokter.dcNamaDokter.BoundText & "%' AND year(TglBKM) between '" & Year(frmLaporanJasaDokter.dtpawal.Value) & "' AND '" & Year(frmLaporanJasaDokter.dtpakhir.Value) & "' and month(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "MM") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "MM") & "' and day(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "dd") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "dd") & "' AND KdDetailKomponenR = '01' AND KdRuanganKasir <> '178' AND NamaKomputer ='" & strNamaHostLocal & "' AND TglCetak='" & Format(TglCetak, "yyyy/MM/dd HH:mm:ss") & "'" & _
        "" & sKdJenisPasien & "" & _
        "" & sIdPenjamin & "" & _
        " GROUP BY DokterPemeriksa, JenisPasien, PenjaminPasien"
    End If

    If frmLaporanJasaDokter.optSudahBayar.Value = True Then
    
        strSQLx = "INSERT INTO RekapKomponenBPRemunerasiTM_JasaDokter " & _
                    " SELECT *, '" & strNamaHostLocal & "', '" & Format(TglCetak, "yyyy/MM/dd HH:mm:ss") & "' " & _
                    " FROM V_RekapKomponenBPRemunerasiTM_JasaDokter " & _
                    " WHERE TotalHutangPenjamin >0 and Tunai = 'X' and NoBKK is null and IdPegawai like '%" & frmLaporanJasaDokter.dcNamaDokter.BoundText & "%' AND year(TglBKM) between '" & Year(frmLaporanJasaDokter.dtpawal.Value) & "' AND '" & Year(frmLaporanJasaDokter.dtpakhir.Value) & "' and month(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "MM") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "MM") & "' and day(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "dd") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "dd") & "' AND KdDetailKomponenR = '01' AND KdRuanganKasir <> '178'" & _
                    "" & sKdJenisPasien & "" & _
                    "" & sIdPenjamin & ""

'        dbConn.CommandTimeout = 0
        dbConn.Execute strSQLx
        
        strSQL = "SELECT DokterPemeriksa, JenisPasien, PenjaminPasien, sum(TotalHutangPenjamin) as TotalBiaya " & _
                 "FROM  RekapKomponenBPRemunerasiTM_JasaDokter " & _
        " WHERE TotalHutangPenjamin > 0 and Tunai = 'X' and NoBKK is not null and IdPegawai like '%" & frmLaporanJasaDokter.dcNamaDokter.BoundText & "%' AND year(TglBKM) between '" & Year(frmLaporanJasaDokter.dtpawal.Value) & "' AND '" & Year(frmLaporanJasaDokter.dtpakhir.Value) & "' and month(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "MM") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "MM") & "' and day(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "dd") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "dd") & "' AND KdDetailKomponenR = '01' AND KdRuanganKasir <> '178' AND NamaKomputer ='" & strNamaHostLocal & "' AND TglCetak='" & Format(TglCetak, "yyyy/MM/dd HH:mm:ss") & "'" & _
        "" & sKdJenisPasien & "" & _
        "" & sIdPenjamin & "" & _
        " GROUP BY DokterPemeriksa, JenisPasien, PenjaminPasien"
    End If
End If

If frmLaporanJasaDokter.optTRS.Value = True Then 'TRS dari validasi kasir
    If frmLaporanJasaDokter.optBelumBayar.Value = True Then
    
        strSQLx = "INSERT INTO RekapKomponenBPRemunerasiTM_JasaDokter " & _
                    " SELECT *, '" & strNamaHostLocal & "', '" & Format(TglCetak, "yyyy/MM/dd HH:mm:ss") & "' " & _
                    " FROM V_RekapKomponenBPRemunerasiTM_JasaDokter " & _
                    " WHERE TotalTanggunganRS >0 and Tunai = 'Y' and NoBKK is null and IdPegawai like '%" & frmLaporanJasaDokter.dcNamaDokter.BoundText & "%' AND year(TglBKM) between '" & Year(frmLaporanJasaDokter.dtpawal.Value) & "' AND '" & Year(frmLaporanJasaDokter.dtpakhir.Value) & "' and month(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "MM") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "MM") & "' and day(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "dd") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "dd") & "' AND KdDetailKomponenR = '01' AND KdRuanganKasir <> '178'" & _
                    "" & sKdJenisPasien & "" & _
                    "" & sIdPenjamin & ""

'        dbConn.CommandTimeout = 0
        dbConn.Execute strSQLx
    
        strSQL = "SELECT  DokterPemeriksa, JenisPasien, PenjaminPasien, sum(TotalTanggunganRS) as TotalBiaya " & _
        " FROM  RekapKomponenBPRemunerasiTM_JasaDokter " & _
        " WHERE TotalTanggunganRS > 0 and Tunai = 'Y' and NoBKK is null and IdPegawai like '%" & frmLaporanJasaDokter.dcNamaDokter.BoundText & "%' AND year(TglBKM) between '" & Year(frmLaporanJasaDokter.dtpawal.Value) & "' AND '" & Year(frmLaporanJasaDokter.dtpakhir.Value) & "' and month(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "MM") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "MM") & "' and day(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "dd") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "dd") & "' AND KdDetailKomponenR = '01' AND KdRuanganKasir <> '178' AND NamaKomputer ='" & strNamaHostLocal & "' AND TglCetak='" & Format(TglCetak, "yyyy/MM/dd HH:mm:ss") & "'" & _
        "" & sKdJenisPasien & "" & _
        "" & sIdPenjamin & "" & _
        " GROUP BY DokterPemeriksa, JenisPasien, PenjaminPasien"
    End If
        

    If frmLaporanJasaDokter.optSudahBayar.Value = True Then
    
        strSQLx = "INSERT INTO RekapKomponenBPRemunerasiTM_JasaDokter " & _
                    " SELECT *, '" & strNamaHostLocal & "', '" & Format(TglCetak, "yyyy/MM/dd HH:mm:ss") & "' " & _
                    " FROM V_RekapKomponenBPRemunerasiTM_JasaDokter " & _
                    " WHERE TotalTanggunganRS >0 and Tunai = 'Y' and NoBKK is null and IdPegawai like '%" & frmLaporanJasaDokter.dcNamaDokter.BoundText & "%' AND year(TglBKM) between '" & Year(frmLaporanJasaDokter.dtpawal.Value) & "' AND '" & Year(frmLaporanJasaDokter.dtpakhir.Value) & "' and month(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "MM") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "MM") & "' and day(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "dd") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "dd") & "' AND KdDetailKomponenR = '01' AND KdRuanganKasir <> '178'" & _
                    "" & sKdJenisPasien & "" & _
                    "" & sIdPenjamin & ""

'        dbConn.CommandTimeout = 0
        dbConn.Execute strSQLx
    
        strSQL = "SELECT DokterPemeriksa, JenisPasien, PenjaminPasien, sum(TotalTanggunganRS) as TotalBiaya " & _
                 "FROM  RekapKomponenBPRemunerasiTM_JasaDokter " & _
        " WHERE TotalTanggunganRS > 0 and Tunai = 'Y' and NoBKK is not null and IdPegawai like '%" & frmLaporanJasaDokter.dcNamaDokter.BoundText & "%' AND year(TglBKM) between '" & Year(frmLaporanJasaDokter.dtpawal.Value) & "' AND '" & Year(frmLaporanJasaDokter.dtpakhir.Value) & "' and month(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "MM") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "MM") & "' and day(TglBKM) BETWEEN '" & Format(frmLaporanJasaDokter.dtpawal.Value, "dd") & "' AND '" & Format(frmLaporanJasaDokter.dtpakhir.Value, "dd") & "' AND KdDetailKomponenR = '01' AND KdRuanganKasir <> '178' AND NamaKomputer ='" & strNamaHostLocal & "' AND TglCetak='" & Format(TglCetak, "yyyy/MM/dd HH:mm:ss") & "'" & _
        "" & sKdJenisPasien & "" & _
        "" & sIdPenjamin & "" & _
        " GROUP BY DokterPemeriksa, JenisPasien, PenjaminPasien"
    End If
End If
    openConnection
    'Set adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdUnknown
    
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        
        If frmLaporanJasaDokter.optTotalBiaya.Value = True Then
            If frmLaporanJasaDokter.optBelumBayar.Value = True Then .txtKriteria.SetText "Total Bayar & Belum dibayar Pasien"
            If frmLaporanJasaDokter.optSudahBayar.Value = True Then .txtKriteria.SetText "Total Bayar & Sudah dibayar Pasien"
        End If
        If frmLaporanJasaDokter.optHutangPenjamin.Value = True Then
            If frmLaporanJasaDokter.optBelumBayar.Value = True Then .txtKriteria.SetText "Hutang Penjamin & Klaim Belum dibayar"
            If frmLaporanJasaDokter.optSudahBayar.Value = True Then .txtKriteria.SetText "Hutang Penjamin & Klaim Cair"
        End If
        If frmLaporanJasaDokter.optTRS.Value = True Then
            If frmLaporanJasaDokter.optBelumBayar.Value = True Then .txtKriteria.SetText "Tanggungan RS & Belum divalidasi Kasir"
            If frmLaporanJasaDokter.optSudahBayar.Value = True Then .txtKriteria.SetText "Tanggungan RS & Sudah divalidasi Kasir"
        End If
        
        'If Format(frmLaporanJasaDokter.dtpawal, "MMMM yyyy") <> Format(frmLaporanJasaDokter.dtpakhir, "MMMM yyyy") Then
            .txtPeriode.SetText Format(frmLaporanJasaDokter.dtpawal, "dd MMMM yyyy") & " s/d " & Format(frmLaporanJasaDokter.dtpakhir, "dd MMMM yyyy")
'        Else
'            .txtPeriode.SetText "Bulan : " & Format(frmLaporanJasaDokter.dtpawal, "MMMM yyyy")
'        End If
        
        If frmLaporanJasaDokter.optBelumBayar.Value = True Then
            .txtJudul.SetText "RINCIAN KREDIT JASA MEDIS DOKTER"
        Else
            .txtJudul.SetText "RINCIAN JASA MEDIS DOKTER PIUTANG TERBAYAR"
        End If
        
        .usDokter.SetUnboundFieldSource ("{ado.DokterPemeriksa}")
        .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
        .usPenjamin.SetUnboundFieldSource ("{ado.PenjaminPasien}")
        .ucTarifTotal.SetUnboundFieldSource ("{ado.TotalBiaya}")
        
    End With
    
    CRViewer1.ReportSource = Report
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 1
    End With
    Screen.MousePointer = vbDefault
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
    Call HapusData
    Set frm_cetak_RincianBiayaJasaDokter = Nothing
End Sub

Public Sub HapusData()
On Error GoTo hell
    strSQLx = "Delete from RekapKomponenBPRemunerasiTM_JasaDokter " & _
                    " WHERE NamaKomputer='" & strNamaHostLocal & "' AND TglCetak= '" & Format(TglCetak, "yyyy/MM/dd HH:mm:ss") & "' "
                         
    dbConn.Execute strSQLx
    frmLaporanJasaDokter.cmdCetak.Enabled = True
        
Exit Sub
hell:
    Call msubPesanError
End Sub

