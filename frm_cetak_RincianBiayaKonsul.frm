VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_cetak_RincianBiayaKonsul 
   Caption         =   "Cetak Laporan Cetak Rincian Biaya Sementara"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   Icon            =   "frm_cetak_RincianBiayaKonsul.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   5805
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
Attribute VB_Name = "frm_cetak_RincianBiayaKonsul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New cr_RincianBiayaKonsul
Dim adocomd As New ADODB.Command
Public sNamaKelas As String

Private Sub Form_Load()
On Error GoTo errLoad

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    
    If strCetak2 = "OA" Then
        Call subCetakPesanPelayananOA
    ElseIf strCetak2 = "TM" Then
        Call subCetakPesanPelayananTM
    End If
    
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
    Set rs = Nothing
End Sub

Private Sub subCetakPesanPelayananOA()
    strSQL = "select * " & _
        " from V_JudulRincianBiayaSementara where " & _
        " nopendaftaran ='" & mstrNoPen & "'  "
    Call msubRecFO(rs, strSQL)
    
    With Report
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        .Text37.SetText "Permintaan Resep/ Obat Alkes Pasien " & strNNamaRuangan & ""
        
        .txttglpendaftaran.SetText rs("TglPendaftaran")
        .txtNoCM.SetText rs("nocm")
        .txtnmpasien.SetText rs("nama pasien") & " / " & IIf(rs("JK").value = "P", "Wanita", "Pria")
        .txtumur.SetText rs("umur")
        .txtAlamat.SetText IIf(IsNull(rs("alamat")), "-", rs("alamat"))
        .txtBiasaCito.SetText "Biasa"
        .txtRuanganPengirim.SetText rs("RuanganTerakhir")
        .txtdokterperujuk.SetText frmKonsul_OrderPelayanan.dcNamaPerujuk.Text
        .txtPrintTglBKM.SetText strNKotaRS & ", " & Format(Now, "dd MMMM yyyy")
        .txtKelas.SetText sNamaKelas
    End With
    
    Set dbcmd = New ADODB.Command
    dbcmd.CommandText = "SELECT * FROM V_DaftarDetailOrderOA " _
                    & "WHERE (NoPendaftaran = '" & mstrNoPen & "') and KdRuanganTujuan ='" & mstrKdRuanganORS & "' AND " _
                    & "TglOrder between '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' and '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' "
    dbcmd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, dbcmd
    With Report
      .Field11.Suppress = True
      .unRKe.Suppress = False
      .Text38.SetText "RKe"
      .udtanggal.SetUnboundFieldSource ("{Ado.TglOrder}")
      .uskelas.SetUnboundFieldSource ("{Ado.kelas}")
      .unqty.SetUnboundFieldSource ("{Ado.JmlBarang}")
      .usNoOrder.SetUnboundFieldSource ("{Ado.NoOrder}")
      .ucbiayasatuan.SetUnboundFieldSource ("{Ado.BiayaSatuan}") '("{Ado.harga_item}")
      .unRKe.SetUnboundFieldSource ("{Ado.ResepKe}")
      
      .ucTarifTotal.SetUnboundFieldSource ("{Ado.BiayaTotal}")
      .ustindakan.SetUnboundFieldSource ("{Ado.NamaBarang}")

      strSQL = "SELECT SUM(BiayaTotal) As TotBiayaTotal FROM V_DaftarDetailOrderOA " _
                & "WHERE (NoPendaftaran = '" & mstrNoPen & "') and KdRuanganTujuan ='" & mstrKdRuanganORS & "' AND " _
                & "TglOrder between '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' and '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "'"
      Set rs = Nothing
      rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF = True Then
        .txtPembebasan.SetText 0
        .txtTanggunganRS.SetText 0
        .txtTotalBiaya.SetText 0
        .txtTanggungan.SetText 0
        .txtBayar.SetText 0
    Else
        .txtPembebasan.SetText 0
        .txtTotalBiaya.SetText IIf(rs("TotBiayaTotal").value = 0, 0, Format(rs("TotBiayaTotal").value, "#,###"))
        If IsNull(rs("TotBiayaTotal").value) Then
            .txtTerbilang.SetText NumToText(0)
        Else
            .txtTerbilang.SetText NumToText(IIf(rs("TotBiayaTotal").value = 0, 0, CCur(rs("TotBiayaTotal").value)))
        End If
    End If
    .txtPetugasKasir.SetText strNmPegawai
    .txtIdPetugas.SetText noidpegawai
    End With
End Sub


Private Sub subCetakPesanPelayananTM()

    strSQL = "select * " & _
        " from V_JudulRincianBiayaSementara where " & _
        " nopendaftaran ='" & mstrNoPen & "'  "
    Call msubRecFO(rs, strSQL)
    
    With Report
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        .Text37.SetText "Permintaan Pemeriksaan " & strNamaRuangan & ""
'        .txtdokterpemeriksa.SetText NamaDokterDituju
        
'        .txtnopendaftaran.SetText rs("nopendaftaran")
        .txttglpendaftaran.SetText rs("TglPendaftaran")
        .txtNoCM.SetText rs("nocm")
        .txtnmpasien.SetText rs("nama pasien") & " / " & IIf(rs("JK").value = "P", "Wanita", "Pria")
        .txtumur.SetText rs("umur")
        .txttanggallahir.SetText rs("TglLahir")
        .txtAlamat.SetText IIf(IsNull(rs("alamat")), "-", rs("alamat"))
        .txtdokterperujuk.SetText frmKonsul_OrderPelayanan.dcNamaPerujuk.Text
        If strNStsCITO = "Tidak" Then
            .txtBiasaCito.SetText "Biasa"
        ElseIf strNStsCITO = "Ya" Then
            .txtBiasaCito.SetText "Cito"
        End If
'        .txtklpkpasien.SetText rs("jenispasien")
'        .txtPenjamin.SetText IIf(IsNull(rs("NamaPenjamin")), "Sendiri", rs("NamaPenjamin"))
'        .txtNamaRuangan.SetText mstrNamaRuangan
        .txtRuanganPengirim.SetText rs("RuanganTerakhir")
      '  .txtdokterperujuk.SetText mstrNama
'        .txttanggallahir.SetText rs("TglLahir")
        .txtPrintTglBKM.SetText strNKotaRS & ", " & Format(Now, "dd MMMM yyyy")
        .txtKelas.SetText rs("deskKelas")
'        .txtHeader.SetText rs("nocm") & "/"
'        .txtFooter.SetText rs("nocm") & "/"
'        .txtNoKartu.SetText IIf(IsNull(rs("IdPenjamin")), "-", rs("IdPenjamin"))
'        .txtMasaBerlaku.SetText IIf(IsNull(rs("tglBerlaku")), "-", rs("tglBerlaku"))
    End With
    Set dbcmd = New ADODB.Command
    
    If strCetak = "1" Then
        dbcmd.CommandText = "SELECT * FROM V_RincianTotalDetailBiayaPelayanan " _
        & "WHERE (NoPendaftaran = '" & mstrNoPen & "') AND TglPelayanan between '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' and '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' and KdRuangan='" & mstrKdRuanganORS & "' "

        dbcmd.CommandType = adCmdText
        Report.Database.AddADOCommand dbConn, dbcmd
        With Report
          .Field11.Suppress = False
          .unRKe.Suppress = True
          .Text38.SetText "No."
          .udtanggal.SetUnboundFieldSource ("{Ado.TglPelayanan}")
          .usruang.SetUnboundFieldSource ("{Ado.ruangan}")
          .usjenispelayanan.SetUnboundFieldSource ("{Ado.jenis_item}")
          .uskelas.SetUnboundFieldSource ("{Ado.kelas}")
          .unqty.SetUnboundFieldSource ("{Ado.Jml_Item}")
          .usNoOrder.SetUnboundFieldSource ("{Ado.NoLab_Rad}")
          .ucbiayasatuan.SetUnboundFieldSource ("{Ado.Harga_Item}") '("{Ado.harga_item}")
          '.ucTarifCITO.SetUnboundFieldSource ("{Ado.TarifCITO}")
          .ucTarifTotal.SetUnboundFieldSource ("{Ado.BiayaTotal}")
          .ustindakan.SetUnboundFieldSource ("{Ado.Nama_Item}")
          .usruangantujuan.SetUnboundFieldSource ("{Ado.Ruangan}")

        strSQL = "SELECT SUM(BiayaTotal) As TotBiayaTotal FROM V_RincianTotalDetailBiayaPelayanan " _
          & "WHERE (NoPendaftaran = '" & mstrNoPen & "') and TglPelayanan between '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' and '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' and KdRuangan='" & mstrKdRuanganORS & "' "

          Set rs = Nothing
          rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.EOF = True Then
            .txtPembebasan.SetText 0
            .txtTanggunganRS.SetText 0
            .txtTotalBiaya.SetText 0
            .txtTanggungan.SetText 0
            .txtBayar.SetText 0
        Else
            .txtPembebasan.SetText 0
            .txtTotalBiaya.SetText IIf(rs("TotBiayaTotal").value = 0, 0, Format(rs("TotBiayaTotal").value, "#,###"))
            If IsNull(rs("TotBiayaTotal").value) Then
                .txtTerbilang.SetText NumToText(0)
            Else
                .txtTerbilang.SetText NumToText(IIf(rs("TotBiayaTotal").value = 0, 0, CCur(rs("TotBiayaTotal").value)))
            End If
        End If
        .txtPetugasKasir.SetText strNmPegawai
        .txtIdPetugas.SetText noidpegawai
       End With
    ElseIf strCetak = "0" Then
      dbcomm.CommandText = "SELECT * FROM V_DetailOrderTM " _
        & "WHERE (NoPendaftaran = '" & mstrNoPen & "') and KdRuanganTujuan ='" & mstrKdRuanganORS & "' AND KdKelas = '" & TempKodeKelas & "' AND TglOrder between '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' and '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' "
        dbcomm.CommandType = adCmdText
        Report.Database.AddADOCommand dbConn, dbcomm
        With Report
          .udtanggal.SetUnboundFieldSource ("{Ado.TglOrder}")
'          .usruang.SetUnboundFieldSource ("{Ado.ruangan}")
'          .usjenispelayanan.SetUnboundFieldSource ("{Ado.jenis_item}")
'          .usHeader.SetUnboundFieldSource ("{Ado.header}")
'          .usFooter.SetUnboundFieldSource ("{Ado.header}")
'          .uskelas.SetUnboundFieldSource ("{Ado.kelas}")
          .unqty.SetUnboundFieldSource ("{Ado.JmlPelayanan}")
'          .UsNoOrder.SetUnboundFieldSource ("{Ado.NoOrder}")
          .ucbiayasatuan.SetUnboundFieldSource ("{Ado.BiayaSatuan}") '("{Ado.harga_item}")
          '.ucTarifCITO.SetUnboundFieldSource ("{Ado.TarifCITO}")
          .ucTarifTotal.SetUnboundFieldSource ("{Ado.BiayaTotal}")
          .ustindakan.SetUnboundFieldSource ("{Ado.NamaPelayanan}")
'          .usruangantujuan.SetUnboundFieldSource ("{Ado.Ruangan}")
          strSQL = "SELECT SUM(BiayaTotal) As TotBiayaTotal FROM V_DetailOrderTM " _
          & "WHERE (NoPendaftaran = '" & mstrNoPen & "') and KdRuanganTujuan ='" & mstrKdRuanganORS & "' and KdKelas ='" & TempKodeKelas & "' AND TglOrder between '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' and '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "'"
          Set rs = Nothing
          rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.EOF = True Then
            .txtPembebasan.SetText 0
            .txtTanggunganRS.SetText 0
            .txtTotalBiaya.SetText 0
            .txtTanggungan.SetText 0
            .txtBayar.SetText 0
        Else
            .txtPembebasan.SetText 0
            .txtTotalBiaya.SetText IIf(rs("TotBiayaTotal").value = 0, 0, Format(rs("TotBiayaTotal").value, "#,###"))
            If IsNull(rs("TotBiayaTotal").value) Then
                .txtTerbilang.SetText NumToText(0)
            Else
                .txtTerbilang.SetText NumToText(IIf(rs("TotBiayaTotal").value = 0, 0, CCur(rs("TotBiayaTotal").value)))
            End If
        End If
        .txtPetugasKasir.SetText strNmPegawai
        .txtIdPetugas.SetText noidpegawai
        End With
    End If
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call frmDaftarPasienRJ.PostingHutangPenjaminPasien_AU("U")
    Set frm_cetak_RincianBiayaKonsul = Nothing
    Set rs = Nothing

End Sub
