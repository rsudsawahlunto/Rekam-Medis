VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakCatatanMedis 
   Caption         =   "Medifirst2000 - Cetak Catatan Medis"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12585
   Icon            =   "frmCetakCatatanMedis.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   12585
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7965
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12525
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
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakCatatanMedis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crCetakCatatanMedis

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    Call openConnection
    Set frmCetakCatatanMedis = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = "SELECT * FROM V_CetakRiwayatMedikPasienRJ " _
    & "WHERE NoPendaftaran='" & mstrNoPen & "'"
    adocomd.CommandType = adCmdText

    strSQL = "SELECT * FROM V_CetakCatatanMedikPasien WHERE NoCM='" & mstrNoCM & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

' edit by Dimas 20140508
    With Report
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        .txtNoPen.SetText mstrNoPen
        .txtNoCM.SetText mstrNoCM
        .txtRuang.SetText mstrNamaRuangan
        If IsNull(rs("Nama Pasien").value) = False Then .txtNama.SetText rs("Nama Pasien").value
        .txtTmpLahir.SetText rs("TempatLahir").value & "," & " " & Format(rs("TglLahir").value, "dd MMMM yyyy")
'        If IsNull(rs("TglLahir").value) = False Then .txtTglLahir.SetText rs("TglLahir").value
        If IsNull(rs("JK").value) = False Then
            If rs("JK").value = "P" Then
                .txtL.Font.Strikethrough = True
            ElseIf rs("JK").value = "L" Then
                .txtP.Font.Strikethrough = True
            End If
        End If
        If IsNull(rs("Nama Keluarga").value) = False Then .txtKeluarga.SetText rs("Nama Keluarga").value
        If IsNull(rs("Bin").value) = False Then .txtBin.SetText rs("Bin").value
        If IsNull(rs("Pekerjaan").value) = False Then .TxtPekerjaan.SetText rs("Pekerjaan").value
        If IsNull(rs("Alamat").value) = False Then .txtAlamat.SetText rs("Alamat").value
        If IsNull(rs("RTRW").value) = False Then .txtRTRW.SetText rs("RTRW").value
        If IsNull(rs("Kelurahan").value) = False Then .txtKelurahan.SetText rs("Kelurahan").value
        If IsNull(rs("Kecamatan").value) = False Then .txtKecamatan.SetText rs("Kecamatan").value
        If IsNull(rs("Kota").value) = False Then .txtKota.SetText rs("Kota").value
        .Database.AddADOCommand dbConn, adocomd
'        .udtgl.SetUnboundFieldSource ("{Ado.TglPeriksa}")
'        .usRuanganPemeriksaan.SetUnboundFieldSource ("{Ado.RuangPemeriksaan}")
'        .usKeluhanUtama.SetUnboundFieldSource ("{Ado.KeluhanUtama}")
'        .usDiagnosa.SetUnboundFieldSource ("{Ado.Diagnosa}")
'        .usPengobatan.SetUnboundFieldSource ("{Ado.Pengobatan}")
'        .usKeterangan.SetUnboundFieldSource ("{Ado.Keterangan}")

        .UnTanggal.SetUnboundFieldSource ("{ado.TglPeriksa}")
        .UnRuangPemeriksa.SetUnboundFieldSource ("{ado.RuangPemeriksaan}")
        .UnKeluahan.SetUnboundFieldSource ("{ado.KeluhanUtama}")
        .UnDiagnosa.SetUnboundFieldSource ("{ado.Diagnosa}")
        .UnPengobatan.SetUnboundFieldSource ("{ado.Pengobatan}")
        .UnKeterangan.SetUnboundFieldSource ("{ado.keterangan}")
'        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
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
    Set frmCetakCatatanMedis = Nothing
End Sub



