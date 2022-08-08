VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakDaftarPasienVerifikasiPemakaianAsuransi 
   Caption         =   "Medifirst2000 - Laporan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakDaftarPasienVerifikasiPemakaianAsuransi.frx":0000
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
Attribute VB_Name = "frmCetakDaftarPasienVerifikasiPemakaianAsuransi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crDaftarPasienVerifikasiPemakaianAsuransi

Private Sub Form_Load()
    On Error GoTo errLoad
    Set dbcmd = New ADODB.Command

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Me.Caption = "Medifirst2000 - Cetak Daftar Pasien Verifikasi Pemakaian Asuransi"

    strSQL = "SELECT  * FROM V_VerifikasiPemakaianAsuransiPasien" & _
    " WHERE TglPendaftaran BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' order by NoCM"

    Call msubRecFO(rs, strSQL)
    With Report
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strWebsite & ", " & strEmail
        .txtPeriode.SetText Format(mdTglAwal, "dd MMMM yyyy HH:mm") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm")

        Set dbcmd.ActiveConnection = dbConn
        dbcmd.CommandText = strSQL
        dbcmd.CommandType = adCmdText
        .Database.AddADOCommand dbConn, dbcmd

        .usRuanganPerawatan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .usNoPendaftaran.SetUnboundFieldSource ("{ado.NoPendaftaran}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
        .udtTglPendaftaran.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .usNamaPenjamin.SetUnboundFieldSource ("{ado.NamaPenjamin}")
    End With

    With CRViewer1
        .ReportSource = Report
        .EnableGroupTree = False
        .ViewReport
        .Zoom 100
    End With
    Screen.MousePointer = vbDefault

    Exit Sub
errLoad:
    Call msubPesanError
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    With CRViewer1
        .Top = 0
        .Left = 0
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakDaftarPasienVerifikasiPemakaianAsuransi = Nothing
End Sub

