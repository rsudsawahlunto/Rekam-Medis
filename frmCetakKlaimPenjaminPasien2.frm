VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakKlaimPenjaminPasien2 
   Caption         =   "Cetak Tagihan Biaya Pengobatan / Perawatan"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   Icon            =   "frmCetakKlaimPenjaminPasien2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   6915
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
      EnableRefreshButton=   -1  'True
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
Attribute VB_Name = "frmCetakKlaimPenjaminPasien2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReportTagihanPerawatanRJ As New crTagihanBiayaPengobatan_Perawatan_JnsPasien

Private Sub Form_Load()
    Dim adocmd As New ADODB.Command

    On Error GoTo errLoad

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    With ReportTagihanPerawatanRJ
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        .txtPeriode.SetText Format(FrmDaftarTagihanMaskin2.DTPickerAwal, "dd MMMM yyyy HH:mm") & " s/d " & Format(FrmDaftarTagihanMaskin2.DTPickerAkhir, "dd MMMM yyyy HH:mm")

        adocmd.CommandText = strSQL
        adocmd.CommandType = adCmdText

        .Database.AddADOCommand dbConn, adocmd

        .usPenjamin.SetUnboundFieldSource ("{ado.Penjamin}")
        .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usCost.SetUnboundFieldSource ("{ado.NoRujukan}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
        .udTglPendaftaran.SetUnboundFieldSource ("{ado.TglBKM}")
        .usNamaPeserta.SetUnboundFieldSource ("{ado.NamaPeserta}")
        .usIDPeserta.SetUnboundFieldSource ("{ado.IDPeserta}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .usUnitBagian.SetUnboundFieldSource ("{ado.UnitBagian}")
        .usJenisPelayanan.SetUnboundFieldSource ("{ado.JenisPelayanan}")
        .usNamaPelayanan.SetUnboundFieldSource ("{ado.NamaPelayanan}")
        .ucBiaya.SetUnboundFieldSource ("{ado.TotalHutangPenjamin}")
    End With

    With CRViewer1
        .ReportSource = ReportTagihanPerawatanRJ
        .EnableGroupTree = False
        .Zoom 1
        .ViewReport
    End With

    Screen.MousePointer = vbDefault
    Exit Sub

errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    If vLaporan = "Print" Then frmCetakKlaimPenjaminPasien2.Hide
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakKlaimPenjaminPasien2 = Nothing
End Sub

