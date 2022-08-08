VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakInfoPesanBarangNM 
   Caption         =   "m"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
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
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakInfoPesanBarangNM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rptInfoPesanBarang As New crInfoPesanBarangNM

Private Sub Form_Load()
Dim adocmd As New ADODB.Command
On Error GoTo Errload
Dim strS As String
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    
    Set adocmd.ActiveConnection = dbConn

    With rptInfoPesanBarang
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strWebsite & ", " & strEmail
        
        If strCetak = "Belum Diterima" Then
                    strS = "select * " & _
                    "from V_InfoPemesananBrgRuanganNM " & _
                    "where NoKirim is null and NoKonfirmasi is null and KdRuangan='" & mstrKdRuangan & "' " & _
                    "and([Tgl. Pesan] between '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' and '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "')"
            .Text6.SetText "Jumlah Pesan"
            .Text1.SetText "Tanggal Pesan"
        ElseIf strCetak = "Sudah Diterima" Then
            strS = "select * " & _
                    "from V_InfoPemesananBrgRuanganNMKirimx " & _
                    "where NoKirim is not null and NoKonfirmasi is null and KdRuangan='" & mstrKdRuangan & "' " & _
                    "and(TglKirim between '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' and '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "')"
            .Text6.SetText "Jumlah Terima"
            .Text1.SetText "Tanggal Kirim"
        Else
            strS = "select * " & _
                    "from V_InfoPemesananBrgRuanganNMKirimx " & _
                    "where NoKirim is not null and NoKonfirmasi is not null and KdRuangan='" & mstrKdRuangan & "' " & _
                    "and([Tgl. Pesan] between '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' and '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "')"
            .Text6.SetText "Jumlah Konfirmasi"
            .Text1.SetText "Tanggal Konfimasi"
        End If
        
    
        adocmd.CommandText = strS
        adocmd.CommandType = adCmdText
        .Database.AddADOCommand dbConn, adocmd
        
            .usTglOrder.SetUnboundFieldSource ("{ado.Tgl. Pesan}")
            .usNoOrder.SetUnboundFieldSource ("{ado.No. Pesan}")
            .usNamaRuangan.SetUnboundFieldSource ("{ado.Tujuan}")
'            .usDetailJenisBarang.SetUnboundFieldSource ("{ado.Jenis Barang}")
            .usNamaBarang.SetUnboundFieldSource ("{ado.Nama Barang}")
            If strCetak = "Belum Diterima" Then
            .usJumOrder.SetUnboundFieldSource ("{ado.Jml. Pesan}")
            ElseIf strCetak = "Sudah Diterima" Then
            .usJumOrder.SetUnboundFieldSource ("{ado.Jml. Pesan}")
            Else
            .usJumOrder.SetUnboundFieldSource ("{ado.JmlbarangKonfirmasi}")
            End If
            .usNamaLengkap.SetUnboundFieldSource ("{ado.Nama Pemesan}")
            .usNoKirim.SetUnboundFieldSource ("{ado.NoKirim}")
            .usNoReg.SetUnboundFieldSource ("{ado.NoRegisterAsset}")
        
    End With
        
        'settingreport rptInfoPesanBarang, sPrinter, sDriver, sUkuranKertas, sDuplex, crPortrait
    If vLaporan = "view" Then
        CRViewer1.ReportSource = rptInfoPesanBarang
        CRViewer1.Zoom 1
        CRViewer1.ViewReport
    Else
        rptInfoPesanBarang.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault

Exit Sub
Errload:
    Screen.MousePointer = vbDefault
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakInfoPesanBarangNM = Nothing
End Sub
