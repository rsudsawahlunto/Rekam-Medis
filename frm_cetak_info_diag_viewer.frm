VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_cetak_info_diag_viewer 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   Icon            =   "frm_cetak_info_diag_viewer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   5850
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
Attribute VB_Name = "frm_cetak_info_diag_viewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_info_diagnosa

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    Dim adocomd As New ADODB.Command
    Call openConnection

    adocomd.ActiveConnection = dbConn
    'adocomd.CommandText = "Select * from V_DaftarDiagnosaPasien where nocm = '" & mstrNoCM & "'"
    adocomd.CommandText = "Select DISTINCT NoPendaftaran, NoCM, TglPeriksa, JenisDiagnosa, KdDiagnosa, Diagnosa, [Ruang Periksa], [Dokter Pemeriksa], [Nama Pasien], JK, Umur ,dbo.Ambil_TindPenunjang(NoPendaftaran,'Pen') AS Penunjang" & _
                          ",dbo.Ambil_TindPenunjang(NoPendaftaran,'Obt') AS Obat " & _
                          ",dbo.Ambil_TindPenunjang(NoPendaftaran,'Rua') AS RI from V_DaftarDiagnosaPasien where nocm = '" & Right(mstrNoCM, 6) & "'"
    adocomd.CommandType = adCmdText

    Report.Database.AddADOCommand dbConn, adocomd

    With Report
        .usnocm.SetUnboundFieldSource ("{ado.nocm}")
        .usnama.SetUnboundFieldSource ("{Ado.nama pasien}")
        .usjeniskelamin.SetUnboundFieldSource ("{Ado.jk}")
        .udtgl.SetUnboundFieldSource ("{Ado.TglPeriksa}")
        .uskodeICD.SetUnboundFieldSource ("{Ado.Kddiagnosa}")
        .usdiagICD.SetUnboundFieldSource ("{Ado.diagnosa}")
        .usjenisdiagnosa.SetUnboundFieldSource ("{Ado.jenisdiagnosa}")
        .usdokter.SetUnboundFieldSource ("{ado.dokter pemeriksa}")
        .usruang.SetUnboundFieldSource ("{ado.ruang periksa}")
        .usumur.SetUnboundFieldSource ("{Ado.umur}")
       ' .usICD9.SetUnboundFieldSource ("{ado.DiagnosaTindakan}")
        .usPenunjang.SetUnboundFieldSource ("{ado.Penunjang}")
        .usObat.SetUnboundFieldSource ("{ado.Obat}")
        .usRI.SetUnboundFieldSource ("{ado.RI}")
        .Text1.SetText strNNamaRS
        .Text3.SetText strNAlamatRS
        .Text300.SetText "REKAM MEDIS"
        .Text400.SetText "TELP. " & strNTeleponRS & "  " & strNKotaRS & "  " & "Kode Pos " & " " & strNKodepos & " "
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_cetak_info_diag_viewer = Nothing
    mblnStatusCetakRD = False
End Sub
