VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakIndexDiagnosaPasien 
   Caption         =   "Cetak index diagnosa pasien"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakIndexDiagnosaPasien.frx":0000
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
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakIndexDiagnosaPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rptIndexDiagnosaPasien As New crIndexDiagnosaPasien

Private Sub Form_Load()
    On Error GoTo errLoad

    Dim adocomd As New ADODB.Command

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    With rptIndexDiagnosaPasien
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strWebsite & ", " & strEmail

        If Format(mdTglAwal, "dd MMMM yyyy") = Format(mdTglAkhir, "dd MMMM yyyy") Then
            .txtTanggal.SetText "Tanggal Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy")
        Else
            .txtTanggal.SetText "Periode Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") & " S/d " & Format(mdTglAkhir, "dd MMMM yyyy")
        End If

        Set adocomd.ActiveConnection = dbConn
        strSQL = "SELECT * " & _
        " From V_IndexDiagnosaPasien" & _
        " WHERE TglPeriksa BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' AND (NamaDiagnosa LIKE '%" & frmPeriodeLaporanIndexDiagnosaPasien.txtNamaDiagnosa.Text & "%' OR NamaDiagnosa IS NULL) AND (Alamat LIKE '%" & frmPeriodeLaporanIndexDiagnosaPasien.txtAlamatPasien.Text & "%' OR Alamat IS NULL)"

        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdUnknown
        .Database.AddADOCommand dbConn, adocomd

        .usNamaDiagnosa.SetUnboundFieldSource ("{ado.NamaDiagnosa}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .usUmur.SetUnboundFieldSource ("{ado.Umur}")
        .usRuanganPemeriksaan.SetUnboundFieldSource ("{ado.RuanganPemeriksaan}")
        .udTglPeriksa.SetUnboundFieldSource ("{ado.TglPeriksa}")
        .usNamaDokter.SetUnboundFieldSource ("{ado.DokterPemeriksa}")
        .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")

        settingreport rptIndexDiagnosaPasien, sPrinter, sDriver, sUkuranKertas, sDuplex, crPortrait
    End With

    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = rptIndexDiagnosaPasien
            .ViewReport
            .Zoom (100)
        End With
    Else
        rptIndexDiagnosaPasien.PrintOut False
        Unload Me
    End If
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
    Set frmCetakIndexDiagnosaPasien = Nothing
End Sub

