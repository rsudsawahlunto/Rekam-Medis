VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakHasilLabPA 
   Caption         =   "crhasilrad"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   Icon            =   "frmCetakHasilLabPA.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4635
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7005
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   5805
      DisplayGroupTree=   0   'False
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
      EnableAnimationControl=   0   'False
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
Attribute VB_Name = "frmCetakHasilLabPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crHasilLabPA

Private Sub Form_Load()
    Dim strNamaNamaDiagnosa As String
    Set frmCetakHasilLabPA = Nothing
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    Dim adocomd As New ADODB.Command
    Call openConnection

    strSQL = "SELECT DISTINCT Diagnosa FROM  V_JudulCetakHasilPeriksaLaboratoriumPA WHERE (NoLaboratorium = '" & mstrNoLabRad & "') AND (NOT (Diagnosa IS NULL))"
    Call msubRecFO(rs, strSQL)
    strNamaNamaDiagnosa = ""
    While rs.EOF = False
        strNamaNamaDiagnosa = strNamaNamaDiagnosa & IIf(IsNull(rs("Diagnosa")), "", rs("Diagnosa")) & ", "
        rs.MoveNext
    Wend

    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = "SELECT * " & _
    " from V_CetakHasilPeriksaLaboratoryPA " & _
    " WHERE NoLaboratorium = '" & mstrNoLabRad & "'"
    adocomd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, adocomd

    With Report
        .NoLab.SetUnboundFieldSource ("{ado.NoLaboratorium}")
        .NoPendaftaran.SetUnboundFieldSource ("{ado.NoPendaftaran}")
        .NoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .udTglPendaftaran.SetUnboundFieldSource ("{ado.TglHasil}")

        .NamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
        .Umur.SetUnboundFieldSource ("{ado.Umur}")
        .usJK.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .usAlamat.SetUnboundFieldSource ("{ado.AlamatLengkap}")

        .RuangPerujuk.SetUnboundFieldSource ("{ado.RuanganPerujuk}")
        .AsalPerujuk.SetUnboundFieldSource ("{ado.AsalPasien}")
        .usDokterPerujuk.SetUnboundFieldSource ("{ado.DokterPerujuk}")
        .txtDiagnosa.SetText strNamaNamaDiagnosa

        .usJenisPemeriksaan.SetUnboundFieldSource ("{ado.JenisPeriksa}")
        .usDetailPemeriksaan.SetUnboundFieldSource ("{ado.NamaDetailPeriksa}")
        .usNamaPemeriksaan.SetUnboundFieldSource ("{ado.NamaPelayanan}")
        .Hasil.SetUnboundFieldSource ("{ado.MemoHasilPeriksa}")

        .usKesimpulan.SetUnboundFieldSource ("{ado.Catatan}")
        .udTglHasil.SetUnboundFieldSource ("{ado.TglHasil}")

        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        .Text300.SetText "INSTALASI LABORATORIUM ANATOMI"
        .SelectPrinter sDriver, sPrinter, vbNull
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
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
    Set frmCetakHasilLabPA = Nothing
End Sub

