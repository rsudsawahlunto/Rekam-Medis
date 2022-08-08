VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmUtilitasRS 
   Caption         =   "Indikator Pelayanan Rawat Inap"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   Icon            =   "frmUtilitasRS.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   5985
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
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
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmUtilitasRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crUtilitasRS
Dim reportgrafik As New crUtilitasRSGrafik
Dim adoCommand As New ADODB.Command
Dim strSQL As String

Private Sub Form_Load()
    openConnection
    Set adoCommand.ActiveConnection = dbConn
    Set frmUtilitasRS = Nothing

    Dim tanggal As String
    If Format(FrmPeriodeIndikatorRS.DTPickerAwal, "dd MMMM yyyy") = Format(FrmPeriodeIndikatorRS.DTPickerAkhir, "dd MMMM yyyy") Then
        tanggal = "Tanggal Kunjungan  : " & " " & Format(FrmPeriodeIndikatorRS.DTPickerAwal, "dd MMMM yyyy") '& " S/d " & Format(FrmPeriodeIndikatorRS.DTPickerAkhir, "dd MMMM yyyy")
    Else
        tanggal = "Periode Kunjungan  : " & " " & Format(FrmPeriodeIndikatorRS.DTPickerAwal, "dd MMMM yyyy") & " S/d " & Format(FrmPeriodeIndikatorRS.DTPickerAkhir, "dd MMMM yyyy")
    End If

    If (FrmPeriodeIndikatorRS.cboKriteria = "Per Ruangan") Then
        strSQL = "SELECT NamaRuangan AS Ruangan,AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR, SUM(JmlPasien) AS JmlPasien " _
        & "FROM dbo.v_S_RekapIndikatorPlyn " _
        & "WHERE TglHitung BETWEEN '" & Format(FrmPeriodeIndikatorRS.DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(FrmPeriodeIndikatorRS.DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' " _
        & "GROUP BY NamaRuangan"
        strjudul = "Ruang Pelayanan"
        strKriteria = "PER RUANGAN"

        adoCommand.CommandText = strSQL
        adoCommand.CommandType = adCmdText

        Select Case cetak
            Case "PerRuangan"
                With Report
                    .Database.AddADOCommand dbConn, adoCommand
                    .txtPeriode.SetText tanggal
                    .txtKriteria.SetText strKriteria
                    .usRuangan.SetUnboundFieldSource ("{ado.Ruangan}")
                    .unTOI.SetUnboundFieldSource ("{ado.TOI}")
                    .unBOR.SetUnboundFieldSource ("{ado.BOR}")
                    .unBTO.SetUnboundFieldSource ("{ado.BTO}")
                    .unLOS.SetUnboundFieldSource ("{ado.LOS}")
                    .unGDR.SetUnboundFieldSource ("{ado.GDR}")
                    .unNDR.SetUnboundFieldSource ("{ado.NDR}")
                    .unJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")
                    .Text1.SetText strNNamaRS
                    .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
                    .Text3.SetText strWebsite & ", " & strEmail
                    .SelectPrinter sDriver, sPrinter, vbNull
                    settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
                End With

                CRViewer1.ReportSource = Report

            Case "GrafikPerRuangan"
                With reportgrafik
                    .Database.AddADOCommand dbConn, adoCommand
                    .txtPeriode.SetText tanggal
                    .txtKriteria.SetText strKriteria
                    .usRuangan.SetUnboundFieldSource ("{ado.Ruangan}")
                    .unTOI.SetUnboundFieldSource ("{ado.TOI}")
                    .unBOR.SetUnboundFieldSource ("{ado.BOR}")
                    .unBTO.SetUnboundFieldSource ("{ado.BTO}")
                    .unLOS.SetUnboundFieldSource ("{ado.LOS}")
                    .unGDR.SetUnboundFieldSource ("{ado.GDR}")
                    .unNDR.SetUnboundFieldSource ("{ado.NDR}")
                    .Text1.SetText strNNamaRS
                    .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
                    .Text3.SetText strWebsite & ", " & strEmail
                    .SelectPrinter sDriver, sPrinter, vbNull
                    settingreport reportgrafik, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
                End With
                CRViewer1.ReportSource = reportgrafik
        End Select
    ElseIf (FrmPeriodeIndikatorRS.cboKriteria = "Per Kelas") Then
        strSQL = "SELECT DeskKelas AS Kelas,AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR, SUM(JmlPasien) AS JmlPasien " _
        & "FROM dbo.v_S_RekapIndikatorPlyn " _
        & "WHERE TglHitung BETWEEN '" & Format(FrmPeriodeIndikatorRS.DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(FrmPeriodeIndikatorRS.DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' " _
        & "GROUP BY DeskKelas"
        '\---------------------------------------------------/
        strjudul = "Kelas Pelayanan"
        strKriteria = "PER KELAS PELAYANAN"

        adoCommand.CommandText = strSQL
        adoCommand.CommandType = adCmdText

        Select Case cetak
            Case "PerKelas"
                With Report
                    .Database.AddADOCommand dbConn, adoCommand
                    .txtPeriode.SetText tanggal
                    .txtKriteria.SetText strKriteria
                    .usRuangan.SetUnboundFieldSource ("{ado.Kelas}")
                    .unTOI.SetUnboundFieldSource ("{ado.TOI}")
                    .unBOR.SetUnboundFieldSource ("{ado.BOR}")
                    .unBTO.SetUnboundFieldSource ("{ado.BTO}")
                    .unLOS.SetUnboundFieldSource ("{ado.LOS}")
                    .unGDR.SetUnboundFieldSource ("{ado.GDR}")
                    .unNDR.SetUnboundFieldSource ("{ado.NDR}")
                    .unJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")

                    .Text1.SetText strNNamaRS
                    .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
                    .Text3.SetText strWebsite & ", " & strEmail
                    .SelectPrinter sDriver, sPrinter, vbNull
                    settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
                End With

                CRViewer1.ReportSource = Report
            Case "GrafikPerkelas"
                With reportgrafik
                    .Database.AddADOCommand dbConn, adoCommand
                    .txtPeriode.SetText tanggal
                    .txtKriteria.SetText strKriteria
                    .usRuangan.SetUnboundFieldSource ("{ado.Kelas}")
                    .unTOI.SetUnboundFieldSource ("{ado.TOI}")
                    .unBOR.SetUnboundFieldSource ("{ado.BOR}")
                    .unBTO.SetUnboundFieldSource ("{ado.BTO}")
                    .unLOS.SetUnboundFieldSource ("{ado.LOS}")
                    .unGDR.SetUnboundFieldSource ("{ado.GDR}")
                    .unNDR.SetUnboundFieldSource ("{ado.NDR}")
                    .Text1.SetText strNNamaRS
                    .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
                    .Text3.SetText strWebsite & ", " & strEmail
                    .SelectPrinter sDriver, sPrinter, vbNull
                    settingreport reportgrafik, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
                End With

                CRViewer1.ReportSource = reportgrafik
        End Select
    End If
    With CRViewer1
        .PrintReport
        .DisplayTabs = False
        .DisplayGroupTree = False
        .Zoom 1
        .ViewReport
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

