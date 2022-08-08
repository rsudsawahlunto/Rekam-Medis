VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmMorbiditasRI 
   Caption         =   "Morbiditas "
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmMorbiditasRawatInap.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   -30
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
Attribute VB_Name = "frmMorbiditasRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crMorbiditasRawatInap
Dim adoCommand As New ADODB.Command
Dim strSQL As String

Private Sub Form_Load()
    openConnection
    Set adoCommand.ActiveConnection = dbConn
    Set frmMorbiditasRI = Nothing
    If Format(frmPeriodeKunjungan.DTPickerAwal, "dd MMMM yyyy") = Format(frmPeriodeKunjungan.DTPickerAkhir, "dd MMMM yyyy") Then
        tanggal = "Tanggal Kunjungan  : " & " " & Format(frmPeriodeKunjungan.DTPickerAwal, "dd MMMM yyyy") '& " S/d " & Format(frmPeriodeKunjungan.DTPickerAkhir, "dd MMMM yyyy")
    Else
        tanggal = "Periode Kunjungan  : " & " " & Format(frmPeriodeKunjungan.DTPickerAwal, "dd MMMM yyyy") & " S/d " & Format(frmPeriodeKunjungan.DTPickerAkhir, "dd MMMM yyyy")
    End If

    Select Case ctk
        Case "IGD"
            strSQL = "SELECT NoDTD,NoDTerperinci,NamaDTD,SUM(Kel_Umur1) AS Kel_Umur1,SUM(Kel_Umur2) AS Kel_Umur2,SUM(Kel_Umur3) AS Kel_Umur3,SUM(Kel_Umur4) AS Kel_Umur4,SUM(Kel_Umur5) AS Kel_Umur5,SUM(Kel_Umur6) AS Kel_Umur6,SUM(Kel_Umur7) AS Kel_Umur7,SUM(Kel_Umur8) AS Kel_Umur8,SUM(Kel_L) AS Kel_L,SUM(Kel_P) AS Kel_P,SUM(Kel_H) AS Kel_H,SUM(Kel_M) AS Kel_M " _
            & "FROM v_S_RekapMorbidRI " _
            & "WHERE TglPeriksa BETWEEN '" & Format(frmPeriodeKunjungan.DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmPeriodeKunjungan.DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' " _
            & " and  kdinstalasi = '01' " _
            & "GROUP BY NoDTD,NoDTerperinci,NamaDTD"
            adoCommand.CommandText = strSQL
            adoCommand.CommandType = adCmdText
            With Report
                .Database.AddADOCommand dbConn, adoCommand
                .txtJudul.SetText "DATA KEADAAN MORBIDITAS IGD SURVEILANS TERPADU RUMAH SAKIT"
                .txtJudul2.SetText "FORMULIR RL 2a1"
                .txtPeriode.SetText tanggal
                .usNoDTD.SetUnboundFieldSource ("{ado.NoDTD}")
                .usNoDT.SetUnboundFieldSource ("{ado.NoDTerperinci}")
                .usNamaDTD.SetUnboundFieldSource ("{ado.NamaDTD}")
                .unKel1.SetUnboundFieldSource ("{ado.Kel_Umur1}")
                .unKel2.SetUnboundFieldSource ("{ado.Kel_Umur2}")
                .unKel3.SetUnboundFieldSource ("{ado.Kel_Umur3}")
                .unKel4.SetUnboundFieldSource ("{ado.Kel_Umur4}")
                .unKel5.SetUnboundFieldSource ("{ado.Kel_Umur5}")
                .unKel6.SetUnboundFieldSource ("{ado.Kel_Umur6}")
                .unKel7.SetUnboundFieldSource ("{ado.Kel_Umur7}")
                .unKel8.SetUnboundFieldSource ("{ado.Kel_Umur8}")
                .unKelL.SetUnboundFieldSource ("{ado.Kel_L}")
                .unKelP.SetUnboundFieldSource ("{ado.Kel_P}")
                .unKelH.SetUnboundFieldSource ("{ado.Kel_H}")
                .unKelM.SetUnboundFieldSource ("{ado.Kel_M}")
                .Text1.SetText strNNamaRS
                .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
                .Text3.SetText strWebsite & ", " & strEmail
                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport Report, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
            End With

        Case "RI"
            strSQL = "SELECT NoDTD,NoDTerperinci,NamaDTD,SUM(Kel_Umur1) AS Kel_Umur1,SUM(Kel_Umur2) AS Kel_Umur2,SUM(Kel_Umur3) AS Kel_Umur3,SUM(Kel_Umur4) AS Kel_Umur4,SUM(Kel_Umur5) AS Kel_Umur5,SUM(Kel_Umur6) AS Kel_Umur6,SUM(Kel_Umur7) AS Kel_Umur7,SUM(Kel_Umur8) AS Kel_Umur8,SUM(Kel_L) AS Kel_L,SUM(Kel_P) AS Kel_P,SUM(Kel_H) AS Kel_H,SUM(Kel_M) AS Kel_M " _
            & "FROM v_S_RekapMorbidRI " _
            & "WHERE TglPeriksa BETWEEN '" & Format(frmPeriodeKunjungan.DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmPeriodeKunjungan.DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' " _
            & " and kdinstalasi = '03'" _
            & "GROUP BY NoDTD,NoDTerperinci,NamaDTD"
            adoCommand.CommandText = strSQL
            adoCommand.CommandType = adCmdText
            With Report
                .Database.AddADOCommand dbConn, adoCommand
                .txtJudul.SetText "DATA KEADAAN MORBIDITAS RAWAT INAP SURVEILANS TERPADU RUMAH SAKIT"
                .txtJudul2.SetText "FORMULIR RL 2a1"
                .txtPeriode.SetText tanggal
                .usNoDTD.SetUnboundFieldSource ("{ado.NoDTD}")
                .usNoDT.SetUnboundFieldSource ("{ado.NoDTerperinci}")
                .usNamaDTD.SetUnboundFieldSource ("{ado.NamaDTD}")
                .unKel1.SetUnboundFieldSource ("{ado.Kel_Umur1}")
                .unKel2.SetUnboundFieldSource ("{ado.Kel_Umur2}")
                .unKel3.SetUnboundFieldSource ("{ado.Kel_Umur3}")
                .unKel4.SetUnboundFieldSource ("{ado.Kel_Umur4}")
                .unKel5.SetUnboundFieldSource ("{ado.Kel_Umur5}")
                .unKel6.SetUnboundFieldSource ("{ado.Kel_Umur6}")
                .unKel7.SetUnboundFieldSource ("{ado.Kel_Umur7}")
                .unKel8.SetUnboundFieldSource ("{ado.Kel_Umur8}")
                .unKelL.SetUnboundFieldSource ("{ado.Kel_L}")
                .unKelP.SetUnboundFieldSource ("{ado.Kel_P}")
                .unKelH.SetUnboundFieldSource ("{ado.Kel_H}")
                .unKelM.SetUnboundFieldSource ("{ado.Kel_M}")
                .Text1.SetText strNNamaRS
                .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
                .Text3.SetText strWebsite & ", " & strEmail
                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport Report, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
            End With
    End Select
    
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
    Set frmMorbiditasRI = Nothing
End Sub

