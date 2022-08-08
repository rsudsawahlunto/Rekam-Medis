VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakLaporanIndPlynRS 
   Caption         =   "Medifirst2000 - Cetak"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakLaporanIndPlynRS.frx":0000
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
Attribute VB_Name = "frmCetakLaporanIndPlynRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ReportIndPelKelas As New cr_IndikatorPerKelas
Dim ReportIndPelRuang As New cr_IndikatorPerRuang
Dim ReportIndPelKelasGrafik As New cr_IndikatorPerKelasGrafik
Dim ReportIndPelRuangGrafik As New cr_IndikatorPerRuangGrafik

Private Sub Form_Load()
    Dim tanggal As String
    Dim laporan As String
    Dim adocomd As New ADODB.Command
    On Error GoTo errLoad
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    If frmLapIndPlynRS.cbKriteria.Text = "Per Kelas" Then
        strSQL = "SELECT * " & _
        " FROM V_IndikatorPelayananRSPerKelas" & _
        " WHERE (TglHitung BETWEEN ' " & Format(frmLapIndPlynRS.dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmLapIndPlynRS.dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "') "
        Call msubOpenRecFO(rs, strSQL, dbConn)
        If rs.EOF = True Then
            MsgBox "Data Tidak Ada", vbInformation, "Informasi"
            Exit Sub
        End If

        Call openConnection
        adocomd.ActiveConnection = dbConn
        adocomd.CommandText = "SElect KdRuangan,Ruangan, Kelas, " _
        & " SUM(JmlBed) AS JmlBed, " _
        & " SUM(JmlHariPerawatan) AS JmlHariPerawatan, " _
        & " SUM(JmlPasienOutHidup) AS JmlPasienOutHidup, " _
        & " SUM(JmlPasienOutMati) AS JmlPasienOutMati, " _
        & " SUM(JmlPasienMatiLK48) AS JmlPasienMatiLK48, " _
        & " SUM(JmlPasienMatiLB48) AS JmlPasienMatiLB48, " _
        & " avg(BOR)as BOR,avg(TOI)as TOI,avg(BTO)as BTO, avg(GDR)as GDR,avg(NDR)as NDR from V_IndikatorPelayananRSPerKelas" _
        & " WHERE TglHitung BETWEEN ('" & Format(frmLapIndPlynRS.dtpAwal, "yyyy/mm/dd 00:00:00") & "') AND ('" & Format(frmLapIndPlynRS.dtpAkhir, "yyyy/mm/dd 23:59:59") & "') " _
        & " GROUP BY KdRuangan, Ruangan, Kelas"
        adocomd.CommandType = adCmdText

        If mblnGrafik = False Then
            ReportIndPelKelas.Database.AddADOCommand dbConn, adocomd
            If Format(frmLapIndPlynRS.dtpAwal.value, "dd/mm/yyyy") = Format(frmLapIndPlynRS.dtpAkhir.value, "dd/mm/yyyy") Then
                tanggal = "Tanggal : " & " " & Format(frmLapIndPlynRS.dtpAwal.value, "dd MMMM yyyy") & " S/d " & Format(frmLapIndPlynRS.dtpAkhir.value, "yyyy/mm/dd")
            Else
                tanggal = "Periode : " & " " & Format(frmLapIndPlynRS.dtpAwal.value, "dd MMMM yyyy") & " s/d " & Format(frmLapIndPlynRS.dtpAkhir.value, "dd MMMM yyyy")
            End If
            With ReportIndPelKelas
                .Text1.SetText strNNamaRS & " " & strKelasRS & " " & strKetKelasRS
                .Text2.SetText "KABUPATEN " & strNKotaRS
                .Text3.SetText strNAlamatRS & " " & "Telp." & " " & strNTeleponRS
                .txtTanggal.SetText tanggal
                .txtRuangan.SetText strNNamaRuangan
                .unKelas.SetUnboundFieldSource ("{ado.Kelas}")
                .unBed.SetUnboundFieldSource ("{ado.JmlBed}")
                .unHari.SetUnboundFieldSource ("{ado.JmlHariPerawatan}")
                .unOutHidup.SetUnboundFieldSource ("{ado.JmlPasienOutHidup}")
                .unOutMati.SetUnboundFieldSource ("{ado.JmlPasienOutMati}")
                .unLK48.SetUnboundFieldSource ("{ado.JmlPasienMatiLK48}")
                .unLB48.SetUnboundFieldSource ("{ado.JmlPasienMatiLB48}")
                .unBOR.SetUnboundFieldSource ("{ado.BOR}")
                .unTOI.SetUnboundFieldSource ("{ado.TOI}")
                .unBTO.SetUnboundFieldSource ("{ado.BTO}")
                .unGDR.SetUnboundFieldSource ("{ado.GDR}")
                .unNDR.SetUnboundFieldSource ("{ado.NDR}")
            End With
            If vLaporan = "view" Then
                Screen.MousePointer = vbHourglass
                With CRViewer1
                    .ReportSource = ReportIndPelKelas
                    .ViewReport
                    .Zoom (100)
                End With
            Else
                ReportIndPelKelas.PrintOut False
                Unload Me
            End If
        Else
            'indikator pelayanan per kelas (grafik)
            With ReportIndPelKelasGrafik
                .txtNamaRS.SetText strNNamaRS
                .txtAlamat.SetText strNAlamatRS
                .txtAlamat2.SetText strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS

                If Format(mdTglAwal, "dd MMMM yyyy") = Format(mdTglAkhir, "dd MMMM yyyy") Then
                    .txtTanggal.SetText "Tanggal Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") '& " S/d " & Format(frmregister.DTPickerAkhir, "dd MMMM yyyy")
                Else
                    .txtTanggal.SetText "Periode Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") & " S/d " & Format(mdTglAkhir, "dd MMMM yyyy")
                End If

                .Database.AddADOCommand dbConn, adocomd
                .usRuangan.SetUnboundFieldSource ("{ado.Ruangan}")
                .usKelas.SetUnboundFieldSource ("{ado.Kelas}")
                .unBOR.SetUnboundFieldSource ("{ado.BOR}")
                .unTOI.SetUnboundFieldSource ("{ado.TOI}")
                .unBTO.SetUnboundFieldSource ("{ado.BTO}")
                .unGDR.SetUnboundFieldSource ("{ado.GDR}")
                .unNDR.SetUnboundFieldSource ("{ado.NDR}")
            End With
            If vLaporan = "view" Then
                Screen.MousePointer = vbHourglass
                With CRViewer1
                    .ReportSource = ReportIndPelKelasGrafik
                    .ViewReport
                    .Zoom (100)
                End With
            Else
                ReportIndPelKelasGrafik.PrintOut False
                Unload Me
            End If
        End If

    ElseIf frmLapIndPlynRS.cbKriteria.Text = "Per Ruangan" Then
        strSQL = "SELECT * " & _
        " FROM V_IndikatorPelayananRSPerRuangan" & _
        " WHERE (TglHitung BETWEEN ' " & Format(frmLapIndPlynRS.dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmLapIndPlynRS.dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "') " '"

        Call msubOpenRecFO(rs, strSQL, dbConn)
        If rs.EOF = True Then
            MsgBox "Data Tidak Ada", vbInformation, "Informasi"
            Exit Sub
        End If

        Call openConnection
        adocomd.ActiveConnection = dbConn
        adocomd.CommandText = "SElect KdRuangan, Ruangan, " _
        & " SUM(JmlBed) AS JmlBed, " _
        & " SUM(JmlHariPerawatan) AS JmlHariPerawatan, " _
        & " SUM(JmlPasienOutHidup) AS JmlPasienOutHidup, " _
        & " SUM(JmlPasienOutMati) AS JmlPasienOutMati, " _
        & " SUM(JmlPasienMatiLK48) AS JmlPasienMatiLK48, " _
        & " SUM(JmlPasienMatiLB48) AS JmlPasienMatiLB48, " _
        & " avg(LOS)as LOS, avg(BOR)as BOR, avg(TOI)as TOI, avg(BTO)as BTO, avg(GDR)as GDR, avg(NDR)as NDR " _
        & " from V_IndikatorPelayananRSPerRuangan" _
        & " WHERE TglHitung BETWEEN ('" & Format(frmLapIndPlynRS.dtpAwal, "yyyy/mm/dd 00:00:00") & "') AND ('" & Format(frmLapIndPlynRS.dtpAkhir, "yyyy/mm/dd 23:59:59") & "') " _
        & " GROUP BY KdRuangan, Ruangan"
        adocomd.CommandType = adCmdText

        If mblnGrafik = False Then
            ReportIndPelRuang.Database.AddADOCommand dbConn, adocomd
            If Format(frmLapIndPlynRS.dtpAwal.value, "dd/mm/yyyy") = Format(frmLapIndPlynRS.dtpAkhir.value, "dd/mm/yyyy") Then
                tanggal = "Tanggal : " & " " & Format(frmLapIndPlynRS.dtpAwal.value, "dd MMMM yyyy") & " S/d " & Format(frmLapIndPlynRS.dtpAkhir.value, "yyyy/mm/dd")
            Else
                tanggal = "Periode : " & " " & Format(frmLapIndPlynRS.dtpAwal.value, "dd MMMM yyyy") & " s/d " & Format(frmLapIndPlynRS.dtpAkhir.value, "dd MMMM yyyy")
            End If

            With ReportIndPelRuang
                .Text1.SetText strNNamaRS & " " & strKelasRS & " " & strKetKelasRS
                .Text2.SetText "KABUPATEN " & strNKotaRS
                .Text3.SetText strNAlamatRS & " " & "Telp." & " " & strNTeleponRS
                .txtTanggal.SetText tanggal
                .txtRuangan.SetText strNNamaRuangan
                .unBed.SetUnboundFieldSource ("{ado.JmlBed}")
                .unHari.SetUnboundFieldSource ("{ado.JmlHariPerawatan}")
                .unOutHidup.SetUnboundFieldSource ("{ado.JmlPasienOutHidup}")
                .unOutMati.SetUnboundFieldSource ("{ado.JmlPasienOutMati}")
                .unLK48.SetUnboundFieldSource ("{ado.JmlPasienMatiLK48}")
                .unLB48.SetUnboundFieldSource ("{ado.JmlPasienMatiLB48}")
                .unLOS.SetUnboundFieldSource ("{ado.LOS}")
                .unBOR.SetUnboundFieldSource ("{ado.BOR}")
                .unTOI.SetUnboundFieldSource ("{ado.TOI}")
                .unBTO.SetUnboundFieldSource ("{ado.BTO}")
                .unGDR.SetUnboundFieldSource ("{ado.GDR}")
                .unNDR.SetUnboundFieldSource ("{ado.NDR}")
            End With
            If vLaporan = "view" Then
                Screen.MousePointer = vbHourglass
                With CRViewer1
                    .ReportSource = ReportIndPelRuang
                    .ViewReport
                    .Zoom (100)
                End With
            Else
                ReportIndPelRuang.PrintOut False
                Unload Me
            End If
        Else
            '            indikator pelayanan per ruangan (grafik)
            With ReportIndPelRuangGrafik
                .txtNamaRS.SetText strNNamaRS
                .txtAlamat.SetText strNAlamatRS
                .txtAlamat2.SetText strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS

                If Format(mdTglAwal, "dd MMMM yyyy") = Format(mdTglAkhir, "dd MMMM yyyy") Then
                    .txtTanggal.SetText "Tanggal Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") '& " S/d " & Format(frmregister.DTPickerAkhir, "dd MMMM yyyy")
                Else
                    .txtTanggal.SetText "Periode Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") & " S/d " & Format(mdTglAkhir, "dd MMMM yyyy")
                End If

                .Database.AddADOCommand dbConn, adocomd
                .usRuangan.SetUnboundFieldSource ("{ado.Ruangan}")
                .unBOR.SetUnboundFieldSource ("{ado.BOR}")
                .unTOI.SetUnboundFieldSource ("{ado.TOI}")
                .unBTO.SetUnboundFieldSource ("{ado.BTO}")
                .unGDR.SetUnboundFieldSource ("{ado.GDR}")
                .unNDR.SetUnboundFieldSource ("{ado.NDR}")
            End With
            If vLaporan = "view" Then
                Screen.MousePointer = vbHourglass
                With CRViewer1
                    .ReportSource = ReportIndPelRuangGrafik
                    .ViewReport
                    .Zoom (100)
                End With
            Else
                ReportIndPelRuangGrafik.PrintOut False
                Unload Me
            End If
        End If
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
errLoad:
    Screen.MousePointer = vbDefault
    msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakLaporanIndPlynRS = Nothing
    mblnGrafik = False
End Sub

