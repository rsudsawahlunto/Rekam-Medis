VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmRekapPasienINADRG 
   Caption         =   "konversi INA DRG - Laporan Rekapitulasi"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   Icon            =   "frmRekapPasienINADRG.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   11010
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7725
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10965
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
Attribute VB_Name = "frmRekapPasienINADRG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crRekapPasienINADRG
Dim report2 As New crDetailINADRG
Dim report3 As New crRekapPasienINADRGPlusNama
Dim report4 As New crRekapPasienINADRGLengkap

Private Sub Form_Load()
    If strCetak = "Laporan Detail" Then
        Call DetailINADRG
    ElseIf strCetak = "Laporan Rekap" Then
        Call RekapINADRG
    ElseIf strCetak = "Laporan Rekap Nama" Then
        Call RekapINADRGByNama
    ElseIf strCetak = "Laporan Rekap Lengkap" Then
        Call RekapINADRGLengkap
    End If

End Sub

Private Sub RekapINADRG()
    On Error GoTo hell
    Dim dAwal As Date
    Dim dAkhir As Date
    Dim d As Integer

    Call openConnection

    Set frmRekapPasienINADRG = Nothing
    Set dbcmd = New ADODB.Command
    Set dbcmd = Nothing

    Me.WindowState = 2

    'get periode laporan
    Set rs = Nothing
    Call msubRecFO(rs, "Select TglKeluar From LoadHasilINADRG ORDER BY TglKeluar")
    If rs.EOF = True Then Exit Sub
    dAwal = Format(rs("TglKeluar"), "dd/MM/yyyy")
    For d = 1 To rs.RecordCount
        If Format(dAkhir, "dd/MM/yyyy") <> Format(rs("TglKeluar"), "dd/MM/yyyy") Then
            dAkhir = Format(rs("TglKeluar"), "dd/MM/yyyy")
        End If
        rs.MoveNext
    Next d

    dbcmd.ActiveConnection = dbConn
    dbcmd.CommandText = "Select * From V_LaporanINADRG ORDER BY TglKeluar,DESKRIPSI"
    dbcmd.CommandType = adCmdText

    With Report
        .Database.AddADOCommand dbConn, dbcmd
        .txtNomorKdRS.SetText strKdRS & " / " & strKelasRS
        .txtNamaRS.SetText strNamaRS

        .udTglKeluar.SetUnboundFieldSource ("{ado.TglKeluar}")
        .usNoRM.SetUnboundFieldSource ("{ado.NoRM}")
        .usKdINADRG.SetUnboundFieldSource ("{ado.KdINADRG}")
        .usDesk.SetUnboundFieldSource ("{ado.DESKRIPSI}")
        .ucTarifINADRG.SetUnboundFieldSource ("{ado.TarifINADRG}")
        .ucTarifRS.SetUnboundFieldSource ("{ado.TarifRS}")
    End With

    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
    CRViewer1.Zoom 1
    Screen.MousePointer = vbDefault
    Exit Sub
hell:
    Screen.MousePointer = vbDefault
    Set frmRekapPasienINADRG = Nothing
    Set dbcmd = Nothing
    Call msubPesanError
End Sub

Private Sub DetailINADRG()
    On Error GoTo hell
    Dim dAwal As Date
    Dim dAkhir As Date
    Dim d As Integer

    Call openConnection

    Set frmRekapPasienINADRG = Nothing
    Set dbcmd = New ADODB.Command
    Set dbcmd = Nothing

    Me.WindowState = 2

    dbcmd.ActiveConnection = dbConn
    dbcmd.CommandText = "Select * From V_LaporanINADRGJEP2009 where JenisPerawatan = '" & IIf(frmLap20091201003.dcInstalasi.BoundText = "02", "2", "1") & "' and TglKeluar between '" & Format(frmLap20091201003.dtpAwal, "yyyy/MM/dd 00:00:00") & "' and '" & Format(frmLap20091201003.dtpAkhir, "yyyy/MM/dd 23:59:59") & "' ORDER BY TglKeluar,DESKRIPSI"
    dbcmd.CommandType = adCmdText

    If rsProfil.State = adStateOpen Then rsProfil.Close
    rsProfil.Open "select * from ProfilRS", dbConn, adOpenForwardOnly, adLockReadOnly

    With report2
        .Database.AddADOCommand dbConn, dbcmd
        .txtKdRS.SetText rsProfil("kdRS").value 'strKdRS
        .txtKelasRS.SetText rsProfil("kelasRS").value
        .txtNamaRS.SetText rsProfil("NamaRS").value

        .udTglKeluar.SetUnboundFieldSource ("{ado.TglKeluar}")
        .usNoRM.SetUnboundFieldSource ("{ado.NoRM}")
        .unUmurThn.SetUnboundFieldSource ("{ado.UmurThn}")
        .unUmurHr.SetUnboundFieldSource ("{ado.UmurHari}")
        .udTgllahir.SetUnboundFieldSource ("{ado.TglLahir}")
        .unJK.SetUnboundFieldSource ("{ado.JK}")
        .unKelasPerawatan.SetUnboundFieldSource ("{ado.KelasPerawatan}")
        .udTglMasuk.SetUnboundFieldSource ("{ado.TglMasuk}")
        .unJenisPerawatan.SetUnboundFieldSource ("{ado.JenisPerawatan}")
        .unCaraPulang.SetUnboundFieldSource ("{ado.CaraPulang}")
        .unLOS.SetUnboundFieldSource ("{ado.LOS}")
        .unBeratLahir.SetUnboundFieldSource ("{ado.BeratLahir}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaLengkap}")

        .usDX1.SetUnboundFieldSource ("{ado.ICD10_1}")
        .UnboundString1.SetUnboundFieldSource ("{ado.ICD10_2}")
        .UnboundString2.SetUnboundFieldSource ("{ado.ICD10_3}")
        .UnboundString3.SetUnboundFieldSource ("{ado.ICD10_4}")
        .UnboundString4.SetUnboundFieldSource ("{ado.ICD10_5}")
        .UnboundString5.SetUnboundFieldSource ("{ado.ICD10_6}")
        .UnboundString6.SetUnboundFieldSource ("{ado.ICD10_7}")
        .UnboundString7.SetUnboundFieldSource ("{ado.ICD10_8}")
        .UnboundString8.SetUnboundFieldSource ("{ado.ICD10_9}")
        .UnboundString9.SetUnboundFieldSource ("{ado.ICD10_10}")
        .UnboundString10.SetUnboundFieldSource ("{ado.ICD10_11}")
        .UnboundString11.SetUnboundFieldSource ("{ado.ICD10_12}")
        .UnboundString12.SetUnboundFieldSource ("{ado.ICD10_13}")
        .UnboundString13.SetUnboundFieldSource ("{ado.ICD10_14}")
        .UnboundString14.SetUnboundFieldSource ("{ado.ICD10_15}")
        .UnboundString15.SetUnboundFieldSource ("{ado.ICD10_16}")
        .UnboundString16.SetUnboundFieldSource ("{ado.ICD10_17}")
        .UnboundString17.SetUnboundFieldSource ("{ado.ICD10_18}")
        .UnboundString18.SetUnboundFieldSource ("{ado.ICD10_19}")
        .UnboundString19.SetUnboundFieldSource ("{ado.ICD10_20}")
        .UnboundString20.SetUnboundFieldSource ("{ado.ICD10_21}")
        .UnboundString21.SetUnboundFieldSource ("{ado.ICD10_22}")
        .UnboundString22.SetUnboundFieldSource ("{ado.ICD10_23}")
        .UnboundString23.SetUnboundFieldSource ("{ado.ICD10_24}")
        .UnboundString24.SetUnboundFieldSource ("{ado.ICD10_25}")
        .UnboundString25.SetUnboundFieldSource ("{ado.ICD10_26}")
        .UnboundString26.SetUnboundFieldSource ("{ado.ICD10_27}")
        .UnboundString27.SetUnboundFieldSource ("{ado.ICD10_28}")
        .UnboundString28.SetUnboundFieldSource ("{ado.ICD10_29}")
        .UnboundString29.SetUnboundFieldSource ("{ado.ICD10_30}")

        .UnboundString30.SetUnboundFieldSource ("{ado.ICD9_1}")
        .UnboundString31.SetUnboundFieldSource ("{ado.ICD9_2}")
        .UnboundString32.SetUnboundFieldSource ("{ado.ICD9_3}")
        .UnboundString33.SetUnboundFieldSource ("{ado.ICD9_4}")
        .UnboundString34.SetUnboundFieldSource ("{ado.ICD9_5}")
        .UnboundString35.SetUnboundFieldSource ("{ado.ICD9_6}")
        .UnboundString36.SetUnboundFieldSource ("{ado.ICD9_7}")
        .UnboundString37.SetUnboundFieldSource ("{ado.ICD9_8}")
        .UnboundString38.SetUnboundFieldSource ("{ado.ICD9_9}")
        .UnboundString39.SetUnboundFieldSource ("{ado.ICD9_10}")
        .UnboundString40.SetUnboundFieldSource ("{ado.ICD9_11}")
        .UnboundString41.SetUnboundFieldSource ("{ado.ICD9_12}")
        .UnboundString42.SetUnboundFieldSource ("{ado.ICD9_13}")
        .UnboundString43.SetUnboundFieldSource ("{ado.ICD9_14}")
        .UnboundString44.SetUnboundFieldSource ("{ado.ICD9_15}")
        .UnboundString45.SetUnboundFieldSource ("{ado.ICD9_16}")
        .UnboundString46.SetUnboundFieldSource ("{ado.ICD9_17}")
        .UnboundString47.SetUnboundFieldSource ("{ado.ICD9_18}")
        .UnboundString48.SetUnboundFieldSource ("{ado.ICD9_19}")
        .UnboundString49.SetUnboundFieldSource ("{ado.ICD9_20}")
        .UnboundString50.SetUnboundFieldSource ("{ado.ICD9_21}")
        .UnboundString51.SetUnboundFieldSource ("{ado.ICD9_22}")
        .UnboundString52.SetUnboundFieldSource ("{ado.ICD9_23}")
        .UnboundString53.SetUnboundFieldSource ("{ado.ICD9_24}")
        .UnboundString54.SetUnboundFieldSource ("{ado.ICD9_25}")
        .UnboundString55.SetUnboundFieldSource ("{ado.ICD9_26}")
        .UnboundString56.SetUnboundFieldSource ("{ado.ICD9_27}")
        .UnboundString57.SetUnboundFieldSource ("{ado.ICD9_28}")
        .UnboundString58.SetUnboundFieldSource ("{ado.ICD9_29}")
        .UnboundString59.SetUnboundFieldSource ("{ado.ICD9_30}")

        .usKdINADRG.SetUnboundFieldSource ("{ado.KdINADRG}")
        .usDesk.SetUnboundFieldSource ("{ado.DESKRIPSI}")
        .unALOS.SetUnboundFieldSource ("{ado.ALOS}")
        .ucTarifINADRG.SetUnboundFieldSource ("{ado.TarifINADRG}")
    End With

    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = report2
    CRViewer1.ViewReport
    CRViewer1.Zoom 1
    Screen.MousePointer = vbDefault
    Exit Sub
hell:
    Screen.MousePointer = vbDefault
    Set frmRekapPasienINADRG = Nothing
    Set dbcmd = Nothing
    Call msubPesanError
End Sub

Private Sub RekapINADRGByNama()
    On Error GoTo hell
    Dim dAwal As Date
    Dim dAkhir As Date
    Dim d As Integer

    Call openConnection

    Set frmRekapPasienINADRG = Nothing
    Set dbcmd = New ADODB.Command
    Set dbcmd = Nothing

    Me.WindowState = 2

    'get periode laporan
    Set rs = Nothing
    Call msubRecFO(rs, "Select TglKeluar From LoadHasilINADRG ORDER BY TglKeluar")
    If rs.EOF = True Then Exit Sub
    dAwal = Format(rs("TglKeluar"), "dd/MM/yyyy")
    For d = 1 To rs.RecordCount
        If Format(dAkhir, "dd/MM/yyyy") <> Format(rs("TglKeluar"), "dd/MM/yyyy") Then
            dAkhir = Format(rs("TglKeluar"), "dd/MM/yyyy")
        End If
        rs.MoveNext
    Next d

    dbcmd.ActiveConnection = dbConn
    dbcmd.CommandText = "Select * From V_LaporanINADRG ORDER BY TglKeluar,DESKRIPSI"
    dbcmd.CommandType = adCmdText

    With report3
        .Database.AddADOCommand dbConn, dbcmd
        .txtNomorKdRS.SetText strKdRS & " / " & strKelasRS
        .txtNamaRS.SetText strNamaRS

        'periode
        .txtTglAwal.SetText CStr(dAwal)
        .txtTglAkhir.SetText CStr(dAkhir)

        .udTglKeluar.SetUnboundFieldSource ("{ado.TglKeluar}")
        .usNoRM.SetUnboundFieldSource ("{ado.NoRM}")
        .usKdINADRG.SetUnboundFieldSource ("{ado.KdINADRG}")
        .usDesk.SetUnboundFieldSource ("{ado.DESKRIPSI}")
        .ucTarifINADRG.SetUnboundFieldSource ("{ado.TarifINADRG}")
        .ucTarifRS.SetUnboundFieldSource ("{ado.TarifRS}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaLengkap}")
    End With

    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = report3
    CRViewer1.ViewReport
    CRViewer1.Zoom 1
    Screen.MousePointer = vbDefault
    Exit Sub
hell:
    Screen.MousePointer = vbDefault
    Set frmRekapPasienINADRG = Nothing
    Set dbcmd = Nothing
    Call msubPesanError
End Sub

Private Sub RekapINADRGLengkap()
    On Error GoTo hell
    Dim dAwal As Date
    Dim dAkhir As Date
    Dim d As Integer

    Call openConnection

    Set frmRekapPasienINADRG = Nothing
    Set dbcmd = New ADODB.Command
    Set dbcmd = Nothing

    Me.WindowState = 2

    If rsProfil.State = adStateOpen Then rsProfil.Close
    rsProfil.Open "select * from ProfilRS", dbConn, adOpenForwardOnly, adLockReadOnly

    dbcmd.ActiveConnection = dbConn
    dbcmd.CommandText = "Select * From V_LaporanINADRGJEP2009 where JenisPerawatan = '" & IIf(frmLap20091201004.dcInstalasi.BoundText = "02", "2", "1") & "' and TglKeluar between '" & Format(frmLap20091201004.dtpAwal, "yyyy/MM/dd 00:00:00") & "' and '" & Format(frmLap20091201004.dtpAkhir, "yyyy/MM/dd 23:59:59") & "' ORDER BY TglKeluar,DESKRIPSI"
    dbcmd.CommandType = adCmdText

    With report4
        .Database.AddADOCommand dbConn, dbcmd
        .txtNomorKdRS.SetText rsProfil("kdRS").value & " / " & rsProfil("kelasRS").value
        .txtNamaRS.SetText rsProfil("NamaRS").value

        .udTglKeluar.SetUnboundFieldSource ("{ado.TglKeluar}")
        .usNoRM.SetUnboundFieldSource ("{ado.NoRM}")
        .usKdINADRG.SetUnboundFieldSource ("{ado.KdINADRG}")
        .usDesk.SetUnboundFieldSource ("{ado.DESKRIPSI}")
        .ucTarifINADRG.SetUnboundFieldSource ("{ado.TarifINADRG}")
        .ucTarifRS.SetUnboundFieldSource ("{ado.TarifRS}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaLengkap}")
        .usDXU10.SetUnboundFieldSource ("{ado.ICD10_1}")
        .usDiagnosaUtama.SetUnboundFieldSource ("{ado.NamaDiagnosa}")
    End With

    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = report4
    CRViewer1.ViewReport
    CRViewer1.Zoom 1
    Screen.MousePointer = vbDefault

    Exit Sub
hell:
    Screen.MousePointer = vbDefault
    Set frmRekapPasienINADRG = Nothing
    Set dbcmd = Nothing
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRekapPasienINADRG = Nothing
End Sub
