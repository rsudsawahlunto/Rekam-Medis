VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmLapRL5 
   Caption         =   "Medifirst2000 - Laporan RL 5"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmLapRL5.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
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
Attribute VB_Name = "frmLapRL5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crLapRL5

Private Sub Form_Load()
    On Error GoTo hell

    Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Set frmLapRL5 = Nothing

    Set rs = Nothing
    Call msubRecFO(rs, "Select * From TempDataKegiatanKesehatanLingkungan_RL5")
    If rs.EOF Then
        MsgBox "Ada masalah waktu penyimpanan data, hubungi administrator", vbCritical, "Error"
        Exit Sub
    End If

    With Report
        .txtNamaRS.SetText rs("NamaRS")
        .txtKdRS.SetText rs("KdRS")

        If rs("DokAmdal") = 0 Then .txtDokAmdalAda.SetText "X" Else .txtDokAmdalTdkAda.SetText "X"
        If rs("DokUKL") = 0 Then .txtDokUKLAda.SetText "X" Else .txtDokUKLTdkAda.SetText "X"

        If rs("InstLimbahCair") = 0 Then .txtInstLimbahCairAda.SetText "X" Else .txtInstLimbahCairTdkAda.SetText "X"
        If rs("IPAL") = 0 Then .txtIPALAda.SetText "X" Else .txtIPALTdkAda.SetText "X"
        If rs("BKuartal1") = 0 Then .txtBKuartal1MS.SetText "X" Else .txtBKuartal1TMS.SetText "X"
        If rs("BKuartal2") = 0 Then .txtBKuartal2MS.SetText "X" Else .txtBKuartal2TMS.SetText "X"
        If rs("BKuartal3") = 0 Then .txtBKuartal3MS.SetText "X" Else .txtBKuartal3TMS.SetText "X"
        .txtThnIPAL.SetText rs("ThnIPAL")
        If Trim(rs("BSumberDana")) = "APBN" Then
            .txtBSumberDanaAPBN.SetText "X"
        ElseIf Trim(rs("BSumberDana")) = "APBD" Then
            .txtBSumberDanaAPBD.SetText "X"
        ElseIf Trim(rs("BSumberDana")) = "BLN" Then
            .txtBSumberDanaBLN.SetText "X"
        ElseIf Trim(rs("BSumberDana")) = "RS" Then
            .txtBSumberDanaRS.SetText "X"
        End If

        If rs("SaranaInsinerator") = 0 Then .txtSaranaInsineratorAda.SetText "X" Else .txtSaranaInsineratorTdkAda.SetText "X"
        If rs("Insinerator") = 0 Then .txtInsineratorYa.SetText "X" Else .txtInsineratorTdk.SetText "X"
        If rs("Permenkes") = 0 Then .txtPermenkesYa.SetText "X" Else .txtPermenkesTdk.SetText "X"
        .txtC4.SetText rs("C4")
        .txtC5.SetText rs("C5")
        .txtThnInsinerator.SetText rs("ThnInsinerator")
        If Trim(rs("CSumberDana")) = "APBN" Then
            .txtCSumberDanaAPBN.SetText "X"
        ElseIf Trim(rs("CSumberDana")) = "APBD" Then
            .txtCSumberDanaAPBD.SetText "X"
        ElseIf Trim(rs("CSumberDana")) = "BLN" Then
            .txtCSumberDanaBLN.SetText "X"
        ElseIf Trim(rs("CSumberDana")) = "RS" Then
            .txtCSumberDanaRS.SetText "X"
        End If

        If rs("SaranaAirBersih") = 0 Then .txtAirBersihPDAM.SetText "X" Else .txtAirBersihSumurBor.SetText "X"
        If rs("Kuantitas") = 0 Then .txtKuantitasCkp.SetText "X" Else .txtKuantitasTdkCkp.SetText "X"
        If rs("Kontinuitas") = 0 Then .txtKontinuitasYa.SetText "X" Else .txtKontinuitasTdk.SetText "X"
        If rs("DKuartal1") = 0 Then .txtDKuartal1MS.SetText "X" Else .txtDKuartal1TMS.SetText "X"
        If rs("DKuartal2") = 0 Then .txtDKuartal2MS.SetText "X" Else .txtDKuartal2TMS.SetText "X"
        If rs("DKuartal3") = 0 Then .txtDKuartal3MS.SetText "X" Else .txtDKuartal3TMS.SetText "X"

        .txtKota.SetText strNKotaRS
        Set dbRst = Nothing
        Call msubRecFO(dbRst, "select [Nama Lengkap] from V_M_DataPegawaiNew where KdJabatan='A01'")
        If dbRst.EOF Then MsgBox "Jabatan direktur tidak ada, hubungi administrator", vbCritical, "Error"
        .txtNamaDirekturRS.SetText dbRst(0)
        .txtNamaPegawai.SetText strNmPegawai
    End With

    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 1
    End With
    Screen.MousePointer = vbDefault

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dbConn.Execute "Delete From TempDataKegiatanKesehatanLingkungan_RL5"
    Set frmLapRL5 = Nothing
    frmDataKegiatanKesehatanLingkunganRL5.cmdBatal_Click
End Sub
