VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmRL2A 
   Caption         =   "Morbiditas"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   Icon            =   "frmRL2A.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   8325
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
Attribute VB_Name = "frmRL2A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crRL2A
Dim adoCommand As New ADODB.Command

Private Sub Form_Load()
    On Error GoTo hell

    openConnection

    Set frmRL2A = Nothing

    adoCommand.CommandText = strSQL
    adoCommand.CommandType = adCmdText

    'Triwulan 1
    If frmLapRL2A.Check1.value = vbChecked And frmLapRL2A.Option1.value = True Then
        With Report
            .Database.AddADOCommand dbConn, adoCommand

            .txtJudul.SetText "DATA KEADAAN MORBIDITAS PASIEN RAWAT INAP RUMAH SAKIT"
            .Text1.SetText strNNamaRS
            .txtJudul2.SetText "FORMULIR RL2a"
            .Text3.SetText " I "
            .Text48.SetText CStr(Format(mdTglAwal, "yyyy"))
            .usNoDTD.SetUnboundFieldSource ("{ado.NoDTD}")
            .usGroup.SetUnboundFieldSource ("{ado.Grup}")
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
        End With
        strSQL = "select * " & _
        " from V_Koders  "

        Call msubRecFO(rs, strSQL)

        With Report
            .Text49.SetText rs("NO1")
            .Text50.SetText rs("NO2")
            .Text51.SetText rs("NO3")
            .Text52.SetText rs("NO4")
            .Text53.SetText rs("NO5")
            .Text54.SetText rs("NO6")
            .Text55.SetText rs("NO7")
        End With
        If vLaporan = "view" Then
            Screen.MousePointer = vbHourglass
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
    End If

    'Triwulan 2
    If frmLapRL2A.Check1.value = vbChecked And frmLapRL2A.Option2.value = True Then
        With Report
            .Database.AddADOCommand dbConn, adoCommand

            .txtJudul.SetText "DATA KEADAAN MORBIDITAS PASIEN RAWAT INAP RUMAH SAKIT"
            .Text1.SetText strNNamaRS
            .txtJudul2.SetText "FORMULIR RL2a"
            .Text3.SetText " II "
            .Text48.SetText CStr(Format(mdTglAwal, "yyyy"))
            .usNoDTD.SetUnboundFieldSource ("{ado.NoDTD}")
            .usGroup.SetUnboundFieldSource ("{ado.Grup}")
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
        End With

        strSQL = "select * " & _
        " from V_Koders  "

        Call msubRecFO(rs, strSQL)

        With Report
            .Text49.SetText rs("NO1")
            .Text50.SetText rs("NO2")
            .Text51.SetText rs("NO3")
            .Text52.SetText rs("NO4")
            .Text53.SetText rs("NO5")
            .Text54.SetText rs("NO6")
            .Text55.SetText rs("NO7")
        End With

        If vLaporan = "view" Then
            Screen.MousePointer = vbHourglass
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
    End If

    'Triwulan 3
    If frmLapRL2A.Check1.value = vbChecked And frmLapRL2A.Option3.value = True Then
        With Report
            .Database.AddADOCommand dbConn, adoCommand

            .txtJudul.SetText "DATA KEADAAN MORBIDITAS PASIEN RAWAT INAP RUMAH SAKIT"
            .Text1.SetText strNNamaRS
            .txtJudul2.SetText "FORMULIR RL2a"
            .Text3.SetText " III "
            .Text48.SetText CStr(Format(mdTglAwal, "yyyy"))
            .usNoDTD.SetUnboundFieldSource ("{ado.NoDTD}")
            .usGroup.SetUnboundFieldSource ("{ado.Grup}")
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
        End With

        strSQL = "select * " & _
        " from V_Koders  "

        Call msubRecFO(rs, strSQL)

        With Report
            .Text49.SetText rs("NO1")
            .Text50.SetText rs("NO2")
            .Text51.SetText rs("NO3")
            .Text52.SetText rs("NO4")
            .Text53.SetText rs("NO5")
            .Text54.SetText rs("NO6")
            .Text55.SetText rs("NO7")
        End With
        If vLaporan = "view" Then
            Screen.MousePointer = vbHourglass
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
    End If

    'Triwulan 4
    If frmLapRL2A.Check1.value = vbChecked And frmLapRL2A.Option4.value = True Then
        With Report
            .Database.AddADOCommand dbConn, adoCommand

            .txtJudul.SetText "DATA KEADAAN MORBIDITAS PASIEN RAWAT INAP RUMAH SAKIT"
            .Text1.SetText strNNamaRS
            .txtJudul2.SetText "FORMULIR RL2a"
            .Text3.SetText " IV "
            .Text48.SetText CStr(Format(mdTglAwal, "yyyy"))
            .usNoDTD.SetUnboundFieldSource ("{ado.NoDTD}")
            .usGroup.SetUnboundFieldSource ("{ado.Grup}")
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
        End With

        strSQL = "select * " & _
        " from V_Koders  "

        Call msubRecFO(rs, strSQL)

        With Report
            .Text49.SetText rs("NO1")
            .Text50.SetText rs("NO2")
            .Text51.SetText rs("NO3")
            .Text52.SetText rs("NO4")
            .Text53.SetText rs("NO5")
            .Text54.SetText rs("NO6")
            .Text55.SetText rs("NO7")
        End With
        If vLaporan = "view" Then
            Screen.MousePointer = vbHourglass
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
    End If

    'Periode
    If frmLapRL2A.Check1.value = vbUnchecked Then
        With Report
            .Database.AddADOCommand dbConn, adoCommand

            .txtJudul.SetText "DATA KEADAAN MORBIDITAS PASIEN RAWAT INAP RUMAH SAKIT"
            .Text1.SetText strNNamaRS
            .txtJudul2.SetText "FORMULIR RL2a"
            .usNoDTD.SetUnboundFieldSource ("{ado.NoDTD}")
            .usGroup.SetUnboundFieldSource ("{ado.Grup}")
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
        End With

        strSQL = "select * " & _
        " from V_Koders  "

        Call msubRecFO(rs, strSQL)

        With Report
            .Text49.SetText rs("NO1")
            .Text50.SetText rs("NO2")
            .Text51.SetText rs("NO3")
            .Text52.SetText rs("NO4")
            .Text53.SetText rs("NO5")
            .Text54.SetText rs("NO6")
            .Text55.SetText rs("NO7")
        End With
        If vLaporan = "view" Then
            Screen.MousePointer = vbHourglass
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
    End If
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
    Set frmRL2A = Nothing
End Sub

