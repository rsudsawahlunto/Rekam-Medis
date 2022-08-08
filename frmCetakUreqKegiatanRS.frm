VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakUreqKegiatanRS 
   Caption         =   "Medifirst2000 - Cetak"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakUreqKegiatanRS.frx":0000
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
Attribute VB_Name = "frmCetakUreqKegiatanRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoCommand As New ADODB.Command
Dim ReportBulan As New crUreqDataKegiatanRS

Private Sub Form_Load()
    On Error GoTo errLoad
    openConnection

    Set frmCetakUreqKegiatanRS = Nothing

    Me.WindowState = 2

    adoCommand.CommandText = strSQL
    adoCommand.CommandType = adCmdText

    'Triwulan 1
    If frmUreqKegiatanRS.Check1.value = vbChecked And frmUreqKegiatanRS.Option1.value = True Then
        With ReportBulan
            .Database.AddADOCommand dbConn, adoCommand

            .txtjudul.SetText "DATA KEGIATAN RUMAH SAKIT"
            .Text51.SetText " I "
            .Text30.SetText "FORMULIR RI 1"
            .Text47.SetText strNNamaRS
            .Text48.SetText "1. PELAYANAN RAWAT INAP"

            .UsKdSubInstalasi.SetUnboundFieldSource ("{ado.KdSubInstalasi}")
            .usSubInstalasi.SetUnboundFieldSource ("{ado.2}")

            .UnboundNumber3.SetUnboundFieldSource Format("{ado.3}")
            .UnboundNumber4.SetUnboundFieldSource ("{ado.4}")
            .UnboundNumber5.SetUnboundFieldSource ("{ado.5}")
            .UnboundNumber6.SetUnboundFieldSource ("{ado.6}")
            .UnboundNumber7.SetUnboundFieldSource ("{ado.7}")
            .UnboundNumber8.SetUnboundFieldSource ("{ado.8}")
            .UnboundNumber9.SetUnboundFieldSource ("{ado.9}")
            .UnboundNumber10.SetUnboundFieldSource Format("({ado.3}+ {ado.4}) - ({ado.5} + {ado.8})")
            .UnboundNumber11.SetUnboundFieldSource ("{ado.11}")
            .UnboundNumber12.SetUnboundFieldSource ("{ado.12}")
            .UnboundNumber13.SetUnboundFieldSource ("{ado.13}")
            .UnboundNumber14.SetUnboundFieldSource ("{ado.14}")
            .UnboundNumber15.SetUnboundFieldSource ("{ado.15}")
            .UnboundNumber16.SetUnboundFieldSource ("{ado.16}")
        End With
        strSQL = "select * " & _
        " from V_Koders  "

        Call msubRecFO(rs, strSQL)

        With ReportBulan
            .Text49.SetText rs("NO1")
            .Text9.SetText rs("NO2")
            .Text11.SetText rs("NO3")
            .Text24.SetText rs("NO4")
            .Text25.SetText rs("NO5")
            .Text26.SetText rs("NO6")
            .Text27.SetText rs("NO7")
        End With

        Screen.MousePointer = vbHourglass
        If vLaporan = "view" Then
            With CRViewer1
                .ReportSource = ReportBulan
                .ViewReport
                .Zoom 1
            End With
        Else
            ReportBulan.PrintOut False
            Unload Me
        End If
        Screen.MousePointer = vbDefault
    End If

    'Triwulan 2
    If frmUreqKegiatanRS.Check1.value = vbChecked And frmUreqKegiatanRS.Option2.value = True Then
        With ReportBulan
            .Database.AddADOCommand dbConn, adoCommand

            .txtjudul.SetText "DATA KEGIATAN RUMAH SAKIT"
            .Text51.SetText " II "
            .Text30.SetText "FORMULIR RI 1"
            .Text47.SetText strNNamaRS
            .Text48.SetText "1. PELAYANAN RAWAT INAP"
            .UsKdSubInstalasi.SetUnboundFieldSource ("{ado.KdSubInstalasi}")
            .usSubInstalasi.SetUnboundFieldSource ("{ado.2}")

            .UnboundNumber3.SetUnboundFieldSource ("{ado.3}")
            .UnboundNumber4.SetUnboundFieldSource ("{ado.4}")
            .UnboundNumber5.SetUnboundFieldSource ("{ado.5}")
            .UnboundNumber6.SetUnboundFieldSource ("{ado.6}")
            .UnboundNumber7.SetUnboundFieldSource ("{ado.7}")
            .UnboundNumber8.SetUnboundFieldSource ("{ado.8}")
            .UnboundNumber9.SetUnboundFieldSource ("{ado.9}")
            .UnboundNumber10.SetUnboundFieldSource Format("({ado.3}+ {ado.4}) - ({ado.5} + {ado.8})")
            .UnboundNumber11.SetUnboundFieldSource ("{ado.11}")
            .UnboundNumber12.SetUnboundFieldSource ("{ado.12}")
            .UnboundNumber13.SetUnboundFieldSource ("{ado.13}")
            .UnboundNumber14.SetUnboundFieldSource ("{ado.14}")
            .UnboundNumber15.SetUnboundFieldSource ("{ado.15}")
            .UnboundNumber16.SetUnboundFieldSource ("{ado.16}")
        End With
        strSQL = "select * " & _
        " from V_Koders  "

        Call msubRecFO(rs, strSQL)

        With ReportBulan
            .Text49.SetText rs("NO1")
            .Text9.SetText rs("NO2")
            .Text11.SetText rs("NO3")
            .Text24.SetText rs("NO4")
            .Text25.SetText rs("NO5")
            .Text26.SetText rs("NO6")
            .Text27.SetText rs("NO7")
        End With

        Screen.MousePointer = vbHourglass
        If vLaporan = "view" Then
            With CRViewer1
                .ReportSource = ReportBulan
                .ViewReport
                .Zoom 1
            End With
        Else
            ReportBulan.PrintOut False
            Unload Me
        End If
        Screen.MousePointer = vbDefault
    End If

    'Triwulan 3
    If frmUreqKegiatanRS.Check1.value = vbChecked And frmUreqKegiatanRS.Option3.value = True Then
        With ReportBulan
            .Database.AddADOCommand dbConn, adoCommand

            .txtjudul.SetText "DATA KEGIATAN RUMAH SAKIT"
            .Text51.SetText " III "
            .Text30.SetText "FORMULIR RI 1"
            .Text47.SetText strNNamaRS
            .Text48.SetText "1. PELAYANAN RAWAT INAP"
            .UsKdSubInstalasi.SetUnboundFieldSource ("{ado.KdSubInstalasi}")
            .usSubInstalasi.SetUnboundFieldSource ("{ado.2}")

            .UnboundNumber3.SetUnboundFieldSource Format("{ado.3}")
            .UnboundNumber4.SetUnboundFieldSource ("{ado.4}")
            .UnboundNumber5.SetUnboundFieldSource ("{ado.5}")
            .UnboundNumber6.SetUnboundFieldSource ("{ado.6}")
            .UnboundNumber7.SetUnboundFieldSource ("{ado.7}")
            .UnboundNumber8.SetUnboundFieldSource ("{ado.8}")
            .UnboundNumber9.SetUnboundFieldSource ("{ado.9}")
            .UnboundNumber10.SetUnboundFieldSource Format("({ado.3}+ {ado.4}) - ({ado.5} + {ado.8})")
            .UnboundNumber11.SetUnboundFieldSource ("{ado.11}")
            .UnboundNumber12.SetUnboundFieldSource ("{ado.12}")
            .UnboundNumber13.SetUnboundFieldSource ("{ado.13}")
            .UnboundNumber14.SetUnboundFieldSource ("{ado.14}")
            .UnboundNumber15.SetUnboundFieldSource ("{ado.15}")
            .UnboundNumber16.SetUnboundFieldSource ("{ado.16}")
        End With
        strSQL = "select * " & _
        " from V_Koders  "

        Call msubRecFO(rs, strSQL)

        With ReportBulan
            .Text49.SetText rs("NO1")
            .Text9.SetText rs("NO2")
            .Text11.SetText rs("NO3")
            .Text24.SetText rs("NO4")
            .Text25.SetText rs("NO5")
            .Text26.SetText rs("NO6")
            .Text27.SetText rs("NO7")
        End With

        Screen.MousePointer = vbHourglass
        If vLaporan = "view" Then
            With CRViewer1
                .ReportSource = ReportBulan
                .ViewReport
                .Zoom 1
            End With
        Else
            ReportBulan.PrintOut False
            Unload Me
        End If
        Screen.MousePointer = vbDefault
    End If

    'Triwulan 4
    If frmUreqKegiatanRS.Check1.value = vbChecked And frmUreqKegiatanRS.Option4.value = True Then
        With ReportBulan
            .Database.AddADOCommand dbConn, adoCommand

            .txtjudul.SetText "DATA KEGIATAN RUMAH SAKIT"
            .Text51.SetText " IV "
            .Text30.SetText "FORMULIR RI 1"
            .Text47.SetText strNNamaRS
            .Text48.SetText "1. PELAYANAN RAWAT INAP"
            .UsKdSubInstalasi.SetUnboundFieldSource ("{ado.KdSubInstalasi}")
            .usSubInstalasi.SetUnboundFieldSource ("{ado.2}")

            .UnboundNumber3.SetUnboundFieldSource Format("{ado.3}")
            .UnboundNumber4.SetUnboundFieldSource ("{ado.4}")
            .UnboundNumber5.SetUnboundFieldSource ("{ado.5}")
            .UnboundNumber6.SetUnboundFieldSource ("{ado.6}")
            .UnboundNumber7.SetUnboundFieldSource ("{ado.7}")
            .UnboundNumber8.SetUnboundFieldSource ("{ado.8}")
            .UnboundNumber9.SetUnboundFieldSource ("{ado.9}")
            .UnboundNumber10.SetUnboundFieldSource Format("({ado.3}+ {ado.4}) - ({ado.5} + {ado.8})")
            .UnboundNumber11.SetUnboundFieldSource ("{ado.11}")
            .UnboundNumber12.SetUnboundFieldSource ("{ado.12}")
            .UnboundNumber13.SetUnboundFieldSource ("{ado.13}")
            .UnboundNumber14.SetUnboundFieldSource ("{ado.14}")
            .UnboundNumber15.SetUnboundFieldSource ("{ado.15}")
            .UnboundNumber16.SetUnboundFieldSource ("{ado.16}")
        End With
        strSQL = "select * " & _
        " from V_Koders  "

        Call msubRecFO(rs, strSQL)

        With ReportBulan
            .Text49.SetText rs("NO1")
            .Text9.SetText rs("NO2")
            .Text11.SetText rs("NO3")
            .Text24.SetText rs("NO4")
            .Text25.SetText rs("NO5")
            .Text26.SetText rs("NO6")
            .Text27.SetText rs("NO7")
        End With

        Screen.MousePointer = vbHourglass
        If vLaporan = "view" Then
            With CRViewer1
                .ReportSource = ReportBulan
                .ViewReport
                .Zoom 1
            End With
        Else
            ReportBulan.PrintOut False
            Unload Me
        End If
        Screen.MousePointer = vbDefault
    End If

    'Periode
    If frmUreqKegiatanRS.Check1.value = vbUnchecked Then
        With ReportBulan
            .Database.AddADOCommand dbConn, adoCommand

            .txtjudul.SetText "DATA KEGIATAN RUMAH SAKIT"
            .Text30.SetText "FORMULIR RI 1"
            .Text47.SetText strNNamaRS
            .Text48.SetText "1. PELAYANAN RAWAT INAP"
            .UsKdSubInstalasi.SetUnboundFieldSource ("{ado.KdSubInstalasi}")
            .usSubInstalasi.SetUnboundFieldSource ("{ado.2}")

            .UnboundNumber3.SetUnboundFieldSource Format("{ado.3}")
            .UnboundNumber4.SetUnboundFieldSource ("{ado.4}")
            .UnboundNumber5.SetUnboundFieldSource ("{ado.5}")
            .UnboundNumber6.SetUnboundFieldSource ("{ado.6}")
            .UnboundNumber7.SetUnboundFieldSource ("{ado.7}")
            .UnboundNumber8.SetUnboundFieldSource ("{ado.8}")
            .UnboundNumber9.SetUnboundFieldSource ("{ado.9}")
            .UnboundNumber10.SetUnboundFieldSource Format("({ado.3}+ {ado.4}) - ({ado.5} + {ado.8})")
            .UnboundNumber11.SetUnboundFieldSource ("{ado.11}")
            .UnboundNumber12.SetUnboundFieldSource ("{ado.12}")
            .UnboundNumber13.SetUnboundFieldSource ("{ado.13}")
            .UnboundNumber14.SetUnboundFieldSource ("{ado.14}")
            .UnboundNumber15.SetUnboundFieldSource ("{ado.15}")
            .UnboundNumber16.SetUnboundFieldSource ("{ado.16}")
        End With
        strSQL = "select * " & _
        " from V_Koders  "

        Call msubRecFO(rs, strSQL)

        With ReportBulan
            .Text49.SetText rs("NO1")
            .Text9.SetText rs("NO2")
            .Text11.SetText rs("NO3")
            .Text24.SetText rs("NO4")
            .Text25.SetText rs("NO5")
            .Text26.SetText rs("NO6")
            .Text27.SetText rs("NO7")
        End With

        Screen.MousePointer = vbHourglass
        If vLaporan = "view" Then
            With CRViewer1
                .ReportSource = ReportBulan
                .ViewReport
                .Zoom 1
            End With
        Else
            ReportBulan.PrintOut False
            Unload Me
        End If
        Screen.MousePointer = vbDefault
    End If

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
    Set frmCetakUreqKegiatanRS = Nothing
End Sub

