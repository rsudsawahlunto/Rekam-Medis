VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakDataRiwayatPemeriksaanPasien 
   Caption         =   "Medifirst2000 - Laporan"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmCetakDataRiwayatPemeriksaanPasien.frx":0000
   KeyPreview      =   -1  'True
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakDataRiwayatPemeriksaanPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crCetakDataRiwayatPemeriksaanPasien

Private Sub Form_Load()
    On Error GoTo hell

    Call openConnection
    Set dbcmd = New ADODB.Command
    Set frmCetakDataRiwayatPemeriksaanPasien = Nothing
    Me.WindowState = 2

    With dbcmd
        .ActiveConnection = dbConn
        .CommandText = "SELECT distinct * FROM V_DataDetailRiwayatPemeriksaanPasien where NoPendaftaran='" & mstrNoPen & "' AND NoCM='" & mstrNoCM & "'"
        .CommandType = adCmdText
    End With

    strSQL = "SELECT * FROM V_DataRiwayatPemeriksaanPasienN where NoPendaftaran='" & mstrNoPen & "' AND NoCM='" & mstrNoCM & "'"
    Set dbRst = Nothing
    Call msubRecFO(dbRst, strSQL)

    With Report
        .Database.AddADOCommand dbConn, dbcmd

        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtWebsite.SetText strWebsite & ", " & strEmail

        If dbRst.EOF = False Then
            .txtNoCM.SetText IIf(IsNull(dbRst("NoCM")), "", dbRst("NoCM"))
            .txtNamaPasien.SetText IIf(IsNull(dbRst("NamaPasien")), "", dbRst("NamaPasien"))
            .txtAlamatPasien.SetText IIf(IsNull(dbRst("Alamat")), "", dbRst("Alamat"))
            .txtnopendaftaran.SetText IIf(IsNull(dbRst("NoPendaftaran")), "", dbRst("NoPendaftaran"))
            .txtTmptTglLahir.SetText IIf(IsNull(dbRst("TempatTglLahir")), "", dbRst("TempatTglLahir")) & " ," & Format(dbRst("TglLahir"), "DD/mm/yyyy")
            .txtumur.SetText IIf(IsNull(dbRst("Umur")), "", dbRst("Umur"))
        End If

        .usJudul.SetUnboundFieldSource ("{ado.Judul}")
        .usA.SetUnboundFieldSource ("{ado.A}")
        .usB.SetUnboundFieldSource ("{ado.B}")
        .usC.SetUnboundFieldSource ("{ado.C}")
        .usD.SetUnboundFieldSource ("{ado.D}")
        .usE.SetUnboundFieldSource ("{ado.E}")
        .usF.SetUnboundFieldSource ("{ado.F}")
        .usG.SetUnboundFieldSource ("{ado.G}")
        .usH.SetUnboundFieldSource ("{ado.H}")
        .usI.SetUnboundFieldSource ("{ado.I}")
        .usJ.SetUnboundFieldSource ("{ado.J}")
    End With

    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = Report
    If vLaporan = "view" Then
        With CRViewer1

            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault

    Set dbcmd = Nothing

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
    Set frmCetakDataRiwayatPemeriksaanPasien = Nothing
End Sub
