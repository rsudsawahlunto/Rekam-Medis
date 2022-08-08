VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakNilaiPersediaan 
   Caption         =   "Form Cetak Nilai Persediaan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakNilaiPersediaan.frx":0000
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
Attribute VB_Name = "frmCetakNilaiPersediaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crCetakNilaiPersediaan
Dim Report1 As New crCetakNilaiPersediaanNM
Public chkHarga As Integer

Private Sub Form_Load()
    On Error GoTo errLoad
    Dim adocomd As New ADODB.Command

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    With Report
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        .txtPeriodeClosing.SetText "Periode Closing " & frmNilaiPersediaan.dcNoClosing.Text
        .txtNamaRuangan.SetText mstrNamaRuangan
    End With

    strSQL = ""
    If frmNilaiPersediaan.ChkPerjenis.value = 0 Then
        strSQL = "SELECT * FROM V_DataStokBarangMedisRekap " & _
        " WHERE KdRuangan = '" & mstrKdRuangan & "' AND (TglClosing = '" & Format(frmNilaiPersediaan.dcNoClosing.BoundText, "yyyy/MM/dd hh:mm:ss") & "') AND StokReal<> 0" & _
        " ORDER By JenisBarang, NamaBarang"
    Else
        strSQL = "SELECT * FROM V_DataStokBarangMedisRekap " & _
        " WHERE JenisBarang like '" & frmNilaiPersediaan.dcJenisBarang.Text & "%' and KdRuangan = '" & mstrKdRuangan & "' AND (TglClosing = '" & Format(frmNilaiPersediaan.dcNoClosing.BoundText, "yyyy/MM/dd HH:mm:ss") & "')" & _
        " ORDER By JenisBarang, NamaBarang"
    End If

    Set dbcmd = New ADODB.Command

    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText

    Report.Database.AddADOCommand dbConn, dbcmd
    With Report
        .usNoTerima.SetUnboundFieldSource ("{ado.NoTerima}")
        .usNamaBarang.SetUnboundFieldSource ("{ado.NamaBarang}")
        .usAsalBarang.SetUnboundFieldSource ("{ado.AsalBarang}")
        .udTglKadaluarsa.SetUnboundFieldSource ("{ado.TglKadaluarsa}")
        .unStok.SetUnboundFieldSource ("{ado.StokReal}")
        If chkHarga = 1 Then
            .Text44.Suppress = False
            .ucHargaNetto2.Suppress = False
            .ucHargaNetto2.SetUnboundFieldSource ("{ado.HargaNetto1}")
            .Text9.Suppress = False
            .ucTotal.Suppress = False
            .Field1.Suppress = False
            .Field6.Suppress = False
            .ucTotal.SetUnboundFieldSource ("{ado.TotalNetto1}")
        Else
            .Text44.Suppress = True
            .ucHargaNetto2.Suppress = True
            .ucHargaNetto2.SetUnboundFieldSource ("{ado.HargaNetto1}")
            .Text9.Suppress = True
            .ucTotal.Suppress = True
            .Field1.Suppress = True
            .Field6.Suppress = True
            .ucTotal.SetUnboundFieldSource ("{ado.TotalNetto1}")
        End If
    End With

    With CRViewer1
        .EnableGroupTree = True
        .ReportSource = Report
        .ViewReport
        .Zoom 1
    End With

    Screen.MousePointer = vbDefault
    Exit Sub
errLoad:
    Screen.MousePointer = vbDefault
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakNilaiPersediaan = Nothing
    strCetak = ""
End Sub

