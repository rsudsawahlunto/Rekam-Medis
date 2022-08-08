VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakKondisiBarang 
   Caption         =   "Medifrst2000 - Laporan Kondisi Barang Non Medis"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakKondisiBarang.frx":0000
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
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakKondisiBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rptKondisiBarang As New CrKondisiBarang

Private Sub Form_Load()

    Dim adocmd As New ADODB.Command

    On Error GoTo errLoad

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    Set adocmd.ActiveConnection = dbConn
    
    strSQL = "SELECT NamaBarang AS [Nama Barang], NamaAsal AS [Asal], DetailJenisBarang AS [Jenis Barang], Kondisi, JmlBarang, KdBarang, KdAsal, KdDetailJenisBarang, KdRuangan, KdKondisi, NoRegisterAsset " & _
    " FROM V_KondisiBarangNonMedis1 " & _
    " WHERE NamaBarang LIKE '%" & frmKondisiBarangNM.txtCariBarang & "%' AND KdRuangan = '" & mstrKdRuangan & "'"

    With rptKondisiBarang
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strWebsite & ", " & strEmail

        adocmd.CommandText = strSQL
        adocmd.CommandType = adCmdText
        .Database.AddADOCommand dbConn, adocmd

        .usNamaBarang.SetUnboundFieldSource ("{ado.Nama Barang}")
        .usAsal.SetUnboundFieldSource ("{ado.Asal}")
        .usJenisBahan.SetUnboundFieldSource ("{ado.Jenis Barang}")
        .usKondisi.SetUnboundFieldSource ("{ado.Kondisi}")
        .unJml.SetUnboundFieldSource ("{ado.JmlBarang}")
        .usNoRegisterAsset.SetUnboundFieldSource ("{ado.NoRegisterAsset}")
    End With

    Screen.MousePointer = vbDefault

    If vLaporan = "view" Then
        CRViewer1.ReportSource = rptKondisiBarang
        CRViewer1.Zoom 1
        CRViewer1.ViewReport
    Else
        rptKondisiBarang.PrintOut False
        Unload Me
    End If
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
    Set frmCetakKondisiBarang = Nothing
End Sub

