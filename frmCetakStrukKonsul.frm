VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakStrukKonsul 
   Caption         =   "Cetak Struk Konsul"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCetakStrukKonsul.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   6765
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
Attribute VB_Name = "frmCetakStrukKonsul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crCetakStrukKonsul

Private Sub Form_Load()
On Error GoTo errLoad
    Dim adocomd As New ADODB.Command

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Call openConnection

    strSQL = "SELECT MessageToDay FROM MasterDataPendukung"
    Call msubRecFO(rs, strSQL)

    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = "SELECT * from V_CetakStrukPasienKonsul " _
                         & " WHERE (NoPendaftaran) =('" & mstrNoPen & "')  AND [Ruangan Perujuk] = '" & frmTransaksiPasien.dgKonsul.Columns("Ruangan Perujuk") & "' AND [Ruangan Tujuan] = '" & frmTransaksiPasien.dgKonsul.Columns("Ruangan Tujuan") & "'"
    adocomd.CommandType = adCmdText
    
    Report.Database.AddADOCommand dbConn, adocomd
    
    Dim tanggal As String
    tanggal = Format(TglPeriodeAwal, "MMMM yyyy")
    
    With Report
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS
        .Text3.SetText strEmail
        .udtgl.SetUnboundFieldSource ("{ado.TglDirujuk}")
        .usnodft.SetUnboundFieldSource ("{ado.NoPendaftaran}")
        .usNoCM.SetUnboundFieldSource ("{Ado.nocm}")
        .usAlias.SetUnboundFieldSource ("{Ado.Alias}")
        .usAlamat.SetUnboundFieldSource ("{Ado.Alamat}")
        .usPenjamin.SetUnboundFieldSource ("{Ado.JenisPasien}")
        .USJenisKelas.SetUnboundFieldSource ("{ado.Kelas}")
        .usruangperiksa.SetUnboundFieldSource ("{Ado.Ruangan Perujuk}")
        .usRuanganTujuan.SetUnboundFieldSource ("{Ado.Ruangan Tujuan}")
        .usDokterPerujuk.SetUnboundFieldSource ("{Ado.Dokter Perujuk}")
        
        .usnoantri.SetUnboundFieldSource ("{Ado.No. Urut}")
        .ususer.SetText (strNmPegawai)
        .txtMessage.SetText (rs(0).value)
        .PrintOut False
    End With

        Screen.MousePointer = vbHourglass
            If vLaporan = "view" Then
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

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakStrukKonsul = Nothing
End Sub
