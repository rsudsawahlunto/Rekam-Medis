VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakSuratKeterangan 
   Caption         =   "Cetak Surat Keterangan"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "frmCetakSuratKeterangan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   11280
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
Attribute VB_Name = "frmCetakSuratKeterangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crSuratKeterangan

Private Sub Form_Load()
    On Error GoTo errLoad
    Dim adocomd As New ADODB.Command
    Dim bln As String

    bln = Format(Now, "MM")
    Select Case bln
        Case "01"
            bln = "I"
        Case "02"
            bln = "II"
        Case "03"
            bln = "III"
        Case "04"
            bln = "IV"
        Case "05"
            bln = "V"
        Case "06"
            bln = "VI"
        Case "07"
            bln = "VII"
        Case "08"
            bln = "VIII"
        Case "09"
            bln = "IX"
        Case "10"
            bln = "X"
        Case "11"
            bln = "XI"
        Case "12"
            bln = "XII"
    End Select

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Call openConnection
    Set Report = Nothing
    adocomd.ActiveConnection = dbConn
    strSQL = "SELECT '   /' + NoCM as NoSurat,  TglMasuk, TglPulang, NamaPasien, JK, Umur, Pekerjaan, Jln, RTRW, Desa," & _
    "Kec,Kota,Propinsi, DokterPemeriksa from V_InfoSuratKeterangan " & _
    " WHERE (NoPendaftaran) =('" & mstrNoPen & "')"
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText
    Set rs = Nothing
    Set rs = dbConn.Execute(strSQL)
    Report.Database.AddADOCommand dbConn, adocomd

    With Report
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text35.SetText strKelasRS & " " & strKetKelasRS
        .Text6.SetText strNKotaRS & ", menerangkan dengan sesungguhnya, bahwa:"
        .Text22.SetText "dirawat di RSUD " & strNKotaRS & ","
        .Text25.SetText strNKotaRS & ", " & Format(Date, "dd mmmm yyyy")

        If IsNull(rs("TglPulang")) Then
            .TxtPulang.Font.Strikethrough = True
        Else
            .TxtDirawat.Font.Strikethrough = True
        End If
        .udTglMasuk.SetUnboundFieldSource ("{ado.TglMasuk}")
        .unnoCM.SetUnboundFieldSource ("{ado.NoSurat}")
        .usDokterPemeriksa.SetText mstrNamaDokter
        .txtTahun.SetText Format(Now, "yyyy")
        .txtBln.SetText bln
        .usNama.SetUnboundFieldSource ("{ado.NamaPasien}")
       ' .ucNamaDokter.SetText mstrNamaDokter
        .usJK.SetUnboundFieldSource ("{Ado.JK}")
        .usUmur.SetUnboundFieldSource ("{Ado.Umur}")
        .UnJln.SetUnboundFieldSource ("{ado.Jln}")
        .unRT.SetUnboundFieldSource ("{ado.RTRW}")
        .unDesa.SetUnboundFieldSource ("{ado.Desa}")
        .UnKec.SetUnboundFieldSource ("{ado.Kec}")
        .usAlamat.SetUnboundFieldSource ("{Ado.Kota}")
        .usPekerjaan.SetUnboundFieldSource ("{Ado.Pekerjaan}")
        '.usDokterPemeriksa.SetUnboundFieldSource ("{Ado.DokterPemeriksa}")
    End With

    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
    CRViewer1.EnableGroupTree = False
    CRViewer1.Zoom (100)
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

