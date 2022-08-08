VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakSuratKeterangan3 
   Caption         =   "Kondisi Barang Non Medis"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakSuratKeterangan3.frx":0000
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
Attribute VB_Name = "frmCetakSuratKeterangan3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rptSuratKeterangan As New crSuratKeterangan3

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    On Error GoTo errLoad

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    With rptSuratKeterangan

        .txtTanggal2.SetText Format(frmSuratKeterangan3C.dtpAwal, "dd MMMM yyyy")

        .txtNama.SetText (frmSuratKeterangan3C.txtNama.Text)

        If frmSuratKeterangan3C.txtSex.Text = "L" Then
            .txtJenisKelamin.SetText "Laki-Laki"
        Else
            .txtJenisKelamin.SetText "Perempuan "
        End If

        .txtTtl.SetText (frmSuratKeterangan3C.txtAlamat.Text)

        .txtTindakan.SetText (frmSuratKeterangan3C.txtTindakan.Text)
        .txtExpertise.SetText (frmSuratKeterangan3C.txtKesimpulan2)

        .txtKiriman.SetText (frmSuratKeterangan3C.txtKiriman.Text)
        .txtNoPendaftaran.SetText (frmSuratKeterangan3C.txtNoCM.Text)
        .txtAdress.SetText (frmSuratKeterangan3C.txtAlamat.Text)
    End With

    CRViewer1.ReportSource = rptSuratKeterangan
    CRViewer1.Zoom 1
    CRViewer1.ViewReport
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
