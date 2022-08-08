VERSION 5.00
Begin VB.Form frm_cetak_KesehatanJiwa 
   Caption         =   "Laporan Cetak Kesehatan Jiwa"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frm_cetak_KesehatanJiwa.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.PictureBox CRViewer1 
      Height          =   7005
      Left            =   0
      ScaleHeight     =   6945
      ScaleWidth      =   5745
      TabIndex        =   0
      Top             =   0
      Width           =   5805
   End
End
Attribute VB_Name = "frm_cetak_KesehatanJiwa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crKesehatanJiwa

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
Me.WindowState = 2
Set frm_cetak_KesehatanJiwa = Nothing

Dim adocomd As New ADODB.Command
Call openConnection

adocomd.ActiveConnection = dbConn

adocomd.CommandText = strSQL

adocomd.CommandType = adCmdText
Report.Database.AddADOCommand dbConn, adocomd


 With Report
            
    .txtNamaRS.SetText strNNamaRS
    .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    '.txtWebsite.SetText strWebsite & ", " & strEmail

    .txtTanggal.SetText Format(frmKesehatanJiwa.DTPickerAwal.Value, "dd/MM/yyyy")
    .txtTanggal2.SetText Format(frmKesehatanJiwa.DTPickerAkhir.Value, "dd/MM/yyyy")

'Report.Text3.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS

'    .usJenisPemeriksaan.SetUnboundFieldSource ("{Ado.}")
'    .usSederhana.SetUnboundFieldSource ("{Ado.}")
'    .usSedang.SetUnboundFieldSource ("{Ado.}")
'    .usCanggih.SetUnboundFieldSource ("{Ado.}")
'    .ucTotal.SetUnboundFieldSource ("{Ado.}")

End With
    
    CRViewer1.ReportSource = Report


    Screen.MousePointer = vbHourglass

With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 1
End With

    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm_cetak_KesehatanJiwa = Nothing
End Sub
