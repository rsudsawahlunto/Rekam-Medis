VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakOdontoGram 
   Caption         =   "Medifirst2000 - Cetak Odontogram"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5880
   Icon            =   "frmCetakOdontoGram.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   5880
   WindowState     =   2  'Maximized
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
Attribute VB_Name = "frmCetakOdontoGram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fso As New Scripting.FileSystemObject
Dim rpt As crOdontoGram
Dim WithEvents sect As CRAXDRT.Section
Attribute sect.VB_VarHelpID = -1

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo tangani
    fso.DeleteFile (App.Path & "\tempbitmap.bmp")
    Set sect = Nothing
    Exit Sub
tangani:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Dim adocomd As New ADODB.Command
    Call openConnection

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

    Set rpt = New crOdontoGram
    Set sect = rpt.Sections.Item("Section10")

    rpt.txtNamaRS.SetText strNNamaRS
    rpt.txtAlamatRS.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    rpt.txtWebsiteRS.SetText strWebsite & ", " & strEmail

    strSQL = "select * from V_CetakOdontoGram where NoPendaftaran='" & mstrNoPen & "'"

    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText

    rpt.Database.AddADOCommand dbConn, adocomd

    With rpt
        .txtTglPendaftaran.SetText frmDiagramOdonto.txtTglDaftar.Text
        .txtUmur.SetText frmDiagramOdonto.txtThn.Text & " thn"
        .txtKelasPelayanan.SetText frmDiagramOdonto.txtKls.Text
        .txtJenisPasien.SetText frmDiagramOdonto.txtJenisPasien.Text

        .txtNoSurat.SetText "Nomor:                 /RSUD/" + bln + "/" + Trim(str(Year(Now)))

'        .txtPekerjaan.SetText ctkHasilPekerjaan
'        .txtAlamat.SetText ctkHasilTHTAlamat
        .UnAlamat.SetUnboundFieldSource ("{ado.Alamat}")
        .UnPekerjaan.SetUnboundFieldSource ("{ado.Pekerjaan}")

        .usNoPendaftaran.SetUnboundFieldSource ("{ado.NoPendaftaran}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .udTglPeriksa.SetUnboundFieldSource ("{ado.TglPeriksa}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
        .usJK.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .usDokterPeriksa.SetUnboundFieldSource ("{ado.NamaDokter}")
        .usRuangan.SetUnboundFieldSource ("{ado.NamaRuangan}")
        .usInstalasi.SetUnboundFieldSource ("{ado.NamaSubInstalasi}")
        .usKeterangan.SetUnboundFieldSource ("{ado.Keterangan}")
    End With
    CRViewer1.ReportSource = rpt
    CRViewer1.ViewReport
    CRViewer1.Zoom 1
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub sect_Format(ByVal pFormattingInfo As Object)
    Dim bmp As StdPicture
    With sect.ReportObjects
        Set .Item("picOdonto").FormattedPicture = LoadPicture(App.Path & "\tempbitmap.bmp") 'default
    End With

End Sub

