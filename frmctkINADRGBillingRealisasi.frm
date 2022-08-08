VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmctkINADRGBillingRealisasi 
   Caption         =   "Medifirst2000 - Cetak"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmctkINADRGBillingRealisasi.frx":0000
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
Attribute VB_Name = "frmctkINADRGBillingRealisasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReportBulan2 As New crINADRGBillingRealisasi

Private Sub Form_Load()
    On Error GoTo errLoad

    Dim tanggal As String
    Dim laporan As String
    Dim adocomd As New ADODB.Command

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    Call openConnection

    adocomd.ActiveConnection = dbConn
    If frmINADRGBillingRealisasi.dcInstalasi.BoundText = "02" Then 'RJ
        adocomd.CommandText = "select distinct TglMasuk, TglKeluar, NoCM, NamaPasien, TotalINADRG, TarifINADRG  from V_INADRG_RealisasiOK " & _
        " WHERE JenisPerawatan = '2' and TglKeluar Between '" & Format(frmINADRGBillingRealisasi.dtpAwal.value, "yyyy/MM/dd 00:00:00") & " ' and" & _
        " '" & Format(frmINADRGBillingRealisasi.dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "' order by NoCM"
    End If
    If frmINADRGBillingRealisasi.dcInstalasi.BoundText = "03" Then 'RI
        adocomd.CommandText = "select distinct TglMasuk, TglKeluar, NoCM, NamaPasien, TotalINADRG, TarifINADRG  from V_INADRG_RealisasiOK " & _
        " WHERE JenisPerawatan = '1' and TglKeluar Between '" & Format(frmINADRGBillingRealisasi.dtpAwal.value, "yyyy/MM/dd 00:00:00") & " ' and" & _
        " '" & Format(frmINADRGBillingRealisasi.dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "' order by NoCM"
    End If

    adocomd.CommandType = adCmdText

    tanggal = Format(frmINADRGBillingRealisasi.dtpAwal.value, "MMMM yyyy")

    With ReportBulan2
        .Database.AddADOCommand dbConn, adocomd

        .txtNamaRS.SetText strNNamaRS & " " & strKelasRS & " " & strKetKelasRS
        .txtAlamat.SetText "KABUPATEN " & strNKotaRS
        .txtAlamat2.SetText strNAlamatRS & " " & "Telp." & " " & strNTeleponRS

        .txtInstalasi.SetText strNNamaInstalasi

        .txtTanggalPilih1.SetText Format(frmINADRGBillingRealisasi.dtpAwal.value, "dd MMMM yyyy")
        .txtTanggalPilih2.SetText Format(frmINADRGBillingRealisasi.dtpAkhir.value, "dd MMMM yyyy")

        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
        .unJmlBilling.SetUnboundFieldSource ("{ado.TotalINADRG}")
        .unINADRG.SetUnboundFieldSource ("{ado.TarifINADRG}")
    End With
    CRViewer1.ReportSource = ReportBulan2
    With CRViewer1
        .Zoom 1
        .ViewReport
    End With
    Screen.MousePointer = vbDefault

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
    Set frmctkINADRGBillingRealisasi = Nothing
End Sub

