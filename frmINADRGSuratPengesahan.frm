VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmINADRGSuratPengesahan 
   Caption         =   "Surat Pengesahan"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmINADRGSuratPengesahan.frx":0000
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
Attribute VB_Name = "frmINADRGSuratPengesahan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim report2 As New crDetailINADRGSP

Private Sub Form_Load()
    On Error GoTo hell
    Dim dAwal As Date
    Dim dAkhir As Date
    Dim d As Integer

    Call openConnection

    Set frmINADRGSuratPengesahan = Nothing
    Set dbcmd = New ADODB.Command
    Set dbcmd = Nothing

    Me.WindowState = 2

    dbcmd.ActiveConnection = dbConn
    dbcmd.CommandText = "Select * From V_LaporanINADRGJEP2009 where TglKeluar between '" & Format(frmINADRGLoadSuratPengesahan.dtpAwal.value, "yyyy-MM-dd 00:00:00.000") & "' and '" & Format(frmINADRGLoadSuratPengesahan.dtpAkhir.value, "yyyy-MM-dd HH:mm:ss.999") & "' ORDER BY TglKeluar,DESKRIPSI"
    dbcmd.CommandType = adCmdText

    If rsProfil.State = adStateOpen Then rsProfil.Close
    rsProfil.Open "select * from ProfilRS", dbConn, adOpenForwardOnly, adLockReadOnly

    With report2
        .Database.AddADOCommand dbConn, dbcmd

        .txtNamaRS.SetText strNNamaRS
        .txtAlamatRS.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .txtKelas.SetText strKelasRS
        .Text53.SetText strKetKelasRS & " " & strNKotaRS & " Menerangkan Pasien atas nama :"
        .Text71.SetText "Benar di rawat di Rumah Sakit Umum Daerah Kelas " & strKelasRS & " " & strKetKelasRS & " " & strNKotaRS & " dengan"
        .dtTglKeluar.SetText strNKotaRS & ", " & Format(Date, "dd mmmm yyyy")

        .usNoRM.SetUnboundFieldSource ("{ado.NoRM}")
        .unUmurThn.SetUnboundFieldSource ("{ado.UmurThn}")
        .unUmurHr.SetUnboundFieldSource ("{ado.UmurHari}")
        .strTglLahir.SetUnboundFieldSource ("{ado.TglLahir}")
        .strTglMasuk.SetUnboundFieldSource ("{ado.TglMasuk}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaLengkap}")
        .usKampung.SetUnboundFieldSource ("{ado.Alamat}")
        .usDesa.SetUnboundFieldSource ("{ado.Kelurahan}")
        .usRT.SetUnboundFieldSource ("{ado.RTRW}")
        .usKecamatan.SetUnboundFieldSource ("{ado.Kecamatan}")
        .usKabupaten.SetUnboundFieldSource ("{ado.Kota}")

        .usDX1.SetUnboundFieldSource ("{ado.ICD10_1}")
        .UnboundString1.SetUnboundFieldSource ("{ado.ICD10_2}")
        .UnboundString2.SetUnboundFieldSource ("{ado.ICD10_3}")
        .UnboundString3.SetUnboundFieldSource ("{ado.ICD10_4}")
        .UnboundString30.SetUnboundFieldSource ("{ado.ICD9_1}")
        .UnboundString31.SetUnboundFieldSource ("{ado.ICD9_2}")
        .UnboundString32.SetUnboundFieldSource ("{ado.ICD9_3}")
        .UnboundString60.SetUnboundFieldSource ("{ado.NamaDiagnosa}")
        .UnboundString61.SetUnboundFieldSource ("{ado.NamaDiagnosa1}")
        .UnboundString62.SetUnboundFieldSource ("{ado.NamaDiagnosa2}")
        .UnboundString63.SetUnboundFieldSource ("{ado.NamaDiagnosa3}")
        .UnboundString64.SetUnboundFieldSource ("{ado.DiagnosaTindakan}")
        .UnboundString65.SetUnboundFieldSource ("{ado.DiagnosaTindakan1}")
        .UnboundString66.SetUnboundFieldSource ("{ado.DiagnosaTindakan2}")
    End With

    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = report2
    CRViewer1.ViewReport
    CRViewer1.Zoom 1
    Screen.MousePointer = vbDefault
    Exit Sub
hell:
    Screen.MousePointer = vbDefault
    Set frmRekapPasienINADRG = Nothing
    Set dbcmd = Nothing
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmINADRGSuratPengesahan = Nothing
End Sub
