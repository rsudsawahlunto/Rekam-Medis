VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_cetak_label_viewer 
   Caption         =   "Cetal Label Registrasi"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_cetak_label_viewer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   5880
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
Attribute VB_Name = "frm_cetak_label_viewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report2 As New Cr_cetakLabel
Dim Report As New cr_CetakLabel_New
Dim DB As CRekamMedis

Dim p As Printer
Dim p2 As Printer

Private Sub Form_Load()
On Error Resume Next
    Dim adocomd As New ADODB.Command
    
    Set frm_cetak_label_viewer = Nothing
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Call openConnection

    strSQL = "SELECT MessageToDay FROM MasterDataPendukung"
    Call msubRecFO(rs, strSQL)

    adocomd.ActiveConnection = dbConn
    If (mstrKdInstalasi = "02") Or (mstrKdInstalasi = "11") Or (mstrKdInstalasi = "06") Then
        adocomd.CommandText = "SELECT * from V_CetakLabelRegistrasiPasienMRSNew " _
        & " WHERE (NoPendaftaran) =('" & mstrNoPen & "')"
    Else
        adocomd.CommandText = "SELECT * from V_CetakLabelRegistrasiPasienRINew " _
        & " WHERE NoPendaftaran ='" & mstrNoPen & "'"
    End If
    adocomd.CommandType = adCmdText

    Report.Database.AddADOCommand dbConn, adocomd
    
    Dim tanggal As String
    tanggal = Format(TglPeriodeAwal, "MMMM yyyy")

    With Report
        .text1.SetText strNNamaRS
        .text2.SetText strNAlamatRS & ", " & strNKotaRS
        .text3.SetText strEmail
        .udtgl.SetUnboundFieldSource ("{ado.tglmasuk}")
        .usnodft.SetUnboundFieldSource ("{ado.nopendaftaran}")
        .usnocm.SetUnboundFieldSource ("{Ado.nocm}")
        .usnocmerm.SetUnboundFieldSource ("{Ado.nocmterm}")
        .usruangperiksa.SetUnboundFieldSource ("{Ado.Ruang Tujuan}")

        .usnmpasien.SetUnboundFieldSource ("{Ado.nama pasien}")
        .usJK.SetUnboundFieldSource ("{ado.jk}")
        .usTglLahir.SetUnboundFieldSource Format(("{ado.TglLahir}"), "dd-mm-yyyy")
        .usAlamat.SetUnboundFieldSource ("{Ado.Alamat}")
        .unRT.SetUnboundFieldSource ("{Ado.RTRW}")
        .usKelurahan.SetUnboundFieldSource ("{Ado.Kelurahan}")
        .usKecamatan.SetUnboundFieldSource ("{Ado.Kecamatan}")

        .usPenjamin.SetUnboundFieldSource ("{Ado.JenisPasien}")
        .USJenisKelas.SetUnboundFieldSource ("{ado.DetailJenisJasaPelayanan}")
        .usNoBed.SetUnboundFieldSource ("{Ado.NoBed}")

        If mstrKdInstalasi = "03" Then
            .txtNoRuangan.SetText "No. Kamar"
            .usNoRuangan.SetUnboundFieldSource ("{Ado.NamaKamar}")
            .txtNoBed.Suppress = False
            .usNoBed.Suppress = False
        Else
            .txtNoRuangan.SetText "No. Ruangan"
            .usNoRuangan.SetUnboundFieldSource ("{Ado.No. Ruangan}")
            .txtNoBed.Suppress = True
            .usNoBed.Suppress = True
        End If

        .usnoantri.SetUnboundFieldSource ("{Ado.NoAntrian}")
        .ususer.SetText (strNmPegawai)
        .txtMessage.SetText (rs(0).value)
    
        NamaCR = "Cr_cetakLabel"
        If sp_CetakLaporan(NamaCR, mstrNoPen, mstrKdRuangan, strIDPegawai) = False Then Exit Sub
        .txtKendali.SetText strCetakKendaliLaporan & " Cetakan ke: " & intCetakKe
        
' untuk cetak 2 printer sekaligus
    Dim tempPrint1 As String
    Dim strDeviceName As String
    Dim strDriverName As String
    Dim strPort As String
    
'    Dim tempPrint2 As String
'    Dim strDeviceName2 As String
'    Dim strDriverName2 As String
'    Dim strPort2 As String

    
'    tempPrint1 = ReadINI("Default Printer", "Printer1", "", "C:\SettingPrinter.ini")
    
    tempPrint1 = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "PrinterLabel1")
'    tempPrint2 = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "PrinterLabel2 ")
    
'    tempPrint2 = ReadINI("Default Printer", "Printer2", "", "C:\SettingPrinter.ini")

    For Each p In Printers
            strDeviceName = tempPrint1
'            strDeviceName2 = tempPrint2
            strDriverName = p.DriverName
'            strDriverName2 = p.DriverName
            strPort = p.Port
'            strPort2 = p.Port

If tempPrint1 <> "" Then
            Report.SelectPrinter strDriverName, strDeviceName, strPort
            Report.PrintOut False
End If
'            Report.SelectPrinter strDriverName2, strDeviceName2, strPort2
'            Report.PrintOut False

        Exit For

    Next
    
    tempPrint2 = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "PrinterLabel2 ")
    If tempPrint2 = "" Then Exit Sub
    Call cetak_tracer
    Unload Me
    Screen.MousePointer = vbDefault
End With



Exit Sub

End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Width = ScaleWidth
    CRViewer1.Height = ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_cetak_label_viewer = Nothing
End Sub

Private Sub cetak_tracer()
On Error Resume Next
    Dim adocomd As New ADODB.Command
    
'    Set frm_cetak_label_viewer = Nothing
'    Screen.MousePointer = vbHourglass
'    Me.WindowState = 2
    Call openConnection

    strSQL = "SELECT MessageToDay FROM MasterDataPendukung"
    Call msubRecFO(rs, strSQL)

    adocomd.ActiveConnection = dbConn
    If (mstrKdInstalasi = "02") Or (mstrKdInstalasi = "11") Or (mstrKdInstalasi = "06") Then
        adocomd.CommandText = "SELECT * from V_CetakLabelRegistrasiPasienMRSNew " _
        & " WHERE (NoPendaftaran) =('" & mstrNoPen & "')"
    Else
        adocomd.CommandText = "SELECT * from V_CetakLabelRegistrasiPasienRINew " _
        & " WHERE NoPendaftaran ='" & mstrNoPen & "'"
    End If
    adocomd.CommandType = adCmdText

    Report2.Database.AddADOCommand dbConn, adocomd
    
    Dim tanggal As String
    tanggal = Format(TglPeriodeAwal, "MMMM yyyy")

    With Report2
        .text1.SetText strNNamaRS
        .text2.SetText strNAlamatRS & ", " & strNKotaRS
        .text3.SetText strEmail
        .udtgl.SetUnboundFieldSource ("{ado.tglmasuk}")
        .usnodft.SetUnboundFieldSource ("{ado.nopendaftaran}")
        .usnocm.SetUnboundFieldSource ("{Ado.nocm}")
        .usnocmerm.SetUnboundFieldSource ("{Ado.nocmterm}")
        .usruangperiksa.SetUnboundFieldSource ("{Ado.Ruang Tujuan}")

        .usnmpasien.SetUnboundFieldSource ("{Ado.nama pasien}")
        .usJK.SetUnboundFieldSource ("{ado.jk}")
        .usUmur.SetUnboundFieldSource ("{ado.Umur}")
        .usAlamat.SetUnboundFieldSource ("{Ado.Alamat}")
        .unRT.SetUnboundFieldSource ("{Ado.RTRW}")
        .usKelurahan.SetUnboundFieldSource ("{Ado.Kelurahan}")
        .usKecamatan.SetUnboundFieldSource ("{Ado.Kecamatan}")

        .usPenjamin.SetUnboundFieldSource ("{Ado.JenisPasien}")
        .USJenisKelas.SetUnboundFieldSource ("{ado.DetailJenisJasaPelayanan}")
        .usNoBed.SetUnboundFieldSource ("{Ado.NoBed}")

        If mstrKdInstalasi = "03" Then
            .txtNoRuangan.SetText "No. Kamar"
            .usNoRuangan.SetUnboundFieldSource ("{Ado.NamaKamar}")
            .txtNoBed.Suppress = False
            .usNoBed.Suppress = False
        Else
            .txtNoRuangan.SetText "No. Ruangan"
            .usNoRuangan.SetUnboundFieldSource ("{Ado.No. Ruangan}")
            .txtNoBed.Suppress = True
            .usNoBed.Suppress = True
        End If

        .usnoantri.SetUnboundFieldSource ("{Ado.NoAntrian}")
        .ususer.SetText (strNmPegawai)
'        .txtMessage.SetText (rs(0).value)
    
        NamaCR = "Cr_cetakLabel"
        If sp_CetakLaporan(NamaCR, mstrNoPen, mstrKdRuangan, strIDPegawai) = False Then Exit Sub
        .txtKendali.SetText strCetakKendaliLaporan & " Cetakan ke: " & intCetakKe
            
    Dim tempPrint2 As String
    Dim strDeviceName2 As String
    Dim strDriverName2 As String
    Dim strPort2 As String
    
    tempPrint2 = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "PrinterLabel2 ")
    
    For Each p In Printers
            strDeviceName2 = tempPrint2
            strDriverName2 = p.DriverName
            strPort2 = p.Port
            
            Report2.SelectPrinter strDriverName2, strDeviceName2, strPort2
            Report2.PrintOut False
        Exit For

    Next
    Unload Me
    Screen.MousePointer = vbDefault
End With

Exit Sub
End Sub
