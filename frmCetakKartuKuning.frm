VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakKartuKuning 
   Caption         =   "frmCetakKartu"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13080
   Icon            =   "frmCetakKartuKuning.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   13080
   WindowState     =   2  'Maximized
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
Attribute VB_Name = "frmCetakKartuKuning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crCetakKartuPasienKuning
Dim p As Printer

Private Sub Form_Load()
    On Error GoTo hell
    Call openConnection

    strSQL = "SELECT NoCM, NamaLengkap, Alamat, JenisKelamin, Propinsi, Kota, Kecamatan, Kelurahan, TglDaftarMembership, Umur,RTRW" & _
    " FROM   V_CetakKartuKuningPasien" & _
    " WHERE NoCM = '" & mstrNoCM & "'"

    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    
    If rs.EOF = False Then
        With Report
            .txtNoCM.SetText rs("NoCM")
            .txtNamaPasien.SetText rs("NamaLengkap")
            .txtTgl.SetText rs("TglDaftarMembership")
    
            .txtumur.SetText IIf(IsNull(rs("Umur")), "-", rs("Umur"))
            If rs("JenisKelamin").value = "L" Then
                .txtJK.SetText "Laki-Laki"
            Else
                .txtJK.SetText "Perempuan"
            End If
    
            .txtAlamat.SetText IIf(IsNull(rs("Alamat")), "-", rs("Alamat"))
            .txtKota.SetText IIf(IsNull(rs("Kota")), "-", rs("Kota"))
            .txtKecamatan.SetText IIf(IsNull(rs("Kecamatan")), "-", rs("Kecamatan"))
            .txtLingkungan.SetText IIf(IsNull(rs("Kelurahan")), "-", rs("Kelurahan"))
            .txtRTRW.SetText IIf(IsNull(rs("RTRW")), "-", rs("RTRW"))
        End With
        Screen.MousePointer = vbHourglass
        
' untuk cetak DARI SETTINGAN PRINTER

Dim tempPrint1 As String
Dim strDeviceName As String
Dim strDriverName As String
Dim strPort As String
Dim Posisi, z, Urutan As Integer

Dim sPrinter1 As String
    
    sPrinter1 = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "Printer1")
    
        Urutan = 0
        For z = 1 To Len(sPrinter1)
            If Mid(sPrinter1, z, 1) = ";" Then
                Urutan = Urutan + 1
                Posisi = z
                ReDim Preserve arrPrinter(Urutan)
                arrPrinter(Urutan).intUrutan = Urutan
                arrPrinter(Urutan).intPosisi = Posisi
                If Urutan = 1 Then
                    arrPrinter(Urutan).strNamaPrinter = Mid(sPrinter1, 1, z - 1)
                Else
                    arrPrinter(Urutan).strNamaPrinter = Mid(sPrinter1, arrPrinter(Urutan - 1).intPosisi + 1, z - arrPrinter(Urutan - 1).intPosisi - 1)
                End If
             
             
            For Each p In Printers
                    strDeviceName = arrPrinter(Urutan).strNamaPrinter
                    strDriverName = p.DriverName
                    strPort = p.Port
        
                    Report.SelectPrinter strDriverName, strDeviceName, strPort
                    Report.PrintOut False
                    Screen.MousePointer = vbDefault

            Exit For
            
            Next
        End If
    Next z
      Unload Me
   
    
'    tempPrint1 = ReadINI("Default Printer", "PrinterKartu", "", "C:\SettingPrinter.ini")
'            With CRViewer1
'
'                For Each p In Printers
'                        strDeviceName = tempPrint1
'                        strDriverName = p.DriverName
'                        strPort = p.Port
'
'                        Report.SelectPrinter strDriverName, strDeviceName, strPort
'            '            Report.PrintOut False
'
'                        .ReportSource = Report
'                        .ViewReport
'                        .EnableGroupTree = False
'                        .Zoom 1
'
'                         Screen.MousePointer = vbDefault
'                Exit For
'
'                Next
'            End With
'
'        With CRViewer1
'            .ReportSource = Report
'            .ViewReport
'            .EnableGroupTree = False
'            .Zoom 1
'        End With
'        Screen.MousePointer = vbDefault
    Else
        MsgBox "Data pasien harus dilengkapi", vbInformation
    End If
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakKartuKuning = Nothing
End Sub

