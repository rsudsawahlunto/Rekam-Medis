VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmViewerLaporanforBukuRegisterPelayanan 
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   Icon            =   "FrmViewerLaporanforBukuRegisterPelayanan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   5850
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
      EnablePrintButton=   0   'False
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
Attribute VB_Name = "FrmViewerLaporanforBukuRegisterPelayanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim reportBukuRegisterPasien As New crBukuRegisterPelayananPasien
Dim adocomd As New ADODB.Command
Dim tanggal As String
Dim p As Printer
Dim tempPrint1 As String
Dim strDeviceName As String
Dim strDriverName As String
Dim strPort As String
Dim Posisi, z, Urutan As Integer
Public strFilter As String
Public sPrinterLegal As String

Private Sub Form_Load()
    Set adocomd = New ADODB.Command
    Set adocomd = Nothing
    adocomd.ActiveConnection = dbConn

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Dim tanggal As String

      Call BkRegister

End Sub

Private Sub BkRegister()

    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText
    
    With reportBukuRegisterPasien
        .Text16.SetText strNNamaRS
        .Text18.SetText strNAlamatRS
        .TxtJudul.SetText "BUKU REGISTER PELAYANAN"
        .Text19.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
        .Database.AddADOCommand dbConn, adocomd
        .txtTgl.SetText Format(FrmBukuRegisterPelayanan.DTPickerAwal, "dd/MM/yyyy") & "  s/d  " & Format(FrmBukuRegisterPelayanan.DTPickerAkhir, "dd/MM/yyyy")
      
        .usNoPendaftaran.SetUnboundFieldSource "{ado.NoPendaftaran}"
        .usNoCM.SetUnboundFieldSource "{ado.NoCM}"
        .usPasien.SetUnboundFieldSource "{ado.NamaPasien}"
        .usJK.SetUnboundFieldSource "{ado.JK}"
        .usJenisPasien.SetUnboundFieldSource "{ado.JenisPasien}"
        .usKelas.SetUnboundFieldSource "{ado.Kelas}"
                
        .udtTglPelayanan.SetUnboundFieldSource "{ado.TglPelayanan}"
        .usJenisPelayanan.SetUnboundFieldSource "{ado.JenisPelayanan}"
        .usRK.SetUnboundFieldSource "{ado.R/K}"
        .usNamaPelayanan.SetUnboundFieldSource "{ado.NamaPelayanan}"
        .usAsalPelayanan.SetUnboundFieldSource "{ado.AsalPelayanan}"
        .usQty.SetUnboundFieldSource "{ado.qty}"
'

        .ucHarga.SetUnboundFieldSource "{ado.HargaSatuan}"
        .ucHargaCito.SetUnboundFieldSource "{ado.HargaCito}"
        .ucHargaService.SetUnboundFieldSource "{ado.HargaService}"
        .ucTotalBiaya.SetUnboundFieldSource "{ado.TotalBiaya}"
        .ucHutangPenjamin.SetUnboundFieldSource "{ado.JmlHutangPenjamin}"
        .ucTanggunganRS.SetUnboundFieldSource "{ado.JmlTanggunganRS}"

        .ucDiskon.SetUnboundFieldSource "{ado.JmlDiskon}"
        .ucHarusDibayar.SetUnboundFieldSource "{ado.TotalHarusDibayar}"
        .usDokterOperator.SetUnboundFieldSource "{ado.DokterOperator}"
        .usDokterAnastesi.SetUnboundFieldSource "{ado.DokterAnastesi}"
        .usDokterPendamping.SetUnboundFieldSource "{ado.DokterPendamping}"
        .usRuanganPelayanan.SetUnboundFieldSource "{ado.Ruangan}"
        .udtTglBkm.SetUnboundFieldSource "{ado.TglBkm}"
'
    End With
    
    sPrinterLegal = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "PrinterLegal")
    
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
        .ReportSource = reportBukuRegisterPasien
        .ViewReport
        .Zoom (100)

         Screen.MousePointer = vbDefault
        End With
    Else
        Urutan = 0
        For z = 1 To Len(sPrinterLegal)
            If Mid(sPrinterLegal, z, 1) = ";" Then
                Urutan = Urutan + 1
                Posisi = z
                ReDim Preserve arrPrinter(Urutan)
                arrPrinter(Urutan).intUrutan = Urutan
                arrPrinter(Urutan).intPosisi = Posisi
                If Urutan = 1 Then
                    arrPrinter(Urutan).strNamaPrinter = Mid(sPrinterLegal, 1, z - 1)
                Else
                    arrPrinter(Urutan).strNamaPrinter = Mid(sPrinterLegal, arrPrinter(Urutan - 1).intPosisi + 1, z - arrPrinter(Urutan - 1).intPosisi - 1)
                End If
             
             
            For Each p In Printers
                    strDeviceName = arrPrinter(Urutan).strNamaPrinter
                    strDriverName = p.DriverName
                    strPort = p.Port
        
                    reportBukuRegisterPasien.SelectPrinter strDriverName, strDeviceName, strPort
                    reportBukuRegisterPasien.PrintOut False
                    Screen.MousePointer = vbDefault

            Exit For
            
            Next
        End If
    Next z
      Unload Me
    End If
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strStatus = ""
    strFilter = ""
    Set FrmViewerLaporanforBukuRegisterPelayanan = Nothing
End Sub

