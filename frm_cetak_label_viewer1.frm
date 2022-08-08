VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_cetak_label_viewer1 
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
Attribute VB_Name = "frm_cetak_label_viewer1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New Cr_cetakLabel
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
        adocomd.CommandText = "SELECT * from V_CetakLabelRegistrasiPasienMRS " _
        & " WHERE (NoPendaftaran) =('" & mstrNoPen & "')"
    Else
        adocomd.CommandText = "SELECT * from V_CetakLabelRegistrasiPasienRI " _
        & " WHERE NoPendaftaran ='" & mstrNoPen & "'"
    End If
    adocomd.CommandType = adCmdText

    Report.Database.AddADOCommand dbConn, adocomd

    Dim tanggal As String
    tanggal = Format(TglPeriodeAwal, "MMMM yyyy")

    With Report
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS
        .Text3.SetText strEmail
        .udtgl.SetUnboundFieldSource ("{ado.tglmasuk}")
        .usnodft.SetUnboundFieldSource ("{ado.nopendaftaran}")
        .usNoCM.SetUnboundFieldSource ("{Ado.nocm}")
        .usruangperiksa.SetUnboundFieldSource ("{Ado.Ruang Tujuan}")

        .usNmPasien.SetUnboundFieldSource ("{Ado.nama pasien}")
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
        .usUser.SetText (strNmPegawai)
        .txtMessage.SetText (rs(0).value)
    
        NamaCR = "Cr_cetakLabel"
        If sp_CetakLaporan(NamaCR, mstrNoPen, mstrKdRuangan, strIDPegawai) = False Then Exit Sub
        .txtKendali.SetText strCetakKendaliLaporan & " Cetakan ke: " & intCetakKe
        
' untuk cetak 2 printer sekaligus
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
