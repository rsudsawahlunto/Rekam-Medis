VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakDaftarPasienS 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   10365
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20205
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
Attribute VB_Name = "frmCetakDaftarPasienS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New cr_DaftarKunjPasienMskBWilayahBJekel
Dim Report2 As New cr_DaftarKunjPasienMskBJenisBWilayah

Private Sub Form_Load()
    Set frmCetakDaftarPasienS = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    If strCetak = "WilayahJekel" Then
        With Report
            .Database.AddADOCommand dbConn, adocomd
        
            If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
'                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
            Else
'                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
            End If

            .txtNamaRS.SetText strNNamaRS
            .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
            .txtAlamat2.SetText strEmail

            .usBulan.SetUnboundFieldSource ("{ado.TglMasuk}")
            .UsRuangan.SetUnboundFieldSource ("{ado.RuanganPerawatan}")
            .usWilayah.SetUnboundFieldSource ("{ado.Judul}")
            .usKelas.SetUnboundFieldSource ("{ado.Kelas}")
            .usJekel.SetUnboundFieldSource ("{ado.JK}")
            .usJml.SetUnboundFieldSource ("{ado.Jml}")
            .txtjudul.SetText ""
            If sUkuranKertas = "" Then
                sUkuranKertas = "5"
                sOrientasKertas = "2"
                sDuplex = "0"
            End If
            settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
        End With
        Me.WindowState = 2
        Screen.MousePointer = vbHourglass
            With CRViewer1
                .ReportSource = Report
                .ViewReport
                .Zoom (100)
            End With
        Screen.MousePointer = vbDefault
    
    Else
        With Report2
            .Database.AddADOCommand dbConn, adocomd

            If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
'                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
            Else
'                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
            End If
            .txtNamaRS.SetText strNNamaRS
            .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
            .txtAlamat2.SetText strEmail

            .usBulan.SetUnboundFieldSource ("{ado.Periode}")
            .usWilayah.SetUnboundFieldSource ("{ado.Judul}")
            .usKelas.SetUnboundFieldSource ("{ado.Kelas}")
            .usPenjamin.SetUnboundFieldSource ("{ado.NamaPenjamin}")
            .usJml.SetUnboundFieldSource ("{ado.Jml}")
            .txtjudul.SetText ""
            If sUkuranKertas = "" Then
                sUkuranKertas = "5"
                sOrientasKertas = "2"
                sDuplex = "0"
            End If
        End With
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = Report2
            .ViewReport
            .Zoom (100)
        End With
        Screen.MousePointer = vbDefault
    End If
End Sub
