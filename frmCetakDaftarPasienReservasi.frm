VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakDaftarPasienReservasi 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "frmCetakDaftarPasienReservasi.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
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
      EnableAnimationControl=   -1  'True
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
Attribute VB_Name = "frmCetakDaftarPasienReservasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crCetakDaftarPasienReservasi

Private Sub Form_Load()
On Error GoTo errLoad
Me.WindowState = 2
    Screen.MousePointer = vbHourglass
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn
    
 Me.Caption = "Medifirst2000 - Cetak Daftar Pasien Askes"
 Set Report = New crCetakDaftarPasienReservasi
    
    
    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText
    With Report
        .Database.AddADOCommand dbConn, dbcmd
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        
        .usTglPesan.SetUnboundFieldSource ("{ado.Tgl Pesan}")
        .usTglPeriksa.SetUnboundFieldSource ("{ado.TglMasuk}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaLengkap}")
        .usNoAntrian.SetUnboundFieldSource ("{ado.NoAntrian}")
        .usRuangan.SetUnboundFieldSource ("{ado.NamaRuangan}")
        .usDokter.SetUnboundFieldSource ("{ado.NamaDokter}")
        .usPoli.SetUnboundFieldSource ("{ado.RuanganPoli}")
        .usNoTlp.SetUnboundFieldSource ("{ado.NoTlp}")
        .usKeterangan.SetUnboundFieldSource ("{ado.Keterangan}")
        
        
        .txtDptAwal.SetText frmDaftarReservasiPasien.dtpAwal.value
        .txtDptAkhir.SetText frmDaftarReservasiPasien.dtpAkhir.value
        
        If frmDaftarReservasiPasien.chkStatus = vbUnchecked Then
            .txtStatus.SetText "Belum Di Registrasikan"
        Else
            .txtStatus.SetText "Sudah Di Registrasikan"
        End If
    End With


   
        
    
CRViewer1.ReportSource = Report
With CRViewer1
               .ReportSource = Report
               .ViewReport
               .Zoom 1
    
    End With
    Screen.MousePointer = vbDefault
Exit Sub
errLoad:
    Screen.MousePointer = vbDefault
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub



