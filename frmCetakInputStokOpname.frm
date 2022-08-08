VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakInputStokOpname 
   Caption         =   "Form Cetak Stok Opname"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakInputStokOpname.frx":0000
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
Attribute VB_Name = "frmCetakInputStokOpname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crCetakInputStokOpname
Dim Report1 As New crCetakInputStokOpnameNM
Public chkHarga As Integer

Private Sub Form_Load()
    On Error GoTo errLoad
    Dim adocomd As New ADODB.Command

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    If mstrKdKelompokBarang = "02" Then     'medis
        With Report
            .Text1.SetText strNNamaRS
            .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
            .Text3.SetText strWebsite & ", " & strEmail

            .txtNamaRuangan.SetText mstrNamaRuangan
        End With
    ElseIf mstrKdKelompokBarang = "01" Then 'non medis
        With Report1
            .Text1.SetText strNNamaRS
            .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
            .Text3.SetText strWebsite & ", " & strEmail

            .txtNamaRuangan.SetText mstrNamaRuangan
        End With
    End If

    Set dbcmd = New ADODB.Command
    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText

    If mstrKdKelompokBarang = "02" Then     'medis
        Report.Database.AddADOCommand dbConn, dbcmd
        With Report
            .usJenisBarang.SetUnboundFieldSource ("{ado.JenisBarang}")
            .usNamaBarang.SetUnboundFieldSource ("{ado.NamaBarang}")
            .unStokSistem.SetUnboundFieldSource ("{ado.StokSystem}")
            If chkHarga = 1 Then
                .Text44.Suppress = False
                .ucHargaNetto1.Suppress = False
                .ucHargaNetto1.SetUnboundFieldSource ("{ado.HargaNetto1}")
            Else
                .Text44.Suppress = True
                .ucHargaNetto1.Suppress = True
            End If
        End With

        If vLaporan = "view" Then
            With CRViewer1
                .ReportSource = Report
                .EnableGroupTree = True
                .ViewReport
                .Zoom 1
            End With
        Else
            Report.PrintOut False
            Unload Me
        End If
    ElseIf mstrKdKelompokBarang = "01" Then 'non medis
        Report1.Database.AddADOCommand dbConn, dbcmd
        With Report1
             .usJenisBarang.SetUnboundFieldSource ("{ado.JenisBarang}")
            .usNamaBarang.SetUnboundFieldSource ("{ado.NamaBarang}")
            .usMerk.SetUnboundFieldSource ("{ado.Merk}")
            .usType.SetUnboundFieldSource ("{ado.Type}")
            .usBahan.SetUnboundFieldSource ("{ado.Bahan}")
            .ucHargaNetto.SetUnboundFieldSource ("{ado.HargaNetto}")
            .ucHargaNetto2.SetUnboundFieldSource ("{ado.HargaNetto2}")
            .usNoRegisterAsset.SetUnboundFieldSource ("{ado.NoRegisterAsset}")
            .usJumlahReal.SetUnboundFieldSource ("{ado.StokSystem}")
            .unStok.SetUnboundFieldSource ("{ado.StokIsi}")
        End With

        If vLaporan = "view" Then
            With CRViewer1
                .ReportSource = Report1
                .EnableGroupTree = True
                .ViewReport
                .Zoom 1
            End With
        Else
            Report1.PrintOut False
            Unload Me
        End If
    End If

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

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakInputStokOpname = Nothing
End Sub
