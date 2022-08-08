VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakNilaiPersediaanNM 
   Caption         =   "Form Cetak Nilai Persediaan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakNilaiPersediaanNM.frx":0000
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
Attribute VB_Name = "frmCetakNilaiPersediaanNM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report1 As New crCetakNilaiPersediaanNM

Private Sub Form_Load()
    On Error GoTo Errload
    Dim adocomd As New ADODB.Command

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    If mstrKdKelompokBarang = "02" Then     'medis
    ElseIf mstrKdKelompokBarang = "01" Then 'non medis
        With Report1
            .Text1.SetText strNNamaRS
            .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
            .Text3.SetText strWebsite & ", " & strEmail
            .txtPeriodeClosing.SetText "Periode Closing " & mdtglclosing
            .txtNamaRuangan.SetText mstrNamaRuangan
        End With
    End If

    Set dbcmd = New ADODB.Command

    If frmNilaiPersediaanNM.ChkPerjenis.value = 0 Then
        strSQL = "SELECT * FROM V_DataStokBarangNonMedisRekapx " & _
        " WHERE KdRuangan = '" & mstrKdRuangan & "' AND (TglClosing = '" & Format(frmNilaiPersediaanNM.dcNoClosing.BoundText, "yyyy/MM/dd hh:mm:ss") & "') AND StokReal<> 0" & _
        " ORDER By JenisBarang, NamaBarang"
    Else
        strSQL = "SELECT * FROM V_DataStokBarangNonMedisRekapx " & _
        " WHERE JenisBarang like '" & frmNilaiPersediaanNM.dcJenisBarang.Text & "%' and KdRuangan = '" & mstrKdRuangan & "' AND (TglClosing = '" & Format(frmNilaiPersediaanNM.dcNoClosing.BoundText, "yyyy/MM/dd HH:mm:ss") & "')" & _
        " ORDER By JenisBarang, NamaBarang"
    End If

    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText

    If mstrKdKelompokBarang = "02" Then     'medis
    ElseIf mstrKdKelompokBarang = "01" Then     'non medis
        Report1.Database.AddADOCommand dbConn, dbcmd
        With Report1
            .usNamaBarang.SetUnboundFieldSource ("{ado.NamaBarang}")
            .usJenisBarang.SetUnboundFieldSource ("{ado.JenisBarang}")
            .ucHargaNetto.SetUnboundFieldSource ("{ado.HargaNetto}")
            .usMerk.SetUnboundFieldSource ("{ado.Merk}")
            .usType.SetUnboundFieldSource ("{ado.Type}")
            .usBahan.SetUnboundFieldSource ("{ado.Bahan}")
            .unStok.SetUnboundFieldSource ("{ado.StokReal}")
            .uctotal.SetUnboundFieldSource ("{ado.TotalNetto}")
            .UsNoRegisterAsset.SetUnboundFieldSource ("{ado.NoRegisterAsset}")
        End With

        With CRViewer1
            .EnableGroupTree = True
            .ReportSource = Report1
            .ViewReport
            .Zoom 1
        End With
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakNilaiPersediaanNM = Nothing
End Sub

