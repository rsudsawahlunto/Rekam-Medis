VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmTenagaGizi 
   Caption         =   "Medifirst2000 - Data Ketenagaan"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   Icon            =   "frmTenagaGizi.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   8325
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
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
Attribute VB_Name = "frmTenagaGizi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crtenagagizi
Dim adoCommand As New ADODB.Command

Private Sub Form_Load()
    On Error GoTo hell
    Set dbcmd = New ADODB.Command

    strSQL = "SELECT  kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak," & _
    "(sum(jmldpkfull)+(jmldpbfull)+(jmldaerahfull)+(jmlpnkfull)+(jmlabrifull)+(jmldeplainfull)+(jmlpttfull)+(jmlswastafull)+(jmlkontrak)) as subtotal1, " & _
    "jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart," & _
    "(sum(jmldpkpart)+(jmldpbpart)+(jmldaerahpart)+(jmlpnkpart)+(jmlabripart)+(jmldeplainpart)+(jmlpttpart)+(jmlswastapart))as subtotal2," & _
    "jmlhonorer " & _
    "FROM V_Ketenagaan WHERE  (KdkualifikasiJurusan IN ('0024', '0025','0026','0027','0029','0030','0181')) " & _
    "GROUP BY  kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak,jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart,jmlhonorer"
    Call msubRecFO(rs, strSQL)

    openConnection

    Set frmTenagaGizi = Nothing

    adoCommand.CommandText = strSQL
    adoCommand.CommandType = adCmdText

    With Report
        .Database.AddADOCommand dbConn, adoCommand

        .Text1.SetText strNNamaRS
        .Text2.SetText "A.JUMLAH TENAGA KESEHATAN MENURUT JENIS"
        .Text3.SetText "5. TENAGA GIZI"

        .unkualifikasi.SetUnboundFieldSource ("{ado.Kualifikasijurusan}")
        .undepkes.SetUnboundFieldSource ("{ado.jmldpkfull}")
        .unprop.SetUnboundFieldSource ("{ado.jmldpbfull}")
        .unkota.SetUnboundFieldSource ("{ado.jmldaerahfull}")
        .undepdiknas.SetUnboundFieldSource ("{ado.jmlpnkfull}")
        .untni.SetUnboundFieldSource ("{ado.jmlabrifull}")
        .unbumn.SetUnboundFieldSource ("{ado.jmldeplainfull}")
        .unptt.SetUnboundFieldSource ("{ado.jmlpttfull}")
        .unswasta.SetUnboundFieldSource ("{ado.jmlswastafull}")
        .unkontrak.SetUnboundFieldSource ("{ado.jmlkontrak}")
        .unsubtotal.SetUnboundFieldSource ("{ado.subtotal1}")

        .undepkes2.SetUnboundFieldSource ("{ado.jmldpkpart}")
        .unprop2.SetUnboundFieldSource ("{ado.jmldpbpart}")
        .unkota2.SetUnboundFieldSource ("{ado.jmldaerahpart}")
        .undepdiknas2.SetUnboundFieldSource ("{ado.jmlpnkpart}")
        .untni2.SetUnboundFieldSource ("{ado.jmlabripart}")
        .unbumn2.SetUnboundFieldSource ("{ado.jmldeplainpart}")
        .unptt2.SetUnboundFieldSource ("{ado.jmlpttpart}")
        .unswasta2.SetUnboundFieldSource ("{ado.jmlswastapart}")
        .unsubtotal2.SetUnboundFieldSource ("{ado.subtotal2}")

        .unhonorer.SetUnboundFieldSource ("{ado.jmlhonorer}")
        .untotal.SetUnboundFieldSource Format("{ado.subtotal1}+ {ado.subtotal2}+ {ado.jmlhonorer}")
    End With

    strSQL = "select * " & _
    " from V_Koders  "

    Call msubRecFO(rs, strSQL)

    With Report
        .Text6.SetText rs("NO1")
        .Text7.SetText rs("NO2")
        .Text16.SetText rs("NO3")
        .Text17.SetText rs("NO4")
        .Text21.SetText rs("NO5")
        .Text22.SetText rs("NO6")
        .Text35.SetText rs("NO7")
    End With

    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If

    Screen.MousePointer = vbDefault
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
    Set frmTenagaGizi = Nothing
End Sub

