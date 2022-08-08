VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakAntrianReservasi 
   Caption         =   "frmCetakAntrianReservasi"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
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
Attribute VB_Name = "frmCetakAntrianReservasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crCetakAntrianReservasi
Dim DB As CRekamMedis

Private Sub Form_Load()
On Error GoTo errLoad
    Dim adocomd As New ADODB.Command
    Set DB = New CRekamMedis
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    
    Call openConnection
    With frmDaftarReservasiPasien
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = "select NoReservasi,NoCM,NamaLengkap,NoAntrian,TglMasuk,RuanganPoli,Keterangan,NamaRuangan,NamaOperator,NamaDokter " & _
                          "from V_DaftarReservasiPasien_New where NoAntrian = '" & .dgDaftarReservasiPasien.Columns(0).value & "' and " & _
                          "NamaLengkap = '" & .dgDaftarReservasiPasien.Columns(4).value & "' " & _
                          "And NamaDokter = '" & .dgDaftarReservasiPasien.Columns(5).value & "' " & _
                          "And NamaRuangan = '" & .dgDaftarReservasiPasien.Columns(7).value & "' " & _
                          "And TglMasuk between '" & Format(.dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(.dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "' order by NoAntrian"
    adocomd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, adocomd
    End With
    
    Dim tanggal As String
    tanggal = Format(TglPeriodeAwal, "MMMM yyyy") '& " S/d " & Format(frmregister.DTPickerAkhir, "dd MMMM yyyy")
    
    With Report
        .usRuang.SetUnboundFieldSource ("{ado.NamaRuangan}")
        .usNama.SetUnboundFieldSource ("{ado.NamaLengkap}")
        .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
        .unNomerAntrian.SetUnboundFieldSource ("{ado.NoAntrian}")
        .usUser.SetUnboundFieldSource ("{ado.NamaOperator}")
        .usDokter.SetUnboundFieldSource ("{ado.NamaDokter}")
        
'        .SelectPrinter "winspool", DB.DefaultPrinterLabelRJ, vbNull
        .PrintOut False
    End With
    
Screen.MousePointer = vbHourglass
Screen.MousePointer = vbDefault
Unload Me
Call sp_Reservasi_Temp(dbcmd)
Exit Sub
errLoad:
    Call msubPesanError
End Sub
Private Sub sp_Reservasi_Temp(ByVal dbcmd As ADODB.Command)
Dim KdJenisPoliklinik As String
Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoReservasi", adInteger, adParamInput, , Null)
        .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 6, frmDaftarReservasiPasien.dgDaftarReservasiPasien.Columns(3).value)
        .Parameters.Append .CreateParameter("NamaLengkap", adVarChar, adParamInput, 50, frmDaftarReservasiPasien.dgDaftarReservasiPasien.Columns(4).value)
        .Parameters.Append .CreateParameter("NoAntrian", adChar, adParamInput, 6, frmDaftarReservasiPasien.dgDaftarReservasiPasien.Columns(0).value)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(frmDaftarReservasiPasien.dgDaftarReservasiPasien.Columns(2).value, "yyyy/MM/dd HH:mm:ss"))
        strSQL = "Select * from JenisPoliklinik where JenisPoliklinik='" & Trim(frmDaftarReservasiPasien.dgDaftarReservasiPasien.Columns(8).value) & "'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
        KdJenisPoliklinik = rs(0).value
        
        End If
        .Parameters.Append .CreateParameter("KdJenisPoliklinik", adVarChar, adParamInput, 2, KdJenisPoliklinik)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 50, Null)
        .Parameters.Append .CreateParameter("NamaRuangan", adVarChar, adParamInput, 50, frmDaftarReservasiPasien.dgDaftarReservasiPasien.Columns(7).value)
        .Parameters.Append .CreateParameter("NamaOperator", adVarChar, adParamInput, 50, strIDPegawai)
        .Parameters.Append .CreateParameter("NamaDokter", adVarChar, adParamInput, 50, frmDaftarReservasiPasien.dgDaftarReservasiPasien.Columns(5).value)
        

        .ActiveConnection = dbConn
        .CommandText = "Add_Reservasi_Temp"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute
        
    If .Parameters("return_value").value <> 0 Then
        MsgBox "Error - Ada Kesalahan Dalam Penyimpanan Data, Hubungi Administrator", vbCritical, "Error"
    Else
    End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
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
Set frmCetakAntrianReservasi = Nothing

End Sub


