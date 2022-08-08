VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakLapSaldoBarangNM 
   Caption         =   "Medifrst2000 - Laporan Saldo Barang"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   Icon            =   "frmCetakLapSaldoBarangNM.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4635
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7005
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   5805
      DisplayGroupTree=   0   'False
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
      EnableAnimationControl=   0   'False
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
Attribute VB_Name = "frmCetakLapSaldoBarangNM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crLapSaldoBarangNM
Dim Judul1, Judul2, Judul3 As String
Dim strGroup As String
Dim strNama As String
Dim strIsiGroup As String

Private Sub Form_Load()
On Error GoTo errLoad
Dim adocomd As New ADODB.Command

    Judul1 = "LAPORAN SALDO BARANG NON MEDIS (PER HARI)"
    Judul2 = "LAPORAN SALDO BARANG NON MEDIS (PERBULAN)"
    Judul3 = "LAPORAN SALDO BARANG NON MEDIS (PERTAHUN)"

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Call openConnection
    
    Select Case strCetak
        Case "Hari"
            Call LaporanPerHari
        Case "Bulan"
            Call LaporanPerBulan
        Case "Tahun"
            Call LaporanPerTahun
        Case "Total"
            Call LaporanTotal
        Case Else
            MsgBox "Pilih dulu mau per Hari, per Bulan, atau per Tahun.", vbExclamation, "Validasi"
            Exit Sub
    End Select
    
    With Report
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail

        .txtGroupBy.SetText strGroup
        .txtTanggal.SetText strCetak
        .txtPeriode.SetText CStr(Format(mdTglAwal, "dd MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "dd MMMM yyyy"))
        
        .udTanggal.SetUnboundFieldSource Format(("{ado.TglTransaksi}"), "dd MMMM yyyy")
        If strGroup <> "" Then .usGroupBy.SetUnboundFieldSource ("{ado." & strGroup & "}")
        .usNamaBarang.SetUnboundFieldSource ("{ado.NamaBarang}")
        
        .txtRuanganLogin.SetText mstrNamaRuangan
        .txtUser.SetText strNmPegawai
    End With
    
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (100)
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

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakLapSaldoBarangNM = Nothing
End Sub

Private Sub LaporanPerHari()
Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    
    If strGroup = "" Then
        adocomd.CommandText = "SELECT * FROM V_DataTransaksiBarangNM " & _
            " WHERE TglTransaksi BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
            " AND NamaBarang LIKE '%" & strNama & "%' AND KdRuangan = '" & mstrKdRuangan & "'"
    Else
        adocomd.CommandText = "SELECT * FROM V_DataTransaksiBarangNM " & _
            " WHERE TglTransaksi BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
            " AND " & strGroup & " LIKE '" & strIsiGroup & "%' AND NamaBarang LIKE '%" & strNama & "%' AND KdRuangan = '" & mstrKdRuangan & "'"
    End If
    
    adocomd.CommandType = adCmdText
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .txtJudul.SetText Judul1
        .unAwal.SetUnboundFieldSource ("{ado.JmlStokAwalPB}")
        .unMasuk.SetUnboundFieldSource ("{ado.JmlTerima}")
        .unKeluar.SetUnboundFieldSource ("{ado.JmlKeluar}")
        .unSaldo.SetUnboundFieldSource Format(("{ado.JmlStokAwalPB} + {ado.JmlTerima} - {ado.JmlKeluar}"), "###,##0")
        .ucNetto.SetUnboundFieldSource Format(("{ado.HargaNetto}"), "###,##0")
        .ucTotal.SetUnboundFieldSource Format(("{ado.JmlStokAwalPB} + {ado.JmlTerima} - {ado.JmlKeluar} * {ado.HargaNetto}"), "###,##0")
    End With
End Sub

Private Sub LaporanPerBulan()
Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    
    If strGroup = "" Then
        adocomd.CommandText = "SELECT {fn MONTHNAME (TglTransaksi)} AS TglTransaksi,  AsalBarang, DetailJenisBarang,  NamaBarang, JmlStokAwalPB, JmlTerima, JmlKeluar, HargaNetto FROM V_DataTransaksiBarangNM " & _
            " WHERE TglTransaksi BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
            " AND NamaBarang LIKE '%" & strNama & "%' AND KdRuangan = '" & mstrKdRuangan & "'"
    Else
        adocomd.CommandText = "SELECT {fn MONTHNAME (TglTransaksi)} AS TglTransaksi,  AsalBarang, DetailJenisBarang,  NamaBarang, JmlStokAwalPB, JmlTerima, JmlKeluar, HargaNetto FROM V_DataTransaksiBarangNM " & _
            " WHERE TglTransaksi BETWEEN '" & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
            " AND " & strGroup & " LIKE '" & strIsiGroup & "%' AND NamaBarang LIKE '%" & strNama & "%' AND KdRuangan = '" & mstrKdRuangan & "'"
    End If
    
    adocomd.CommandType = adCmdText
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .txtJudul.SetText Judul2
        .unAwal.SetUnboundFieldSource ("{ado.JmlStokAwalPB}")
        .unMasuk.SetUnboundFieldSource ("{ado.JmlTerima}")
        .unKeluar.SetUnboundFieldSource ("{ado.JmlKeluar}")
        .unSaldo.SetUnboundFieldSource Format(("{ado.JmlStokAwalPB} + {ado.JmlTerima} - {ado.JmlKeluar}"), "###,##0")
        .ucNetto.SetUnboundFieldSource Format(("{ado.HargaNetto}"), "###,##0")
        .ucTotal.SetUnboundFieldSource Format(("{ado.JmlStokAwalPB} + {ado.JmlTerima} - {ado.JmlKeluar} * {ado.HargaNetto}"), "###,##0")
    End With
End Sub

Private Sub LaporanPerTahun()
Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    
    If strGroup = "" Then
        adocomd.CommandText = "SELECT *, {fn YEAR (TglTransaksi) } AS TglTransaksi,  AsalBarang, DetailJenisBarang,  NamaBarang, JmlStokAwalPB, JmlTerima, JmlKeluar, HargaNetto FROM V_DataTransaksiBarangNM " & _
            " WHERE TglTransaksi BETWEEN '" & Format(mdTglAwal, "yyyy/01/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
            " AND NamaBarang LIKE '%" & strNama & "%' AND KdRuangan = '" & mstrKdRuangan & "'"
    Else
        adocomd.CommandText = "SELECT *, {fn YEAR (TglTransaksi) } AS TglTransaksi, Pabrik, AsalBarang, DetailJenisBarang,  NamaBarang, JmlStokAwalPB, JmlTerima, JmlKeluar, HargaNetto FROM V_DataTransaksiBarangNM " & _
            " WHERE TglTransaksi BETWEEN '" & Format(mdTglAwal, "yyyy/01/01 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
            " AND " & strGroup & " LIKE '" & strIsiGroup & "%' AND NamaBarang LIKE '%" & strNama & "%' AND KdRuangan = '" & mstrKdRuangan & "'"
    End If
    
    adocomd.CommandType = adCmdText
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .txtJudul.SetText Judul3
        .unAwal.SetUnboundFieldSource ("{ado.JmlStokAwalPT}")
        .unMasuk.SetUnboundFieldSource ("{ado.JmlTerima}")
        .unKeluar.SetUnboundFieldSource ("{ado.JmlKeluar}")
        .unSaldo.SetUnboundFieldSource Format(("{ado.JmlStokAwalPT} + {ado.JmlTerima} - {ado.JmlKeluar}"), "###,##0")
        .ucNetto.SetUnboundFieldSource Format(("{ado.HargaNetto}"), "###,##0")
        .ucTotal.SetUnboundFieldSource Format(("{ado.JmlStokAwalPT} + {ado.JmlTerima} - {ado.JmlKeluar} * {ado.HargaNetto}"), "###,##0")
    End With
End Sub

Private Sub LaporanTotal()
Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    
    If strGroup = "" Then
        adocomd.CommandText = "SELECT * FROM V_DataTransaksiBarangNM " & _
            " WHERE NamaBarang LIKE '%" & strNama & "%' AND KdRuangan = '" & mstrKdRuangan & "'"
    Else
        adocomd.CommandText = "SELECT * FROM V_DataTransaksiBarangNM " & _
            " WHERE " & strGroup & " LIKE '" & strIsiGroup & "%' AND NamaBarang LIKE '%" & strNama & "%' AND KdRuangan = '" & mstrKdRuangan & "'"
    End If
    
    adocomd.CommandType = adCmdText
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .txtJudul.SetText Judul1
        .unAwal.SetUnboundFieldSource ("{ado.JmlStokAwalPT}")
        .unMasuk.SetUnboundFieldSource ("{ado.JmlTerima}")
        .unKeluar.SetUnboundFieldSource ("{ado.JmlKeluar}")
        .unSaldo.SetUnboundFieldSource Format(("{ado.JmlStokAwalPT} + {ado.JmlTerima} - {ado.JmlKeluar}"), "###,##0")
        .ucNetto.SetUnboundFieldSource Format(("{ado.HargaNetto}"), "###,##0")
        .ucTotal.SetUnboundFieldSource Format(("{ado.JmlStokAwalPT} + {ado.JmlTerima} - {ado.JmlKeluar} * {ado.HargaNetto}"), "###,##0")
    End With
End Sub
