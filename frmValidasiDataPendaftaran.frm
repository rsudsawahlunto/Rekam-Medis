VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmValidasiDataPendaftaran 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 -Validasi Data"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmValidasiDataPendaftaran.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   13500
   Begin VB.ComboBox cbPulang 
      Appearance      =   0  'Flat
      Height          =   330
      ItemData        =   "frmValidasiDataPendaftaran.frx":0CCA
      Left            =   11880
      List            =   "frmValidasiDataPendaftaran.frx":0CD4
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox cbKeluar 
      Appearance      =   0  'Flat
      Height          =   330
      ItemData        =   "frmValidasiDataPendaftaran.frx":0CDE
      Left            =   10560
      List            =   "frmValidasiDataPendaftaran.frx":0CE8
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pbData 
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   7200
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Max             =   200
      Scrolling       =   1
   End
   Begin VB.OptionButton optIGD 
      Caption         =   "Pasien IGD"
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.OptionButton optRI 
      Caption         =   "Pasien Rawat Inap"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdPerbaiki 
      Caption         =   "&Perbaiki Data"
      Height          =   495
      Left            =   8400
      TabIndex        =   5
      Top             =   7200
      Width           =   1695
   End
   Begin VB.TextBox txtIsi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   4320
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdValidasiData 
      Caption         =   "&Validasi Data"
      Height          =   495
      Left            =   10080
      TabIndex        =   6
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   11760
      TabIndex        =   7
      Top             =   7200
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   5415
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   9551
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      Appearance      =   0
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmValidasiDataPendaftaran.frx":0CF2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   11640
      Picture         =   "frmValidasiDataPendaftaran.frx":36B3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmValidasiDataPendaftaran.frx":443B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12135
   End
End
Attribute VB_Name = "frmValidasiDataPendaftaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim substrNomorRetur As String
Dim subbolSimpan As Boolean
Dim sqlQuery As String
Dim rsQuery As New ADODB.recordset
Dim i, j, iData As Integer

Private Sub subLoadData(Optional s_Kriteria As String)
    On Error Resume Next
    Dim strStatusData As String
    Dim QueryRI As String
    Dim rsRI As New ADODB.recordset
    Dim QueryIGD As String
    Dim rsIGD As New ADODB.recordset

    Call subSetGrid

    If optRI.value = True Then ' ri
        QueryRI = " SELECT NoPendaftaran, NoCM, NamaPasien, RuanganPelayanan, TglMasuk, KelasPelayanan, StatusKeluar, StatusPulang,NoPakai" & _
        " FROM V_DaftarPasienRI4Validasi Where (StatusKeluar = '" & cbKeluar & "'  AND StatusPulang ='" & cbPulang & "' )  "
        Call msubRecFO(rsRI, QueryRI)
        If rsRI.EOF = True Then Exit Sub

        For i = 1 To rsRI.RecordCount
            pbData.value = i
            pbData.Max = rsRI.RecordCount
            DoEvents
            With fgData
                .TextMatrix(i, 0) = rsRI("NoPendaftaran")
                .TextMatrix(i, 1) = rsRI("NoCM")
                .TextMatrix(i, 2) = rsRI("NamaPasien")
                .TextMatrix(i, 3) = rsRI("RuanganPelayanan")

                .TextMatrix(i, 4) = rsRI("TglMasuk")
                .TextMatrix(i, 5) = rsRI("KelasPelayanan")
                If IsNull(rsRI("StatusKeluar")) Then
                    .TextMatrix(i, 6) = "-"
                    .TextMatrix(i, 9) = "-"
                Else
                    .TextMatrix(i, 6) = rsRI("StatusKeluar")
                    .TextMatrix(i, 9) = rsRI("StatusKeluar")
                End If

                If IsNull(rsRI("StatusPulang")) Then
                    .TextMatrix(i, 7) = "-"
                    .TextMatrix(i, 10) = "-"
                Else
                    .TextMatrix(i, 7) = rsRI("StatusPulang")
                    .TextMatrix(i, 10) = rsRI("StatusPulang")
                End If

                .TextMatrix(i, 8) = rsRI("NoPakai")

                For j = 0 To 7
                    .Col = j
                    If .Col = 6 Then ' status keluar
                        If .TextMatrix(i, 6) = "T" Then
                            .Row = i
                            .CellBackColor = vbRed
                            .CellForeColor = vbWhite
                        End If
                    End If
                    If .Col = 7 Then ' status pulang
                        If .TextMatrix(i, 7) = "T" Then
                            .Row = i
                            .CellBackColor = vbRed
                            .CellForeColor = vbWhite
                        End If
                    End If
                Next j

                .Rows = .Rows + 1
                rsRI.MoveNext
            End With
            pbData.value = Int(pbData.value) + 1
        Next i

    Else ' IGD
        Set rsIGD = Nothing
        QueryIGD = " SELECT     NoPendaftaran, NoCM, NamaPasien, RuanganPelayanan, TglMasuk, KelasPelayanan, StatusPulang" & _
        " FROM    V_DaftarPasienIGD4Validasi Where  StatusPulang ='" & cbPulang & "' "
        Call msubRecFO(rsIGD, QueryIGD)
        If rsIGD.EOF = True Then Exit Sub

        For i = 1 To rsIGD.RecordCount
            pbData.value = i
            pbData.Max = rsIGD.RecordCount
            DoEvents
            With fgData
                .TextMatrix(i, 0) = rsIGD("NoPendaftaran")
                .TextMatrix(i, 1) = rsIGD("NoCM")
                .TextMatrix(i, 2) = rsIGD("NamaPasien")
                .TextMatrix(i, 3) = rsIGD("RuanganPelayanan")

                .TextMatrix(i, 4) = rsIGD("TglMasuk")
                .TextMatrix(i, 5) = rsIGD("KelasPelayanan")

                If IsNull(rsIGD("StatusPulang")) Then
                    .TextMatrix(i, 6) = "-"
                    .TextMatrix(i, 7) = "-"
                Else
                    .TextMatrix(i, 6) = rsIGD("StatusPulang")
                    .TextMatrix(i, 7) = rsIGD("StatusPulang")
                End If

                For j = 0 To 7
                    .Col = j
                    If .Col = 6 Then ' status pulang
                        If .TextMatrix(i, 6) = "T" Then
                            .Row = i
                            .CellBackColor = vbRed
                            .CellForeColor = vbWhite
                        End If
                    End If

                Next j

                .Rows = .Rows + 1
                rsIGD.MoveNext
            End With
            pbData.value = Int(pbData.value) + 1
        Next i

    End If

    MsgBox "Load Data sukses ", vbInformation, vbOK, "Informasi"
    pbData.value = 0

End Sub

Private Sub subLoadText()
    txtIsi.Left = fgData.Left
    Select Case fgData.Col
        Case 6, 7
            txtIsi.MaxLength = 1
        Case Else
            Exit Sub
    End Select

    For i = 0 To fgData.Col - 1
        txtIsi.Left = txtIsi.Left + fgData.ColWidth(i)
    Next i
    txtIsi.Visible = True
    txtIsi.Top = fgData.Top - 7

    For i = 0 To fgData.Row - 1
        txtIsi.Top = txtIsi.Top + fgData.RowHeight(i)
    Next i

    If fgData.TopRow > 1 Then
        txtIsi.Top = txtIsi.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If

    txtIsi.Width = fgData.ColWidth(fgData.Col)
    txtIsi.Height = fgData.RowHeight(fgData.Row)

    txtIsi.Visible = True
    txtIsi.SelStart = Len(txtIsi.Text)
    txtIsi.SetFocus
End Sub

Private Sub subSetGrid()
    If optRI.value = True Then
        With fgData
            .Rows = 2
            .Cols = 11

            .RowHeight(0) = 400
            .TextMatrix(0, 0) = "No Pendaftaran"
            .TextMatrix(0, 1) = "NoCM"
            .TextMatrix(0, 2) = "Nama Pasien"
            .TextMatrix(0, 3) = "Ruang Pelayanan"
            .TextMatrix(0, 4) = "Tgl Masuk"
            .TextMatrix(0, 5) = "Kelas Pelayanan"
            .TextMatrix(0, 6) = "Status Keluar"
            .TextMatrix(0, 7) = "Status Pulang"
            .TextMatrix(0, 8) = "No Pakai"

            .ColWidth(0) = 1400
            .ColAlignment(0) = flexAlignLeftCenter
            .ColAlignment(1) = flexAlignLeftCenter
            .ColWidth(1) = 1000
            .ColWidth(2) = 3000
            .ColWidth(3) = 2200
            .ColWidth(4) = 1800
            .ColWidth(5) = 1200
            .ColWidth(6) = 1200
            .ColWidth(7) = 1200
            .ColAlignment(6) = flexAlignCenterCenter
            .ColAlignment(7) = flexAlignCenterCenter
            .ColWidth(8) = 0
            .ColWidth(9) = 0
            .ColWidth(10) = 0
        End With
    Else
        With fgData
            .Rows = 2
            .Cols = 9

            .RowHeight(0) = 400
            .TextMatrix(0, 0) = "No Pendaftaran"
            .TextMatrix(0, 1) = "NoCM"
            .TextMatrix(0, 2) = "Nama Pasien"
            .TextMatrix(0, 3) = "Ruang Pelayanan"
            .TextMatrix(0, 4) = "Tgl Masuk"
            .TextMatrix(0, 5) = "Kelas Pelayanan"
            .TextMatrix(0, 6) = "Status Pulang"
            .TextMatrix(0, 7) = "Status Pulang temp"

            .ColWidth(0) = 1400
            .ColAlignment(0) = flexAlignLeftCenter
            .ColAlignment(1) = flexAlignLeftCenter
            .ColWidth(1) = 1000
            .ColWidth(2) = 3000
            .ColWidth(3) = 2200
            .ColWidth(4) = 2000
            .ColWidth(5) = 1400
            .ColWidth(6) = 1300
            .ColAlignment(6) = flexAlignCenterCenter
            .ColWidth(7) = 0
        End With
    End If
End Sub

Private Sub cbKeluar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cbPulang.SetFocus
End Sub

Private Sub cmdPerbaiki_Click()
    On Error GoTo hell_
    If optRI.value = True Then
        With fgData
            If MsgBox("Yakin akan memperbaiki data pasien : " & .TextMatrix(.Row, 2), vbInformation + vbYesNo, "validasi") = vbNo Then Exit Sub
            If .TextMatrix(.Row, 6) <> .TextMatrix(.Row, 9) Then
                sqlQuery = "update PemakaianKamar set StatusKeluar ='" & .TextMatrix(.Row, 6) & "' where NoPendaftaran ='" & .TextMatrix(.Row, 0) & "'  and NoCM ='" & .TextMatrix(.Row, 1) & "' and NoPakai = '" & .TextMatrix(.Row, 8) & "' "  ' "
                Call msubRecFO(rsQuery, sqlQuery)
                Set rsQuery = Nothing
            End If

            If .TextMatrix(.Row, 7) <> .TextMatrix(.Row, 10) Then
                strSQL = "select NoPakai from PasienKeluarKamar where NoPakai = '" & .TextMatrix(.Row, 8) & "'"
                Call msubRecFO(rs, strSQL)
                If rs.EOF = True Then MsgBox "Pasien Belum Keluar Kamar", vbExclamation, "Validasi": Exit Sub
                sqlQuery = "update RegistrasiRI set StatusPulang ='" & .TextMatrix(.Row, 7) & "' where NoPendaftaran ='" & .TextMatrix(.Row, 0) & "'  "
                Call msubRecFO(rsQuery, sqlQuery)
                Set rsQuery = Nothing
            End If
        End With
    Else
        With fgData
            If MsgBox("Yakin akan memperbaiki data pasien : " & .TextMatrix(.Row, 2), vbInformation + vbYesNo, "validasi") = vbNo Then Exit Sub
            If .TextMatrix(.Row, 6) <> .TextMatrix(.Row, 7) Then
                sqlQuery = "update RegistrasiIGD set StatusPulang ='" & .TextMatrix(.Row, 6) & "' where NoPendaftaran ='" & .TextMatrix(.Row, 0) & "'  "
                Call msubRecFO(rsQuery, sqlQuery)
                Set rsQuery = Nothing
            End If

        End With
    End If

    MsgBox "Proses perbaiki data berhasil ", vbInformation, "Informasi"
    cmdValidasiData.SetFocus

    Exit Sub
hell_:
    msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdValidasiData_Click()
    On Error Resume Next
    Call subLoadData
End Sub

Private Sub fgData_DblClick()
    Call fgData_KeyDown(13, 0)
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)

    Select Case KeyCode
        Case 13
            If fgData.TextMatrix(fgData.Row, 2) = "" Then Exit Sub
            Call subLoadText
            txtIsi.Text = Trim(fgData.TextMatrix(fgData.Row, fgData.Col))
            txtIsi.SelStart = 0
            txtIsi.SelLength = Len(txtIsi.Text)
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    optRI.value = True
    Call subLoadData
    pbData.value = 0.0001
End Sub

Private Sub optIGD_Click()
    Call optRI_Click
End Sub

Private Sub optIGD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cbKeluar.SetFocus
End Sub

Private Sub optRI_Click()
    If optRI.value = True Then
        cbKeluar.Visible = True
        cbPulang.Visible = True
        optRI.FontBold = True
        optIGD.FontBold = False
    Else
        cbKeluar.Visible = False
        cbPulang.Visible = True
        optRI.FontBold = False
        optIGD.FontBold = True
    End If
End Sub

Private Sub optRI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cbKeluar.SetFocus
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If KeyAscii = 13 Then

        fgData.TextMatrix(fgData.Row, fgData.Col) = txtIsi.Text
        txtIsi.Visible = False

        If fgData.RowPos(fgData.Row) >= fgData.Height - 360 Then
            fgData.SetFocus
            SendKeys "{DOWN}"
            Exit Sub
        End If
        fgData.SetFocus
    ElseIf KeyAscii = 27 Then
        txtIsi.Visible = False
        fgData.SetFocus
    End If
End Sub

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

