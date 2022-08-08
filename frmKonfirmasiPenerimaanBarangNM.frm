VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmKonfirmasiPenerimaanBarangNM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Konfirmasi Penerimaan Barang"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   12495
   Begin VB.Frame Frame2 
      Caption         =   "Data Order"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   12255
      Begin VB.TextBox txtNoOrder 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtRuanganTujuanPemesanan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtNoKirim 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   6240
         MaxLength       =   15
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtRuanganPengirim 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   6240
         MaxLength       =   15
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtNoKonfirmasi 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   10080
         MaxLength       =   15
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtKdRuanganPengirim 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   6600
         MaxLength       =   15
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpTglOrder 
         Height          =   330
         Left            =   2040
         TabIndex        =   14
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   133300227
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSComCtl2.DTPicker dtpTglKonfirmasi 
         Height          =   330
         Left            =   10080
         TabIndex        =   16
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   78249987
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan Tujuan Pesanan"
         Height          =   210
         Index           =   10
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Order"
         Height          =   210
         Index           =   0
         Left            =   840
         TabIndex        =   22
         Top             =   360
         Width           =   750
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Order"
         Height          =   210
         Index           =   1
         Left            =   840
         TabIndex        =   21
         Top             =   720
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Kirim"
         Height          =   210
         Index           =   2
         Left            =   4560
         TabIndex        =   20
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan Pengirim"
         Height          =   210
         Index           =   3
         Left            =   4560
         TabIndex        =   19
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Konfirmasi"
         Height          =   210
         Index           =   4
         Left            =   8640
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Konfirmasi"
         Height          =   210
         Index           =   5
         Left            =   8640
         TabIndex        =   17
         Top             =   720
         Width           =   1125
      End
   End
   Begin VB.TextBox txtNamaFormPengirim 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   7680
      TabIndex        =   1
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Frame Frame0 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   12255
      Begin VB.TextBox txtIsi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   3015
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   5318
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   9255
      TabIndex        =   2
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   10800
      TabIndex        =   3
      Top             =   5520
      Width           =   1455
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   7
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
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10695
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmKonfirmasiPenerimaanBarangNM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBatal_Click()
    Call subKosong
    Call subSetGrid
    dtpTglOrder.SetFocus
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo Errload
Dim i As Integer
     For i = 1 To fgData.Rows - 1
        With fgData
            If .TextMatrix(i, 4) = "0" Or .TextMatrix(i, 4) = "" Then MsgBox "Silahkan isi jumlah terima", vbExclamation, "Validasi": Exit Sub
        End With
    Next i

'    If Val(fgData.TextMatrix(fgData.Row, 4)) = "" Then MsgBox "Jumlah terima harus diisi", vbExclamation, "Validasi": Exit Sub
    If txtNoKonfirmasi.Text = "" Then
        If sp_Konfirmasi() = False Then Exit Sub
    End If
    
    If mstrKdKelompokBarang = "01" Then
        For i = 1 To fgData.Rows - 1
        With fgData
            If sp_DetailTerimaBarangRuanganNM(.TextMatrix(i, 6), .TextMatrix(i, 7), .TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 5), .TextMatrix(i, 8), .TextMatrix(i, 9)) = False Then Exit Sub
        End With
    Next i
    
    Else
    
    For i = 1 To fgData.Rows - 1
        With fgData
            If sp_DetailTerimaBarangRuangan(.TextMatrix(i, 6), .TextMatrix(i, 7), .TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 5), .TextMatrix(i, 8)) = False Then Exit Sub
        End With
    Next i
    End If

    MsgBox "No Konfirmasi : " & txtNoKonfirmasi.Text, vbInformation, "Informasi"
    Call Add_HistoryLoginActivity("Add_StrukOrder+Add_DetailOrderRuangan")
    Call cmdBatal_Click

Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Function sp_Konfirmasi() As Boolean
On Error GoTo Errload
    sp_Konfirmasi = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoKonfirmasi", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("TglKonfirmasi", adDate, adParamInput, , Format(dtpTglKonfirmasi.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("NamaKonfirmasi", adVarChar, adParamInput, 150, Null)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutKode", adChar, adParamOutput, 10, Null)
    
        .ActiveConnection = dbConn
        .CommandText = "Add_Konfirmasi"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").value <> 0 Then
            MsgBox "Error - Ada kesalahan dalam penyimpanan data struk terima, Hubungi administrator", vbCritical, "Error"
            sp_Konfirmasi = False
        Else
            txtNoKonfirmasi.Text = .Parameters("OutKode").value
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
Errload:
    sp_Konfirmasi = False
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Call msubPesanError
End Function

Private Function sp_DetailTerimaBarangRuangan(f_KdBarang As String, f_KdAsal As String, f_JmlKirim As String, _
                                              f_JmlTerima As String, f_Keterangan As String, f_NoTerima As String) As Boolean
On Error GoTo hell
    sp_DetailTerimaBarangRuangan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, txtNoKonfirmasi.Text)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("JmlKirim", adDouble, adParamInput, , f_JmlKirim)
        .Parameters.Append .CreateParameter("JmlTerima", adDouble, adParamInput, , f_JmlTerima)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder.Text)
        .Parameters.Append .CreateParameter("NoKirim", adChar, adParamInput, 10, txtNoKirim.Text)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, f_Keterangan)
        .Parameters.Append .CreateParameter("KdRuanganPenerima", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdRuanganPengirim", adChar, adParamInput, 3, txtKdRuanganPengirim.Text)
        .Parameters.Append .CreateParameter("NoKonfirmasi", adChar, adParamInput, 10, txtNoKonfirmasi.Text)
        .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, f_NoTerima)
        
        .ActiveConnection = dbConn
        .CommandText = "Add_DetailTerimaBarangRuangan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").value <> 0 Then
            MsgBox "Error - Ada kesalahan dalam penyimpanan data, Hubungi administrator", vbCritical, "Error"
            sp_DetailTerimaBarangRuangan = False
        Else
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
hell:
    sp_DetailTerimaBarangRuangan = False
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Call msubPesanError
End Function

Private Function sp_DetailTerimaBarangRuanganNM(f_KdBarang As String, f_KdAsal As String, f_JmlKirim As String, _
                                              f_JmlTerima As String, f_Keterangan As String, f_NoTerima As String, f_NoRegister As String) As Boolean
On Error GoTo hell
    sp_DetailTerimaBarangRuanganNM = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, txtNoKonfirmasi.Text)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("JmlKirim", adDouble, adParamInput, , f_JmlKirim)
        .Parameters.Append .CreateParameter("JmlTerima", adDouble, adParamInput, , f_JmlTerima)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder.Text)
        .Parameters.Append .CreateParameter("NoKirim", adChar, adParamInput, 10, txtNoKirim.Text)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, f_Keterangan)
        .Parameters.Append .CreateParameter("KdRuanganPenerima", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdRuanganPengirim", adChar, adParamInput, 3, txtKdRuanganPengirim.Text)
        .Parameters.Append .CreateParameter("NoKonfirmasi", adChar, adParamInput, 10, txtNoKonfirmasi.Text)
        .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, f_NoTerima)
        .Parameters.Append .CreateParameter("NoRegisterAset", adVarChar, adParamInput, 15, f_NoRegister)
        
        .ActiveConnection = dbConn
        .CommandText = "Add_DetailTerimaBarangRuanganNonMedis"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").value <> 0 Then
            MsgBox "Error - Ada kesalahan dalam penyimpanan data, Hubungi administrator", vbCritical, "Error"
            sp_DetailTerimaBarangRuanganNM = False
        Else
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
hell:
    sp_DetailTerimaBarangRuanganNM = False
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Call msubPesanError
End Function

Private Sub cmdTutup_Click()
'    If subbolSimpan = False Then
'        If MsgBox("Simpan data Pemakaian Obat dan Alat Kesehatan", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
'            Call cmdSimpan_Click
'            Exit Sub
'        End If
'    End If
    Unload Me
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
            Call SubLoadText
            txtIsi.Text = Trim(fgData.TextMatrix(fgData.Row, fgData.Col))
            txtIsi.SelStart = 0
            txtIsi.SelLength = Len(txtIsi.Text)
        
        Case vbKeyDelete
            If fgData.Row = fgData.Rows - 1 Then Exit Sub
            fgData.RemoveItem fgData.Row
    
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
'    If KeyAscii = 27 Then dgObatAlkes.Visible = False
End Sub

Private Sub Form_Load()
On Error GoTo Errload
    Call PlayFlashMovie(Me)

    Call centerForm(Me, MDIUtama)
'    dtpTglOrder.Value = Now
    dtpTglKonfirmasi.value = Now
    Call subSetGrid
    
Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
If txtNamaFormPengirim.Text = "frmInfoPesanBarangNM" Then
    frmInfoPesanBarangNM.Enabled = True
    frmInfoPesanBarangNM.cmdTampilkan_Click
'Else
'    frmInfoPesanBarang.Enabled = True
'    frmInfoPesanBarangNM.cmdTampilkan_Click
End If
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
Dim i As Integer
    If KeyAscii = 13 Then
        Select Case fgData.Col
            Case 4
                If Val(txtIsi.Text) = 0 Then
                    txtIsi.Text = 0
                Else
                    fgData.TextMatrix(fgData.Row, 4) = txtIsi.Text
                    If Val(fgData.TextMatrix(fgData.Row, 4)) > Val(fgData.TextMatrix(fgData.Row, 3)) Then
                        MsgBox "Jumlah terima lebih besar dari jumlah kirim", vbInformation, vbOKOnly
                        txtIsi.Text = ""
                        txtIsi.SetFocus
                        Exit Sub
                    End If
                End If
            Case 5
                If KeyAscii = 13 Then
                    fgData.TextMatrix(fgData.Row, 5) = txtIsi.Text
                    txtIsi.Visible = False
                End If
        End Select
                        
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
    
    Select Case fgData.Col
        Case 4
            If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
        Case 5
            If KeyAscii = 13 Then
                fgData.TextMatrix(fgData.Row, 5) = txtIsi.Text
                txtIsi.Visible = False
            End If
    End Select
Exit Sub
End Sub

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

Private Sub subKosong()
    txtNoOrder.Text = ""
    dtpTglOrder.value = Now
    txtRuanganTujuanPemesanan.Text = ""
    txtNoKirim.Text = ""
    txtRuanganPengirim.Text = ""
    txtNoKonfirmasi.Text = ""
End Sub

Private Sub subSetGrid()
On Error GoTo Errload
    With fgData
        .Clear
        .Rows = 2
        .Cols = 10
        
        .RowHeight(0) = 500
        
        .TextMatrix(0, 0) = "Nama Barang"
        .TextMatrix(0, 1) = "Asal Barang"
        .TextMatrix(0, 2) = "Jml Order"
        .TextMatrix(0, 3) = "Jml Kirim"
        .TextMatrix(0, 4) = "Jml Terima"
        .TextMatrix(0, 5) = "Keterangan"
        .TextMatrix(0, 6) = "KdBarang"
        .TextMatrix(0, 7) = "KdAsal"
        .TextMatrix(0, 8) = "NoTerima"
        .TextMatrix(0, 9) = "NoRegister"
    
        .ColWidth(0) = 3800
        .ColWidth(1) = 1800
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 3350
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
     
            
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignCenterCenter
    End With

Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub SubLoadText()
Dim i As Integer
    txtIsi.Left = fgData.Left
    Select Case fgData.Col
        Case 4
            txtIsi.MaxLength = 5
        Case 5
            txtIsi.MaxLength = 50
        Case Else
            Exit Sub
    End Select
    
    Select Case fgData.Col
        Case 4
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
        Case 5
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
    End Select
End Sub


