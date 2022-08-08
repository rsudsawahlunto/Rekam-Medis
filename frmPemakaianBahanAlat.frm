VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPemakaianBahanAlat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pemakaian Bahan dan Alat"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPemakaianBahanAlat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmPemakaianBahanAlat.frx":0CCA
   ScaleHeight     =   6750
   ScaleWidth      =   13110
   Begin VB.TextBox txtKdRuanganPerujuk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   3600
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   5880
      Width           =   13095
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   9480
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   11280
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame8 
      Height          =   3855
      Left            =   0
      TabIndex        =   2
      Top             =   2040
      Width           =   13095
      Begin MSDataListLib.DataCombo dcPelayanan 
         Height          =   330
         Left            =   2640
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid dgObatAlkes 
         Height          =   2535
         Left            =   840
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   16
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcAsalBarang 
         Height          =   330
         Left            =   5640
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtIsi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   7800
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid fgAlkes 
         Height          =   3615
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   6376
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Pemakaian"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   13095
      Begin VB.TextBox txtKeperluan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4320
         TabIndex        =   5
         Top             =   480
         Width           =   5175
      End
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   9720
         TabIndex        =   15
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
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
         CustomFormat    =   "dd MMMM yyyy HH:mm"
         Format          =   137035779
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSDataListLib.DataCombo dcNamaPelayanan 
         Height          =   330
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Pemakaian"
         Height          =   210
         Left            =   9720
         TabIndex        =   16
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Pemakaian Untuk "
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Keperluan"
         Height          =   210
         Left            =   4320
         TabIndex        =   6
         Top             =   240
         Width           =   810
      End
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
      Picture         =   "frmPemakaianBahanAlat.frx":1994
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   11280
      Picture         =   "frmPemakaianBahanAlat.frx":4355
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPemakaianBahanAlat.frx":50DD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmPemakaianBahanAlat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tempStatusTampil As Boolean
Dim subJenisHargaNetto  As Integer

Private Sub dcNamaPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcNamaPelayanan.MatchedWithList = True Then txtKeperluan.SetFocus
        strSQL = "select  kdpelayananrs,namapelayanan from V_ListPemakaianBahan  WHERE (namapelayanan LIKE '%" & dcNamaPelayanan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcNamaPelayanan.BoundText = rs(0).value
        dcNamaPelayanan.Text = rs(1).value
    End If
End Sub

Private Sub dcPelayanan_Change()
    On Error GoTo errLoad
    fgAlkes.TextMatrix(fgAlkes.Row, 0) = dcPelayanan.Text
    fgAlkes.TextMatrix(fgAlkes.Row, 12) = dcPelayanan.BoundText
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcPelayanan_Change
        dcPelayanan.Visible = False
        fgAlkes.Col = 1
        fgAlkes.SetFocus
    End If
End Sub

Private Sub dcPelayanan_LostFocus()
    dcPelayanan.Visible = False
End Sub

Private Sub cmdSimpan_Click()
    Dim i As Integer
    On Error GoTo aa
     If Periksa("datacombo", dcNamaPelayanan, "Nama Pelayanan kosong") = False Then Exit Sub
    If fgAlkes.TextMatrix(1, 10) = "" Then Exit Sub
    
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn
    With fgAlkes
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 10) = "" Then GoTo lanjut_
            If sp_PemakaianBahanAlat(.TextMatrix(i, 10), .TextMatrix(i, 11), .TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 6), .TextMatrix(i, 8)) = False Then Exit Sub
lanjut_:
        Next i
    End With

    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    cmdSimpan.Enabled = False
    Call subLoadGridSource
    Exit Sub
aa:
    msubPesanError
End Sub

Private Function sp_PemakaianBahanAlat(f_KdBarang As String, f_KdAsal As String, _
    f_Satuan As String, f_Jumlah As Double, f_HargaSatuan As Currency, f_HargaBeli As String) As Boolean
    On Error GoTo errLoad
    Dim i As Integer
    sp_PemakaianBahanAlat = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdPelayananRS", adVarChar, adParamInput, 6, dcNamaPelayanan.BoundText)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("TglPemakaian", adDate, adParamInput, , Format(dtpTglPeriksa.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("SatuanJml", adChar, adParamInput, 1, f_Satuan)
        .Parameters.Append .CreateParameter("JmlBarang", adDouble, adParamInput, , CDbl(f_Jumlah))

        .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , f_HargaSatuan)
        .Parameters.Append .CreateParameter("Keperluan", adVarChar, adParamInput, 200, IIf(txtKeperluan.Text = "", Null, txtKeperluan.Text))
        .Parameters.Append .CreateParameter("HargaBeli", adCurrency, adParamInput, , f_HargaBeli)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("idUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        '.Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, strNoTerima)
        .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, "0000000000")

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PemakaianBahanAlat"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_PemakaianBahanAlat = False
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Exit Function
errLoad:
    sp_PemakaianBahanAlat = False
    msubPesanError
End Function

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdHapus_Click()
    With fgAlkes
        If .Row = .Rows Then Exit Sub
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        If .TextMatrix(.Row, 12) = "SudahAda" Then Exit Sub
        msubRemoveItem fgAlkes, .Row
    End With
End Sub

Private Sub dgObatAlkes_DblClick()
    On Error GoTo errLoad
    Dim i As Integer
    Dim tempSettingDataPendukung As Integer
    Dim curHargaBrg As Currency
    Dim strNoTerima As String

    For i = 0 To fgAlkes.Rows - 1
        If dgObatAlkes.Columns("KdBarang") = fgAlkes.TextMatrix(i, 10) And dgObatAlkes.Columns("KdAsal") = fgAlkes.TextMatrix(i, 11) Then
            MsgBox "Data tersebut sudah diinput", vbExclamation, "Validasi"
            dgObatAlkes.Visible = False
            fgAlkes.SetFocus: fgAlkes.Row = i
            Exit Sub
        End If
    Next i

    strNoTerima = ""
    Set rsb = Nothing
    Call msubRecFO(rsb, "select dbo.TakeNoFIFO_F('" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "','" & mstrKdRuangan & "') as NoFIFO")
    strNoTerima = IIf(IsNull(rsb("NoFIFO")), "0000000000", rsb("NoFIFO"))

    For i = 0 To fgAlkes.Rows - 1
        If dgObatAlkes.Columns("KdBarang") = fgAlkes.TextMatrix(i, 10) And dgObatAlkes.Columns("KdAsal") = fgAlkes.TextMatrix(i, 11) Then
            MsgBox "Data tersebut sudah diinput", vbExclamation, "Validasi"
            dgObatAlkes.Visible = False
            fgAlkes.SetFocus: fgAlkes.Row = i
            Exit Sub
        End If
    Next i

    tempStatusTampil = True
    fgAlkes.TextMatrix(fgAlkes.Row, 1) = dgObatAlkes.Columns("NamaBarang")
    fgAlkes.TextMatrix(fgAlkes.Row, 2) = dgObatAlkes.Columns("AsalBarang")
    fgAlkes.TextMatrix(fgAlkes.Row, 3) = dgObatAlkes.Columns("Satuan")
    fgAlkes.TextMatrix(fgAlkes.Row, 10) = dgObatAlkes.Columns("KdBarang")
    fgAlkes.TextMatrix(fgAlkes.Row, 11) = dgObatAlkes.Columns("KdAsal")

    fgAlkes.TextMatrix(fgAlkes.Row, 9) = strNoTerima
    curHargaBrg = 0

    strSQL = ""
    Set rsb = Nothing
    strSQL = "SELECT dbo.FB_TakeHargaNettoOA('2222222222','01','" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "','" & dgObatAlkes.Columns("Satuan") & "', '" & mstrKdRuangan & "','" & strNoTerima & "') AS HargaBarang"
    Call msubRecFO(rsb, strSQL)
    If rsb.EOF = True Then curHargaBrg = 0 Else curHargaBrg = rsb(0).value
    fgAlkes.TextMatrix(fgAlkes.Row, 6) = Format(curHargaBrg, "#,###")

    strSQL = ""
    Set rs = Nothing
    strSQL = "Select JmlStok as Stok From StokRuangan Where KdBarang='" & dgObatAlkes.Columns("KdBarang") & "' and KdAsal='" & dgObatAlkes.Columns("KdAsal") & "' and KdRuangan='" & mstrKdRuangan & "'"
    Call msubRecFO(rsb, strSQL)
    If rsb.EOF Then
        fgAlkes.TextMatrix(fgAlkes.Row, 5) = 0
    Else
        fgAlkes.TextMatrix(fgAlkes.Row, 5) = IIf(IsNull(rsb("Stok")), 0, rsb("Stok"))
    End If

    tempStatusTampil = False
    dgObatAlkes.Visible = False
    fgAlkes.TextMatrix(fgAlkes.Row, 4) = 0
    fgAlkes.TextMatrix(fgAlkes.Row, 7) = fgAlkes.TextMatrix(fgAlkes.Row, 4) * fgAlkes.TextMatrix(fgAlkes.Row, 6)
    fgAlkes.TextMatrix(fgAlkes.Row, 8) = curHargaBrg
    fgAlkes.SetFocus
    fgAlkes.Col = 4

    Exit Sub
errLoad:
End Sub

Private Sub dgObatAlkes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dgObatAlkes_DblClick
    End If
End Sub

Private Sub dtpTglPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then fgAlkes.SetFocus: fgAlkes.Col = 1
End Sub

Private Sub fgAlkes_DblClick()
    Call fgAlkes_KeyDown(13, 0)
End Sub

Private Sub fgAlkes_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim strKdBrg As String
    Dim strKdAsal As String

    Select Case KeyCode
        Case 13
            If fgAlkes.Col = fgAlkes.Cols - 1 Then
                If fgAlkes.TextMatrix(fgAlkes.Row, 2) <> "" Then
                    If fgAlkes.TextMatrix(fgAlkes.Rows - 1, 2) <> "" Then fgAlkes.Rows = fgAlkes.Rows + 1
                    fgAlkes.Row = fgAlkes.Rows - 1
                    fgAlkes.Col = 1
                Else
                    fgAlkes.Col = 1
                End If
            Else
                For i = 0 To fgAlkes.Cols - 2
                    If fgAlkes.Col = fgAlkes.Cols - 1 Then Exit For
                    fgAlkes.Col = fgAlkes.Col + 1
                    If fgAlkes.ColWidth(fgAlkes.Col) > 0 Then Exit For
                Next i
            End If
            fgAlkes.SetFocus

            If fgAlkes.Col > 7 Then
                fgAlkes.Rows = fgAlkes.Rows + 1
                fgAlkes.Row = fgAlkes.Rows - 1
                fgAlkes.Col = 0
                fgAlkes.SetFocus
            End If

        Case 27
            dgObatAlkes.Visible = False

        Case vbKeyDelete
            'validasi FIFO
            If bolStatusFIFO = True Then
                If fgAlkes.CellBackColor = vbRed Then
                    MsgBox "Data yang barisnya berwarna merah tidak bisa di edit", vbExclamation, "validasi"
                    fgAlkes.SetFocus
                    Exit Sub
                End If

                With fgAlkes
                    i = .Rows - 1
                    strKdBrg = .TextMatrix(.Row, 10)
                    strKdAsal = .TextMatrix(.Row, 11)
                    Do While i <> 0 'khusus utk delete dr keyboard diset 0 agar ke cek keseluruhannya
                        If .TextMatrix(i, 10) <> "" Then
                            If (strKdBrg = .TextMatrix(i, 10)) And (strKdAsal = .TextMatrix(i, 11)) Then
                                If .Row = .Rows Then Exit Sub
                                If .Row = 0 Then Exit Sub
                                .Row = i
                                If .Rows = 2 Then
                                    For i = 0 To .Cols - 1
                                        .TextMatrix(1, i) = ""
                                    Next i
                                    Exit Sub
                                Else
                                    .RemoveItem .Row
                                End If
                                .Row = i - 1
                            End If
                        End If
                        i = i - 1
                    Loop
                End With
            Else
                fgAlkes.RemoveItem fgAlkes.Row
            End If
            'end FIFO
    End Select
End Sub

Private Sub fgAlkes_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    txtIsi.Text = ""
    If Not (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Or KeyAscii = 32 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
        Exit Sub
    End If

    Select Case fgAlkes.Col
        Case 0 'nama pemeriksaan
            Call subLoadDataCombo(dcPelayanan)

        Case 1 'Nama Barang
            txtIsi.MaxLength = 0
            Call subLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)

        Case 2 'asal barang
            Call subLoadDataCombo(dcAsalBarang)

        Case 3 'satauan hasil

        Case 4 'Jumlah
            txtIsi.MaxLength = 4
            Call subLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpTglPeriksa.value = Format(Now, "yyyy/MMMM/dd HH:mm:ss")
    Call subLoadGridSource
    Call subLoadDcSource

    subJenisHargaNetto = 1

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadText()
    Dim i As Integer
    txtIsi.Left = fgAlkes.Left
    For i = 0 To fgAlkes.Col - 1
        txtIsi.Left = txtIsi.Left + fgAlkes.ColWidth(i)
    Next i
    txtIsi.Visible = True
    txtIsi.Top = fgAlkes.Top - 7

    For i = 0 To fgAlkes.Row - 1
        txtIsi.Top = txtIsi.Top + fgAlkes.RowHeight(i)
    Next i

    If fgAlkes.TopRow > 1 Then
        txtIsi.Top = txtIsi.Top - ((fgAlkes.TopRow - 1) * fgAlkes.RowHeight(1))
    End If

    txtIsi.Width = fgAlkes.ColWidth(fgAlkes.Col)
    txtIsi.Height = fgAlkes.RowHeight(fgAlkes.Row)

    txtIsi.Visible = True
    txtIsi.SelStart = Len(txtIsi.Text)
    txtIsi.SetFocus
End Sub

Private Sub txtIsi_Change()
    On Error GoTo errLoad
    Dim i As Integer
    Select Case fgAlkes.Col
        Case 0  ' nama pemeriksaan

        Case 1 ' nama barang
            If tempStatusTampil = True Then Exit Sub
            strSQL = "execute CariBarangNStokMedis_V '" & txtIsi.Text & "%','" & mstrKdRuangan & "'"
            Call msubRecFO(dbRst, strSQL)

            Set dgObatAlkes.DataSource = dbRst
            With dgObatAlkes
                For i = 0 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next i

                .Columns("KdBarang").Width = 1500
                .Columns("NamaBarang").Width = 3000
                .Columns("JenisBarang").Width = 1500
                .Columns("Kekuatan").Width = 1000
                .Columns("AsalBarang").Width = 1000
                .Columns("Satuan").Width = 675

                .Top = txtIsi.Top + txtIsi.Height
                .Left = txtIsi.Left
                .Visible = True
            End With

        Case Else
            dgObatAlkes.Visible = False
            Exit Sub
    End Select

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDataCombo(s_DcName As Object)
    Dim i As Integer
    s_DcName.Left = fgAlkes.Left
    For i = 0 To fgAlkes.Col - 1
        s_DcName.Left = s_DcName.Left + fgAlkes.ColWidth(i)
    Next i
    s_DcName.Visible = True
    s_DcName.Top = fgAlkes.Top - 7

    For i = 0 To fgAlkes.Row - 1
        s_DcName.Top = s_DcName.Top + fgAlkes.RowHeight(i)
    Next i

    If fgAlkes.TopRow > 1 Then
        s_DcName.Top = s_DcName.Top - ((fgAlkes.TopRow - 1) * fgAlkes.RowHeight(1))
    End If

    s_DcName.Width = fgAlkes.ColWidth(fgAlkes.Col)
    s_DcName.Height = fgAlkes.RowHeight(fgAlkes.Row)

    s_DcName.Visible = True
    s_DcName.SetFocus
End Sub

Private Sub txtIsi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then If dgObatAlkes.Visible = True Then If dgObatAlkes.ApproxCount > 0 Then dgObatAlkes.SetFocus
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    Dim i, j As Integer
    Dim dblSelisih As Double
    Dim intRowTemp As Integer
    Dim strNoTerima As String
    Dim curHargaBrg As Currency
    Dim dblSelisihNow As Double
    Dim dblJmlStokMax As Double
    Dim strKdBrg As String
    Dim strKdAsal As String
    Dim dblJmlTerkecil As Double
    Dim dblTotalStokK As Double
    
    If KeyAscii = 39 Then KeyAscii = 0

    If KeyAscii = 13 Then
        With fgAlkes
            Select Case fgAlkes.Col
                Case 0
                    If dgObatAlkes.Visible = True Then
                        dgObatAlkes.SetFocus
                        Exit Sub
                    Else
                        .SetFocus
                        .Col = 1
                    End If

                Case 1
                    If dgObatAlkes.Visible = True Then
                        dgObatAlkes.SetFocus
                        Exit Sub
                    Else
                        .SetFocus
                        .Col = 4
                    End If
                Case 4
                        
                    
                    If fgAlkes.TextMatrix(fgAlkes.Row, 1) = "" Then
                        MsgBox "Obat harus diisi", vbExclamation, "Validasi":
                        txtIsi.Text = ""
                        Exit Sub
                        
                    End If
                        
                    If fgAlkes.TextMatrix(fgAlkes.Row, 4) = "" Then
                        MsgBox "Jumlah Obat harus diisi", vbExclamation, "Validasi":
                         txtIsi.Text = ""
                        Exit Sub
                       
                    End If
                        
                    If Trim(txtIsi.Text) = "," Then txtIsi.Text = 0
                    If Trim(txtIsi.Text) = "" Then txtIsi.Text = 0

                    If (.TextMatrix(.Row, 3) = "S") Then
                        If CDbl(txtIsi.Text) > CDbl(.TextMatrix(.Row, 5)) Then
                            MsgBox "Jumlah lebih besar dari stock (" & .TextMatrix(.Row, 5) & ")", vbExclamation, "Validasi"
                            txtIsi.SelStart = 0: txtIsi.SelLength = Len(txtIsi.Text)
                            Exit Sub
                        End If
                    ElseIf (.TextMatrix(.Row, 3) = "K") Then
                        Set rs = Nothing
                        strSQL = "Select JmlTerkecil From MasterBarang Where KdBarang = '" & .TextMatrix(.Row, 10) & "'"
                        Call msubRecFO(rs, strSQL)
                        dblJmlTerkecil = IIf(rs.EOF, 1, rs(0).value)

                        dblTotalStokK = dblJmlTerkecil * .TextMatrix(.Row, 5)
                        If Val(txtIsi.Text) > Val(dblTotalStokK) Then
                            MsgBox "Jumlah lebih besar dari stock (" & .TextMatrix(.Row, 5) & ")", vbExclamation, "Validasi"
                            txtIsi.SelStart = 0: txtIsi.SelLength = Len(txtIsi.Text)
                            Exit Sub
                        End If
                    End If

                    'add for FIFO validasi jika terjadi edit jml stok, hapus otomatis
                    If bolStatusFIFO = True Then
                        If Trim(.TextMatrix(.Row, 4)) <> "" Then
                            i = .Rows - 1
                            strKdBrg = .TextMatrix(.Row, 10)
                            strKdAsal = .TextMatrix(.Row, 11)
                            Do While i <> 1
                                If .TextMatrix(i, 10) <> "" Then
                                    If (strKdBrg = .TextMatrix(i, 10)) And (strKdAsal = .TextMatrix(i, 11)) Then
                                        .Row = i
                                        If .CellBackColor = vbRed Then
                                            .RemoveItem (.Row)
                                            .Row = i - 1
                                        End If
                                    End If
                                End If
                                i = i - 1
                            Loop

                            For i = 1 To .Rows - 1
                                If (strKdBrg = .TextMatrix(i, 10)) And (strKdAsal = .TextMatrix(i, 11)) Then
                                    .Row = i
                                    Exit For
                                End If
                            Next i
                        End If

                        .SetFocus
                        intRowTemp = 0
                    End If

                    'add for FIFO jika jml yg diinput melebihi stok penerimaan, mk otomatis muncul di row selanjutnya
                    If bolStatusFIFO = True Then
                        Set dbRst = Nothing
                        Call msubRecFO(dbRst, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & .TextMatrix(.Row, 10) & "','" & .TextMatrix(.Row, 11) & "','" & .TextMatrix(.Row, 9) & "') as stok")
                        If .TextMatrix(.Row, 3) = "S" Then
                            dblSelisih = dbRst(0) - CDbl(txtIsi.Text)
                        Else
                            dblSelisih = (dbRst(0) * dblJmlTerkecil) - CDbl(txtIsi.Text)
                        End If
                        If dblSelisih < 0 Then
                            If .TextMatrix(.Row, 3) = "S" Then
                                txtIsi.Text = dbRst(0)
                            Else
                                txtIsi.Text = dbRst(0) * dblJmlTerkecil
                            End If
                        Else
                            Set dbRst = Nothing
                            strSQL = "Select JmlStok as Stok From StokRuangan Where KdBarang='" & .TextMatrix(.Row, 10) & "' and KdAsal='" & .TextMatrix(.Row, 11) & "' and KdRuangan='" & mstrKdRuangan & "'"
                            Call msubRecFO(dbRst, strSQL)
                            If dbRst.EOF Then
                                .TextMatrix(.Row, 5) = 0
                            Else
                                .TextMatrix(.Row, 5) = IIf(IsNull(dbRst("Stok")), 0, dbRst("Stok"))
                            End If
                        End If
                    End If
                    'end FIFO

                    .TextMatrix(.Row, .Col) = txtIsi.Text
                    .TextMatrix(.Row, 7) = CCur(.TextMatrix(.Row, 6)) * CDbl(.TextMatrix(.Row, 4))

                    'add for FIFO jika jml yg diinput melebihi stok penerimaan, mk otomatis muncul di row selanjutnya
                    If bolStatusFIFO = True Then
                        If dblSelisih < 0 Then
                            With fgAlkes
                                strSQL = "select NoTerima As NoFIFO,JmlStokMax from V_StokRuanganFIFO where KdBarang='" & .TextMatrix(.Row, 10) & "' and KdAsal='" & .TextMatrix(.Row, 11) & "' and NoTerima<>'" & .TextMatrix(.Row, 9) & "' and JmlStok<>0 order by TglTerima asc"
                                Set dbRst = Nothing
                                Call msubRecFO(dbRst, strSQL)
                                If dbRst.EOF = False Then
                                    dbRst.MoveFirst
                                    For i = 1 To dbRst.RecordCount
                                        .Rows = .Rows + 1

                                        intRowTemp = .Row
                                        If .TextMatrix(.Rows - 2, 10) = "" Then
                                            .Row = .Rows - 2
                                        Else
                                            .Row = .Rows - 1
                                        End If
                                        For j = 0 To .Cols - 1
                                            .Col = j
                                            .CellBackColor = vbRed
                                            .CellForeColor = vbWhite
                                        Next j

                                        .Row = intRowTemp
                                        intRowTemp = 0
                                        If .TextMatrix(.Rows - 2, 2) = "" Then
                                            intRowTemp = .Rows - 2
                                        Else
                                            intRowTemp = .Rows - 1
                                        End If

                                        .TextMatrix(intRowTemp, 0) = .TextMatrix(.Row, 0)
                                        .TextMatrix(intRowTemp, 1) = .TextMatrix(.Row, 1)
                                        .TextMatrix(intRowTemp, 2) = .TextMatrix(.Row, 2)
                                        .TextMatrix(intRowTemp, 3) = .TextMatrix(.Row, 3)
                                        .TextMatrix(intRowTemp, 10) = .TextMatrix(.Row, 10)
                                        .TextMatrix(intRowTemp, 11) = .TextMatrix(.Row, 11)

                                        strNoTerima = dbRst("NoFIFO")
                                        .TextMatrix(intRowTemp, 9) = strNoTerima

                                        strSQL = ""
                                        Set rsb = Nothing
                                        strSQL = "SELECT dbo.FB_TakeHargaNettoOA('2222222222','01','" & .TextMatrix(intRowTemp, 10) & "','" & .TextMatrix(intRowTemp, 11) & "','" & .TextMatrix(intRowTemp, 3) & "', '" & mstrKdRuangan & "','" & .TextMatrix(intRowTemp, 9) & "') AS HargaBarang"
                                        Call msubRecFO(rsb, strSQL)
                                        If rsb.EOF = True Then curHargaBrg = 0 Else curHargaBrg = rsb(0).value

                                        .TextMatrix(intRowTemp, 6) = curHargaBrg
                                        .TextMatrix(intRowTemp, 8) = curHargaBrg

                                        .TextMatrix(intRowTemp, 4) = Abs(dblSelisih)

                                        If .TextMatrix(intRowTemp, 3) = "S" Then
                                            dblSelisih = Abs(dblSelisih) - CDbl(dbRst("JmlStokMax"))
                                        Else
                                            dblSelisih = Abs(dblSelisih) - CDbl(dbRst("JmlStokMax") * dblJmlTerkecil)
                                        End If
                                        If dblSelisih >= 0 Then
                                            If .TextMatrix(intRowTemp, 3) = "S" Then
                                                .TextMatrix(intRowTemp, 4) = dbRst("JmlStokMax")
                                            Else
                                                .TextMatrix(intRowTemp, 4) = dbRst("JmlStokMax") * dblJmlTerkecil
                                            End If
                                        End If

                                        .TextMatrix(intRowTemp, 7) = CCur(.TextMatrix(intRowTemp, 6)) * CDbl(.TextMatrix(intRowTemp, 4))

                                        If dblSelisih <= 0 Then Exit For
                                        dbRst.MoveNext
                                    Next i
                                End If
                            End With
                        End If
                    End If
                    'end fifo

                    .SetFocus
                    .Col = 5

            End Select
        End With

        txtIsi.Visible = False

        If fgAlkes.RowPos(fgAlkes.Row) >= fgAlkes.Height - 360 Then
            fgAlkes.SetFocus
            SendKeys "{DOWN}"
            Exit Sub
        End If

    ElseIf KeyAscii = 27 Then
        txtIsi.Visible = False
        dgObatAlkes.Visible = False
        fgAlkes.SetFocus
    End If

    If fgAlkes.Col = 4 Then
        If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",")) Then KeyAscii = 0
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    strSQL = "select  kdpelayananrs,namapelayanan from V_ListPemakaianBahan  " '"
    Call msubDcSource(dcNamaPelayanan, dbRst, strSQL)

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()
    With fgAlkes
        .clear
        .Rows = 2
        .Cols = 14
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Nama Barang"
        .TextMatrix(0, 2) = "Asal Barang"
        .TextMatrix(0, 3) = "Satuan"
        .TextMatrix(0, 4) = "Jumlah"
        .TextMatrix(0, 5) = "Stok"
        .TextMatrix(0, 6) = "Harga Satuan"
        .TextMatrix(0, 7) = "Total Harga"
        .TextMatrix(0, 8) = "Harga Beli"
        .TextMatrix(0, 9) = "NoTerima"
        .TextMatrix(0, 10) = "KdBarang"
        .TextMatrix(0, 11) = "KdAsal"
        .TextMatrix(0, 12) = "KdPelayanRS"
        .TextMatrix(0, 13) = "KdRuangan"

        .ColWidth(0) = 0
        .ColWidth(1) = 3700
        .ColWidth(2) = 1800
        .ColWidth(3) = 1000
        .ColWidth(4) = 800
        .ColWidth(5) = 800
        .ColWidth(6) = 1800
        .ColWidth(7) = 2000
        .ColWidth(8) = 1200
        .ColWidth(9) = 0
        .ColWidth(10) = 0 'KdBarang
        .ColWidth(11) = 0 'KdAsal
        .ColWidth(12) = 0 'KdPelayanRS
        .ColWidth(13) = 0 'KdRuangan

    End With
End Sub

Private Sub txtKeperluan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglPeriksa.SetFocus
End Sub

