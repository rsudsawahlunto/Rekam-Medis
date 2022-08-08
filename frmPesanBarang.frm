VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPesanBarang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pemesanan Barang"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPesanBarang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   7215
   Begin VB.Frame frameDataBrg 
      Caption         =   "Data Barang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   6735
      Begin MSDataGridLib.DataGrid dgBarang 
         Height          =   1935
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Height          =   670
      Left            =   0
      TabIndex        =   19
      Top             =   5760
      Width           =   7215
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   330
         Left            =   5520
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   330
         Left            =   3960
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   0
      TabIndex        =   16
      Top             =   2040
      Width           =   7215
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4920
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtJmlPesan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3960
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtNamaBrg 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   3615
      End
      Begin MSFlexGridLib.MSFlexGrid GridBarang 
         Height          =   2535
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   50
         Cols            =   6
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   8577768
         ForeColorFixed  =   -2147483627
         ForeColorSel    =   -2147483628
         BackColorBkg    =   16777215
         FocusRect       =   0
         HighLight       =   2
         FillStyle       =   1
         GridLines       =   3
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jml. Pesan"
         Height          =   210
         Left            =   3960
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nama Barang"
         Height          =   210
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Pemesanan"
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
      Left            =   0
      TabIndex        =   11
      Top             =   960
      Width           =   7215
      Begin MSDataListLib.DataCombo dcRuanganTujuan 
         Height          =   330
         Left            =   4200
         TabIndex        =   2
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox txtNoPesan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpTglPesan 
         Height          =   330
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
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
         Format          =   116064259
         UpDown          =   -1  'True
         CurrentDate     =   38070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. Pemesanan"
         Height          =   210
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pemesanan"
         Height          =   210
         Left            =   1920
         TabIndex        =   13
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tujuan Pemesanan"
         Height          =   210
         Left            =   4200
         TabIndex        =   12
         Top             =   360
         Width           =   1560
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   20
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   5400
      Picture         =   "frmPesanBarang.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPesanBarang.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPesanBarang.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmPesanBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFilter As String
Dim strkdbarang As String
Dim intJmlBarang As Integer

Private Sub cmdHapus_Click()
    With GridBarang
        If .Row = .Rows Then Exit Sub
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        .RemoveItem .Row
        intRowNow = .Rows - 1
    End With
End Sub

Private Sub cmdSimpan_Click()
    Dim adoCommand As New ADODB.Command
    Dim adoComm As New ADODB.Command
    For i = 1 To GridBarang.Rows - 2
        With adoCommand
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
            .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
            .Parameters.Append .CreateParameter("TglOrder", adDate, adParamInput, , Format(dtpTglPesan.value, "yyyy/MM/dd HH:mm:ss"))
            .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strIDPegawai)
            .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, dcRuanganTujuan.BoundText)
            .Parameters.Append .CreateParameter("OutputNoOrder", adChar, adParamOutput, 10, Null)

            .ActiveConnection = dbConn
            .CommandText = "dbo.Add_StrukOrderRuangan"
            .CommandType = adCmdStoredProc
            .Execute
            If Not (.Parameters("RETURN_VALUE").value = 0) Then
                MsgBox "Ada Kesalahan dalam Penyimpanan Data", vbCritical, "Validasi"
            Else
                txtNoPesan = .Parameters("OutputNoOrder").value
            End If
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
        End With
    Next i

    For i = 1 To GridBarang.Rows - 2
        With adoComm
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
            .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoPesan)
            .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, GridBarang.TextMatrix(i, 0))
            .Parameters.Append .CreateParameter("JmlOrder", adInteger, adParamInput, , GridBarang.TextMatrix(i, 2))

            .ActiveConnection = dbConn
            .CommandText = "dbo.Add_DetailOrderRuangan"
            .CommandType = adCmdStoredProc
            .Execute
            If Not (.Parameters("RETURN_VALUE").value = 0) Then
                MsgBox "Ada Kesalahan dalam Penyimpanan Data", vbCritical, "Validasi"
            Else
            End If
            Call deleteADOCommandParameters(adoComm)
            Set adoComm = Nothing
        End With
    Next i

    Call Add_HistoryLoginActivity("Add_StrukOrderRuangan+Add_DetailOrderRuangan")
    MsgBox "Pemasukan Data Sukses", vbInformation, "Informasi"
    cmdSimpan.Enabled = False
End Sub

Private Sub cmdTambah_Click()
    If txtJmlPesan.Text = "" Then
        MsgBox "Jumlah Pesan Harus Diisi", vbInformation, "Informasi"
        txtJmlPesan.SetFocus
        Exit Sub
    End If
    Dim i As Integer
    With GridBarang
        If strkdbarang = "" Then Exit Sub
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = strkdbarang Then Exit Sub
        Next i
        intRowNow = .Rows - 1

        .TextMatrix(intRowNow, 0) = strkdbarang
        .TextMatrix(intRowNow, 1) = txtNamaBrg.Text
        .TextMatrix(intRowNow, 2) = CInt(txtJmlPesan.Text)
        .Rows = .Rows + 1
        .SetFocus
    End With
    txtJmlPesan = ""
    txtNamaBrg.Text = ""
    txtNamaBrg.SetFocus
    frameDataBrg.Visible = False
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcRuanganTujuan_Change()
    txtNamaBrg.Text = ""
End Sub

Private Sub dcRuanganTujuan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcRuanganTujuan.MatchedWithList = True Then txtNamaBrg.SetFocus
        strSQL = "select kdruangan, namaruangan from v_ruangantujuanorderbrg where (NamaRuangan LIKE '%" & dcRuanganTujuan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcRuanganTujuan.Text = ""
            Exit Sub
        End If
        dcRuanganTujuan.BoundText = rs(0).value
        dcRuanganTujuan.Text = rs(1).value
    End If
End Sub

Private Sub dgBarang_DblClick()
    Call dgBarang_KeyPress(13)
End Sub

Private Sub dgBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlBarang = 0 Then Exit Sub

        txtNamaBrg.Text = dgBarang.Columns(1).value
        strkdbarang = dgBarang.Columns(0).value
        If strkdbarang = "" Then
            MsgBox "Pilih dulu Nama Barang yg ingin di proses", vbCritical, "Validasi"
            txtNamaBrg.Text = ""
            dgBarang.SetFocus
            Exit Sub
        End If
        frameDataBrg.Visible = False
    End If
    If KeyAscii = 27 Then
        frameDataBrg.Visible = False
    End If
End Sub

Private Sub dtpTglPesan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        dcRuanganTujuan.SetFocus
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglPesan.value = Now
    Set rs = Nothing
    rs.Open "select * from v_ruangantujuanorderbrg", dbConn, adOpenDynamic, adLockOptimistic
    Set dcRuanganTujuan.RowSource = rs
    dcRuanganTujuan.ListField = rs.Fields(1).Name
    dcRuanganTujuan.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
    Call setgrid

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtJmlPesan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdTambah_Click
    End If
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtNamaBrg_Change()
    If dcRuanganTujuan.BoundText = "038" Then
        strFilter = "WHERE [nama barang] like '" & txtNamaBrg.Text & "%' and StatusBarang='0'"
    Else
        strFilter = "WHERE [nama barang] like '" & txtNamaBrg.Text & "%' and StatusBarang='1'"
    End If
    strkdbarang = ""
    frameDataBrg.Visible = True
    Call subLoadbarang
End Sub

Private Sub subLoadbarang()
    On Error Resume Next
    strSQL = "SELECT * FROM V_DataBarang " & strFilter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlBarang = rs.RecordCount
    Set dgBarang.DataSource = rs
    Call SetGridBarang
End Sub

Private Sub setgrid()
    With GridBarang
        .clear
        .Rows = 2
        .Cols = 3
        .TextMatrix(0, 0) = "Kode Barang"
        .TextMatrix(0, 1) = "Nama Barang"
        .TextMatrix(0, 2) = "Jumlah Pesan"
        .ColWidth(0) = 1500
        .ColWidth(1) = 3500
        .ColWidth(2) = 1400
    End With
End Sub

Private Sub SetGridBarang()
    With dgBarang
        .Columns(0).Width = 1000
        .Columns(1).Width = 2500
        .Columns(2).Width = 1200
        .Columns(3).Width = 950
        .Columns(4).Width = 0
        .Columns(5).Width = 0
        .Columns(6).Width = 0
    End With
End Sub

Private Sub txtNamaBrg_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 27 Then
        txtNamaBrg = ""
        frameDataBrg.Visible = False
    End If
    If KeyAscii = 13 Then
        dgBarang.SetFocus
    End If
    If KeyAscii = 39 Then KeyAscii = 0
hell:
End Sub
