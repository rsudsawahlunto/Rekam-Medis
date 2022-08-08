VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmImunisasiJenisKontrasepsi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Imunisasi & Jenis Kontrasepsi"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImunisasiJenisKontrasepsi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   8670
   Begin VB.CommandButton cmTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   7005
      TabIndex        =   3
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   7680
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Imunisasi"
      TabPicture(0)   =   "frmImunisasiJenisKontrasepsi.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Jenis Kontrasepsi"
      TabPicture(1)   =   "frmImunisasiJenisKontrasepsi.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   5895
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   8415
         Begin VB.TextBox txtKodeExternal 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   24
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txtNamaExternal 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   23
            Top             =   1440
            Width           =   6615
         End
         Begin VB.CheckBox CheckStatusEnbl 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   6840
            TabIndex        =   22
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtCariImunisasi 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   18
            Top             =   5400
            Width           =   2355
         End
         Begin VB.TextBox txtKodeImunisasi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   240
            MaxLength       =   3
            TabIndex        =   12
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtNamaImunisasi 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   8
            Top             =   480
            Width           =   6975
         End
         Begin MSDataGridLib.DataGrid dgImunisasi 
            Height          =   3375
            Left            =   240
            TabIndex        =   9
            Top             =   1920
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   5953
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   15
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
                  LCID            =   1033
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
                  LCID            =   1033
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
         Begin VB.Label Label12 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cari Data"
            Height          =   210
            Index           =   4
            Left            =   240
            TabIndex        =   17
            Top             =   5400
            Width           =   720
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
            Height          =   210
            Index           =   1
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Nama Imunisasi"
            Height          =   210
            Index           =   0
            Left            =   1200
            TabIndex        =   10
            Top             =   240
            Width           =   1230
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5775
         Left            =   -74880
         TabIndex        =   5
         Top             =   600
         Width           =   8415
         Begin VB.TextBox txtKodeExternal1 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   29
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txtNamaExternal1 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   28
            Top             =   1440
            Width           =   6615
         End
         Begin VB.CheckBox CheckStatusEnbl1 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   6840
            TabIndex        =   27
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtCariJenisKontrasepsi 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   20
            Top             =   5280
            Width           =   2355
         End
         Begin VB.TextBox txtKodeJenisKontrasepsi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   240
            MaxLength       =   2
            TabIndex        =   16
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtJenisKontrasepsi 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   13
            Top             =   480
            Width           =   6975
         End
         Begin MSDataGridLib.DataGrid dgKontrasepsi 
            Height          =   3255
            Left            =   240
            TabIndex        =   6
            Top             =   1920
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   5741
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   15
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
                  LCID            =   1033
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
                  LCID            =   1033
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
         Begin VB.Label Label2 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cari Data"
            Height          =   210
            Index           =   5
            Left            =   240
            TabIndex        =   19
            Top             =   5280
            Width           =   720
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
            Height          =   210
            Index           =   3
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Kontrasepsi"
            Height          =   210
            Index           =   2
            Left            =   1200
            TabIndex        =   14
            Top             =   240
            Width           =   1380
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   21
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
      Left            =   6840
      Picture         =   "frmImunisasiJenisKontrasepsi.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmImunisasiJenisKontrasepsi.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmImunisasiJenisKontrasepsi.frx":444B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmImunisasiJenisKontrasepsi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function sp_Imunisasi(f_Status As String) As Boolean
    On Error GoTo errLoad
    sp_Imunisasi = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdImunisasi", adChar, adParamInput, 3, txtKodeImunisasi.Text)
        .Parameters.Append .CreateParameter("NamaImunisasi", adVarChar, adParamInput, 50, Trim(txtNamaImunisasi.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKodeExternal.Text))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNamaExternal.Text))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_Imunisasi"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_Imunisasi = False
        Else
            Call Add_HistoryLoginActivity("AUD_Imunisasi")
        End If
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    sp_Imunisasi = False
    Call msubPesanError
End Function

Private Function sp_JenisKontrasepsi(f_Status As String) As Boolean
    On Error GoTo errLoad
    sp_JenisKontrasepsi = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdJenisKontrasepsi", adChar, adParamInput, 2, txtKodeJenisKontrasepsi.Text)
        .Parameters.Append .CreateParameter("JenisKontrasepsi", adVarChar, adParamInput, 50, Trim(txtJenisKontrasepsi.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKodeExternal1.Text))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNamaExternal1.Text))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl1.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_JenisKontrasepsi"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_JenisKontrasepsi = False
        Else
            Call Add_HistoryLoginActivity("AUD_JenisKontrasepsi")
        End If
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    sp_JenisKontrasepsi = False
    Call msubPesanError
End Function

Private Sub subKosong()
    Select Case SSTab1.Tab
        Case 0
            txtKodeImunisasi.Text = ""
            txtNamaImunisasi.Text = ""
            txtCariImunisasi.Text = ""
            txtKodeExternal.Text = ""
            txtNamaExternal.Text = ""
            CheckStatusEnbl.value = 1
        Case 1
            txtKodeJenisKontrasepsi.Text = ""
            txtJenisKontrasepsi.Text = ""
            txtCariJenisKontrasepsi.Text = ""
            txtKodeExternal1.Text = ""
            txtNamaExternal1.Text = ""
            CheckStatusEnbl1.value = 1
    End Select
End Sub

Private Sub subLoadDataGrid()
    On Error GoTo errLoad
    Select Case SSTab1.Tab
        Case 0
            Call msubRecFO(rs, "SELECT * FROM Imunisasi WHERE NamaImunisasi LIKE '%" & txtCariImunisasi.Text & "%'")
            Set dgImunisasi.DataSource = rs
            With dgImunisasi
                .Columns(0).Width = 1500
                .Columns(1).Width = 4000
                .Columns(0).Caption = "Kode"
                .Columns(1).Caption = "Nama Imunisasi"
            End With
        Case 1
            Call msubRecFO(rs, "SELECT * FROM JenisKontrasepsi WHERE JenisKontrasepsi LIKE '%" & txtCariJenisKontrasepsi.Text & "%'")
            Set dgKontrasepsi.DataSource = rs
            With dgKontrasepsi
                .Columns(0).Width = 1500
                .Columns(1).Width = 4000
                .Columns(0).Caption = "Kode"
                .Columns(1).Caption = "Nama Kontrasepsi"
            End With
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdBatal_Click()
    Call subKosong
    Call subLoadDataGrid
    Select Case SSTab1.Tab
        Case 0
            txtNamaImunisasi.SetFocus
        Case 1
            txtJenisKontrasepsi.SetFocus
    End Select
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errLoad

    Select Case SSTab1.Tab
        Case 0
            If txtKodeImunisasi.Text = "" Then
                MsgBox "Pilih data yang akan dihapus", vbExclamation, "validasi"
                Exit Sub
            End If
            If MsgBox("Apakah anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
            If sp_Imunisasi("D") = False Then Exit Sub
            MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
            Call cmdBatal_Click
        Case 1
            If txtKodeJenisKontrasepsi.Text = "" Then
                MsgBox "Pilih data yang akan dihapus", vbExclamation, "validasi"
                Exit Sub
            End If
            If MsgBox("Apakah anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
            If sp_JenisKontrasepsi("D") = False Then Exit Sub
            MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
            Call cmdBatal_Click
    End Select

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtNamaImunisasi, "Nama imunisasi kosong") = False Then Exit Sub
            If sp_Imunisasi("A") = False Then Exit Sub
            MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
            Call cmdBatal_Click
        Case 1
            If Periksa("text", txtJenisKontrasepsi, "Nama jenis kontrasepsi kosong") = False Then Exit Sub
            If sp_JenisKontrasepsi("A") = False Then Exit Sub
            MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
            Call cmdBatal_Click
    End Select

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub dgImunisasi_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgImunisasi
    WheelHook.WheelHook dgImunisasi
End Sub

Private Sub dgImunisasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaImunisasi.SetFocus
End Sub

Private Sub dgImunisasi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errLoad

    txtKodeImunisasi.Text = dgImunisasi.Columns(0)
    txtNamaImunisasi.Text = dgImunisasi.Columns(1)
    txtKodeExternal.Text = dgImunisasi.Columns("KodeExternal").value
    txtNamaExternal.Text = dgImunisasi.Columns("NamaExternal").value
    If dgImunisasi.Columns("StatusEnabled") = "" Then
        CheckStatusEnbl.value = 0
    ElseIf dgImunisasi.Columns("StatusEnabled") = 0 Then
        CheckStatusEnbl.value = 0
    ElseIf dgImunisasi.Columns("StatusEnabled") = 1 Then
        CheckStatusEnbl.value = 1
    End If

    Exit Sub
errLoad:
End Sub

Private Sub dgKontrasepsi_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKontrasepsi
    WheelHook.WheelHook dgKontrasepsi
End Sub

Private Sub dgKontrasepsi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJenisKontrasepsi.SetFocus
End Sub

Private Sub dgKontrasepsi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errLoad

    txtKodeJenisKontrasepsi.Text = dgKontrasepsi.Columns(0)
    txtJenisKontrasepsi.Text = dgKontrasepsi.Columns(1)
    txtKodeExternal1.Text = dgKontrasepsi.Columns("KodeExternal").value
    txtNamaExternal1.Text = dgKontrasepsi.Columns("NamaExternal").value
    If dgKontrasepsi.Columns("StatusEnabled") = "" Then
        CheckStatusEnbl1.value = 0
    ElseIf dgKontrasepsi.Columns("StatusEnabled") = 0 Then
        CheckStatusEnbl1.value = 0
    ElseIf dgKontrasepsi.Columns("StatusEnabled") = 1 Then
        CheckStatusEnbl1.value = 1
    End If
    Exit Sub
errLoad:
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    SSTab1.Tab = 0
    Call cmdBatal_Click

    Exit Sub
errLoad:
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call cmdBatal_Click
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case SSTab1.Tab
            Case 0
                txtNamaImunisasi.SetFocus
            Case 1
                txtJenisKontrasepsi.SetFocus
        End Select
    End If
End Sub

Private Sub txtCariImunisasi_Change()
    Call subLoadDataGrid
End Sub

Private Sub txtCariJenisKontrasepsi_Change()
    Call subLoadDataGrid
End Sub

Private Sub txtJenisKontrasepsi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then txtKodeExternal1.SetFocus
End Sub

Private Sub txtJenisKontrasepsi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal1.SetFocus
End Sub

Private Sub txtNamaImunisasi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then txtKodeExternal.SetFocus
End Sub

Private Sub txtNamaImunisasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal.SetFocus
End Sub

Private Sub txtKodeExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl.SetFocus
End Sub

Private Sub CheckStatusEnbl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal.SetFocus
End Sub

Private Sub txtNamaExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKodeExternal1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl1.SetFocus
End Sub

Private Sub CheckStatusEnbl1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal1.SetFocus
End Sub

Private Sub txtNamaExternal1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub
