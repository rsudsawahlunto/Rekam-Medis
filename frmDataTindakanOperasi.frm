VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmDataTindakanOperasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Tindakan Operasi"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDataTindakanOperasi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   8880
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      TabIndex        =   18
      Top             =   960
      Width           =   8895
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   7080
         Width           =   1575
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   5640
         TabIndex        =   8
         Top             =   7080
         Width           =   1575
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   7080
         Width           =   1575
      End
      Begin VB.CommandButton cmTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   7200
         TabIndex        =   10
         Top             =   7080
         Width           =   1575
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6735
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   11880
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Jenis Operasi"
         TabPicture(0)   =   "frmDataTindakanOperasi.frx":0CCA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "List Tindakan Operasi"
         TabPicture(1)   =   "frmDataTindakanOperasi.frx":0CE6
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   6015
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   8415
            Begin VB.CheckBox CheckStatusEnbl1 
               Alignment       =   1  'Right Justify
               Caption         =   "Status Aktif"
               Height          =   255
               Left            =   6840
               TabIndex        =   15
               Top             =   1440
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.TextBox txtKodeExternal1 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2280
               MaxLength       =   5
               TabIndex        =   14
               Top             =   1440
               Width           =   1815
            End
            Begin VB.TextBox txtNamaExternal1 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2280
               MaxLength       =   5
               TabIndex        =   16
               Top             =   1800
               Width           =   5895
            End
            Begin VB.TextBox txtKdOps 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   330
               Left            =   2280
               MaxLength       =   5
               TabIndex        =   11
               Top             =   360
               Width           =   1215
            End
            Begin MSDataListLib.DataCombo dcJenisOperasi 
               Height          =   330
               Left            =   2280
               TabIndex        =   12
               Top             =   720
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
            End
            Begin VB.TextBox txtTindakan 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2280
               MaxLength       =   50
               TabIndex        =   13
               Top             =   1080
               Width           =   5895
            End
            Begin MSDataGridLib.DataGrid dgTindakanOperasi 
               Height          =   3615
               Left            =   240
               TabIndex        =   17
               Top             =   2280
               Width           =   7935
               _ExtentX        =   13996
               _ExtentY        =   6376
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
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Kode External"
               Height          =   210
               Left            =   240
               TabIndex        =   30
               Top             =   1440
               Width           =   1140
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Nama External"
               Height          =   210
               Left            =   240
               TabIndex        =   29
               Top             =   1800
               Width           =   1170
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Kode "
               Height          =   210
               Left            =   240
               TabIndex        =   25
               Top             =   360
               Width           =   480
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Nama Tindakan Operasi"
               Height          =   210
               Left            =   240
               TabIndex        =   24
               Top             =   1080
               Width           =   1905
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Jenis Operasi"
               Height          =   210
               Left            =   240
               TabIndex        =   23
               Top             =   720
               Width           =   1050
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   6015
            Left            =   -74880
            TabIndex        =   20
            Top             =   600
            Width           =   8415
            Begin VB.CheckBox CheckStatusEnbl 
               Alignment       =   1  'Right Justify
               Caption         =   "Status Aktif"
               Height          =   255
               Left            =   6840
               TabIndex        =   4
               Top             =   1560
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.TextBox txtKodeExternal 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   5
               TabIndex        =   3
               Top             =   1560
               Width           =   1815
            End
            Begin VB.TextBox txtNamaExternal 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   5
               TabIndex        =   5
               Top             =   1920
               Width           =   6255
            End
            Begin VB.TextBox txtSingkatan 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   5
               TabIndex        =   2
               Top             =   1200
               Width           =   1695
            End
            Begin VB.TextBox txtKodeJenis 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   330
               Left            =   1920
               MaxLength       =   2
               TabIndex        =   0
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox txtJenisOperasi 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   50
               TabIndex        =   1
               Top             =   840
               Width           =   6255
            End
            Begin MSDataGridLib.DataGrid dgJenisOperasi 
               Height          =   3495
               Left            =   240
               TabIndex        =   6
               Top             =   2400
               Width           =   7935
               _ExtentX        =   13996
               _ExtentY        =   6165
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
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Kode External"
               Height          =   210
               Left            =   240
               TabIndex        =   32
               Top             =   1560
               Width           =   1140
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Nama External"
               Height          =   210
               Left            =   240
               TabIndex        =   31
               Top             =   1920
               Width           =   1170
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Singkatan"
               Height          =   210
               Left            =   240
               TabIndex        =   27
               Top             =   1200
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Kode "
               Height          =   210
               Left            =   240
               TabIndex        =   26
               Top             =   480
               Width           =   480
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nama Jenis Operasi"
               Height          =   210
               Left            =   240
               TabIndex        =   21
               Top             =   840
               Width           =   1560
            End
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   28
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
      Left            =   7080
      Picture         =   "frmDataTindakanOperasi.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDataTindakanOperasi.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataTindakanOperasi.frx":444B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmDataTindakanOperasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.recordset
Dim rst As ADODB.recordset
Dim s As String, kdJnsOP1 As String * 2, kdJnsOP2 As String * 2
Dim KdTindakan As String * 5

Private Sub cmdBatal_Click()
    On Error Resume Next
    Select Case SSTab1.Tab
        Case 0
            txtKodeJenis.Text = ""
            txtJenisOperasi.Text = ""
            txtSingkatan.Text = ""
            txtKodeExternal.Text = ""
            txtNamaExternal.Text = ""
            CheckStatusEnbl.value = 1
            txtJenisOperasi.SetFocus
        Case 1
            dcJenisOperasi.BoundText = ""
            txtKdOps.Text = ""
            txtTindakan.Text = ""
            txtKodeExternal1.Text = ""
            txtNamaExternal1.Text = ""
            CheckStatusEnbl1.value = 1
            txtTindakan.SetFocus
    End Select
    Call ListData
    Call subLoadDcSource
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errLoad
    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtJenisOperasi, "Nama Jenis operasi Kosong") = False Then Exit Sub
            If sp_JenisOperasi("D") = False Then Exit Sub

        Case 1
            If Periksa("datacombo", dcJenisOperasi, "Nama Jenis operasi Kosong") = False Then Exit Sub
            If Periksa("text", txtTindakan, "Nama Tindakan operasi Kosong") = False Then Exit Sub
            If sp_ListTindakanOperasi("D") = False Then Exit Sub

    End Select
    MsgBox "Data berhasil dihapus", vbInformation, "Informasi"
    Call cmdBatal_Click
    Call ListData
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo xxx
    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtJenisOperasi, "Nama Jenis operasi Kosong") = False Then Exit Sub
            If sp_JenisOperasi("A") = False Then Exit Sub

        Case 1
            If Periksa("datacombo", dcJenisOperasi, "Nama Jenis operasi Kosong") = False Then Exit Sub
            If Periksa("text", txtTindakan, "Nama Tindakan operasi Kosong") = False Then Exit Sub
            If sp_ListTindakanOperasi("A") = False Then Exit Sub

    End Select
    MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
    Call cmdBatal_Click
    Call ListData
    Exit Sub
xxx:
    Call msubPesanError
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub dcJenisOperasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcJenisOperasi.MatchedWithList = True Then txtTindakan.SetFocus
        strSQL = "SELECT KdJenisOperasi, JenisOperasi FROM JenisOperasi where StatusEnabled='1'  and (JenisOperasi LIKE '%" & dcJenisOperasi.Text & "%')order by JenisOperasi"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcJenisOperasi.Text = ""
            Exit Sub
        End If
        dcJenisOperasi.BoundText = rs(0).value
        dcJenisOperasi.Text = rs(1).value
    End If
End Sub

Private Sub dgJenisOperasi_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgJenisOperasi
    WheelHook.WheelHook dgJenisOperasi
End Sub

Private Sub dgJenisOperasi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKodeJenis.Text = dgJenisOperasi.Columns(0).value
    txtJenisOperasi.Text = dgJenisOperasi.Columns(1).value
    txtSingkatan.Text = dgJenisOperasi.Columns(2).value
    txtKodeExternal.Text = dgJenisOperasi.Columns(4).value
    txtNamaExternal.Text = dgJenisOperasi.Columns(5).value
    If dgJenisOperasi.Columns(6) = "" Then
        CheckStatusEnbl.value = 0
    ElseIf dgJenisOperasi.Columns(6) = 0 Then
        CheckStatusEnbl.value = 0
    ElseIf dgJenisOperasi.Columns(6) = 1 Then
        CheckStatusEnbl.value = 1
    End If
End Sub

Private Sub dgTindakanOperasi_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgTindakanOperasi
    WheelHook.WheelHook dgTindakanOperasi
End Sub

Private Sub dgTindakanOperasi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdOps.Text = dgTindakanOperasi.Columns(0).value
    txtTindakan.Text = dgTindakanOperasi.Columns(1).value
    dcJenisOperasi.BoundText = dgTindakanOperasi.Columns(2).value
    txtKodeExternal1.Text = dgTindakanOperasi.Columns(4).value
    txtNamaExternal1.Text = dgTindakanOperasi.Columns(5).value
    If dgTindakanOperasi.Columns(6) = "" Then
        CheckStatusEnbl1.value = 0
    ElseIf dgTindakanOperasi.Columns(6) = 0 Then
        CheckStatusEnbl1.value = 0
    ElseIf dgTindakanOperasi.Columns(6) = 1 Then
        CheckStatusEnbl1.value = 1
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    Select Case KeyCode
        Case vbKey0
            If strCtrlKey = 4 Then SSTab1.SetFocus: SSTab1.Tab = 0
        Case vbKey1
            If strCtrlKey = 4 Then SSTab1.SetFocus: SSTab1.Tab = 1
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call ListData
    Call subLoadDcSource
    Call cmdBatal_Click

    SSTab1.Tab = 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call cmdBatal_Click
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        Select Case SSTab1.Tab
            Case 0
                txtJenisOperasi.SetFocus
            Case 1
                dcJenisOperasi.SetFocus
        End Select
    End If
End Sub

Sub ListData()
    On Error GoTo errLoad
    Select Case SSTab1.Tab
        Case 0
            Set rs = Nothing
            strSQL = "select * from JenisOperasi"
            Set rs = New ADODB.recordset
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgJenisOperasi.DataSource = rs
            dgJenisOperasi.Columns(0).Caption = "Kode Jenis Operasi"
            dgJenisOperasi.Columns(0).Width = 1500
            dgJenisOperasi.Columns(1).Caption = "Nama Jenis Operasi"
            dgJenisOperasi.Columns(1).Width = 5000
            dgJenisOperasi.Columns(2).Width = 1000
            Set rs = Nothing

        Case 1
            Set rs = Nothing
            strSQL = "select * from V_ListTindakanOperasi"
            Set rst = New ADODB.recordset
            rst.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgTindakanOperasi.DataSource = rst
            dgTindakanOperasi.Columns(0).Caption = "Kode Jenis Operasi"
            dgTindakanOperasi.Columns(0).Width = 1000
            dgTindakanOperasi.Columns(1).Caption = "Nama Tindakan Operasi"
            dgTindakanOperasi.Columns(1).Width = 4500
            dgTindakanOperasi.Columns(2).Caption = "Kode Jenis Operasi"
            dgTindakanOperasi.Columns(2).Width = 1000
            dgTindakanOperasi.Columns(3).Caption = "Jenis Operasi"
            dgTindakanOperasi.Columns(3).Width = 4500
            dgTindakanOperasi.Columns(2).Width = 0
            Set rs = Nothing
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function sp_JenisOperasi(f_Status As String) As Boolean
    On Error GoTo errSp
    sp_JenisOperasi = False
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdJenisOperasi", adChar, adParamInput, 2, txtKodeJenis.Text)
        .Parameters.Append .CreateParameter("JenisOperasi", adVarChar, adParamInput, 50, Trim(txtJenisOperasi.Text))
        .Parameters.Append .CreateParameter("Singkatan", adVarChar, adParamInput, 5, Trim(txtSingkatan.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKodeExternal.Text))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNamaExternal.Text))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_JenisOperasi"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan data", vbCritical, "Validasi"
            Call deleteADOCommandParameters(dbcmd)
            Set dbcmd = Nothing
            sp_JenisOperasi = False
        Else
            sp_JenisOperasi = True
            Call Add_HistoryLoginActivity("AUD_JenisOperasi")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errSp:
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    msubPesanError
End Function

Private Function sp_ListTindakanOperasi(f_Status As String) As Boolean
    On Error GoTo errSp
    sp_ListTindakanOperasi = False
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdTindakanOperasi", adChar, adParamInput, 5, txtKdOps.Text)
        .Parameters.Append .CreateParameter("NamaTindakanOperasi", adVarChar, adParamInput, 50, Trim(txtTindakan.Text))
        .Parameters.Append .CreateParameter("KdJenisOperasi", adChar, adParamInput, 2, dcJenisOperasi.BoundText)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKodeExternal1.Text))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNamaExternal1.Text))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl1.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_ListTindakanOperasi"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan data", vbCritical, "Validasi"
            Call deleteADOCommandParameters(dbcmd)
            Set dbcmd = Nothing
            sp_ListTindakanOperasi = False
        Else
            sp_ListTindakanOperasi = True
            Call Add_HistoryLoginActivity("AUD_ListTindakanOperasi")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errSp:
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    msubPesanError
End Function

Private Sub subLoadDcSource()
    On Error GoTo errLoad
    Call msubDcSource(dcJenisOperasi, rs, "SELECT KdJenisOperasi, JenisOperasi FROM JenisOperasi where StatusEnabled='1' order by JenisOperasi")
    Exit Sub
errLoad:
    Call msubPesanError("subLoadDcSource")
End Sub

Private Sub txtJenisOperasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtSingkatan.SetFocus
End Sub

Private Sub txtSingkatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal.SetFocus
End Sub

Private Sub txtTindakan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal1.SetFocus
End Sub

Private Sub txtKodeExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal.SetFocus
End Sub

Private Sub CheckStatusEnbl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl.SetFocus
End Sub

Private Sub txtKodeExternal1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal1.SetFocus
End Sub

Private Sub CheckStatusEnbl1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaExternal1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl1.SetFocus
End Sub
