VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmMasterDiagnosaKeperawatan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Diagnosa Keperawatan"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMasterDiagnosaKeperawatan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9870
   Begin VB.Frame fraDiagnosa 
      Caption         =   "Data Diagnosa"
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
      Left            =   0
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   9855
      Begin MSDataGridLib.DataGrid dgDiagnosa 
         Height          =   2775
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   4895
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
   End
   Begin VB.TextBox txtKdDiagnosaKeperawatan 
      Height          =   375
      Left            =   0
      MaxLength       =   10
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   15
      Top             =   6840
      Width           =   9855
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   6840
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   5400
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   8280
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "F1 - Cetak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   930
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   9855
      Begin VB.CheckBox CheckStatusEnbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         Height          =   255
         Left            =   8280
         TabIndex        =   6
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtNamaExternal 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   1320
         Width           =   6015
      End
      Begin VB.TextBox txtKodeExternal 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtKdDiagnosa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6120
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtDiagnosa 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2640
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   600
         Width           =   6975
      End
      Begin MSDataGridLib.DataGrid dgDiagnosaKeperawatan 
         Height          =   3855
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   6800
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
      Begin MSDataListLib.DataCombo dcKdDiagnosa 
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Nama External"
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Kode External"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Kode Diagnosa"
         Height          =   210
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Diagnosa Keperawatan"
         Height          =   210
         Left            =   2760
         TabIndex        =   13
         Top             =   360
         Width           =   1860
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   19
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
      Left            =   8040
      Picture         =   "frmMasterDiagnosaKeperawatan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterDiagnosaKeperawatan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterDiagnosaKeperawatan.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmMasterDiagnosaKeperawatan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilterDiagnosa As String
Dim intJmlDiagnosa As Integer
Dim mstrKdDiagnosa As String

Private Sub cmdBatal_Click()
    Call Clear
    Call subLoadGridSource
    'txtKdDiagnosa.SetFocus
    dcKdDiagnosa.SetFocus
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errLoad
    If txtKdDiagnosa.Text = "" Then
        MsgBox "Pilih dulu Diagnosa yang akan dihapus.", vbOKOnly, "Validasi"
        Exit Sub
    End If
    If MsgBox("Apakah anda yakin akan mengapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub

    If sp_DiagnosaKeperawatan("D") = False Then Exit Sub

    Call cmdBatal_Click
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    If Periksa("datacombo", dcKdDiagnosa, "Kode Diagnosa kosong") = False Then Exit Sub
    If Periksa("text", txtDiagnosa, "Diagnosa Keperawatan kosong") = False Then Exit Sub

    If sp_DiagnosaKeperawatan("A") = False Then Exit Sub

    Call cmdBatal_Click
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdCetak_Click()
    frmCetakMastDiagnosaKeperawatan.Show
End Sub

Private Sub dcKdDiagnosa_Click(Area As Integer)
Call subLoadAskep
End Sub

Private Sub dgDiagnosa_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDiagnosa
    WheelHook.WheelHook dgDiagnosa
End Sub

Private Sub dgDiagnosa_DblClick()
    Call dgDiagnosa_KeyPress(13)
End Sub

Private Sub dgDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDiagnosa = 0 Then Exit Sub
        txtKdDiagnosa.Text = dgDiagnosa.Columns(0).value
        mstrKdDiagnosa = dgDiagnosa.Columns(0).value
        If mstrKdDiagnosa = "" Then
            MsgBox "Pilih dulu Diagnosa-nya", vbCritical, "Validasi"
            txtKdDiagnosa.Text = ""
            dgDiagnosa.SetFocus
            Exit Sub
        End If
        fraDiagnosa.Visible = False
        txtDiagnosa.SetFocus
    End If
    If KeyAscii = 27 Then
        fraDiagnosa.Visible = False
    End If
End Sub

Private Sub dgDiagnosaKeperawatan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDiagnosaKeperawatan
    WheelHook.WheelHook dgDiagnosaKeperawatan
End Sub

Private Sub dgDiagnosaKeperawatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDiagnosa.SetFocus
    If KeyAscii = 27 Then fraDiagnosa.Visible = False
End Sub

Private Sub dgDiagnosaKeperawatan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgDiagnosaKeperawatan.ApproxCount = 0 Then Exit Sub
    txtKdDiagnosa.Text = dgDiagnosaKeperawatan.Columns(2).value
    dcKdDiagnosa.Text = dgDiagnosaKeperawatan.Columns(7).value 'Nama Askep
    
    txtDiagnosa.Text = dgDiagnosaKeperawatan.Columns(1).value
    txtKdDiagnosaKeperawatan.Text = dgDiagnosaKeperawatan.Columns(0).value
    txtKodeExternal.Text = dgDiagnosaKeperawatan.Columns(3).value
    txtNamaExternal.Text = dgDiagnosaKeperawatan.Columns(4).value
    If dgDiagnosaKeperawatan.Columns(5) = "" Then
        CheckStatusEnbl.value = 0
    ElseIf dgDiagnosaKeperawatan.Columns(5) = 0 Then
        CheckStatusEnbl.value = 0
    ElseIf dgDiagnosaKeperawatan.Columns(5) = 1 Then
        CheckStatusEnbl.value = 1
    End If
    fraDiagnosa.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            If dgDiagnosaKeperawatan.ApproxCount = 0 Then Exit Sub
            frmCetakMastDiagnosaKeperawatan.Show
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call openConnection
    Call Clear
    Call subLoadGridSource
    Call loadKdDiagnosa

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub loadKdDiagnosa()
    On Error GoTo errLoad

        strSQL = "select KdAskep, NamaAskep from AsuhanKeperawatan "
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        Set dcKdDiagnosa.RowSource = rs
        dcKdDiagnosa.ListField = rs(1).Name
        dcKdDiagnosa.BoundColumn = rs(0).Name

        dcKdDiagnosa.SetFocus
    Exit Sub
errLoad:
End Sub

Private Sub subLoadGridSource()
    On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "select * from DiagnosaKeperawatan inner join AsuhanKeperawatan ON DiagnosaKeperawatan.KdAskep = AsuhanKeperawatan.KdAskep"
    rs.Open strSQL, dbConn, adOpenDynamic, adLockOptimistic
    Set dgDiagnosaKeperawatan.DataSource = rs
    With dgDiagnosaKeperawatan
        .Columns(0).Caption = "Kd Diagnosa Keperawatan"
        .Columns(0).Width = 1600
        .Columns(1).Caption = "Diagnosa Keperawatan"
        .Columns(1).Width = 5900
        .Columns(2).Caption = "Kode Askep"
        .Columns(2).Width = 1200
    End With
    Set rs = Nothing

    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub Clear()
    txtKdDiagnosa.Text = ""
    txtDiagnosa.Text = ""
    txtKdDiagnosaKeperawatan.Text = ""
    txtKodeExternal.Text = ""
    txtNamaExternal.Text = ""
    CheckStatusEnbl.value = 1
    dcKdDiagnosa.Text = ""
End Sub

Private Sub txtDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal.SetFocus
End Sub

Private Sub txtDiagnosa_LostFocus()
    Dim i As Integer
    Dim tempText As String

    tempText = Trim(txtDiagnosa.Text)
    txtDiagnosa.Text = ""
    For i = 1 To Len(tempText)
        If Asc(Mid(tempText, i, 1)) <> 10 And Asc(Mid(tempText, i, 1)) <> 13 Then
            txtDiagnosa.Text = txtDiagnosa.Text & Mid(tempText, i, 1)
        End If
    Next i
End Sub

Private Sub txtKdDiagnosa_Change()
    On Error GoTo errLoad
'    If txtKdDiagnosa.Text = "" Then Exit Sub
    strFilterDiagnosa = "WHERE KdAskep like '%" & txtKdDiagnosa.Text & "%' or NamaAskep like '%" & txtKdDiagnosa.Text & "%'"
    fraDiagnosa.Visible = True
    Call subLoadAskep

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadAskep()
    On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "select KdAskep, NamaAskep from AsuhanKeperawatan " & strFilterDiagnosa
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDiagnosa = rs.RecordCount
    Set dgDiagnosa.DataSource = rs
    With dgDiagnosa
        .Columns(0).Caption = "Kode Askep"
        .Columns(0).Width = 1200
        .Columns(1).Caption = "Nama Askep"
        .Columns(1).Width = 7500
    End With
    fraDiagnosa.Left = 0
    fraDiagnosa.Top = 1920
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub subLoadDiagnosa()
    On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "select KdDiagnosa, NamaDiagnosa from Diagnosa " & strFilterDiagnosa
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDiagnosa = rs.RecordCount
    Set dgDiagnosa.DataSource = rs
    With dgDiagnosa
        .Columns(0).Caption = "Kode Diagnosa"
        .Columns(0).Width = 1200
        .Columns(1).Caption = "Nama Diagnosa"
        .Columns(1).Width = 7500
    End With
    fraDiagnosa.Left = 0
    fraDiagnosa.Top = 1920
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub txtKdDiagnosa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If fraDiagnosa.Visible = False Then Exit Sub
        dgDiagnosa.SetFocus
    End If
End Sub

Private Sub txtKdDiagnosa_KeyPress(KeyAscii As Integer)
    On Error GoTo errorLahYaw
    If KeyAscii = 13 Then
        If intJmlDiagnosa = 0 Then Exit Sub
        If fraDiagnosa.Visible = True Then
            dgDiagnosa.SetFocus
        Else
            txtDiagnosa.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        fraDiagnosa.Visible = False
    End If
errorLahYaw:
End Sub

Private Function sp_DiagnosaKeperawatan(f_Status As String) As Boolean
    sp_DiagnosaKeperawatan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdDiagnosaKeperawatan", adVarChar, adParamInput, 10, Trim(txtKdDiagnosaKeperawatan.Text))
        .Parameters.Append .CreateParameter("DiagnosaKeperawatan", adVarChar, adParamInput, 500, Trim(txtDiagnosa.Text))
        .Parameters.Append .CreateParameter("OutputKode", adVarChar, adParamOutput, 10, Null)
        .Parameters.Append .CreateParameter("KdAskep", adVarChar, adParamInput, 4, dcKdDiagnosa.BoundText)
        '.Parameters.Append .CreateParameter("KdAskep", adVarChar, adParamInput, 4, txtKdDiagnosa.Text)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 500, txtNamaExternal.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_DiagnosaKeperawatan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            If f_Status = "A" Then
                MsgBox "Gagal menyimpan data", vbCritical, "Validasi"
            Else
                MsgBox "Gagal menghapus data", vbCritical, "Validasi"
            End If
            sp_DiagnosaKeperawatan = False
        End If

        If f_Status = "A" Then
            txtKdDiagnosaKeperawatan.Text = .Parameters("OutputKode").value
            MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
        Else
            MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub txtKodeExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal.SetFocus
End Sub

Private Sub txtNamaExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl.SetFocus
End Sub

Private Sub CheckStatusEnbl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub
