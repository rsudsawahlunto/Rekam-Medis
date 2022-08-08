VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmTujuanNRencanaTindakan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Tujuan & Rencana Tindakan"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTujuanNRencanaTindakan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   13920
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
      Height          =   855
      Left            =   0
      TabIndex        =   19
      Top             =   7080
      Width           =   13935
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   11310
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   495
         Left            =   10095
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   495
         Left            =   8880
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   12525
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   2295
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
         TabIndex        =   21
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cari Diagnosa Keperawatan"
         Height          =   210
         Left            =   1560
         TabIndex        =   20
         Top             =   120
         Width           =   2205
      End
   End
   Begin VB.Frame fraDiagnosa 
      Caption         =   "Data Diagnosa Keperawatan"
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
      Left            =   960
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   11175
      Begin MSDataGridLib.DataGrid dgDiagnosaKeperawatan 
         Height          =   2775
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   22
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
      Height          =   6135
      Left            =   0
      TabIndex        =   13
      Top             =   960
      Width           =   13935
      Begin VB.TextBox txtDiagnosaKeperawatan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   7095
      End
      Begin VB.TextBox txtMemo 
         Appearance      =   0  'Flat
         Height          =   765
         Left            =   240
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1320
         Width           =   13455
      End
      Begin VB.TextBox txtKdDiagnosa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dcDetailAsKep 
         Height          =   330
         Left            =   8880
         TabIndex        =   4
         Top             =   600
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcPenyebab 
         Height          =   330
         Left            =   8880
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid dgTindakan 
         Height          =   3615
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   6376
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Diagnosa Keperawatan"
         Height          =   210
         Index           =   0
         Left            =   1680
         TabIndex        =   18
         Top             =   360
         Width           =   1860
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Detail Asuhan Keperawatan"
         Height          =   210
         Index           =   0
         Left            =   8880
         TabIndex        =   17
         Top             =   360
         Width           =   2250
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Memo Asuhan keperawatan"
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   2280
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Kode Diagnosa"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Penyebab"
         Height          =   210
         Index           =   1
         Left            =   8880
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   810
      End
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12120
      Picture         =   "frmTujuanNRencanaTindakan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTujuanNRencanaTindakan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12255
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmTujuanNRencanaTindakan.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmTujuanNRencanaTindakan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilterDiagnosa As String
Dim intJmlDiagnosa As Integer
Dim subKdDetailAskep As String
Dim subKdDiagnosaKeperawatan As String

Private Sub cmdBatal_Click()
    On Error GoTo errLoad
    Call subLoadDcSource
    Call clear
    Call subLoadGridSource
    txtDiagnosaKeperawatan.SetFocus
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdHapus_Click()
    If txtDiagnosaKeperawatan.Text = "" Then
        MsgBox "Pilih dulu Diagnosa yang akan dihapus (Dobel klik)", vbOKOnly, "Validasi"
        Exit Sub
    End If
    If MsgBox("Apakah anda yakin akan mengapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub

    If sp_Tindakan("M") = False Then Exit Sub

    Call cmdBatal_Click
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    If Periksa("text", txtDiagnosaKeperawatan, "Diagnosa Keperawatan kosong") = False Then Exit Sub
    If Periksa("datacombo", dcDetailAsKep, "Detail Asuhan Keperawatan kosong") = False Then Exit Sub
    If Periksa("text", txtMemo, "Memo Asuhan Keperawatan kosong") = False Then Exit Sub

    If sp_Tindakan("A") = False Then Exit Sub
    Call Add_HistoryLoginActivity("AUD_TujuanNRencanaTindakan")
    Call cmdBatal_Click
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdCetak_Click()
    frmCetakTujuanNRencanaTindakan.Show
End Sub

Private Sub dcDetailAsKep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        subKdDetailAskep = dcDetailAsKep.BoundText
        If dcDetailAsKep.MatchedWithList = True Then txtMemo.SetFocus
        strSQL = "select kddetailaskep, detailaskep from DetailDiagnosaKeperawatan where StatusEnabled='1' and (detailaskep LIKE '%" & dcDetailAsKep.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcDetailAsKep.Text = ""
            Exit Sub
        End If
        dcDetailAsKep.BoundText = rs(0).value
        dcDetailAsKep.Text = rs(1).value
    End If
End Sub

Private Sub dcPenyebab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcPenyebab.MatchedWithList = True Then dcDetailAsKep.SetFocus
        strSQL = "select kdpenyebab, penyebabaskep from PenyebabDiagnosaKeperawatan where StatusEnabled='1' and (penyebabaskep LIKE '%" & dcPenyebab.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcPenyebab.Text = ""
            Exit Sub
        End If
        dcPenyebab.BoundText = rs(0).value
        dcPenyebab.Text = rs(1).value
    End If
End Sub

Private Sub dgDiagnosaKeperawatan_DblClick()
    Call dgDiagnosaKeperawatan_KeyPress(13)
    WheelHook.WheelUnHook
    Set MyProperty = dgDiagnosaKeperawatan
    WheelHook.WheelHook dgDiagnosaKeperawatan
End Sub

Private Sub dgDiagnosaKeperawatan_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        txtDiagnosaKeperawatan.Text = dgDiagnosaKeperawatan.Columns(1).value
        txtKdDiagnosa.Text = dgDiagnosaKeperawatan.Columns(0).value
        fraDiagnosa.Visible = False
        dcPenyebab.SetFocus
    End If
    If KeyAscii = 27 Then
        fraDiagnosa.Visible = False
    End If
    Exit Sub
hell:
End Sub

Private Sub dgTindakan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgTindakan
    WheelHook.WheelHook dgTindakan
End Sub

Private Sub dgTindakan_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    Me.Caption = dgDiagnosaKeperawatan.Columns(ColIndex).Width
End Sub

Private Sub dgTindakan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtMemo.SetFocus
    If KeyAscii = 27 Then fraDiagnosa.Visible = False
End Sub

Private Sub dgTindakan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    subTampil = True
    txtKdDiagnosa.Text = dgTindakan.Columns("KdDiagnosaKeperawatan")
    txtDiagnosaKeperawatan.Text = dgTindakan.Columns("DiagnosaKeperawatan")
    dcDetailAsKep.BoundText = dgTindakan.Columns("KdDetailAsKep")
    txtMemo.Text = dgTindakan.Columns("MemoAsKep")
    fraDiagnosa.Visible = False
    subTampil = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            If dgTindakan.ApproxCount = 0 Then Exit Sub
            frmCetakTujuanNRencanaTindakan.Show
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
    Call subLoadDcSource
    Call clear
    Call subLoadGridSource

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()
    On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "select KdDiagnosaKeperawatan,KdDetailAsKep,DiagnosaKeperawatan,PenyebabAsKep,DetailAsKep,MemoAsKep from V_TujuanNRencanaTindakan WHERE DiagnosaKeperawatan LIKE '%" & txtParameter.Text & "%' and StatusEnabled='1' and Expr1='1' and Expr2='1'"
    rs.Open strSQL, dbConn, adOpenDynamic, adLockOptimistic
    Set dgTindakan.DataSource = rs
    With dgTindakan
        .Columns("KdDiagnosaKeperawatan").Width = 0
        .Columns("KdDetailAsKep").Width = 0
        .Columns("DiagnosaKeperawatan").Width = 6800
        .Columns("PenyebabAsKep").Width = 2500
        .Columns("DetailAsKep").Width = 1500
        .Columns("MemoAsKep").Width = 4000
    End With
    Set rs = Nothing

    cmdHapus.Visible = True

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub clear()
    txtKdDiagnosa.Text = ""
    txtDiagnosaKeperawatan.Text = ""
    dcPenyebab.BoundText = ""
    dcDetailAsKep.Text = ""
    txtMemo.Text = ""
    fraDiagnosa.Visible = False
End Sub

Private Sub txtMemo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtDiagnosaKeperawatan_Change()
    If subTampil = True Then Exit Sub
    strFilterDiagnosa = "WHERE DiagnosaKeperawatan like '%" & txtDiagnosaKeperawatan.Text & "%'"
    fraDiagnosa.Visible = True
    Call subLoadDiagnosa
End Sub

Private Sub subLoadDiagnosa()
    On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "select * from DiagnosaKeperawatan " & strFilterDiagnosa
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDiagnosa = rs.RecordCount
    Set dgDiagnosaKeperawatan.DataSource = rs
    With dgDiagnosaKeperawatan
        .Columns(0).Caption = "Kode Diagnosa Keperawatan"
        .Columns(0).Width = 1200
        .Columns(1).Caption = "Diagnosa Keperawatan"
        .Columns(1).Width = 5200
    End With
    fraDiagnosa.Left = 0
    fraDiagnosa.Top = 1920
    Exit Sub
errLoad:
End Sub

Private Function sp_Tindakan(f_Status As String) As Boolean
    On Error GoTo hell
    sp_Tindakan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdDiagnosaKeperawatan", adVarChar, adParamInput, 10, txtKdDiagnosa.Text)
        .Parameters.Append .CreateParameter("KdDetailAsKep", adVarChar, adParamInput, 3, dcDetailAsKep.BoundText)
        .Parameters.Append .CreateParameter("MemoAsKep", adVarChar, adParamInput, 1000, Trim(txtMemo.Text))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_TujuanNRencanaTindakan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            If f_Status = "A" Then
                MsgBox "Gagal menyimpan data", vbCritical, "Validasi"
            Else
                MsgBox "Gagal menghapus data", vbCritical, "Validasi"
            End If
            sp_Tindakan = False
        End If

        If f_Status = "A" Then
            
        Else
            MsgBox "Berhasil menghapus data", vbOKOnly, "Validasi"
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError
End Function

Private Sub txtMemo_LostFocus()
    Dim i As Integer
    Dim tempText As String

    tempText = Trim(txtMemo.Text)
    txtMemo.Text = ""
    For i = 1 To Len(tempText)
        If Asc(Mid(tempText, i, 1)) <> 10 And Asc(Mid(tempText, i, 1)) <> 13 Then
            txtMemo.Text = txtMemo.Text & Mid(tempText, i, 1)
        End If
    Next i
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    Call msubDcSource(dcDetailAsKep, rs, "select * from DetailDiagnosaKeperawatan where StatusEnabled='1'")
    Call msubDcSource(dcPenyebab, rs, "select * from PenyebabDiagnosaKeperawatan where StatusEnabled='1'")

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtParameter_Change()

    strSQL = "select * from V_TujuanNRencanaTindakan WHERE DiagnosaKeperawatan LIKE '%" & txtParameter.Text & "%' and StatusEnabled='1' and Expr1='1' and Expr2='1'"
    Call msubRecFO(rs, strSQL)
    Set dgTindakan.DataSource = rs
    With dgTindakan
        .Columns("KdDiagnosaKeperawatan").Width = 0
        .Columns("KdDetailAsKep").Width = 0
        .Columns("DiagnosaKeperawatan").Width = 6800
        .Columns("PenyebabAsKep").Width = 1000
        .Columns("DetailAsKep").Width = 1000
        .Columns("MemoAsKep").Width = 4000
    End With
    Set rs = Nothing
End Sub
