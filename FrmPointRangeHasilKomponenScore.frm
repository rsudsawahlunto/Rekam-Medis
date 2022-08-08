VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form FrmPointRangeHasilKomponenScore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Range Hasil Komponen Score"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPointRangeHasilKomponenScore.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8565
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
      Height          =   1935
      Left            =   0
      TabIndex        =   11
      Top             =   1005
      Width           =   8535
      Begin VB.TextBox txtPoint 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5280
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtHasilMax 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtRangeHasil 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox txtHasilMin 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtKdRangeHasil 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         MaxLength       =   7
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   600
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dcKomponenScore 
         Height          =   330
         Left            =   2880
         TabIndex        =   4
         Top             =   1320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Point"
         Height          =   210
         Index           =   4
         Left            =   5280
         TabIndex        =   18
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Komponen Score"
         Height          =   210
         Index           =   1
         Left            =   2880
         TabIndex        =   17
         Top             =   1080
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasil Maximal"
         Height          =   210
         Index           =   3
         Left            =   1560
         TabIndex        =   16
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Range Hasil"
         Height          =   210
         Index           =   2
         Left            =   1560
         TabIndex        =   15
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasil Minimum"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1110
      End
   End
   Begin VB.Frame Frame3 
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
      TabIndex        =   13
      Top             =   5760
      Width           =   8535
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   6885
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   5160
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid dgData 
      Height          =   2655
      Left            =   0
      TabIndex        =   6
      Top             =   3000
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
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
      Left            =   6720
      Picture         =   "FrmPointRangeHasilKomponenScore.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "FrmPointRangeHasilKomponenScore.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "FrmPointRangeHasilKomponenScore.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "FrmPointRangeHasilKomponenScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    Call msubDcSource(dcKomponenScore, rs, "SELECT KdKomponenScore, NamaKomponenScore FROM KomponenScore where StatusEnabled='1'")

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdBatal_Click()
    On Error GoTo errLoad
    Call subKosong
    Call subLoadDcSource
    Call subLoadGridSource
    Exit Sub
errLoad:
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errLoad

    If txtKdRangeHasil.Text = "" Then Exit Sub
    If MsgBox("Apakah anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    If sp_RangeHasil("D") = False Then Exit Sub

    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"
    Call cmdBatal_Click

    Exit Sub
errLoad:
    MsgBox "Penghapusan Gagal, Data Sudah Terpakai !", vbOKOnly, "Informasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    If Periksa("text", txtRangeHasil, "Range hasil kosong") = False Then Exit Sub
    If Periksa("text", txtHasilMin, "Range hasil minimum kosong") = False Then Exit Sub
    If Periksa("text", txtHasilMax, "Range hasil maximum kosong") = False Then Exit Sub
    If Periksa("datacombo", dcKomponenScore, "Komponen score kosong") = False Then Exit Sub
    If Periksa("text", txtPoint, "Point kosong") = False Then Exit Sub

    If sp_RangeHasil("A") = False Then Exit Sub
    MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcKomponenScore_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKomponenScore.MatchedWithList = True Then txtPoint.SetFocus
        strSQL = "SELECT KdKomponenScore, NamaKomponenScore FROM KomponenScore where StatusEnabled='1' and (NamaKomponenScore LIKE '%" & dcKomponenScore.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcKomponenScore.Text = ""
            txtPoint.SetFocus
            Exit Sub
        End If
        dcKomponenScore.BoundText = rs(0).value
        dcKomponenScore.Text = rs(1).value
    End If
End Sub

Private Sub dgData_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgData
    WheelHook.WheelHook dgData
End Sub

Private Sub dgData_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errLoad

    If dgData.ApproxCount = 0 Then Exit Sub
    txtKdRangeHasil.Text = dgData.Columns("KdRangeHasil")
    txtRangeHasil.Text = dgData.Columns("RangeHasil")
    txtHasilMin.Text = dgData.Columns("HasilMin")
    txtHasilMax.Text = dgData.Columns("HasilMax")
    dcKomponenScore.BoundText = dgData.Columns("KdKomponenScore")
    txtPoint.Text = dgData.Columns("Point")

    Exit Sub
errLoad:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()
    strSQL = "SELECT dbo.PointRangeHasilKomponenScore.RangeHasil, dbo.PointRangeHasilKomponenScore.HasilMin, dbo.PointRangeHasilKomponenScore.HasilMax, dbo.KomponenScore.NamaKomponenScore, dbo.PointRangeHasilKomponenScore.Point, dbo.PointRangeHasilKomponenScore.KdRangeHasil, dbo.PointRangeHasilKomponenScore.KdKomponenScore " & _
    " FROM dbo.KomponenScore INNER JOIN dbo.PointRangeHasilKomponenScore ON dbo.KomponenScore.KdKomponenScore = dbo.PointRangeHasilKomponenScore.KdKomponenScore"
    Call msubRecFO(rs, strSQL)
    With dgData
        Set .DataSource = rs
    End With
End Sub

Sub subKosong()
    txtKdRangeHasil.Text = ""
    txtRangeHasil.Text = ""
    txtHasilMin.Text = ""
    txtHasilMax.Text = ""
    dcKomponenScore.BoundText = ""
    txtPoint.Text = ""
End Sub

Private Function sp_RangeHasil(f_Status As String) As Boolean
    sp_RangeHasil = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdRangeHasil", adVarChar, adParamInput, 7, txtKdRangeHasil.Text)
        .Parameters.Append .CreateParameter("RangeHasil", adVarChar, adParamInput, 50, Trim(txtRangeHasil.Text))
        .Parameters.Append .CreateParameter("HasilMin", adDouble, adParamInput, , CDbl(txtHasilMin.Text))
        .Parameters.Append .CreateParameter("HasilMax", adDouble, adParamInput, , CDbl(txtHasilMax.Text))
        .Parameters.Append .CreateParameter("Point", adDouble, adParamInput, , CDbl(txtPoint.Text))
        .Parameters.Append .CreateParameter("KdKomponenScore", adVarChar, adParamInput, 4, dcKomponenScore.BoundText)
        .Parameters.Append .CreateParameter("OutputKdRangeHasil", adVarChar, adParamOutput, 7, Null)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_PointRangeHasilKomponenScore"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_RangeHasil = False
        Else
            txtKdRangeHasil.Text = .Parameters("OutputKdRangeHasil")
            Call Add_HistoryLoginActivity("AUD_PointRangeHasilKomponenScore")
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub txtHasilMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcKomponenScore.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtHasilMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtHasilMax.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtPoint_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtRangeHasil_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtHasilMin.SetFocus
End Sub

