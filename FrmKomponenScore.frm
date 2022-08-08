VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form FrmKomponenScore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Komponen Score"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmKomponenScore.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   12765
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
      Height          =   1815
      Left            =   0
      TabIndex        =   12
      Top             =   1005
      Width           =   12735
      Begin VB.TextBox txtKodeExternal 
         Appearance      =   0  'Flat
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
         Left            =   4200
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtNamaExternal 
         Appearance      =   0  'Flat
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
         Left            =   6120
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1320
         Width           =   4815
      End
      Begin VB.CheckBox CheckStatusEnbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         Height          =   255
         Left            =   11040
         TabIndex        =   6
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtNamaKomponenScore 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   6735
      End
      Begin VB.TextBox txtKdKomponenScore 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         MaxLength       =   4
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   600
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dcSatuanHasil 
         Height          =   330
         Left            =   8400
         TabIndex        =   2
         Top             =   600
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
      Begin MSDataListLib.DataCombo dcKomponenKlinis 
         Height          =   330
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Kode External"
         Height          =   210
         Left            =   4200
         TabIndex        =   20
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "NamaExternal"
         Height          =   210
         Left            =   6120
         TabIndex        =   19
         Top             =   1080
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Komponen Klinis"
         Height          =   210
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Satuan Hasil"
         Height          =   210
         Left            =   8400
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nama Komponen Score"
         Height          =   210
         Index           =   0
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   1920
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
      TabIndex        =   16
      Top             =   6480
      Width           =   12735
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   6120
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   11040
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   7800
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   9360
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid dgData 
      Height          =   3495
      Left            =   0
      TabIndex        =   7
      Top             =   2880
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   6165
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
      TabIndex        =   18
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
      Left            =   10920
      Picture         =   "FrmKomponenScore.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "FrmKomponenScore.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "FrmKomponenScore.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "FrmKomponenScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    strSQL = "SELECT dbo.KomponenKlinis.KdSatuanHasil, dbo.SatuanHasil.SatuanHasil " & _
    " FROM  dbo.KomponenKlinis INNER JOIN dbo.SatuanHasil ON dbo.KomponenKlinis.KdSatuanHasil = dbo.SatuanHasil.KdSatuanHasil where dbo.SatuanHasil.StatusEnabled='1' order by SatuanHasil"
    Call msubDcSource(dcSatuanHasil, rs, strSQL)
    Call msubDcSource(dcKomponenKlinis, rs, "SELECT KdKomponenKlinis, KomponenKlinis FROM  KomponenKlinis where StatusEnabled='1' order by KomponenKlinis")

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

    If txtKdKomponenScore.Text = "" Then Exit Sub
    If MsgBox("Apakah anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    If sp_KomponenScore("D") = False Then Exit Sub

    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"
    Call cmdBatal_Click

    Exit Sub
errLoad:
    MsgBox "Penghapusan Gagal, Data Sudah Terpakai !", vbOKOnly, "Informasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    If Periksa("text", txtNamaKomponenScore, "Nama komponen score kosong") = False Then Exit Sub
    If sp_KomponenScore("A") = False Then Exit Sub

    MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcKomponenKlinis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKomponenKlinis.MatchedWithList = True Then txtKodeExternal.SetFocus
        strSQL = "SELECT KdKomponenKlinis, KomponenKlinis FROM  KomponenKlinis where StatusEnabled='1'  and (KomponenKlinis LIKE '%" & dcKomponenKlinis.Text & "%')order by KomponenKlinis"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcKomponenKlinis.Text = ""
            Exit Sub
        End If
        dcKomponenKlinis.BoundText = rs(0).value
        dcKomponenKlinis.Text = rs(1).value
    End If
End Sub

Private Sub dcSatuanHasil_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcSatuanHasil.BoundText
    strSQL = "SELECT dbo.KomponenKlinis.KdSatuanHasil, dbo.SatuanHasil.SatuanHasil  " & _
    " FROM  dbo.KomponenKlinis INNER JOIN dbo.SatuanHasil ON dbo.KomponenKlinis.KdSatuanHasil = dbo.SatuanHasil.KdSatuanHasil WHERE (dbo.KomponenKlinis.KdKomponenKlinis = '" & dcKomponenKlinis.BoundText & "')"
    Call msubDcSource(dcSatuanHasil, rs, strSQL)
    dcSatuanHasil.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcSatuanHasil_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcSatuanHasil.MatchedWithList = True Then dcKomponenKlinis.SetFocus
        strSQL = "SELECT dbo.KomponenKlinis.KdSatuanHasil, dbo.SatuanHasil.SatuanHasil " & _
        " FROM  dbo.KomponenKlinis INNER JOIN dbo.SatuanHasil ON dbo.KomponenKlinis.KdSatuanHasil = dbo.SatuanHasil.KdSatuanHasil where dbo.SatuanHasil.StatusEnabled='1'  and (SatuanHasil LIKE '%" & dcSatuanHasil.Text & "%')order by SatuanHasil"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcSatuanHasil.Text = ""
            dcKomponenKlinis.SetFocus
            Exit Sub
        End If
        dcSatuanHasil.BoundText = rs(0).value
        dcSatuanHasil.Text = rs(1).value
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
    txtKdKomponenScore.Text = dgData.Columns("Kode Komponen Score")
    txtNamaKomponenScore.Text = dgData.Columns("Nama Komponen Score")
    dcKomponenKlinis.BoundText = dgData.Columns("Komponen Klinis")
    dcSatuanHasil.BoundText = dgData.Columns("Satuan Hasil")
    txtKodeExternal.Text = dgData.Columns("KodeExternal").value
    txtNamaExternal.Text = dgData.Columns("NamaExternal").value
    If dgData.Columns("StatusEnabled") = "" Then
        CheckStatusEnbl.value = 0
    ElseIf dgData.Columns("StatusEnabled") = 0 Then
        CheckStatusEnbl.value = 0
    ElseIf dgData.Columns("StatusEnabled") = 1 Then
        CheckStatusEnbl.value = 1
    End If

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
    strSQL = " SELECT dbo.KomponenScore.KdKomponenScore, dbo.KomponenScore.NamaKomponenScore, dbo.SatuanHasil.SatuanHasil, " & _
    " dbo.KomponenKlinis.KomponenKlinis, dbo.SatuanHasil.KdSatuanHasil, dbo.KomponenKlinis.KdKomponenKlinis, dbo.KomponenScore.KodeExternal, " & _
    " dbo.KomponenScore.NamaExternal, dbo.KomponenScore.StatusEnabled " & _
    " FROM dbo.KomponenScore LEFT OUTER JOIN " & _
    " dbo.KomponenKlinis ON dbo.KomponenScore.KdKomponenKlinis = dbo.KomponenKlinis.KdKomponenKlinis LEFT OUTER JOIN " & _
    " dbo.SatuanHasil ON dbo.KomponenScore.KdSatuanHasil = dbo.SatuanHasil.KdSatuanHasil "

    Call msubRecFO(rs, strSQL)
    With dgData
        Set .DataSource = rs
        .Columns(0).Width = 1500
        .Columns(0).Caption = "Kode Komponen Score"
        .Columns(1).Width = 4500
        .Columns(1).Caption = "Nama Komponen Score"
        .Columns(2).Width = 2000
        .Columns(2).Caption = "Satuan Hasil"
        .Columns(3).Width = 4000
        .Columns(3).Caption = "Komponen Klinis"
        .Columns(4).Width = 0
        .Columns(5).Width = 0
    End With
End Sub

Sub subKosong()
    txtKdKomponenScore.Text = ""
    txtNamaKomponenScore.Text = ""
    dcSatuanHasil.BoundText = ""
    dcKomponenKlinis.BoundText = ""
    txtKodeExternal.Text = ""
    txtNamaExternal.Text = ""
    CheckStatusEnbl.value = 1
End Sub

Private Function sp_KomponenScore(f_Status As String) As Boolean
    sp_KomponenScore = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKomponenScore", adVarChar, adParamInput, 4, txtKdKomponenScore.Text)
        .Parameters.Append .CreateParameter("NamaKomponenScore", adVarChar, adParamInput, 100, Trim(txtNamaKomponenScore.Text))
        .Parameters.Append .CreateParameter("KdSatuanHasil", adChar, adParamInput, 2, IIf(dcSatuanHasil.BoundText = "", Null, dcSatuanHasil.BoundText))
        .Parameters.Append .CreateParameter("KdKomponenKlinis", adVarChar, adParamInput, 3, IIf(dcKomponenKlinis.BoundText = "", Null, dcKomponenKlinis.BoundText))
        .Parameters.Append .CreateParameter("OutputKode", adVarChar, adParamOutput, 4, Null)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKodeExternal.Text))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNamaExternal.Text))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_KomponenScore"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_KomponenScore = False
        Else
            txtKdKomponenScore.Text = .Parameters("OutputKode")
            Call Add_HistoryLoginActivity("AUD_KomponenScore")
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub txtNamaKomponenScore_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcSatuanHasil.SetFocus
End Sub

Private Sub txtKodeExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal.SetFocus
End Sub

Private Sub txtNamaExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl.SetFocus
End Sub

Private Sub CheckStatusEnbl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

