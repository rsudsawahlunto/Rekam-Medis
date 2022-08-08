VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form FrmKomponenKlinis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Komponen Klinis"
   ClientHeight    =   7485
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
   Icon            =   "FrmKomponenKlinis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
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
      Height          =   2775
      Left            =   0
      TabIndex        =   11
      Top             =   1005
      Width           =   8535
      Begin VB.TextBox txtKodeExternal 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtNamaExternal 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   4
         Top             =   2280
         Width           =   6615
      End
      Begin VB.CheckBox CheckStatusEnbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         Height          =   255
         Left            =   6960
         TabIndex        =   5
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtNamaKomponenKlinis 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1320
         Width           =   8055
      End
      Begin VB.TextBox txtKdKomponenKlinis 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   600
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dcSatuanHasil 
         Height          =   330
         Left            =   1680
         TabIndex        =   1
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
      Begin VB.Label Label12 
         Caption         =   "Kode External"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Nama External"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Satuan Hasil"
         Height          =   210
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Komponen Klinis"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1320
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
      TabIndex        =   14
      Top             =   6720
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
      Height          =   2775
      Left            =   0
      TabIndex        =   6
      Top             =   3840
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4895
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
      TabIndex        =   16
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
      Picture         =   "FrmKomponenKlinis.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "FrmKomponenKlinis.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "FrmKomponenKlinis.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "FrmKomponenKlinis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    Call msubDcSource(dcSatuanHasil, rs, "SELECT KdSatuanHasil, SatuanHasil FROM SatuanHasil where StatusEnabled='1'")

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdBatal_Click()
    On Error GoTo errLoad
    Call subKosong
    Call subLoadDcSource
    Call subLoadGridSource
    txtNamaKomponenKlinis.SetFocus
    Exit Sub
errLoad:
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errLoad

    If txtKdKomponenKlinis.Text = "" Then Exit Sub
    If MsgBox("Apakah anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    If sp_KomponenKlinis("D") = False Then Exit Sub

    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"
    Call cmdBatal_Click

    Exit Sub
errLoad:
    MsgBox "Penghapusan Gagal, Data Sudah Terpakai !", vbOKOnly, "Informasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    If Periksa("text", txtNamaKomponenKlinis, "Nama komponen klinis kosong") = False Then Exit Sub
    If sp_KomponenKlinis("A") = False Then Exit Sub

    MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcSatuanHasil_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcSatuanHasil.MatchedWithList = True Then txtNamaKomponenKlinis.SetFocus
        strSQL = "SELECT KdSatuanHasil, SatuanHasil FROM SatuanHasil where StatusEnabled='1' and (SatuanHasil LIKE '%" & dcSatuanHasil.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcSatuanHasil.Text = ""
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

Private Sub dgData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcSatuanHasil.SetFocus
End Sub

Private Sub dgData_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errLoad

    If dgData.ApproxCount = 0 Then Exit Sub
    txtKdKomponenKlinis.Text = dgData.Columns("KdKomponenKlinis")
    txtNamaKomponenKlinis.Text = dgData.Columns("KomponenKlinis")
    dcSatuanHasil.BoundText = dgData.Columns("kdSatuanHasil")

    txtKodeExternal.Text = dgData.Columns("KodeExternal").value
    txtNamaExternal.Text = dgData.Columns("NamaExternal").value
    If dgData.Columns("StatusEnabled").value = "" Then
        CheckStatusEnbl.value = 0
    ElseIf dgData.Columns("StatusEnabled").value = 0 Then
        CheckStatusEnbl.value = 0
    ElseIf dgData.Columns("StatusEnabled").value = 1 Then
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
    strSQL = "SELECT dbo.KomponenKlinis.KomponenKlinis, dbo.SatuanHasil.SatuanHasil, dbo.KomponenKlinis.KdKomponenKlinis, dbo.KomponenKlinis.KdSatuanHasil, " & _
    "dbo.KomponenKlinis.KodeExternal, dbo.KomponenKlinis.NamaExternal, dbo.KomponenKlinis.StatusEnabled " & _
    "FROM dbo.KomponenKlinis LEFT OUTER JOIN dbo.SatuanHasil ON dbo.KomponenKlinis.KdSatuanHasil = dbo.SatuanHasil.KdSatuanHasil"
    Call msubRecFO(rs, strSQL)
    With dgData
        Set .DataSource = rs
        .Columns(0).Width = 5900
        .Columns(1).Width = 2000
    End With
End Sub

Sub subKosong()
    txtKdKomponenKlinis.Text = ""
    txtNamaKomponenKlinis.Text = ""
    dcSatuanHasil.BoundText = ""
    txtKodeExternal.Text = ""
    txtNamaExternal.Text = ""
    CheckStatusEnbl.value = 1
End Sub

Private Function sp_KomponenKlinis(f_Status As String) As Boolean
    sp_KomponenKlinis = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKomponenKlinis", adVarChar, adParamInput, 3, txtKdKomponenKlinis.Text)
        .Parameters.Append .CreateParameter("KomponenKlinis", adVarChar, adParamInput, 50, Trim(txtNamaKomponenKlinis.Text))
        .Parameters.Append .CreateParameter("KdSatuanHasil", adChar, adParamInput, 2, IIf(dcSatuanHasil.BoundText = "", Null, dcSatuanHasil.BoundText))
        .Parameters.Append .CreateParameter("OutputKode", adVarChar, adParamOutput, 3, Null)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKodeExternal.Text))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNamaExternal.Text))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_KomponenKlinis"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_KomponenKlinis = False
        Else
            txtKdKomponenKlinis.Text = .Parameters("OutputKode")
            Call Add_HistoryLoginActivity("AUD_KomponenKlinis")
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub txtKdKomponenKlinis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcSatuanHasil.SetFocus
End Sub

Private Sub txtNamaKomponenKlinis_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then txtKodeExternal.SetFocus
End Sub

Private Sub txtNamaKomponenKlinis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
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

