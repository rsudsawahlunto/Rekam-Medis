VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRujukan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Rujukan"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRujukan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9495
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   0
      TabIndex        =   30
      Top             =   4560
      Width           =   9495
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   6480
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Rujukan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      TabIndex        =   22
      Top             =   2040
      Width           =   9495
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         MaxLength       =   10
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo dcTempatPerujuk 
         Height          =   330
         Left            =   5520
         TabIndex        =   8
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
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
      Begin MSDataListLib.DataCombo dcRujukanAsal 
         Height          =   330
         Left            =   2160
         TabIndex        =   7
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
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
      Begin MSComCtl2.DTPicker dtpTglRujuk 
         Height          =   330
         Left            =   240
         TabIndex        =   9
         Top             =   1320
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
         Format          =   392757251
         UpDown          =   -1  'True
         CurrentDate     =   37813
      End
      Begin MSDataListLib.DataCombo dcNamaPerujuk 
         Height          =   330
         Left            =   2280
         TabIndex        =   10
         Top             =   1320
         Width           =   3495
         _ExtentX        =   6165
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
      Begin MSDataListLib.DataCombo dcDiagnosa 
         Height          =   330
         Left            =   5880
         TabIndex        =   11
         Top             =   1320
         Width           =   3375
         _ExtentX        =   5953
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Diagnosa (Penyakit) Rujukan"
         Height          =   210
         Index           =   8
         Left            =   5880
         TabIndex        =   29
         Top             =   1080
         Width           =   2325
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Nama Perujuk"
         Height          =   210
         Index           =   7
         Left            =   2280
         TabIndex        =   28
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Rujukan Asal"
         Height          =   210
         Index           =   3
         Left            =   2160
         TabIndex        =   27
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Nama Tempat Perujuk"
         Height          =   210
         Index           =   4
         Left            =   5520
         TabIndex        =   26
         Top             =   360
         Width           =   1830
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Dirujuk"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Nama Perujuk = Nama Dokter/Bidan/Mantri/Dukun/Paranormal "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   1920
         Width           =   5475
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Nama Tempat Perujuk = Nama Puskesmas/Nama Klinik/Tempat Dokter Praktek/Nama Rumah Sakit "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   8505
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Pasien"
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
      TabIndex        =   14
      Top             =   960
      Width           =   9495
      Begin VB.Frame Frame4 
         Caption         =   "Umur"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6600
         TabIndex        =   15
         Top             =   360
         Width           =   2775
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   5
            Top             =   188
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   4
            Top             =   188
            Width           =   375
         End
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   240
            MaxLength       =   6
            TabIndex        =   3
            Top             =   188
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Left            =   2400
            TabIndex        =   18
            Top             =   240
            Width           =   165
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Left            =   1560
            TabIndex        =   17
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Left            =   720
            TabIndex        =   16
            Top             =   240
            Width           =   285
         End
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5160
         MaxLength       =   9
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   12
         TabIndex        =   0
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   5160
         TabIndex        =   21
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   2160
         TabIndex        =   20
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   585
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   32
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
      Left            =   7680
      Picture         =   "frmRujukan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRujukan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRujukan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmRujukan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub subLoadTempatRujukan()
    On Error GoTo errLoad
    Dim tempKode As String
    mstrKdInstalasiPerujuk = frmRegistrasiAll.dcInstalasi.BoundText

    tempKode = frmRegistrasiAll.dcRujukanRI.BoundText
    If tempKode = "08" Or tempKode = "09" Or tempKode = "10" Or tempKode = "01" Or tempKode = "06" Or tempKode = "07" Or tempKode = "11" Or tempKode = "12" Then
        strSQL = "Select KdRuangan,NamaRuangan From Ruangan where StatusEnabled='1'"
        Call msubDcSource(dcTempatPerujuk, rs, strSQL)
    ElseIf tempKode = "02" Or tempKode = "03" Or tempKode = "04" Or tempKode = "05" Then
        strSQL = "select KdDetailRujukanAsal,DetailRujukanAsal from dbo.DetailRujukanAsal where KdRujukanAsal = '" & tempKode & "' and StatusEnabled='1'"
        Call msubDcSource(dcTempatPerujuk, rs, strSQL)
    End If
    If rs.EOF = False Then dcTempatPerujuk.Text = rs(1)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    Call sp_Rujukan(dbcmd)
    Call subEnableControl(False)
    If strRegistrasi = "RJ" Then
        frmRegistrasiAll.cmdRujukan.Enabled = False
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcDiagnosa_GotFocus()
    strSQL = "SELECT NamaDiagnosa FROM Diagnosa where StatusEnabled='1' ORDER BY NamaDiagnosa"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcDiagnosa.RowSource = rs
    dcDiagnosa.ListField = rs.Fields(0).Name
    Set rs = Nothing
End Sub

Private Sub dcDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub dcDiagnosa_LostFocus()
    dcDiagnosa = StrConv(dcDiagnosa, vbProperCase)
End Sub

Private Sub dcNamaPerujuk_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcNamaPerujuk.BoundText
    strSQL = "SELECT NamaDokter FROM V_DaftarDokter ORDER BY NamaDokter"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcNamaPerujuk.RowSource = rs
    dcNamaPerujuk.ListField = rs.Fields(0).Name
    Set rs = Nothing
    dcNamaPerujuk.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcNamaPerujuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dcDiagnosa.SetFocus
End Sub

Private Sub dcNamaPerujuk_LostFocus()
    dcNamaPerujuk = StrConv(dcNamaPerujuk, vbProperCase)
End Sub

Private Sub dcRujukanAsal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcTempatPerujuk.SetFocus
End Sub

Private Sub dcTempatPerujuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dtpTglRujuk.SetFocus
End Sub

Private Sub dcTempatPerujuk_LostFocus()
    dcTempatPerujuk = StrConv(dcTempatPerujuk, vbProperCase)
End Sub

Private Sub dtpTglRujuk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcNamaPerujuk.SetFocus
End Sub

Private Sub Form_Activate()
    dcTempatPerujuk.SetFocus
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglRujuk.value = Now
    Call subLoadTempatRujukan
    dcRujukanAsal.Enabled = False
End Sub

'untuk cek validasi
Private Function funcCekValidasi() As Boolean
    If dcRujukanAsal.Text = "" Then
        MsgBox "Pilihan Rujukan Asal harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        dcRujukanAsal.SetFocus
        Exit Function
    End If
    If dcTempatPerujuk.Text = "" Then
        MsgBox "Pilihan Tempat Perujuk harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        dcTempatPerujuk.SetFocus
        Exit Function
    End If
    funcCekValidasi = True
End Function

'Store procedure untuk mengisi identitas pasien
Private Sub sp_Rujukan(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("NoRujukan", adVarChar, adParamInput, 30, Null)
        .Parameters.Append .CreateParameter("KdRujukanAsal", adChar, adParamInput, 2, mstrKdInstalasiPerujuk)
        .Parameters.Append .CreateParameter("SubRujukanAsal", adVarChar, adParamInput, 100, dcTempatPerujuk.Text)
        .Parameters.Append .CreateParameter("NamaPerujuk", adVarChar, adParamInput, 50, dcNamaPerujuk.Text)
        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , Format(dtpTglRujuk.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("DiagnosaRujukan", adVarChar, adParamInput, 100, dcDiagnosa.Text)

        .ActiveConnection = dbConn
        .CommandText = "AU_Rujukan"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam Pemasukan Data Rujukan", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("AU_Rujukan")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
        mstrKdInstalasiPerujuk = ""
    End With
    Exit Sub
End Sub

'untuk enable control
Private Sub subEnableControl(blnStatus As Boolean)
    dcRujukanAsal.Enabled = blnStatus
    dcTempatPerujuk.Enabled = blnStatus
    dcNamaPerujuk.Enabled = blnStatus
    dtpTglRujuk.Enabled = blnStatus
    dcDiagnosa.Enabled = blnStatus
    cmdSimpan.Enabled = blnStatus
End Sub

