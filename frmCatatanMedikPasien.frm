VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCatatanMedikPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Catatan Medik Pasien"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCatatanMedikPasien.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   9270
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   32
      Top             =   6720
      Width           =   9255
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   465
         Left            =   7320
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   465
         Left            =   5280
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraDokter 
      Caption         =   "Data Dokter"
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
      Left            =   120
      TabIndex        =   31
      Top             =   3240
      Visible         =   0   'False
      Width           =   9015
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   2055
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   3625
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
   Begin VB.Frame fraCatatanMedis 
      Caption         =   "Catatan Medik Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   0
      TabIndex        =   25
      Top             =   2160
      Width           =   9255
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2400
         TabIndex        =   8
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   2040
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   3720
         Width           =   6975
      End
      Begin VB.TextBox txtPengobatan 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   2040
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2520
         Width           =   6975
      End
      Begin VB.TextBox txtKeluhanUtama 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   2040
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1320
         Width           =   6975
      End
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   118292483
         CurrentDate     =   38077
      End
      Begin MSDataListLib.DataCombo dcTriase 
         Height          =   330
         Left            =   7200
         TabIndex        =   10
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo dcImunisasi 
         Height          =   330
         Left            =   5400
         TabIndex        =   35
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Imunisasi"
         Height          =   210
         Index           =   2
         Left            =   5400
         TabIndex        =   36
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Triase"
         Height          =   210
         Index           =   1
         Left            =   7200
         TabIndex        =   33
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Pemeriksa"
         Height          =   210
         Index           =   0
         Left            =   2400
         TabIndex        =   30
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   210
         Left            =   2040
         TabIndex        =   29
         Top             =   3480
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pengobatan"
         Height          =   210
         Left            =   2040
         TabIndex        =   28
         Top             =   2280
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Keluhan Utama"
         Height          =   210
         Left            =   2040
         TabIndex        =   27
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Pemeriksaan"
         Height          =   210
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1710
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   1215
      Left            =   0
      TabIndex        =   16
      Top             =   960
      Width           =   9255
      Begin VB.Frame Frame5 
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
         Height          =   580
         Left            =   6600
         TabIndex        =   17
         Top             =   360
         Width           =   2415
         Begin VB.TextBox txtHari 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   6
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   900
            MaxLength       =   6
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   4
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2130
            TabIndex        =   20
            Top             =   270
            Width           =   150
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1350
            TabIndex        =   19
            Top             =   270
            Width           =   210
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   555
            TabIndex        =   18
            Top             =   270
            Width           =   240
         End
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1800
         TabIndex        =   23
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   2880
         TabIndex        =   22
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   5280
         TabIndex        =   21
         Top             =   360
         Width           =   1065
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   34
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
      Left            =   7560
      Picture         =   "frmCatatanMedikPasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1755
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmCatatanMedikPasien.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmCatatanMedikPasien.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmCatatanMedikPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilterDokter As String
Dim intJmlDokter As Integer

Private Sub cmdSimpan_Click()
    Dim adocomd As New ADODB.Command
    If mstrKdDokter = "" Then
        MsgBox "Silahkan pilih Dokter pemeriksa ", vbExclamation, "Validasi"
        txtDokter.SetFocus
        Exit Sub
    End If
    cmdSimpan.Enabled = False
    fraCatatanMedis.Enabled = False
    Call sp_CatatanMedisRJ(adocomd)
End Sub

Private Sub cmdTutup_Click()
    If cmdSimpan.Enabled = True Then
        If txtNoPendaftaran.Text <> "" Then
            If MsgBox("Simpan catatan medik", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
                Call cmdSimpan_Click
                Exit Sub
            End If
        End If
    End If
    Unload Me
End Sub

Private Sub dcTriase_Change()
    Select Case LCase(dcTriase.Text)
        Case "hijau"
            dcTriase.BackColor = vbGreen
        Case "kuning"
            dcTriase.BackColor = vbYellow
        Case "merah"
            dcTriase.BackColor = vbRed
        Case "biru"
            dcTriase.BackColor = vbBlue
        Case Else
            dcTriase.BackColor = vbWhite
    End Select
End Sub

Private Sub dcTriase_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKeluhanUtama.SetFocus
End Sub

Private Sub dcTriase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcTriase.MatchedWithList = True Then dcTriase.SetFocus
        strSQL = "SELECT KdTriase, NamaTriase FROM Triase where StatusEnabled='1' and (NamaTriase LIKE '%" & dcTriase.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcTriase.Text = ""
            Exit Sub
        End If
        dcTriase.BoundText = rs(0).value
        dcTriase.Text = rs(1).value
    End If
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        txtDokter.Text = dgDokter.Columns(1).value
        mstrKdDokter = dgDokter.Columns(0).value
        If mstrKdDokter = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        fraDokter.Visible = False
        dcTriase.SetFocus
    End If
    If KeyAscii = 27 Then
        fraDokter.Visible = False
    End If
End Sub

Private Sub dtpTglPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtDokter.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglPeriksa.value = Now

    strSQL = "SELECT KdTriase, NamaTriase FROM Triase where StatusEnabled='1'"
    Call msubDcSource(dcTriase, rs, strSQL)
    If rs.EOF = False Then dcTriase.BoundText = rs(0).value

    strSQL = "SELECT KdImunisasi, NamaImunisasi FROM Imunisasi where StatusEnabled='1'"
    Call msubDcSource(dcImunisasi, rs, strSQL)
    If rs.EOF = False Then dcImunisasi.BoundText = rs(0).value

    With frmTransaksiPasien
        txtNoPendaftaran = .txtNoPendaftaran.Text
        txtNoCM = .txtNoCM.Text
        txtNamaPasien = .txtNamaPasien.Text
        txtSex.Text = .txtSex.Text
        txtThn = .txtThn.Text
        txtBln = .txtBln.Text
        txtHari = .txtHr.Text
    End With

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmTransaksiPasien.Enabled = True
    Call frmTransaksiPasien.subLoadRiwayatCatatanMedis
End Sub

Private Sub txtDokter_Change()
    strFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
    mstrKdDokter = ""
    fraDokter.Visible = True
    Call subLoadDokter
End Sub

Private Sub txtDokter_GotFocus()
    txtDokter.SelStart = 0
    txtDokter.SelLength = Len(txtDokter.Text)
    fraDokter.Visible = True
    strFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
    Call subLoadDokter
End Sub

Private Sub txtDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If fraDokter.Visible = False Then Exit Sub
        dgDokter.SetFocus
    End If
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 27 Then
        fraDokter.Visible = False
    End If
    If KeyAscii = 13 Then
        If fraDokter.Visible = True Then
            dgDokter.SetFocus
        Else
            dcTriase.SetFocus
        End If
    End If
    Call SetKeyPressToChar(KeyAscii)
hell:
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter()
    On Error Resume Next
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & strFilterDokter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDokter = rs.RecordCount
    Set dgDokter.DataSource = rs
    With dgDokter
        .Columns(0).Width = 1500
        .Columns(1).Width = 4000
        .Columns(2).Width = 400
        .Columns(3).Alignment = dbgCenter
        .Columns(3).Width = 2000
    End With
End Sub

Private Sub txtKeluhanUtama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPengobatan.SetFocus
End Sub

Private Sub txtKeluhanUtama_LostFocus()
    Dim i As Integer
    Dim tempText As String

    tempText = Trim(txtKeluhanUtama.Text)
    txtKeluhanUtama.Text = ""
    For i = 1 To Len(tempText)
        If Asc(Mid(tempText, i, 1)) <> 10 And Asc(Mid(tempText, i, 1)) <> 13 Then
            txtKeluhanUtama.Text = txtKeluhanUtama.Text & Mid(tempText, i, 1)
        End If
    Next i
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKeterangan_LostFocus()
    Dim i As Integer
    Dim tempText As String

    tempText = Trim(txtKeterangan.Text)
    txtKeterangan.Text = ""
    For i = 1 To Len(tempText)
        If Asc(Mid(tempText, i, 1)) <> 10 And Asc(Mid(tempText, i, 1)) <> 13 Then
            txtKeterangan.Text = txtKeterangan.Text & Mid(tempText, i, 1)
        End If
    Next i
End Sub

Private Sub txtPengobatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub txtPengobatan_LostFocus()
    Dim i As Integer
    Dim tempText As String

    tempText = Trim(txtPengobatan.Text)
    txtPengobatan.Text = ""
    For i = 1 To Len(tempText)
        If Asc(Mid(tempText, i, 1)) <> 10 And Asc(Mid(tempText, i, 1)) <> 13 Then
            txtPengobatan.Text = txtPengobatan.Text & Mid(tempText, i, 1)
        End If
    Next i
End Sub

'Store procedure untuk mengisi catatan medis pasien
Private Sub sp_CatatanMedisRJ(ByVal adoCommand As ADODB.Command)
    On Error GoTo errLoad

    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dtpTglPeriksa.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, mstrKdDokter)
        .Parameters.Append .CreateParameter("KeluhanUtama", adVarChar, adParamInput, 1000, txtKeluhanUtama.Text)
        .Parameters.Append .CreateParameter("KdImunisasi", adChar, adParamInput, 3, IIf(dcImunisasi.Text = "", Null, dcImunisasi.BoundText))
        .Parameters.Append .CreateParameter("Pengobatan", adVarChar, adParamInput, 1000, txtPengobatan.Text)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 1000, txtKeterangan.Text)
        .Parameters.Append .CreateParameter("KdTriase", adChar, adParamInput, 2, IIf(dcTriase.BoundText = "", Null, dcTriase.BoundText))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AU_CatatanMedikPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Catatan Medis Pasien", vbCritical, "Validasi"
        Else
            MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
            Call Add_HistoryLoginActivity("AU_CatatanMedikPasien")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub
