VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPasienGDKecelakaan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Kecelakaan Gawat Darurat"
   ClientHeight    =   5190
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
   Icon            =   "frmPasienGDKecelakaan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9495
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   26
      Top             =   4320
      Width           =   9495
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   465
         Left            =   5400
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   465
         Left            =   7440
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Kecelakaan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   14
      Top             =   2040
      Width           =   9495
      Begin VB.TextBox txtTempatKecelakaan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2160
         MaxLength       =   150
         TabIndex        =   11
         Top             =   1800
         Width           =   7095
      End
      Begin VB.TextBox txtNamaKecelakaan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2160
         MaxLength       =   150
         TabIndex        =   10
         Top             =   1440
         Width           =   7095
      End
      Begin MSComCtl2.DTPicker dtpTglKecelakaan 
         Height          =   330
         Left            =   2160
         TabIndex        =   9
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   118751235
         UpDown          =   -1  'True
         CurrentDate     =   38076
      End
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   2160
         TabIndex        =   7
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   118751235
         UpDown          =   -1  'True
         CurrentDate     =   38076
      End
      Begin MSDataListLib.DataCombo dcPerawat 
         Height          =   330
         Left            =   2160
         TabIndex        =   8
         Top             =   720
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "DataCombo1"
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Perawat"
         Height          =   210
         Index           =   11
         Left            =   240
         TabIndex        =   29
         Top             =   780
         Width           =   675
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Periksa"
         Height          =   210
         Index           =   10
         Left            =   240
         TabIndex        =   28
         Top             =   420
         Width           =   1260
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nama Kecelakaan"
         Height          =   210
         Index           =   8
         Left            =   240
         TabIndex        =   27
         Top             =   1500
         Width           =   1410
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Tempat Kecelakaan"
         Height          =   210
         Index           =   9
         Left            =   240
         TabIndex        =   25
         Top             =   1860
         Width           =   1605
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Kecelakaan"
         Height          =   210
         Index           =   7
         Left            =   240
         TabIndex        =   24
         Top             =   1140
         Width           =   1605
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
      Height          =   1095
      Left            =   0
      TabIndex        =   15
      Top             =   960
      Width           =   9495
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5520
         MaxLength       =   9
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   975
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
         Left            =   6840
         TabIndex        =   16
         Top             =   360
         Width           =   2415
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
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Index           =   4
            Left            =   550
            TabIndex        =   19
            Top             =   277
            Width           =   285
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Index           =   5
            Left            =   1350
            TabIndex        =   18
            Top             =   277
            Width           =   240
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Index           =   6
            Left            =   2130
            TabIndex        =   17
            Top             =   270
            Width           =   165
         End
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Index           =   3
         Left            =   5520
         TabIndex        =   23
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Index           =   2
         Left            =   3000
         TabIndex        =   22
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Index           =   1
         Left            =   1800
         TabIndex        =   21
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   30
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
      Picture         =   "frmPasienGDKecelakaan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPasienGDKecelakaan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPasienGDKecelakaan.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmPasienGDKecelakaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    If Periksa("datacombo", dcPerawat, "Nama pemeriksa kosong") = False Then Exit Sub
    If Periksa("text", txtNamaKecelakaan, "Nama kecelakaan kosong") = False Then Exit Sub

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("TglPeriksa", adDate, adParamInput, , Format(dtpTglPeriksa.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("NamaKecelakaan", adVarChar, adParamInput, 150, Trim(txtNamaKecelakaan.Text))
        .Parameters.Append .CreateParameter("TglKecelakaan", adDate, adParamInput, , Format(dtpTglKecelakaan.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TempatKecelakaan", adVarChar, adParamInput, 150, IIf(Len(Trim(txtTempatKecelakaan.Text)) = 0, Null, Trim(txtTempatKecelakaan.Text)))
        .Parameters.Append .CreateParameter("IdPegawai", adVarChar, adParamInput, 10, dcPerawat.BoundText)
        .Parameters.Append .CreateParameter("IdUser", adVarChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AU_KasusKecelakaan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("AU_KasusKecelakaan")
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    cmdSimpan.Enabled = False
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan data kecelakaan", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub dcPerawat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglKecelakaan.SetFocus
End Sub

Private Sub dcPerawat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcPerawat.MatchedWithList = True Then dtpTglKecelakaan.SetFocus
        strSQL = " SELECT IdPegawai, [Nama Pemeriksa]" & _
        " From V_DaftarPemeriksaPasien" & _
        " and (Nama Pemeriksa LIKE '%" & dcPerawat.Text & "%')ORDER BY [Nama Pemeriksa] "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcPerawat.Text = ""
            Exit Sub
        End If
        dcPerawat.BoundText = rs(0).value
        dcPerawat.Text = rs(1).value
    End If
End Sub

Private Sub dtpTglKecelakaan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtNamaKecelakaan.SetFocus
End Sub

Private Sub dtpTglPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcPerawat.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglKecelakaan.value = Now
    dtpTglPeriksa.value = Now

    strSQL = "SELECT IdPegawai, [Nama Pemeriksa]" & _
    " From V_DaftarPemeriksaPasien" & _
    " ORDER BY [Nama Pemeriksa]"
    Call msubDcSource(dcPerawat, rs, strSQL)
    If rs.EOF = False Then dcPerawat.BoundText = strIDPegawaiAktif

    txtNoPendaftaran = frmDaftarPasienGD.dgDaftarPasienGD.Columns(0)
    txtNoCM = frmDaftarPasienGD.dgDaftarPasienGD.Columns(1)
    txtNamaPasien = frmDaftarPasienGD.dgDaftarPasienGD.Columns(2)
    If frmDaftarPasienGD.dgDaftarPasienGD.Columns(3).value = "P" Then
        txtSex.Text = "Perempuan"
    Else
        txtSex.Text = "Laki-laki"
    End If
    txtThn = frmDaftarPasienGD.dgDaftarPasienGD.Columns(11)
    txtBln = frmDaftarPasienGD.dgDaftarPasienGD.Columns(12)
    txtHari = frmDaftarPasienGD.dgDaftarPasienGD.Columns(13)
    TxtTglMasuk = frmDaftarPasienGD.dgDaftarPasienGD.Columns(7)
    mstrKdSubInstalasi = frmDaftarPasienGD.dgDaftarPasienGD.Columns("KdSubInstalasi").value

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmTransaksiPasien.Enabled = True
    Call frmTransaksiPasien.subLoadRiwayatKecelakaan
End Sub

Private Sub txtNamaKecelakaan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtTempatKecelakaan.SetFocus
End Sub

Private Sub txtTempatKecelakaan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSimpan.SetFocus
End Sub

