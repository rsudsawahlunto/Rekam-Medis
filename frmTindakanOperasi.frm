VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTindakanOperasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Catatan Tindakan Operasi"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTindakanOperasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   11310
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   31
      Top             =   5880
      Width           =   11295
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   465
         Left            =   7680
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   465
         Left            =   9480
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame framDiagnosa 
      Caption         =   "Data Tindakan Operasi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      TabIndex        =   25
      Top             =   1920
      Width           =   11295
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
         Height          =   1935
         Left            =   6120
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   9495
         Begin MSDataGridLib.DataGrid dgDokter 
            Height          =   1335
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   2355
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
      End
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6120
         MaxLength       =   50
         TabIndex        =   10
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtJenisOperasi 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   9000
         MaxLength       =   50
         TabIndex        =   12
         Top             =   480
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpTglDirujuk 
         Height          =   330
         Left            =   4080
         TabIndex        =   9
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   118882307
         UpDown          =   -1  'True
         CurrentDate     =   38077
      End
      Begin MSComCtl2.DTPicker dtpTglMulai 
         Height          =   330
         Left            =   2040
         TabIndex        =   8
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   118882307
         UpDown          =   -1  'True
         CurrentDate     =   38077
      End
      Begin MSComctlLib.ListView lvwTindakanOperasi 
         Height          =   2535
         Left            =   2040
         TabIndex        =   13
         Top             =   1200
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nama Tindakan Operasi"
            Object.Width           =   13229
         EndProperty
      End
      Begin VB.TextBox txtKdDokter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6840
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Penanggung Jawab"
         Height          =   210
         Left            =   6120
         TabIndex        =   30
         Top             =   240
         Width           =   2220
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Daftar Tindakan Operasi"
         Height          =   210
         Left            =   2040
         TabIndex        =   29
         Top             =   960
         Width           =   1950
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Operasi"
         Height          =   210
         Left            =   9000
         TabIndex        =   28
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Selesai Operasi"
         Height          =   210
         Left            =   4080
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Mulai Operasi"
         Height          =   210
         Left            =   2040
         TabIndex        =   26
         Top             =   240
         Width           =   1425
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
      Height          =   975
      Left            =   0
      TabIndex        =   16
      Top             =   960
      Width           =   11295
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtNoIBS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
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
         Left            =   8640
         TabIndex        =   17
         Top             =   240
         Width           =   2535
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   7
            Top             =   250
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            MaxLength       =   6
            TabIndex        =   6
            Top             =   250
            Width           =   375
         End
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            MaxLength       =   6
            TabIndex        =   5
            Top             =   250
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Left            =   2280
            TabIndex        =   20
            Top             =   300
            Width           =   165
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Left            =   1440
            TabIndex        =   19
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Left            =   600
            TabIndex        =   18
            Top             =   300
            Width           =   285
         End
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7440
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   4320
         MaxLength       =   50
         TabIndex        =   3
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3120
         MaxLength       =   6
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "No. Bedah Sentral"
         Height          =   210
         Left            =   1560
         TabIndex        =   24
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   7440
         TabIndex        =   23
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   4320
         TabIndex        =   22
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   3120
         TabIndex        =   21
         Top             =   240
         Width           =   585
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   35
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
      Left            =   9480
      Picture         =   "frmTindakanOperasi.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmTindakanOperasi.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTindakanOperasi.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmTindakanOperasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilter As String
Dim intJmlDokter As Integer
Dim intJmlTindDipilih As Integer
Dim strKdTindOperasi() As String
Dim i, j As Integer
Dim itemAll As Object

Private Sub cmdSimpan_Click()
    If funcCekValidasi = False Then Exit Sub
    On Error GoTo errSimpan
    strSQL = "DELETE FROM DetailTindakanOperasi WHERE NoIBS='" & txtNoIBS.Text & "'"
    dbConn.Execute strSQL
    strSQL = "DELETE FROM CatatanTindakanOperasi WHERE NoIBS='" & txtNoIBS.Text & "'"
    dbConn.Execute strSQL
    'Insert ke tabel CatatanTindakanOperasi
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoIBS", adChar, adParamInput, 10, txtNoIBS.Text)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("TglOperasi", adDate, adParamInput, , Format(dtpTglMulai.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, txtKdDokter.Text)
        .Parameters.Append .CreateParameter("TglSelesai", adDate, adParamInput, , Format(dtpTglDirujuk.value, "yyyy/MM/dd HH:mm:ss"))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_CatatanTindakanOperasi"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan Catatan Tindakan Operasi", vbCritical, "Validasi"
            Call deleteADOCommandParameters(dbcmd)
            Set dbcmd = Nothing
            Exit Sub

        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing

        'Insert ke tabel DetailTindakanOperasi
        For i = 1 To lvwTindakanOperasi.ListItems.Count
            If lvwTindakanOperasi.ListItems(i).Checked = True Then
                If funcDetailTindakanOperasi(dbcmd, Trim(Right(lvwTindakanOperasi.ListItems(i).Key, Len(lvwTindakanOperasi.ListItems(i).Key) - 1))) = False Then Exit Sub
            End If
        Next i
    End With
    Call Add_HistoryLoginActivity("Add_CatatanTindakanOperasi+Add_DetailTindakanOperasi")
    frmTransaksiPasien.subLoadRiwayatOperasi
    MsgBox "Pemasukan Catatan Tindakan Operasi Pasien Sukses", vbExclamation, "Validasi"
    framDiagnosa.Enabled = False
    cmdSimpan.Enabled = False
    Exit Sub
errSimpan:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan data tindakan operasi", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub dgDokter_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDokter
    WheelHook.WheelHook dgDokter
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        txtDokter.Text = dgDokter.Columns(1).value
        txtKdDokter.Text = dgDokter.Columns(0).value
        If txtKdDokter.Text = "" Then
            MsgBox "Pilih dulu Dokter Penanggungjawab yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        fraDokter.Visible = False
    End If
End Sub

Private Sub dtpTglDirujuk_GotFocus()
    fraDokter.Visible = False
End Sub

Private Sub dtpTglDirujuk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtDokter.SetFocus
End Sub

Private Sub dtpTglMulai_GotFocus()
    fraDokter.Visible = False
End Sub

Private Sub dtpTglMulai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglDirujuk.SetFocus
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    intJmlTindDipilih = 0
    dtpTglDirujuk.value = Now
    subLoadLvw
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmTransaksiPasien.Enabled = True
End Sub

Private Sub lvwTindakanOperasi_GotFocus()
    fraDokter.Visible = False
End Sub

Private Sub lvwTindakanOperasi_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim blnSelected As Boolean
    If Item.Checked = True Then
        intJmlTindDipilih = intJmlTindDipilih + 1
        ReDim Preserve strKdTindOperasi(intJmlTindDipilih)
        strKdTindOperasi(intJmlTindDipilih) = Right(Item.Key, Len(Item.Key) - 1)
    Else
        blnSelected = False
        For i = 1 To intJmlTindDipilih
            If strKdTindOperasi(i) = Right(Item.Key, Len(Item.Key) - 1) Then blnSelected = True
            If blnSelected = True Then
                If i = intJmlTindDipilih Then
                    strKdTindOperasi(i) = ""
                Else
                    strKdTindOperasi(i) = strKdTindOperasi(i + 1)
                End If
            End If
        Next i
        intJmlTindDipilih = intJmlTindDipilih - 1
    End If
End Sub

Private Sub lvwTindakanOperasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtDokter_Change()
    strFilter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
    txtKdDokter.Text = ""
    Call subLoadDokter
End Sub

Private Sub txtDokter_GotFocus()
    fraDokter.Visible = True
    If txtDokter.Text = "" Then strFilter = ""
    Call subLoadDokter
End Sub

Private Sub txtDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then Call txtDokter_KeyPress(13)
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        dgDokter.SetFocus
    End If
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        txtDokter.Text = ""
        txtKdDokter.Text = ""
        fraDokter.Visible = False
    End If
    Call SetKeyPressToChar(KeyAscii)
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter()
    On Error GoTo errLoad
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & strFilter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDokter = rs.RecordCount
    Set dgDokter.DataSource = rs
    With dgDokter
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

'untuk loading data listview tindakan
Private Sub subLoadLvw()
    On Error GoTo errLoad
    strSQL = "SELECT KdTindakanOperasi,NamaTindakanOperasi FROM ListTindakanOperasi WHERE KdJenisOperasi='" & mstrKdJenisOperasi & "' ORDER BY NamaTindakanOperasi"
    msubRecFO rs, strSQL
    lvwTindakanOperasi.ListItems.clear
    Do While rs.EOF = False
        Set itemAll = lvwTindakanOperasi.ListItems.Add(, "A" & rs(0).value, rs(1).value)
        rs.MoveNext
    Loop
    If intJmlTindDipilih = 0 Then Exit Sub
    For i = 1 To lvwTindakanOperasi.ListItems.Count
        For j = 1 To intJmlTindDipilih
            If lvwTindakanOperasi.ListItems(i).Key = strKdTindOperasi(j) Then lvwTindakanOperasi.ListItems(i).Checked = True
        Next j
    Next i
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

'untuk cek validasi
Private Function funcCekValidasi() As Boolean
    If txtKdDokter.Text = "" Then
        MsgBox "Pilihan Dokter harus diisi sesuai data daftar dokter", vbCritical, "Validasi"
        funcCekValidasi = False
        txtDokter.SetFocus
        Exit Function
    End If
    Dim blnChecked As Boolean
    blnChecked = False
    For i = 1 To lvwTindakanOperasi.ListItems.Count
        If lvwTindakanOperasi.ListItems(i).Checked = True Then blnChecked = True
    Next i
    If blnChecked = False Then
        MsgBox "Pilihan Tindakan Operasi harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        lvwTindakanOperasi.SetFocus
        Exit Function
    End If
    funcCekValidasi = True
End Function

'untuk penyimpanan ke tabel DetailTindakanOperasi
Private Function funcDetailTindakanOperasi(adoCommand As ADODB.Command, strKdTindakan As String) As Boolean
    On Error GoTo errSimpan
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoIBS", adChar, adParamInput, 10, txtNoIBS.Text)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("KdTindakanOperasi", adChar, adParamInput, 5, strKdTindakan)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DetailTindakanOperasi"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan Detail Tindakan Operasi", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoCommand)
            Set dbcmd = Nothing
            funcDetailTindakanOperasi = False
            Exit Function

        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    funcDetailTindakanOperasi = True
    Exit Function
errSimpan:
    Call msubPesanError
    funcDetailTindakanOperasi = False
End Function

Private Sub txtJenisOperasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lvwTindakanOperasi.SetFocus
End Sub
