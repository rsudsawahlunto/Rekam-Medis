VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDaftarAntrianPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Antrian Pasien"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarAntrianPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   14295
   Begin VB.Frame fraCari 
      Caption         =   "Cari Data Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   0
      TabIndex        =   11
      Top             =   7200
      Width           =   14295
      Begin VB.CommandButton cmdHapusRegistrasi 
         Caption         =   "&Hapus Data"
         Enabled         =   0   'False
         Height          =   450
         Left            =   5040
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdPasienDirujuk 
         Caption         =   "&Masuk Poliklinik"
         Height          =   450
         Left            =   6855
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdUbahJenisPasien 
         Caption         =   "&Ubah Jenis Pasien"
         Height          =   450
         Left            =   8670
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdUpdateRegistrasi 
         Caption         =   "Ubah &Registrasi"
         Height          =   450
         Left            =   10485
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   600
         TabIndex        =   5
         Top             =   400
         Width           =   3735
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   450
         Left            =   12300
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukan Nama Pasien /  No.CM / Ruangan"
         Height          =   240
         Index           =   0
         Left            =   600
         TabIndex        =   12
         Top             =   165
         Width           =   3675
      End
   End
   Begin VB.Frame fraDaftar 
      Caption         =   "Daftar Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   0
      TabIndex        =   13
      Top             =   960
      Width           =   14295
      Begin VB.Frame fraCetakLabel 
         Caption         =   "Jumlah Baris Label"
         Height          =   1335
         Left            =   11400
         TabIndex        =   20
         Top             =   4680
         Visible         =   0   'False
         Width           =   2655
         Begin VB.CommandButton cmdCetakLabel 
            Caption         =   "Cetak"
            Height          =   375
            Left            =   1560
            TabIndex        =   24
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Batal"
            Height          =   375
            Left            =   1560
            TabIndex        =   23
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtJml 
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Text            =   "1"
            Top             =   480
            Width           =   975
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   375
            Left            =   1200
            TabIndex        =   22
            Top             =   480
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Value           =   1
            Max             =   100
            Min             =   1
            Enabled         =   -1  'True
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8415
         TabIndex        =   14
         Top             =   165
         Width           =   5775
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   156565507
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   2
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   156565507
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   15
            Top             =   360
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgDaftarAntrianPasien 
         Height          =   5175
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   9128
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   19
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
            Size            =   9.75
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
      Begin MSDataListLib.DataCombo dcStatusPeriksa 
         Height          =   360
         Left            =   6480
         TabIndex        =   0
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.Label LblJumData 
         AutoSize        =   -1  'True
         Caption         =   "Data 0 / 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Status Periksa"
         Height          =   240
         Index           =   1
         Left            =   6480
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   8070
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4154
            Text            =   "Cetak Label (F1)"
            TextSave        =   "Cetak Label (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4154
            Text            =   "Ubah Data Pasien (Shift + F4)"
            TextSave        =   "Ubah Data Pasien (Shift + F4)"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4154
            Text            =   "Cetak Daftar Antrian Pasien (Shift+F6)"
            TextSave        =   "Cetak Daftar Antrian Pasien (Shift+F6)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4154
            Text            =   "Refresh Data (F5)"
            TextSave        =   "Refresh Data (F5)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4154
            Text            =   "Cetak SEP (F9)"
            TextSave        =   "Cetak SEP (F9)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4154
            Text            =   "Cetak Label2 (F11)"
            TextSave        =   "Cetak Label2 (F11)"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      Left            =   12480
      Picture         =   "frmDaftarAntrianPasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarAntrianPasien.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarAntrianPasien.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmDaftarAntrianPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intJumlahPrint As Integer
Dim tempPembayaranKe As Integer
Dim printLabel As Printer
Dim Barcode39 As clsBarCode39
Dim X As String



Public Sub cmdCari_Click()
    On Error GoTo errLoad
    lblJumData.Caption = "Data 0 / 0"
    If dtpAwal.Day <> dtpAkhir.Day Or dtpAwal.Month <> dtpAkhir.Month Or dtpAwal.Year <> dtpAkhir.Year Then
        strSQL = "select top 100 * " & _
        " from V_DaftarAntrianPasienMRS_IRM " & _
        " where ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR Ruangan like '%" & txtParameter.Text & "%') and TglMasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "' and  [Status Periksa] = '" & dcStatusPeriksa.Text & "'" & _
        " order by Ruangan, TglMasuk, [No. Urut]"
    Else
        strSQL = "select * " & _
        " from V_DaftarAntrianPasienMRS_IRM " & _
        " where ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR Ruangan like '%" & txtParameter.Text & "%') and TglMasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "' and  [Status Periksa] = '" & dcStatusPeriksa.Text & "'" & _
        " order by Ruangan, TglMasuk, [No. Urut]"
    End If
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic

    Set dgDaftarAntrianPasien.DataSource = rs
    Call SetGridAntrianPasien
    lblJumData.Caption = "Data 0 / " & dgDaftarAntrianPasien.ApproxCount
    
    If dcStatusPeriksa.Text = "Sudah" Then
        cmdUbahJenisPasien.Enabled = False
        cmdUpdateRegistrasi.Enabled = False
        cmdPasienDirujuk.Enabled = False
    Else
        cmdUbahJenisPasien.Enabled = True
        cmdUpdateRegistrasi.Enabled = True
        cmdPasienDirujuk.Enabled = True
    End If
    
    
    If dgDaftarAntrianPasien.ApproxCount > 0 Then
        dgDaftarAntrianPasien.SetFocus
    Else
        dcStatusPeriksa.SetFocus
    End If
    Exit Sub
errLoad:
End Sub

Private Sub cmdCetakLabel_Click()
Dim i As Integer

For i = 1 To txtjml.Text
    Call printerLabel
Next i

fraCetakLabel.Visible = False
End Sub

Private Sub cmdHapusRegistrasi_Click()
    On Error GoTo errLoad

    If dgDaftarAntrianPasien.ApproxCount = 0 Then Exit Sub
    If MsgBox("Anda yakin akan menghapus data registrasi pasien", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, dgDaftarAntrianPasien.Columns("No. Registrasi"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 10, dgDaftarAntrianPasien.Columns("KdRuangan"))
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , dgDaftarAntrianPasien.Columns("TglMasuk"))

        .ActiveConnection = dbConn
        .CommandText = "Add_DeleteRegistrasiPasienMRS"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Add_DeleteRegistrasiPasienMRS")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Call cmdCari_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdPasienDirujuk_Click()
    On Error GoTo errLoad
    If dgDaftarAntrianPasien.ApproxCount < 1 Then Exit Sub

    If dgDaftarAntrianPasien.Columns("KdInstalasi").value <> "02" Then
        MsgBox "Proses ini hanya untuk Pasien yang mendaftar ke Ruang Perawatan RJ", vbInformation, "Informasi"
        dgDaftarAntrianPasien.SetFocus
        Exit Sub
    End If
    'validasi status periksa
    If dgDaftarAntrianPasien.Columns("Status Periksa") = "" Then Exit Sub 'status periksa
    If LCase(dgDaftarAntrianPasien.Columns("Status Periksa")) = "sedang" Then
        MsgBox "Pasien sedang dalam proses", vbExclamation, "Validasi"
        Exit Sub
    ElseIf LCase(dgDaftarAntrianPasien.Columns("Status Periksa")) = "sudah" Then
        MsgBox "Pasien sudah selesai diproses", vbExclamation, "Validasi"
        Exit Sub
    End If

    mstrNoPen = dgDaftarAntrianPasien.Columns("No. Registrasi") 'no pendaftaran
    mstrKdRuanganPasien = dgDaftarAntrianPasien.Columns("KdRuangan").value 'Kode Ruangan Pasien
    mstrNamaRuanganPasien = dgDaftarAntrianPasien.Columns("KdRuangan").value 'Kode Ruangan Pasien

    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If

    frmRegistrasiRJ.txtNoPendaftaran = dgDaftarAntrianPasien.Columns("No. Registrasi")
    frmRegistrasiRJ.subTampilData (dgDaftarAntrianPasien.Columns("No. Registrasi"))
    frmRegistrasiRJ.Show
    frmDaftarAntrianPasien.Enabled = False
      strSQL = "select Value from SettingGlobal where Prefix='PathSdkAntrian'"
    Call msubRecFO(rs, strSQL)
      Dim path As String
    If Not rs.EOF Then
        If rs(0).value <> "" Then
            path = rs(0).value
        End If
    End If
    
    strSQL = "select StatusAntrian from SettingDataUmum"
    Call msubRecFO(rs, strSQL)
    Dim coba As Long
    
    If Not rs.EOF Then
        If rs(0).value = "1" Then
            If Dir(path) <> "" Then
                path = path + "endpoint:" & Chr(34) & "net.tcp://192.168.0.11:5556/Queue" & Chr(34) & " Type:" & Chr(34) & "Counting Patient" & Chr(34) & " NoAntrian:" & dgDaftarAntrianPasien.Columns(1).value
                coba = Shell(path, vbNormalFocus)
                Call cmdCari_Click
            End If
        End If
    End If
   
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdUbahJenisPasien_Click()
    On Error GoTo errLoad
    If dgDaftarAntrianPasien.ApproxCount = 0 Then Exit Sub
    Call subLoadFormJP
    Exit Sub
errLoad:
End Sub

Private Sub cmdUpdateRegistrasi_Click()
    On Error GoTo hell

    frmRegistrasiUpdate.txtNoPendaftaran = dgDaftarAntrianPasien.Columns(3)
    Call frmRegistrasiUpdate.txtNoPendaftaran_KeyPress(13)
    frmRegistrasiUpdate.Show
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub Command4_Click()
fraCetakLabel.Visible = False
End Sub

Private Sub dcStatusPeriksa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcStatusPeriksa.MatchedWithList = True Then dtpAwal.SetFocus
        strSQL = "Select kdstatusperiksa, statusperiksa From StatusPeriksaPasien Where StatusEnabled='1' and (statusperiksa LIKE '%" & dcStatusPeriksa.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcStatusPeriksa.Text = ""
            Exit Sub
        End If
        dcStatusPeriksa.BoundText = rs(0).value
        dcStatusPeriksa.Text = rs(1).value
    End If
End Sub

Private Sub dgDaftarAntrianPasien_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDaftarAntrianPasien
    WheelHook.WheelHook dgDaftarAntrianPasien
    
    fraCetakLabel.Visible = False
End Sub

Private Sub dgDaftarAntrianPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdUpdateRegistrasi.SetFocus
End Sub

Private Sub dgDaftarAntrianPasien_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    lblJumData.Caption = "Data " & dgDaftarAntrianPasien.Bookmark & " / " & dgDaftarAntrianPasien.ApproxCount
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Activate()
    Call cmdCari_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errLoad
    Dim strShiftKey As String
    strShiftKey = (Shift + vbShiftMask)
    Select Case KeyCode
        Case vbKeyF1
            If dgDaftarAntrianPasien.ApproxCount = 0 Then Exit Sub
            intJumlahPrint = 1
            mstrNoPen = dgDaftarAntrianPasien.Columns(3).value
            mstrKdInstalasi = dgDaftarAntrianPasien.Columns(20).value
            frm_cetak_label_viewer.Show
'            frm_cetak_label_viewer.Cetaklangsung
        Case vbKeyF4
            If strShiftKey = 2 Then
                strPasien = "Lama"
                mstrNoCM = dgDaftarAntrianPasien.Columns(4).value
                frmforUbahRegistrasi.Show
            End If
        
        Case vbKeyF5
            Call cmdCari_Click
        
        Case vbKeyF6
            If strShiftKey = 2 Then
                If dgDaftarAntrianPasien.ApproxCount = 0 Then Exit Sub
                frmCetakDaftarAntrianPasien.Show
            End If
            
        Case vbKeyF9
            mstrNoPen = dgDaftarAntrianPasien.Columns("No. Registrasi")
            strSQL = "select *  from SettingGlobal where Prefix = 'KdKelompokPasienUmum'"

            Call msubRecFO(rsCek, strSQL)
            If rsCek.EOF = False Then
                strSQL1 = "SELECT * FROM PemakaianAsuransi where NoPendaftaran = '" & mstrNoPen & "'"
                Call msubRecFO(rs1, strSQL1)
                mstrNoSJP = rs1("NoSJP")
                If mstrNoSJP = "" Then
                    MsgBox "No SJP kosong", vbExclamation, "Validasi"
                    Exit Sub
                End If
                vLaporan = "view"
                frmViewerSJP.Show
            End If
        Case vbKeyF11
            txtjml.Text = 1
            UpDown1.value = 1
            fraCetakLabel.Visible = True
    End Select
    Exit Sub
errLoad:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)

    mblnFormDaftarAntrian = True

    dtpAwal.value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.value = Now
    dcStatusPeriksa.BoundText = ""

    If mblnAdmin = True Then
        cmdHapusRegistrasi.Enabled = True
    Else
        cmdHapusRegistrasi.Enabled = False
    End If

    Call subLoadDcSource
    Call cmdCari_Click
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad
    strSQL = "Select * From StatusPeriksaPasien Where StatusEnabled='1'"
    Call msubDcSource(dcStatusPeriksa, rs, strSQL)
    If rs.EOF = False Then dcStatusPeriksa.BoundText = rs(0).value
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub SetGridAntrianPasien()
    With dgDaftarAntrianPasien
        .Columns(0).Width = 1500 'ruangan
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1900  'tgl masuk
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 750 'no urut
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Caption = "No. Registrasi"
        .Columns(3).Width = 1200 'no pendaftaran
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Caption = "No .CM"
        .Columns(4).Width = 1500 'no cm
        .Columns(4).Alignment = dbgCenter
        .Columns(5).Width = 2000 'nama pasien
        .Columns(5).Alignment = dbgCenter
        .Columns(6).Width = 400 'jk
        .Columns(6).Alignment = dbgCenter
        .Columns(7).Width = 1700 'umur
        .Columns(7).Alignment = dbgCenter
        .Columns(8).Width = 0 'alamat lengkap
        .Columns(9).Width = 0 'status pasien
        .Columns(10).Width = 1500 'jenis pasien
        .Columns(11).Width = 0 'kd ruangan
        .Columns(12).Width = 0 'kd subinstalasi
        .Columns(13).Width = 0 'kd kelas
        .Columns(14).Width = 0 'umur tahun
        .Columns(15).Width = 0 'umur bulan
        .Columns(16).Width = 0 'umur hari
        .Columns(17).Width = 1000 'kelas
        .Columns(18).Width = 0 'namainstalasi
        .Columns(19).Width = 0 'KdInstalasi
        .Columns(20).Width = 0 'IdPenjamin
        .Columns(21).Width = 0 'TglLahir
        .Columns(22).Caption = "Status Periksa"
        .Columns(22).Width = 700 'status periksa
        .Columns(22).Alignment = dbgCenter
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnFormDaftarAntrian = False
End Sub

Private Sub txtjml_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
        Beep
        KeyAscii = 0
End If
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdCari_Click
        txtParameter.SetFocus
    End If
End Sub

'untuk load data pasien di form ubah jenis pasien
Private Sub subLoadFormJP()

    mstrNoPen = dgDaftarAntrianPasien.Columns("No. Registrasi").value
    mstrNoCM = dgDaftarAntrianPasien.Columns("No .CM").value
    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If

    With frmUbahJenisPasien
        .Show
        .txtNamaFormPengirim.Text = Me.Name
        .txtNoCM.Text = dgDaftarAntrianPasien.Columns("No .CM").value
        .txtNamaPasien.Text = dgDaftarAntrianPasien.Columns("Nama Pasien").value
        If dgDaftarAntrianPasien.Columns("JK").value = "P" Then
            .txtJK.Text = "Perempuan"
        Else
            .txtJK.Text = "Laki-laki"
        End If
        .txtThn.Text = dgDaftarAntrianPasien.Columns("UmurTahun").value
        .txtBln.Text = dgDaftarAntrianPasien.Columns("UmurBulan").value
        .txtHr.Text = dgDaftarAntrianPasien.Columns("UmurHari").value
        .txttglpendaftaran.Text = dgDaftarAntrianPasien.Columns("TglMasuk").value
        .lblNoPendaftaran.Visible = False
        .txtNoPendaftaran.Visible = False
        .dcJenisPasien.BoundText = mstrKdJenisPasien
        .dcPenjamin.BoundText = mstrKdPenjaminPasien
    End With
End Sub

Private Sub printerLabel()
    If dgDaftarAntrianPasien.ApproxCount = 0 Then Exit Sub
    Dim tempPrint As String
    Dim KertasLabel As String
    
    tempPrint = ReadINI("Default Printer", "Printer Label", "", "C:\Setting.ini")
    KertasLabel = ReadINI("Default Printer", "Kertas Label", "", "C:\Setting.ini")
    
    For Each printLabel In Printers
        If Right(printLabel.DeviceName, Len(tempPrint)) = tempPrint Then X = tempPrint: Exit For
    Next
    
    If X = "" Then MsgBox "Printer label tidak terdeteksi, harap periksa lagi", vbInformation, "Informasi": Exit Sub
    
    If printLabel.DeviceName = X Then
        Set Printer = printLabel
    Else
        MsgBox "Printer label tidak terdeteksi, harap periksa lagi", vbInformation, "Informasi"
        Exit Sub
    End If
    
    Printer.Font = "Arial Narrow"
    If KertasLabel = "2" Then
    '    Printer.CurrentY = 0
    '    Printer.FontName = "Tahoma"
        Printer.FontSize = 9
        Printer.FontBold = True
    '
        Printer.CurrentY = 0 + 170
        Printer.CurrentX = 100
        Printer.Print "      RSUD SAWAHLUNTO"
        Printer.CurrentY = 0 + 170
        Printer.CurrentX = 3100
        Printer.Print "      RSUD SAWAHLUNTO"
        
        Set rs3 = Nothing
        strSQL3 = "SELECT SUBSTRING('" & dgDaftarAntrianPasien.Columns(4).value & "',5,2)+'-'+SUBSTRING('" & dgDaftarAntrianPasien.Columns(4).value & "',3,2)+'-'+SUBSTRING('" & dgDaftarAntrianPasien.Columns(4).value & "',1,2)"
        Call msubRecFO(rs3, strSQL3)
        
        Printer.CurrentY = 300 + 170
        Printer.CurrentX = 100
        Printer.FontSize = 9
        Printer.FontBold = True
        Printer.Print rs3(0).value
    
        Printer.CurrentY = 300 + 170
        Printer.CurrentX = 3100
        Printer.Print rs3(0).value
        
        Set Barcode39 = New clsBarCode39
        With Barcode39
            .CurrentX = 0  '400 - 150
            .CurrentY = 400 + 170 'sip ' jarak barcode dari atas ke bawah makin dikit makin ke atas
            
            .NarrowX = 12 'Val(txtNarrowX.Text)
            .BarcodeHeight = 300 'Val(txtHeight.Text)
            .ShowBox = 0
            .Barcode = Right(dgDaftarAntrianPasien.Columns(4).value, 6)
            If .ErrNumber <> 0 Then
                MsgBox "Error: It contain invalid barcode charater", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            .Draw Printer
            
        End With
        
        With Barcode39
            .CurrentX = 3000  '400 - 150
            .CurrentY = 400 + 170 'sip ' jarak barcode dari atas ke bawah makin dikit makin ke atas
            
            .NarrowX = 12 'Val(txtNarrowX.Text)
            .BarcodeHeight = 300 'Val(txtHeight.Text)
            .ShowBox = 0
            .Barcode = Right(dgDaftarAntrianPasien.Columns(4).value, 6)
            If .ErrNumber <> 0 Then
                MsgBox "Error: It contain invalid barcode charater", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            .Draw Printer
            
        End With
        
        Printer.CurrentY = 850 + 170
        Printer.CurrentX = 100
        Printer.FontSize = 8
        Printer.FontBold = False
        Printer.Print dgDaftarAntrianPasien.Columns(5).value & " (" & dgDaftarAntrianPasien.Columns(6).value & ")"
    
        Printer.CurrentY = 850 + 170
        Printer.CurrentX = 3100
        Printer.Print dgDaftarAntrianPasien.Columns(5).value & " (" & dgDaftarAntrianPasien.Columns(6).value & ")"
        
        Set rs1 = Nothing
        strSQL1 = "SELECT TglLahir FROM Pasien WHERE NoCM='" & Right(dgDaftarAntrianPasien.Columns(4).value, 6) & "'"
        Call msubRecFO(rs1, strSQL1)
        
        Printer.CurrentY = 1030 + 170
        Printer.CurrentX = 100
        Printer.Print "Tgl. Lahir " & rs1(0).value & " (" & dgDaftarAntrianPasien.Columns(14).value & " th)"
    
        Printer.CurrentY = 1030 + 170
        Printer.CurrentX = 3100
        Printer.Print "Tgl. Lahir " & rs1(0).value & " (" & dgDaftarAntrianPasien.Columns(14).value & " th)"
        Printer.EndDoc
    Else
        Set rs3 = Nothing
        strSQL3 = "SELECT SUBSTRING('" & dgDaftarAntrianPasien.Columns(4).value & "',5,2)+'-'+SUBSTRING('" & dgDaftarAntrianPasien.Columns(4).value & "',3,2)+'-'+SUBSTRING('" & dgDaftarAntrianPasien.Columns(4).value & "',1,2)"
        Call msubRecFO(rs3, strSQL3)
        
        Printer.CurrentY = 0 + 180
        Printer.CurrentX = 50
        Printer.FontSize = 7
        Printer.FontBold = True
        Printer.Print rs3(0).value
    
        Printer.CurrentY = 0 + 180
        Printer.CurrentX = 2050
        Printer.Print rs3(0).value
        
        Printer.CurrentY = 0 + 180
        Printer.CurrentX = 4050
        Printer.Print rs3(0).value
        
        Printer.CurrentY = 250 + 140
        Printer.CurrentX = 50
        Printer.FontSize = 6
        Printer.FontBold = True
        Printer.Print dgDaftarAntrianPasien.Columns(5).value & " (" & dgDaftarAntrianPasien.Columns(6).value & ")"
    
        Printer.CurrentY = 250 + 140
        Printer.CurrentX = 2050
        Printer.Print dgDaftarAntrianPasien.Columns(5).value & " (" & dgDaftarAntrianPasien.Columns(6).value & ")"
        
        Printer.CurrentY = 250 + 140
        Printer.CurrentX = 4050
        Printer.Print dgDaftarAntrianPasien.Columns(5).value & " (" & dgDaftarAntrianPasien.Columns(6).value & ")"
        
        Set rs1 = Nothing
        strSQL1 = "SELECT TglLahir FROM Pasien WHERE NoCM='" & Right(dgDaftarAntrianPasien.Columns(4).value, 6) & "'"
        Call msubRecFO(rs1, strSQL1)
        
'        Printer.FontBold = False
        Printer.CurrentY = 400 + 140
        Printer.CurrentX = 50
        Printer.Print "Tgl. Lahir " & rs1(0).value & " (" & dgDaftarAntrianPasien.Columns(14).value & " th)"
    
        Printer.CurrentY = 400 + 140
        Printer.CurrentX = 2050
        Printer.Print "Tgl. Lahir " & rs1(0).value & " (" & dgDaftarAntrianPasien.Columns(14).value & " th)"
        
        Printer.CurrentY = 400 + 140
        Printer.CurrentX = 4050
        Printer.Print "Tgl. Lahir " & rs1(0).value & " (" & dgDaftarAntrianPasien.Columns(14).value & " th)"
        
        Set Barcode39 = New clsBarCode39
        
        With Barcode39
            
            .CurrentX = 0  '400 - 150
            .CurrentY = 500 + 100 'sip ' jarak barcode dari atas ke bawah makin dikit makin ke atas
            
            .NarrowX = 12 'Val(txtNarrowX.Text)
            .BarcodeHeight = 150 'Val(txtHeight.Text)
            
            .ShowBox = 0
            .Barcode = Right(dgDaftarAntrianPasien.Columns(4).value, 6)
            If .ErrNumber <> 0 Then
                MsgBox "Error: It contain invalid barcode charater", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            .Draw Printer
            
        End With
        
        With Barcode39
            .CurrentX = 2000  '400 - 150
            .CurrentY = 500 + 100 'sip ' jarak barcode dari atas ke bawah makin dikit makin ke atas
            
            .NarrowX = 12 'Val(txtNarrowX.Text)
            .BarcodeHeight = 150 'Val(txtHeight.Text)
           
            .ShowBox = 0
            .Barcode = Right(dgDaftarAntrianPasien.Columns(4).value, 6)
            If .ErrNumber <> 0 Then
                MsgBox "Error: It contain invalid barcode charater", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            .Draw Printer
            
        End With
        
        With Barcode39
            .CurrentX = 4000  '400 - 150
            .CurrentY = 500 + 100 'sip ' jarak barcode dari atas ke bawah makin dikit makin ke atas
            
            .NarrowX = 12 'Val(txtNarrowX.Text)
            .BarcodeHeight = 150 'Val(txtHeight.Text)
           
            .ShowBox = 0
            .Barcode = Right(dgDaftarAntrianPasien.Columns(4).value, 6)
            If .ErrNumber <> 0 Then
                MsgBox "Error: It contain invalid barcode charater", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            .Draw Printer
            
        End With
        Printer.EndDoc
    End If
End Sub

Private Sub UpDown1_Change()
    txtjml.Text = UpDown1.value
End Sub

