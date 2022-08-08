VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCekKepesertaanBPJSVclaim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Informasi BPJS"
   ClientHeight    =   7935
   ClientLeft      =   6195
   ClientTop       =   4245
   ClientWidth     =   12675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCekKepesertaanBPJSVclaim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   12675
   Begin VB.Frame fraSettingKoneksi 
      Caption         =   "Setting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   3960
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txtUrlBPJS 
         Height          =   435
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   2640
         Width           =   6015
      End
      Begin VB.TextBox txtKodeRS 
         Height          =   435
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   1920
         Width           =   6015
      End
      Begin VB.TextBox txtPasswordKey 
         Height          =   435
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   1200
         Width           =   6015
      End
      Begin VB.TextBox txtConsumerID 
         Height          =   435
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   480
         Width           =   6015
      End
      Begin VB.CommandButton cmdSimpanSetting 
         Caption         =   "Simpan"
         Height          =   375
         Left            =   3240
         TabIndex        =   23
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton cmdTutupSetting 
         Caption         =   "Tutup"
         Height          =   375
         Left            =   4680
         TabIndex        =   22
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "URL Server BPJS"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2400
         Width           =   4575
      End
      Begin VB.Label Label4 
         Caption         =   "Kode Rumah Sakit"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   4575
      End
      Begin VB.Label Label3 
         Caption         =   "Password Key"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Consumer ID"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame fraHapusSEP 
      Caption         =   "Hapus SEP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3960
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdBatalKetHapus 
         Caption         =   "Batal"
         Height          =   375
         Left            =   4800
         TabIndex        =   16
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdSimpanKetHapus 
         Caption         =   "Simpan"
         Height          =   375
         Left            =   3360
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtKetHapus 
         Height          =   735
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   480
         Width           =   6015
      End
      Begin VB.Label Label2 
         Caption         =   "Silakan masukkan penyebab hapus SEP"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pilih Jenis"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3735
      Begin VB.OptionButton optCekNoRujukanbyNoKartu 
         Caption         =   "Cek Rujukan by No. Kartu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         Width           =   3135
      End
      Begin VB.OptionButton optSimpanNoPendaftaran 
         Caption         =   "Simpan No.Pendaftaran Ke Service Lokal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   35
         Top             =   5400
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdSetting 
         BackColor       =   &H000000FF&
         Caption         =   "Setting Koneksi"
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   6120
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.OptionButton optRiwayatSEP 
         Caption         =   "Riwayat SEP Peserta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   5040
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.OptionButton optCekNoRujukan 
         Caption         =   "Cek No. Rujukan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   3135
      End
      Begin VB.OptionButton optUpdateTglPulang 
         Caption         =   "Update Tanggal Pulang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   4680
         Width           =   2535
      End
      Begin VB.OptionButton optHapusSEP 
         Caption         =   "Hapus SEP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4320
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton optCekNoSEP 
         Caption         =   "Cek No. SEP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   1815
      End
      Begin VB.OptionButton optBPJS 
         Caption         =   "Cek Peserta by No. Kartu BPJS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3255
      End
      Begin VB.OptionButton optNIK 
         Caption         =   "Cek Peserta by NIK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.TextBox txtNoBPJS 
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   12720
      TabIndex        =   1
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
   Begin VB.Frame Frame3 
      Caption         =   "Parameter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      TabIndex        =   6
      Top             =   1200
      Width           =   8655
      Begin VB.CommandButton cmdValidasi 
         BackColor       =   &H0000FF00&
         Caption         =   "&Proses"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtParameter 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   4455
      End
      Begin MSComCtl2.DTPicker dtpTglPulang 
         Height          =   315
         Left            =   4680
         TabIndex        =   10
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
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
         Format          =   127664131
         UpDown          =   -1  'True
         CurrentDate     =   37694
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Pulang"
         Height          =   210
         Index           =   0
         Left            =   4680
         TabIndex        =   12
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kartu Peserta"
         Height          =   210
         Index           =   30
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5535
      Left            =   3960
      TabIndex        =   5
      Top             =   2280
      Width           =   8655
      Begin MSFlexGridLib.MSFlexGrid fgPeserta 
         Height          =   5175
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   9128
         _Version        =   393216
         AllowUserResizing=   3
      End
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmCekKepesertaanBPJSVclaim.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   10440
      Picture         =   "frmCekKepesertaanBPJSVclaim.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2235
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmCekKepesertaanBPJSVclaim.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14415
   End
   Begin VB.Menu mnKlikKanan 
      Caption         =   "Menu Klik Kanan"
      Visible         =   0   'False
      Begin VB.Menu mnCopy 
         Caption         =   "Copy"
      End
   End
End
Attribute VB_Name = "frmCekKepesertaanBPJSVclaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnKartuAktif As Boolean
Dim HanyaLewat As Boolean
Dim strHapusNoSEP As String
'---------- detailKartuBPJS------------------
Dim noKartu As String
Dim nik As String
Dim nama As String
Dim pisa As String
Dim sex As String
Dim tgllahir As String
Dim tglCetakKartu As String
Dim kdProvider As String
Dim nmProvider As String
Dim kdCabang As String
Dim nmCabang As String
Dim kdJenisPeserta As String
Dim nmJenisPeserta As String
Dim kdKelas As String
Dim nmKelas As String
'----------------------------
Dim statusPeserta As String

Private Sub cmdBatalKetHapus_Click()
    fraHapusSEP.Visible = False
    strHapusNoSEP = ""
End Sub

Private Sub cmdSetting_Click()
    fraSettingKoneksi.Visible = True
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    txtConsumerID.SetFocus
    
    strSQL3 = "Select Value From SettingGlobal where Prefix In('ConsumerID','PasswordKey','KodeRS','UrlGenerateSEP')"
    Call msubRecFO(rs3, strSQL3)
    
    If rs3.EOF = False Then
        txtConsumerID.Text = rs3(0).value
        rs3.MoveNext
        txtKodeRS.Text = rs3(0).value
        rs3.MoveNext
        txtPasswordKey.Text = rs3(0).value
        rs3.MoveNext
        txtUrlBPJS.Text = rs3(0).value
    End If
    
End Sub

Private Sub cmdSimpanKetHapus_Click()
On Error Resume Next
    If Periksa("text", txtKetHapus, "") = False Then
        MsgBox "Silakan isi keterangan penyebab hapus SEP", vbCritical, "Hapus SEP"
        Exit Sub
    End If

    If (Dir("C:\SDK\vclaim\result.tlb") <> "") Then
        Dim context As ContextVclaim
        Dim sep() As String
        Set context = New ContextVclaim
        strSQL = "Select Value From SettingGlobal where Prefix In('ConsumerID','PasswordKey')"
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = False Then
            context.ConsumerID = rs(0).value
            rs.MoveNext
            context.PasswordKey = rs(0).value
        End If
        
        strSQL = "SELECT Value FROM SettingGlobal where Prefix='UrlGenerateSEP'"
        Call msubRecFO(rs, strSQL)
        Dim URL  As String
        If rs.EOF = False Then
            URL = rs.Fields(0)
            context.URL = URL
        End If
        
        sep = context.DeleteSep(strHapusNoSEP, strIDPegawaiAktif)

        If InStr(1, UCase(sep(0)), "GAGAL") = 0 And InStr(1, UCase(sep(0)), "ERROR") = 0 Then
            strSQL = "select NoPendaftaran,TglSJP,NoKartuPeserta from V_CetakSuratJaminanPelayanan  where NoSJP='" & strHapusNoSEP & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
                'Untuk mendapatkan PPKRujukan
                strSQL1 = "select kdProvider from DetailKartuBPJS  where noKartu='" & rs.Fields("NoKartuPeserta") & "'"
                Call msubRecFO(rs1, strSQL1)
    
                'Untuk mendapatkan KdLakaLantas
                strSQL2 = "select KdLakaLantas from PemakaianAsuransiCatatan  where NoSJP='" & strHapusNoSEP & "'"
                Call msubRecFO(rs2, strSQL2)
    
                dbConn.Execute "INSERT INTO dbo.DaftarHapusSEP( NoPendaftaran ,NoSJP ,TglSJP ,KdLakaLantas ,PPKRujukan ,TglActivity,Keterangan,IDPegawai) " & _
                "VALUES  ( '" & rs.Fields("NoPendaftaran") & "' , '" & strHapusNoSEP & "','" & Format(rs.Fields("TglSJP"), "yyyy-MM-dd hh:mm:ss") & "' , " & IIf(IsNull(rs2.Fields("KdLakaLantas")), "-", rs2.Fields("KdLakaLantas")) & " , '" & IIf(IsNull(rs1.Fields("kdProvider")), "-", rs1.Fields("kdProvider")) & "' , GETDATE(),'" & Trim(txtKetHapus.Text) & "','" & strIDPegawaiAktif & "' )"
            
            End If
            
            strSQL = "UPDATE dbo.PemakaianAsuransi SET NoSJP='-' WHERE NoSJP='" & strHapusNoSEP & "'"
            Call msubRecFO(rs, strSQL)
            
            MsgBox "Hapus No. SEP Berhasil.! " & vbCrLf & "dengan No. SEP: " & strHapusNoSEP, vbInformation, "Hapus No. SEP"
        Else
            MsgBox "Hapus No. SEP " & vbCrLf & Replace(sep(0), "message:", ""), vbInformation, "Validasi"
            Debug.Print sep(0)
            Exit Sub
    
        End If
    Else
        MsgBox "error", vbInformation, ""
    End If

    fraHapusSEP.Visible = False
End Sub

Private Sub cmdSimpanSetting_Click()
    If Periksa("text", txtConsumerID, "Consumer ID Masih Kosong") = False Then Exit Sub
    If Periksa("text", txtKodeRS, "Kode Rumah Sakit Masih Kosong") = False Then Exit Sub
    If Periksa("text", txtPasswordKey, "Password Key Masih Kosong") = False Then Exit Sub
    If Periksa("text", txtUrlBPJS, "Url Server BPJS Masih Kosong") = False Then Exit Sub
    
    dbConn.Execute "UPDATE dbo.SettingGlobal SET Value='" & Trim(txtConsumerID.Text) & "' WHERE Prefix='ConsumerID'"
    dbConn.Execute "UPDATE dbo.SettingGlobal SET Value='" & Trim(txtKodeRS.Text) & "' WHERE Prefix='KodeRS'"
    dbConn.Execute "UPDATE dbo.SettingGlobal SET Value='" & Trim(txtPasswordKey.Text) & "' WHERE Prefix='PasswordKey'"
    dbConn.Execute "UPDATE dbo.SettingGlobal SET Value='" & Trim(txtUrlBPJS.Text) & "' WHERE Prefix='UrlGenerateSEP'"
    
    MsgBox "Koneksi Berhasil Disimpan", vbInformation, ""
    
'    fraSettingKoneksi.Visible = False
'    Frame1.Enabled = True
'    Frame2.Enabled = True
'    Frame3.Enabled = True
End Sub

Private Sub cmdTutupSetting_Click()
    fraSettingKoneksi.Visible = False
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
End Sub

Private Sub cmdValidasi_Click()
On Error GoTo hell
    If Periksa("text", txtParameter, "No. Kartu/No. SEP Masih Kosong") = False Then Exit Sub
    txtNoBPJS.Text = ""
    blnKartuAktif = False
    Screen.MousePointer = vbHourglass
    
    If optBPJS.value = True Then
        Call ValidateKartuPeserta("Kartu BPJS")
        
    ElseIf optNIK.value = True Then
        Call ValidateKartuPeserta("NIK")
        
    ElseIf optCekNoRujukan.value = True Then
        Call CekDataRujukanByNoRujukan(Trim(txtParameter.Text))
        
    ElseIf optCekNoRujukanbyNoKartu.value = True Then
        Call CekDataRujukanByNoRujukan(Trim(txtParameter.Text))
        
    ElseIf optCekNoSEP.value = True Then
        Call CekNoSEPPeserta(Trim(txtParameter.Text))
        
    ElseIf optHapusSEP.value = True Then
        If Len(Trim(txtParameter.Text)) < 5 Then MsgBox "Silakan cek kembali No SEP pasien", vbInformation + vbOKOnly, "Hapus SEP": Exit Sub
        If MsgBox("Apakah anda yakin akan menghapus SEP?", vbQuestion + vbYesNo, "Hapus SEP") = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
        fraHapusSEP.Visible = True
        strHapusNoSEP = Trim(txtParameter.Text)
        txtKetHapus.Text = ""
        txtKetHapus.SetFocus
        
    ElseIf optUpdateTglPulang.value = True Then
        Call UpdateTglPulang(Trim(txtParameter.Text))
        
    End If
    
''    If blnKartuAktif = True Then
''        MsgBox "Pasien tersebut aktif", vbInformation, "Cek Kepesertaan BPJS"
''    End If
    
    Screen.MousePointer = vbDefault
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub ValidateKartuPeserta(strJenisID As String)
On Error GoTo hell

    If (Dir("C:\SDK\vclaim\result.tlb") <> "") Then
        Dim context As ContextVclaim
        Set context = New ContextVclaim
        Dim result() As String
        strSQL = "Select Value From SettingGlobal where Prefix In('ConsumerID','PasswordKey')"
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = False Then
            context.ConsumerID = rs(0).value
            rs.MoveNext
            context.PasswordKey = rs(0).value
        End If
        
        strSQL = "SELECT Value FROM SettingGlobal where Prefix='UrlGenerateSEP'"
        Call msubRecFO(rs, strSQL)
        Dim URL  As String
        If rs.EOF = False Then
            URL = rs.Fields(0)
            context.URL = URL
        End If
              
        If strJenisID = "Kartu BPJS" And optBPJS.value = True Then
            result = context.CariPesertaByNoKartuBpjs(txtParameter.Text, Format(Now, "yyyy-MM-dd"))
        ElseIf strJenisID = "NIK" Then
            result = context.CariPesertaByNik(txtParameter.Text, Format(Now, "yyyy-MM-dd"))
        End If
    End If
    
        Call fillGridWithRiwayatPasienByRow(fgPeserta, result)
        Dim i As Long
        For i = LBound(result) To UBound(result)
            Debug.Print (result(i))
            Dim arr() As String
            arr = Split(result(i), ":")
            Select Case arr(0)
                    Case "error"
                        MsgBox arr(1), vbExclamation, "Cek Kepesertaan BPJS"
                        Exit Sub
                    Case "MR-NOKARTU"
                        blnKartuAktif = True
                        txtNoBPJS.Text = arr(1)
                        noKartu = arr(1)
                        Debug.Print "NoKartu : " & arr(1)
                    Case "MR-NIK"
                         nik = arr(1)
                    Case "MR-NAMA"
                         nama = arr(1)
                    Case "PROVUMUM-NMPROVIDER"
                         nmProvider = arr(1)
                    Case "STATUSPESERTA-TGLLAHIR"
                         tgllahir = arr(1)
                    Case "PROVUMUM-KDPROVIDER"
                         ppkRujukan = arr(1)
                         kdProvider = arr(1)
                    Case "HAKKELAS-KETERANGAN"
                         nmKelas = arr(1)
                    Case "JENISPESERTA-KETERANGAN"
                         nmJenisPeserta = arr(1)
                    Case "JENISPESERTA-KODE"
                         kdJenisPeserta = arr(1)
                    Case "MR-PISA"
                        pisa = arr(1)
                    Case "PROVUMUM-SEX"
                         sex = arr(1)
                    Case "kdCabang"
                         kdCabang = arr(1)
                    Case "nmCabang"
                         nmCabang = arr(1)
                    Case "HAKKELAS-KODE"
                         kdKelas = arr(1)
                    Case "STATUSPESERTA-TGLCETAKKARTU"
                         tglCetakKartu = arr(1)
                    Case "STATUSPESERTA-KETERANGAN"
                        statusPeserta = arr(1)
            End Select
        Next i
        
        If UBound(result) = -1 Then
            blnKartuAktif = False
        Else
            MsgBox statusPeserta, vbInformation, "Cek Kepesertaan BPJS"
            If sp_detailkartubpjs(noKartu, nik, nama, pisa, sex, tgllahir, tglCetakKartu, kdProvider, nmProvider, kdCabang, nmCabang, kdJenisPeserta, nmJenisPeserta, kdKelas, nmKelas) = False Then Exit Sub
        End If
        
Exit Sub
hell:
    MsgBox "Koneksi Bridging Bermasalah", vbCritical, "Validasi"
End Sub

Private Sub fgPeserta_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intCtrlShift As Integer
    intCtrlShift = vbCtrlMask + Shift
    Select Case KeyCode
        Case vbKeyC
            If intCtrlShift = 4 Then
                Clipboard.Clear
                Clipboard.SetText fgPeserta.Clip
            End If
    End Select

End Sub

Private Sub fgPeserta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnKlikKanan
    End If
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    txtNoBPJS.Text = ""
    blnKartuAktif = False
    HanyaLewat = True
    optBPJS.value = True
    HanyaLewat = False
    dtpTglPulang.value = Now
    strHapusNoSEP = ""
'    If strIDPegawaiAktif = "L009000119" Or strIDPegawaiAktif = "8888888888" Then
        cmdSetting.Visible = True
'    End If
         
    'Call WheelHookFlexGrid(Me.hWnd)
    
End Sub

Sub fillGridWithRiwayatPasien(vFG As MSFlexGrid, vResult() As String)
    With vFG
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .cols = 1
        .rows = 2
'        Call subSetGrid
        Dim row As Integer
        Dim col As Integer
        Dim rows As Integer
        Dim cols As Integer
        Dim i As Integer
        row = 1
        For i = 0 To UBound(vResult)
            Dim arrResult() As String
'            Debug.Assert i <> 9
            arrResult = Split(vResult(i), ":")
            col = isHeaderExist(vFG, arrResult(0))
            If col > -1 Then
                'col = isHeaderExist(vFG, arrResult(0))
                
                If .TextMatrix(row, col) <> "" Then 'KALO TEXTMATRIX TARGET SUDAH ADA ISI BERARTI KITA HARUS TAMBAH
                                                    'ROWS/PINDAH KE BARIS SELANJUTNYA
                                                    'diharapkan kolom pertama adalah kolom yang selalu memiliki nilai
                    .rows = .rows + 1
                    row = .rows - 1
                End If
                .TextMatrix(row, col) = arrResult(1)
            Else 'KALAU ADA KOLOM BARU
                col = .cols - 1
                .TextMatrix(0, col) = arrResult(0) 'BERI HEADER BARU
                .TextMatrix(row, col) = arrResult(1)
                .cols = .cols + 1
                
            End If
        Next i
        .cols = .cols - 1 'MENGHILANGKAN KOLOM YG KELEBIHAN
    End With
End Sub

Function isHeaderExist(vFG As MSFlexGrid, strHeader As String) As Integer
    isHeaderExist = -1
    With vFG
        Dim col As Integer
        For col = 0 To vFG.cols - 1
            If UCase(.TextMatrix(0, col)) = UCase(strHeader) Then
                isHeaderExist = col
                Exit Function
            End If
        Next col
    End With
End Function


Sub fillGridWithRiwayatPasienByRow(vFG As MSFlexGrid, vResult() As String)
    With vFG
        .Clear
        .Redraw = False
        .cols = 3
        .rows = UBound(vResult) + 2
        .FixedRows = 1
        .FixedCols = 1
        .ColWidth(0) = 300
        .RowHeight(0) = 300
        
        .ColWidth(1) = 3200
        .ColWidth(2) = 6000

        Dim row As Integer
        Dim col As Integer
        Dim rows As Integer
        Dim cols As Integer
        Dim i As Integer
        Dim j As Integer
        row = 1
        j = 0
        For i = 0 To UBound(vResult)
            Dim arrResult() As String
            arrResult = Split(vResult(i), ":")
            If vResult(i) = "=============" Then j = j + 1: GoTo lewati
            If Left(vResult(i), 2) = ">>" Then
                .TextMatrix(i - j + 1, 1) = arrResult(0)
                GoTo lewati
            End If
            .TextMatrix(i - j + 1, 1) = arrResult(0)
            .TextMatrix(i - j + 1, 2) = arrResult(1)
lewati:
        Next i
        .ColAlignment(1) = flexAlignLeftTop
        .ColAlignment(2) = flexAlignLeftTop
        .Redraw = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call WheelUnHookFlexGrid(Me.hWnd)
End Sub

Private Sub mnCopy_Click()
    Clipboard.Clear
    Clipboard.SetText fgPeserta.TextMatrix(fgPeserta.row, fgPeserta.col)
End Sub

Private Sub optBPJS_Click()
    If HanyaLewat = True Then Exit Sub
    lbl(30).Caption = "No. Kartu Peserta"
    txtParameter.SetFocus
    fraHapusSEP.Visible = False
    Call bersih
End Sub

Private Sub optCekNoRujukan_Click()
    lbl(30).Caption = "No. Rujukan"
    txtParameter.SetFocus
    fraHapusSEP.Visible = False
    Call bersih
End Sub

Private Sub optCekNoRujukanbyNoKartu_Click()
    lbl(30).Caption = "No. Kartu BPJS Peserta"
    txtParameter.SetFocus
    fraHapusSEP.Visible = False
    Call bersih
End Sub

Private Sub optCekNoSEP_Click()
    lbl(30).Caption = "No. SEP Peserta"
    txtParameter.SetFocus
    fraHapusSEP.Visible = False
    Call bersih
End Sub

Private Sub optHapusSEP_Click()
    lbl(30).Caption = "No. SEP Peserta"
    txtParameter.SetFocus
    Call bersih
End Sub



Private Sub optNIK_Click()
    lbl(30).Caption = "NIK Peserta"
    txtParameter.SetFocus
    fraHapusSEP.Visible = False
    Call bersih
End Sub

Private Sub optRiwayatSEP_Click()
    lbl(30).Caption = "No. Kartu BPJS Peserta"
    txtParameter.SetFocus
    fraHapusSEP.Visible = False
End Sub

Private Sub optSimpanNoPendaftaran_Click()
    lbl(30).Caption = "No. Pendaftaran"
    txtParameter.SetFocus
    fraHapusSEP.Visible = False
End Sub



Private Sub optUpdateTglPulang_Click()
    lbl(30).Caption = "No. SEP Peserta"
    txtParameter.SetFocus
    fraHapusSEP.Visible = False
    Call bersih
End Sub

Private Sub txtKetHapus_LostFocus()
    If optHapusSEP.value = False Then fraHapusSEP.Visible = False
End Sub

Private Sub txtNoBPJS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdValidasi.SetFocus
End If
End Sub

Private Sub UpdateTglPulang(strNoSEP As String)
On Error Resume Next

    If Len(Trim(strNoSEP)) < 19 Then
        Call msubPesanError("Silakan verifikasi NoSEP aktif")
        Exit Sub
    End If
    
    If (Dir("C:\SDK\vclaim\result.tlb") <> "") Then
        Dim context As ContextVclaim
        Dim sep() As String
        Set context = New ContextVclaim
        strSQL = "Select Value From SettingGlobal where Prefix In('ConsumerID','PasswordKey')"
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = False Then
            context.ConsumerID = rs(0).value
            rs.MoveNext
            context.PasswordKey = rs(0).value
        End If
        
        strSQL = "SELECT Value FROM SettingGlobal where Prefix='UrlGenerateSEP'"
        Call msubRecFO(rs, strSQL)
        Dim URL  As String
        If rs.EOF = False Then
            URL = rs.Fields(0)
            context.URL = URL
        End If
        
        strSQL = "Select Value From SettingGlobal where Prefix ='KodeRS'"
        Call msubRecFO(rs, strSQL)
        
        sep = context.UpdateTgPulangSep(strNoSEP, Format(dtpTglPulang.value, "yyyy-MM-dd"), strIDPegawaiAktif)
        
        Dim i As Long
        For i = LBound(sep) To UBound(sep)
            Debug.Print (sep(i))
            Dim arr() As String
            arr = Split(sep(i), ":")
            Select Case arr(0)
                Case "error"
                    MsgBox arr(1), vbExclamation, "Update Tanggal Pulang Pasien BPJS"
                    Exit Sub
            End Select
        Next i
        
        MsgBox "Update Tanggal Pulang berhasil" & vbCrLf & sep(0), vbInformation, "Update Tanggal Pulang Pasien BPJS"
        
    Else
        MsgBox "error", vbInformation, ""
    End If
End Sub

Private Sub CekNoSEPPeserta(strNoSEP As String)
On Error Resume Next
    If Len(Trim$(strNoSEP)) < 4 Then
        Call msubPesanError("Silakan verifikasi NoSEP aktif")
        Exit Sub
    End If
    
    If (Dir("C:\SDK\vclaim\result.tlb") <> "") Then
        Dim context As ContextVclaim
        Dim sep() As String
        Set context = New ContextVclaim
        strSQL = "Select Value From SettingGlobal where Prefix In('ConsumerID','PasswordKey')"
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = False Then
            context.ConsumerID = rs(0).value
            rs.MoveNext
            context.PasswordKey = rs(0).value
        End If
        
        strSQL = "SELECT Value FROM SettingGlobal where Prefix='UrlGenerateSEP'"
        Call msubRecFO(rs, strSQL)
        Dim URL  As String
        If rs.EOF = False Then
            URL = rs.Fields(0)
            context.URL = URL
        End If
        
        sep = context.CariSEP(strNoSEP)
        
        Call fillGridWithRiwayatPasienByRow(fgPeserta, sep)
        
        Dim i As Long
        For i = LBound(result) To UBound(result)
            Debug.Print (result(i))
            Dim arr() As String
            arr = Split(sep(i), ":")
            Select Case arr(0)
                Case "error"
                    MsgBox Replace(sep(i), "error:", ""), vbExclamation, "Cari SEP"
                    Exit Sub
            End Select
        Next i
        
    Else
        MsgBox "error", vbInformation, ""
    End If
End Sub

Private Sub CekDataRujukanByNoRujukan(strNoRujukan As String)
On Error Resume Next

    If Len(Trim$(strNoRujukan)) < 4 Then
        Call msubPesanError("Silakan verifikasi No. Rujukan aktif")
        Exit Sub
    End If
    
    If (Dir("C:\SDK\vclaim\result.tlb") <> "") Then
        Dim context As ContextVclaim
        Dim result() As String
        Set context = New ContextVclaim
        strSQL = "Select Value From SettingGlobal where Prefix In('ConsumerID','PasswordKey')"
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = False Then
            context.ConsumerID = rs(0).value
            rs.MoveNext
            context.PasswordKey = rs(0).value
        End If
        
        strSQL = "SELECT Value FROM SettingGlobal where Prefix='UrlGenerateSEP'"
        Call msubRecFO(rs, strSQL)
        Dim URL  As String
        If rs.EOF = False Then
            URL = rs.Fields(0)
            context.URL = URL
        End If
        
        If optCekNoRujukan.value = True Then
            result = context.RujukanPcareByNoRujukan(strNoRujukan)
        ElseIf optCekNoRujukanbyNoKartu.value = True Then
            result = context.RujukanPcareByNoKartu(strNoRujukan)
        End If
        
        Call fillGridWithRiwayatPasienByRow(fgPeserta, result)
        
        Dim i As Long
        For i = LBound(result) To UBound(result)
            Debug.Print (result(i))
            Dim arr() As String
            arr = Split(result(i), ":")
            Select Case arr(0)
                Case "error"
                    MsgBox arr(1), vbExclamation, "Cari Rujukan Pasien BPJS"
                    Exit Sub
            End Select
        Next i
        
    Else
        MsgBox "error", vbInformation, ""
    End If
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    cmdValidasi.SetFocus
End If
End Sub

' MouseWheel FlexGrid===========================================================================
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
'  Dim ctl As Control
'  Dim bHandled As Boolean
'  Dim bOver As Boolean
'
'  For Each ctl In Controls
'    ' Is the mouse over the control
'    On Error Resume Next
'    bOver = (ctl.Visible And IsOver(ctl.hWnd, Xpos, Ypos))
'    On Error GoTo 0
'
'    If bOver Then
'      ' If so, respond accordingly
'      bHandled = True
'      Select Case True
'
'        Case TypeOf ctl Is MSFlexGrid
'          FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
'
'        Case Else
'          bHandled = False
'
'      End Select
'      If bHandled Then Exit Sub
'    End If
'    bOver = False
'  Next ctl
'
'  ' Scroll was not handled by any controls, so treat as a general message send to the form
'  Me.Caption = "Form Scroll " & IIf(Rotation < 0, "Down", "Up")
End Sub

Private Function sp_detailkartubpjs(f_noKartu As String, f_nik As String, f_nama As String, f_pisa As String, f_sex As String, f_TglLahir As String, f_tglCetakKartu As String, f_kdProvider As String, f_nmProvider As String, f_kdCabang As String, f_nmCabang As String, f_kdJenisPeserta As String, f_nmJenisPeserta As String, f_kdKelas As String, f_nmKelas As String)
On Error GoTo errLoad
    sp_detailkartubpjs = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("noKartu", adChar, adParamInput, 13, f_noKartu)
        .Parameters.Append .CreateParameter("nik", adChar, adParamInput, 16, f_nik)
        .Parameters.Append .CreateParameter("nama", adVarChar, adParamInput, 50, f_nama)
        .Parameters.Append .CreateParameter("pisa", adVarChar, adParamInput, 3, f_pisa)
        .Parameters.Append .CreateParameter("sex", adVarChar, adParamInput, 3, f_sex)
        .Parameters.Append .CreateParameter("tglLahir", adVarChar, adParamInput, 20, f_TglLahir)
        .Parameters.Append .CreateParameter("tglCetakKartu", adVarChar, adParamInput, 20, f_tglCetakKartu)
        .Parameters.Append .CreateParameter("kdProvider", adVarChar, adParamInput, 20, f_kdProvider)
        .Parameters.Append .CreateParameter("nmProvider", adVarChar, adParamInput, 50, f_nmProvider)
        .Parameters.Append .CreateParameter("kdCabang", adVarChar, adParamInput, 20, f_kdCabang)
        .Parameters.Append .CreateParameter("nmCabang", adVarChar, adParamInput, 50, f_nmCabang)
        .Parameters.Append .CreateParameter("kdJenisPeserta", adVarChar, adParamInput, 20, f_kdJenisPeserta)
        .Parameters.Append .CreateParameter("nmJenisPeserta", adVarChar, adParamInput, 50, f_nmJenisPeserta)
        .Parameters.Append .CreateParameter("kdKelas", adVarChar, adParamInput, 10, f_kdKelas)
        .Parameters.Append .CreateParameter("nmKelas", adVarChar, adParamInput, 20, f_nmKelas)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_DetailKartuBPJS"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute

        Call Add_HistoryLoginActivity("AUD_detailkartubpjs")
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    sp_detailkartubpjs = False
    Call msubPesanError
End Function

Public Sub bersih()
    txtParameter.Text = ""
End Sub
