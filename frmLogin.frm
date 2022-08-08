VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Login "
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSqlIdentifikasi 
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtUserName 
      Height          =   495
      Left            =   1320
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtPwd 
      Height          =   495
      Left            =   1320
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtServerName 
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDatabaseName 
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   5175
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   240
         Top             =   1560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton cmdBatal 
         Cancel          =   -1  'True
         Caption         =   "&Batal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtUserID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   2
         Tag             =   "*"
         Top             =   690
         Width           =   3015
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Lanjutkan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   360
         Left            =   1800
         TabIndex        =   3
         Top             =   1020
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ruangan :"
         Height          =   210
         Left            =   840
         TabIndex        =   10
         Top             =   1110
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nama Pemakai :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   405
         TabIndex        =   7
         Top             =   420
         Width           =   1290
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Kata Kunci :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   6
         Top             =   750
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   1905
      Left            =   0
      Picture         =   "frmLogin.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5205
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBatal_Click()
    End
End Sub

Private Sub cmdOK_Click()
    Dim adoCommand As New ADODB.Command

    If Periksa("datacombo", dcRuangan, "Nama ruangan kosong") = False Then Exit Sub

    mstrKdRuangan = dcRuangan.BoundText
    mstrNamaRuangan = dcRuangan.Text
    mstrKdRuanganLogin = dcRuangan.BoundText
    Set rs = Nothing
    rs.Open "Select NamaRS,Alamat,KotaKodyaKab,KodePos,Telepon,NamaFileLogoRS, Website, Email, KelasRS, KetKelasRS, KdRS from ProfilRS", dbConn, adOpenStatic, adLockReadOnly
    On Error Resume Next
    mstrKdInstalasiNonMedis = "05"

    strNNamaRS = rs(0).value
    strNAlamatRS = rs(1).value
    strNKotaRS = rs(2).value
    strNKodepos = rs(3).value
    strNTeleponRS = rs(4).value
    strNamaFileLogoRS = rs(5).value
    strWebsite = rs(6).value
    strEmail = rs(7).value
    strKelasRS = rs(8).value
    strKetKelasRS = rs(9).value

    Set rs = Nothing

    strUser = txtUserID.Text
    strPass = txtPassword.Text
    strQuery = "SELECT IdPegawai, cast(Username as varchar)as Username , cast(Password as varchar)as Password, Status, KdKategoryUser from Login"
    Set rslogin = Nothing
    With rslogin
        adoCommand.ActiveConnection = dbConn
        adoCommand.CommandText = strQuery
        adoCommand.CommandType = adCmdText
        Set .Source = adoCommand
        .Open
        'check recordset
        If rslogin.RecordCount = 0 Then Exit Sub
    End With

    rslogin.MoveFirst

    Do While rslogin.EOF = False
        If UCase(strUser) = UCase(rslogin!UserName) And strPass = Crypt(rslogin!Password) Then
            strIDPegawaiAktif = rslogin!idpegawai
            strIDPegawai = rslogin!idpegawai

            If UCase(strUser) = "ADMIN" Then
                mblnAdmin = True
            Else
                mblnAdmin = False
            End If

            strQuery = "SELECT * FROM LoginAplikasi WHERE IdPegawai = '" & strIDPegawai & "'"
            Set rsLoginApp = Nothing
            With rsLoginApp
                adoCommand.ActiveConnection = dbConn
                adoCommand.CommandText = strQuery
                adoCommand.CommandType = adCmdText
                Set .Source = adoCommand
                .Open
                'check recordset
                If rsLoginApp.RecordCount = 0 Then
                    MsgBox "Anda tidak mempunyai akses untuk membuka aplikasi ini", vbCritical, "Aplikasi Error"
                    Exit Sub
                End If
            End With
            rsLoginApp.MoveFirst
            Do While rsLoginApp.EOF = False
                'Untuk Aplikasi ganti sesuai keperluan
                '**************************************
                If rsLoginApp!KdAplikasi = "020" Then GoTo UserPermited
                '**************************************
                rsLoginApp.MoveNext
            Loop
            MsgBox "Anda tidak mempunyai akses untuk membuka aplikasi ini", vbCritical, "Aplikasi Error"
            Exit Sub

UserPermited:
            strPassEn = Crypt(txtPassword)
            strQuery = "UPDATE Login SET IdPegawai ='" & _
            strIDPegawai & "', UserName ='" & _
            strUser & "',Password ='" & strPassEn & _
            "',Status = '1' WHERE (IdPegawai = '" & strIDPegawai & "')"
            adoCommand.CommandText = strQuery
            adoCommand.CommandType = adCmdText
            adoCommand.Execute

            strSQL = "SELECT NamaLengkap FROM DataPegawai WHERE idPegawai ='" & rslogin!idpegawai & "' AND kdJenisPegawai ='004'" 'staf SIMRS
            Set rs = Nothing
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount = 0 Then
                boolStafSIMRS = False
            Else
                boolStafSIMRS = True
            End If

            'form utama ganti sesuai keperluan
            '**************************************
            Call GetIdPegawai
            UserID = noidpegawai
            strSQL = "SELECT KdKategoryUser FROM Login WHERE IdPegawai='" & UserID & "'"
            Set rs = Nothing
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs(0).value = "01" Then
                mblnAdmin = True
            Else
                mblnAdmin = False
            End If

            strNamaHostLocal = Winsock1.LocalHostName
            strKdAplikasi = "020"
            dTglLogin = Now
            Call subSp_HistoryLoginAplikasi("A")

            MDIUtama.Show
            Unload Me
            Exit Sub
        End If
        rslogin.MoveNext
    Loop
    MsgBox "Anda salah memasukkan username atau password", vbCritical, "Salah user/password"
End Sub

Private Sub dcRuangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcRuangan.MatchedWithList = True Then cmdOK.SetFocus
        strSQL = "select KdRuangan, NamaRuangan from V_LoginAplikasiRekamMedis WHERE (NamaRuangan LIKE '%" & dcRuangan.Text & "%') order by KdRuangan ASC"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcRuangan.BoundText = rs(0).value
        dcRuangan.Text = rs(1).value
    End If
End Sub

Private Sub Form_Load()
    Dim adoCommand As New ADODB.Command
    On Error GoTo errLogin

    strServerName = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Server Name")
    strDatabaseName = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Database Name")
    strUserName = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "User Name")
    strPassword = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Password Name")
    strSQLIdentifikasi = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "SQLIdentifikasi")
    txtServerName.Text = strServerName
    txtDatabaseName.Text = strDatabaseName
    txtuserName.Text = strUserName
    txtPwd.Text = strPassword
    txtSqlIdentifikasi.Text = strSQLIdentifikasi
    strServerName = txtServerName.Text
    strDatabaseName = txtDatabaseName.Text
    strUserName = txtuserName.Text
    strPassword = txtPwd.Text
    strSQLIdentifikasi = txtSqlIdentifikasi.Text

    If txtServerName.Text = "Error" Then
        MsgBox "Tidak ada nama server"
        frmSetServer.Show
        Unload Me
        Exit Sub
    End If
    Set dbConn = Nothing
    openConnection
    If blnError = True Then Exit Sub

    strSQL = "select KdRuangan, NamaRuangan FROM V_LoginAplikasiRekamMedis Order by KdRuangan ASC"
    Call msubDcSource(dcRuangan, rs, strSQL)

    strSQL = "select StatusAntrian from SettingDataUmum"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Or IsNull(rs.Fields("StatusAntrian")) Or rs.Fields("StatusAntrian") = 0 Then
        bolAntrian = False
    Else
        bolAntrian = True
    End If
    Exit Sub
errLogin:
    MsgBox Err.Number & " - " & Err.Description
End Sub


Private Sub Image1_DblClick()
    Unload Me
    frmSetServer.Show
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Dim StrValid As String
    'IDEM atau hampir sama dgn txtUserID_KeyPress
    StrValid = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789"
    If KeyAscii = 13 Then
        dcRuangan.SetFocus
    ElseIf KeyAscii = vbKeyBack Then
        Exit Sub
    ElseIf KeyAscii = vbKeyDelete Then
        Exit Sub
    End If
    If InStr(StrValid, Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeySpace Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
    'Periksa setiap karakter yg diketikkan ke kotak UserID
    Dim StrValid As String
    'Ini adalah string yg diperbolehkan utk diinput
    'Anda bisa saja menggantinya ssd keinginan Anda
    StrValid = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789"
    If KeyAscii = 13 Then 'Jika ditekan Enter pd keyboard
        txtPassword.SetFocus   'pindahkan kursor ke txtPassword
        '     SendKeys "{Home}+{End}" 'Highlight teks kalau sudah ada
    End If
    If InStr(StrValid, Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeySpace Then
        KeyAscii = 0  'Jika diinput karakter yg tdk valid, diam saja
    End If
End Sub

Private Sub SetNothing()
    Set dbConn = Nothing
    Set rslogin = Nothing
    Set rsLoginApp = Nothing
    Call openConnection
    rslogin.Open "SELECT dataPegawai.NamaLengkap FROM Login INNER JOIN dataPegawai ON Login.IdPegawai = dataPegawai.IdPegawai where dataPegawai.IdPegawai ='" & strIDPegawai & "'", dbConn, adOpenStatic, adLockOptimistic
    If rslogin.RecordCount = 0 Then
        MDIUtama.StatusBar1.Panels(1).Text = " "
    Else
        MDIUtama.StatusBar1.Panels(1).Text = rslogin(0).value
    End If
    Set dbConn = Nothing
    Set rslogin = Nothing
End Sub

