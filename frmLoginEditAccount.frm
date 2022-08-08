VERSION 5.00
Begin VB.Form frmLoginEditAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Ganti Kata Kunci"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "frmLoginEditAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5670
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   4440
      Width           =   5655
      Begin VB.CommandButton cmdRubah 
         Caption         =   "&Ubah"
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
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Tutu&p"
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
         Left            =   4080
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   5
      Top             =   3000
      Width           =   5655
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
         Left            =   2640
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtUser 
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
         Left            =   2640
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtPassword2 
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
         Left            =   2640
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         Caption         =   "Password"
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
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         Caption         =   "Nama User"
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
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   870
      End
      Begin VB.Label lblPassword2 
         AutoSize        =   -1  'True
         Caption         =   "Ketik Password Sekali Lagi"
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
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   2115
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   2160
      Width           =   5655
      Begin VB.Label lblNamaPegawai 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pegawai"
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
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label lblNama 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2640
         TabIndex        =   10
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Image Image1 
      Height          =   2265
      Left            =   0
      Picture         =   "frmLoginEditAccount.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5685
   End
End
Attribute VB_Name = "frmLoginEditAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoCommand As New ADODB.Command

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRubah_Click()
    strUser = Trim(txtUser)
    intLenUser = Len(strUser)
    strPass = txtPassword
    strPass2 = txtPassword2
    strPassEn = Crypt(txtPassword.Text)

    If intLenUser = 0 Then
        MsgBox "User tidak boleh dikosongkan", vbCritical, "User kosong"
        txtUser.SetFocus
        Exit Sub
    End If

    If strPass <> strPass2 Then
        MsgBox "Dua password yang anda masukkan tidak sama", vbCritical, "Password tidak sama"
        txtPassword.SetFocus
        txtPassword = ""
        txtPassword2 = ""
        Exit Sub
    End If

    Set rsLoginCompare = Nothing
    strQuery = "SELECT * from Login WHERE (Username = '" & strUser & "')"
    adoCommand.CommandText = strQuery
    adoCommand.CommandType = adCmdText
    adoCommand.Execute
    Set rsLoginCompare.Source = adoCommand
    rsLoginCompare.Open

    'Belum ada username dengan nama tersebut, nama tersebut boleh dipakai sebagai username
    If rsLoginCompare.RecordCount = 0 Then GoTo OldUser
    If rsLoginCompare!idpegawai = strIDPegawai Then
        GoTo OldUser
    Else
        MsgBox "Username sudah ada, pilih username yang lain", vbCritical, "Username error"
        txtUser.SetFocus
        txtUser = ""
    End If
    Exit Sub

OldUser:
    strQuery = "UPDATE Login SET IdPegawai ='" & _
    strIDPegawai & "', UserName =cast('" & _
    strUser & "' as varbinary),Password =cast('" & strPassEn & _
    "' as varbinary) WHERE (IdPegawai = '" & strIDPegawai & "')"
    adoCommand.CommandText = strQuery
    adoCommand.CommandType = adCmdText
    adoCommand.Execute
    Unload Me
    Exit Sub

EmptyName:
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    openConnection
    adoCommand.ActiveConnection = dbConn
    strQuery = "SELECT IdPegawai, NamaLengkap FROM dataPegawai WHERE (IdPegawai = '" & Trim(strIDPegawaiAktif) & "')"
    adoCommand.CommandText = strQuery
    adoCommand.CommandType = adCmdText
    Set rsPegawai.Source = adoCommand
    rsPegawai.Open
    'check recordset
    If rsPegawai.RecordCount = 0 Then Exit Sub
    lblNama = rsPegawai!NamaLengkap
    rsPegawai.Close

    Set rslogin = Nothing
    strQuery = "SELECT IdPegawai, cast(Username as varchar)as Username , cast(Password as varchar)as Password, Status, KdKategoryUser from Login WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
    adoCommand.CommandText = strQuery
    adoCommand.CommandType = adCmdText
    adoCommand.Execute
    Set rslogin.Source = adoCommand
    rslogin.Open

    txtUser = rslogin!UserName
    txtPassword = Crypt(rslogin!Password)
    txtPassword2 = Crypt(rslogin!Password)
    strUserLama = txtUser
End Sub

Private Sub SetNothing()
    Set dbConn = Nothing
    Set adoCommand = Nothing
    Set rsPegawai = Nothing
    Set rslogin = Nothing
    Set rsLoginCompare = Nothing
End Sub

Private Sub txtPassword_Click()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtPassword_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPassword2_Click()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtPassword2_GotFocus()
'    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtPassword2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdRubah_Click
End Sub

Private Sub txtUser_Click()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtUser_GotFocus()
'    SendKeys "{Home}+{End}"
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

