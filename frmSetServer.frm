VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Setting Server & Database"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   14
      Top             =   1800
      Width           =   5175
      Begin VB.TextBox txtServerName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtDatabase 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   3135
      End
      Begin VB.CommandButton cmdTestConnection 
         Caption         =   "T&est Connection"
         Height          =   375
         Left            =   2400
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cbDatabase2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1800
         TabIndex        =   15
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Database"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Server"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         TabIndex        =   17
         Top             =   315
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdLoadReg 
      Caption         =   "&Load dari Registry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdLanjut 
      Caption         =   "&Lanjut"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   7200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   0
      TabIndex        =   7
      Top             =   3480
      Width           =   5175
      Begin VB.TextBox txtuserName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   450
         Width           =   885
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   880
         Width           =   765
      End
   End
   Begin VB.CheckBox chkSQL 
      Caption         =   "SQL Server Authentication"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   4695
   End
   Begin VB.CommandButton cmdSaveReg 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   5160
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   5580
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Not Connected"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
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
   End
   Begin VB.Image Image1 
      Height          =   2025
      Left            =   0
      Picture         =   "frmSetServer.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5205
   End
End
Attribute VB_Name = "frmSetServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbDatabase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSaveReg.SetFocus
End Sub

Private Sub chkSQL_Click()
    If chkSQL.value = 1 Then
        Frame2.Enabled = True
    Else
        Frame2.Enabled = False
        txtuserName.Text = ""
        txtPassword.Text = ""
    End If
End Sub

Private Sub chkSQL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkSQL.value = 1 Then
            txtuserName.SetFocus
        Else
            cmdSaveReg.SetFocus
        End If
    End If
End Sub

Private Sub cmdBatal_Click()
    Unload Me
End Sub

Private Sub cmdConnect_Click()
    On Error GoTo errhandler
    If txtServerName.Text = "" Then
        GoTo errhandler
    End If
    SB.SimpleText = "Connecting..."
    Call Open_ListDB
    Set rsRecordset = New ADODB.recordset
    Set rsRecordset = dbConn.Execute("sp_databases")
    Do Until rsRecordset.EOF
        cbDatabase.AddItem (rsRecordset.Fields("Database_Name"))
        rsRecordset.MoveNext
    Loop
    cmdConnect.Enabled = False
    Exit Sub
errhandler:
    Call MsgBox("Server does not contain SQL Server", vbOKOnly, "Error")
    cmdConnect.Enabled = True
    txtServerName.Enabled = True
    txtServerName.Text = ""
    SB.SimpleText = "Not Connected"
    Exit Sub
End Sub

Private Sub cmdLanjut_Click()
    On Error GoTo errorLanjut
    frmLogin.Show
    Exit Sub
errorLanjut:
End Sub

Private Sub cmdLoadReg_Click()
    On Error Resume Next

    strSQLIdentifikasi = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "SQLIdentifikasi")
    If strSQLIdentifikasi = 0 Then
        strServerName = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Server Name")
        strDatabaseName = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Database Name")

        txtServerName.Text = strServerName
        txtDatabase.Text = strDatabaseName
    Else

        strServerName = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Server Name")
        strDatabaseName = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Database Name")

        txtServerName.Text = strServerName
        txtDatabase.Text = strDatabaseName

        strUserName = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "User Name")
        strPassword = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Password Name")

        txtuserName.Text = strUserName
        txtPassword.Text = strPassword

        chkSQL.value = 1
    End If
End Sub

Private Sub cmdSaveReg_Click()
    On Error GoTo errorLanjut

    If chkSQL.value <> 1 Then
        strServerName = txtServerName.Text
        strDatabaseName = txtDatabase.Text
        strSQLIdentifikasi = 0
        Call CreateKey("HKEY_CURRENT_USER\Software\Medifirst2000")
        Call CreateKey("HKEY_CURRENT_USER\Software\Medifirst2000\Standard")
        Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Server Name", strServerName)
        Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Database Name", strDatabaseName)
        Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "SQLIdentifikasi", strSQLIdentifikasi)
    Else
        strServerName = txtServerName.Text
        strDatabaseName = txtDatabase.Text
        strUserName = txtuserName.Text
        strPassword = txtPassword.Text
        strSQLIdentifikasi = 1
        Call CreateKey("HKEY_CURRENT_USER\Software\Medifirst2000")
        Call CreateKey("HKEY_CURRENT_USER\Software\Medifirst2000\Standard")
        Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Server Name", strServerName)
        Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Database Name", strDatabaseName)
        Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "User Name", strUserName)
        Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Password Name", strPassword)
        Call SetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "SQLIdentifikasi", strSQLIdentifikasi)
    End If

    frmLogin.Show
    Unload Me
    Exit Sub
errorLanjut:
End Sub

Private Sub cmdTestConnection_Click()
    Dim dbConn As New ADODB.connection
    Dim myConSTR As String

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    dbConn.CursorLocation = adUseServer

    'edit to SQL 2005
    myConSTR = "Provider=SQLNCLI10;Integrated Security=SSPI;DataTypeCompatibility=80;Persist Security Info=False;Initial Catalog=" & cbDatabase & ";Data Source=" & txtServerName
    dbConn.Open myConSTR
    Screen.MousePointer = vbDefault
    If Err Then
        MsgBox "SQL Connection Failed: " & Err.Description, vbCritical, "Error.."
        cmdSimpan.Enabled = False
    Else
        MsgBox "SQL Connection Success" & vbCrLf & "Click 'Save Connection' to save setting into registry", vbInformation, "Success.."
        cmdSimpan.Enabled = True
        cmdSimpan.Default = True
        cmdSimpan.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call cmdLoadReg_Click
End Sub

Private Sub txtDatabase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkSQL.SetFocus
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSaveReg.SetFocus
End Sub

Private Sub Open_ListDB()
    Dim ServerName
    Dim i
    Dim cmd As New ADODB.Command
    Dim DB As String

    ServerName = txtServerName.Text
    Set cmd = New ADODB.Command
    Set dbConn = New ADODB.connection

    With dbConn
        .Provider = "MSDASQL;DRIVER={SQL Server};SERVER=" & ServerName & ";trusted_connection=yes;database=" & DB & ""
        .Open
    End With

    cbDatabase.Enabled = True
    SB.SimpleText = "Connected to " & txtServerName.Text & ""
    Exit Sub
End Sub

Private Sub txtDatabaseName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSaveReg.SetFocus
End Sub

Private Sub txtServerName_GotFocus()
    txtServerName.SetFocus
    txtServerName.SelStart = 0
    txtServerName.SelLength = Len(txtServerName.Text)
End Sub

Private Sub txtServerName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDatabase.SetFocus
End Sub

Private Sub txtuserName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtPassword.SetFocus
End Sub
