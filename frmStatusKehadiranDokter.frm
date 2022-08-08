VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmStatusHadirDokter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status Kehadiran Dokter"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStatusKehadiranDokter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7770
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   0
      TabIndex        =   11
      Top             =   3000
      Width           =   7695
      Begin VB.CommandButton cmdKeluar 
         Caption         =   "K&eluar"
         Height          =   495
         Left            =   5760
         TabIndex        =   16
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   3960
         TabIndex        =   15
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   495
         Left            =   2040
         TabIndex        =   14
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid dgStatusHadir 
         Height          =   2295
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   7695
      Begin VB.CheckBox chkStatusEnabled 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Enabled"
         Height          =   375
         Left            =   6000
         TabIndex        =   10
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txtNamaExternal 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2880
         TabIndex        =   9
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtKdExternal 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtNamaKehadiran 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   5535
      End
      Begin VB.TextBox txtKdStatusHadir 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Nama External"
         Height          =   375
         Left            =   2880
         TabIndex        =   8
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Kode External"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Status Kehadiran"
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Status Hadir"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
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
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmStatusKehadiranDokter.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmStatusHadirDokter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================= Create by Dayz ====================================================
Option Explicit
Dim adoCommand As New ADODB.Command

Private Sub cmdBatal_Click()
On Error Resume Next
    Call Clear
    Call setgrid
Exit Sub
'gabril:
'    Call msubPesanError
End Sub

Private Sub cmdHapus_Click()
On Error GoTo gabril
    If Periksa("text", txtNamaKehadiran, "Silahkan isi Nama Status Kehadiran") = False Then Exit Sub
    If MsgBox("Apakah anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    If sp_StatusHadir("D") = False Then Exit Sub
    
    MsgBox "Data Berhasil Dihapus", vbInformation, "Medifirst2000"
    
    Call cmdBatal_Click
Exit Sub
gabril:
    Call msubPesanError
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo gabril
    If Periksa("text", txtNamaKehadiran, "Silahkan isi Nama Status Kehadiran") = False Then Exit Sub
    If sp_StatusHadir("A") = False Then Exit Sub
    
    MsgBox "Penyimpanan Berhasil", vbInformation, "Medifirst2000"
    
    Call cmdBatal_Click
Exit Sub
gabril:
    Call msubPesanError
End Sub
Private Function sp_StatusHadir(f_Status As String) As Boolean
sp_StatusHadir = True
    
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdStatusHadir", adChar, adParamInput, 2, txtKdStatusHadir.Text)
        .Parameters.Append .CreateParameter("StatusHadir", adVarChar, adParamInput, 50, Trim(txtNamaKehadiran.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 50, txtKdExternal.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNamaExternal.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStatusEnabled.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        
        .ActiveConnection = dbConn
        .CommandText = "AUD_KehadiranDokter"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data Kehadiran Dokter", vbCritical, "Validasi"
            sp_StatusHadir = False
        End If
        
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Sub dgStatusHadir_Click()
On Error GoTo gabril
If dgStatusHadir.ApproxCount = 0 Then Exit Sub
    txtKdStatusHadir = dgStatusHadir.Columns(0).value
    txtNamaKehadiran = dgStatusHadir.Columns(1).value
    txtKdExternal = dgStatusHadir.Columns(2).value
    txtNamaExternal = dgStatusHadir.Columns(3).value
Exit Sub
gabril:
   ' Call msubPesanError
End Sub

Private Sub dgStatusHadir_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   Call dgStatusHadir_Click
End Sub

Private Sub Form_Activate()
    txtNamaKehadiran.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo gabril
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call cmdBatal_Click
Exit Sub
gabril:
    Call msubPesanError
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Private Sub Clear()
On Error GoTo gabril
    txtKdStatusHadir.Text = ""
    txtNamaKehadiran.Text = ""
    txtKdExternal.Text = ""
    txtNamaExternal.Text = ""
Exit Sub
gabril:
    Call msubPesanError
End Sub
Private Sub setgrid()
On Error GoTo gabril
    
strSQL = "select * from KehadiranDokter"
Call msubRecFO(rs, strSQL)
Set dgStatusHadir.DataSource = rs
    
With dgStatusHadir
    .Columns(0).Caption = "Kode Status"
    .Columns(1).Caption = "Nama Status"
    .Columns(2).Caption = "Kode External"
    .Columns(3).Caption = "Nama External"
    .Columns(4).Caption = "Status Enabled"
    
    .Columns(0).Alignment = dbgCenter
    .Columns(4).Alignment = dbgCenter
    
    .Columns(0).Width = 1200
    .Columns(1).Width = 1800
    .Columns(2).Width = 1200
    .Columns(3).Width = 1800
    .Columns(4).Width = 1100
End With
Exit Sub
gabril:
    Call msubPesanError
End Sub
Private Sub txtNamaKehadiran_KeyPress(KeyAscii As Integer)
On Error GoTo gabril
    If KeyAscii = 13 Then
        txtKdExternal.SetFocus
    End If
Exit Sub
gabril:
    msubPesanError
End Sub
Private Sub txtKdExternal_KeyPress(KeyAscii As Integer)
On Error GoTo gabril
    If KeyAscii = 13 Then
        txtNamaExternal.SetFocus
    End If
Exit Sub
gabril:
    msubPesanError
End Sub
Private Sub txtNamaExternal_KeyPress(KeyAscii As Integer)
On Error GoTo gabril
    If KeyAscii = 13 Then
        chkStatusEnabled.SetFocus
    End If
Exit Sub
gabril:
    msubPesanError
End Sub
Private Sub chkStatusEnabled_KeyPress(KeyAscii As Integer)
On Error GoTo gabril
    If KeyAscii = 13 Then
        cmdSimpan.SetFocus
    End If
Exit Sub
gabril:
    msubPesanError
End Sub
