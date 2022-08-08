VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAsuransi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Asuransi Pasien"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAsuransi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   8880
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   0
      TabIndex        =   26
      Top             =   4320
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5953
      _Version        =   393216
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   17
      Top             =   3360
      Width           =   8775
      Begin VB.CommandButton cmdBatal 
         Caption         =   "Batal"
         Height          =   495
         Left            =   3000
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   4920
         TabIndex        =   19
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         Height          =   495
         Left            =   6840
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   8775
      Begin VB.TextBox txtIdPeserta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         MaxLength       =   16
         TabIndex        =   22
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtAlamat 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         MaxLength       =   100
         TabIndex        =   12
         Top             =   1800
         Width           =   3855
      End
      Begin VB.TextBox txtNamaPeserta 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1200
         Width           =   3975
      End
      Begin VB.TextBox txtIdAsuransi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5280
         MaxLength       =   25
         TabIndex        =   4
         Top             =   480
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker dtptglLahir 
         Height          =   330
         Left            =   6360
         TabIndex        =   10
         Top             =   1200
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd-MMMM-yyyy"
         Format          =   131072003
         CurrentDate     =   38077
      End
      Begin MSDataListLib.DataCombo dcGolongan 
         Height          =   330
         Left            =   4080
         TabIndex        =   13
         Top             =   1800
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcInstitusi 
         Height          =   330
         Left            =   6360
         TabIndex        =   14
         Top             =   1800
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker dtpTglDaftar 
         Height          =   330
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   131072003
         UpDown          =   -1  'True
         CurrentDate     =   38077
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         TabIndex        =   6
         Top             =   3360
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo dcPenjamin 
         Height          =   330
         Left            =   2280
         TabIndex        =   27
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label11 
         Caption         =   "Tanggal Daftar"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "ID Penjamin"
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "ID Peserta"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Institusi Asal"
         Height          =   255
         Left            =   6360
         TabIndex        =   16
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Golongan"
         Height          =   255
         Left            =   4080
         TabIndex        =   15
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Alamat"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Tanggal Lahir"
         Height          =   255
         Left            =   6360
         TabIndex        =   9
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Nama Peserta"
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "ID Asuransi"
         Height          =   255
         Left            =   5280
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "No. CM"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3120
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Penjamin"
         Height          =   255
         Left            =   4680
         TabIndex        =   2
         Top             =   3000
         Visible         =   0   'False
         Width           =   2055
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
   Begin VB.Image Image2 
      Height          =   960
      Left            =   7320
      Picture         =   "frmAsuransi.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmAsuransi.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmAsuransi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBatal_Click()
    Call Clear
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo errLoad
If txtIdAsuransi.Text = "" Then
    MsgBox "ID Asuransi harus diisi", vbInformation
    txtIdAsuransi.SetFocus
Exit Sub

ElseIf dcPenjamin.Text = "" Then
    MsgBox "Nama Penjamin harus diisi", vbInformation
    dcPenjamin.SetFocus
Exit Sub

ElseIf dcGolongan.Text = "" Then
    MsgBox "Nama Golongan harus diisi", vbInformation
    dcGolongan.SetFocus
Exit Sub

ElseIf dcInstitusi.Text = "" Then
    MsgBox "Nama Institusi harus diisi", vbInformation
    dcInstitusi.SetFocus
Exit Sub

ElseIf txtIdPeserta.Text = "" Then
    MsgBox "ID Peserta harus diisi", vbInformation
    txtIdPeserta.SetFocus
Exit Sub

ElseIf txtNamaPeserta.Text = "" Then
    MsgBox "Nama Peserta harus diisi", vbInformation
    txtNamaPeserta.SetFocus
Exit Sub

End If
    If sp_Asuransi("A") = False Then Exit Sub
Exit Sub
errLoad:
    Call msubPesanError
End Sub
Private Function sp_Asuransi(f_Status) As Boolean
    sp_RencanaAskep = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, Trim(dcPenjamin.BoundText))
        .Parameters.Append .CreateParameter("IdAsuransi", adVarChar, adParamInput, 25, txtIdAsuransi.Text)
        .Parameters.Append .CreateParameter("TglDaftar", adDate, adParamInput, , Format(dtpTglDaftar.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, Null)
        .Parameters.Append .CreateParameter("NamaPeserta", adVarChar, adParamInput, 50, txtNamaPeserta.Text)
        .Parameters.Append .CreateParameter("IdPeserta", adVarChar, adParamInput, 16, txtIdPeserta.Text)
        .Parameters.Append .CreateParameter("KdGolongan", adChar, adParamInput, 2, dcGolongan.BoundText)
        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(dtptglLahir.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Alamat", adVarChar, adParamInput, 100, txtAlamat.Text)
        .Parameters.Append .CreateParameter("KdInstitusiAsal", adVarChar, adParamInput, 4, dcInstitusi.BoundText)
        .Parameters.Append .CreateParameter("QAsuransiPasien", adInteger, adParamInput, , Null)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_DataPesertaAsuransi"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "ada Kesalahan dalam penyimpanan data", vbInformation
            sp_Asuransi = False
        Else
            If f_Status = "A" Then
                MsgBox "Pendaftaran Asuransi Berhasil tersimpan", vbInformation
            Else
                MsgBox "Data Berhasil Dihapus", vbInformation
            End If
            cmdTutup.SetFocus
            Call Clear
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcGolongan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcInstitusi.SetFocus
End Sub

Private Sub dcInstitusi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub dcPenjamin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtIdAsuransi.SetFocus
End Sub

Private Sub dtptglLahir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtAlamat.SetFocus
End Sub

Private Sub dtptglLahir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAlamat.SetFocus
End Sub

Private Sub Form_Activate()
    dcPenjamin.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    
    dtptglLahir.value = Now
    dtpTglDaftar.value = Now
    
    Call Clear
    Call subLoadDcSource

Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub subLoadDcSource()
On Error GoTo errLoad

    Call msubDcSource(dcGolongan, rs, "Select KdGolongan,NamaGolongan from GolonganAsuransi where StatusEnabled=1")
    Call msubDcSource(dcInstitusi, rs, "Select KdInstitusiAsal,InstitusiAsal from InstitusiAsalPasien where StatusEnabled=1")
    Call msubDcSource(dcPenjamin, rs, "Select IdPenjamin,NamaPenjamin from Penjamin where StatusEnabled=1")
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub
Private Sub Clear()
On Error GoTo errLoad
    txtNoCM.Text = ""
    txtNamaPeserta.Text = ""
    txtAlamat.Text = ""
    txtIdAsuransi.Text = ""
    txtIdPeserta.Text = ""
    dcGolongan.Text = ""
    dcInstitusi.Text = ""
    dcPenjamin.Text = ""
Exit Sub
errLoad:
    Call msubPesanError
End Sub
Private Sub txtAlamat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcGolongan.SetFocus
End Sub

Private Sub txtIdAsuransi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtIdPeserta.SetFocus
End Sub
Private Sub txtIdPeserta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaPeserta.SetFocus
End Sub
Private Sub txtNamaPeserta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtptglLahir.SetFocus
End Sub
