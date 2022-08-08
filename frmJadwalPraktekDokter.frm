VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmJadwalPraktekDokter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jadwal Praktek Dokter"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJadwalPraktekDokter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   8565
   Begin VB.TextBox txtKdDokter 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8640
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   3855
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
      Height          =   2055
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Visible         =   0   'False
      Width           =   7815
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   1455
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   16
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   0
      TabIndex        =   20
      Top             =   3480
      Width           =   8535
      Begin VB.CommandButton cmdKeluar 
         Caption         =   "Kelua&r"
         Height          =   495
         Left            =   6480
         TabIndex        =   27
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   4320
         TabIndex        =   26
         Top             =   3840
         Width           =   2055
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   495
         Left            =   2280
         TabIndex        =   25
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   3840
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid dgJadwalPraktek 
         Height          =   3495
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   6165
         _Version        =   393216
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   8535
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5040
         TabIndex        =   19
         Top             =   1800
         Width           =   3375
      End
      Begin MSDataListLib.DataCombo dcStatus 
         Height          =   330
         Left            =   2280
         TabIndex        =   17
         Top             =   1800
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSMask.MaskEdBox meJamselesai 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "H:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   4
         EndProperty
         Height          =   330
         Left            =   1200
         TabIndex        =   15
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meJamMulai 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "H:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   4
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   330
         Left            =   5040
         TabIndex        =   11
         Top             =   1080
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox txtKdPraktek 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox cboHari 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmJadwalPraktekDokter.frx":0CCA
         Left            =   3000
         List            =   "frmJadwalPraktekDokter.frx":0CE3
         TabIndex        =   5
         Top             =   360
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker dtpTanggal 
         Height          =   375
         Left            =   5880
         TabIndex        =   3
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   108134400
         UpDown          =   -1  'True
         CurrentDate     =   41487
      End
      Begin VB.Label Label9 
         Caption         =   "Keterangan"
         Height          =   255
         Left            =   5040
         TabIndex        =   18
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Status Praktek"
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Jam Selesai"
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Jam Mulai"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Ruangan"
         Height          =   255
         Left            =   5040
         TabIndex        =   10
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Nama Dokter"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Kode Praktek Dokter"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Hari"
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Tanggal"
         Height          =   255
         Left            =   5880
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
      Picture         =   "frmJadwalPraktekDokter.frx":0D19
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmJadwalPraktekDokter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFilter As String
Dim adoCommand As New ADODB.Command
Private Sub cboHari_KeyPress(KeyAscii As Integer)
On Error GoTo gabril
    If KeyAscii = 13 Then
        txtDokter.SetFocus
    Else
    End If
    Exit Sub
gabril:
    Call msubPesanError
End Sub


Private Sub dcNamaPoli_KeyPress(KeyAscii As Integer)
On Error GoTo gabril
    If KeyAscii = 13 Then
        meJamMulai.SetFocus
    Else
    End If
    Exit Sub
gabril:
    Call msubPesanError
End Sub

Private Sub dcRuangan_KeyPress(KeyAscii As Integer)
On Error GoTo gabril
    If KeyAscii = 13 Then
        meJamMulai.SetFocus
    Else
    End If
Exit Sub
gabril:
    Call msubPesanError
End Sub

Private Sub dgJadwalPraktek_Click()
On Error GoTo gabril
With dgJadwalPraktek
    If dgJadwalPraktek.ApproxCount = 0 Then Exit Sub
    txtKdPraktek.Text = .Columns(0).value
    cboHari.Text = .Columns(1).value
    dtpTanggal.value = .Columns(2).value
    txtDokter.Text = .Columns(3).value
    dcRuangan.Text = .Columns(4).value
    meJamMulai.Text = .Columns(5).value
    meJamselesai.Text = .Columns(6).value
    dcStatus.Text = .Columns(7).value
    txtKeterangan.Text = .Columns(8).value
    fraDokter.Visible = False
End With
Exit Sub
gabril:
   ' Call msubPesanError
End Sub

Private Sub dgJadwalPraktek_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call dgJadwalPraktek_Click
End Sub

'
'Private Sub dgJadwalPraktek_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'On Error GoTo gabril
'With dgJadwalPraktek
'    txtKdPraktek.Text = .Columns(0).value
'    cboHari.Text = .Columns(1).value
'    dtpTanggal.value = .Columns(2).value
'    txtDokter.Text = .Columns(3).value
'    dcRuangan.Text = .Columns(4).value
'    dcNamaPoli.Text = .Columns(5).value
'    meJamMulai.Text = .Columns(6).value
'    meJamselesai.Text = .Columns(7).value
'    dcStatus.Text = .Columns(8).value
'    txtKeterangan.Text = .Columns(9).value
'    fraDokter.Visible = False
'End With
'Exit Sub
'gabril:
'    Call msubPesanError
'End Sub

Private Sub meJamMulai_KeyPress(KeyAscii As Integer)
On Error GoTo gabril
    If KeyAscii = 13 Then
        meJamselesai.SetFocus
    Else
    End If
    Exit Sub
gabril:
    Call msubPesanError
End Sub
Private Sub meJamSelesai_KeyPress(KeyAscii As Integer)
On Error GoTo gabril
    If KeyAscii = 13 Then
        dcStatus.SetFocus
    Else
    End If
    Exit Sub
gabril:
    Call msubPesanError
End Sub
Private Sub dcStatus_KeyPress(KeyAscii As Integer)
On Error GoTo gabril
    If KeyAscii = 13 Then
        cmdSimpan.SetFocus
    Else
    End If
    Exit Sub
gabril:
    Call msubPesanError
End Sub
Private Sub cmdBatal_Click()
On Error Resume Next
   ' Call setgrid
    Call Clear
     
Exit Sub
'gabril:
'    Call msubPesanError
End Sub

Private Sub cmdHapus_Click()
On Error GoTo gabril
    If Periksa("text", txtDokter, "Silahkan isi Nama Dokter") = False Then Exit Sub
    If MsgBox("Apakah anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    If sp_JadwalPraktekDokter("D") = False Then Exit Sub
    
    MsgBox "Data berhasil di hapus", vbInformation, "Medifirst2000"
    
    Call cmdBatal_Click
Exit Sub
gabril:
    Call msubPesanError
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo handel
    If cboHari.Text = "" Then
    MsgBox "Silahkan isi hari praktek"
    cboHari.SetFocus
    Exit Sub
    End If
    If Periksa("text", txtDokter, "Silahkan isi Nama Dokter") = False Then Exit Sub
    If Periksa("datacombo", dcRuangan, "Silahkan isi ruangan praktek") = False Then Exit Sub
    
If meJamMulai.Text = "__:__" Or meJamselesai.Text = "__:__" Then
    MsgBox "Lengkapi Jam Praktek Terlebih Dahulu", vbExclamation, "Medifirst2000"
    meJamMulai.SetFocus
Else
    If sp_JadwalPraktekDokter("A") = False Then Exit Sub
    
    MsgBox "Penyimpanan Berhasil", vbInformation, "Medifirst2000"
    
    Call cmdBatal_Click
End If

Exit Sub
handel:
    Call msubPesanError
End Sub
Private Function sp_JadwalPraktekDokter(f_Status As String) As Boolean
sp_JadwalPraktekDokter = True
    
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdPraktek", adChar, adParamInput, 3, txtKdPraktek.Text)
        .Parameters.Append .CreateParameter("Hari", adVarChar, adParamInput, 7, Trim(cboHari.Text))
        .Parameters.Append .CreateParameter("Tgl", adDBDate, adParamInput, , Format(dtpTanggal.value, "dd/MMMM/yyyy HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtKdDokter.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, Trim(dcRuangan.BoundText))
'        .Parameters.Append .CreateParameter("RuanganPoli", adChar, adParamInput, 2, Trim(IIf((dcNamaPoli.BoundText = ""), "-", dcNamaPoli.BoundText)))
        If meJamMulai.Text <> "__:__" Then
            .Parameters.Append .CreateParameter("JamMulai", adVarChar, adParamInput, 5, meJamMulai.Text)
        Else
            .Parameters.Append .CreateParameter("JamMulai", adVarChar, adParamInput, 5, Null)
        End If
        
        If meJamselesai.Text <> "__:__" Then
            .Parameters.Append .CreateParameter("JamSelesai", adVarChar, adParamInput, 5, meJamselesai.Text)
        Else
            .Parameters.Append .CreateParameter("JamSelesai", adVarChar, adParamInput, 5, Null)
        End If
'        .Parameters.Append .CreateParameter("JamSelesai", adTinyInt, adParamInput, , Format(meJamselesai.Text, "HH:mm"))
        .Parameters.Append .CreateParameter("KdStatusHadir", adChar, adParamInput, 2, dcStatus.BoundText)
        .Parameters.Append .CreateParameter("Keterangan", adChar, adParamInput, 50, txtKeterangan.Text)

        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        
        .ActiveConnection = dbConn
        .CommandText = "AUD_JadwalPraktekDokter"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data Kehadiran Dokter", vbCritical, "Validasi"
            sp_JadwalPraktekDokter = False
        End If
        
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

'Private Sub dgJadwal_Click()
'On Error GoTo gabril
'If dgJadwalPraktek.ApproxCount = 0 Then Exit Sub
'    txtKdStatusHadir = dgJadwalPraktek.Columns(0).Value
'    txtNamaKehadiran = dgJadwalPraktek.Columns(1).Value
'    txtKdExternal = dgJadwalPraktek.Columns(2).Value
'    txtNamaExternal = dgJadwalPraktek.Columns(3).Value
'Exit Sub
'gabril:
'    Call msubPesanError
'End Sub

Private Sub Form_Activate()
    cboHari.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo gabril
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call dcSource
   ' Call cmdBatal_Click
    Call loadGridSource
Exit Sub
gabril:
    Call msubPesanError
End Sub
Private Sub dcSource()
On Error GoTo gabril
    'Call msubDcSource(dcRuangan, rs, "Select KdRuangan,NamaRuangan,KdInstalasi from Ruangan where KdInstalasi in ('02','03','06','09')")
    Call msubDcSource(dcRuangan, rs, "Select KdRuangan,NamaRuangan,KdInstalasi from Ruangan where KdInstalasi in ('02')")
    dcRuangan.Text = ""
    dcRuangan.BoundText = rs(0).value
    Set rs = Nothing
    Call msubDcSource(dcStatus, rs, "Select * from KehadiranDokter where StatusEnabled='1'")
    Exit Sub
gabril:
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Private Sub Clear()
On Error GoTo gabril
    txtKdPraktek.Text = ""
    cboHari.Text = ""
    dtpTanggal.value = Now
    txtDokter.Text = ""
    dcRuangan.Text = ""
    meJamMulai.Text = "__:__"
    meJamselesai.Text = "__:__"
    dcStatus.Text = ""
    txtKeterangan.Text = ""
    Call loadGridSource
Exit Sub
gabril:
    Call msubPesanError
End Sub
Public Sub loadGridSource()
On Error GoTo gabril
    
strSQL = "select KdPraktek,Hari,Tgl,NamaLengkap,NamaRuangan,JamMulai,JamSelesai,StatusHadir,Keterangan from V_JadwalPraktekDokter"
Call msubRecFO(rs, strSQL)

Set dgJadwalPraktek.DataSource = rs

With dgJadwalPraktek
    .Columns(0).Caption = "Kode"
    .Columns(1).Caption = "Hari"
    .Columns(2).Caption = "Tanggal"
    .Columns(3).Caption = "Nama Dokter"
    .Columns(4).Caption = "Ruangan Praktek"
    .Columns(5).Caption = "Jam Mulai"
    .Columns(6).Caption = "Jam Selesai"
    .Columns(7).Caption = "Status Kehadiran"
    .Columns(8).Caption = "Keterangan"
    
    .Columns(0).Width = 0
    .Columns(1).Width = 900
    '.Columns(2).Width = 1800
    .Columns(2).Width = 0
    .Columns(3).Width = 3000
    .Columns(4).Width = 1800
    .Columns(5).Width = 1100
    .Columns(6).Width = 1100
    .Columns(7).Width = 1100
    .Columns(8).Width = 1100
    
    
End With
Exit Sub
gabril:
    Call msubPesanError
End Sub
Private Sub txtDokter_Change()
On Error GoTo gabril
    strFilter = "WHERE NamaDokter like '%" & txtDokter.Text & "%' and KdJenisPegawai='001'"
'    txtDokter.Text = ""
    Call subLoadDokter
    fraDokter.Left = 120
    fraDokter.Top = Frame1.Top + txtDokter.Top + txtDokter.Height
    If txtDokter.Text = "" Then
        fraDokter.Visible = False
    Else
        fraDokter.Visible = True
    End If
 
'   Me.Height = 8415
    Call centerForm(Me, MDIUtama)
Exit Sub
gabril:
    Call msubPesanError
End Sub
Private Sub txtDokter_GotFocus()
    Call txtDokter_Change
End Sub

Private Sub txtDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If fraDokter.Visible = False Then Exit Sub
        dgDokter.SetFocus
    End If
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
On Error GoTo gabril
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        dgDokter.SetFocus
    End If
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 27 Then
        fraDokter.Visible = False
'        Me.Height = 8415
    End If
    Call SetKeyPressToChar(KeyAscii)
gabril:
End Sub
Private Sub subLoadDokter()
    On Error Resume Next
    strSQL = "SELECT NamaDokter AS [Nama Dokter],KodeDokter AS [Kode Dokter],JK,Jabatan FROM V_DaftarDokter " & strFilter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDokter = rs.RecordCount
    Set dgDokter.DataSource = rs
    With dgDokter
        .Columns(0).Width = 3000 'nama dokter
        .Columns(1).Width = 0 'kode dokter
        .Columns(2).Width = 400
        .Columns(3).Width = 3300
    End With
    txtKdDokter.Text = rs(1).value
End Sub
Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        txtDokter.Text = dgDokter.Columns("Nama Dokter").value
        txtKdDokter.Text = dgDokter.Columns("Kode Dokter").value
        If txtDokter.Text = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        fraDokter.Visible = False
        'Me.Height = 8415
        Call centerForm(Me, MDIUtama)
    End If
End Sub

