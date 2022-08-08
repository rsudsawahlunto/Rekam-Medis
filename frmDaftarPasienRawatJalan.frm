VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDaftarPasienRawatJalan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Rawat Jalan"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienRawatJalan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   14670
   Begin VB.Frame frameJudul 
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
      TabIndex        =   10
      Top             =   1080
      Width           =   14655
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
         Left            =   8760
         TabIndex        =   11
         Top             =   120
         Width           =   5775
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   345
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   127008771
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   345
            Left            =   3480
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   127008771
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   12
            Top             =   307
            Width           =   255
         End
      End
      Begin VB.Label LblJumData 
         AutoSize        =   -1  'True
         Caption         =   "10 / 100 Data"
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
         TabIndex        =   14
         Top             =   720
         Width           =   1155
      End
   End
   Begin MSDataGridLib.DataGrid dgPasienRJ 
      Height          =   5175
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
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
            LCID            =   1033
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
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
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
      Height          =   975
      Left            =   0
      TabIndex        =   8
      Top             =   7320
      Width           =   14655
      Begin VB.TextBox txtJnsPasien 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6360
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtRuangan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3240
         TabIndex        =   5
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   450
         Left            =   12960
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pasien"
         Height          =   210
         Left            =   6360
         TabIndex        =   16
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan"
         Height          =   210
         Left            =   3240
         TabIndex        =   15
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Nama Pasien / No. CM"
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2640
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   8340
      Width           =   14670
      _ExtentX        =   25876
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   12885
            Text            =   "Cetak Daftar Pasien Daftar (F11)"
            TextSave        =   "Cetak Daftar Pasien Daftar (F11)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   12885
            Text            =   "Refresh Data (F5)"
            TextSave        =   "Refresh Data (F5)"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   17
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
      Left            =   12840
      Picture         =   "frmDaftarPasienRawatJalan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPasienRawatJalan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPasienRawatJalan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDaftarPasienRawatJalan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCari_Click()
On Error GoTo hell
    lblJumData.Caption = "0/0"
    Set rs = Nothing
    strSQL = "select * from V_LaporanPasienRawatJalan where ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%') AND (TglMasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "') AND NamaRuangan LIKE '%" & txtRuangan & "%' AND JenisPasien LIKE '%" & txtJnsPasien & "%'"
    rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    Set dgPasienRJ.DataSource = rs
    Call SetGridPasienRJ
    lblJumData.Caption = "1 / " & dgPasienRJ.ApproxCount & " Data"
    If dgPasienRJ.ApproxCount = 0 Then dtpAwal.SetFocus Else dgPasienRJ.SetFocus
Exit Sub
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgPasienRJ_Click()
WheelHook.WheelUnHook
        Set MyProperty = dgPasienRJ
        WheelHook.WheelHook dgPasienRJ
End Sub

Private Sub dgPasienRJ_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    lblJumData.Caption = dgPasienRJ.Bookmark & " / " & dgPasienRJ.ApproxCount & " Data"
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo hell
Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    Select Case KeyCode
        Case vbKeyF5
            Call cmdCari_Click
        Case vbKeyF11
            If dgPasienRJ.ApproxCount = 0 Then Exit Sub
            mdTglAwal = dtpAwal.value: mdTglAkhir = dtpAkhir.value
            frmCetakDaftarPasienRawatJalan.Show
    End Select
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpAwal.value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.value = Now
    frameJudul.Caption = "Daftar Pasien Rawat Jalan "
    Call cmdCari_Click
Exit Sub
errLoad:
Call msubPesanError
End Sub

Sub SetGridPasienRJ()
    With dgPasienRJ
        .Columns(0).Width = 600
        .Columns(0).Caption = "No Urut"
        .Columns(1).Width = 800
        .Columns(1).Caption = "No CM"
        .Columns(2).Width = 3000
        .Columns(3).Width = 800
        .Columns(3).Alignment = dbgCenter
        .Columns(3).Caption = "Kunjungan"
        .Columns(4).Width = 1590
        .Columns(5).Width = 300
        .Columns(6).Width = 1590
        .Columns(7).Width = 1700
        .Columns(8).Width = 3000
        .Columns(9).Width = 600
        .Columns(9).Caption = "Kasus"
        .Columns(10).Width = 1900
    End With
End Sub

Private Sub txtJnsPasien_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdCari_Click
    txtJnsPasien.SetFocus
End If
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdCari_Click
        txtParameter.SetFocus
    End If
End Sub

Private Sub txtRuangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            Call cmdCari_Click
            txtRuangan.SetFocus
    End If
End Sub
