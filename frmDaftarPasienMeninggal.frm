VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDaftarPasienMeninggal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Meninggal"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienMeninggal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   14340
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   10800
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   8235
      Width           =   14340
      _ExtentX        =   25294
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   12788
            MinWidth        =   12788
            Text            =   "Cetak Surat Keterangan Meninggal [ F1 ] "
            TextSave        =   "Cetak Surat Keterangan Meninggal [ F1 ] "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   12788
            MinWidth        =   12788
            Text            =   " Cetak Daftar Pasien Meninggal [ Shift + F1 ]"
            TextSave        =   " Cetak Daftar Pasien Meninggal [ Shift + F1 ]"
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
   Begin VB.Frame frameJudul 
      Caption         =   "Daftar Pasien Meninggal di "
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
      TabIndex        =   9
      Top             =   960
      Width           =   14295
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
         Left            =   8400
         TabIndex        =   10
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
            Height          =   375
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   127074307
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   127074307
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   11
            Top             =   322
            Width           =   255
         End
      End
   End
   Begin MSDataGridLib.DataGrid dgPasienMeninggal 
      Height          =   5295
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   9340
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
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   7320
      Width           =   14295
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   12495
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   450
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan  Nama Pasien / No.CM"
         Height          =   210
         Left            =   1560
         TabIndex        =   8
         Top             =   195
         Width           =   2640
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   13
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
      Picture         =   "frmDaftarPasienMeninggal.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPasienMeninggal.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPasienMeninggal.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDaftarPasienMeninggal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCari_Click()
    On Error GoTo hell
    strSQL = "select distinct * from V_DaftarPasienMeninggal where([Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%' OR NoCM LIKE '%" & txtParameter.Text & "%') AND TglMeninggal between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "' "
    Call msubRecFO(rs, strSQL)
    Set dgPasienMeninggal.DataSource = rs
    Call SetGridPasienMeninggal
    If dgPasienMeninggal.ApproxCount = 0 Then dtpAwal.SetFocus Else dgPasienMeninggal.SetFocus
    Exit Sub
hell:
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    cmdCetak.Enabled = False
    strSQL = " select distinct * from V_DaftarPasienMeninggal where([Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%' OR Nocm LIKE '%" & txtParameter.Text & "%') AND TglMeninggal between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbExclamation, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    Else

        Set FrmViewerLaporan = Nothing
        cetak = "DPMeninggal"
        FrmViewerLaporan.Show
        cmdCetak.Enabled = True
    End If
    Exit Sub
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgPasienMeninggal_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPasienMeninggal
    WheelHook.WheelHook dgPasienMeninggal
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errLoad
    Dim strShiftKey As String
    strShiftKey = (Shift + vbShiftMask)
    Select Case KeyCode
        Case vbKeyF1
            If strShiftKey = 2 Then
                Call cmdCetak_Click
            Else
                cmdCetak.Enabled = False
                strSQL = " select * from V_DaftarPasienMeninggal where NOCM='" & dgPasienMeninggal.Columns("NoCM").value & "'  AND TglMeninggal between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'"
                Set rs = Nothing
                rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
                If rs.RecordCount = 0 Then
                    MsgBox "Tidak ada data", vbExclamation, "Validasi"
                    cmdCetak.Enabled = True
                    Exit Sub
                Else
                    Set FrmViewerLaporan = Nothing
                    cetak = "PMeninggal"
                    FrmViewerLaporan.Show
                    cmdCetak.Enabled = True
                End If
            End If
    End Select
    Exit Sub
errLoad:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    frameJudul.Caption = "Daftar Pasien Meninggal"
    dtpAkhir.value = Now
    dtpAwal.value = Format(Now, "dd MMM yyyy 00:00:00")
    Call cmdCari_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub SetGridPasienMeninggal()
    With dgPasienMeninggal
        .Columns(0).Width = 0
        .Columns(1).Width = 800
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 1800
        .Columns(3).Width = 300
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Width = 1400
        .Columns(5).Width = 3500
        .Columns(6).Width = 1590
        .Columns(7).Width = 1590
        .Columns(8).Width = 2300
        .Columns(9).Width = 0
        .Columns(10).Width = 2500
        .Columns(11).Width = 0
        .Columns(12).Width = 0
    End With
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdCari_Click
        txtParameter.SetFocus
    End If
End Sub

Private Sub txtParameter_LostFocus()
    txtParameter.Text = StrConv(txtParameter.Text, vbProperCase)
End Sub
