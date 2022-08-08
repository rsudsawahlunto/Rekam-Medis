VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPeriodeLaporanIndexDiagnosaPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Laporan Index Diagnosa Pasien"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPeriodeLaporanIndexDiagnosaPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   14805
   Begin VB.Frame Frame3 
      Height          =   6375
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   14775
      Begin VB.TextBox txtAlamatPasien 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3840
         TabIndex        =   1
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtNamaDiagnosa 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   3495
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
         Left            =   8640
         TabIndex        =   10
         Top             =   240
         Width           =   6075
         Begin VB.CommandButton cmdcari 
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
            TabIndex        =   4
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPickerAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   2
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
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
            CustomFormat    =   "dd MMMM, yyyy"
            Format          =   115998723
            UpDown          =   -1  'True
            CurrentDate     =   37956
         End
         Begin MSComCtl2.DTPicker DTPickerAkhir 
            Height          =   375
            Left            =   3600
            TabIndex        =   3
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
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
            CustomFormat    =   "dd MMMM, yyyy"
            Format          =   115998723
            UpDown          =   -1  'True
            CurrentDate     =   37956
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3240
            TabIndex        =   11
            Top             =   322
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgData 
         Height          =   5175
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   9128
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   2
         RowHeight       =   15
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
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Alamat Pasien"
         Height          =   210
         Index           =   1
         Left            =   3840
         TabIndex        =   13
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama Diagnosa"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
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
      Left            =   0
      TabIndex        =   8
      Top             =   7320
      Width           =   14805
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   12960
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   11160
         TabIndex        =   6
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label LblJumData 
         AutoSize        =   -1  'True
         Caption         =   "10 / 20 Data"
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
         Top             =   240
         Width           =   1050
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   15
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
      Left            =   12960
      Picture         =   "frmPeriodeLaporanIndexDiagnosaPasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPeriodeLaporanIndexDiagnosaPasien.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPeriodeLaporanIndexDiagnosaPasien.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmPeriodeLaporanIndexDiagnosaPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCari_Click()
    On Error GoTo errLoad
    LblJumData.Caption = ""

    strSQL = "SELECT KodeICD, NamaDiagnosa, NoCM, NamaPasien, JK, Umur, Alamat, RuanganPemeriksaan, TglPeriksa, DokterPemeriksa, JenisDiagnosa, JenisPasien" & _
    " From V_IndexDiagnosaPasien" & _
    " WHERE TglPeriksa BETWEEN '" & Format(DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' AND (NamaDiagnosa LIKE '%" & txtNamaDiagnosa.Text & "%' OR NamaDiagnosa IS NULL) AND (Alamat LIKE '%" & txtAlamatPasien.Text & "%' OR Alamat IS NULL)"
    Call msubRecFO(rs, strSQL)
    Set dgData.DataSource = rs
    With dgData
        .Columns("NamaDiagnosa").Width = 2800
        .Columns("NamaPasien").Width = 2000
        .Columns("Umur").Width = 1300
        .Columns("Alamat").Width = 4500
        .Columns("RuanganPemeriksaan").Width = 1800
        .Columns("JenisDiagnosa").Width = 2000
        .Columns("JenisPasien").Width = 1100
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    cmdCetak.Enabled = False
    mdTglAwal = DTPickerAwal.value
    mdTglAkhir = DTPickerAkhir.value

    strSQL = "SELECT * " & _
    " From V_IndexDiagnosaPasien" & _
    " WHERE TglPeriksa BETWEEN '" & Format(DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' AND (NamaDiagnosa LIKE '%" & txtNamaDiagnosa.Text & "%' OR NamaDiagnosa IS NULL) AND (Alamat LIKE '%" & txtAlamatPasien.Text & "%' OR Alamat IS NULL)"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbExclamation, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    Else
        vLaporan = ""
        If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
        frmCetakIndexDiagnosaPasien.Show
        cmdCetak.Enabled = True
    End If
    Exit Sub
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgData_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgData
    WheelHook.WheelHook dgData
End Sub

Private Sub dgData_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    Me.Caption = dgData.Columns(ColIndex).Width
End Sub

Private Sub dgData_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    LblJumData.Caption = dgData.Bookmark & " / " & dgData.ApproxCount & " Data"
End Sub

Private Sub DTPickerAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdcari.SetFocus
End Sub

Private Sub DTPickerAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DTPickerAkhir.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    DTPickerAwal.value = Now
    DTPickerAkhir.value = Now
    Call cmdCari_Click
End Sub

Private Sub txtAlamatPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then DTPickerAwal.SetFocus
End Sub

Private Sub txtNamaDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAlamatPasien.SetFocus
End Sub

