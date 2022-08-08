VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmInformasiTarifPelayanan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Informasi Tarif Pelayanan"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInformasiTarifPelayanan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   11205
   Begin VB.Frame Frame3 
      Height          =   6135
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   11175
      Begin MSDataGridLib.DataGrid dgData 
         Height          =   5535
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   9763
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
         Left            =   105
         TabIndex        =   9
         Top             =   5835
         Width           =   1050
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
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   7080
      Width           =   11205
      Begin VB.TextBox txtNamaPemeriksaan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtKelasPelayanan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3840
         TabIndex        =   1
         Top             =   480
         Width           =   3015
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   9240
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   7440
         TabIndex        =   3
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Masukan Nama Pemeriksaan"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   2265
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Kelas"
         Height          =   210
         Index           =   1
         Left            =   3840
         TabIndex        =   7
         Top             =   240
         Width           =   405
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   10
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
      Left            =   9360
      Picture         =   "frmInformasiTarifPelayanan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmInformasiTarifPelayanan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmInformasiTarifPelayanan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmInformasiTarifPelayanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCari_Click()
    On Error GoTo errLoad
    LblJumData.Caption = ""

    strSQL = "SELECT [Jenis Pelayanan], [Nama Pelayanan], [Kelas Pelayanan], [Tarif Pelayanan]" & _
    " From V_InfoTarifPelayanan" & _
    " WHERE [Nama Pelayanan] LIKE '%" & txtNamaPemeriksaan.Text & "%' AND [Kelas Pelayanan] LIKE '%" & txtKelasPelayanan.Text & "%' and StatusEnabled='1' and Expr1='1' and Expr2='1'" & _
    " ORDER BY [Jenis Pelayanan], [Nama Pelayanan]"
    Call msubRecFO(rs, strSQL)
    LblJumData.Caption = rs.RecordCount & " Data"

    Set dgData.DataSource = rs
    With dgData
        .Columns("Jenis Pelayanan").Width = 3200
        .Columns("Jenis Pelayanan").Caption = "Jenis Pemeriksaan"

        .Columns("Nama Pelayanan").Width = 4200
        .Columns("Nama Pelayanan").Caption = "Nama Pemeriksaan"

        .Columns("Kelas Pelayanan").Width = 1500
        .Columns("Kelas Pelayanan").Caption = "Kelas"

        .Columns("Tarif Pelayanan").Width = 1300
        .Columns("Tarif Pelayanan").NumberFormat = "#,###"
        .Columns("Tarif Pelayanan").Alignment = dbgRight
        .Columns("Tarif Pelayanan").Caption = "Tarif"
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo errLoad
    cmdCetak.Enabled = False
    mstrFilter = ""
    strSQL = "SELECT * " & _
    " From V_InfoTarifPelayanan WHERE [Nama Pelayanan] LIKE '%" & txtNamaPemeriksaan.Text & "%' AND [Kelas Pelayanan] LIKE '%" & txtKelasPelayanan.Text & "%'" & _
    " "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbExclamation, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    Else
        vLaporan = ""
        If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    End If

    mstrFilter = " WHERE [Nama Pelayanan] LIKE '%" & txtNamaPemeriksaan.Text & "%' AND [Kelas Pelayanan] LIKE '%" & txtKelasPelayanan.Text & "%'"
    cmdCetak.Enabled = True
    frmCetakInformasiTarifPelayanan.Show

errLoad:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgData_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgData
    WheelHook.WheelHook dgData
End Sub

Private Sub dgData_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    LblJumData.Caption = dgData.Bookmark & " / " & dgData.ApproxCount & " Data"
End Sub

Private Sub DTPickerAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
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
    Call cmdCari_Click
End Sub

Private Sub txtKelasPelayanan_Change()
    Call cmdCari_Click
End Sub

Private Sub txtKelasPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then If cmdCetak.Enabled = True Then cmdCetak.SetFocus
End Sub

Private Sub txtNamaPemeriksaan_Change()
    Call cmdCari_Click
End Sub

Private Sub txtNamaPemeriksaan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKelasPelayanan.SetFocus
End Sub

