VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmDetailDiagnosaKeperawatan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Detail Diagnosa Keperawatan"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDetailDiagnosaKeperawatan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   7470
   Begin VB.TextBox txtOuputKode 
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   17
      Top             =   7560
      Width           =   7455
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   6120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   4920
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "F1 - Cetak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   930
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   20
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Detail Diagnosa Keperawatan"
      TabPicture(0)   =   "frmDetailDiagnosaKeperawatan.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Penyebab Diagnosa Keperawatan"
      TabPicture(1)   =   "frmDetailDiagnosaKeperawatan.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   -74880
         TabIndex        =   30
         Top             =   600
         Width           =   7095
         Begin VB.TextBox txtKodeExternal1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   240
            TabIndex        =   13
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox txtNamaExternal1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   240
            TabIndex        =   14
            Top             =   2040
            Width           =   5175
         End
         Begin VB.CheckBox CheckStatusEnbl1 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   5520
            TabIndex        =   15
            Top             =   2160
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtpenyebab 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   500
            TabIndex        =   12
            Top             =   600
            Width           =   5055
         End
         Begin VB.TextBox txtkdpenyebab 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   240
            MaxLength       =   7
            TabIndex        =   11
            Top             =   600
            Width           =   1455
         End
         Begin MSDataGridLib.DataGrid dgPenyebab 
            Height          =   3135
            Left            =   240
            TabIndex        =   16
            Top             =   2520
            Width           =   6600
            _ExtentX        =   11642
            _ExtentY        =   5530
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
         Begin VB.Label Label9 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Penyebab Diagnosa Keperawatan"
            Height          =   210
            Left            =   1800
            TabIndex        =   32
            Top             =   360
            Width           =   2730
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Kode Penyebab"
            Height          =   210
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   1290
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   7095
         Begin VB.CheckBox CheckStatusEnbl 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   5520
            TabIndex        =   5
            Top             =   2160
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtNamaExternal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   240
            TabIndex        =   4
            Top             =   2040
            Width           =   5175
         End
         Begin VB.TextBox txtKodeExternal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   240
            TabIndex        =   3
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtKdDetailAskep 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   240
            MaxLength       =   7
            TabIndex        =   1
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtDetailAskep 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   500
            TabIndex        =   2
            Top             =   600
            Width           =   5055
         End
         Begin MSDataGridLib.DataGrid dgDetailAskep 
            Height          =   3135
            Left            =   240
            TabIndex        =   6
            Top             =   2520
            Width           =   6600
            _ExtentX        =   11642
            _ExtentY        =   5530
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
         Begin VB.Label Label7 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Kode Detail"
            Height          =   210
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Detail Asuhan Keperawatan"
            Height          =   210
            Left            =   1800
            TabIndex        =   28
            Top             =   360
            Width           =   2250
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   -74760
         TabIndex        =   21
         Top             =   480
         Width           =   6855
         Begin VB.TextBox txtKdKelompokPegawai 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   330
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   23
            Top             =   360
            Width           =   4215
         End
         Begin VB.TextBox txtKelompokPegawai 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2400
            MaxLength       =   50
            TabIndex        =   22
            Top             =   720
            Width           =   4215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Kode Kelompok Pegawai"
            Height          =   210
            Left            =   240
            TabIndex        =   25
            Top             =   405
            Width           =   2010
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Kelompok Pegawai"
            Height          =   210
            Left            =   240
            TabIndex        =   24
            Top             =   780
            Width           =   1530
         End
      End
      Begin MSDataGridLib.DataGrid dgKelompokPegawai 
         Height          =   3855
         Left            =   -74760
         TabIndex        =   26
         Top             =   1800
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6800
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   5760
      Picture         =   "frmDetailDiagnosaKeperawatan.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1755
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDetailDiagnosaKeperawatan.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDetailDiagnosaKeperawatan.frx":444B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDetailDiagnosaKeperawatan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilterDiagnosa As String
Dim intJmlDetail As Integer
Dim mstrKdDetail As String

Private Sub cmdBatal_Click()
    On Error GoTo errLoad

    Call clear
    Call subLoadGridSource

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdHapus_Click()

    On Error GoTo errLoad

    If MsgBox("Yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub

    Select Case SSTab1.Tab
        Case 0
            If txtKdDetailAskep.Text = "" Then
                MsgBox "Pilih dulu data yang akan dihapus", vbOKOnly, "Validasi"
                Exit Sub
            End If

            If sp_DetailAskep("D") = False Then Exit Sub
        Case 1
            If txtkdpenyebab.Text = "" Then
                MsgBox "Pilih dulu data yang akan dihapus", vbOKOnly, "Validasi"
                Exit Sub
            End If

            If sp_PenyebabAskep("D") = False Then Exit Sub
    End Select

    Call cmdBatal_Click

    Exit Sub

errLoad:
    Call msubPesanError

End Sub

Private Sub cmdSimpan_Click()

    On Error GoTo errLoad

    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtDetailAskep, "Detail Askep kosong") = False Then Exit Sub
            If sp_DetailAskep("A") = False Then Exit Sub
        Case 1
            If Periksa("text", txtpenyebab, "Penyebab Askep kosong") = False Then Exit Sub
            If sp_PenyebabAskep("A") = False Then Exit Sub
    End Select

    Call cmdBatal_Click

    Exit Sub

errLoad:
    Call msubPesanError

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgDetailAskep_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDetailAskep
    WheelHook.WheelHook dgDetailAskep
End Sub

Private Sub dgDetailAskep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDetailAskep.SetFocus
End Sub

Private Sub dgKelompokPegawai_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKelompokPegawai
    WheelHook.WheelHook dgKelompokPegawai
End Sub

Private Sub dgPenyebab_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPenyebab
    WheelHook.WheelHook dgPenyebab
End Sub

Private Sub dgPenyebab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtpenyebab.SetFocus
End Sub

Private Sub dgDetailAskep_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    On Error Resume Next

    txtKdDetailAskep.Text = dgDetailAskep.Columns(0).value
    txtDetailAskep.Text = dgDetailAskep.Columns(1).value
    txtKodeExternal.Text = dgDetailAskep.Columns(2).value
    txtNamaExternal.Text = dgDetailAskep.Columns(3).value
    If dgDetailAskep.Columns(4) = "" Then
        CheckStatusEnbl.value = 0
    ElseIf dgDetailAskep.Columns(4) = 0 Then
        CheckStatusEnbl.value = 0
    ElseIf dgDetailAskep.Columns(4) = 1 Then
        CheckStatusEnbl.value = 1
    End If

End Sub

Private Sub dgPenyebab_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    On Error Resume Next

    txtkdpenyebab.Text = dgPenyebab.Columns(0).value
    txtpenyebab.Text = dgPenyebab.Columns(1).value
    txtKodeExternal1.Text = dgPenyebab.Columns(2).value
    txtNamaExternal1.Text = dgPenyebab.Columns(3).value
    If dgPenyebab.Columns(4) = "" Then
        CheckStatusEnbl1.value = 0
    ElseIf dgPenyebab.Columns(4) = 0 Then
        CheckStatusEnbl1.value = 0
    ElseIf dgPenyebab.Columns(4) = 1 Then
        CheckStatusEnbl1.value = 1
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()

    On Error GoTo errLoad

    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call openConnection
    Call clear
    Call subLoadGridSource

    SSTab1.Tab = 0
    Exit Sub

errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()

    On Error GoTo errLoad

    Select Case SSTab1.Tab
        Case 0
            Set rs = Nothing
            strSQL = "select * from DetailDiagnosaKeperawatan"
            rs.Open strSQL, dbConn, adOpenDynamic, adLockOptimistic
            Set dgDetailAskep.DataSource = rs
            With dgDetailAskep
                .Columns(0).Caption = "Kd Detail Askep"
                .Columns(0).Width = 1500
                .Columns(1).Caption = "Detail Askep"
                .Columns(1).Width = 4800
            End With
            Set rs = Nothing
        Case 1
            Set rs = Nothing
            strSQL = "select * from PenyebabDiagnosaKeperawatan"
            rs.Open strSQL, dbConn, adOpenDynamic, adLockOptimistic
            Set dgPenyebab.DataSource = rs
            With dgPenyebab
                .Columns(0).Caption = "Kode Penyebab"
                .Columns(0).Width = 1500
                .Columns(1).Caption = "Penyebab Askep"
                .Columns(1).Width = 4800
            End With
    End Select

    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub clear()

    On Error Resume Next

    Select Case SSTab1.Tab
        Case 0
            txtKdDetailAskep.Text = ""
            txtDetailAskep.Text = ""
            txtDetailAskep.SetFocus
            txtKodeExternal.Text = ""
            txtNamaExternal.Text = ""
            CheckStatusEnbl.value = 1
        Case 1
            txtkdpenyebab.Text = ""
            txtpenyebab.Text = ""
            txtpenyebab.SetFocus
            txtKodeExternal1.Text = ""
            txtNamaExternal1.Text = ""
            CheckStatusEnbl1.value = 1
    End Select

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call cmdBatal_Click
End Sub

Private Sub txtDetailAskep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal.SetFocus
End Sub

Private Sub txtKodeExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal.SetFocus
End Sub

Private Sub txtKodeExternal1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal1.SetFocus
End Sub

Private Sub txtNamaExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl.SetFocus
End Sub

Private Sub CheckStatusEnbl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdsimpan.SetFocus
End Sub

Private Sub CheckStatusEnbl1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdsimpan.SetFocus
End Sub

Private Sub txtNamaExternal1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl1.SetFocus
End Sub

Private Sub txtpenyebab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal1.SetFocus
End Sub

Private Function sp_DetailAskep(f_Status As String) As Boolean

    sp_DetailAskep = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdDetailAskep", adVarChar, adParamInput, 3, txtKdDetailAskep.Text)
        .Parameters.Append .CreateParameter("DetailAskep", adVarChar, adParamInput, 50, Trim(txtDetailAskep.Text))
        .Parameters.Append .CreateParameter("OutputKode", adVarChar, adParamOutput, 3, Null)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNamaExternal.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_DetailDiagnosaKeperawatan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            If f_Status = "A" Then
                MsgBox "Gagal menyimpan data", vbCritical, "Validasi"
            Else
                MsgBox "Gagal menghapus data", vbCritical, "Validasi"
            End If
            sp_DetailAskep = False
        End If

        If f_Status = "A" Then
            txtOuputKode.Text = .Parameters("OutputKode").value
            MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
        Else
            MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
        End If

        Call Add_HistoryLoginActivity("AUD_DetailDiagnosaKeperawatan")
        Call deleteADOCommandParameters(dbcmd)

        Set dbcmd = Nothing
    End With

End Function

Private Function sp_PenyebabAskep(f_Status As String) As Boolean

    sp_PenyebabAskep = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdPenyebab", adVarChar, adParamInput, 3, txtkdpenyebab.Text)
        .Parameters.Append .CreateParameter("PenyebabAskep", adVarChar, adParamInput, 50, Trim(txtpenyebab.Text))
        .Parameters.Append .CreateParameter("OutputKode", adVarChar, adParamOutput, 3, Null)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal1.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNamaExternal1.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl1.value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_PenyebabDiagnosaKeperawatan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            If f_Status = "A" Then
                MsgBox "Gagal menyimpan data", vbCritical, "Validasi"
            Else
                MsgBox "Gagal menghapus data", vbCritical, "Validasi"
            End If
            sp_PenyebabAskep = False
        End If

        If f_Status = "A" Then
            txtOuputKode.Text = .Parameters("OutputKode").value
        Else
            MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
        End If
        Call Add_HistoryLoginActivity("AUD_PenyebabDiagnosaKeperawatan")
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

End Function
