VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frminfokamar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Informasi Kamar Rawat Inap"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frinfokamar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   10215
   Begin VB.Frame frameDetailKmr 
      Caption         =   "Informasi Detail Kamar Rawat Inap"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   10215
      Begin VB.CommandButton cmdTutup2 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8400
         TabIndex        =   20
         Top             =   6000
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dgDetailKmrRI 
         Height          =   5535
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   9763
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
   End
   Begin VB.Frame frametarif 
      Caption         =   "Informasi Tarif Pelayanan Kamar Rawat Inap"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   10215
      Begin VB.CommandButton cmdtutuptarif 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8400
         TabIndex        =   18
         Top             =   6000
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dGridTarifKamar 
         Height          =   5535
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   9763
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
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   7080
      Width           =   10215
      Begin VB.CommandButton cmdInfoDetailKmr 
         Caption         =   "Info &Detail Kamar"
         Height          =   375
         Left            =   4560
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdInfoTarif 
         Caption         =   "&Info Tarif"
         Height          =   375
         Left            =   6480
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   8280
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6135
      Left            =   0
      TabIndex        =   13
      Top             =   960
      Width           =   10215
      Begin VB.CheckBox chkSemua 
         Caption         =   "Semua Ruangan"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid dGridKamar 
         Height          =   3615
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
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
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCariKamar 
         Caption         =   "&Cari"
         Height          =   495
         Left            =   9120
         TabIndex        =   4
         Top             =   450
         Width           =   975
      End
      Begin VB.Frame Frame5 
         Caption         =   "Status Bed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   5400
         TabIndex        =   14
         Top             =   360
         Width           =   3495
         Begin VB.OptionButton optKmrIsi 
            Caption         =   "Terisi"
            Height          =   210
            Left            =   2520
            TabIndex        =   3
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optKmrKosong 
            Caption         =   "Kosong"
            Height          =   210
            Left            =   1200
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
      End
      Begin MSDataListLib.DataCombo dtCboKlsPlynn 
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
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
      Begin MSDataListLib.DataCombo dtCboNamaSMF 
         Height          =   330
         Left            =   2520
         TabIndex        =   1
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ruang Perawatan"
         Height          =   210
         Left            =   2520
         TabIndex        =   16
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Pelayanan"
         Height          =   210
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1275
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   21
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
      Left            =   8400
      Picture         =   "frinfokamar.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frinfokamar.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frinfokamar.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frminfokamar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSemua_Click()
    If chkSemua.value = 1 Then
        Label3.Visible = False
        dtCboNamaSMF.Visible = False
    Else
        Label3.Visible = True
        dtCboNamaSMF.Visible = True
    End If
End Sub

Private Sub cmdCariKamar_Click()
    On Error GoTo errLoad
    Dim kdKelas As String
    Dim kdRu As String

    If chkSemua.value = 1 Then
        If (Me.optKmrIsi.value = True Or Me.optKmrKosong.value = True) And (Me.dtCboKlsPlynn.Text = "") Then
            MsgBox "Pilih Dulu Kelas Pelayanan !", vbOKOnly + vbExclamation, "Informasi"
            Exit Sub
        End If
    Else
        If (Me.optKmrIsi.value = True Or Me.optKmrKosong.value = True) And (Me.dtCboKlsPlynn.Text = "" Or Me.dtCboNamaSMF.Text = "") Then
            MsgBox "Pilih Dulu Kelas atau Ruang Perawatan !", vbOKOnly + vbExclamation, "Informasi"
            Exit Sub
        End If
    End If

    kdKelas = Me.dtCboKlsPlynn.BoundText
    kdRu = Me.dtCboNamaSMF.BoundText

    Set rs = New recordset
    If Me.optKmrKosong.value = True Then
        If chkSemua.value = 1 Then
            rs.Open "Select [Kelas Pelayanan],Ruangan,NoKamar,NoBed,Status,NamaPasien,JenisKelamin From V_InformasiKamarRawatInap_New Where (KdKelas = '" & kdKelas & "') and StatusEnabled='1' and Expr1='1' and Expr2='1' and Expr3='1'" _
            & " And (Status = 'Kosong')", dbConn, adOpenDynamic, adLockOptimistic
        Else
            rs.Open "Select [Kelas Pelayanan],Ruangan,NoKamar,NoBed,Status,NamaPasien,JenisKelamin From V_InformasiKamarRawatInap_New Where (KdKelas = '" & kdKelas & "') And (KdRuangan = '" & kdRu & "') and StatusEnabled='1' and Expr1='1' and Expr2='1' and Expr3='1'" _
            & " And (Status = 'Kosong')", dbConn, adOpenDynamic, adLockOptimistic
        End If
        Set Me.dGridKamar.DataSource = rs
        Call setgrid
    Else
        If chkSemua.value = 1 Then
            rs.Open "Select [Kelas Pelayanan],Ruangan,NoKamar,NoBed,Status,NamaPasien,JenisKelamin From V_InformasiKamarRawatInap_New Where (KdKelas = '" & kdKelas & "') and StatusEnabled='1' and Expr1='1' and Expr2='1' and Expr3='1'" _
            & " And (Status = 'Isi')", dbConn, adOpenDynamic, adLockOptimistic
        Else
            rs.Open "Select [Kelas Pelayanan],Ruangan,NoKamar,NoBed,Status,NamaPasien,JenisKelamin From V_InformasiKamarRawatInap_New Where (KdKelas = '" & kdKelas & "') And (KdRuangan = '" & kdRu & "') and StatusEnabled='1' and Expr1='1' and Expr2='1' and Expr3='1'" _
            & " And (Status = 'Isi')", dbConn, adOpenDynamic, adLockOptimistic
        End If
        Set Me.dGridKamar.DataSource = rs
        Call setgrid
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdInfoDetailKmr_Click()
    On Error GoTo hell
    Dim rs4 As New ADODB.recordset
    frameDetailKmr.Visible = True
    cmdInfoDetailKmr.Visible = False
    rs4.Open "select distinct Ruangan,Kelas,dbo.Ambil_JumlahKamar(KdRuangan,KdKelas) as TotalKamar,dbo.Ambil_JumlahBedIsi(Ruangan,Kelas) as BedTerisi,dbo.Ambil_JumlahBedKosong(Ruangan,Kelas) as BedKosong,dbo.Ambil_JumlahBed(Ruangan,Kelas) as TotalBed from V_RuanganKelas where StatusEnabled='1'", dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgDetailKmrRI.DataSource = rs4
    With dgDetailKmrRI
        .Columns(0).Width = 2200
        .Columns(1).Width = 1900
        .Columns(2).Alignment = dbgRight
        .Columns(2).Caption = "Jml. Kamar"
        .Columns(2).Width = 1100
        .Columns(3).Caption = "Jml. Bed Terisi"
        .Columns(3).Alignment = dbgRight
        .Columns(3).Width = 1400
        .Columns(4).Caption = "Jml. Bed Kosong"
        .Columns(4).Alignment = dbgRight
        .Columns(4).Width = 1500
        .Columns(5).Caption = "Total Bed"
        .Columns(5).Alignment = dbgRight
        .Columns(5).Width = 1000
    End With
    Exit Sub
hell:
End Sub

Private Sub cmdInfoTarif_Click()
    On Error GoTo hell
    Dim rs3 As New ADODB.recordset
    frametarif.Visible = True
    cmdInfoTarif.Visible = False
    rs3.Open "select * from V_InformasiTarifKamarRawatInap", dbConn, adOpenForwardOnly, adLockReadOnly
    Set dGridTarifKamar.DataSource = rs3
    With dGridTarifKamar
        .Columns(0).Width = 2900
        .Columns(1).Width = 2100
        .Columns(2).Width = 2100
        .Columns(4).Alignment = dbgRight
        .Columns(3).Width = 1200
        .Columns(4).Width = 800
    End With
    Exit Sub
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdTutup2_Click()
    frameDetailKmr.Visible = False
    cmdInfoDetailKmr.Visible = True
End Sub

Private Sub cmdtutuptarif_Click()
    frametarif.Visible = False
    cmdInfoTarif.Visible = True
End Sub

Private Sub dgDetailKmrRI_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDetailKmrRI
    WheelHook.WheelHook dgDetailKmrRI
End Sub

Private Sub dGridKamar_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dGridKamar
    WheelHook.WheelHook dGridKamar
End Sub

Private Sub dGridTarifKamar_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dGridTarifKamar
    WheelHook.WheelHook dGridTarifKamar
End Sub

Private Sub dtCboKlsPlynn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dtCboKlsPlynn.MatchedWithList = True Then
            If chkSemua.value = False Then
                dtCboNamaSMF.SetFocus
            Else
                optKmrKosong.SetFocus
            End If
        End If
        strSQL = "select distinct KdKelas, [Kelas Pelayanan] from V_InformasiKamarRawatInap where Expr2='1' and ([Kelas Pelayanan] LIKE '%" & dtCboKlsPlynn.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dtCboKlsPlynn.Text = ""
            Exit Sub
        End If
        dtCboKlsPlynn.BoundText = rs(0).value
        dtCboKlsPlynn.Text = rs(1).value

    End If
    
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub dtCboNamaSMF_GotFocus()
    Dim rs2 As New ADODB.recordset
    rs2.Open "select distinct Ruangan,KdRuangan from V_InformasiKamarRawatInap where KdKelas='" & Me.dtCboKlsPlynn.BoundText & "' and Expr1='1'", dbConn, adOpenForwardOnly, adLockReadOnly
    Set Me.dtCboNamaSMF.RowSource = rs2
    Me.dtCboNamaSMF.ListField = "Ruangan"
    Me.dtCboNamaSMF.BoundColumn = "KdRuangan"
End Sub

Private Sub dtCboNamaSMF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dtCboNamaSMF.MatchedWithList = True Then optKmrKosong.SetFocus
        strSQL = "select distinct kdRuangan,Ruangan from V_InformasiKamarRawatInap where KdKelas='" & Me.dtCboKlsPlynn.BoundText & "' and Expr1='1' and (Ruangan LIKE '%" & dtCboNamaSMF.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dtCboNamaSMF.Text = ""
            Exit Sub
        End If
        dtCboNamaSMF.BoundText = rs(0).value
        dtCboNamaSMF.Text = rs(1).value
    End If
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Dim rs1 As New ADODB.recordset
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call openConnection

    '//populate dtcbo klsplynn........
    rs1.Open "select distinct [Kelas Pelayanan],KdKelas from V_InformasiKamarRawatInap where Expr2='1'", dbConn, adOpenForwardOnly, adLockReadOnly
    Set Me.dtCboKlsPlynn.RowSource = rs1
    Me.dtCboKlsPlynn.ListField = "Kelas Pelayanan"
    Me.dtCboKlsPlynn.BoundColumn = "KdKelas"
    optKmrKosong.value = True

    '//populate dtcbo smf........
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
    Set rs1 = Nothing
    Set rs2 = Nothing
End Sub

Sub setgrid()
    With dGridKamar
        .Columns(0).Width = 2200
        .Columns(1).Width = 2500
        .Columns(2).Width = 1300
        .Columns(2).Alignment = dbgCenter
        .Columns(2).Caption = "   No. Kamar"
        .Columns(3).Width = 1300
        .Columns(3).Alignment = dbgCenter
        .Columns(3).Caption = "   No. Bed"
        .Columns(4).Width = 1400
        .Columns(4).Alignment = dbgCenter
        .Columns(4).Caption = "  Status Bed"
        .Columns(5).Caption = "Nama Pasien"
        .Columns(5).Width = 2000
        .Columns(6).Caption = "JK"
        .Columns(6).Width = 700
        
    End With
End Sub

Private Sub optKmrIsi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdCariKamar_Click
    End If
End Sub

Private Sub optKmrKosong_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdCariKamar_Click
    End If
End Sub

