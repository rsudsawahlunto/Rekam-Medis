VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDaftarPolaKematianPenyakit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pola Kematian Menurut Penyakit"
   ClientHeight    =   8565
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
   Icon            =   "frmDaftarPolaKematianPenyakit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   14340
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
      TabIndex        =   7
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
         Left            =   11040
         TabIndex        =   8
         Top             =   120
         Width           =   3135
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
            CustomFormat    =   "MMMM yyyy"
            Format          =   61800451
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   1
            Top             =   240
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "MMMM yyyy"
            Format          =   61800451
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   9
            Top             =   322
            Visible         =   0   'False
            Width           =   255
         End
      End
   End
   Begin MSDataGridLib.DataGrid dgPasienMeninggal 
      Height          =   5535
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   14295
      _ExtentX        =   25215
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
      TabIndex        =   6
      Top             =   7560
      Width           =   14295
      Begin VB.CommandButton cmdCetak 
         Caption         =   "C&etak"
         Height          =   495
         Left            =   10800
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   12495
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   4800
      Picture         =   "frmDaftarPolaKematianPenyakit.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPolaKematianPenyakit.frx":431A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmDaftarPolaKematianPenyakit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim subTanggalTerakhir As Integer

Private Sub cmdcari_Click()
On Error GoTo hell

    Select Case dtpAkhir.Month
        Case 1, 3, 5, 7, 8, 10, 12
            subTanggalTerakhir = 31
        Case 4, 6, 9, 11
            subTanggalTerakhir = 30
        Case 2
            subTanggalTerakhir = 28
    End Select

    strSQL = "SELECT NamaDiagnosa, Sum([< 1 Tahun]) AS [< 1 Tahun], sum([1 - 4 Tahun]) AS [1 - 4 Tahun], SUM(SemuaUmur) AS SemuaUmur" & _
        " From V_PolaKematianMenurutPenyakit " & _
        " Where month(tglmeninggal) = '" & Month(dtpAwal.Value) & "' and year(tglmeninggal) = '" & Year(dtpAwal.Value) & "'" & _
        " GROUP BY NamaDiagnosa ORDER BY semuaumur DESC"
    Call msubRecFO(rs, strSQL)
    Set dgPasienMeninggal.DataSource = rs
    Call SetGridPasienMeninggal
    If dgPasienMeninggal.ApproxCount = 0 Then dtpAwal.SetFocus Else dgPasienMeninggal.SetFocus
hell:
End Sub

Private Sub cmdCetak_Click()
    cmdCetak.Enabled = False
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    
    Select Case dtpAkhir.Month
        Case 1, 3, 5, 7, 8, 10, 12
            subTanggalTerakhir = 31
        Case 4, 6, 9, 11
            subTanggalTerakhir = 30
        Case 2
            subTanggalTerakhir = 28
    End Select

    strSQL = "SELECT NamaDiagnosa, Sum([< 1 Tahun]) AS [< 1 Tahun], sum([1 - 4 Tahun]) AS [1 - 4 Tahun], SUM(SemuaUmur) AS SemuaUmur" & _
        " From V_PolaKematianMenurutPenyakit " & _
        " Where month(tglmeninggal) = '" & Month(dtpAwal.Value) & "' and year(tglmeninggal) = '" & Year(dtpAwal.Value) & "'" & _
        " GROUP BY NamaDiagnosa ORDER BY semuaumur DESC"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbExclamation, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If

    Set frmCetakPolaKematianPerPenyakit = Nothing
    frmCetakPolaKematianPerPenyakit.Show
    cmdCetak.Enabled = True
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgPasienMeninggal_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then cmdCetak.SetFocus
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
    If KeyCode = 13 Then cmdCari.SetFocus ' dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo errLoad

    Call centerForm(Me, MDIUtama)
    dtpAkhir.Value = Now
    dtpAwal.Value = Format(Now, "dd MMM yyyy 00:00:00")
    Call cmdcari_Click

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub SetGridPasienMeninggal()
    With dgPasienMeninggal
        .Columns(0).Width = 5000
        .Columns(1).Width = 1200
        .Columns(2).Width = 1200
        .Columns(3).Width = 1200
    End With
End Sub

Private Sub txtParameter_Change()
'    Call cmdCari_Click
End Sub

Private Sub txtParameter_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        Call cmdcari_Click
        txtParameter.SetFocus
    End If
End Sub

Private Sub txtParameter_LostFocus()
    txtParameter.Text = StrConv(txtParameter.Text, vbProperCase)
End Sub
