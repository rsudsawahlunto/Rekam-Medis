VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRL4Sub4_PenyebabKecelakaanRJNew2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL4B Data Keadaan Morbiditas Pasien RJ Penyebab Kecelakaan"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6165
   Icon            =   "frmRL4Sub4_PenyebabKecelakaanRJNew2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6165
   Begin VB.Frame fraButton 
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
      TabIndex        =   0
      Top             =   3000
      Width           =   6165
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   1905
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
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
   Begin VB.Frame Frame2 
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
      Height          =   1935
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   6135
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   5895
         Begin MSComCtl2.DTPicker dtptahun 
            Height          =   375
            Left            =   1560
            TabIndex        =   6
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            CustomFormat    =   "MMMM yyyy"
            Format          =   135659523
            UpDown          =   -1  'True
            CurrentDate     =   40544
         End
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3840
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   490
      Scrolling       =   1
   End
   Begin VB.Label lblPersen 
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   3960
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL4Sub4_PenyebabKecelakaanRJNew2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmRL4Sub4_PenyebabKecelakaanRJNew2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim awal As String
Dim akhir As String

'Special Buat Excel
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel.Range
Dim oResizeRange As Excel.Range
Dim j As String
Dim xx As Integer
'Special Buat Excel

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)

    dtptahun.value = Now
    dtptahun.CustomFormat = "MMMM yyyy"
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo error

    ProgressBar1.value = ProgressBar1.Min
    lblPersen.Caption = "0 %"
    Screen.MousePointer = vbHourglass

    'Hal1
    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.path & "\RL 4B_penyakit rawat jalan(sebab).xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    For xx = 2 To 31
        With oSheet
            .Cells(xx, 3) = rsb("KdRS").value
            .Cells(xx, 2) = rsb("KotaKodyaKab").value
            .Cells(xx, 4) = rsb("NamaRS").value
            .Cells(xx, 5) = Format(dtptahun.value, "YYYY")
        End With
    Next xx

    Set rs = Nothing
    strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, NamaDTD, NoDTerperinci, isnull(sum(Kel_Umur0L), 0) as Kel_Umur0L, isnull(sum(Kel_Umur0P), 0) as Kel_Umur0P, isnull(sum(Kel_Umur1L), 0) as Kel_Umur1L,isnull(sum(Kel_Umur1P), 0) as Kel_Umur1P, isnull(sum(Kel_Umur2L), 0) as Kel_Umur2L,isnull(sum(Kel_Umur2P), 0) as Kel_Umur2P, " _
    & "isnull(sum(Kel_Umur3L), 0) as Kel_Umur3L, isnull(sum(Kel_Umur3P), 0) as Kel_Umur3P, isnull(sum(Kel_Umur4L), 0) as Kel_Umur4L, isnull(sum(Kel_Umur4P), 0) as Kel_Umur4P, isnull(sum(Kel_Umur5L), 0) as Kel_Umur5L, isnull(sum(Kel_Umur5P), 0) as Kel_Umur5P, isnull(sum(Kel_Umur6L), 0) as Kel_Umur6L, " _
    & "isnull(sum(Kel_Umur6P), 0) as Kel_Umur6P,isnull(sum(Kel_Umur7L), 0) as Kel_Umur7L, isnull(sum(Kel_Umur7P), 0) as Kel_Umur7P, isnull(sum(Kel_Umur8L), 0) as Kel_Umur8L, isnull(sum(Kel_Umur8P), 0) as Kel_Umur8P, isnull(sum(Kel_L), 0) as Kel_L, isnull(sum(Kel_P), 0) as Kel_P, isnull(sum(Kel_L), 0) + isnull(sum(Kel_P), 0) as Kel_H, isnull(sum(Kel_M), 0) AS Kel_M, " _
    & "isnull(sum(Kel_L), 0) + isnull(sum(Kel_P), 0) as Total FROM RL4_02New as a left outer join " _
    & "(SELECT Diagnosa.NoDTD from PeriksaDiagnosa inner join Diagnosa on PeriksaDiagnosa.KdDiagnosa = Diagnosa.KdDiagnosa where Month(TglPeriksa) = '" & dtptahun.Year & "' and Year(TglPeriksa) = '" & dtptahun.Year & "') as b ON a.NoDTD = b.NoDTD " _
    & "where a.qnodtd between '979' and'1008'" _
    & "Group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD "
    Call msubRecFO(rs, strSQL)

    ProgressBar1.Min = 0
    ProgressBar1.Max = rs.RecordCount
    ProgressBar1.value = 0

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        j = 2
        Call setcell
    End If

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    msubPesanError
End Sub

Private Sub setcell()
    While Not rs.EOF
        With oSheet
            .Cells(j, 10) = Trim(IIf(IsNull(rs!Kel_Umur0L.value), 0, (rs!Kel_Umur0L.value)))
            .Cells(j, 11) = Trim(IIf(IsNull(rs!Kel_Umur0P.value), 0, (rs!Kel_Umur0P.value)))
            .Cells(j, 12) = Trim(IIf(IsNull(rs!Kel_Umur1L.value), 0, (rs!Kel_Umur1L.value)))
            .Cells(j, 13) = Trim(IIf(IsNull(rs!Kel_Umur1P.value), 0, (rs!Kel_Umur1P.value)))
            .Cells(j, 14) = Trim(IIf(IsNull(rs!Kel_Umur2L.value), 0, (rs!Kel_Umur2L.value)))
            .Cells(j, 15) = Trim(IIf(IsNull(rs!Kel_Umur2P.value), 0, (rs!Kel_Umur2P.value)))
            .Cells(j, 16) = Trim(IIf(IsNull(rs!Kel_Umur3L.value), 0, (rs!Kel_Umur3L.value)))
            .Cells(j, 17) = Trim(IIf(IsNull(rs!Kel_Umur3P.value), 0, (rs!Kel_Umur3P.value)))
            .Cells(j, 18) = Trim(IIf(IsNull(rs!Kel_Umur4L.value), 0, (rs!Kel_Umur4L.value)))
            .Cells(j, 19) = Trim(IIf(IsNull(rs!Kel_Umur4P.value), 0, (rs!Kel_Umur4P.value)))
            .Cells(j, 20) = Trim(IIf(IsNull(rs!Kel_Umur5L.value), 0, (rs!Kel_Umur5L.value)))
            .Cells(j, 21) = Trim(IIf(IsNull(rs!Kel_Umur5P.value), 0, (rs!Kel_Umur5P.value)))
            .Cells(j, 22) = Trim(IIf(IsNull(rs!Kel_Umur6L.value), 0, (rs!Kel_Umur6L.value)))
            .Cells(j, 23) = Trim(IIf(IsNull(rs!Kel_Umur6P.value), 0, (rs!Kel_Umur6P.value)))
            .Cells(j, 24) = Trim(IIf(IsNull(rs!Kel_Umur7L.value), 0, (rs!Kel_Umur7L.value)))
            .Cells(j, 25) = Trim(IIf(IsNull(rs!Kel_Umur7P.value), 0, (rs!Kel_Umur7P.value)))
            .Cells(j, 26) = Trim(IIf(IsNull(rs!Kel_Umur8L.value), 0, (rs!Kel_Umur8L.value)))
            .Cells(j, 27) = Trim(IIf(IsNull(rs!Kel_Umur8P.value), 0, (rs!Kel_Umur8P.value)))
            .Cells(j, 28) = Trim(IIf(IsNull(rs!Kel_L.value), 0, (rs!Kel_L.value)))
            .Cells(j, 29) = Trim(IIf(IsNull(rs!Kel_P.value), 0, (rs!Kel_P.value)))
            .Cells(j, 30) = Trim(IIf(IsNull(rs!Kel_H.value), 0, (rs!Kel_H.value)))
            .Cells(j, 31) = Trim(IIf(IsNull(rs!Kel_M.value), 0, (rs!Kel_M.value)))
        End With
        j = j + 1

        ProgressBar1.value = Int(ProgressBar1.value) + 1
        lblPersen.Caption = Int(ProgressBar1.value * 100 / ProgressBar1.Max) & " %"

        rs.MoveNext

        If rs.EOF = True Then Exit Sub
    Wend
End Sub

