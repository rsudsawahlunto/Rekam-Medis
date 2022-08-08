VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRL4Sub4_aNew2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL4A Data Keadaan Morbiditas Pasien RI"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6165
   Icon            =   "frmRL4Sub4_aNew2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
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
      TabIndex        =   5
      Top             =   3120
      Width           =   6165
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   1905
      End
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
      TabIndex        =   0
      Top             =   1080
      Width           =   6135
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5895
         Begin MSComCtl2.DTPicker dtpTahunAkhir 
            Height          =   375
            Left            =   3240
            TabIndex        =   2
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
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
            CustomFormat    =   "dd MMM yyyy"
            Format          =   119209987
            UpDown          =   -1  'True
            CurrentDate     =   40544
         End
         Begin MSComCtl2.DTPicker dtpTahunAwal 
            Height          =   375
            Left            =   720
            TabIndex        =   9
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
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
            CustomFormat    =   "dd MMM yyyy"
            Format          =   3801091
            UpDown          =   -1  'True
            CurrentDate     =   40544
         End
         Begin VB.Label Label1 
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   10
            Top             =   600
            Width           =   255
         End
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   490
      Scrolling       =   1
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   8
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
      Left            =   5520
      TabIndex        =   4
      Top             =   3600
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL4Sub4_aNew2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmRL4Sub4_aNew2"
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

    dtptahunawal.value = Now
    dtpTahunAkhir.value = Now
'    dtpTahunAwal.CustomFormat = "MMM yyyyy"
'    dtpTahunAkhir.CustomFormat = "MMM yyyyy"
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo error

    Set oXL = CreateObject("Excel.Application")
    Set oWB = oXL.Workbooks.Open(App.path & "\RL 4.A Penyakit Rawat Inap.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    For xx = 2 To 508
        With oSheet
            .Cells(xx, 3) = rsb("KdRS").value
            .Cells(xx, 2) = rsb("KotaKodyaKab").value
            .Cells(xx, 4) = rsb("NamaRS").value
            .Cells(xx, 5) = Format(dtptahunawal.value, "YYYY")
        End With
    Next xx

    Set rs = Nothing
'    strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, NamaDTD, NoDTerperinci, isnull(sum(Kel_Umur0L), 0) as Kel_Umur0L, isnull(sum(Kel_Umur0P), 0) as Kel_Umur0P, isnull(sum(Kel_Umur1L), 0) as Kel_Umur1L,isnull(sum(Kel_Umur1P), 0) as Kel_Umur1P, isnull(sum(Kel_Umur2L), 0) as Kel_Umur2L,isnull(sum(Kel_Umur2P), 0) as Kel_Umur2P, " _
'    & "isnull(sum(Kel_Umur3L), 0) as Kel_Umur3L, isnull(sum(Kel_Umur3P), 0) as Kel_Umur3P, isnull(sum(Kel_Umur4L), 0) as Kel_Umur4L, isnull(sum(Kel_Umur4P), 0) as Kel_Umur4P, isnull(sum(Kel_Umur5L), 0) as Kel_Umur5L, isnull(sum(Kel_Umur5P), 0) as Kel_Umur5P, isnull(sum(Kel_Umur6L), 0) as Kel_Umur6L, " _
'    & "isnull(sum(Kel_Umur6P), 0) as Kel_Umur6P,isnull(sum(Kel_Umur7L), 0) as Kel_Umur7L, isnull(sum(Kel_Umur7P), 0) as Kel_Umur7P, isnull(sum(Kel_Umur8L), 0) as Kel_Umur8L, isnull(sum(Kel_Umur8P), 0) as Kel_Umur8P, isnull(sum(Kel_L), 0) as Kel_L, isnull(sum(Kel_P), 0) as Kel_P, isnull(sum(Kel_L), 0) + isnull(sum(Kel_P), 0) as Kel_H, isnull(sum(Kel_M), 0) AS Kel_M, " _
'    & "isnull(sum(Kel_L), 0) + isnull(sum(Kel_P), 0) as Total FROM RL4_01New as a left outer join " _
'    & "(SELECT Diagnosa.NoDTD from PeriksaDiagnosa inner join Diagnosa on PeriksaDiagnosa.KdDiagnosa = Diagnosa.KdDiagnosa where Month(TglPeriksa) = '" & dtptahun.Month & "' AND Year(TglPeriksa) = '" & dtptahun.Year & "') as b ON a.NoDTD = b.NoDTD " _
'    & "where a.qnodtd between '482' and'978'" _
'    & "Group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD "

strSQL = "SELECT a.NoDTD, a.QNoDTD, 'Grup' = case when a.NoDTD < = '298' then '0' else '1' end, NamaDTD, NoDTerperinci, isnull(sum(Kel_Umur0L), 0) as Kel_Umur0L, isnull(sum(Kel_Umur0P), 0) as Kel_Umur0P, isnull(sum(Kel_Umur1L), 0) as Kel_Umur1L,isnull(sum(Kel_Umur1P), 0) as Kel_Umur1P, isnull(sum(Kel_Umur2L), 0) as Kel_Umur2L,isnull(sum(Kel_Umur2P), 0) as Kel_Umur2P, " _
    & "isnull(sum(Kel_Umur3L), 0) as Kel_Umur3L, isnull(sum(Kel_Umur3P), 0) as Kel_Umur3P, isnull(sum(Kel_Umur4L), 0) as Kel_Umur4L, isnull(sum(Kel_Umur4P), 0) as Kel_Umur4P, isnull(sum(Kel_Umur5L), 0) as Kel_Umur5L, isnull(sum(Kel_Umur5P), 0) as Kel_Umur5P, isnull(sum(Kel_Umur6L), 0) as Kel_Umur6L, " _
    & "isnull(sum(Kel_Umur6P), 0) as Kel_Umur6P,isnull(sum(Kel_Umur7L), 0) as Kel_Umur7L, isnull(sum(Kel_Umur7P), 0) as Kel_Umur7P, isnull(sum(Kel_Umur8L), 0) as Kel_Umur8L, isnull(sum(Kel_Umur8P), 0) as Kel_Umur8P, isnull(sum(Kel_L), 0) as Kel_L, isnull(sum(Kel_P), 0) as Kel_P, isnull(sum(Kel_L), 0) + isnull(sum(Kel_P), 0) as Kel_H, isnull(sum(Kel_M), 0) AS Kel_M, " _
    & "isnull(sum(Kel_L), 0) + isnull(sum(Kel_P), 0) as Total FROM RL4_01New as a " _
    & "where a.qnodtd between '482' and'978' and (TglPeriksa Between '" & Format(dtptahunawal.value, "yyyy-mm-dd 00:00:00") & "' AND '" & Format(dtpTahunAkhir.value, "yyyy-mm-dd 23:59:59") & "') " _
    & "Group by a.NoDTD, a.NamaDTD, a.NoDTerperinci, a.QNoDTD "
    Call msubRecFO(rs, strSQL)
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
'        If rs!qnodtd = "482" Then j = 2
'        If rs!qnodtd = "483" Then j = 3
'        If rs!qnodtd = "484" Then j = 4
'        If rs!qnodtd = "485" Then j = 5
'        If rs!qnodtd = "486" Then j = 6
'        If rs!qnodtd = "487" Then j = 7
'        If rs!qnodtd = "488" Then j = 8
'        If rs!qnodtd = "489" Then j = 9
'        If rs!qnodtd = "490" Then j = 10
'        If rs!qnodtd = "491" Then j = 11
'        If rs!qnodtd = "492" Then j = 12
'        If rs!qnodtd = "493" Then j = 13
'        If rs!qnodtd = "494" Then j = 14
'        If rs!qnodtd = "495" Then j = 15
'        If rs!qnodtd = "496" Then j = 16
'        If rs!qnodtd = "497" Then j = 17
'        If rs!qnodtd = "498" Then j = 18
'        If rs!qnodtd = "499" Then j = 19
'        If rs!qnodtd = "500" Then j = 20
'        If rs!qnodtd = "501" Then j = 21
'        If rs!qnodtd = "501" Then j = 21
'        If rs!qnodtd = "502" Then j = 22
'        If rs!qnodtd = "503" Then j = 23
'        If rs!qnodtd = "504" Then j = 24
'        If rs!qnodtd = "505" Then j = 25
'        If rs!qnodtd = "506" Then j = 26
'        If rs!qnodtd = "507" Then j = 27
'        If rs!qnodtd = "508" Then j = 28
'        If rs!qnodtd = "509" Then j = 29
'        If rs!qnodtd = "510" Then j = 30
'        If rs!qnodtd = "511" Then j = 31
'        If rs!qnodtd = "512" Then j = 32
'        If rs!qnodtd = "513" Then j = 33
'        If rs!qnodtd = "514" Then j = 34
'        If rs!qnodtd = "515" Then j = 35
'        If rs!qnodtd = "516" Then j = 36
'        If rs!qnodtd = "517" Then j = 37
'        If rs!qnodtd = "518" Then j = 38
'        If rs!qnodtd = "519" Then j = 39
'        If rs!qnodtd = "520" Then j = 40
'        If rs!qnodtd = "521" Then j = 41
'        If rs!qnodtd = "522" Then j = 42
'        If rs!qnodtd = "523" Then j = 43
'        If rs!qnodtd = "524" Then j = 44
'        If rs!qnodtd = "525" Then j = 45
'        If rs!qnodtd = "526" Then j = 46
'        If rs!qnodtd = "527" Then j = 47
'        If rs!qnodtd = "528" Then j = 48
'        If rs!qnodtd = "529" Then j = 49
'        If rs!qnodtd = "530" Then j = 50
'        If rs!qnodtd = "531" Then j = 51
'        If rs!qnodtd = "532" Then j = 52
'        If rs!qnodtd = "533" Then j = 53
'        If rs!qnodtd = "534" Then j = 54
'        If rs!qnodtd = "535" Then j = 55
'        If rs!qnodtd = "536" Then j = 56
'        If rs!qnodtd = "537" Then j = 57
'        If rs!qnodtd = "538" Then j = 58
'        If rs!qnodtd = "539" Then j = 59
'        If rs!qnodtd = "540" Then j = 60
'        If rs!qnodtd = "541" Then j = 61
'        If rs!qnodtd = "542" Then j = 62
'        If rs!qnodtd = "543" Then j = 63
'        If rs!qnodtd = "544" Then j = 64
'        If rs!qnodtd = "545" Then j = 65
'        If rs!qnodtd = "546" Then j = 66
'        If rs!qnodtd = "547" Then j = 67
'        If rs!qnodtd = "548" Then j = 68
'        If rs!qnodtd = "549" Then j = 69
'        If rs!qnodtd = "550" Then j = 70
'        If rs!qnodtd = "551" Then j = 71
'        If rs!qnodtd = "552" Then j = 72
'        If rs!qnodtd = "553" Then j = 73
'        If rs!qnodtd = "554" Then j = 74
'        If rs!qnodtd = "555" Then j = 75
'        If rs!qnodtd = "556" Then j = 76
'        If rs!qnodtd = "557" Then j = 77
'        If rs!qnodtd = "558" Then j = 78
'        If rs!qnodtd = "559" Then j = 79
'        If rs!qnodtd = "560" Then j = 80
'        If rs!qnodtd = "561" Then j = 81
'        If rs!qnodtd = "562" Then j = 82
'        If rs!qnodtd = "563" Then j = 83
'        If rs!qnodtd = "564" Then j = 84
'        If rs!qnodtd = "565" Then j = 85
'        If rs!qnodtd = "566" Then j = 86
'        If rs!qnodtd = "567" Then j = 87
'        If rs!qnodtd = "568" Then j = 88
'        If rs!qnodtd = "569" Then j = 89
'        If rs!qnodtd = "570" Then j = 90
'        If rs!qnodtd = "571" Then j = 91
'        If rs!qnodtd = "572" Then j = 92
'        If rs!qnodtd = "573" Then j = 93
'        If rs!qnodtd = "574" Then j = 94
'        If rs!qnodtd = "575" Then j = 95
'        If rs!qnodtd = "576" Then j = 96
'        If rs!qnodtd = "577" Then j = 97
'        If rs!qnodtd = "578" Then j = 98
'        If rs!qnodtd = "579" Then j = 99
'        If rs!qnodtd = "580" Then j = 100
'        If rs!qnodtd = "581" Then j = 101
'        If rs!qnodtd = "582" Then j = 102
'        If rs!qnodtd = "583" Then j = 103
'        If rs!qnodtd = "584" Then j = 104
'        If rs!qnodtd = "585" Then j = 105
'        If rs!qnodtd = "586" Then j = 106
'        If rs!qnodtd = "587" Then j = 107
'        If rs!qnodtd = "588" Then j = 108
'        If rs!qnodtd = "589" Then j = 109
'        If rs!qnodtd = "590" Then j = 110
'        If rs!qnodtd = "591" Then j = 111
'        If rs!qnodtd = "592" Then j = 112
'        If rs!qnodtd = "593" Then j = 113
'        If rs!qnodtd = "594" Then j = 114
'        If rs!qnodtd = "595" Then j = 115
'        If rs!qnodtd = "596" Then j = 116
'        If rs!qnodtd = "597" Then j = 117
'        If rs!qnodtd = "598" Then j = 118
'        If rs!qnodtd = "599" Then j = 119
'        If rs!qnodtd = "600" Then j = 120
'        If rs!qnodtd = "601" Then j = 121
'        If rs!qnodtd = "602" Then j = 122
'        If rs!qnodtd = "603" Then j = 123
'        If rs!qnodtd = "604" Then j = 124
'        If rs!qnodtd = "605" Then j = 125
'        If rs!qnodtd = "606" Then j = 126
'        If rs!qnodtd = "607" Then j = 127
'        If rs!qnodtd = "608" Then j = 128
'        If rs!qnodtd = "609" Then j = 129
'        If rs!qnodtd = "610" Then j = 130
'        If rs!qnodtd = "611" Then j = 131
'        If rs!qnodtd = "612" Then j = 132
'        If rs!qnodtd = "613" Then j = 133
'        If rs!qnodtd = "614" Then j = 134
'        If rs!qnodtd = "615" Then j = 135
'        If rs!qnodtd = "616" Then j = 136
'        If rs!qnodtd = "617" Then j = 137
'        If rs!qnodtd = "618" Then j = 138
'        If rs!qnodtd = "619" Then j = 139
'        If rs!qnodtd = "620" Then j = 140
'        If rs!qnodtd = "621" Then j = 141
'        If rs!qnodtd = "622" Then j = 142
'        If rs!qnodtd = "623" Then j = 143
'        If rs!qnodtd = "624" Then j = 144
'        If rs!qnodtd = "625" Then j = 145
'        If rs!qnodtd = "626" Then j = 146
'        If rs!qnodtd = "627" Then j = 147
'        If rs!qnodtd = "628" Then j = 148
'        If rs!qnodtd = "629" Then j = 149
'        If rs!qnodtd = "630" Then j = 150
'        If rs!qnodtd = "631" Then j = 151
'        If rs!qnodtd = "632" Then j = 152
'        If rs!qnodtd = "633" Then j = 153
'        If rs!qnodtd = "634" Then j = 154
'        If rs!qnodtd = "635" Then j = 155
'        If rs!qnodtd = "636" Then j = 156
'        If rs!qnodtd = "637" Then j = 157
'        If rs!qnodtd = "638" Then j = 158
'        If rs!qnodtd = "639" Then j = 159
'        If rs!qnodtd = "640" Then j = 160
'        If rs!qnodtd = "641" Then j = 161
'        If rs!qnodtd = "642" Then j = 162
'        If rs!qnodtd = "643" Then j = 163
'        If rs!qnodtd = "644" Then j = 164
'        If rs!qnodtd = "645" Then j = 165
'        If rs!qnodtd = "646" Then j = 166
'        If rs!qnodtd = "647" Then j = 167
'        If rs!qnodtd = "648" Then j = 168
'        If rs!qnodtd = "649" Then j = 169
'        If rs!qnodtd = "650" Then j = 170
'        If rs!qnodtd = "651" Then j = 171
'        If rs!qnodtd = "652" Then j = 172
'        If rs!qnodtd = "653" Then j = 173
'        If rs!qnodtd = "654" Then j = 174
'        If rs!qnodtd = "655" Then j = 175
'        If rs!qnodtd = "656" Then j = 176
'        If rs!qnodtd = "657" Then j = 177
'        If rs!qnodtd = "658" Then j = 178
'        If rs!qnodtd = "659" Then j = 179
'        If rs!qnodtd = "660" Then j = 180
'        If rs!qnodtd = "661" Then j = 181
'        If rs!qnodtd = "662" Then j = 182
'        If rs!qnodtd = "663" Then j = 183
'        If rs!qnodtd = "664" Then j = 184
'        If rs!qnodtd = "665" Then j = 185
'        If rs!qnodtd = "666" Then j = 186
'        If rs!qnodtd = "667" Then j = 187
'        If rs!qnodtd = "668" Then j = 188
'        If rs!qnodtd = "669" Then j = 189
'        If rs!qnodtd = "670" Then j = 190
'        If rs!qnodtd = "671" Then j = 191
'        If rs!qnodtd = "672" Then j = 192
'        If rs!qnodtd = "673" Then j = 193
'        If rs!qnodtd = "674" Then j = 194
'        If rs!qnodtd = "675" Then j = 195
'        If rs!qnodtd = "676" Then j = 196
'        If rs!qnodtd = "677" Then j = 197
'        If rs!qnodtd = "678" Then j = 198
'        If rs!qnodtd = "679" Then j = 199
'        If rs!qnodtd = "680" Then j = 200
'        If rs!qnodtd = "681" Then j = 201
'        If rs!qnodtd = "682" Then j = 202
'        If rs!qnodtd = "683" Then j = 203
'        If rs!qnodtd = "684" Then j = 204
'        If rs!qnodtd = "685" Then j = 205
'        If rs!qnodtd = "686" Then j = 206
'        If rs!qnodtd = "687" Then j = 207
'        If rs!qnodtd = "688" Then j = 208
'        If rs!qnodtd = "689" Then j = 209
'        If rs!qnodtd = "690" Then j = 210
'        If rs!qnodtd = "691" Then j = 211
'        If rs!qnodtd = "692" Then j = 212
'        If rs!qnodtd = "693" Then j = 213
'        If rs!qnodtd = "694" Then j = 214
'        If rs!qnodtd = "695" Then j = 215
'        If rs!qnodtd = "696" Then j = 216
'        If rs!qnodtd = "697" Then j = 217
'        If rs!qnodtd = "698" Then j = 218
'        If rs!qnodtd = "699" Then j = 219
'        If rs!qnodtd = "700" Then j = 220
'        If rs!qnodtd = "701" Then j = 221
'        If rs!qnodtd = "702" Then j = 222
'        If rs!qnodtd = "703" Then j = 223
'        If rs!qnodtd = "704" Then j = 224
'        If rs!qnodtd = "705" Then j = 225
'        If rs!qnodtd = "706" Then j = 226
'        If rs!qnodtd = "707" Then j = 227
'        If rs!qnodtd = "708" Then j = 228
'        If rs!qnodtd = "709" Then j = 229
'        If rs!qnodtd = "710" Then j = 230
'        If rs!qnodtd = "711" Then j = 231
'        If rs!qnodtd = "712" Then j = 232
'        If rs!qnodtd = "713" Then j = 233
'        If rs!qnodtd = "714" Then j = 234
'        If rs!qnodtd = "715" Then j = 235
'        If rs!qnodtd = "716" Then j = 236
'        If rs!qnodtd = "717" Then j = 237
'        If rs!qnodtd = "718" Then j = 238
'        If rs!qnodtd = "719" Then j = 239
'        If rs!qnodtd = "720" Then j = 240
'        If rs!qnodtd = "721" Then j = 241
'        If rs!qnodtd = "722" Then j = 242
'        If rs!qnodtd = "723" Then j = 243
'        If rs!qnodtd = "724" Then j = 244
'        If rs!qnodtd = "725" Then j = 245
'        If rs!qnodtd = "726" Then j = 246
'        If rs!qnodtd = "727" Then j = 247
'        If rs!qnodtd = "728" Then j = 248
'        If rs!qnodtd = "729" Then j = 249
'        If rs!qnodtd = "730" Then j = 250
'        If rs!qnodtd = "731" Then j = 251
'        If rs!qnodtd = "732" Then j = 252
'        If rs!qnodtd = "733" Then j = 253
'        If rs!qnodtd = "734" Then j = 254
'        If rs!qnodtd = "735" Then j = 255
'        If rs!qnodtd = "736" Then j = 256
'        If rs!qnodtd = "737" Then j = 257
'        If rs!qnodtd = "738" Then j = 258
'        If rs!qnodtd = "739" Then j = 259
'        If rs!qnodtd = "740" Then j = 260
'        If rs!qnodtd = "741" Then j = 261
'        If rs!qnodtd = "742" Then j = 262
'        If rs!qnodtd = "743" Then j = 263
'        If rs!qnodtd = "744" Then j = 264
'        If rs!qnodtd = "745" Then j = 265
'        If rs!qnodtd = "746" Then j = 266
'        If rs!qnodtd = "747" Then j = 267
'        If rs!qnodtd = "748" Then j = 268
'        If rs!qnodtd = "749" Then j = 269
'        If rs!qnodtd = "750" Then j = 270
'        If rs!qnodtd = "751" Then j = 271
'        If rs!qnodtd = "752" Then j = 272
'        If rs!qnodtd = "753" Then j = 273
'        If rs!qnodtd = "754" Then j = 274
'        If rs!qnodtd = "755" Then j = 275
'        If rs!qnodtd = "756" Then j = 276
'        If rs!qnodtd = "757" Then j = 277
'        If rs!qnodtd = "758" Then j = 278
'        If rs!qnodtd = "759" Then j = 279
'        If rs!qnodtd = "760" Then j = 280
'        If rs!qnodtd = "761" Then j = 281
'        If rs!qnodtd = "762" Then j = 282
'        If rs!qnodtd = "763" Then j = 283
'        If rs!qnodtd = "764" Then j = 284
'        If rs!qnodtd = "765" Then j = 285
'        If rs!qnodtd = "766" Then j = 286
'        If rs!qnodtd = "767" Then j = 287
'        If rs!qnodtd = "768" Then j = 288
'        If rs!qnodtd = "769" Then j = 289
'        If rs!qnodtd = "770" Then j = 290
'        If rs!qnodtd = "771" Then j = 291
'        If rs!qnodtd = "772" Then j = 292
'        If rs!qnodtd = "773" Then j = 293
'        If rs!qnodtd = "774" Then j = 294
'        If rs!qnodtd = "775" Then j = 295
'        If rs!qnodtd = "776" Then j = 296
'        If rs!qnodtd = "777" Then j = 297
'        If rs!qnodtd = "778" Then j = 298
'        If rs!qnodtd = "779" Then j = 299
'        If rs!qnodtd = "780" Then j = 300
'        If rs!qnodtd = "781" Then j = 301
'        If rs!qnodtd = "782" Then j = 302
'        If rs!qnodtd = "783" Then j = 303
'        If rs!qnodtd = "784" Then j = 304
'        If rs!qnodtd = "785" Then j = 305
'        If rs!qnodtd = "786" Then j = 306
'        If rs!qnodtd = "787" Then j = 307
'        If rs!qnodtd = "788" Then j = 308
'        If rs!qnodtd = "789" Then j = 309
'        If rs!qnodtd = "790" Then j = 310
'        If rs!qnodtd = "791" Then j = 311
'        If rs!qnodtd = "792" Then j = 312
'        If rs!qnodtd = "793" Then j = 313
'        If rs!qnodtd = "794" Then j = 314
'        If rs!qnodtd = "795" Then j = 315
'        If rs!qnodtd = "796" Then j = 316
'        If rs!qnodtd = "797" Then j = 317
'        If rs!qnodtd = "798" Then j = 318
'        If rs!qnodtd = "799" Then j = 319
'        If rs!qnodtd = "800" Then j = 320
'        If rs!qnodtd = "801" Then j = 321
'        If rs!qnodtd = "802" Then j = 322
'        If rs!qnodtd = "803" Then j = 323
'        If rs!qnodtd = "804" Then j = 324
'        If rs!qnodtd = "805" Then j = 325
'        If rs!qnodtd = "806" Then j = 326
'        If rs!qnodtd = "807" Then j = 327
'        If rs!qnodtd = "808" Then j = 328
'        If rs!qnodtd = "809" Then j = 329
'        If rs!qnodtd = "810" Then j = 330
'        If rs!qnodtd = "811" Then j = 331
'        If rs!qnodtd = "812" Then j = 332
'        If rs!qnodtd = "813" Then j = 333
'        If rs!qnodtd = "814" Then j = 334
'        If rs!qnodtd = "815" Then j = 335
'        If rs!qnodtd = "816" Then j = 336
'        If rs!qnodtd = "817" Then j = 337
'        If rs!qnodtd = "818" Then j = 338
'        If rs!qnodtd = "819" Then j = 339
'        If rs!qnodtd = "820" Then j = 340
'        If rs!qnodtd = "821" Then j = 341
'        If rs!qnodtd = "822" Then j = 342
'        If rs!qnodtd = "823" Then j = 343
'        If rs!qnodtd = "824" Then j = 344
'        If rs!qnodtd = "825" Then j = 345
'        If rs!qnodtd = "826" Then j = 346
'        If rs!qnodtd = "827" Then j = 347
'        If rs!qnodtd = "828" Then j = 348
'        If rs!qnodtd = "829" Then j = 349
'        If rs!qnodtd = "830" Then j = 350
'        If rs!qnodtd = "831" Then j = 351
'        If rs!qnodtd = "832" Then j = 352
'        If rs!qnodtd = "833" Then j = 353
'        If rs!qnodtd = "834" Then j = 354
'        If rs!qnodtd = "835" Then j = 355
'        If rs!qnodtd = "836" Then j = 356
'        If rs!qnodtd = "837" Then j = 357
'        If rs!qnodtd = "838" Then j = 358
'        If rs!qnodtd = "839" Then j = 359
'        If rs!qnodtd = "840" Then j = 360
'        If rs!qnodtd = "841" Then j = 361
'        If rs!qnodtd = "842" Then j = 362
'        If rs!qnodtd = "843" Then j = 363
'        If rs!qnodtd = "844" Then j = 364
'        If rs!qnodtd = "845" Then j = 365
'        If rs!qnodtd = "846" Then j = 366
'        If rs!qnodtd = "847" Then j = 367
'        If rs!qnodtd = "848" Then j = 368
'        If rs!qnodtd = "849" Then j = 369
'        If rs!qnodtd = "850" Then j = 370
'        If rs!qnodtd = "851" Then j = 371
'        If rs!qnodtd = "852" Then j = 372
'        If rs!qnodtd = "853" Then j = 373
'        If rs!qnodtd = "854" Then j = 374
'        If rs!qnodtd = "855" Then j = 375
'        If rs!qnodtd = "856" Then j = 376
'        If rs!qnodtd = "857" Then j = 377
'        If rs!qnodtd = "858" Then j = 378
'        If rs!qnodtd = "859" Then j = 379
'        If rs!qnodtd = "860" Then j = 380
'        If rs!qnodtd = "861" Then j = 381
'        If rs!qnodtd = "862" Then j = 382
'        If rs!qnodtd = "863" Then j = 383
'        If rs!qnodtd = "864" Then j = 384
'        If rs!qnodtd = "865" Then j = 385
'        If rs!qnodtd = "866" Then j = 386
'        If rs!qnodtd = "867" Then j = 387
'        If rs!qnodtd = "868" Then j = 388
'        If rs!qnodtd = "869" Then j = 389
'        If rs!qnodtd = "870" Then j = 390
'        If rs!qnodtd = "871" Then j = 391
'        If rs!qnodtd = "872" Then j = 392
'        If rs!qnodtd = "873" Then j = 393
'        If rs!qnodtd = "874" Then j = 394
'        If rs!qnodtd = "875" Then j = 395
'        If rs!qnodtd = "876" Then j = 396
'        If rs!qnodtd = "877" Then j = 397
'        If rs!qnodtd = "878" Then j = 398
'        If rs!qnodtd = "879" Then j = 399
'        If rs!qnodtd = "880" Then j = 400
'        If rs!qnodtd = "881" Then j = 401
'        If rs!qnodtd = "882" Then j = 402
'        If rs!qnodtd = "883" Then j = 403
'        If rs!qnodtd = "884" Then j = 404
'        If rs!qnodtd = "885" Then j = 405
'        If rs!qnodtd = "886" Then j = 406
'        If rs!qnodtd = "887" Then j = 407
'        If rs!qnodtd = "888" Then j = 408
'        If rs!qnodtd = "889" Then j = 409
'        If rs!qnodtd = "890" Then j = 410
'        If rs!qnodtd = "891" Then j = 411
'        If rs!qnodtd = "892" Then j = 412
'        If rs!qnodtd = "893" Then j = 413
'        If rs!qnodtd = "894" Then j = 414
'        If rs!qnodtd = "895" Then j = 415
'        If rs!qnodtd = "896" Then j = 416
'        If rs!qnodtd = "897" Then j = 417
'        If rs!qnodtd = "898" Then j = 418
'        If rs!qnodtd = "899" Then j = 419
'        If rs!qnodtd = "900" Then j = 420
'        If rs!qnodtd = "901" Then j = 421
'        If rs!qnodtd = "902" Then j = 422
'        If rs!qnodtd = "903" Then j = 423
'        If rs!qnodtd = "904" Then j = 424
'        If rs!qnodtd = "905" Then j = 425
'        If rs!qnodtd = "906" Then j = 426
'        If rs!qnodtd = "907" Then j = 427
'        If rs!qnodtd = "908" Then j = 428
'        If rs!qnodtd = "909" Then j = 429
'        If rs!qnodtd = "910" Then j = 430
'        If rs!qnodtd = "911" Then j = 431
'        If rs!qnodtd = "912" Then j = 432
'        If rs!qnodtd = "913" Then j = 433
'        If rs!qnodtd = "914" Then j = 434
'        If rs!qnodtd = "915" Then j = 435
'        If rs!qnodtd = "916" Then j = 436
'        If rs!qnodtd = "917" Then j = 437
'        If rs!qnodtd = "918" Then j = 438
'        If rs!qnodtd = "919" Then j = 439
'        If rs!qnodtd = "920" Then j = 440
'        If rs!qnodtd = "921" Then j = 441
'        If rs!qnodtd = "922" Then j = 442
'        If rs!qnodtd = "923" Then j = 443
'        If rs!qnodtd = "924" Then j = 444
'        If rs!qnodtd = "925" Then j = 445
'        If rs!qnodtd = "926" Then j = 446
'        If rs!qnodtd = "927" Then j = 447
'        If rs!qnodtd = "928" Then j = 448
'        If rs!qnodtd = "929" Then j = 449
'        If rs!qnodtd = "930" Then j = 450
'        If rs!qnodtd = "931" Then j = 451
'        If rs!qnodtd = "932" Then j = 452
'        If rs!qnodtd = "933" Then j = 453
'        If rs!qnodtd = "934" Then j = 454
'        If rs!qnodtd = "935" Then j = 455
'        If rs!qnodtd = "936" Then j = 456
'        If rs!qnodtd = "937" Then j = 457
'        If rs!qnodtd = "938" Then j = 458
'        If rs!qnodtd = "939" Then j = 459
'        If rs!qnodtd = "940" Then j = 460
'        If rs!qnodtd = "941" Then j = 461
'        If rs!qnodtd = "942" Then j = 462
'        If rs!qnodtd = "943" Then j = 463
'        If rs!qnodtd = "944" Then j = 464
'        If rs!qnodtd = "945" Then j = 465
'        If rs!qnodtd = "946" Then j = 466
'        If rs!qnodtd = "947" Then j = 467
'        If rs!qnodtd = "948" Then j = 468
'        If rs!qnodtd = "949" Then j = 469
'        If rs!qnodtd = "950" Then j = 470
'        If rs!qnodtd = "951" Then j = 471
'        If rs!qnodtd = "952" Then j = 472
'        If rs!qnodtd = "953" Then j = 473
'        If rs!qnodtd = "954" Then j = 474
'        If rs!qnodtd = "955" Then j = 475
'        If rs!qnodtd = "956" Then j = 476
'        If rs!qnodtd = "957" Then j = 477
'        If rs!qnodtd = "958" Then j = 478
'        If rs!qnodtd = "959" Then j = 479
'        If rs!qnodtd = "960" Then j = 480
'        If rs!qnodtd = "961" Then j = 481
'        If rs!qnodtd = "962" Then j = 482
'        If rs!qnodtd = "963" Then j = 483
'        If rs!qnodtd = "964" Then j = 484
'        If rs!qnodtd = "965" Then j = 485
'        If rs!qnodtd = "966" Then j = 486
'        If rs!qnodtd = "967" Then j = 487
'        If rs!qnodtd = "968" Then j = 488
'        If rs!qnodtd = "969" Then j = 489
'        If rs!qnodtd = "970" Then j = 490
'        If rs!qnodtd = "971" Then j = 491
'        If rs!qnodtd = "972" Then j = 492
'        If rs!qnodtd = "973" Then j = 493
'        If rs!qnodtd = "974" Then j = 494
'        If rs!qnodtd = "975" Then j = 495
'        If rs!qnodtd = "976" Then j = 496
'        If rs!qnodtd = "977" Then j = 497
'        If rs!qnodtd = "978" Then j = 498
'        If rs!qnodtd = "979" Then j = 499
'        If rs!qnodtd = "980" Then j = 500
'        If rs!qnodtd = "981" Then j = 501
'        If rs!qnodtd = "982" Then j = 502
'        If rs!qnodtd = "983" Then j = 503
'        If rs!qnodtd = "984" Then j = 504
'        If rs!qnodtd = "985" Then j = 505
'        If rs!qnodtd = "986" Then j = 506
'        If rs!qnodtd = "987" Then j = 507
'        If rs!qnodtd = "988" Then j = 508
'        If rs!qnodtd = "989" Then j = 509
'        If rs!qnodtd = "990" Then j = 510
'        If rs!qnodtd = "991" Then j = 511
'        If rs!qnodtd = "992" Then j = 512
'        If rs!qnodtd = "993" Then j = 513
'        If rs!qnodtd = "994" Then j = 514
'        If rs!qnodtd = "995" Then j = 515
'        If rs!qnodtd = "996" Then j = 516
'        If rs!qnodtd = "997" Then j = 517
'        If rs!qnodtd = "998" Then j = 518
'        If rs!qnodtd = "999" Then j = 519
'        If rs!qnodtd = "1000" Then j = 520
'        If rs!qnodtd = "1001" Then j = 521
'        If rs!qnodtd = "1002" Then j = 522
'        If rs!qnodtd = "1003" Then j = 523
'        If rs!qnodtd = "1004" Then j = 524
'        If rs!qnodtd = "1005" Then j = 525
'        If rs!qnodtd = "1006" Then j = 526
'        If rs!qnodtd = "1007" Then j = 527

        If rs!qnodtd = "482" Then j = 2
        If rs!qnodtd = "483" Then j = 3
        If rs!qnodtd = "484" Then j = 4
        If rs!qnodtd = "485" Then j = 5
        If rs!qnodtd = "486" Then j = 6
        If rs!qnodtd = "487" Then j = 7
        If rs!qnodtd = "488" Then j = 8
        If rs!qnodtd = "489" Then j = 9
        If rs!qnodtd = "490" Then j = 10
        If rs!qnodtd = "491" Then j = 11
        If rs!qnodtd = "492" Then j = 12
        If rs!qnodtd = "493" Then j = 13
        If rs!qnodtd = "494" Then j = 14
        If rs!qnodtd = "495" Then j = 15
        If rs!qnodtd = "496" Then j = 16
        If rs!qnodtd = "497" Then j = 17
        If rs!qnodtd = "498" Then j = 18
        If rs!qnodtd = "499" Then j = 19
        If rs!qnodtd = "500" Then j = 20
        If rs!qnodtd = "501" Then j = 21
        If rs!qnodtd = "502" Then j = 22
        If rs!qnodtd = "503" Then j = 23
        If rs!qnodtd = "504" Then j = 24
        If rs!qnodtd = "505" Then j = 25
        If rs!qnodtd = "506" Then j = 26
        If rs!qnodtd = "508" Then j = 27
        If rs!qnodtd = "509" Then j = 28
        If rs!qnodtd = "510" Then j = 29
        If rs!qnodtd = "511" Then j = 30
        If rs!qnodtd = "512" Then j = 31
        If rs!qnodtd = "513" Then j = 32
        If rs!qnodtd = "514" Then j = 33
        If rs!qnodtd = "515" Then j = 34
        If rs!qnodtd = "516" Then j = 35
        If rs!qnodtd = "517" Then j = 36
        If rs!qnodtd = "518" Then j = 37
        If rs!qnodtd = "519" Then j = 38
        If rs!qnodtd = "520" Then j = 39
        If rs!qnodtd = "521" Then j = 40
        If rs!qnodtd = "522" Then j = 41
        If rs!qnodtd = "523" Then j = 42
        If rs!qnodtd = "524" Then j = 43
        If rs!qnodtd = "525" Then j = 44
        If rs!qnodtd = "526" Then j = 45
        If rs!qnodtd = "527" Then j = 46
        If rs!qnodtd = "528" Then j = 47
        If rs!qnodtd = "529" Then j = 48
        If rs!qnodtd = "530" Then j = 49
        If rs!qnodtd = "531" Then j = 50
        If rs!qnodtd = "532" Then j = 51
        If rs!qnodtd = "533" Then j = 52
        If rs!qnodtd = "534" Then j = 53
        If rs!qnodtd = "535" Then j = 54
        If rs!qnodtd = "536" Then j = 55
        If rs!qnodtd = "537" Then j = 56
        If rs!qnodtd = "538" Then j = 57
        If rs!qnodtd = "539" Then j = 58
        If rs!qnodtd = "540" Then j = 59
        If rs!qnodtd = "541" Then j = 60
        If rs!qnodtd = "1031" Then j = 61
        If rs!qnodtd = "1043" Then j = 62
        If rs!qnodtd = "1044" Then j = 63
        If rs!qnodtd = "1045" Then j = 64
        If rs!qnodtd = "1046" Then j = 65
        If rs!qnodtd = "1047" Then j = 66
        If rs!qnodtd = "1048" Then j = 67
        If rs!qnodtd = "542" Then j = 68
        If rs!qnodtd = "543" Then j = 69
        If rs!qnodtd = "544" Then j = 70
        If rs!qnodtd = "545" Then j = 71
        If rs!qnodtd = "546" Then j = 72
        If rs!qnodtd = "547" Then j = 73
        If rs!qnodtd = "548" Then j = 74
        If rs!qnodtd = "549" Then j = 75
        If rs!qnodtd = "550" Then j = 76
        If rs!qnodtd = "551" Then j = 77
        If rs!qnodtd = "552" Then j = 78
        If rs!qnodtd = "553" Then j = 79
        If rs!qnodtd = "554" Then j = 80
        If rs!qnodtd = "555" Then j = 81
        If rs!qnodtd = "556" Then j = 82
        If rs!qnodtd = "557" Then j = 83
        If rs!qnodtd = "558" Then j = 84
        If rs!qnodtd = "559" Then j = 85
        If rs!qnodtd = "560" Then j = 86
        If rs!qnodtd = "561" Then j = 87
        If rs!qnodtd = "562" Then j = 88
        If rs!qnodtd = "563" Then j = 89
        If rs!qnodtd = "564" Then j = 90
        If rs!qnodtd = "565" Then j = 91
        If rs!qnodtd = "566" Then j = 92
        If rs!qnodtd = "567" Then j = 93
        If rs!qnodtd = "568" Then j = 94
        If rs!qnodtd = "569" Then j = 95
        If rs!qnodtd = "570" Then j = 96
        If rs!qnodtd = "571" Then j = 97
        If rs!qnodtd = "572" Then j = 98
        If rs!qnodtd = "573" Then j = 99
        If rs!qnodtd = "574" Then j = 100
        If rs!qnodtd = "575" Then j = 101
        If rs!qnodtd = "576" Then j = 102
        If rs!qnodtd = "577" Then j = 103
        If rs!qnodtd = "578" Then j = 104
        If rs!qnodtd = "580" Then j = 105
        If rs!qnodtd = "581" Then j = 106
        If rs!qnodtd = "582" Then j = 107
        If rs!qnodtd = "583" Then j = 108
        If rs!qnodtd = "584" Then j = 109
        If rs!qnodtd = "585" Then j = 110
        If rs!qnodtd = "586" Then j = 111
        If rs!qnodtd = "587" Then j = 112
        If rs!qnodtd = "588" Then j = 113
        If rs!qnodtd = "589" Then j = 114
        If rs!qnodtd = "590" Then j = 115
        If rs!qnodtd = "591" Then j = 116
        If rs!qnodtd = "592" Then j = 117
        If rs!qnodtd = "593" Then j = 118
        If rs!qnodtd = "594" Then j = 119
        If rs!qnodtd = "595" Then j = 120
        If rs!qnodtd = "596" Then j = 121
        If rs!qnodtd = "597" Then j = 122
        If rs!qnodtd = "598" Then j = 123
        If rs!qnodtd = "599" Then j = 124
        If rs!qnodtd = "600" Then j = 125
        If rs!qnodtd = "601" Then j = 126
        If rs!qnodtd = "602" Then j = 127
        If rs!qnodtd = "603" Then j = 128
        If rs!qnodtd = "604" Then j = 129
        If rs!qnodtd = "605" Then j = 130
        If rs!qnodtd = "606" Then j = 131
        If rs!qnodtd = "607" Then j = 132
        If rs!qnodtd = "608" Then j = 133
        If rs!qnodtd = "609" Then j = 134
        If rs!qnodtd = "610" Then j = 135
        If rs!qnodtd = "611" Then j = 136
        If rs!qnodtd = "612" Then j = 137
        If rs!qnodtd = "613" Then j = 138
        If rs!qnodtd = "614" Then j = 139
        If rs!qnodtd = "615" Then j = 140
        If rs!qnodtd = "616" Then j = 141
        If rs!qnodtd = "617" Then j = 142
        If rs!qnodtd = "618" Then j = 143
        If rs!qnodtd = "619" Then j = 144
        If rs!qnodtd = "620" Then j = 145
        If rs!qnodtd = "621" Then j = 146
        If rs!qnodtd = "622" Then j = 147
        If rs!qnodtd = "623" Then j = 148
        If rs!qnodtd = "624" Then j = 149
        If rs!qnodtd = "625" Then j = 150
        If rs!qnodtd = "626" Then j = 151
        If rs!qnodtd = "627" Then j = 152
        If rs!qnodtd = "628" Then j = 153
        If rs!qnodtd = "629" Then j = 154
        If rs!qnodtd = "630" Then j = 155
        If rs!qnodtd = "631" Then j = 156
        If rs!qnodtd = "632" Then j = 157
        If rs!qnodtd = "633" Then j = 158
        If rs!qnodtd = "634" Then j = 159
        If rs!qnodtd = "635" Then j = 160
        If rs!qnodtd = "636" Then j = 161
        If rs!qnodtd = "637" Then j = 162
        If rs!qnodtd = "638" Then j = 163
        If rs!qnodtd = "639" Then j = 164
        If rs!qnodtd = "640" Then j = 165
        If rs!qnodtd = "641" Then j = 166
        If rs!qnodtd = "642" Then j = 167
        If rs!qnodtd = "643" Then j = 168
        If rs!qnodtd = "644" Then j = 169
        If rs!qnodtd = "645" Then j = 170
        If rs!qnodtd = "646" Then j = 171
        If rs!qnodtd = "647" Then j = 172
        If rs!qnodtd = "648" Then j = 173
        If rs!qnodtd = "649" Then j = 174
        If rs!qnodtd = "650" Then j = 175
        If rs!qnodtd = "651" Then j = 176
        If rs!qnodtd = "652" Then j = 177
        If rs!qnodtd = "653" Then j = 178
        If rs!qnodtd = "654" Then j = 179
        If rs!qnodtd = "655" Then j = 180
        If rs!qnodtd = "656" Then j = 181
        If rs!qnodtd = "657" Then j = 182
        If rs!qnodtd = "658" Then j = 183
        If rs!qnodtd = "659" Then j = 184
        If rs!qnodtd = "660" Then j = 185
        If rs!qnodtd = "661" Then j = 186
        If rs!qnodtd = "662" Then j = 187
        If rs!qnodtd = "663" Then j = 188
        If rs!qnodtd = "664" Then j = 189
        If rs!qnodtd = "665" Then j = 190
        If rs!qnodtd = "666" Then j = 191
        If rs!qnodtd = "667" Then j = 192
        If rs!qnodtd = "668" Then j = 193
        If rs!qnodtd = "669" Then j = 194
        If rs!qnodtd = "670" Then j = 195
        If rs!qnodtd = "671" Then j = 196
        If rs!qnodtd = "672" Then j = 197
        If rs!qnodtd = "673" Then j = 198
        If rs!qnodtd = "674" Then j = 199
        If rs!qnodtd = "675" Then j = 200
        If rs!qnodtd = "1049" Then j = 201
        If rs!qnodtd = "1050" Then j = 202
        If rs!qnodtd = "676" Then j = 203
        If rs!qnodtd = "677" Then j = 204
        If rs!qnodtd = "678" Then j = 205
        If rs!qnodtd = "679" Then j = 206
        If rs!qnodtd = "680" Then j = 207
        If rs!qnodtd = "681" Then j = 208
        If rs!qnodtd = "682" Then j = 209
        If rs!qnodtd = "683" Then j = 210
        If rs!qnodtd = "685" Then j = 211
        If rs!qnodtd = "686" Then j = 212
        If rs!qnodtd = "687" Then j = 213
        If rs!qnodtd = "688" Then j = 214
        If rs!qnodtd = "689" Then j = 215
        If rs!qnodtd = "690" Then j = 216
        If rs!qnodtd = "691" Then j = 217
        If rs!qnodtd = "692" Then j = 218
        If rs!qnodtd = "693" Then j = 219
        If rs!qnodtd = "694" Then j = 220
        If rs!qnodtd = "695" Then j = 221
        If rs!qnodtd = "696" Then j = 222
        If rs!qnodtd = "697" Then j = 223
        If rs!qnodtd = "698" Then j = 224
        If rs!qnodtd = "699" Then j = 225
        If rs!qnodtd = "700" Then j = 226
        If rs!qnodtd = "701" Then j = 227
        If rs!qnodtd = "702" Then j = 228
        If rs!qnodtd = "703" Then j = 229
        If rs!qnodtd = "704" Then j = 230
        If rs!qnodtd = "705" Then j = 231
        If rs!qnodtd = "706" Then j = 232
        If rs!qnodtd = "707" Then j = 233
        If rs!qnodtd = "708" Then j = 234
        If rs!qnodtd = "709" Then j = 235
        If rs!qnodtd = "710" Then j = 236
        If rs!qnodtd = "711" Then j = 237
        If rs!qnodtd = "712" Then j = 238
        If rs!qnodtd = "713" Then j = 239
        If rs!qnodtd = "714" Then j = 240
        If rs!qnodtd = "715" Then j = 241
        If rs!qnodtd = "716" Then j = 242
        If rs!qnodtd = "717" Then j = 243
        If rs!qnodtd = "718" Then j = 244
        If rs!qnodtd = "719" Then j = 245
        If rs!qnodtd = "720" Then j = 246
        If rs!qnodtd = "721" Then j = 247
        If rs!qnodtd = "722" Then j = 248
        If rs!qnodtd = "723" Then j = 249
        If rs!qnodtd = "724" Then j = 250
        If rs!qnodtd = "725" Then j = 251
        If rs!qnodtd = "726" Then j = 252
        If rs!qnodtd = "727" Then j = 253
        If rs!qnodtd = "728" Then j = 254
        If rs!qnodtd = "729" Then j = 255
        If rs!qnodtd = "730" Then j = 256
        If rs!qnodtd = "731" Then j = 257
        If rs!qnodtd = "732" Then j = 258
        If rs!qnodtd = "733" Then j = 259
        If rs!qnodtd = "734" Then j = 260
        If rs!qnodtd = "735" Then j = 261
        If rs!qnodtd = "736" Then j = 262
        If rs!qnodtd = "738" Then j = 263
        If rs!qnodtd = "739" Then j = 264
        If rs!qnodtd = "740" Then j = 265
        If rs!qnodtd = "741" Then j = 266
        If rs!qnodtd = "742" Then j = 267
        If rs!qnodtd = "743" Then j = 268
        If rs!qnodtd = "744" Then j = 269
        If rs!qnodtd = "745" Then j = 270
        If rs!qnodtd = "746" Then j = 271
        If rs!qnodtd = "747" Then j = 272
        If rs!qnodtd = "748" Then j = 273
        If rs!qnodtd = "749" Then j = 274
        If rs!qnodtd = "750" Then j = 275
        If rs!qnodtd = "751" Then j = 276
        If rs!qnodtd = "1052" Then j = 277
        If rs!qnodtd = "752" Then j = 278
        If rs!qnodtd = "753" Then j = 279
        If rs!qnodtd = "754" Then j = 280
        If rs!qnodtd = "755" Then j = 281
        If rs!qnodtd = "756" Then j = 282
        If rs!qnodtd = "757" Then j = 283
        If rs!qnodtd = "758" Then j = 284
        If rs!qnodtd = "759" Then j = 285
        If rs!qnodtd = "760" Then j = 286
        If rs!qnodtd = "761" Then j = 287
        If rs!qnodtd = "762" Then j = 288
        If rs!qnodtd = "763" Then j = 289
        If rs!qnodtd = "764" Then j = 290
        If rs!qnodtd = "765" Then j = 291
        If rs!qnodtd = "766" Then j = 292
        If rs!qnodtd = "767" Then j = 293
        If rs!qnodtd = "768" Then j = 294
        If rs!qnodtd = "769" Then j = 295
        If rs!qnodtd = "770" Then j = 296
        If rs!qnodtd = "771" Then j = 297
        If rs!qnodtd = "772" Then j = 298
        If rs!qnodtd = "773" Then j = 299
        If rs!qnodtd = "774" Then j = 300
        If rs!qnodtd = "775" Then j = 301
        If rs!qnodtd = "776" Then j = 302
        If rs!qnodtd = "777" Then j = 303
        If rs!qnodtd = "778" Then j = 304
        If rs!qnodtd = "779" Then j = 305
        If rs!qnodtd = "780" Then j = 306
        If rs!qnodtd = "781" Then j = 307
        If rs!qnodtd = "782" Then j = 308
        If rs!qnodtd = "783" Then j = 309
        If rs!qnodtd = "784" Then j = 310
        If rs!qnodtd = "785" Then j = 311
        If rs!qnodtd = "786" Then j = 312
        If rs!qnodtd = "787" Then j = 313
        If rs!qnodtd = "788" Then j = 314
        If rs!qnodtd = "789" Then j = 315
        If rs!qnodtd = "790" Then j = 316
        If rs!qnodtd = "791" Then j = 317
        If rs!qnodtd = "792" Then j = 318
        If rs!qnodtd = "793" Then j = 319
        If rs!qnodtd = "794" Then j = 320
        If rs!qnodtd = "795" Then j = 321
        If rs!qnodtd = "796" Then j = 322
        If rs!qnodtd = "797" Then j = 323
        If rs!qnodtd = "798" Then j = 324
        If rs!qnodtd = "799" Then j = 325
        If rs!qnodtd = "800" Then j = 326
        If rs!qnodtd = "801" Then j = 327
        If rs!qnodtd = "802" Then j = 328
        If rs!qnodtd = "803" Then j = 329
        If rs!qnodtd = "804" Then j = 330
        If rs!qnodtd = "805" Then j = 331
        If rs!qnodtd = "806" Then j = 332
        If rs!qnodtd = "807" Then j = 333
        If rs!qnodtd = "808" Then j = 334
        If rs!qnodtd = "809" Then j = 335
        If rs!qnodtd = "810" Then j = 336
        If rs!qnodtd = "811" Then j = 337
        If rs!qnodtd = "812" Then j = 338
        If rs!qnodtd = "813" Then j = 339
        If rs!qnodtd = "814" Then j = 340
        If rs!qnodtd = "815" Then j = 341
        If rs!qnodtd = "816" Then j = 342
        If rs!qnodtd = "817" Then j = 343
        If rs!qnodtd = "818" Then j = 344
        If rs!qnodtd = "819" Then j = 345
        If rs!qnodtd = "820" Then j = 346
        If rs!qnodtd = "821" Then j = 347
        If rs!qnodtd = "822" Then j = 348
        If rs!qnodtd = "823" Then j = 349
        If rs!qnodtd = "824" Then j = 350
        If rs!qnodtd = "825" Then j = 351
        If rs!qnodtd = "1054" Then j = 352
        If rs!qnodtd = "826" Then j = 353
        If rs!qnodtd = "827" Then j = 354
        If rs!qnodtd = "1081" Then j = 355
        If rs!qnodtd = "1082" Then j = 356
        If rs!qnodtd = "828" Then j = 357
        If rs!qnodtd = "829" Then j = 358
        If rs!qnodtd = "830" Then j = 359
        If rs!qnodtd = "831" Then j = 360
        If rs!qnodtd = "832" Then j = 361
        If rs!qnodtd = "833" Then j = 362
        If rs!qnodtd = "834" Then j = 363
        If rs!qnodtd = "835" Then j = 364
        If rs!qnodtd = "836" Then j = 365
        If rs!qnodtd = "1083" Then j = 366
        If rs!qnodtd = "837" Then j = 367
        If rs!qnodtd = "838" Then j = 368
        If rs!qnodtd = "839" Then j = 369
        If rs!qnodtd = "840" Then j = 370
        If rs!qnodtd = "841" Then j = 371
        If rs!qnodtd = "842" Then j = 372
        If rs!qnodtd = "843" Then j = 373
        If rs!qnodtd = "844" Then j = 374
        If rs!qnodtd = "845" Then j = 375
        If rs!qnodtd = "846" Then j = 376
        If rs!qnodtd = "847" Then j = 377
        If rs!qnodtd = "848" Then j = 378
        If rs!qnodtd = "849" Then j = 379
        If rs!qnodtd = "850" Then j = 380
        If rs!qnodtd = "851" Then j = 381
        If rs!qnodtd = "852" Then j = 382
        If rs!qnodtd = "853" Then j = 383
        If rs!qnodtd = "854" Then j = 384
        If rs!qnodtd = "855" Then j = 385
        If rs!qnodtd = "856" Then j = 386
        If rs!qnodtd = "857" Then j = 387
        If rs!qnodtd = "858" Then j = 388
        If rs!qnodtd = "859" Then j = 389
        If rs!qnodtd = "860" Then j = 390
        If rs!qnodtd = "861" Then j = 391
        If rs!qnodtd = "862" Then j = 392
        If rs!qnodtd = "863" Then j = 393
        If rs!qnodtd = "864" Then j = 394
        If rs!qnodtd = "865" Then j = 395
        If rs!qnodtd = "866" Then j = 396
        If rs!qnodtd = "867" Then j = 397
        If rs!qnodtd = "868" Then j = 398
        If rs!qnodtd = "869" Then j = 399
        If rs!qnodtd = "870" Then j = 400
        If rs!qnodtd = "872" Then j = 401
        If rs!qnodtd = "873" Then j = 402
        If rs!qnodtd = "874" Then j = 403
        If rs!qnodtd = "875" Then j = 404
        If rs!qnodtd = "876" Then j = 405
        If rs!qnodtd = "877" Then j = 406
        If rs!qnodtd = "878" Then j = 407
        If rs!qnodtd = "879" Then j = 408
        If rs!qnodtd = "880" Then j = 409
        If rs!qnodtd = "881" Then j = 410
        If rs!qnodtd = "882" Then j = 411
        If rs!qnodtd = "883" Then j = 412
        If rs!qnodtd = "884" Then j = 413
        If rs!qnodtd = "885" Then j = 414
        If rs!qnodtd = "886" Then j = 415
        If rs!qnodtd = "887" Then j = 416
        If rs!qnodtd = "888" Then j = 417
        If rs!qnodtd = "889" Then j = 418
        If rs!qnodtd = "890" Then j = 419
        If rs!qnodtd = "891" Then j = 420
        If rs!qnodtd = "892" Then j = 421
        If rs!qnodtd = "893" Then j = 422
        If rs!qnodtd = "894" Then j = 423
        If rs!qnodtd = "895" Then j = 424
        If rs!qnodtd = "896" Then j = 425
        If rs!qnodtd = "897" Then j = 426
        If rs!qnodtd = "898" Then j = 427
        If rs!qnodtd = "899" Then j = 428
        If rs!qnodtd = "900" Then j = 429
        If rs!qnodtd = "901" Then j = 430
        If rs!qnodtd = "902" Then j = 431
        If rs!qnodtd = "903" Then j = 432
        If rs!qnodtd = "904" Then j = 433
        If rs!qnodtd = "905" Then j = 434
        If rs!qnodtd = "906" Then j = 435
        If rs!qnodtd = "907" Then j = 436
        If rs!qnodtd = "908" Then j = 437
        If rs!qnodtd = "909" Then j = 438
        If rs!qnodtd = "910" Then j = 439
        If rs!qnodtd = "911" Then j = 440
        If rs!qnodtd = "912" Then j = 441
        If rs!qnodtd = "913" Then j = 442
        If rs!qnodtd = "914" Then j = 443
        If rs!qnodtd = "915" Then j = 444
        If rs!qnodtd = "916" Then j = 445
        If rs!qnodtd = "917" Then j = 446
        If rs!qnodtd = "918" Then j = 447
        If rs!qnodtd = "919" Then j = 448
        If rs!qnodtd = "920" Then j = 449
        If rs!qnodtd = "921" Then j = 450
        If rs!qnodtd = "922" Then j = 451
        If rs!qnodtd = "923" Then j = 452
        If rs!qnodtd = "924" Then j = 453
        If rs!qnodtd = "925" Then j = 454
        If rs!qnodtd = "926" Then j = 455
        If rs!qnodtd = "927" Then j = 456
        If rs!qnodtd = "928" Then j = 457
        If rs!qnodtd = "929" Then j = 458
        If rs!qnodtd = "930" Then j = 459
        If rs!qnodtd = "931" Then j = 460
        If rs!qnodtd = "932" Then j = 461
        If rs!qnodtd = "933" Then j = 462
        If rs!qnodtd = "934" Then j = 463
        If rs!qnodtd = "935" Then j = 464
        If rs!qnodtd = "936" Then j = 465
        If rs!qnodtd = "937" Then j = 466
        If rs!qnodtd = "938" Then j = 467
        If rs!qnodtd = "939" Then j = 468
        If rs!qnodtd = "940" Then j = 469
        If rs!qnodtd = "941" Then j = 470
        If rs!qnodtd = "942" Then j = 471
        If rs!qnodtd = "943" Then j = 472
        If rs!qnodtd = "944" Then j = 473
        If rs!qnodtd = "945" Then j = 474
        If rs!qnodtd = "946" Then j = 475
        If rs!qnodtd = "947" Then j = 476
        If rs!qnodtd = "948" Then j = 477
        If rs!qnodtd = "949" Then j = 478
        If rs!qnodtd = "950" Then j = 479
        If rs!qnodtd = "951" Then j = 480
        If rs!qnodtd = "952" Then j = 481
        If rs!qnodtd = "953" Then j = 482
        If rs!qnodtd = "1084" Then j = 483
        If rs!qnodtd = "954" Then j = 484
        If rs!qnodtd = "955" Then j = 485
        If rs!qnodtd = "956" Then j = 486
        If rs!qnodtd = "957" Then j = 487
        If rs!qnodtd = "958" Then j = 488
        If rs!qnodtd = "959" Then j = 489
        If rs!qnodtd = "960" Then j = 490
        If rs!qnodtd = "961" Then j = 491
        If rs!qnodtd = "962" Then j = 492
        If rs!qnodtd = "963" Then j = 493
        If rs!qnodtd = "964" Then j = 494
        If rs!qnodtd = "965" Then j = 495
        If rs!qnodtd = "966" Then j = 496
        If rs!qnodtd = "967" Then j = 497
        If rs!qnodtd = "968" Then j = 498
        If rs!qnodtd = "969" Then j = 499
        If rs!qnodtd = "970" Then j = 500
        If rs!qnodtd = "971" Then j = 501
        If rs!qnodtd = "972" Then j = 502
        If rs!qnodtd = "973" Then j = 503
        If rs!qnodtd = "974" Then j = 504
        If rs!qnodtd = "975" Then j = 505
        If rs!qnodtd = "976" Then j = 506
        If rs!qnodtd = "977" Then j = 507
        If rs!qnodtd = "978" Then j = 508

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
        rs.MoveNext

        If rs.EOF = True Then Exit Sub

        If rs!qnodtd = "541" Then
            rs.MoveNext
        ElseIf rs!qnodtd = "751" Then
            rs.MoveNext
        End If
    Wend
End Sub

