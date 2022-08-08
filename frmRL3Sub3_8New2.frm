VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRL3Sub3_8New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.8 Pemeriksaan Laboratorium"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6135
   Icon            =   "frmRL3Sub3_8New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6135
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   6135
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   1320
         Width           =   1905
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   1320
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtptahun 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   600
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
         CustomFormat    =   "yyyy"
         Format          =   138477571
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3000
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   184
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
      Left            =   5520
      TabIndex        =   5
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL3Sub3_8New2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmRL3Sub3_8New2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Special Buat Excel
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel.Range
Dim oResizeRange As Excel.Range
Dim i, j, k, l, xx As Integer
Dim w, X, Y, z As String
'Special Buat Excel

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)

    dtptahun.value = Now
    dtptahun.CustomFormat = "yyyyy"
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo error

    ProgressBar1.value = ProgressBar1.Min
    lblPersen.Caption = "0 %"
    Screen.MousePointer = vbHourglass

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.path & "\RL 3.8_laboratorium.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "select NomorUrut,JenisKegiatan,COUNT(NoPendaftaran) as Jumlah " & _
                 "from RL3_08New2 where YEAR(TglPelayanan)='" & dtptahun.Year & "' And NomorUrut IS NOT Null " & _
                 "Group by NomorUrut,JenisKegiatan order by NomorUrut"
        
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    
'    For i = 1 To 184
    
    For i = 1 To rsb.RecordCount
    
        Select Case rsb!NomorUrut
                '1 Hematologi
                '1.1 Sitologi Sel Darah-------------------------------------------------------------------------------------------
            Case 1
                j = 4
                w = "and KdPelayananRS = '069006'"
            Case 2
                j = 5
                w = "and KdPelayananRS = '068013'"
            Case 3
                j = 6
                w = "and KdPelayananRS = ''"
            Case 4
                j = 7
                w = "and KdPelayananRS = '067022'"
            Case 5
                j = 8
                w = "and KdPelayananRS = ''"
            Case 6
                j = 9
                w = "and KdPelayananRS = ''"
            Case 7
                j = 10
                w = "and KdPelayananRS = '068042'"
            Case 8
                j = 11
                w = "and KdPelayananRS = '068017'"
                '1.2 Sitokimia Darah
            Case 9
                j = 13
                w = "and KdPelayananRS = ''"
            Case 10
                j = 14
                w = "and KdPelayananRS = ''"
            Case 11
                j = 15
                w = "and KdPelayananRS = ''"
            Case 12
                j = 16
                w = "and KdPelayananRS = ''"
            Case 13
                j = 17
                w = "and KdPelayananRS = ''"
            Case 14
                j = 18
                w = "and KdPelayananRS = ''"
                '1.3 Analisa Hb
            Case 15
                j = 20
                w = "and KdPelayananRS = ''"
            Case 16
                j = 21
                w = "and KdPelayananRS = ''"
            Case 17
                j = 22
                w = "and KdPelayananRS = ''"
                '1.4 Perbankan Darah
            Case 18
                j = 24
                w = "and KdPelayananRS = ''"
            Case 19
                j = 25
                w = "and KdPelayananRS = '069024'"
            Case 20
                j = 26
                w = "and KdPelayananRS = ''"
            Case 21
                j = 27
                w = "and KdPelayananRS = ''"
                '1.5 Hemostasis
            Case 22
                j = 29
                w = "and KdPelayananRS = ''"
            Case 23
                j = 30
                w = "and KdPelayananRS = ''"
            Case 24
                j = 31
                w = "and KdPelayananRS = ''"
            Case 25
                j = 32
                w = "and KdPelayananRS = ''"
            Case 26
                j = 33
                w = "and KdPelayananRS = ''"
            Case 27
                j = 34
                w = "and KdPelayananRS = ''"
            Case 28
                j = 35
                w = "and KdPelayananRS = ''"
            Case 29
                j = 36
                w = "and KdPelayananRS = ''"
            Case 30
                j = 37
                w = "and KdPelayananRS = ''"
            Case 31
                j = 38
                w = "and KdPelayananRS = ''"
            Case 32
                j = 39
                w = "and KdPelayananRS = '064024'"
            Case 33
                j = 40
                w = "and KdPelayananRS = ''"
            Case 34
                j = 41
                w = "and KdPelayananRS = '076026'"
            Case 35
                j = 42
                w = "and KdPelayananRS = ''"
            Case 36
                j = 43
                w = "and KdPelayananRS = ''"
            Case 37
                j = 44
                w = "and KdPelayananRS = ''"
            Case 38
                j = 45
                w = "and KdPelayananRS = ''"
            Case 39
                j = 46
                w = "and KdPelayananRS = '068044'"
            Case 40
                j = 47
                w = "and KdPelayananRS = ''"
            Case 41
                j = 48
                w = "and KdPelayananRS = ''"
            Case 42
                j = 49
                w = "and KdPelayananRS = ''"
            Case 43
                j = 50
                w = "and KdPelayananRS = ''"
                '1.6 Pemeriksaan Lain
            Case 44
                j = 52
                w = "and KdPelayananRS = ''"
            Case 45
                j = 53
                w = "and KdPelayananRS = ''"
            Case 46
                j = 54
                w = "and KdPelayananRS = ''"
            Case 47
                j = 55
                w = "and KdPelayananRS = ''"
            Case 48
                j = 56
                w = "and KdPelayananRS = ''"
            Case 49
                j = 57
                w = "and KdPelayananRS = '068022'"
            Case 50
                j = 58
                w = "and KdPelayananRS = ''"
            Case 51
                j = 59
                w = "and KdPelayananRS = ''"
                '2 Kimia Klinik
                '2.1 Protein dan NPN----------------------------------------------------------------------------------------------
            Case 52
                j = 62
                w = "and KdPelayananRS = '065001'"
            Case 53
                j = 63
                w = "and KdPelayananRS = ''"
            Case 54
                j = 64
                w = "and KdPelayananRS = '066003'"
            Case 55
                j = 65
                w = "and KdPelayananRS = '068002'"
            Case 56
                j = 66
                w = "and KdPelayananRS = ''"
            Case 57
                j = 67
                w = "and KdPelayananRS = '076008'"
            Case 58
                j = 68
                w = "and KdPelayananRS = ''"
            Case 59
                j = 69
                w = "and KdPelayananRS = '064017'"
            Case 60
                j = 70
                w = "and KdPelayananRS = ''"
            Case 61
                j = 71
                w = "and KdPelayananRS = ''"
            Case 62
                j = 72
                w = "and KdPelayananRS = ''"
            Case 63
                j = 73
                w = "and KdPelayananRS = ''"
            Case 64
                j = 74
                w = "and KdPelayananRS = '068038'"
            Case 65
                j = 75
                w = "and KdPelayananRS = ''"
            Case 66
                j = 76
                w = "and KdPelayananRS = '067029'"
            Case 67
                j = 77
                w = "and KdPelayananRS = ''"
            Case 68
                j = 78
                w = "and KdPelayananRS = ''"
            Case 69
                j = 79
                w = "and KdPelayananRS = '065011'"
            Case 70
                j = 80
                w = "and KdPelayananRS = '065018'"
            Case 71
                j = 81
                w = "and KdPelayananRS = '068051'"
            Case 72
                j = 82
                w = "and KdPelayananRS = '068052'"
                '2.2 Karbohidrat
            Case 73
                j = 88
                w = "and KdPelayananRS = '067002'"
            Case 74
                j = 85
                w = "and KdPelayananRS = ''"
            Case 75
                j = 86
                w = "and KdPelayananRS = ''"
            Case 76
                j = 87
                w = "and KdPelayananRS = ''"
            Case 77
                j = 88
                w = "and KdPelayananRS = ''"
                '2.3 Lipid, Lipoprotein, Apoprotein
            Case 78
                j = 90
                w = "and KdPelayananRS = ''"
            Case 79
                j = 91
                w = "and KdPelayananRS = ''"
            Case 80
                j = 92
                w = "and KdPelayananRS = ''"
            Case 81
                j = 93
                w = "and KdPelayananRS = ''"
            Case 82
                j = 94
                w = "and KdPelayananRS = '066012'"
            Case 83
                j = 95
                w = "and KdPelayananRS = '066026'"
            Case 84
                j = 96
                w = "and KdPelayananRS = ''"
            Case 85
                j = 97
                w = "and KdPelayananRS = ''"
            Case 86
                j = 98
                w = "and KdPelayananRS = '059015' or KdPelayananRS='008024' or KdPelayananRS='066028' or KdPelayananRS='164016'"
                '2.4 Enzim
            Case 87
                j = 100
                w = "and KdPelayananRS= '066001'"
            Case 88
                j = 101
                w = "and KdPelayananRS = ''"
            Case 89
                j = 102
                w = "and KdPelayananRS = ''"
            Case 90
                j = 103
                w = "and KdPelayananRS = ''"
            Case 91
                j = 104
                w = "and KdPelayananRS = ''"
            Case 92
                j = 105
                w = "and KdPelayananRS= '063003'"
            Case 93
                j = 106
                w = "and KdPelayananRS = ''"
            Case 94
                j = 107
                w = "and KdPelayananRS = '060010' or KdPelayananRS = '008010'"
            Case 95
                j = 108
                w = "and KdPelayananRS = ''"
            Case 96
                j = 109
                w = "and KdPelayananRS = ''"
            Case 97
                j = 110
                w = "and KdPelayananRS = ''"
            Case 98
                j = 111
                w = "and KdPelayananRS = ''"
            Case 99
                j = 112
                w = "and KdPelayananRS = ''"
            Case 100
                j = 113
                w = "and KdPelayananRS = ''"
            Case 101
                j = 114
                w = "and KdPelayananRS = ''"
            Case 102
                j = 115
                w = "and KdPelayananRS = '064018'"
                '2.5 Mikronutrient dan Monitoring Kadar Terapi Obat
            Case 103
                j = 117
                w = "and KdPelayananRS = ''"
            Case 104
                j = 118
                w = "and KdPelayananRS = ''"
            Case 105
                j = 119
                w = "and KdPelayananRS = ''"
            Case 106
                j = 120
                w = "and KdPelayananRS = ''"
            Case 107
                j = 121
                w = "and KdPelayananRS = ''"
            Case 108
                j = 122
                w = "and KdPelayananRS = ''"
            Case 109
                j = 123
                w = "and KdPelayananRS = ''"
            Case 110
                j = 124
                w = "and KdPelayananRS = ''"
            Case 111
                j = 125
                w = "and KdPelayananRS = ''"
            Case 112
                j = 126
                w = "and KdPelayananRS = ''"
            Case 113
                j = 127
                w = "and KdPelayananRS = ''"
            Case 114
                j = 128
                w = "and KdPelayananRS = ''"
            Case 115
                j = 129
                w = "and KdPelayananRS = ''"
            Case 116
                j = 130
                w = "and KdPelayananRS = '066016'"
            Case 117
                j = 131
                w = "and KdPelayananRS = ''"
            Case 118
                j = 132
                w = "and KdPelayananRS = ''"
            Case 119
                j = 133
                w = "and KdPelayananRS = ''"
            Case 120
                j = 134
                w = "and KdPelayananRS = ''"
            Case 121
                j = 135
                w = "and KdPelayananRS = ''"
            Case 122
                j = 136
                w = "and KdPelayananRS = ''"
            Case 123
                j = 137
                w = "and KdPelayananRS = ''"
                '2.6 Elektrolit
            Case 124
                j = 139
                w = "and KdPelayananRS = ''"
            Case 125
                j = 140
                w = "and KdPelayananRS = '076015' or KdPelayananRS = '066009'"
            Case 126
                j = 141
                w = "and KdPelayananRS = '066010'"
            Case 127
                j = 142
                w = "and KdPelayananRS = '066011'"
            Case 128
                j = 143
                w = "and KdPelayananRS = '066017'"
            Case 129
                j = 144
                w = "and KdPelayananRS = '066016'"
                '2.7 Fungsi Organ
            Case 130
                j = 146
                w = "and KdPelayananRS = ''"
            Case 131
                j = 147
                w = "and KdPelayananRS = ''"
            Case 132
                j = 148
                w = "and KdPelayananRS = ''"
            Case 133
                j = 149
                w = "and KdPelayananRS = '065008'"
            Case 134
                j = 150
                w = "and KdPelayananRS = ''"
            Case 135
                j = 151
                w = "and KdPelayananRS = ''"
            Case 136
                j = 152
                w = "and KdPelayananRS = '066010'"
            Case 137
                j = 153
                w = "and KdPelayananRS = ''"
            Case 138
                j = 154
                w = "and KdPelayananRS = ''"
                '2.8 Hormon dan Fungsi Endokrin
            Case 139
                j = 156
                w = "and KdPelayananRS = ''"
            Case 140
                j = 157
                w = "and KdPelayananRS = ''"
            Case 141
                j = 158
                w = "and KdPelayananRS = ''"
            Case 142
                j = 159
                w = "and KdPelayananRS = ''"
            Case 143
                j = 160
                w = "and KdPelayananRS = ''"
            Case 144
                j = 161
                w = "and KdPelayananRS = '062003'"
            Case 145
                j = 162
                w = "and KdPelayananRS = '062002'"
            Case 146
                j = 163
                w = "and KdPelayananRS = ''"
            Case 147
                j = 164
                w = "and KdPelayananRS = ''"
            Case 148
                j = 165
                w = "and KdPelayananRS = ''"
            Case 149
                j = 166
                w = "and KdPelayananRS = ''"
            Case 150
                j = 167
                w = "and KdPelayananRS = ''"
            Case 151
                j = 168
                w = "and KdPelayananRS = ''"
            Case 152
                j = 169
                w = "and KdPelayananRS = ''"
            Case 153
                j = 170
                w = "and KdPelayananRS = ''"
            Case 154
                j = 171
                w = "and KdPelayananRS = ''"
            Case 155
                j = 172
                w = "and KdPelayananRS = ''"
            Case 156
                j = 173
                w = "and KdPelayananRS = '068020' or KdPelayananRS = '068020'"
            Case 157
                j = 174
                w = "and KdPelayananRS = ''"
            Case 158
                j = 175
                w = "and KdPelayananRS = ''"
            Case 159
                j = 176
                w = "and KdPelayananRS = ''"
            Case 160
                j = 177
                w = "and KdPelayananRS = ''"
            Case 161
                j = 178
                w = "and KdPelayananRS = '062008'"
            Case 162
                j = 179
                w = "and KdPelayananRS = '062009'"
            Case 163
                j = 180
                w = "and KdPelayananRS = ''"
            Case 164
                j = 181
                w = "and KdPelayananRS = '062010'"
            Case 165
                j = 182
                w = "and KdPelayananRS = ''"
            Case 166
                j = 183
                w = "and KdPelayananRS = ''"
            Case 167
                j = 184
                w = "and KdPelayananRS = ''"
            Case 168
                j = 185
                w = "and KdPelayananRS = ''"
            Case 169
                j = 186
                w = "and KdPelayananRS = ''"
            Case 170
                j = 187
                w = "and KdPelayananRS = ''"
                '2.9 Pemeriksaan Lain
            Case 171
                j = 189
                w = "and KdPelayananRS = ''"
            Case 172
                j = 190
                w = "and KdPelayananRS = ''"
            Case 173
                j = 191
                w = "and KdPelayananRS = ''"
            Case 174
                j = 192
                w = "and KdPelayananRS = ''"
            Case 175
                j = 193
                w = "and KdPelayananRS = ''"
            Case 176
                j = 194
                w = "and KdPelayananRS = ''"
            Case 177
                j = 195
                w = "and KdPelayananRS = ''"
            Case 178
                j = 196
                w = "and KdPelayananRS = ''"
            Case 179
                j = 197
                w = "and KdPelayananRS = ''"
            Case 180
                j = 198
                w = "and KdPelayananRS = ''"
            Case 181
                j = 199
                w = "and KdPelayananRS = ''"
            Case 182
                j = 200
                w = "and KdPelayananRS = '067037'"
            Case 183
                j = 201
                w = "and KdPelayananRS = ''"
            Case 184
                j = 202
                w = "and KdPelayananRS = ''"
        End Select

'        strSQL = "SELECT sum (JmlPasien) as jmlpasien From RL3_08New " & _
'        " where Year(Tglpelayanan) = '" & dtptahun.Year & "' " & w & ""
        
'        Set rsb = Nothing
'        rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        With oSheet
            If rsb("Jumlah").value <> "" Then
                .Cells(j, 8) = rsb("Jumlah").value
            Else
                .Cells(j, 8) = "0"
            End If
        End With

        ProgressBar1.value = Int(ProgressBar1.value) + 1
        lblPersen.Caption = Int(ProgressBar1.value / rsb.RecordCount * 100) & " %"
        
        rsb.MoveNext
    Next i

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    For xx = 2 To 202
        With oSheet
            .Cells(xx, 3) = rsb("KdRS").value
            .Cells(xx, 2) = rsb("KotaKodyaKab").value
            .Cells(xx, 4) = rsb("NamaRS").value
            .Cells(xx, 5) = Format(dtptahun.value, "YYYY")
        End With
    Next xx

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    Call msubPesanError
    Screen.MousePointer = vbDefault
End Sub

