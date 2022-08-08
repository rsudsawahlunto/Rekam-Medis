VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRL51 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL5,1 Data Peralatan Medik Rumah Sakit"
   ClientHeight    =   3540
   ClientLeft      =   7725
   ClientTop       =   3840
   ClientWidth     =   6300
   Icon            =   "frmRL5,1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6300
   Begin VB.OptionButton Option3 
      Caption         =   "Hal. 3"
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Hal. 2"
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Hal. 1"
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   1680
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
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
         CustomFormat    =   " yyyy"
         Format          =   127139843
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
   End
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
      Left            =   -120
      TabIndex        =   0
      Top             =   2760
      Width           =   6405
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
         Left            =   3480
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
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
      Caption         =   "Halaman"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   6255
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   3360
      Picture         =   "frmRL5,1.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2955
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRL5,1.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL5,1.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmRL51"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project/reference/microsoft excel 12.0 object library
'Selalu gunakan format file excel 2003  .xls sebagai standar agar pengguna excel 2003 atau diatasnya dpt menggunakan report laporannya
'Catatan: Format excel 2000 tidak dpt mengoperasikan beberapa fungsi yg ada pada excell 2003 atau diatasnya

Option Explicit

'Special Buat Excel
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel.Range
Dim oResizeRange As Excel.Range
Dim j As String
'Special Buat Excel

Private Sub cmdCetak_Click()
    On Error GoTo hell

    If Option1.value = True Then

        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 5,1 Hal.1.xls")
        Set oSheet = oWB.ActiveSheet

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("g6", "g7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("q6", "q7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "select * from RL5a where kdbarang <= 000000068 "
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 14
            Call setcell
        End If

    ElseIf Option2.value = True Then

        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 5,1 Hal.2.xls")
        Set oSheet = oWB.ActiveSheet

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("g6", "g7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("q6", "q7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "select * from RL5a where kdbarang between '000000069' and '000000133' "
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 14
            Call setcell
        End If

    ElseIf Option3.value = True Then

        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL 5,1 Hal.3.xls")
        Set oSheet = oWB.ActiveSheet

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("g6", "g7")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("q6", "q7")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing
        strSQL = "select * from RL5a where kdbarang between '000000134' and '000000181' "
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 14
            Call setcell
        End If
    End If
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
End Sub

Private Sub setcell()
    While Not rs.EOF
        With oSheet
            .Cells(j, 8) = Trim(IIf(IsNull(rs![<5].value), 0, (rs![<5].value)))
            .Cells(j, 9) = Trim(IIf(IsNull(rs![5-10].value), 0, (rs![5-10].value)))
            .Cells(j, 10) = Trim(IIf(IsNull(rs![>10].value), 0, (rs![>10].value)))
            .Cells(j, 11) = Trim(IIf(IsNull(rs!KapasitasRata.value), 0, (rs!KapasitasRata.value)))
            .Cells(j, 12) = Trim(IIf(IsNull(rs!Baik.value), 0, (rs!Baik.value)))
            .Cells(j, 13) = Trim(IIf(IsNull(rs!RusakRingan.value), 0, (rs!RusakRingan.value)))
            .Cells(j, 14) = Trim(IIf(IsNull(rs!RusakBerat.value), 0, (rs!RusakBerat.value)))
            .Cells(j, 15) = Trim(IIf(IsNull(rs!IjinAda.value), 0, (rs!IjinAda.value)))
            .Cells(j, 16) = Trim(IIf(IsNull(rs!IjinTidakAda.value), 0, (rs!IjinTidakAda.value)))
            .Cells(j, 17) = Trim(IIf(IsNull(rs!SertifikatAda.value), 0, (rs!SertifikatAda.value)))
            .Cells(j, 18) = Trim(IIf(IsNull(rs!SertifikatTidakAda.value), 0, (rs!SertifikatTidakAda.value)))
        End With
        j = j + 1
        rs.MoveNext
    Wend
End Sub
