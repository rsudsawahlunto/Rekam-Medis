VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRL3Sub3_2New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.2 Kunjungan Rawat Darurat"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmRL3Sub3_2New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6135
   Begin VB.Frame Frame2 
      Caption         =   "Filter Periode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   6135
      Begin VB.OptionButton optPeriode 
         Caption         =   "Tahun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optPeriode 
         Caption         =   "Tanggal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   6135
      Begin VB.CommandButton cmdTutup 
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
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
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
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   1425
      End
      Begin MSComCtl2.DTPicker dtptahun 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   240
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
         CustomFormat    =   "yyyy"
         Format          =   138084355
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtpTgl 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
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
         Format          =   116916227
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtpTgl 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
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
         Format          =   116916227
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin VB.Label Label1 
         Caption         =   "s/d"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   240
         Width           =   255
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
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL3Sub3_2New2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
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
End
Attribute VB_Name = "frmRL3Sub3_2New2"
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
Dim Cell22 As String
Dim Cell23 As String
Dim Cell24 As String
Dim Cell25 As String
Dim Cell26 As String
Dim Cell27 As String
Dim Cell28 As String

'Special Buat Excel
Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    
    dtpTgl(0).Visible = True
    dtpTgl(1).Visible = True
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
    Set oWB = oXL.Workbooks.Open(App.path & "\RL 3.2_Rawat darurat.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    For xx = 2 To 6
        With oSheet
            .Cells(xx, 1) = rsb("KodeExternal").value
            .Cells(xx, 2) = rsb("KotaKodyaKab").value
            .Cells(xx, 3) = rsb("KdRS").value
            .Cells(xx, 4) = rsb("NamaRS").value
            .Cells(xx, 5) = Format(dtptahun.value, "YYYY")
        End With
    Next xx

    Set rsx = Nothing
    strSQL = "select JenisPelayanan, SUM(Rujukan) as Rujukan, SUM(NonRujukan) as NonRujukan, SUM(Dirawat) as Dirawat, " & _
    " SUM(Dirujuk) as Dirujuk, SUM(pULANG) as PULANG, SUM(mATIDIigd) as mATIDIigd, SUM(Mati) as Mati " & _
    " from RL3_02New2 where Year(Tglmasuk) = '" & dtptahun.Year & "' or tglmasuk is null " & _
    " group by JenisPelayanan "
    Call msubRecFO(rsx, strSQL)

    ProgressBar1.Min = 0
    ProgressBar1.Max = rsx.RecordCount
    ProgressBar1.value = 0

    rsx.MoveFirst

    For i = 1 To rsx.RecordCount
        If rsx!JenisPelayanan = "Bedah" Then
            j = 2
        ElseIf rsx!JenisPelayanan = "Non Bedah" Then
            j = 3
        ElseIf rsx!JenisPelayanan = "Kebidanan" Then
            j = 4
        ElseIf rsx!JenisPelayanan = "Psikiatrik" Then
            j = 5
        ElseIf rsx!JenisPelayanan = "Anak" Then
            j = 6
        End If

        Cell22 = oSheet.Cells(j, 8).value
        Cell23 = oSheet.Cells(j, 9).value
        Cell24 = oSheet.Cells(j, 10).value
        Cell25 = oSheet.Cells(j, 11).value
        Cell26 = oSheet.Cells(j, 12).value
        Cell27 = oSheet.Cells(j, 13).value
        Cell28 = oSheet.Cells(j, 14).value

        If rsx!JenisPelayanan = "Bedah" Then
            With oSheet
                .Cells(j, 8) = Trim(rsx![Rujukan] + Cell22)
                .Cells(j, 9) = Trim(rsx![NonRujukan] + Cell23)
                .Cells(j, 10) = Trim(rsx![DiRawat] + Cell24)
                .Cells(j, 11) = Trim(rsx![DiRujuk] + Cell25)
                .Cells(j, 12) = Trim(rsx![Pulang] + Cell26)
                .Cells(j, 13) = Trim(rsx![MatiDiIGD] + Cell27)
                .Cells(j, 14) = Trim(rsx![Mati] + Cell28)
            End With
        ElseIf rsx!JenisPelayanan = "Non Bedah" Then
            With oSheet
                .Cells(j, 8) = Trim(rsx![Rujukan] + Cell22)
                .Cells(j, 9) = Trim(rsx![NonRujukan] + Cell23)
                .Cells(j, 10) = Trim(rsx![DiRawat] + Cell24)
                .Cells(j, 11) = Trim(rsx![DiRujuk] + Cell25)
                .Cells(j, 12) = Trim(rsx![Pulang] + Cell26)
                .Cells(j, 13) = Trim(rsx![MatiDiIGD] + Cell27)
                .Cells(j, 14) = Trim(rsx![Mati] + Cell28)
            End With
        ElseIf rsx!JenisPelayanan = "Kebidanan" Then
            With oSheet
                .Cells(j, 8) = Trim(rsx![Rujukan] + Cell22)
                .Cells(j, 9) = Trim(rsx![NonRujukan] + Cell23)
                .Cells(j, 10) = Trim(rsx![DiRawat] + Cell24)
                .Cells(j, 11) = Trim(rsx![DiRujuk] + Cell25)
                .Cells(j, 12) = Trim(rsx![Pulang] + Cell26)
                .Cells(j, 13) = Trim(rsx![MatiDiIGD] + Cell27)
                .Cells(j, 14) = Trim(rsx![Mati] + Cell28)
            End With
        ElseIf rsx!JenisPelayanan = "Psikiatrik" Then
            With oSheet
                .Cells(j, 8) = Trim(rsx![Rujukan] + Cell22)
                .Cells(j, 9) = Trim(rsx![NonRujukan] + Cell23)
                .Cells(j, 10) = Trim(rsx![DiRawat] + Cell24)
                .Cells(j, 11) = Trim(rsx![DiRujuk] + Cell25)
                .Cells(j, 12) = Trim(rsx![Pulang] + Cell26)
                .Cells(j, 13) = Trim(rsx![MatiDiIGD] + Cell27)
                .Cells(j, 14) = Trim(rsx![Mati] + Cell28)
            End With
        ElseIf rsx!JenisPelayanan = "Anak" Then
            With oSheet
                .Cells(j, 8) = Trim(rsx![Rujukan] + Cell22)
                .Cells(j, 9) = Trim(rsx![NonRujukan] + Cell23)
                .Cells(j, 10) = Trim(rsx![DiRawat] + Cell24)
                .Cells(j, 11) = Trim(rsx![DiRujuk] + Cell25)
                .Cells(j, 12) = Trim(rsx![Pulang] + Cell26)
                .Cells(j, 13) = Trim(rsx![MatiDiIGD] + Cell27)
                .Cells(j, 14) = Trim(rsx![Mati] + Cell28)
            End With
        End If

        rsx.MoveNext

        ProgressBar1.value = Int(ProgressBar1.value) + 1
        lblPersen.Caption = Int(ProgressBar1.value * 100 / ProgressBar1.Max) & " %"
    Next i

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    MsgBox "Data Tidak Ada", vbInformation, "Validasi"
    Screen.MousePointer = vbDefault
End Sub

