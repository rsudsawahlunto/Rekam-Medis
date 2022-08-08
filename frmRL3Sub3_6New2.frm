VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRL3Sub3_6New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.6 Kegiatan Pembedahan"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6135
   Icon            =   "frmRL3Sub3_6New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6135
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   6135
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   3240
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   1080
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtptahun 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   480
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
         Format          =   137691139
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   1
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
      TabIndex        =   2
      Top             =   3000
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
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
      TabIndex        =   3
      Top             =   3120
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL3Sub3_6New2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmRL3Sub3_6New2"
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
Dim Cell1, Cell2, Cell3, Cell4 As String

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
    Set oWB = oXL.Workbooks.Open(App.path & "\RL 3.6_pembedahan.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    For xx = 2 To 15
        With oSheet
            .Cells(xx, 1) = rsb("KodeExternal")
            .Cells(xx, 2) = rsb("KotaKodyaKab").value
            .Cells(xx, 3) = rsb("KdRS").value
            .Cells(xx, 4) = rsb("NamaRS").value
            .Cells(xx, 5) = Format(dtptahun.value, "YYYY")
        End With
    Next xx

    Set rsx = Nothing
'    strSQL = "Select distinct * from RL3_06New where Year(Tglpelayanan) between '" _
'    & dtptahun.Year & "' and '" & dtptahun.Year & "' or tglpelayanan is null"
    strSQL = "select NamaTindakan, SUM(Khusus) as Khusus,SUM(Besar) as Besar,SUM(Sedang) as Sedang,SUM(Kecil) as Kecil, KdJenis " & _
             "from RL3_06New2 where YEAR(TglPelayanan)='" & dtptahun.Year & "' group by NamaTindakan, KdJenis"
    Call msubRecFO(rsx, strSQL)

    ProgressBar1.Min = 0
    ProgressBar1.Max = rsx.RecordCount
    ProgressBar1.value = 0

    rsx.MoveFirst

    For i = 1 To rsx.RecordCount

'        If rsx!Spesialisasi = "Bedah" Then
        If rsx!KdJenis = "01" Then
            j = 2
'        ElseIf rsx!Spesialisasi = "Obstetrik &Ginekologi" Then
        ElseIf rsx!KdJenis = "02" Then
            j = 3
'        ElseIf rsx!Spesialisasi = "Bedah Saraf" Then
        ElseIf rsx!KdJenis = "03" Then
            j = 4
'        ElseIf rsx!Spesialisasi = "THT" Then
        ElseIf rsx!KdJenis = "04" Then
            j = 5
'        ElseIf rsx!Spesialisasi = "Mata" Then
        ElseIf rsx!KdJenis = "05" Then
            j = 6
'        ElseIf rsx!Spesialisasi = "Kulit & Kelamin" Then
        ElseIf rsx!KdJenis = "06" Then
            j = 7
'        ElseIf rsx!Spesialisasi = "Gigi & Mulut" Then
        ElseIf rsx!KdJenis = "07" Then
            j = 8
'        ElseIf rsx!Spesialisasi = "Bedah Anak" Then
        ElseIf rsx!KdJenis = "08" Then
            j = 9
'        ElseIf rsx!Spesialisasi = "Kardiologi" Then 'Kardiovaskuler
        ElseIf rsx!KdJenis = "09" Then
            j = 10
'        ElseIf rsx!Spesialisasi = "Ortopedi" Then 'Bedah Orthopedi
        ElseIf rsx!KdJenis = "10" Then
            j = 11
        ElseIf rsx!Spesialisasi = "Thorak" Then
        ElseIf rsx!KdJenis = "11" Then
            j = 12
''        ElseIf rsx!Spesialisasi = "Digestive" Then
        ElseIf rsx!KdJenis = "12" Then
            j = 13
'        ElseIf rsx!Spesialisasi = "Urologi" Then
        ElseIf rsx!KdJenis = "13" Then
            j = 14
'        ElseIf rsx!Spesialisasi = "Lain-lain" Then
        ElseIf rsx!KdJenis = "14" Then
            j = 15
        Else
            j = 15
        End If
        
        With oSheet
            .Cells(j, 9) = IIf(IsNull(rsx![Khusus]), 0, rsx![Khusus])
            .Cells(j, 10) = IIf(IsNull(rsx![Besar]), 0, rsx![Besar])
            .Cells(j, 11) = IIf(IsNull(rsx![Sedang]), 0, rsx![Sedang])
            .Cells(j, 12) = IIf(IsNull(rsx![Kecil]), 0, rsx![Kecil])
        End With

'        Cell1 = oSheet.Cells(j, 9).value
'        Cell2 = oSheet.Cells(j, 10).value
'        Cell3 = oSheet.Cells(j, 11).value
'        Cell4 = oSheet.Cells(j, 12).value
'
'        If rsx!Spesialisasi = "Bedah" Then
'            With oSheet
'                .Cells(j, 9) = Trim(rsx![Khusus] + Cell1)
'                .Cells(j, 10) = Trim(rsx![Besar] + Cell2)
'                .Cells(j, 11) = Trim(rsx![Sedang] + Cell3)
'                .Cells(j, 12) = Trim(rsx![Kecil] + Cell4)
'            End With
'        ElseIf rsx!Spesialisasi = "Obstetrik &Ginekologi" Then
'            With oSheet
'                .Cells(j, 9) = Trim(rsx![Khusus] + Cell1)
'                .Cells(j, 10) = Trim(rsx![Besar] + Cell2)
'                .Cells(j, 11) = Trim(rsx![Sedang] + Cell3)
'                .Cells(j, 12) = Trim(rsx![Kecil] + Cell4)
'            End With
'        ElseIf rsx!Spesialisasi = "Bedah Saraf" Then
'            With oSheet
'                .Cells(j, 9) = Trim(rsx![Khusus] + Cell1)
'                .Cells(j, 10) = Trim(rsx![Besar] + Cell2)
'                .Cells(j, 11) = Trim(rsx![Sedang] + Cell3)
'                .Cells(j, 12) = Trim(rsx![Kecil] + Cell4)
'            End With
'        ElseIf rsx!Spesialisasi = "THT" Then
'            With oSheet
'                .Cells(j, 9) = Trim(rsx![Khusus] + Cell1)
'                .Cells(j, 10) = Trim(rsx![Besar] + Cell2)
'                .Cells(j, 11) = Trim(rsx![Sedang] + Cell3)
'                .Cells(j, 12) = Trim(rsx![Kecil] + Cell4)
'            End With
'        ElseIf rsx!Spesialisasi = "Mata" Then
'            With oSheet
'                .Cells(j, 9) = Trim(rsx![Khusus] + Cell1)
'                .Cells(j, 10) = Trim(rsx![Besar] + Cell2)
'                .Cells(j, 11) = Trim(rsx![Sedang] + Cell3)
'                .Cells(j, 12) = Trim(rsx![Kecil] + Cell4)
'            End With
'        ElseIf rsx!Spesialisasi = "Kulit & Kelamin" Then
'            With oSheet
'                .Cells(j, 9) = Trim(rsx![Khusus] + Cell1)
'                .Cells(j, 10) = Trim(rsx![Besar] + Cell2)
'                .Cells(j, 11) = Trim(rsx![Sedang] + Cell3)
'                .Cells(j, 12) = Trim(rsx![Kecil] + Cell4)
'            End With
'        ElseIf rsx!Spesialisasi = "Gigi & Mulut" Then
'            With oSheet
'                .Cells(j, 9) = Trim(rsx![Khusus] + Cell1)
'                .Cells(j, 10) = Trim(rsx![Besar] + Cell2)
'                .Cells(j, 11) = Trim(rsx![Sedang] + Cell3)
'                .Cells(j, 12) = Trim(rsx![Kecil] + Cell4)
'            End With
'        ElseIf rsx!Spesialisasi = "Bedah Anak" Then
'            With oSheet
'                .Cells(j, 9) = Trim(rsx![Khusus] + Cell1)
'                .Cells(j, 10) = Trim(rsx![Besar] + Cell2)
'                .Cells(j, 11) = Trim(rsx![Sedang] + Cell3)
'                .Cells(j, 12) = Trim(rsx![Kecil] + Cell4)
'            End With
'        ElseIf rsx!Spesialisasi = "Kardiologi" Then
'            With oSheet
'                .Cells(j, 9) = Trim(rsx![Khusus] + Cell1)
'                .Cells(j, 10) = Trim(rsx![Besar] + Cell2)
'                .Cells(j, 11) = Trim(rsx![Sedang] + Cell3)
'                .Cells(j, 12) = Trim(rsx![Kecil] + Cell4)
'            End With
'        ElseIf rsx!Spesialisasi = "Ortopedi" Then
'            With oSheet
'                .Cells(j, 9) = Trim(rsx![Khusus] + Cell1)
'                .Cells(j, 10) = Trim(rsx![Besar] + Cell2)
'                .Cells(j, 11) = Trim(rsx![Sedang] + Cell3)
'                .Cells(j, 12) = Trim(rsx![Kecil] + Cell4)
'            End With
'        ElseIf rsx!Spesialisasi = "Thorak" Then
'            With oSheet
'                .Cells(j, 9) = Trim(rsx![Khusus] + Cell1)
'                .Cells(j, 10) = Trim(rsx![Besar] + Cell2)
'                .Cells(j, 11) = Trim(rsx![Sedang] + Cell3)
'                .Cells(j, 12) = Trim(rsx![Kecil] + Cell4)
'            End With
'        ElseIf rsx!Spesialisasi = "Digestive" Then
'            With oSheet
'                .Cells(j, 9) = Trim(rsx![Khusus] + Cell1)
'                .Cells(j, 10) = Trim(rsx![Besar] + Cell2)
'                .Cells(j, 11) = Trim(rsx![Sedang] + Cell3)
'                .Cells(j, 12) = Trim(rsx![Kecil] + Cell4)
'            End With
'        ElseIf rsx!Spesialisasi = "Urologi" Then
'            With oSheet
'                .Cells(j, 9) = Trim(rsx![Khusus] + Cell1)
'                .Cells(j, 10) = Trim(rsx![Besar] + Cell2)
'                .Cells(j, 11) = Trim(rsx![Sedang] + Cell3)
'                .Cells(j, 12) = Trim(rsx![Kecil] + Cell4)
'            End With
'        ElseIf rsx!Spesialisasi = "Lain-lain" Then
'            With oSheet
'                .Cells(j, 9) = Trim(rsx![Khusus] + Cell1)
'                .Cells(j, 10) = Trim(rsx![Besar] + Cell2)
'                .Cells(j, 11) = Trim(rsx![Sedang] + Cell3)
'                .Cells(j, 12) = Trim(rsx![Kecil] + Cell4)
'            End With
'        Else
'            With oSheet
'                .Cells(j, 9) = Trim(rsx![Khusus] + Cell1)
'                .Cells(j, 10) = Trim(rsx![Besar] + Cell2)
'                .Cells(j, 11) = Trim(rsx![Sedang] + Cell3)
'                .Cells(j, 12) = Trim(rsx![Kecil] + Cell4)
'            End With
'        End If
        rsx.MoveNext

        ProgressBar1.value = ProgressBar1.value + 1
        lblPersen.Caption = Int(ProgressBar1.value * 100 / ProgressBar1.Max) & " %"
    Next i

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    Call msubPesanError
    Screen.MousePointer = vbDefault
End Sub

