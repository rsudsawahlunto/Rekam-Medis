VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRL3Sub3_2New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.2 Kunjungan Rawat Darurat"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmRL3Sub3_2New.frx":0000
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
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   1320
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   720
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   134807555
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   720
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   133169155
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin VB.Label Label1 
         Caption         =   "s/d"
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   840
         Width           =   375
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
      Picture         =   "frmRL3Sub3_2New.frx":0CCA
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
Attribute VB_Name = "frmRL3Sub3_2New"
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
Dim i, j, k, l As Integer
Dim w, X, y, z As String
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

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpAwal.value = Format(Now, "dd/mm/yyyy")
    dtpAkhir.value = Format(Now, "dd/mm/yyyy")

    ProgressBar1.value = ProgressBar1.Min
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo error

    ProgressBar1.value = ProgressBar1.Min
    lblPersen.Caption = "0 %"
    Screen.MousePointer = vbHourglass

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.Path & "\Formulir RL 3.2.xlsx")
    Set oSheet = oWB.ActiveSheet

    Set rsx = Nothing
    strSQL = "Select * from RL3_02New where Tglmasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "'or tglmasuk is null"
    Call msubRecFO(rsx, strSQL)

    rsx.MoveFirst

    For i = 1 To rsx.RecordCount
        ProgressBar1.Max = rsx.RecordCount

        If rsx!JenisPelayanan = "Bedah" Then
            j = 15
        ElseIf rsx!JenisPelayanan = "NonBedah" Then
            j = 16
        ElseIf rsx!JenisPelayanan = "Kebidanan" Then
            j = 17
        ElseIf rsx!JenisPelayanan = "Psikiatrik" Then
            j = 18
        ElseIf rsx!JenisPelayanan = "Anak" Then
            j = 19
        End If

        Cell22 = oSheet.Cells(j, 5).value
        Cell23 = oSheet.Cells(j, 6).value
        Cell24 = oSheet.Cells(j, 7).value
        Cell25 = oSheet.Cells(j, 8).value
        Cell26 = oSheet.Cells(j, 9).value
        Cell27 = oSheet.Cells(j, 10).value
        Cell28 = oSheet.Cells(j, 11).value

        If rsx!JenisPelayanan = "Bedah" Then
            With oSheet
                .Cells(j, 5) = Trim(rsx![Rujukan] + Cell22)
                .Cells(j, 6) = Trim(rsx![NonRujukan] + Cell23)
                .Cells(j, 7) = Trim(rsx![DiRawat] + Cell24)
                .Cells(j, 8) = Trim(rsx![DiRujuk] + Cell25)
                .Cells(j, 9) = Trim(rsx![Pulang] + Cell26)
                .Cells(j, 10) = Trim(rsx![MatiDiIGD] + Cell27)
                .Cells(j, 11) = Trim(rsx![Mati] + Cell28)
            End With
        ElseIf rsx!JenisPelayanan = "NonBedah" Then
            With oSheet
                .Cells(j, 5) = Trim(rsx![Rujukan] + Cell22)
                .Cells(j, 6) = Trim(rsx![NonRujukan] + Cell23)
                .Cells(j, 7) = Trim(rsx![DiRawat] + Cell24)
                .Cells(j, 8) = Trim(rsx![DiRujuk] + Cell25)
                .Cells(j, 9) = Trim(rsx![Pulang] + Cell26)
                .Cells(j, 10) = Trim(rsx![MatiDiIGD] + Cell27)
                .Cells(j, 11) = Trim(rsx![Mati] + Cell28)
            End With
        ElseIf rsx!JenisPelayanan = "Kebidanan" Then
            With oSheet
                .Cells(j, 5) = Trim(rsx![Rujukan] + Cell22)
                .Cells(j, 6) = Trim(rsx![NonRujukan] + Cell23)
                .Cells(j, 7) = Trim(rsx![DiRawat] + Cell24)
                .Cells(j, 8) = Trim(rsx![DiRujuk] + Cell25)
                .Cells(j, 9) = Trim(rsx![Pulang] + Cell26)
                .Cells(j, 10) = Trim(rsx![MatiDiIGD] + Cell27)
                .Cells(j, 11) = Trim(rsx![Mati] + Cell28)
            End With
        ElseIf rsx!JenisPelayanan = "Psikiatrik" Then
            With oSheet
                .Cells(j, 5) = Trim(rsx![Rujukan] + Cell22)
                .Cells(j, 6) = Trim(rsx![NonRujukan] + Cell23)
                .Cells(j, 7) = Trim(rsx![DiRawat] + Cell24)
                .Cells(j, 8) = Trim(rsx![DiRujuk] + Cell25)
                .Cells(j, 9) = Trim(rsx![Pulang] + Cell26)
                .Cells(j, 10) = Trim(rsx![MatiDiIGD] + Cell27)
                .Cells(j, 11) = Trim(rsx![Mati] + Cell28)
            End With
        ElseIf rsx!JenisPelayanan = "Anak" Then
            With oSheet
                .Cells(j, 5) = Trim(rsx![Rujukan] + Cell22)
                .Cells(j, 6) = Trim(rsx![NonRujukan] + Cell23)
                .Cells(j, 7) = Trim(rsx![DiRawat] + Cell24)
                .Cells(j, 8) = Trim(rsx![DiRujuk] + Cell25)
                .Cells(j, 9) = Trim(rsx![Pulang] + Cell26)
                .Cells(j, 10) = Trim(rsx![MatiDiIGD] + Cell27)
                .Cells(j, 11) = Trim(rsx![Mati] + Cell28)
            End With
        End If
        rsx.MoveNext
        ProgressBar1.value = Int(ProgressBar1.value) + 1
        lblPersen.Caption = Int(ProgressBar1.value / rsx.RecordCount * 100) & " %"
    Next i

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With oSheet
        .Cells(7, 4) = rsb("KdRS").value
        .Cells(8, 4) = rsb("NamaRS").value
        .Cells(9, 4) = Right(dtpAwal.value, 4)
    End With

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    MsgBox "Data Tidak Ada", vbInformation, "Validasi"
    Screen.MousePointer = vbDefault
End Sub
