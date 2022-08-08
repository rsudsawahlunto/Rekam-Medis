VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRL3Sub3_4New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.4 Kegiatan Kebidanan"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5445
   Icon            =   "frmRL3Sub3_4New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5445
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   5415
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtptahun 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
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
         Format          =   115408899
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
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
      TabIndex        =   1
      Top             =   3000
      Width           =   4575
      _ExtentX        =   8070
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
      Left            =   4680
      TabIndex        =   2
      Top             =   3120
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL3Sub3_4New2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmRL3Sub3_4New2"
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
Dim i, ii, j, k, l, xx As Integer
Dim w, x, y, z As String
Dim Cell1, Cell2, Cell3, Cell4, Cell5, Cell6, Cell7, Cell8, Cell9, Cell10, Cell11, Cell12, Cell13, Cell14, Cell15 As String
Dim awal As String
Dim akhir As String

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
    Set oWB = oXL.Workbooks.Open(App.Path & "\RL 3.4_kebidanan.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    For xx = 2 To 13
        With oSheet
            .Cells(xx, 1) = Format(dtptahun.value, "YYYY")
            .Cells(xx, 2) = rsb("KodeExternal").value
            .Cells(xx, 3) = rsb("KdRS").value
            .Cells(xx, 4) = rsb("NamaRS").value
            .Cells(xx, 5) = rsb("KotaKodyaKab").value
        End With
    Next xx

    Set rsx = Nothing
    strSQL = "Select * from RL3_04cNew where Year(Tglmasuk) between '" _
    & dtptahun.Year & "' AND '" & dtptahun.Year & "' or tglmasuk is null"
    Call msubRecFO(rsx, strSQL)

    Set rsb = Nothing

    strSQL = "Select * from RL3_04bNew where Year(TglPeriksa) between '" _
    & dtptahun.Year & "' and '" & dtptahun.Year & "' or tglPeriksa is null"

    Call msubRecFO(rsb, strSQL)

    k = rsx.RecordCount + rsb.RecordCount

    If rsx.RecordCount = 0 And rsb.RecordCount = 0 Then
        MsgBox "Data Tidak Ada", vbInformation, "Validasi"
        Screen.MousePointer = vbDefault
        Exit Sub
    ElseIf rsx.RecordCount > 0 And rsb.RecordCount > 0 Then
        GoTo kebidanan
    ElseIf rsx.RecordCount > 0 Then
        GoTo kebidanan
    ElseIf rsb.RecordCount > 0 Then
        GoTo imunisasi
    End If

kebidanan:
    rsx.MoveFirst
    For i = 1 To rsx.RecordCount
        If rsx!TindakanMedis = "Persalinan Normal" Then
            j = 2
        ElseIf rsx!TindakanMedis = "Sectio Caesaria" Then
            j = 3
        ElseIf rsx!TindakanMedis = "Persalinan dengan Komplikasi" Then
            j = 4
        ElseIf rsx!TindakanMedis = "Perdarahan Sebelum Persalinan" Then
            j = 5
        ElseIf rsx!TindakanMedis = "Perdarahan Sedudah Persalinan" Then
            j = 6
        ElseIf rsx!TindakanMedis = "Pre Eclampsi" Then
            j = 7
        ElseIf rsx!TindakanMedis = "Eclampsi" Then
            j = 8
        ElseIf rsx!TindakanMedis = "Infeksi" Then
            j = 9
        ElseIf rsx!TindakanMedis = "Lain - Lain" Then
            j = 10
        ElseIf rsx!TindakanMedis = "Abortus" Then
            j = 11
        End If

        Cell1 = oSheet.Cells(j, 8).value
        Cell2 = oSheet.Cells(j, 9).value
        Cell3 = oSheet.Cells(j, 10).value
        Cell4 = oSheet.Cells(j, 11).value
        Cell5 = oSheet.Cells(j, 12).value
        Cell6 = oSheet.Cells(j, 13).value
        Cell7 = oSheet.Cells(j, 15).value
        Cell8 = oSheet.Cells(j, 16).value
        Cell9 = oSheet.Cells(j, 18).value
        Cell10 = oSheet.Cells(j, 19).value
        Cell11 = oSheet.Cells(j, 21).value

        If rsx!TindakanMedis = "Persalinan Normal" Then
            With oSheet
                .Cells(j, 8) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 9) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 10) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 11) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 15) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 16) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 18) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 19) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 21) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Sectio Caesaria" Then
            With oSheet
                .Cells(j, 8) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 9) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 10) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 11) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 15) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 16) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 18) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 19) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 21) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Persalinan dengan Komplikasi" Then
            With oSheet
                .Cells(j, 8) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 9) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 10) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 11) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 15) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 16) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 18) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 19) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 21) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Perdarahan Sebelum Persalinan" Then
            With oSheet
                .Cells(j, 8) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 9) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 10) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 11) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 15) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 16) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 18) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 19) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 21) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Perdarahan Sedudah Persalinan" Then
            With oSheet
                .Cells(j, 8) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 9) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 10) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 11) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 15) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 16) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 18) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 19) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 21) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Pre Eclampsi" Then
            With oSheet
                .Cells(j, 8) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 9) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 10) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 11) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 15) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 16) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 18) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 19) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 21) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Eclampsi" Then
            With oSheet
                .Cells(j, 8) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 9) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 10) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 11) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 15) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 16) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 18) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 19) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 21) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Infeksi" Then
            With oSheet
                .Cells(j, 8) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 9) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 10) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 11) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 15) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 16) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 18) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 19) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 21) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Lain - Lain" Then
            With oSheet
                .Cells(j, 8) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 9) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 10) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 11) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 15) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 16) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 18) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 19) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 21) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Abortus" Then
            With oSheet
                .Cells(j, 8) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 9) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 10) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 11) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 15) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 16) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 18) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 19) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 21) = Trim(rsx![RujukAtas] + Cell11)
            End With
        End If
        rsx.MoveNext
    Next i

    If rsb.RecordCount = 0 Then
        oXL.Visible = True
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        GoTo imunisasi
    End If

imunisasi:

    ProgressBar1.Min = 0
    ProgressBar1.Max = rsb.RecordCount
    ProgressBar1.value = 0

    rsb.MoveFirst
    For ii = 1 To rsb.RecordCount

        If rsb!NamaImunisasi = "TT1" Then
            j = 12
        ElseIf rsb!NamaImunisasi = "TT2" Then
            j = 13
        End If

        Cell12 = oSheet.Cells(j, 8).value
        Cell13 = oSheet.Cells(j, 9).value
        Cell14 = oSheet.Cells(j, 10).value
        Cell15 = oSheet.Cells(j, 11).value

        If rsb!NamaImunisasi = "TT1" Then
            With oSheet
                .Cells(j, 8) = Trim(rsb![JmlRujukanRS] + Cell12)
                .Cells(j, 9) = Trim(rsb![JmlRujukanBidan] + Cell13)
                .Cells(j, 10) = Trim(rsb![JmlRujukanPskms] + Cell14)
                .Cells(j, 11) = Trim(rsb![JmlRujukanFaskes] + Cell15)
            End With
        ElseIf rsb!NamaImunisasi = "TT2" Then
            With oSheet
                .Cells(j, 8) = Trim(rsb![JmlRujukanRS] + Cell12)
                .Cells(j, 9) = Trim(rsb![JmlRujukanBidan] + Cell13)
                .Cells(j, 10) = Trim(rsb![JmlRujukanPskms] + Cell14)
                .Cells(j, 11) = Trim(rsb![JmlRujukanFaskes] + Cell15)
            End With
        End If

        rsb.MoveNext

        ProgressBar1.value = ProgressBar1.value + 1
        lblPersen.Caption = Int(ProgressBar1.value * 100 / ProgressBar1.Max) & " %"
    Next ii

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    MsgBox "Data Tidak Ada", vbInformation, "Validasi"
    Screen.MousePointer = vbDefault
End Sub

