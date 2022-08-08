VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRL3Sub3_4New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.4 Kegiatan Kebidanan"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6135
   Icon            =   "frmRL3Sub3_4New.frx":0000
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
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   720
         TabIndex        =   3
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
         Format          =   130285571
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   375
         Left            =   3240
         TabIndex        =   4
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
         Format          =   129892355
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin VB.Label Label1 
         Caption         =   "s/d"
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   840
         Width           =   375
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   6
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
      TabIndex        =   7
      Top             =   3000
      Width           =   5415
      _ExtentX        =   9551
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
      Left            =   5520
      TabIndex        =   8
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL3Sub3_4New.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmRL3Sub3_4New"
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
Dim Cell1, Cell2, Cell3, Cell4, Cell5, Cell6, Cell7, Cell8, Cell9, Cell10, Cell11, Cell12, Cell13, Cell14, Cell15 As String

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
    Set oWB = oXL.Workbooks.Open(App.Path & "\Formulir RL 3.4.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With oSheet
        .Cells(5, 4) = rsb("KdRS").value
        .Cells(6, 4) = rsb("NamaRS").value
        .Cells(7, 4) = Right(dtpAwal.value, 4)
    End With

    Set rsx = Nothing
    strSQL = "Select * from RL3_04cNew where Tglmasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "'or tglmasuk is null"
    Call msubRecFO(rsx, strSQL)

    Set rsb = Nothing
    strSQL = "Select * from RL3_04bNew where TglPeriksa between '" & Format(dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "' or tglPeriksa is null"
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
        If rsb.RecordCount = 0 Then
            ProgressBar1.Max = rsx.RecordCount
        Else
            ProgressBar1.Max = k
        End If

        If rsx!TindakanMedis = "Persalinan Normal" Then
            j = 14
        ElseIf rsx!TindakanMedis = "Perd Sbl Persalinan" Then
            j = 16
        ElseIf rsx!TindakanMedis = "Perd Sdh Persalinan" Then
            j = 17
        ElseIf rsx!TindakanMedis = "Pre Eclampsi" Then
            j = 18
        ElseIf rsx!TindakanMedis = "Eclampsi" Then
            j = 19
        ElseIf rsx!TindakanMedis = "Infeksi" Then
            j = 20
        ElseIf rsx!TindakanMedis = "Lain-lain" Then
            j = 21
        ElseIf rsx!TindakanMedis = "Sectio Caesaria" Then
            j = 22
        ElseIf rsx!TindakanMedis = "Abortus" Then
            j = 23
        End If

        Cell1 = oSheet.Cells(j, 5).value
        Cell2 = oSheet.Cells(j, 6).value
        Cell3 = oSheet.Cells(j, 7).value
        Cell4 = oSheet.Cells(j, 8).value
        Cell5 = oSheet.Cells(j, 9).value
        Cell6 = oSheet.Cells(j, 10).value
        Cell7 = oSheet.Cells(j, 12).value
        Cell8 = oSheet.Cells(j, 13).value
        Cell9 = oSheet.Cells(j, 15).value
        Cell10 = oSheet.Cells(j, 16).value
        Cell11 = oSheet.Cells(j, 18).value

        If rsx!TindakanMedis = "Persalinan Normal" Then
            With oSheet
                .Cells(j, 5) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 6) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 7) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 8) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 9) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 10) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 15) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 16) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Perd Sbl Persalinan" Then
            With oSheet
                .Cells(j, 5) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 6) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 7) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 8) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 9) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 10) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 15) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 16) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Perd Sdh Persalinan" Then
            With oSheet
                .Cells(j, 5) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 6) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 7) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 8) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 9) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 10) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 15) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 16) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Pre Eclampsi" Then
            With oSheet
                .Cells(j, 5) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 6) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 7) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 8) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 9) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 10) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 15) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 16) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Eclampsi" Then
            With oSheet
                .Cells(j, 5) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 6) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 7) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 8) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 9) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 10) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 15) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 16) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Infeksi" Then
            With oSheet
                .Cells(j, 5) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 6) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 7) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 8) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 9) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 10) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 15) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 16) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Lain-lain" Then
            With oSheet
                .Cells(j, 5) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 6) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 7) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 8) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 9) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 10) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 15) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 16) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Sectio Caesaria" Then
            With oSheet
                .Cells(j, 5) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 6) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 7) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 8) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 9) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 10) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 15) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 16) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell11)
            End With
        ElseIf rsx!TindakanMedis = "Abortus" Then
            With oSheet
                .Cells(j, 5) = Trim(rsx![JmlRujukanRS] + Cell1)
                .Cells(j, 6) = Trim(rsx![JmlRujukanBidan] + Cell2)
                .Cells(j, 7) = Trim(rsx![JmlRujukanPskms] + Cell3)
                .Cells(j, 8) = Trim(rsx![JmlRujukanFaskes] + Cell4)
                .Cells(j, 9) = Trim(rsx![JmlHidupRujukan] + Cell5)
                .Cells(j, 10) = Trim(rsx![MatiRujukan] + Cell6)
                .Cells(j, 12) = Trim(rsx![JmlHidupRujukan] + Cell7)
                .Cells(j, 13) = Trim(rsx![MatiRujukan] + Cell8)
                .Cells(j, 15) = Trim(rsx![JmlHidupNonRujukan] + Cell9)
                .Cells(j, 16) = Trim(rsx![MatiNonRujukan] + Cell10)
                .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell11)
            End With
        End If
        rsx.MoveNext
        ProgressBar1.value = Int(ProgressBar1.value) + 1
        If rsb.RecordCount = 0 Then
            lblPersen.Caption = Int(ProgressBar1.value / rsx.RecordCount * 100) & " %"
        Else
            lblPersen.Caption = Int(ProgressBar1.value / k * 100) & " %"
        End If
    Next i

    If rsb.RecordCount = 0 Then
        oXL.Visible = True
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        GoTo imunisasi
    End If

imunisasi:
    rsb.MoveFirst

    For j = 1 To rsb.RecordCount
        If rsx.RecordCount = 0 Then
            ProgressBar1.Max = rsb.RecordCount
        Else
            ProgressBar1.Max = k
        End If

        If rsb!NamaImunisasi = "TT1" Then
            j = 24
        ElseIf rsb!NamaImunisasi = "TT2" Then
            j = 25
        End If

        Cell12 = oSheet.Cells(j, 5).value
        Cell13 = oSheet.Cells(j, 6).value
        Cell14 = oSheet.Cells(j, 7).value
        Cell15 = oSheet.Cells(j, 8).value

        If rsb!NamaImunisasi = "TT1" Then
            With oSheet
                .Cells(j, 5) = Trim(rsb![JmlRujukanRS] + Cell12)
                .Cells(j, 6) = Trim(rsb![JmlRujukanBidan] + Cell13)
                .Cells(j, 7) = Trim(rsb![JmlRujukanPskms] + Cell14)
                .Cells(j, 8) = Trim(rsb![JmlRujukanFaskes] + Cell15)
            End With
        ElseIf rsb!NamaImunisasi = "TT2" Then
            With oSheet
                .Cells(j, 5) = Trim(rsb![JmlRujukanRS] + Cell12)
                .Cells(j, 6) = Trim(rsb![JmlRujukanBidan] + Cell13)
                .Cells(j, 7) = Trim(rsb![JmlRujukanPskms] + Cell14)
                .Cells(j, 8) = Trim(rsb![JmlRujukanFaskes] + Cell15)
            End With
        End If
        rsb.MoveNext
        ProgressBar1.value = Int(ProgressBar1.value) + 1
        If rsx.RecordCount = 0 Then
            lblPersen.Caption = Int(ProgressBar1.value / rsb.RecordCount * 100) & " %"
        Else
            lblPersen.Caption = Int(ProgressBar1.value / k * 100) & " %"
        End If
    Next j

    If rsx.RecordCount <> 0 Then
        oXL.Visible = True
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        MsgBox "Data Tidak Ada", vbInformation, "Validasi"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Exit Sub
error:
    Call msubPesanError
    Screen.MousePointer = vbDefault
End Sub
