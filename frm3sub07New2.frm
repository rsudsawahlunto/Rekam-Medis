VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm3sub07New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.07 Kegiatan Radiologi"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6135
   Icon            =   "frm3sub07New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3465
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
         Format          =   137756675
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
      Top             =   3045
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   17
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
      Left            =   5280
      TabIndex        =   5
      Top             =   3120
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm3sub07New2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frm3sub07New2"
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
Dim j, xx As Integer
Dim Cell1 As String

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

    ProgressBar1.value = ProgressBar1.Min
    lblPersen.Caption = "0 %"
    Screen.MousePointer = vbHourglass

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.path & "\RL 3.7_radiologi.xlsx")
    Set oSheet = oWB.ActiveSheet

    Set rsx = Nothing
'    strSQL = "Select distinct * from RL3_07aNew where Year(TglPelayanan) = '" & dtptahun.Year & "'"
    strSQL = "Select JenisKegiatan, COUNT(NoPendaftaran) AS Jumlah, KdJenis from RL3_07New2 where Year(TglPelayanan) = '" & dtptahun.Year & "' group by JenisKegiatan, KdJenis"
    Call msubRecFO(rsx, strSQL)

    If rsx.RecordCount > 0 Then
        rsx.MoveFirst

        While Not rsx.EOF
'            If rsx!Judul = "Foto tanpa bahan kontras" Then
            If rsx!KdJenis = "01" Then
                j = 2
''            ElseIf rsx!Judul = "Foto dengan bahan kontras" Then
            ElseIf rsx!KdJenis = "02" Then
                j = 3
'            ElseIf rsx!Judul = "Foto dengan rol film" Then
            ElseIf rsx!KdJenis = "03" Then
                j = 4
'            ElseIf rsx!Judul = "Flouroskopi" Then
            ElseIf rsx!KdJenis = "04" Then
                j = 5
'            ElseIf rsx!Judul = "Foto Gigi" Then
            ElseIf rsx!KdJenis = "05" Then
                j = 6
'            ElseIf rsx!Judul = "Dento alveolair" Then
            ElseIf rsx!KdJenis = "06" Then
                j = 7
'            ElseIf rsx!Judul = "Panoramic" Then
            ElseIf rsx!KdJenis = "07" Then
                j = 8
'            ElseIf rsx!Judul = "Cephalographi" Then
            ElseIf rsx!KdJenis = "08" Then
                j = 9
'            ElseIf rsx!Judul = "CT Scan" Then
            ElseIf rsx!KdJenis = "09" Then
                j = 10
'            ElseIf rsx!Judul = "C.T. Scan Dikepala" Then
            ElseIf rsx!KdJenis = "10" Then
                j = 11
'            ElseIf rsx!Judul = "C.T. Scan Diluar kepala" Then
            ElseIf rsx!KdJenis = "11" Then
                j = 12
'            ElseIf rsx!Judul = "Lymphografi" Then
            ElseIf rsx!KdJenis = "12" Then
                j = 13
'            ElseIf rsx!Judul = "Angiograpi" Then
            ElseIf rsx!KdJenis = "13" Then
                j = 14
'            ElseIf rsx!Judul = "Lain-lain" Then
            ElseIf rsx!KdJenis = "14" Then
                j = 15
            ElseIf rsx!KdJenis = "15" Then
                j = 16
            ElseIf rsx!KdJenis = "16" Then
                j = 17
            ElseIf rsx!KdJenis = "18" Then
                j = 18
            End If
                
                oSheet.Cells(j, 8) = rsx!Jumlah
'            Cell1 = oSheet.Cells(j, 8).value
'
'            If rsx!Judul = "Foto tanpa bahan kontras" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(rsx!Jumlah + Cell1)
'                End With
'            ElseIf rsx!Judul = "Foto dengan bahan kontras" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(rsx!Jumlah + Cell1)
'                End With
'            ElseIf rsx!Judul = "Foto dengan rol film" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(rsx!Jumlah + Cell1)
'                End With
'            ElseIf rsx!Judul = "Flouroskopi" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(rsx!Jumlah + Cell1)
'                End With
'            ElseIf rsx!Judul = "Dento alveolair" Then
'                With oSheet
'                    .Cells(6, 8) = Trim(rsx!Jumlah + Cell1)
'                End With
'            ElseIf rsx!Judul = "Panoramic" Then
'                With oSheet
'                    .Cells(6, 8) = Trim(rsx!Jumlah + Cell1)
'                End With
'            ElseIf rsx!Judul = "Cephalographi" Then
'                With oSheet
'                    .Cells(6, 8) = Trim(rsx!Jumlah + Cell1)
'                End With
'            ElseIf rsx!Judul = "C.T. Scan Dikepala" Then
'                With oSheet
'                    .Cells(7, 8) = Trim(rsx!Jumlah + Cell1)
'                End With
'            ElseIf rsx!Judul = "C.T. Scan Diluar kepala" Then
'                With oSheet
'                    .Cells(7, 8) = Trim(rsx!Jumlah + Cell1)
'                End With
'            ElseIf rsx!Judul = "Lymphografi" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(rsx!Jumlah + Cell1)
'                End With
'            ElseIf rsx!Judul = "Angiograpi" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(rsx!Jumlah + Cell1)
'                End With
'            ElseIf rsx!Judul = "Lain-lain" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(rsx!Jumlah + Cell1)
'                End With
'            End If
            rsx.MoveNext
        Wend
    End If

'    Set dbRst = Nothing
'    strSQL = "Select distinct * from RL3_07bNew where Year(TglPelayanan) = '" & dtptahun.Year & "'"
'    Call msubRecFO(dbRst, strSQL)
'
'    If dbRst.RecordCount > 0 Then
'        dbRst.MoveFirst
'
'        While Not dbRst.EOF
'            If dbRst!Judul = "Jumlah Kegiatan Radiotherapi" Then
'                j = 11
'            ElseIf dbRst!Judul = "Lain-lain" Then
'                j = 12
'            End If
'
'            Cell1 = oSheet.Cells(j, 8).value
'
'            If dbRst!Judul = "Jumlah Kegiatan Radiotherapi" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!Jumlah + Cell1)
'                End With
'            ElseIf dbRst!Judul = "Lain-lain" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!Jumlah + Cell1)
'                End With
'            End If
'            dbRst.MoveNext
'        Wend
'    End If
'
'    Set dbRst = Nothing
'    strSQL = "Select distinct * from RL3_07cNew where Year(TglPelayanan) = '" & dtptahun.Year & "'"
'    Call msubRecFO(dbRst, strSQL)
'
'    If dbRst.RecordCount > 0 Then
'        dbRst.MoveFirst
'
'        While Not dbRst.EOF
'            If dbRst!Judul = "Jumlah Kegiatan Diagnostik" Then
'                j = 13
'            ElseIf dbRst!Judul = "Jumlah Kegiatan Therapi" Then
'                j = 14
'            ElseIf dbRst!Judul = "Lain-lain" Then
'                j = 15
'            End If
'
'            If oSheet.Cells(j, 8).value = "" Then Cell1 = 0 Else Cell1 = oSheet.Cells(j, 8).value
'
'            If dbRst!Judul = "Jumlah Kegiatan Diagnostik" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!Jumlah + Cell1)
'                End With
'            ElseIf dbRst!Judul = "Jumlah Kegiatan Therapi" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!Jumlah + Cell1)
'                End With
'            ElseIf dbRst!Judul = "Lain-lain" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!Jumlah + Cell1)
'                End With
'            End If
'            dbRst.MoveNext
'        Wend
'    End If
'
'    Set dbRst = Nothing
'    strSQL = "Select distinct * from RL3_07dNew where Year(TglPelayanan) = '" & dtptahun.Year & "'"
'    Call msubRecFO(dbRst, strSQL)
'
'    If dbRst.RecordCount > 0 Then
'        dbRst.MoveFirst
'
'        While Not dbRst.EOF
'            If dbRst!Judul = "USG" Then
'                j = 16
'            ElseIf dbRst!Judul = "MRI" Then
'                j = 17
'            ElseIf dbRst!Judul = "Lain-lain" Then
'                j = 18
'            End If
'
'            If oSheet.Cells(j, 8).value = "" Then Cell1 = 0 Else Cell1 = oSheet.Cells(j, 8).value
'
'            If dbRst!Judul = "USG" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!Jumlah + Cell1)
'                End With
'            ElseIf dbRst!Judul = "MRI" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!Jumlah + Cell1)
'                End With
'            ElseIf dbRst!Judul = "Lain-lain" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!Jumlah + Cell1)
'                End With
'            End If
'
'            dbRst.MoveNext
'        Wend
'    End If

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    ProgressBar1.Min = 0
    ProgressBar1.Max = 17
    ProgressBar1.value = 0

    For xx = 2 To 18
        With oSheet
            .Cells(xx, 3) = rsb("KdRS").value
            .Cells(xx, 2) = rsb("KotaKodyaKab").value
            .Cells(xx, 4) = rsb("NamaRS").value
            .Cells(xx, 5) = Format(dtptahun.value, "YYYY")
        End With
        ProgressBar1.value = Int(ProgressBar1.value) + 1
        lblPersen.Caption = Int(ProgressBar1.value * 100 / ProgressBar1.Max) & " %"
    Next xx

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    MsgBox "Data Tidak Ada", vbInformation, "Validasi"
    Screen.MousePointer = vbDefault
End Sub
