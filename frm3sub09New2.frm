VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm3sub09New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.09 Pelayanan Rehabilitasi Medik"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6135
   Icon            =   "frm3sub09New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
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
         Format          =   126025731
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
      TabIndex        =   5
      Top             =   3120
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
      TabIndex        =   6
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm3sub09New2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frm3sub09New2"
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
Dim Cell6 As String
Dim Cell11 As String

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
    Set oWB = oXL.Workbooks.Open(App.path & "\RL 3.9_rehab medik.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    For xx = 2 To 48
        With oSheet
            .Cells(xx, 3) = rsb("KdRS").value
            .Cells(xx, 2) = rsb("KotaKodyaKab").value
            .Cells(xx, 4) = rsb("NamaRS").value
            .Cells(xx, 5) = Format(dtptahun.value, "YYYY")
        End With
    Next xx

    Set dbRst = Nothing
'    strSQL = "Select distinct * from RL3_09New "
    strSQL = "select NamaTindakan, COUNT(NoPendaftaran) AS Jumlah, KdJenisTindakan from RL3_09New2 WHERE YEAR(TglPelayanan) = '" & dtptahun.Year & "' group by NamaTindakan, KdJenisTindakan"
    Call msubRecFO(dbRst, strSQL)

    ProgressBar1.Min = 0
    ProgressBar1.Max = dbRst.RecordCount
    ProgressBar1.value = 0

    If dbRst.RecordCount > 0 Then

        dbRst.MoveFirst

        While Not dbRst.EOF
'            If dbRst!TindakanMedis = "Gait Analyzer" Then
            If dbRst!KdJenisTindakan = "01" Then
                j = 3
'            ElseIf dbRst!TindakanMedis = "EMG" Then
            ElseIf dbRst!KdJenisTindakan = "02" Then
                j = 4
'            ElseIf dbRst!TindakanMedis = "Uro Dinamic" Then
            ElseIf dbRst!KdJenisTindakan = "03" Then
                j = 5
'            ElseIf dbRst!TindakanMedis = "Sideback" Then
            ElseIf dbRst!KdJenisTindakan = "04" Then
                j = 6
'            ElseIf dbRst!TindakanMedis = "E N Tree" Then
            ElseIf dbRst!KdJenisTindakan = "05" Then
                j = 7
'            ElseIf dbRst!TindakanMedis = "Spyrometer" Then
            ElseIf dbRst!KdJenisTindakan = "06" Then
                j = 8
'            ElseIf dbRst!TindakanMedis = "Static Bicycle" Then
            ElseIf dbRst!KdJenisTindakan = "07" Then
                j = 9
'            ElseIf dbRst!TindakanMedis = "Tread Mill" Then
            ElseIf dbRst!KdJenisTindakan = "08" Then
                j = 10
'            ElseIf dbRst!TindakanMedis = "Body Platismograf" Then
            ElseIf dbRst!KdJenisTindakan = "09" Then
                j = 11
            ElseIf dbRst!KdJenisTindakan = "10" Then
                j = 12
'
'            ElseIf dbRst!TindakanMedis = "Latihan Fisik" Then
            ElseIf dbRst!KdJenisTindakan = "11" Then
                j = 14
'            ElseIf dbRst!TindakanMedis = "Aktinoterapi" Then
            ElseIf dbRst!KdJenisTindakan = "12" Then
                j = 15
'            ElseIf dbRst!TindakanMedis = "Elektroterapi" Then
            ElseIf dbRst!KdJenisTindakan = "13" Then
                j = 16
'            ElseIf dbRst!TindakanMedis = "Hidroterapi" Then
            ElseIf dbRst!KdJenisTindakan = "14" Then
                j = 17
'            ElseIf dbRst!TindakanMedis = "Traksi Lumbal & Cervical" Then
            ElseIf dbRst!KdJenisTindakan = "15" Then
                j = 18
'            ElseIf dbRst!TindakanMedis = "Fisioterapi Lain" Then
            ElseIf dbRst!KdJenisTindakan = "16" Then
                j = 19

'            ElseIf dbRst!TindakanMedis = "Snoosien Room" Then
            ElseIf dbRst!KdJenisTindakan = "17" Then
                j = 21
'            ElseIf dbRst!TindakanMedis = "Sensori Integrasi" Then
            ElseIf dbRst!KdJenisTindakan = "18" Then
                j = 22
'            ElseIf dbRst!TindakanMedis = "Latihan Aktivitas Kehidupan Sehari-hari" Then
            ElseIf dbRst!KdJenisTindakan = "19" Then
                j = 23

'            ElseIf dbRst!TindakanMedis = "Proper Body Mekanik" Then
            ElseIf dbRst!KdJenisTindakan = "20" Then
                j = 24
'            ElseIf dbRst!TindakanMedis = "Pembuatan Alat lontar & Adaptasi Alat" Then
            ElseIf dbRst!KdJenisTindakan = "21" Then
                j = 25
'            ElseIf dbRst!TindakanMedis = "Analisa Persiapan Kerja" Then
            ElseIf dbRst!KdJenisTindakan = "22" Then
                j = 26
'            ElseIf dbRst!TindakanMedis = "Latihan Relaksasi" Then
            ElseIf dbRst!KdJenisTindakan = "23" Then
                j = 27
'            ElseIf dbRst!TindakanMedis = "Analisa & Intervensi, Persepsi, Kognitif" Then
            ElseIf dbRst!KdJenisTindakan = "24" Then
                j = 28
            ElseIf dbRst!KdJenisTindakan = "25" Then
                j = 29

'            ElseIf dbRst!TindakanMedis = "Fungsi Bicara" Then
            ElseIf dbRst!KdJenisTindakan = "26" Then
                j = 31
'            ElseIf dbRst!TindakanMedis = "Fungsi Bahasa & Laku" Then
            ElseIf dbRst!KdJenisTindakan = "27" Then
                j = 32
'            ElseIf dbRst!TindakanMedis = "Fungsi Menelan" Then
            ElseIf dbRst!KdJenisTindakan = "28" Then
                j = 33
            ElseIf dbRst!KdJenisTindakan = "29" Then
                j = 34

'            ElseIf dbRst!TindakanMedis = "Psikologi Anak" Then
            ElseIf dbRst!KdJenisTindakan = "30" Then
                j = 36
'            ElseIf dbRst!TindakanMedis = "Psikologi Dewasa" Then
            ElseIf dbRst!KdJenisTindakan = "31" Then
                j = 37
            ElseIf dbRst!KdJenisTindakan = "32" Then
                j = 38

'            ElseIf dbRst!TindakanMedis = "Evaluasi Lingkungan Rumah" Then
            ElseIf dbRst!KdJenisTindakan = "33" Then
                j = 40
'            ElseIf dbRst!TindakanMedis = "Evaluasi Ekonomi" Then
            ElseIf dbRst!KdJenisTindakan = "34" Then
                j = 41
'            ElseIf dbRst!TindakanMedis = "Evaluasi Pekerjaan" Then
            ElseIf dbRst!KdJenisTindakan = "35" Then
                j = 42
            ElseIf dbRst!KdJenisTindakan = "36" Then
                j = 43
                
'            ElseIf dbRst!TindakanMedis = "Pembuatan Alat Bantu" Then
            ElseIf dbRst!KdJenisTindakan = "37" Then
                j = 45
'            ElseIf dbRst!TindakanMedis = "Pembuatan Alat Anggota Tiruan" Then
            ElseIf dbRst!KdJenisTindakan = "38" Then
                j = 46
'            ElseIf dbRst!TindakanMedis = "Ortotik Prostetik Lain" Then
            ElseIf dbRst!KdJenisTindakan = "39" Then
                j = 47
            End If
            
            With oSheet
                If dbRst!KdJenisTindakan <> "" Then
                    .Cells(j, 8) = dbRst!Jumlah
                End If
            End With

'            Cell6 = oSheet.Cells(j, 8).value
'
'            If dbRst!TindakanMedis = "Gait Analyzer" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "EMG" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Uro Dinamic" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Sideback" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "E N Tree" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Spyrometer" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Static Bicycle" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Tread Mill" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Body Platismograf" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'
'            ElseIf dbRst!TindakanMedis = "Latihan Fisik" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Aktinoterapi" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Elektroterapi" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Hidroterapi" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Traksi Lumbal & Cervical" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Fisioterapi Lain" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'
'            ElseIf dbRst!TindakanMedis = "Snoosien Room" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Sensori Integrasi" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Latihan Aktivitas Kehidupan Sehari-hari" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'
'            ElseIf dbRst!TindakanMedis = "Proper Body Mekanik" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Pembuatan Alat lontar & Adaptasi Alat" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Analisa Persiapan Kerja" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Latihan Relaksasi" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Analisa & Intervensi, Persepsi, Kognitif" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'
'            ElseIf dbRst!TindakanMedis = "Fungsi Bicara" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Fungsi Bahasa & Laku" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Fungsi Menelan" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'
'            ElseIf dbRst!TindakanMedis = "Psikologi Anak" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Psikologi Dewasa" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'
'            ElseIf dbRst!TindakanMedis = "Evaluasi Lingkungan Rumah" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Evaluasi Ekonomi" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Evaluasi Pekerjaan" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'
'            ElseIf dbRst!TindakanMedis = "Pembuatan Alat Bantu" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Pembuatan Alat Anggota Tiruan" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            ElseIf dbRst!TindakanMedis = "Ortotik Prostetik Lain" Then
'                With oSheet
'                    .Cells(j, 8) = Trim(dbRst!JmlTindakan + Cell6)
'                End With
'            End If
'
'            With oSheet
'                .Cells(48, 8) = Trim(dbRst!KunjunganRumah + Cell6)
'            End With

            dbRst.MoveNext

            ProgressBar1.value = Int(ProgressBar1.value) + 1
            lblPersen.Caption = Int(ProgressBar1.value * 100 / ProgressBar1.Max) & " %"
        Wend
    End If

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    MsgBox "Data Tidak Ada", vbInformation, "Validasi"
    Screen.MousePointer = vbDefault
End Sub
