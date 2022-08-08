VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm3sub09New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.09 Pelayanan Rehabilitasi Medik"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6135
   Icon            =   "frm3sub09New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
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
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy "
         Format          =   130285571
         UpDown          =   -1  'True
         CurrentDate     =   38212
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   375
         Left            =   3240
         TabIndex        =   6
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   133300227
         UpDown          =   -1  'True
         CurrentDate     =   38212
      End
      Begin VB.Label Label1 
         Caption         =   "s/d"
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
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
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm3sub09New.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frm3sub09New"
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
Dim j As Integer
Dim Cell6 As String
Dim Cell11 As String

'Special Buat Excel
Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpAwal.value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.value = Now
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo error

    Screen.MousePointer = vbHourglass

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.Path & "\Formulir RL 3.9.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With oSheet
        .Cells(7, 4) = rsb("KdRS").value
        .Cells(8, 4) = rsb("NamaRS").value
        .Cells(9, 4) = Right(dtpAwal.value, 4)
    End With

    Set dbRst = Nothing
    strSQL = "Select distinct * from RL3_09New where TglPelayanan between '" & Format(dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "'"

    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF
            If dbRst!TindakanMedis = "Gait Analyzer" Then
                j = 13
            ElseIf dbRst!TindakanMedis = "EMG" Then
                j = 14
            ElseIf dbRst!TindakanMedis = "Uro Dinamic" Then
                j = 15
            ElseIf dbRst!TindakanMedis = "Sideback" Then
                j = 16
            ElseIf dbRst!TindakanMedis = "E N Tree" Then
                j = 17
            ElseIf dbRst!TindakanMedis = "Spyrometer" Then
                j = 18
            ElseIf dbRst!TindakanMedis = "Static Bicycle" Then
                j = 19
            ElseIf dbRst!TindakanMedis = "Tread Mill" Then
                j = 20
            ElseIf dbRst!TindakanMedis = "Body Platismograf" Then
                j = 21

            ElseIf dbRst!TindakanMedis = "Latihan Fisik" Then
                j = 24
            ElseIf dbRst!TindakanMedis = "Aktinoterapi" Then
                j = 25
            ElseIf dbRst!TindakanMedis = "Elektroterapi" Then
                j = 26
            ElseIf dbRst!TindakanMedis = "Hidroterapi" Then
                j = 27
            ElseIf dbRst!TindakanMedis = "Traksi Lumbal & Cervical" Then
                j = 28
            ElseIf dbRst!TindakanMedis = "Fisioterapi Lain" Then
                j = 29

            ElseIf dbRst!TindakanMedis = "Snoosien Room" Then
                j = 31
            ElseIf dbRst!TindakanMedis = "Sensori Integrasi" Then
                j = 32
            ElseIf dbRst!TindakanMedis = "Latihan Aktivitas Kehidupan Sehari-hari" Then
                j = 33

            ElseIf dbRst!TindakanMedis = "Proper Body Mekanik" Then
                j = 34
            ElseIf dbRst!TindakanMedis = "Pembuatan Alat lontar & Adaptasi Alat" Then
                j = 35
            ElseIf dbRst!TindakanMedis = "Analisa Persiapan Kerja" Then
                j = 12
            ElseIf dbRst!TindakanMedis = "Latihan Relaksasi" Then
                j = 13
            ElseIf dbRst!TindakanMedis = "Analisa & Intervensi, Persepsi, Kognitif" Then
                j = 14

            ElseIf dbRst!TindakanMedis = "Fungsi Bicara" Then
                j = 17
            ElseIf dbRst!TindakanMedis = "Fungsi Bahasa & Laku" Then
                j = 18
            ElseIf dbRst!TindakanMedis = "Fungsi Menelan" Then
                j = 19

            ElseIf dbRst!TindakanMedis = "Psikologi Anak" Then
                j = 22
            ElseIf dbRst!TindakanMedis = "Psikologi Dewasa" Then
                j = 23

            ElseIf dbRst!TindakanMedis = "Evaluasi Lingkungan Rumah" Then
                j = 26
            ElseIf dbRst!TindakanMedis = "Evaluasi Ekonomi" Then
                j = 27
            ElseIf dbRst!TindakanMedis = "Evaluasi Pekerjaan" Then
                j = 28

            ElseIf dbRst!TindakanMedis = "Pembuatan Alat Bantu" Then
                j = 31
            ElseIf dbRst!TindakanMedis = "Pembuatan Alat Anggota Tiruan" Then
                j = 32
            ElseIf dbRst!TindakanMedis = "Ortotik Prostetik Lain" Then
                j = 33
            End If

            Cell6 = oSheet.Cells(j, 6).value
            Cell11 = oSheet.Cells(j, 11).value

            If dbRst!TindakanMedis = "Gait Analyzer" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "EMG" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Uro Dinamic" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Sideback" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "E N Tree" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Spyrometer" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Static Bicycle" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Tread Mill" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Body Platismograf" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With

            ElseIf dbRst!TindakanMedis = "Latihan Fisik" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Aktinoterapi" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Elektroterapi" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Hidroterapi" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Traksi Lumbal & Cervical" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Fisioterapi Lain" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With

            ElseIf dbRst!TindakanMedis = "Snoosien Room" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Sensori Integrasi" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Latihan Aktivitas Kehidupan Sehari-hari" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With

            ElseIf dbRst!TindakanMedis = "Proper Body Mekanik" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Pembuatan Alat lontar & Adaptasi Alat" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Analisa Persiapan Kerja" Then
                With oSheet
                    .Cells(j, 11) = Trim(dbRst!JmlTindakan + Cell11)
                End With
            ElseIf dbRst!TindakanMedis = "Latihan Relaksasi" Then
                With oSheet
                    .Cells(j, 11) = Trim(dbRst!JmlTindakan + Cell11)
                End With
            ElseIf dbRst!TindakanMedis = "Analisa & Intervensi, Persepsi, Kognitif" Then
                With oSheet
                    .Cells(j, 11) = Trim(dbRst!JmlTindakan + Cell11)
                End With

            ElseIf dbRst!TindakanMedis = "Fungsi Bicara" Then
                With oSheet
                    .Cells(j, 11) = Trim(dbRst!JmlTindakan + Cell11)
                End With
            ElseIf dbRst!TindakanMedis = "Fungsi Bahasa & Laku" Then
                With oSheet
                    .Cells(j, 11) = Trim(dbRst!JmlTindakan + Cell11)
                End With
            ElseIf dbRst!TindakanMedis = "Fungsi Menelan" Then
                With oSheet
                    .Cells(j, 11) = Trim(dbRst!JmlTindakan + Cell11)
                End With

            ElseIf dbRst!TindakanMedis = "Psikologi Anak" Then
                With oSheet
                    .Cells(j, 11) = Trim(dbRst!JmlTindakan + Cell11)
                End With
            ElseIf dbRst!TindakanMedis = "Psikologi Dewasa" Then
                With oSheet
                    .Cells(j, 11) = Trim(dbRst!JmlTindakan + Cell11)
                End With

            ElseIf dbRst!TindakanMedis = "Evaluasi Lingkungan Rumah" Then
                With oSheet
                    .Cells(j, 11) = Trim(dbRst!JmlTindakan + Cell11)
                End With
            ElseIf dbRst!TindakanMedis = "Evaluasi Ekonomi" Then
                With oSheet
                    .Cells(j, 11) = Trim(dbRst!JmlTindakan + Cell11)
                End With
            ElseIf dbRst!TindakanMedis = "Evaluasi Pekerjaan" Then
                With oSheet
                    .Cells(j, 11) = Trim(dbRst!JmlTindakan + Cell11)
                End With

            ElseIf dbRst!TindakanMedis = "Pembuatan Alat Bantu" Then
                With oSheet
                    .Cells(j, 11) = Trim(dbRst!JmlTindakan + Cell11)
                End With
            ElseIf dbRst!TindakanMedis = "Pembuatan Alat Anggota Tiruan" Then
                With oSheet
                    .Cells(j, 11) = Trim(dbRst!JmlTindakan + Cell11)
                End With
            ElseIf dbRst!TindakanMedis = "Ortotik Prostetik Lain" Then
                With oSheet
                    .Cells(j, 11) = Trim(dbRst!JmlTindakan + Cell11)
                End With
            End If

            With oSheet
                .Cells(34, 11) = Trim(dbRst!KunjunganRumah + Cell11)
            End With

            dbRst.MoveNext
        Wend
    End If

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    MsgBox "Data Tidak Ada", vbInformation, "Validasi"
    Screen.MousePointer = vbDefault
End Sub
