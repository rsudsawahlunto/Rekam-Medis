VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRL1HAL4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL1 Halaman 4"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRL1HAL4.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6345
   Begin VB.Frame Frame3 
      Caption         =   "Triwulan"
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   5655
      Begin VB.CheckBox Check1 
         Caption         =   "Triwulan"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Triwulan2"
         Height          =   495
         Left            =   1560
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Triwulan3"
         Height          =   495
         Left            =   2880
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Triwulan4"
         Height          =   495
         Left            =   4200
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Triwulan1"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtptahun 
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   360
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
         Format          =   117178371
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
   End
   Begin VB.Frame Frame1 
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
      Left            =   3480
      TabIndex        =   5
      Top             =   3120
      Width           =   2295
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
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
         Format          =   117178371
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
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
      Top             =   3120
      Width           =   2295
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
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
         Format          =   117178371
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
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   6285
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
      TabIndex        =   14
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
      Caption         =   "Periode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   15
      Top             =   1080
      Width           =   6255
      Begin VB.Label Label1 
         Caption         =   "s/d"
         Height          =   255
         Left            =   2760
         TabIndex        =   16
         Top             =   2280
         Width           =   375
      End
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   3360
      Picture         =   "frmRL1HAL4.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2955
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRL1HAL4.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL1HAL4.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmRL1HAL4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project/reference/microsoft excel 12.0 object library
'Selalu gunakan format file excel 2003  .xls sebagai standar agar pengguna excel 2003 atau diatasnya dpt menggunakan report laporannya
'Catatan: Format excel 2000 tidak dpt mengoperasikan beberapa fungsi yg ada pada excell 2003 atau diatasnya

Option Explicit

Dim awal As String
Dim akhir As String

'Special Buat Excel
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel.Range
Dim oResizeRange As Excel.Range
Dim j As String
'Special Buat Excel

Dim Cell3 As String
Dim Cell4 As String
Dim Cell5 As String
Dim Cell6 As String
Dim Cell7 As String
Dim Cell8 As String
Dim Cell9 As String
Dim Cell10 As String
Dim Cell11 As String
Dim Cell12 As String
Dim Cell13 As String
Dim Cell14 As String
Dim Cell15 As String
Dim Cell16 As String

Private Sub Check1_Click()
    If Check1.value = 0 Then
        dtpAwal.Enabled = True
        dtpAkhir.Enabled = True
        dtptahun.Enabled = False
        Option1.Enabled = False
        Option2.Enabled = False
        Option3.Enabled = False
        Option4.Enabled = False
        dtpAwal.value = Now
        dtpAkhir.value = Now
        dtpAkhir.CustomFormat = "dd MMMM yyyy"
        dtpAwal.CustomFormat = "dd MMMM yyyy"
    Else
        dtpAwal.Enabled = False
        dtpAkhir.Enabled = False
        dtptahun.Enabled = True
        Option1.Enabled = True
        Option2.Enabled = True
        Option3.Enabled = True
        Option4.Enabled = True
        dtpAkhir.CustomFormat = "MMMM dd"
        dtpAwal.CustomFormat = "MMMM dd"
        dtptahun.value = Now
    End If
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo errLoad

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    oXL.Visible = True
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.path & "\RL1 Hal4.xls")
    Set oSheet = oWB.ActiveSheet

    Set rsb = Nothing
    strSQL = "select * from profilrs"
    Call msubRecFO(rsb, strSQL)

    Set oResizeRange = oSheet.Range("d1", "d2")
    oResizeRange.value = Trim(rsb!KdRs)

    strSQL = "select * from RL1_14 where TglPeriksa between '" & Format(dtpAwal.value, "yyyy/MM/dd") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd") & "'or tglperiksa is null"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF
            If dbRst!kdjeniskontrasepsi = "01" Then
                j = 31
            ElseIf dbRst!kdjeniskontrasepsi = "02" Then
                j = 32
            ElseIf dbRst!kdjeniskontrasepsi = "03" Then
                j = 33
            ElseIf dbRst!kdjeniskontrasepsi = "04" Then
                j = 34
            ElseIf dbRst!kdjeniskontrasepsi = "05" Then
                j = 35
            ElseIf dbRst!kdjeniskontrasepsi = "06" Then
                j = 36
            ElseIf dbRst!kdjeniskontrasepsi = "07" Then
                j = 37
            ElseIf dbRst!kdjeniskontrasepsi = "08" Then
                j = 38
            End If

            Cell3 = oSheet.Cells(j, 3).value
            Cell4 = oSheet.Cells(j, 4).value
            Cell5 = oSheet.Cells(j, 5).value
            Cell7 = oSheet.Cells(j, 7).value
            Cell8 = oSheet.Cells(j, 8).value
            Cell9 = oSheet.Cells(j, 9).value

            If dbRst!kdjeniskontrasepsi = "01" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!bukanrujukan + Cell3)
                    .Cells(j, 4) = Trim(dbRst!rujukanri + Cell4)
                    .Cells(j, 5) = Trim(dbRst!rujukanrj + Cell5)
                    .Cells(j, 7) = Trim(dbRst!kunjunganulang + Cell7)
                    .Cells(j, 8) = Trim(dbRst!jmlefek + Cell8)
                    .Cells(j, 9) = Trim(dbRst!dirujukkeatas + Cell9)
                End With

            ElseIf dbRst!kdjeniskontrasepsi = "02" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!bukanrujukan + Cell3)
                    .Cells(j, 4) = Trim(dbRst!rujukanri + Cell4)
                    .Cells(j, 5) = Trim(dbRst!rujukanrj + Cell5)
                    .Cells(j, 7) = Trim(dbRst!kunjunganulang + Cell7)
                    .Cells(j, 8) = Trim(dbRst!jmlefek + Cell8)
                    .Cells(j, 9) = Trim(dbRst!dirujukkeatas + Cell9)
                End With

            ElseIf dbRst!kdjeniskontrasepsi = "03" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!bukanrujukan + Cell3)
                    .Cells(j, 4) = Trim(dbRst!rujukanri + Cell4)
                    .Cells(j, 5) = Trim(dbRst!rujukanrj + Cell5)
                    .Cells(j, 7) = Trim(dbRst!kunjunganulang + Cell7)
                    .Cells(j, 8) = Trim(dbRst!jmlefek + Cell8)
                    .Cells(j, 9) = Trim(dbRst!dirujukkeatas + Cell9)
                End With

            ElseIf dbRst!kdjeniskontrasepsi = "04" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!bukanrujukan + Cell3)
                    .Cells(j, 4) = Trim(dbRst!rujukanri + Cell4)
                    .Cells(j, 5) = Trim(dbRst!rujukanrj + Cell5)
                    .Cells(j, 7) = Trim(dbRst!kunjunganulang + Cell7)
                    .Cells(j, 8) = Trim(dbRst!jmlefek + Cell8)
                    .Cells(j, 9) = Trim(dbRst!dirujukkeatas + Cell9)
                End With

            ElseIf dbRst!kdjeniskontrasepsi = "05" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!bukanrujukan + Cell3)
                    .Cells(j, 4) = Trim(dbRst!rujukanri + Cell4)
                    .Cells(j, 5) = Trim(dbRst!rujukanrj + Cell5)
                    .Cells(j, 7) = Trim(dbRst!kunjunganulang + Cell7)
                    .Cells(j, 8) = Trim(dbRst!jmlefek + Cell8)
                    .Cells(j, 9) = Trim(dbRst!dirujukkeatas + Cell9)
                End With

            ElseIf dbRst!kdjeniskontrasepsi = "06" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!bukanrujukan + Cell3)
                    .Cells(j, 4) = Trim(dbRst!rujukanri + Cell4)
                    .Cells(j, 5) = Trim(dbRst!rujukanrj + Cell5)
                    .Cells(j, 7) = Trim(dbRst!kunjunganulang + Cell7)
                    .Cells(j, 8) = Trim(dbRst!jmlefek + Cell8)
                    .Cells(j, 9) = Trim(dbRst!dirujukkeatas + Cell9)
                End With

            ElseIf dbRst!kdjeniskontrasepsi = "07" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!bukanrujukan + Cell3)
                    .Cells(j, 4) = Trim(dbRst!rujukanri + Cell4)
                    .Cells(j, 5) = Trim(dbRst!rujukanrj + Cell5)
                    .Cells(j, 7) = Trim(dbRst!kunjunganulang + Cell7)
                    .Cells(j, 8) = Trim(dbRst!jmlefek + Cell8)
                    .Cells(j, 9) = Trim(dbRst!dirujukkeatas + Cell9)
                End With

            ElseIf dbRst!kdjeniskontrasepsi = "08" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!bukanrujukan + Cell3)
                    .Cells(j, 4) = Trim(dbRst!rujukanri + Cell4)
                    .Cells(j, 5) = Trim(dbRst!rujukanrj + Cell5)
                    .Cells(j, 7) = Trim(dbRst!kunjunganulang + Cell7)
                    .Cells(j, 8) = Trim(dbRst!jmlefek + Cell8)
                    .Cells(j, 9) = Trim(dbRst!dirujukkeatas + Cell9)
                End With
            End If
            dbRst.MoveNext
        Wend
    End If

    strSQL = "select * from RL1_16"
    Call msubRecFO(rsx, strSQL)

    If rsx.RecordCount > 0 Then
        rsx.MoveFirst
        j = 21
        While Not rsx.EOF
            With oSheet
                .Cells(j, 13) = Trim(IIf(IsNull(rsx!Jml.value), 0, (rsx!Jml.value)))
            End With
            j = j + 1
            rsx.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select distinct * from RL1_13 "
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF
            If dbRst!TindakanMedis = "Gait Analyzer" Then
                j = 7
            ElseIf dbRst!TindakanMedis = "EMG" Then
                j = 8
            ElseIf dbRst!TindakanMedis = "Uro Dinamic" Then
                j = 9
            ElseIf dbRst!TindakanMedis = "Sideback" Then
                j = 10
            ElseIf dbRst!TindakanMedis = "E N Tree" Then
                j = 11
            ElseIf dbRst!TindakanMedis = "Spyrometer" Then
                j = 12
            ElseIf dbRst!TindakanMedis = "Static Bicycle" Then
                j = 13
            ElseIf dbRst!TindakanMedis = "Tread Mill" Then
                j = 14
            ElseIf dbRst!TindakanMedis = "Body Platismograf" Then
                j = 15

            ElseIf dbRst!TindakanMedis = "Latihan Fisik" Then
                j = 17
            ElseIf dbRst!TindakanMedis = "Aktinoterapi" Then
                j = 18
            ElseIf dbRst!TindakanMedis = "Elektroterapi" Then
                j = 19
            ElseIf dbRst!TindakanMedis = "Hidroterapi" Then
                j = 20
            ElseIf dbRst!TindakanMedis = "Traksi Lumbal & Cervical" Then
                j = 21
            ElseIf dbRst!TindakanMedis = "Fisioterapi Lain" Then
                j = 22

            ElseIf dbRst!TindakanMedis = "Snoosien Room" Then
                j = 24
            ElseIf dbRst!TindakanMedis = "Sensori Integrasi" Then
                j = 25
            ElseIf dbRst!TindakanMedis = "Latihan Aktivitas Kehidupan Sehari-hari" Then
                j = 26

            ElseIf dbRst!TindakanMedis = "Proper Body Mekanik" Then
                j = 6
            ElseIf dbRst!TindakanMedis = "Pembuatan Alat lontar & Adaptasi Alat" Then
                j = 7
            ElseIf dbRst!TindakanMedis = "Analisa Persiapan Kerja" Then
                j = 8
            ElseIf dbRst!TindakanMedis = "Latihan Relaksasi" Then
                j = 9
            ElseIf dbRst!TindakanMedis = "Analisa & Intervensi, Persepsi, Kognitif" Then
                j = 10

            ElseIf dbRst!TindakanMedis = "Fungsi Bicara" Then
                j = 12
            ElseIf dbRst!TindakanMedis = "Fungsi Bahasa & Laku" Then
                j = 13
            ElseIf dbRst!TindakanMedis = "Fungsi Menelan" Then
                j = 14

            ElseIf dbRst!TindakanMedis = "Psikologi Anak" Then
                j = 16
            ElseIf dbRst!TindakanMedis = "Psikologi Dewasa" Then
                j = 17

            ElseIf dbRst!TindakanMedis = "Evaluasi Lingkungan Rumah" Then
                j = 19
            ElseIf dbRst!TindakanMedis = "Evaluasi Ekonomi" Then
                j = 20
            ElseIf dbRst!TindakanMedis = "Evaluasi Pekerjaan" Then
                j = 21

            ElseIf dbRst!TindakanMedis = "Pembuatan Alat Bantu" Then
                j = 23
            ElseIf dbRst!TindakanMedis = "Pembuatan Alat Anggota Tiruan" Then
                j = 24
            ElseIf dbRst!TindakanMedis = "Ortotik Prostetik Lain" Then
                j = 25
            End If

            If oSheet.Cells(j, 3).value = "" Then Cell3 = 0 Else Cell3 = oSheet.Cells(j, 3).value
            If oSheet.Cells(j, 6).value = "" Then Cell6 = 0 Else Cell6 = oSheet.Cells(j, 6).value

            If dbRst!TindakanMedis = "Gait Analyzer" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With
            ElseIf dbRst!TindakanMedis = "EMG" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With
            ElseIf dbRst!TindakanMedis = "Uro Dinamic" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With
            ElseIf dbRst!TindakanMedis = "Sideback" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With
            ElseIf dbRst!TindakanMedis = "E N Tree" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With
            ElseIf dbRst!TindakanMedis = "Spyrometer" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With
            ElseIf dbRst!TindakanMedis = "Static Bicycle" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With
            ElseIf dbRst!TindakanMedis = "Tread Mill" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With
            ElseIf dbRst!TindakanMedis = "Body Platismograf" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With

            ElseIf dbRst!TindakanMedis = "Latihan Fisik" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With
            ElseIf dbRst!TindakanMedis = "Aktinoterapi" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With
            ElseIf dbRst!TindakanMedis = "Elektroterapi" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With
            ElseIf dbRst!TindakanMedis = "Hidroterapi" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With
            ElseIf dbRst!TindakanMedis = "Traksi Lumbal & Cervical" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With
            ElseIf dbRst!TindakanMedis = "Fisioterapi Lain" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With

            ElseIf dbRst!TindakanMedis = "Snoosien Room" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With
            ElseIf dbRst!TindakanMedis = "Sensori Integrasi" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
                End With
            ElseIf dbRst!TindakanMedis = "Latihan Aktivitas Kehidupan Sehari-hari" Then
                With oSheet
                    .Cells(j, 3) = Trim(dbRst!JmlTindakan + Cell3)
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
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Latihan Relaksasi" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Analisa & Intervensi, Persepsi, Kognitif" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With

            ElseIf dbRst!TindakanMedis = "Fungsi Bicara" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Fungsi Bahasa & Laku" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Fungsi Menelan" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With

            ElseIf dbRst!TindakanMedis = "Psikologi Anak" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Psikologi Dewasa" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With

            ElseIf dbRst!TindakanMedis = "Evaluasi Lingkungan Rumah" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Evaluasi Ekonomi" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Evaluasi Pekerjaan" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With

            ElseIf dbRst!TindakanMedis = "Pembuatan Alat Bantu" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Pembuatan Alat Anggota Tiruan" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            ElseIf dbRst!TindakanMedis = "Ortotik Prostetik Lain" Then
                With oSheet
                    .Cells(j, 6) = Trim(dbRst!JmlTindakan + Cell6)
                End With
            End If

            With oSheet
                .Cells(26, 6) = Trim(dbRst!KunjunganRumah + Cell6)
            End With
            dbRst.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select distinct * from RL1_15 "
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF
            If dbRst!JenisPenyuluhan = "Kesehatan Umum" Then
                j = 7
            ElseIf dbRst!JenisPenyuluhan = "Keluarga Berencana" Then
                j = 8
            ElseIf dbRst!JenisPenyuluhan = "Kesehatan Ibu & Anak" Then
                j = 9
            ElseIf dbRst!JenisPenyuluhan = "Gizi" Then
                j = 10
            ElseIf dbRst!JenisPenyuluhan = "Imunisasi" Then
                j = 11
            ElseIf dbRst!JenisPenyuluhan = "Usia Lanjut" Then
                j = 12
            ElseIf dbRst!JenisPenyuluhan = "Penyakit Diare" Then
                j = 13
            ElseIf dbRst!JenisPenyuluhan = "Gigi & Mulut" Then
                j = 14
            ElseIf dbRst!JenisPenyuluhan = "Kesehatan Jiwa" Then
                j = 15
            ElseIf dbRst!JenisPenyuluhan = "NAPZA" Then
                j = 16
            ElseIf dbRst!JenisPenyuluhan = "Lain-Lain" Then
                j = 17
            End If

            If oSheet.Cells(j, 10).value = "" Then Cell10 = 0 Else Cell10 = oSheet.Cells(j, 10).value
            If oSheet.Cells(j, 11).value = "" Then Cell11 = 0 Else Cell11 = oSheet.Cells(j, 11).value
            If oSheet.Cells(j, 12).value = "" Then Cell12 = 0 Else Cell12 = oSheet.Cells(j, 12).value
            If oSheet.Cells(j, 13).value = "" Then Cell13 = 0 Else Cell13 = oSheet.Cells(j, 13).value
            If oSheet.Cells(j, 14).value = "" Then Cell14 = 0 Else Cell14 = oSheet.Cells(j, 14).value
            If oSheet.Cells(j, 15).value = "" Then Cell15 = 0 Else Cell15 = oSheet.Cells(j, 15).value
            If oSheet.Cells(j, 16).value = "" Then Cell16 = 0 Else Cell16 = oSheet.Cells(j, 16).value

            If dbRst!JenisPenyuluhan = "Kesehatan Umum" Then
                With oSheet
                    If dbRst!CaraTindakanP = "Pemasangan Poster" Then
                        .Cells(j, 10) = Trim(dbRst!Jml + Cell10)
                    ElseIf dbRst!CaraTindakanP = "Pemutaran Kaset" Then
                        .Cells(j, 11) = Trim(dbRst!Jml + Cell11)
                    ElseIf dbRst!CaraTindakanP = "Demonstrasi" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!CaraTindakanP = "Pameran" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!CaraTindakanP = "Pelatihan" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    ElseIf dbRst!CaraTindakanP = "Donasi" Then
                        .Cells(j, 15) = Trim(dbRst!Jml + Cell15)
                    ElseIf dbRst!CaraTindakanP = "Konsignasi" Then
                        .Cells(j, 16) = Trim(dbRst!Jml + Cell16)
                    End If
                End With
            ElseIf dbRst!JenisPenyuluhan = "Keluarga Berencana" Then
                With oSheet
                    If dbRst!CaraTindakanP = "Pemasangan Poster" Then
                        .Cells(j, 10) = Trim(dbRst!Jml + Cell10)
                    ElseIf dbRst!CaraTindakanP = "Pemutaran Kaset" Then
                        .Cells(j, 11) = Trim(dbRst!Jml + Cell11)
                    ElseIf dbRst!CaraTindakanP = "Demonstrasi" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!CaraTindakanP = "Pameran" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!CaraTindakanP = "Pelatihan" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    ElseIf dbRst!CaraTindakanP = "Donasi" Then
                        .Cells(j, 15) = Trim(dbRst!Jml + Cell15)
                    ElseIf dbRst!CaraTindakanP = "Konsignasi" Then
                        .Cells(j, 16) = Trim(dbRst!Jml + Cell16)
                    End If
                End With
            ElseIf dbRst!JenisPenyuluhan = "Kesehatan Ibu & Anak" Then
                With oSheet
                    If dbRst!CaraTindakanP = "Pemasangan Poster" Then
                        .Cells(j, 10) = Trim(dbRst!Jml + Cell10)
                    ElseIf dbRst!CaraTindakanP = "Pemutaran Kaset" Then
                        .Cells(j, 11) = Trim(dbRst!Jml + Cell11)
                    ElseIf dbRst!CaraTindakanP = "Demonstrasi" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!CaraTindakanP = "Pameran" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!CaraTindakanP = "Pelatihan" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    ElseIf dbRst!CaraTindakanP = "Donasi" Then
                        .Cells(j, 15) = Trim(dbRst!Jml + Cell15)
                    ElseIf dbRst!CaraTindakanP = "Konsignasi" Then
                        .Cells(j, 16) = Trim(dbRst!Jml + Cell16)
                    End If
                End With
            ElseIf dbRst!JenisPenyuluhan = "Gizi" Then
                With oSheet
                    If dbRst!CaraTindakanP = "Pemasangan Poster" Then
                        .Cells(j, 10) = Trim(dbRst!Jml + Cell10)
                    ElseIf dbRst!CaraTindakanP = "Pemutaran Kaset" Then
                        .Cells(j, 11) = Trim(dbRst!Jml + Cell11)
                    ElseIf dbRst!CaraTindakanP = "Demonstrasi" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!CaraTindakanP = "Pameran" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!CaraTindakanP = "Pelatihan" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    ElseIf dbRst!CaraTindakanP = "Donasi" Then
                        .Cells(j, 15) = Trim(dbRst!Jml + Cell15)
                    ElseIf dbRst!CaraTindakanP = "Konsignasi" Then
                        .Cells(j, 16) = Trim(dbRst!Jml + Cell16)
                    End If
                End With
            ElseIf dbRst!JenisPenyuluhan = "Imunisasi" Then
                With oSheet
                    If dbRst!CaraTindakanP = "Pemasangan Poster" Then
                        .Cells(j, 10) = Trim(dbRst!Jml + Cell10)
                    ElseIf dbRst!CaraTindakanP = "Pemutaran Kaset" Then
                        .Cells(j, 11) = Trim(dbRst!Jml + Cell11)
                    ElseIf dbRst!CaraTindakanP = "Demonstrasi" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!CaraTindakanP = "Pameran" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!CaraTindakanP = "Pelatihan" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    ElseIf dbRst!CaraTindakanP = "Donasi" Then
                        .Cells(j, 15) = Trim(dbRst!Jml + Cell15)
                    ElseIf dbRst!CaraTindakanP = "Konsignasi" Then
                        .Cells(j, 16) = Trim(dbRst!Jml + Cell16)
                    End If
                End With
            ElseIf dbRst!JenisPenyuluhan = "Usia Lanjut" Then
                With oSheet
                    If dbRst!CaraTindakanP = "Pemasangan Poster" Then
                        .Cells(j, 10) = Trim(dbRst!Jml + Cell10)
                    ElseIf dbRst!CaraTindakanP = "Pemutaran Kaset" Then
                        .Cells(j, 11) = Trim(dbRst!Jml + Cell11)
                    ElseIf dbRst!CaraTindakanP = "Demonstrasi" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!CaraTindakanP = "Pameran" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!CaraTindakanP = "Pelatihan" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    ElseIf dbRst!CaraTindakanP = "Donasi" Then
                        .Cells(j, 15) = Trim(dbRst!Jml + Cell15)
                    ElseIf dbRst!CaraTindakanP = "Konsignasi" Then
                        .Cells(j, 16) = Trim(dbRst!Jml + Cell16)
                    End If
                End With
            ElseIf dbRst!JenisPenyuluhan = "Penyakit Diare" Then
                With oSheet
                    If dbRst!CaraTindakanP = "Pemasangan Poster" Then
                        .Cells(j, 10) = Trim(dbRst!Jml + Cell10)
                    ElseIf dbRst!CaraTindakanP = "Pemutaran Kaset" Then
                        .Cells(j, 11) = Trim(dbRst!Jml + Cell11)
                    ElseIf dbRst!CaraTindakanP = "Demonstrasi" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!CaraTindakanP = "Pameran" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!CaraTindakanP = "Pelatihan" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    ElseIf dbRst!CaraTindakanP = "Donasi" Then
                        .Cells(j, 15) = Trim(dbRst!Jml + Cell15)
                    ElseIf dbRst!CaraTindakanP = "Konsignasi" Then
                        .Cells(j, 16) = Trim(dbRst!Jml + Cell16)
                    End If
                End With
            ElseIf dbRst!JenisPenyuluhan = "Gigi & Mulut" Then
                With oSheet
                    If dbRst!CaraTindakanP = "Pemasangan Poster" Then
                        .Cells(j, 10) = Trim(dbRst!Jml + Cell10)
                    ElseIf dbRst!CaraTindakanP = "Pemutaran Kaset" Then
                        .Cells(j, 11) = Trim(dbRst!Jml + Cell11)
                    ElseIf dbRst!CaraTindakanP = "Demonstrasi" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!CaraTindakanP = "Pameran" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!CaraTindakanP = "Pelatihan" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    ElseIf dbRst!CaraTindakanP = "Donasi" Then
                        .Cells(j, 15) = Trim(dbRst!Jml + Cell15)
                    ElseIf dbRst!CaraTindakanP = "Konsignasi" Then
                        .Cells(j, 16) = Trim(dbRst!Jml + Cell16)
                    End If
                End With
            ElseIf dbRst!JenisPenyuluhan = "Kesehatan Jiwa" Then
                With oSheet
                    If dbRst!CaraTindakanP = "Pemasangan Poster" Then
                        .Cells(j, 10) = Trim(dbRst!Jml + Cell10)
                    ElseIf dbRst!CaraTindakanP = "Pemutaran Kaset" Then
                        .Cells(j, 11) = Trim(dbRst!Jml + Cell11)
                    ElseIf dbRst!CaraTindakanP = "Demonstrasi" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!CaraTindakanP = "Pameran" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!CaraTindakanP = "Pelatihan" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    ElseIf dbRst!CaraTindakanP = "Donasi" Then
                        .Cells(j, 15) = Trim(dbRst!Jml + Cell15)
                    ElseIf dbRst!CaraTindakanP = "Konsignasi" Then
                        .Cells(j, 16) = Trim(dbRst!Jml + Cell16)
                    End If
                End With
            ElseIf dbRst!JenisPenyuluhan = "NAPZA" Then
                With oSheet
                    If dbRst!CaraTindakanP = "Pemasangan Poster" Then
                        .Cells(j, 10) = Trim(dbRst!Jml + Cell10)
                    ElseIf dbRst!CaraTindakanP = "Pemutaran Kaset" Then
                        .Cells(j, 11) = Trim(dbRst!Jml + Cell11)
                    ElseIf dbRst!CaraTindakanP = "Demonstrasi" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!CaraTindakanP = "Pameran" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!CaraTindakanP = "Pelatihan" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    ElseIf dbRst!CaraTindakanP = "Donasi" Then
                        .Cells(j, 15) = Trim(dbRst!Jml + Cell15)
                    ElseIf dbRst!CaraTindakanP = "Konsignasi" Then
                        .Cells(j, 16) = Trim(dbRst!Jml + Cell16)
                    End If
                End With
            ElseIf dbRst!JenisPenyuluhan = "Lain-lain" Then
                With oSheet
                    If dbRst!CaraTindakanP = "Pemasangan Poster" Then
                        .Cells(j, 10) = Trim(dbRst!Jml + Cell10)
                    ElseIf dbRst!CaraTindakanP = "Pemutaran Kaset" Then
                        .Cells(j, 11) = Trim(dbRst!Jml + Cell11)
                    ElseIf dbRst!CaraTindakanP = "Demonstrasi" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!CaraTindakanP = "Pameran" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!CaraTindakanP = "Pelatihan" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    ElseIf dbRst!CaraTindakanP = "Donasi" Then
                        .Cells(j, 15) = Trim(dbRst!Jml + Cell15)
                    ElseIf dbRst!CaraTindakanP = "Konsignasi" Then
                        .Cells(j, 16) = Trim(dbRst!Jml + Cell16)
                    End If
                End With
            End If
            dbRst.MoveNext
        Wend
    End If

    strSQL = "select * from RL1_17"
    Call msubRecFO(rsx, strSQL)

    If rsx.RecordCount > 0 Then
        rsx.MoveFirst
        j = 38
        While Not rsx.EOF
            With oSheet
                .Cells(j, 12) = Trim(IIf(IsNull(rsx!JenisKeahlian.value), "", (rsx!JenisKeahlian.value)))
                .Cells(j, 13) = Trim(IIf(IsNull(rsx!Negara.value), "", (rsx!Negara.value)))
                .Cells(j, 14) = Trim(IIf(IsNull(rsx!Status.value), "", (rsx!Status.value)))
                .Cells(j, 15) = Trim(IIf(IsNull(rsx!LamaDomisili.value), "", (rsx!LamaDomisili.value)))
                .Cells(j, 16) = Trim(IIf(IsNull(rsx!JenisPelayanan.value), "", (rsx!JenisPelayanan.value)))
                .Cells(j, 17) = Trim(IIf(IsNull(rsx!Jml.value), 0, (rsx!Jml.value)))
            End With
            j = j + 1
            rsx.MoveNext
        Wend
    End If
    Exit Sub
errLoad:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtptahun_Change()
    dtptahun.MaxDate = Now
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    With Me
        .dtpAwal.value = Now
        .dtpAkhir.value = Now
        .dtptahun.value = Now
    End With
    Check1.value = 1
    Option1.value = 1
End Sub

Private Sub Option1_Click()
    awal = CStr(dtptahun.Year) + "/01/01"
    akhir = CStr(dtptahun.Year) + "/03/31"

    dtpAwal = awal
    dtpAkhir = akhir
End Sub

Private Sub Option2_Click()
    awal = CStr(dtptahun.Year) + "/04/01"
    akhir = CStr(dtptahun.Year) + "/06/30"

    dtpAwal = awal
    dtpAkhir = akhir
End Sub

Private Sub Option3_Click()
    awal = CStr(dtptahun.Year) + "/07/01"
    akhir = CStr(dtptahun.Year) + "/09/30"

    dtpAwal = awal
    dtpAkhir = akhir
End Sub

Private Sub Option4_Click()
    awal = CStr(dtptahun.Year) + "/10/01"
    akhir = CStr(dtptahun.Year) + "/12/31"

    dtpAwal = awal
    dtpAkhir = akhir
End Sub
