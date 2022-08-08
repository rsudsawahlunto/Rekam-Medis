VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRL1HAL6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL1 Halaman 6"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRL1HAL6.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6300
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
         Format          =   115408899
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
         Format          =   115408899
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
         Format          =   115408899
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
      Picture         =   "frmRL1HAL6.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2955
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRL1HAL6.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL1HAL6.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmRL1HAL6"
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

'Untuk Pengganti Group Dijadikan Penginputan Di Cell yg sama
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
Dim Cell17 As String
Dim Cell18 As String
Dim Cell19 As String
Dim Cell20 As String
Dim Cell21 As String
Dim Cell22 As String
'Untuk Pengganti Group Dijadikan Penginputan Di Cell yg sama

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
    Set oWB = oXL.Workbooks.Open(App.Path & "\RL1 Hal6.xls")
    Set oSheet = oWB.ActiveSheet

    Set rsb = Nothing
    strSQL = "select * from profilrs"
    Call msubRecFO(rsb, strSQL)

    Set oResizeRange = oSheet.Range("h1", "h2")
    oResizeRange.value = Trim(rsb!KdRs)

    Set rs = Nothing
    strSQL = " select * from RL1_23 where TglMasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'or TglMasuk is null"
    Call msubRecFO(rs, strSQL)

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        While Not rs.EOF
            If rs!NamaExternal = "Membayar" Then
                j = 7
            ElseIf rs!NamaExternal = "Askes" Then
                j = 9
            ElseIf rs!NamaExternal = "JPKM" Then
                j = 11
            ElseIf rs!NamaExternal = "Kontrak" Then
                j = 12
            ElseIf rs!NamaExternal = "Keringanan" Then
                j = 13
            ElseIf rs!NamaExternal = "Kartu Sehat" Then
                j = 15
            ElseIf rs!NamaExternal = "Keterangan Tidak Mampu" Then
                j = 16
            End If

            Cell7 = oSheet.Cells(j, 7).value
            Cell8 = oSheet.Cells(j, 8).value
            Cell9 = oSheet.Cells(j, 9).value
            Cell10 = oSheet.Cells(j, 10).value
            Cell11 = oSheet.Cells(j, 11).value
            Cell12 = oSheet.Cells(j, 12).value
            Cell13 = oSheet.Cells(j, 13).value
            Cell14 = oSheet.Cells(j, 14).value

            If rs!NamaExternal = "Membayar" Then
                With oSheet
                    .Cells(j, 7) = Trim(IIf(IsNull(rs!jmlpasienkeluar.value), 0, (rs!jmlpasienkeluar.value)) + Cell7)
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!lamadirawat.value), 0, (rs!lamadirawat.value)) + Cell8)
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!jmlpasienrj.value), 0, (rs!jmlpasienrj.value)) + Cell9)
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpasienlab.value), 0, (rs!jmlpasienlab.value)) + Cell10)
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlpasienrad.value), 0, (rs!jmlpasienrad.value)) + Cell11)
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmllainnya.value), 0, (rs!jmllainnya.value)) + Cell12)
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!seharusnya.value), 0, (rs!seharusnya.value)) + Cell13)
                    .Cells(j, 14) = Trim(IIf(IsNull(rs!diterima.value), 0, (rs!diterima.value)) + Cell14)
                End With
            ElseIf rs!NamaExternal = "Askes" Then
                With oSheet
                    .Cells(j, 7) = Trim(IIf(IsNull(rs!jmlpasienkeluar.value), 0, (rs!jmlpasienkeluar.value)) + Cell7)
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!lamadirawat.value), 0, (rs!lamadirawat.value)) + Cell8)
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!jmlpasienrj.value), 0, (rs!jmlpasienrj.value)) + Cell9)
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpasienlab.value), 0, (rs!jmlpasienlab.value)) + Cell10)
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlpasienrad.value), 0, (rs!jmlpasienrad.value)) + Cell11)
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmllainnya.value), 0, (rs!jmllainnya.value)) + Cell12)
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!seharusnya.value), 0, (rs!seharusnya.value)) + Cell13)
                    .Cells(j, 14) = Trim(IIf(IsNull(rs!diterima.value), 0, (rs!diterima.value)) + Cell14)
                End With
            ElseIf rs!NamaExternal = "JPKM" Then
                With oSheet
                    .Cells(j, 7) = Trim(IIf(IsNull(rs!jmlpasienkeluar.value), 0, (rs!jmlpasienkeluar.value)) + Cell7)
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!lamadirawat.value), 0, (rs!lamadirawat.value)) + Cell8)
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!jmlpasienrj.value), 0, (rs!jmlpasienrj.value)) + Cell9)
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpasienlab.value), 0, (rs!jmlpasienlab.value)) + Cell10)
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlpasienrad.value), 0, (rs!jmlpasienrad.value)) + Cell11)
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmllainnya.value), 0, (rs!jmllainnya.value)) + Cell12)
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!seharusnya.value), 0, (rs!seharusnya.value)) + Cell13)
                    .Cells(j, 14) = Trim(IIf(IsNull(rs!diterima.value), 0, (rs!diterima.value)) + Cell14)
                End With
            ElseIf rs!NamaExternal = "Kontrak" Then
                With oSheet
                    .Cells(j, 7) = Trim(IIf(IsNull(rs!jmlpasienkeluar.value), 0, (rs!jmlpasienkeluar.value)) + Cell7)
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!lamadirawat.value), 0, (rs!lamadirawat.value)) + Cell8)
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!jmlpasienrj.value), 0, (rs!jmlpasienrj.value)) + Cell9)
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpasienlab.value), 0, (rs!jmlpasienlab.value)) + Cell10)
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlpasienrad.value), 0, (rs!jmlpasienrad.value)) + Cell11)
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmllainnya.value), 0, (rs!jmllainnya.value)) + Cell12)
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!seharusnya.value), 0, (rs!seharusnya.value)) + Cell13)
                    .Cells(j, 14) = Trim(IIf(IsNull(rs!diterima.value), 0, (rs!diterima.value)) + Cell14)
                End With
            ElseIf rs!NamaExternal = "Keringanan" Then
                With oSheet
                    .Cells(j, 7) = Trim(IIf(IsNull(rs!jmlpasienkeluar.value), 0, (rs!jmlpasienkeluar.value)) + Cell7)
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!lamadirawat.value), 0, (rs!lamadirawat.value)) + Cell8)
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!jmlpasienrj.value), 0, (rs!jmlpasienrj.value)) + Cell9)
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpasienlab.value), 0, (rs!jmlpasienlab.value)) + Cell10)
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlpasienrad.value), 0, (rs!jmlpasienrad.value)) + Cell11)
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmllainnya.value), 0, (rs!jmllainnya.value)) + Cell12)
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!seharusnya.value), 0, (rs!seharusnya.value)) + Cell13)
                    .Cells(j, 14) = Trim(IIf(IsNull(rs!diterima.value), 0, (rs!diterima.value)) + Cell14)
                End With
            ElseIf rs!NamaExternal = "Kartu Sehat" Then
                With oSheet
                    .Cells(j, 7) = Trim(IIf(IsNull(rs!jmlpasienkeluar.value), 0, (rs!jmlpasienkeluar.value)) + Cell7)
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!lamadirawat.value), 0, (rs!lamadirawat.value)) + Cell8)
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!jmlpasienrj.value), 0, (rs!jmlpasienrj.value)) + Cell9)
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpasienlab.value), 0, (rs!jmlpasienlab.value)) + Cell10)
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlpasienrad.value), 0, (rs!jmlpasienrad.value)) + Cell11)
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmllainnya.value), 0, (rs!jmllainnya.value)) + Cell12)
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!seharusnya.value), 0, (rs!seharusnya.value)) + Cell13)
                    .Cells(j, 14) = Trim(IIf(IsNull(rs!diterima.value), 0, (rs!diterima.value)) + Cell14)
                End With
            ElseIf rs!NamaExternal = "Keterangan Tidak Mampu" Then
                With oSheet
                    .Cells(j, 7) = Trim(IIf(IsNull(rs!jmlpasienkeluar.value), 0, (rs!jmlpasienkeluar.value)) + Cell7)
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!lamadirawat.value), 0, (rs!lamadirawat.value)) + Cell8)
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!jmlpasienrj.value), 0, (rs!jmlpasienrj.value)) + Cell9)
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpasienlab.value), 0, (rs!jmlpasienlab.value)) + Cell10)
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlpasienrad.value), 0, (rs!jmlpasienrad.value)) + Cell11)
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmllainnya.value), 0, (rs!jmllainnya.value)) + Cell12)
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!seharusnya.value), 0, (rs!seharusnya.value)) + Cell13)
                    .Cells(j, 14) = Trim(IIf(IsNull(rs!diterima.value), 0, (rs!diterima.value)) + Cell14)
                End With
            End If
            rs.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select distinct * from RL1_24_1 where Tglkirim between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'or tglkirim is null"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF
            If dbRst!kdsubinstalasi = "001" Then
                j = 25
            ElseIf dbRst!kdsubinstalasi = "002" Then
                j = 26
            ElseIf dbRst!kdsubinstalasi = "003" Then
                j = 27
            ElseIf dbRst!kdsubinstalasi = "005" Then
                j = 28
            ElseIf dbRst!kdsubinstalasi = "004" Then
                j = 29
            ElseIf dbRst!kdsubinstalasi = "007" Then
                j = 30
            ElseIf dbRst!kdsubinstalasi = "008" Then
                j = 31
            ElseIf dbRst!kdsubinstalasi = "009" Then
                j = 32
            ElseIf dbRst!kdsubinstalasi = "010" Then
                j = 33
            ElseIf dbRst!kdsubinstalasi = "011" Then
                j = 34
            ElseIf dbRst!kdsubinstalasi = "012" Then
                j = 35
            ElseIf dbRst!kdsubinstalasi = "014" Then
                j = 36
            ElseIf dbRst!kdsubinstalasi = "016" Then
                j = 37
            ElseIf dbRst!Spesialisasi = "Spesialisasi Lain" Then
                j = 38
            End If

            If oSheet.Cells(j, 7).value = "" Then Cell7 = 0 Else Cell7 = oSheet.Cells(j, 7).value
            If oSheet.Cells(j, 8).value = "" Then Cell8 = 0 Else Cell8 = oSheet.Cells(j, 8).value
            If oSheet.Cells(j, 9).value = "" Then Cell9 = 0 Else Cell9 = oSheet.Cells(j, 9).value
            If oSheet.Cells(j, 10).value = "" Then Cell10 = 0 Else Cell10 = oSheet.Cells(j, 10).value
            If dbRst!kdsubinstalasi = "001" Then
                With oSheet
                    If dbRst!RujukanTujuan = "Rumah Sakit" Then
                        .Cells(j, 7) = dbRst!TotalXRS + Cell7
                        .Cells(j, 8) = dbRst!Jml + Cell8

                    ElseIf dbRst!RujukanTujuan = "Puskesmas" Then
                        .Cells(j, 9) = dbRst!TotalXPus + Cell9
                        .Cells(j, 10) = dbRst!Jml + Cell10
                    End If
                End With
            ElseIf dbRst!kdsubinstalasi = "002" Then
                With oSheet
                    If dbRst!RujukanTujuan = "Rumah Sakit" Then
                        .Cells(j, 7) = dbRst!TotalXRS + Cell7
                        .Cells(j, 8) = dbRst!Jml + Cell8

                    ElseIf dbRst!RujukanTujuan = "Puskesmas" Then
                        .Cells(j, 9) = dbRst!TotalXPus + Cell9
                        .Cells(j, 10) = dbRst!Jml + Cell10
                    End If
                End With
            ElseIf dbRst!kdsubinstalasi = "003" Then
                With oSheet
                    If dbRst!RujukanTujuan = "Rumah Sakit" Then
                        .Cells(j, 7) = dbRst!TotalXRS + Cell7
                        .Cells(j, 8) = dbRst!Jml + Cell8

                    ElseIf dbRst!RujukanTujuan = "Puskesmas" Then
                        .Cells(j, 9) = dbRst!TotalXPus + Cell9
                        .Cells(j, 10) = dbRst!Jml + Cell10
                    End If
                End With
            ElseIf dbRst!kdsubinstalasi = "004" Then
                With oSheet
                    If dbRst!RujukanTujuan = "Rumah Sakit" Then
                        .Cells(j, 7) = dbRst!TotalXRS + Cell7
                        .Cells(j, 8) = dbRst!Jml + Cell8

                    ElseIf dbRst!RujukanTujuan = "Puskesmas" Then
                        .Cells(j, 9) = dbRst!TotalXPus + Cell9
                        .Cells(j, 10) = dbRst!Jml + Cell10
                    End If
                End With
            ElseIf dbRst!kdsubinstalasi = "005" Then
                With oSheet
                    If dbRst!RujukanTujuan = "Rumah Sakit" Then
                        .Cells(j, 7) = dbRst!TotalXRS + Cell7
                        .Cells(j, 8) = dbRst!Jml + Cell8

                    ElseIf dbRst!RujukanTujuan = "Puskesmas" Then
                        .Cells(j, 9) = dbRst!TotalXPus + Cell9
                        .Cells(j, 10) = dbRst!Jml + Cell10
                    End If
                End With
            ElseIf dbRst!kdsubinstalasi = "007" Then
                With oSheet
                    If dbRst!RujukanTujuan = "Rumah Sakit" Then
                        .Cells(j, 7) = dbRst!TotalXRS + Cell7
                        .Cells(j, 8) = dbRst!Jml + Cell8

                    ElseIf dbRst!RujukanTujuan = "Puskesmas" Then
                        .Cells(j, 9) = dbRst!TotalXPus + Cell9
                        .Cells(j, 10) = dbRst!Jml + Cell10
                    End If
                End With
            ElseIf dbRst!kdsubinstalasi = "008" Then
                With oSheet
                    If dbRst!RujukanTujuan = "Rumah Sakit" Then
                        .Cells(j, 7) = dbRst!TotalXRS + Cell7
                        .Cells(j, 8) = dbRst!Jml + Cell8

                    ElseIf dbRst!RujukanTujuan = "Puskesmas" Then
                        .Cells(j, 9) = dbRst!TotalXPus + Cell9
                        .Cells(j, 10) = dbRst!Jml + Cell10
                    End If
                End With
            ElseIf dbRst!kdsubinstalasi = "009" Then
                With oSheet
                    If dbRst!RujukanTujuan = "Rumah Sakit" Then
                        .Cells(j, 7) = dbRst!TotalXRS + Cell7
                        .Cells(j, 8) = dbRst!Jml + Cell8

                    ElseIf dbRst!RujukanTujuan = "Puskesmas" Then
                        .Cells(j, 9) = dbRst!TotalXPus + Cell9
                        .Cells(j, 10) = dbRst!Jml + Cell10
                    End If
                End With
            ElseIf dbRst!kdsubinstalasi = "010" Then
                With oSheet
                    If dbRst!RujukanTujuan = "Rumah Sakit" Then
                        .Cells(j, 7) = dbRst!TotalXRS + Cell7
                        .Cells(j, 8) = dbRst!Jml + Cell8

                    ElseIf dbRst!RujukanTujuan = "Puskesmas" Then
                        .Cells(j, 9) = dbRst!TotalXPus + Cell9
                        .Cells(j, 10) = dbRst!Jml + Cell10
                    End If
                End With
            ElseIf dbRst!kdsubinstalasi = "011" Then
                With oSheet
                    If dbRst!RujukanTujuan = "Rumah Sakit" Then
                        .Cells(j, 7) = dbRst!TotalXRS + Cell7
                        .Cells(j, 8) = dbRst!Jml + Cell8

                    ElseIf dbRst!RujukanTujuan = "Puskesmas" Then
                        .Cells(j, 9) = dbRst!TotalXPus + Cell9
                        .Cells(j, 10) = dbRst!Jml + Cell10
                    End If
                End With
            ElseIf dbRst!kdsubinstalasi = "012" Then
                With oSheet
                    If dbRst!RujukanTujuan = "Rumah Sakit" Then
                        .Cells(j, 7) = dbRst!TotalXRS + Cell7
                        .Cells(j, 8) = dbRst!Jml + Cell8

                    ElseIf dbRst!RujukanTujuan = "Puskesmas" Then
                        .Cells(j, 9) = dbRst!TotalXPus + Cell9
                        .Cells(j, 10) = dbRst!Jml + Cell10
                    End If
                End With
            ElseIf dbRst!kdsubinstalasi = "014" Then
                With oSheet
                    If dbRst!RujukanTujuan = "Rumah Sakit" Then
                        .Cells(j, 7) = dbRst!TotalXRS + Cell7
                        .Cells(j, 8) = dbRst!Jml + Cell8

                    ElseIf dbRst!RujukanTujuan = "Puskesmas" Then
                        .Cells(j, 9) = dbRst!TotalXPus + Cell9
                        .Cells(j, 10) = dbRst!Jml + Cell10
                    End If
                End With
            ElseIf dbRst!kdsubinstalasi = "016" Then
                With oSheet
                    If dbRst!RujukanTujuan = "Rumah Sakit" Then
                        .Cells(j, 7) = dbRst!TotalXRS + Cell7
                        .Cells(j, 8) = dbRst!Jml + Cell8

                    ElseIf dbRst!RujukanTujuan = "Puskesmas" Then
                        .Cells(j, 9) = dbRst!TotalXPus + Cell9
                        .Cells(j, 10) = dbRst!Jml + Cell10
                    End If
                End With
            ElseIf dbRst!Spesialisasi = "Spesialisasi Lain" Then
                With oSheet
                    If dbRst!RujukanTujuan = "Rumah Sakit" Then
                        .Cells(j, 7) = dbRst!TotalXRS + Cell7
                        .Cells(j, 8) = dbRst!Jml + Cell8

                    ElseIf dbRst!RujukanTujuan = "Puskesmas" Then
                        .Cells(j, 9) = dbRst!TotalXPus + Cell9
                        .Cells(j, 10) = dbRst!Jml + Cell10
                    End If
                End With
            End If
            dbRst.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select distinct * from RL1_24_2 where Tglterima between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'or tglterima is null"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF
            If dbRst!kdsubinstalasi = "001" Then
                j = 25
            ElseIf dbRst!kdsubinstalasi = "002" Then
                j = 26
            ElseIf dbRst!kdsubinstalasi = "003" Then
                j = 27
            ElseIf dbRst!kdsubinstalasi = "005" Then
                j = 28
            ElseIf dbRst!kdsubinstalasi = "004" Then
                j = 29
            ElseIf dbRst!kdsubinstalasi = "007" Then
                j = 30
            ElseIf dbRst!kdsubinstalasi = "008" Then
                j = 31
            ElseIf dbRst!kdsubinstalasi = "009" Then
                j = 32
            ElseIf dbRst!kdsubinstalasi = "010" Then
                j = 33
            ElseIf dbRst!kdsubinstalasi = "011" Then
                j = 34
            ElseIf dbRst!kdsubinstalasi = "012" Then
                j = 35
            ElseIf dbRst!kdsubinstalasi = "014" Then
                j = 36
            ElseIf dbRst!kdsubinstalasi = "016" Then
                j = 37
            ElseIf dbRst!Spesialisasi = "Spesialisasi Lain" Then
                j = 38
            End If

            If oSheet.Cells(j, 11).value = "" Then Cell11 = 0 Else Cell11 = oSheet.Cells(j, 11).value
            If oSheet.Cells(j, 12).value = "" Then Cell12 = 0 Else Cell12 = oSheet.Cells(j, 12).value
            If oSheet.Cells(j, 13).value = "" Then Cell13 = 0 Else Cell13 = oSheet.Cells(j, 13).value

            If dbRst!kdsubinstalasi = "001" Then
                With oSheet
                    .Cells(j, 11) = dbRst!TotalKunjungan + Cell11
                    .Cells(j, 12) = dbRst!KunjunganAsing + Cell12
                    .Cells(j, 13) = dbRst!Pasien + Cell13
                End With
            ElseIf dbRst!kdsubinstalasi = "002" Then
                With oSheet
                    .Cells(j, 11) = dbRst!TotalKunjungan + Cell11
                    .Cells(j, 12) = dbRst!KunjunganAsing + Cell12
                    .Cells(j, 13) = dbRst!Pasien + Cell13
                End With
            ElseIf dbRst!kdsubinstalasi = "003" Then
                With oSheet
                    .Cells(j, 11) = dbRst!TotalKunjungan + Cell11
                    .Cells(j, 12) = dbRst!KunjunganAsing + Cell12
                    .Cells(j, 13) = dbRst!Pasien + Cell13
                End With
            ElseIf dbRst!kdsubinstalasi = "004" Then
                With oSheet
                    .Cells(j, 11) = dbRst!TotalKunjungan + Cell11
                    .Cells(j, 12) = dbRst!KunjunganAsing + Cell12
                    .Cells(j, 13) = dbRst!Pasien + Cell13
                End With
            ElseIf dbRst!kdsubinstalasi = "005" Then
                With oSheet
                    .Cells(j, 11) = dbRst!TotalKunjungan + Cell11
                    .Cells(j, 12) = dbRst!KunjunganAsing + Cell12
                    .Cells(j, 13) = dbRst!Pasien + Cell13
                End With
            ElseIf dbRst!kdsubinstalasi = "007" Then
                With oSheet
                    .Cells(j, 11) = dbRst!TotalKunjungan + Cell11
                    .Cells(j, 12) = dbRst!KunjunganAsing + Cell12
                    .Cells(j, 13) = dbRst!Pasien + Cell13
                End With
            ElseIf dbRst!kdsubinstalasi = "008" Then
                With oSheet
                    .Cells(j, 11) = dbRst!TotalKunjungan + Cell11
                    .Cells(j, 12) = dbRst!KunjunganAsing + Cell12
                    .Cells(j, 13) = dbRst!Pasien + Cell13
                End With
            ElseIf dbRst!kdsubinstalasi = "009" Then
                With oSheet
                    .Cells(j, 11) = dbRst!TotalKunjungan + Cell11
                    .Cells(j, 12) = dbRst!KunjunganAsing + Cell12
                    .Cells(j, 13) = dbRst!Pasien + Cell13
                End With
            ElseIf dbRst!kdsubinstalasi = "010" Then
                With oSheet
                    .Cells(j, 11) = dbRst!TotalKunjungan + Cell11
                    .Cells(j, 12) = dbRst!KunjunganAsing + Cell12
                    .Cells(j, 13) = dbRst!Pasien + Cell13
                End With
            ElseIf dbRst!kdsubinstalasi = "011" Then
                With oSheet
                    .Cells(j, 11) = dbRst!TotalKunjungan + Cell11
                    .Cells(j, 12) = dbRst!KunjunganAsing + Cell12
                    .Cells(j, 13) = dbRst!Pasien + Cell13
                End With
            ElseIf dbRst!kdsubinstalasi = "012" Then
                With oSheet
                    .Cells(j, 11) = dbRst!TotalKunjungan + Cell11
                    .Cells(j, 12) = dbRst!KunjunganAsing + Cell12
                    .Cells(j, 13) = dbRst!Pasien + Cell13
                End With
            ElseIf dbRst!kdsubinstalasi = "014" Then
                With oSheet
                    .Cells(j, 11) = dbRst!TotalKunjungan + Cell11
                    .Cells(j, 12) = dbRst!KunjunganAsing + Cell12
                    .Cells(j, 13) = dbRst!Pasien + Cell13
                End With
            ElseIf dbRst!kdsubinstalasi = "016" Then
                With oSheet
                    .Cells(j, 11) = dbRst!TotalKunjungan + Cell11
                    .Cells(j, 12) = dbRst!KunjunganAsing + Cell12
                    .Cells(j, 13) = dbRst!Pasien + Cell13
                End With
            ElseIf dbRst!Spesialisasi = "Spesialisasi Lain" Then
                With oSheet
                    .Cells(j, 11) = dbRst!TotalKunjungan + Cell11
                    .Cells(j, 12) = dbRst!KunjunganAsing + Cell12
                    .Cells(j, 13) = dbRst!Pasien + Cell13
                End With
            End If
            dbRst.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select distinct * from RL1_24_3 where Tglmasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'or tglmasuk is null"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF
            If dbRst!kdsubinstalasi = "001" Then
                j = 25
            ElseIf dbRst!kdsubinstalasi = "002" Then
                j = 26
            ElseIf dbRst!kdsubinstalasi = "003" Then
                j = 27
            ElseIf dbRst!kdsubinstalasi = "005" Then
                j = 28
            ElseIf dbRst!kdsubinstalasi = "004" Then
                j = 29
            ElseIf dbRst!kdsubinstalasi = "007" Then
                j = 30
            ElseIf dbRst!kdsubinstalasi = "008" Then
                j = 31
            ElseIf dbRst!kdsubinstalasi = "009" Then
                j = 32
            ElseIf dbRst!kdsubinstalasi = "010" Then
                j = 33
            ElseIf dbRst!kdsubinstalasi = "011" Then
                j = 34
            ElseIf dbRst!kdsubinstalasi = "012" Then
                j = 35
            ElseIf dbRst!kdsubinstalasi = "014" Then
                j = 36
            ElseIf dbRst!kdsubinstalasi = "016" Then
                j = 37
            ElseIf dbRst!Spesialisasi = "Spesialisasi Lain" Then
                j = 38
            End If

            If oSheet.Cells(j, 14).value = "" Then Cell14 = 0 Else Cell14 = oSheet.Cells(j, 14).value
            If oSheet.Cells(j, 15).value = "" Then Cell15 = 0 Else Cell15 = oSheet.Cells(j, 15).value
            If oSheet.Cells(j, 16).value = "" Then Cell16 = 0 Else Cell16 = oSheet.Cells(j, 16).value
            If oSheet.Cells(j, 17).value = "" Then Cell17 = 0 Else Cell17 = oSheet.Cells(j, 17).value
            If oSheet.Cells(j, 18).value = "" Then Cell18 = 0 Else Cell18 = oSheet.Cells(j, 18).value
            If oSheet.Cells(j, 19).value = "" Then Cell19 = 0 Else Cell19 = oSheet.Cells(j, 19).value
            If oSheet.Cells(j, 20).value = "" Then Cell20 = 0 Else Cell20 = oSheet.Cells(j, 20).value
            If oSheet.Cells(j, 21).value = "" Then Cell21 = 0 Else Cell21 = oSheet.Cells(j, 21).value
            If oSheet.Cells(j, 22).value = "" Then Cell22 = 0 Else Cell22 = oSheet.Cells(j, 22).value

            If dbRst!kdsubinstalasi = "001" Then
                With oSheet
                    .Cells(j, 14) = dbRst!DariPuskesmas + Cell14
                    .Cells(j, 15) = dbRst!DariFasilitasLain + Cell15
                    .Cells(j, 16) = dbRst!DariRSLain + Cell16
                    .Cells(j, 17) = dbRst!DikembalikanPuskesmas + Cell17
                    .Cells(j, 18) = dbRst!DikembalikanFasilitasLain + Cell18
                    .Cells(j, 19) = dbRst!DikembalikanRSLain + Cell19
                    .Cells(j, 20) = dbRst!PasienRujukan + Cell20
                    .Cells(j, 21) = dbRst!DatangSendiri + Cell21
                    .Cells(j, 22) = dbRst!DiterimaKembali + Cell22
                End With
            ElseIf dbRst!kdsubinstalasi = "002" Then
                With oSheet
                    .Cells(j, 14) = dbRst!DariPuskesmas + Cell14
                    .Cells(j, 15) = dbRst!DariFasilitasLain + Cell15
                    .Cells(j, 16) = dbRst!DariRSLain + Cell16
                    .Cells(j, 17) = dbRst!DikembalikanPuskesmas + Cell17
                    .Cells(j, 18) = dbRst!DikembalikanFasilitasLain + Cell18
                    .Cells(j, 19) = dbRst!DikembalikanRSLain + Cell19
                    .Cells(j, 20) = dbRst!PasienRujukan + Cell20
                    .Cells(j, 21) = dbRst!DatangSendiri + Cell21
                    .Cells(j, 22) = dbRst!DiterimaKembali + Cell22
                End With
            ElseIf dbRst!kdsubinstalasi = "003" Then
                With oSheet
                    .Cells(j, 14) = dbRst!DariPuskesmas + Cell14
                    .Cells(j, 15) = dbRst!DariFasilitasLain + Cell15
                    .Cells(j, 16) = dbRst!DariRSLain + Cell16
                    .Cells(j, 17) = dbRst!DikembalikanPuskesmas + Cell17
                    .Cells(j, 18) = dbRst!DikembalikanFasilitasLain + Cell18
                    .Cells(j, 19) = dbRst!DikembalikanRSLain + Cell19
                    .Cells(j, 20) = dbRst!PasienRujukan + Cell20
                    .Cells(j, 21) = dbRst!DatangSendiri + Cell21
                    .Cells(j, 22) = dbRst!DiterimaKembali + Cell22
                End With
            ElseIf dbRst!kdsubinstalasi = "004" Then
                With oSheet
                    .Cells(j, 14) = dbRst!DariPuskesmas + Cell14
                    .Cells(j, 15) = dbRst!DariFasilitasLain + Cell15
                    .Cells(j, 16) = dbRst!DariRSLain + Cell16
                    .Cells(j, 17) = dbRst!DikembalikanPuskesmas + Cell17
                    .Cells(j, 18) = dbRst!DikembalikanFasilitasLain + Cell18
                    .Cells(j, 19) = dbRst!DikembalikanRSLain + Cell19
                    .Cells(j, 20) = dbRst!PasienRujukan + Cell20
                    .Cells(j, 21) = dbRst!DatangSendiri + Cell21
                    .Cells(j, 22) = dbRst!DiterimaKembali + Cell22
                End With
            ElseIf dbRst!kdsubinstalasi = "005" Then
                With oSheet
                    .Cells(j, 14) = dbRst!DariPuskesmas + Cell14
                    .Cells(j, 15) = dbRst!DariFasilitasLain + Cell15
                    .Cells(j, 16) = dbRst!DariRSLain + Cell16
                    .Cells(j, 17) = dbRst!DikembalikanPuskesmas + Cell17
                    .Cells(j, 18) = dbRst!DikembalikanFasilitasLain + Cell18
                    .Cells(j, 19) = dbRst!DikembalikanRSLain + Cell19
                    .Cells(j, 20) = dbRst!PasienRujukan + Cell20
                    .Cells(j, 21) = dbRst!DatangSendiri + Cell21
                    .Cells(j, 22) = dbRst!DiterimaKembali + Cell22
                End With
            ElseIf dbRst!kdsubinstalasi = "007" Then
                With oSheet
                    .Cells(j, 14) = dbRst!DariPuskesmas + Cell14
                    .Cells(j, 15) = dbRst!DariFasilitasLain + Cell15
                    .Cells(j, 16) = dbRst!DariRSLain + Cell16
                    .Cells(j, 17) = dbRst!DikembalikanPuskesmas + Cell17
                    .Cells(j, 18) = dbRst!DikembalikanFasilitasLain + Cell18
                    .Cells(j, 19) = dbRst!DikembalikanRSLain + Cell19
                    .Cells(j, 20) = dbRst!PasienRujukan + Cell20
                    .Cells(j, 21) = dbRst!DatangSendiri + Cell21
                    .Cells(j, 22) = dbRst!DiterimaKembali + Cell22
                End With
            ElseIf dbRst!kdsubinstalasi = "008" Then
                With oSheet
                    .Cells(j, 14) = dbRst!DariPuskesmas + Cell14
                    .Cells(j, 15) = dbRst!DariFasilitasLain + Cell15
                    .Cells(j, 16) = dbRst!DariRSLain + Cell16
                    .Cells(j, 17) = dbRst!DikembalikanPuskesmas + Cell17
                    .Cells(j, 18) = dbRst!DikembalikanFasilitasLain + Cell18
                    .Cells(j, 19) = dbRst!DikembalikanRSLain + Cell19
                    .Cells(j, 20) = dbRst!PasienRujukan + Cell20
                    .Cells(j, 21) = dbRst!DatangSendiri + Cell21
                    .Cells(j, 22) = dbRst!DiterimaKembali + Cell22
                End With
            ElseIf dbRst!kdsubinstalasi = "009" Then
                With oSheet
                    .Cells(j, 14) = dbRst!DariPuskesmas + Cell14
                    .Cells(j, 15) = dbRst!DariFasilitasLain + Cell15
                    .Cells(j, 16) = dbRst!DariRSLain + Cell16
                    .Cells(j, 17) = dbRst!DikembalikanPuskesmas + Cell17
                    .Cells(j, 18) = dbRst!DikembalikanFasilitasLain + Cell18
                    .Cells(j, 19) = dbRst!DikembalikanRSLain + Cell19
                    .Cells(j, 20) = dbRst!PasienRujukan + Cell20
                    .Cells(j, 21) = dbRst!DatangSendiri + Cell21
                    .Cells(j, 22) = dbRst!DiterimaKembali + Cell22
                End With
            ElseIf dbRst!kdsubinstalasi = "010" Then
                With oSheet
                    .Cells(j, 14) = dbRst!DariPuskesmas + Cell14
                    .Cells(j, 15) = dbRst!DariFasilitasLain + Cell15
                    .Cells(j, 16) = dbRst!DariRSLain + Cell16
                    .Cells(j, 17) = dbRst!DikembalikanPuskesmas + Cell17
                    .Cells(j, 18) = dbRst!DikembalikanFasilitasLain + Cell18
                    .Cells(j, 19) = dbRst!DikembalikanRSLain + Cell19
                    .Cells(j, 20) = dbRst!PasienRujukan + Cell20
                    .Cells(j, 21) = dbRst!DatangSendiri + Cell21
                    .Cells(j, 22) = dbRst!DiterimaKembali + Cell22
                End With
            ElseIf dbRst!kdsubinstalasi = "011" Then
                With oSheet
                    .Cells(j, 14) = dbRst!DariPuskesmas + Cell14
                    .Cells(j, 15) = dbRst!DariFasilitasLain + Cell15
                    .Cells(j, 16) = dbRst!DariRSLain + Cell16
                    .Cells(j, 17) = dbRst!DikembalikanPuskesmas + Cell17
                    .Cells(j, 18) = dbRst!DikembalikanFasilitasLain + Cell18
                    .Cells(j, 19) = dbRst!DikembalikanRSLain + Cell19
                    .Cells(j, 20) = dbRst!PasienRujukan + Cell20
                    .Cells(j, 21) = dbRst!DatangSendiri + Cell21
                    .Cells(j, 22) = dbRst!DiterimaKembali + Cell22
                End With
            ElseIf dbRst!kdsubinstalasi = "012" Then
                With oSheet
                    .Cells(j, 14) = dbRst!DariPuskesmas + Cell14
                    .Cells(j, 15) = dbRst!DariFasilitasLain + Cell15
                    .Cells(j, 16) = dbRst!DariRSLain + Cell16
                    .Cells(j, 17) = dbRst!DikembalikanPuskesmas + Cell17
                    .Cells(j, 18) = dbRst!DikembalikanFasilitasLain + Cell18
                    .Cells(j, 19) = dbRst!DikembalikanRSLain + Cell19
                    .Cells(j, 20) = dbRst!PasienRujukan + Cell20
                    .Cells(j, 21) = dbRst!DatangSendiri + Cell21
                    .Cells(j, 22) = dbRst!DiterimaKembali + Cell22
                End With
            ElseIf dbRst!kdsubinstalasi = "014" Then
                With oSheet
                    .Cells(j, 14) = dbRst!DariPuskesmas + Cell14
                    .Cells(j, 15) = dbRst!DariFasilitasLain + Cell15
                    .Cells(j, 16) = dbRst!DariRSLain + Cell16
                    .Cells(j, 17) = dbRst!DikembalikanPuskesmas + Cell17
                    .Cells(j, 18) = dbRst!DikembalikanFasilitasLain + Cell18
                    .Cells(j, 19) = dbRst!DikembalikanRSLain + Cell19
                    .Cells(j, 20) = dbRst!PasienRujukan + Cell20
                    .Cells(j, 21) = dbRst!DatangSendiri + Cell21
                    .Cells(j, 22) = dbRst!DiterimaKembali + Cell22
                End With
            ElseIf dbRst!kdsubinstalasi = "016" Then
                With oSheet
                    .Cells(j, 14) = dbRst!DariPuskesmas + Cell14
                    .Cells(j, 15) = dbRst!DariFasilitasLain + Cell15
                    .Cells(j, 16) = dbRst!DariRSLain + Cell16
                    .Cells(j, 17) = dbRst!DikembalikanPuskesmas + Cell17
                    .Cells(j, 18) = dbRst!DikembalikanFasilitasLain + Cell18
                    .Cells(j, 19) = dbRst!DikembalikanRSLain + Cell19
                    .Cells(j, 20) = dbRst!PasienRujukan + Cell20
                    .Cells(j, 21) = dbRst!DatangSendiri + Cell21
                    .Cells(j, 22) = dbRst!DiterimaKembali + Cell22
                End With
            ElseIf dbRst!Spesialisasi = "Spesialisasi Lain" Then
                With oSheet
                    .Cells(j, 14) = dbRst!DariPuskesmas + Cell14
                    .Cells(j, 15) = dbRst!DariFasilitasLain + Cell15
                    .Cells(j, 16) = dbRst!DariRSLain + Cell16
                    .Cells(j, 17) = dbRst!DikembalikanPuskesmas + Cell17
                    .Cells(j, 18) = dbRst!DikembalikanFasilitasLain + Cell18
                    .Cells(j, 19) = dbRst!DikembalikanRSLain + Cell19
                    .Cells(j, 20) = dbRst!PasienRujukan + Cell20
                    .Cells(j, 21) = dbRst!DatangSendiri + Cell21
                    .Cells(j, 22) = dbRst!DiterimaKembali + Cell22
                End With
            End If
            dbRst.MoveNext
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
