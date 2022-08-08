VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRL1HAL3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL1 Halaman 3"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRL1HAL3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6315
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
         Format          =   126418947
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
         Format          =   126418947
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
         Format          =   126418947
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
      Picture         =   "frmRL1HAL3.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2955
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRL1HAL3.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL1HAL3.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmRL1HAL3"
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
Dim Cell7 As String
Dim Cell12 As String
Dim Cell13 As String
Dim Cell14 As String

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
    Set oWB = oXL.Workbooks.Open(App.Path & "\RL1 Hal3.xls")
    Set oSheet = oWB.ActiveSheet

    Set rsb = Nothing
    strSQL = "select * from profilrs"
    Call msubRecFO(rsb, strSQL)

    Set oResizeRange = oSheet.Range("g1", "g2")
    oResizeRange.value = Trim(rsb!KdRs)

    strSQL = "Select * from V_PengadaanObat where TglTerima between '" & Format(dtpAwal.value, "yyyy/MM/dd") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd") & "'"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF

            If dbRst!KdKategoryBarang = "01" Then
                j = 42
            ElseIf dbRst!KdKategoryBarang = "02" Then
                j = 43
            ElseIf dbRst!KdKategoryBarang = "03" Then
                j = 44
            End If

            Cell12 = oSheet.Cells(j, 13).value
            Cell13 = oSheet.Cells(j, 16).value

            If dbRst!KdKategoryBarang = "01" Then
                With oSheet
                    .Cells(j, 13) = Trim(dbRst!jmlnonformularium + Cell12)
                    .Cells(j, 16) = Trim(dbRst!jmlformularium + Cell13)
                End With
            ElseIf dbRst!KdKategoryBarang = "02" Then
                With oSheet
                    .Cells(j, 13) = Trim(dbRst!jmlnonformularium + Cell12)
                    .Cells(j, 16) = Trim(dbRst!jmlformularium + Cell13)
                End With
            ElseIf dbRst!KdKategoryBarang = "03" Then
                With oSheet
                    .Cells(j, 13) = Trim(dbRst!jmlnonformularium + Cell12)
                    .Cells(j, 16) = Trim(dbRst!jmlformularium + Cell13)
                End With
            End If
            dbRst.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select distinct * from RL1_9A where TglPelayanan between '" & Format(dtpAwal.value, "yyyy/MM/dd") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd") & "'"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF

            If dbRst!Judul = "Foto tanpa bahan kontras" Then
                j = 6
            ElseIf dbRst!Judul = "Foto dengan bahan kontras" Then
                j = 7
            ElseIf dbRst!Judul = "Foto dengan rol film" Then
                j = 8
            ElseIf dbRst!Judul = "Flouroskopi" Then
                j = 9
            ElseIf dbRst!Judul = "Dento alveolair" Then
                j = 11
            ElseIf dbRst!Judul = "Panoramic" Then
                j = 12
            ElseIf dbRst!Judul = "Cephalographi" Then
                j = 13
            ElseIf dbRst!Judul = "C.T. Scan Dikepala" Then
                j = 15
            ElseIf dbRst!Judul = "C.T. Scan Diluar kepala" Then
                j = 16
            ElseIf dbRst!Judul = "Lymphografi" Then
                j = 17
            ElseIf dbRst!Judul = "Angiograpi" Then
                j = 18
            ElseIf dbRst!Judul = "Lain-lain" Then
                j = 19
            End If

            If oSheet.Cells(j, 7).value = "" Then Cell7 = 0 Else Cell7 = oSheet.Cells(j, 7).value

            If dbRst!Judul = "Foto tanpa bahan kontras" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!Judul = "Foto dengan bahan kontras" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!Judul = "Foto dengan rol film" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!Judul = "Flouroskopi" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!Judul = "Dento alveolair" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!Judul = "Panoramic" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!Judul = "Cephalographi" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!Judul = "C.T. Scan Dikepala" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!Judul = "C.T. Scan Diluar kepala" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!Judul = "Lymphografi" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!Judul = "Angiograpi" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!Judul = "Lain-lain" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            End If
            dbRst.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select distinct * from RL1_9B where TglPelayanan between '" & Format(dtpAwal.value, "yyyy/MM/dd") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd") & "'"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF

            If dbRst!Judul = "Jumlah Kegiatan Radiotherapi" Then
                j = 24
            ElseIf dbRst!Judul = "Lain-lain" Then
                j = 25

            End If

            If oSheet.Cells(j, 7).value = "" Then Cell7 = 0 Else Cell7 = oSheet.Cells(j, 7).value

            If dbRst!Judul = "Jumlah Kegiatan Radiotherapi" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!Judul = "Lain-lain" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With

            End If
            dbRst.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select distinct * from RL1_9C where TglPelayanan between '" & Format(dtpAwal.value, "yyyy/MM/dd") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd") & "'"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF

            If dbRst!Judul = "Jumlah Kegiatan Diagnostik" Then
                j = 29
            ElseIf dbRst!Judul = "Jumlah Kegiatan Therapi" Then
                j = 30
            ElseIf dbRst!Judul = "Lain-lain" Then
                j = 31

            End If

            If oSheet.Cells(j, 7).value = "" Then Cell7 = 0 Else Cell7 = oSheet.Cells(j, 7).value

            If dbRst!Judul = "Jumlah Kegiatan Diagnostik" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!Judul = "Jumlah Kegiatan Therapi" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!Judul = "Lain-lain" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            End If
            dbRst.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select distinct * from RL1_9D where TglPelayanan between '" & Format(dtpAwal.value, "yyyy/MM/dd") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd") & "'"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF

            If dbRst!Judul = "USG" Then
                j = 35
            ElseIf dbRst!Judul = "MRI" Then
                j = 36
            ElseIf dbRst!Judul = "Lain-lain" Then
                j = 37

            End If

            If oSheet.Cells(j, 7).value = "" Then Cell7 = 0 Else Cell7 = oSheet.Cells(j, 7).value

            If dbRst!Judul = "USG" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!Judul = "MRI" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!Judul = "Lain-lain" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            End If
            dbRst.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select distinct * from RL1_10 where TglPelayanan between '" & Format(dtpAwal.value, "yyyy/MM/dd") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd") & "'"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF

            If dbRst!JenisKegiatan = "Electro Encephalografi (EEG)" Then
                j = 41
            ElseIf dbRst!JenisKegiatan = "Electro Kardiographi (EKG)" Then
                j = 42
            ElseIf dbRst!JenisKegiatan = "Endoskopi (semua bentuk)" Then
                j = 43
            ElseIf dbRst!JenisKegiatan = "Hemodialisa" Then
                j = 44
            ElseIf dbRst!JenisKegiatan = "Densitometri Tulang" Then
                j = 45
            ElseIf dbRst!JenisKegiatan = "Koreksi Fraktur/Dislokasi non Bedah" Then
                j = 46
            ElseIf dbRst!JenisKegiatan = "Pungsi" Then
                j = 47
            ElseIf dbRst!JenisKegiatan = "Spirometri" Then
                j = 48
            ElseIf dbRst!JenisKegiatan = "Test Kulit/Alergi/Histamin" Then
                j = 49
            ElseIf dbRst!JenisKegiatan = "Topometri" Then
                j = 50
            ElseIf dbRst!JenisKegiatan = "Treadmill/Exercise Test" Then
                j = 51
            ElseIf dbRst!JenisKegiatan = "Lain-lain" Then
                j = 52
            End If

            If oSheet.Cells(j, 7).value = "" Then Cell7 = 0 Else Cell7 = oSheet.Cells(j, 7).value

            If dbRst!JenisKegiatan = "Electro Encephalografi (EEG)" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!JenisKegiatan = "Electro Kardiographi (EKG)" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!JenisKegiatan = "Endoskopi (semua bentuk)" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!JenisKegiatan = "Hemodialisa" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!JenisKegiatan = "Densitometri Tulang" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!JenisKegiatan = "Koreksi Fraktur/Dislokasi non Bedah" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!JenisKegiatan = "Pungsi" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!JenisKegiatan = "Spirometri" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!JenisKegiatan = "Test Kulit/Alergi/Histamin" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!JenisKegiatan = "Topometri" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!JenisKegiatan = "Treadmill/Exercise Test" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            ElseIf dbRst!JenisKegiatan = "Lain-lain" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jumlah + Cell7)
                End With
            End If
            dbRst.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select distinct * from RL1_11a where TglPelayanan between '" & Format(dtpAwal.value, "yyyy/MM/dd") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd") & "'"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF
            If dbRst!NamaExternal = "Kimia" Then
                j = 6
            ElseIf dbRst!NamaExternal = "Gula Darah" Then
                j = 7
            ElseIf dbRst!NamaExternal = "Hematologi" Then
                j = 8
            ElseIf dbRst!NamaExternal = "Serologi" Then
                j = 9
            ElseIf dbRst!NamaExternal = "Bakteriologi" Then
                j = 10
            ElseIf dbRst!NamaExternal = "Liquor" Then
                j = 11
            ElseIf dbRst!NamaExternal = "Transudat/Exsudat" Then
                j = 12
            ElseIf dbRst!NamaExternal = "Urine" Then
                j = 13
            ElseIf dbRst!NamaExternal = "Tinja" Then
                j = 14
            ElseIf dbRst!NamaExternal = "Analisa Gas Darah" Then
                j = 15
            ElseIf dbRst!NamaExternal = "Radio Assay" Then
                j = 16
            ElseIf dbRst!NamaExternal = "Cairan Otak" Then
                j = 17
            ElseIf dbRst!NamaExternal = "Cairan Tubuh Lain nya" Then
                j = 18
            ElseIf dbRst!NamaExternal = "Imunologi" Then
                j = 19
            ElseIf dbRst!NamaExternal = "Mikrobiologi Klinik" Then
                j = 20
            ElseIf dbRst!NamaExternal = "Lain-lain" Then
                j = 21
            End If

            If oSheet.Cells(j, 12).value = "" Then Cell12 = 0 Else Cell12 = oSheet.Cells(j, 12).value
            If oSheet.Cells(j, 13).value = "" Then Cell13 = 0 Else Cell13 = oSheet.Cells(j, 13).value
            If oSheet.Cells(j, 14).value = "" Then Cell14 = 0 Else Cell14 = oSheet.Cells(j, 14).value

            If dbRst!NamaExternal = "Kimia" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Gula Darah" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Hematologi" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Serologi" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Bakteriologi" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Liquor" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Transudat/Exsudat" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Urine" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Tinja" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Analisa Gas Darah" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Radio Assay" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Cairan Otak" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Cairan Tubuh Lain nya" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Imunologi" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Mikrobiologi Klinik" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Lain-lain" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            End If
            dbRst.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select distinct * from RL1_11b where TglPelayanan between '" & Format(dtpAwal.value, "yyyy/MM/dd") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd") & "'"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF

            If dbRst!NamaExternal = "Sitologi" Then
                j = 26
            ElseIf dbRst!NamaExternal = "Histologi" Then
                j = 27
            ElseIf dbRst!NamaExternal = "Lain-lain" Then
                j = 28
            End If

            If oSheet.Cells(j, 12).value = "" Then Cell12 = 0 Else Cell12 = oSheet.Cells(j, 12).value
            If oSheet.Cells(j, 13).value = "" Then Cell13 = 0 Else Cell13 = oSheet.Cells(j, 13).value
            If oSheet.Cells(j, 14).value = "" Then Cell14 = 0 Else Cell14 = oSheet.Cells(j, 14).value

            If dbRst!NamaExternal = "Sitologi" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Histologi" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Lain-lain" Then
                With oSheet
                    If dbRst!LevelProduk = "Sederhana" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Sedang" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    ElseIf dbRst!LevelProduk = "Canggih" Then
                        .Cells(j, 14) = Trim(dbRst!Jml + Cell14)
                    End If
                End With
            End If
            dbRst.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select distinct * from RL1_11c where TglPelayanan between '" & Format(dtpAwal.value, "yyyy/MM/dd") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd") & "'"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF
            If dbRst!NamaExternal = "Narkotika" Then
                j = 33
            ElseIf dbRst!NamaExternal = "Psikotropika" Then
                j = 34
            ElseIf dbRst!NamaExternal = "Zat Aditif" Then
                j = 35
            ElseIf dbRst!NamaExternal = "Pestisida" Then
                j = 36
            ElseIf dbRst!NamaExternal = "Zat Toksiologi" Then
                j = 37
            End If

            If oSheet.Cells(j, 12).value = "" Then Cell12 = 0 Else Cell12 = oSheet.Cells(j, 12).value
            If oSheet.Cells(j, 13).value = "" Then Cell13 = 0 Else Cell13 = oSheet.Cells(j, 13).value

            If dbRst!NamaExternal = "Narkotika" Then
                With oSheet
                    If dbRst!LevelProduk = "Skrining" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Konfirmasi" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Psikotropika" Then
                With oSheet
                    If dbRst!LevelProduk = "Skrining" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Konfirmasi" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Zat Aditif" Then
                With oSheet
                    If dbRst!LevelProduk = "Skrining" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Konfirmasi" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Pestisida" Then
                With oSheet
                    If dbRst!LevelProduk = "Skrining" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Konfirmasi" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    End If
                End With
            ElseIf dbRst!NamaExternal = "Zat Toksiologi" Then
                With oSheet
                    If dbRst!LevelProduk = "Skrining" Then
                        .Cells(j, 12) = Trim(dbRst!Jml + Cell12)
                    ElseIf dbRst!LevelProduk = "Konfirmasi" Then
                        .Cells(j, 13) = Trim(dbRst!Jml + Cell13)
                    End If
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
