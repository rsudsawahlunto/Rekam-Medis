VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRL1HAL2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL1 Halaman 2"
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
   Icon            =   "frmRL1HAL2.frx":0000
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
         Format          =   115212291
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
         Format          =   115212291
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
         Format          =   115212291
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
      Picture         =   "frmRL1HAL2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2955
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRL1HAL2.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL1HAL2.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmRL1HAL2"
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
Dim Cell10 As String
Dim Cell22 As String
Dim Cell23 As String
Dim Cell24 As String
Dim Cell25 As String
Dim Cell27 As String
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
    Set oWB = oXL.Workbooks.Open(App.Path & "\RL1 Hal2.xls")
    Set oSheet = oWB.ActiveSheet

    Set rsb = Nothing
    strSQL = "select * from profilrs"
    Call msubRecFO(rsb, strSQL)

    Set oResizeRange = oSheet.Range("g2", "g3")
    oResizeRange.value = Trim(rsb!KdRs)

    strSQL = "Select * from rl1_2 where TglPendaftaran between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'or tglpendaftaran is null"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF

            If dbRst!statuspasien = "Baru" Then
                j = 10
            ElseIf dbRst!statuspasien = "Lama" Then
                j = 11
            End If

            Cell7 = oSheet.Cells(j, 7).value

            If dbRst!statuspasien = "Baru" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jml + Cell7)
                End With
            ElseIf dbRst!statuspasien = "Lama" Then
                With oSheet
                    .Cells(j, 7) = Trim(dbRst!Jml + Cell7)
                End With
            End If
            dbRst.MoveNext
        Wend
    End If

    strSQL = "Select * from rl1_3 where Tglmasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'or tglmasuk is null"
    Call msubRecFO(rs2, strSQL)

    If rs2.RecordCount > 0 Then
        rs2.MoveFirst

        While Not rs2.EOF

            If rs2!kdsubinstalasi = "001" Then
                j = 16
            ElseIf rs2!kdsubinstalasi = "002" Then
                j = 17
            ElseIf rs2!kdsubinstalasi = "003" Then
                j = 20
            ElseIf rs2!kdsubinstalasi = "004" Then
                j = 22
            ElseIf rs2!kdsubinstalasi = "005" Then
                j = 23
            ElseIf rs2!kdsubinstalasi = "006" Then
                j = 25
            ElseIf rs2!kdsubinstalasi = "007" Then
                j = 26
            ElseIf rs2!kdsubinstalasi = "008" Then
                j = 27
            ElseIf rs2!kdsubinstalasi = "026" Then
                j = 28
            ElseIf rs2!kdsubinstalasi = "009" Then
                j = 29
            ElseIf rs2!kdsubinstalasi = "010" Then
                j = 30
            ElseIf rs2!kdsubinstalasi = "011" Then
                j = 31
            ElseIf rs2!kdsubinstalasi = "012" Then
                j = 32
            ElseIf rs2!kdsubinstalasi = "013" Then
                j = 33
            ElseIf rs2!kdsubinstalasi = "014" Then
                j = 34
            ElseIf rs2!kdsubinstalasi = "015" Then
                j = 35
            ElseIf rs2!kdsubinstalasi = "016" Then
                j = 36
            ElseIf rs2!kdsubinstalasi = "017" Then
                j = 37
            ElseIf rs2!kdsubinstalasi = "018" Then
                j = 38
            ElseIf rs2!kdsubinstalasi = "019" Then
                j = 39
            ElseIf rs2!kdsubinstalasi = "020" Then
                j = 40
            ElseIf rs2!kdsubinstalasi = "029" Then
                j = 41
            ElseIf rs2!kdsubinstalasi = "030" Then
                j = 42
            ElseIf rs2!kdsubinstalasi = "031" Then
                j = 43
            End If

            Cell7 = oSheet.Cells(j, 7).value
            Cell8 = oSheet.Cells(j, 8).value

            If rs2!kdsubinstalasi = "001" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "002" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "003" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "004" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "005" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "006" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "007" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "008" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "026" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "009" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "010" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "011" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "012" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "013" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "014" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "015" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "016" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "017" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "018" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "019" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "020" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "029" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "030" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            ElseIf rs2!kdsubinstalasi = "031" Then
                With oSheet
                    .Cells(j, 7) = Trim(rs2!jmlbaru + Cell7)
                    .Cells(j, 8) = Trim(rs2!jmllama + Cell8)
                End With

            End If
            rs2.MoveNext
        Wend
    End If

    strSQL = "Select * from RL1_7 where Tglmasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'or tglmasuk is null"
    Call msubRecFO(rsx, strSQL)

    If rsx.RecordCount > 0 Then
        rsx.MoveFirst

        While Not rsx.EOF

            If rsx!JenisPelayanan = "Bedah" Then
                j = 22
            ElseIf rsx!JenisPelayanan = "NonBedah" Then
                j = 23
            ElseIf rsx!JenisPelayanan = "Kebidanan" Then
                j = 24
            ElseIf rsx!JenisPelayanan = "Psikiatrik" Then
                j = 25
            ElseIf rsx!JenisPelayanan = "Anak" Then
                j = 26
            End If

            Cell22 = oSheet.Cells(j, 22).value
            Cell23 = oSheet.Cells(j, 23).value
            Cell24 = oSheet.Cells(j, 24).value
            Cell25 = oSheet.Cells(j, 25).value
            Cell27 = oSheet.Cells(j, 27).value

            If rsx!JenisPelayanan = "Bedah" Then
                With oSheet
                    .Cells(j, 22) = Trim(rsx![Rujukan] + Cell22)
                    .Cells(j, 23) = Trim(rsx![NonRujukan] + Cell23)
                    .Cells(j, 24) = Trim(rsx![DiRawat] + Cell24)
                    .Cells(j, 25) = Trim(rsx![DiRujuk] + Cell25)
                    .Cells(j, 27) = Trim(rsx![Mati] + Cell27)
                End With
            ElseIf rsx!JenisPelayanan = "NonBedah" Then
                With oSheet
                    .Cells(j, 22) = Trim(rsx![Rujukan] + Cell22)
                    .Cells(j, 23) = Trim(rsx![NonRujukan] + Cell23)
                    .Cells(j, 24) = Trim(rsx![DiRawat] + Cell24)
                    .Cells(j, 25) = Trim(rsx![DiRujuk] + Cell25)
                    .Cells(j, 27) = Trim(rsx![Mati] + Cell27)
                End With
            ElseIf rsx!JenisPelayanan = "Kebidanan" Then
                With oSheet
                    .Cells(j, 22) = Trim(rsx![Rujukan] + Cell22)
                    .Cells(j, 23) = Trim(rsx![NonRujukan] + Cell23)
                    .Cells(j, 24) = Trim(rsx![DiRawat] + Cell24)
                    .Cells(j, 25) = Trim(rsx![DiRujuk] + Cell25)
                    .Cells(j, 27) = Trim(rsx![Mati] + Cell27)
                End With
            ElseIf rsx!JenisPelayanan = "Psikiatrik" Then
                With oSheet
                    .Cells(j, 22) = Trim(rsx![Rujukan] + Cell22)
                    .Cells(j, 23) = Trim(rsx![NonRujukan] + Cell23)
                    .Cells(j, 24) = Trim(rsx![DiRawat] + Cell24)
                    .Cells(j, 25) = Trim(rsx![DiRujuk] + Cell25)
                    .Cells(j, 27) = Trim(rsx![Mati] + Cell27)
                End With
            ElseIf rsx!JenisPelayanan = "Anak" Then
                With oSheet
                    .Cells(j, 22) = Trim(rsx![Rujukan] + Cell22)
                    .Cells(j, 23) = Trim(rsx![NonRujukan] + Cell23)
                    .Cells(j, 24) = Trim(rsx![DiRawat] + Cell24)
                    .Cells(j, 25) = Trim(rsx![DiRujuk] + Cell25)
                    .Cells(j, 27) = Trim(rsx![Mati] + Cell27)
                End With
            End If
            rsx.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select * from RL1_8 where TglPendaftaran between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'or tglpendaftaran is null"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF

            If dbRst![JenisPelayanan] = "Penyakit Dalam" Then
                j = 10
            ElseIf dbRst![JenisPelayanan] = "Neonatal" Then
                j = 12
            ElseIf dbRst![JenisPelayanan] = "Lain-lain" Then
                j = 13
            ElseIf dbRst![JenisPelayanan] = "Obstetri dan Ginekologi" Then
                j = 14
            ElseIf dbRst![JenisPelayanan] = "Saraf" Then
                j = 15
            ElseIf dbRst![JenisPelayanan] = "Jiwa" Then
                j = 16
            End If

            Cell10 = oSheet.Cells(j, 26).value

            If dbRst![JenisPelayanan] = "Penyakit Dalam" Then
                With oSheet
                    .Cells(j, 26) = Trim(dbRst![JmlKunjungan] + Cell10)
                End With
            ElseIf dbRst![JenisPelayanan] = "Neonatal" Then
                With oSheet
                    .Cells(j, 26) = Trim(dbRst![JmlKunjungan] + Cell10)
                End With
            ElseIf dbRst![JenisPelayanan] = "Lain-lain" Then
                With oSheet
                    .Cells(j, 26) = Trim(dbRst![JmlKunjungan] + Cell10)
                End With
            ElseIf dbRst![JenisPelayanan] = "Obstetri dan Ginekologi" Then
                With oSheet
                    .Cells(j, 26) = Trim(dbRst![JmlKunjungan] + Cell10)
                End With
            ElseIf dbRst![JenisPelayanan] = "Saraf" Then
                With oSheet
                    .Cells(j, 26) = Trim(dbRst![JmlKunjungan] + Cell10)
                End With
            ElseIf dbRst![JenisPelayanan] = "Jiwa" Then
                With oSheet
                    .Cells(j, 26) = Trim(dbRst![JmlKunjungan] + Cell10)
                End With
            End If
            dbRst.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select * from RL1_6 where TglMasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'or tglmasuk is null"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF

            If dbRst![JenisPelayanan] = "Psikotest" Then
                j = 10
            ElseIf dbRst![JenisPelayanan] = "Konsultasi" Then
                j = 11
            ElseIf dbRst![JenisPelayanan] = "Terapi Medikamentosa" Then
                j = 12
            ElseIf dbRst![JenisPelayanan] = "Elektro Medik" Then
                j = 13
            ElseIf dbRst![JenisPelayanan] = "Psikoterapi" Then
                j = 14
            ElseIf dbRst![JenisPelayanan] = "Play Therapy" Then
                j = 15
            ElseIf dbRst![JenisPelayanan] = "Rehabilitasi Medik Psikiatrik" Then
                j = 16
            End If

            If oSheet.Cells(j, 22).value = "" Then Cell22 = 0 Else Cell22 = oSheet.Cells(j, 22).value

            If dbRst![JenisPelayanan] = "Psikotest" Then
                With oSheet
                    .Cells(j, 22) = Trim(dbRst![JmlKunjungan] + Cell22)
                End With
            ElseIf dbRst![JenisPelayanan] = "Konsultasi" Then
                With oSheet
                    .Cells(j, 22) = Trim(dbRst![JmlKunjungan] + Cell22)
                End With
            ElseIf dbRst![JenisPelayanan] = "Terapi Medikamentosa" Then
                With oSheet
                    .Cells(j, 22) = Trim(dbRst![JmlKunjungan] + Cell22)
                End With
            ElseIf dbRst![JenisPelayanan] = "Elektro Medik" Then
                With oSheet
                    .Cells(j, 22) = Trim(dbRst![JmlKunjungan] + Cell22)
                End With
            ElseIf dbRst![JenisPelayanan] = "Psikoterapi" Then
                With oSheet
                    .Cells(j, 22) = Trim(dbRst![JmlKunjungan] + Cell22)
                End With
            ElseIf dbRst![JenisPelayanan] = "Play Therapy" Then
                With oSheet
                    .Cells(j, 22) = Trim(dbRst![JmlKunjungan] + Cell22)
                End With
            ElseIf dbRst![JenisPelayanan] = "Rehabilitasi Medik Psikiatrik" Then
                With oSheet
                    .Cells(j, 22) = Trim(dbRst![JmlKunjungan] + Cell22)
                End With
            End If
            dbRst.MoveNext
        Wend
    End If

    Set rsx = Nothing
    strSQL = "Select distinct * from RL1_5 where Tglpelayanan between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'or tglpelayanan is null"
    Call msubRecFO(rsx, strSQL)

    If rsx.RecordCount > 0 Then
        rsx.MoveFirst

        While Not rsx.EOF

            If rsx!Spesialisasi = "Bedah" Then
                j = 39
            ElseIf rsx!Spesialisasi = "Obstetrik &Ginekologi" Then
                j = 40
            ElseIf rsx!Spesialisasi = "Bedah Saraf" Then
                j = 41
            ElseIf rsx!Spesialisasi = "THT" Then
                j = 42
            ElseIf rsx!Spesialisasi = "Mata" Then
                j = 43
            ElseIf rsx!Spesialisasi = "Kulit & Kelamin" Then
                j = 44
            ElseIf rsx!Spesialisasi = "Gigi & Mulut" Then
                j = 45
            ElseIf rsx!Spesialisasi = "Kardiologi" Then
                j = 46
            ElseIf rsx!Spesialisasi = "Ortopedi" Then
                j = 47
            ElseIf rsx!Spesialisasi = "Paru-Paru" Then
                j = 48
            ElseIf rsx!Spesialisasi = "Lain-lain" Then
                j = 49
            End If

            If oSheet.Cells(j, 13).value = "" Then Cell13 = 0 Else Cell13 = oSheet.Cells(j, 13).value
            If oSheet.Cells(j, 14).value = "" Then Cell14 = 0 Else Cell14 = oSheet.Cells(j, 14).value
            If oSheet.Cells(j, 15).value = "" Then Cell15 = 0 Else Cell15 = oSheet.Cells(j, 15).value
            If oSheet.Cells(j, 16).value = "" Then Cell16 = 0 Else Cell16 = oSheet.Cells(j, 16).value
            If oSheet.Cells(j, 17).value = "" Then Cell17 = 0 Else Cell17 = oSheet.Cells(j, 17).value
            If oSheet.Cells(j, 18).value = "" Then Cell18 = 0 Else Cell18 = oSheet.Cells(j, 18).value
            If oSheet.Cells(j, 19).value = "" Then Cell19 = 0 Else Cell19 = oSheet.Cells(j, 19).value
            If oSheet.Cells(j, 20).value = "" Then Cell20 = 0 Else Cell20 = oSheet.Cells(j, 20).value
            If oSheet.Cells(j, 21).value = "" Then Cell21 = 0 Else Cell21 = oSheet.Cells(j, 21).value

            If rsx!Spesialisasi = "Bedah" Then
                With oSheet
                    .Cells(j, 13) = Trim(rsx![KhususBD] + Cell13)
                    .Cells(j, 14) = Trim(rsx![KhususGD] + Cell14)
                    .Cells(j, 15) = Trim(rsx![BesarBD] + Cell15)
                    .Cells(j, 16) = Trim(rsx![BesarGD] + Cell16)
                    .Cells(j, 17) = Trim(rsx![SedangBD] + Cell17)
                    .Cells(j, 18) = Trim(rsx![SedangGD] + Cell18)
                    .Cells(j, 19) = Trim(rsx![KecilBD] + Cell19)
                    .Cells(j, 20) = Trim(rsx![KecilGD] + Cell20)
                    .Cells(j, 21) = Trim(rsx![KecilPoli] + Cell21)
                End With
            ElseIf rsx!Spesialisasi = "Obstetrik &Ginekologi" Then
                With oSheet
                    .Cells(j, 13) = Trim(rsx![KhususBD] + Cell13)
                    .Cells(j, 14) = Trim(rsx![KhususGD] + Cell14)
                    .Cells(j, 15) = Trim(rsx![BesarBD] + Cell15)
                    .Cells(j, 16) = Trim(rsx![BesarGD] + Cell16)
                    .Cells(j, 17) = Trim(rsx![SedangBD] + Cell17)
                    .Cells(j, 18) = Trim(rsx![SedangGD] + Cell18)
                    .Cells(j, 19) = Trim(rsx![KecilBD] + Cell19)
                    .Cells(j, 20) = Trim(rsx![KecilGD] + Cell20)
                    .Cells(j, 21) = Trim(rsx![KecilPoli] + Cell21)
                End With
            ElseIf rsx!Spesialisasi = "Bedah Saraf" Then
                With oSheet
                    .Cells(j, 13) = Trim(rsx![KhususBD] + Cell13)
                    .Cells(j, 14) = Trim(rsx![KhususGD] + Cell14)
                    .Cells(j, 15) = Trim(rsx![BesarBD] + Cell15)
                    .Cells(j, 16) = Trim(rsx![BesarGD] + Cell16)
                    .Cells(j, 17) = Trim(rsx![SedangBD] + Cell17)
                    .Cells(j, 18) = Trim(rsx![SedangGD] + Cell18)
                    .Cells(j, 19) = Trim(rsx![KecilBD] + Cell19)
                    .Cells(j, 20) = Trim(rsx![KecilGD] + Cell20)
                    .Cells(j, 21) = Trim(rsx![KecilPoli] + Cell21)
                End With
            ElseIf rsx!Spesialisasi = "THT" Then
                With oSheet
                    .Cells(j, 13) = Trim(rsx![KhususBD] + Cell13)
                    .Cells(j, 14) = Trim(rsx![KhususGD] + Cell14)
                    .Cells(j, 15) = Trim(rsx![BesarBD] + Cell15)
                    .Cells(j, 16) = Trim(rsx![BesarGD] + Cell16)
                    .Cells(j, 17) = Trim(rsx![SedangBD] + Cell17)
                    .Cells(j, 18) = Trim(rsx![SedangGD] + Cell18)
                    .Cells(j, 19) = Trim(rsx![KecilBD] + Cell19)
                    .Cells(j, 20) = Trim(rsx![KecilGD] + Cell20)
                    .Cells(j, 21) = Trim(rsx![KecilPoli] + Cell21)
                End With
            ElseIf rsx!Spesialisasi = "Mata" Then
                With oSheet
                    .Cells(j, 13) = Trim(rsx![KhususBD] + Cell13)
                    .Cells(j, 14) = Trim(rsx![KhususGD] + Cell14)
                    .Cells(j, 15) = Trim(rsx![BesarBD] + Cell15)
                    .Cells(j, 16) = Trim(rsx![BesarGD] + Cell16)
                    .Cells(j, 17) = Trim(rsx![SedangBD] + Cell17)
                    .Cells(j, 18) = Trim(rsx![SedangGD] + Cell18)
                    .Cells(j, 19) = Trim(rsx![KecilBD] + Cell19)
                    .Cells(j, 20) = Trim(rsx![KecilGD] + Cell20)
                    .Cells(j, 21) = Trim(rsx![KecilPoli] + Cell21)
                End With
            ElseIf rsx!Spesialisasi = "Kulit & Kelamin" Then
                With oSheet
                    .Cells(j, 13) = Trim(rsx![KhususBD] + Cell13)
                    .Cells(j, 14) = Trim(rsx![KhususGD] + Cell14)
                    .Cells(j, 15) = Trim(rsx![BesarBD] + Cell15)
                    .Cells(j, 16) = Trim(rsx![BesarGD] + Cell16)
                    .Cells(j, 17) = Trim(rsx![SedangBD] + Cell17)
                    .Cells(j, 18) = Trim(rsx![SedangGD] + Cell18)
                    .Cells(j, 19) = Trim(rsx![KecilBD] + Cell19)
                    .Cells(j, 20) = Trim(rsx![KecilGD] + Cell20)
                    .Cells(j, 21) = Trim(rsx![KecilPoli] + Cell21)
                End With
            ElseIf rsx!Spesialisasi = "Gigi & Mulut" Then
                With oSheet
                    .Cells(j, 13) = Trim(rsx![KhususBD] + Cell13)
                    .Cells(j, 14) = Trim(rsx![KhususGD] + Cell14)
                    .Cells(j, 15) = Trim(rsx![BesarBD] + Cell15)
                    .Cells(j, 16) = Trim(rsx![BesarGD] + Cell16)
                    .Cells(j, 17) = Trim(rsx![SedangBD] + Cell17)
                    .Cells(j, 18) = Trim(rsx![SedangGD] + Cell18)
                    .Cells(j, 19) = Trim(rsx![KecilBD] + Cell19)
                    .Cells(j, 20) = Trim(rsx![KecilGD] + Cell20)
                    .Cells(j, 21) = Trim(rsx![KecilPoli] + Cell21)
                End With
            ElseIf rsx!Spesialisasi = "Kardiologi" Then
                With oSheet
                    .Cells(j, 13) = Trim(rsx![KhususBD] + Cell13)
                    .Cells(j, 14) = Trim(rsx![KhususGD] + Cell14)
                    .Cells(j, 15) = Trim(rsx![BesarBD] + Cell15)
                    .Cells(j, 16) = Trim(rsx![BesarGD] + Cell16)
                    .Cells(j, 17) = Trim(rsx![SedangBD] + Cell17)
                    .Cells(j, 18) = Trim(rsx![SedangGD] + Cell18)
                    .Cells(j, 19) = Trim(rsx![KecilBD] + Cell19)
                    .Cells(j, 20) = Trim(rsx![KecilGD] + Cell20)
                    .Cells(j, 21) = Trim(rsx![KecilPoli] + Cell21)
                End With
            ElseIf rsx!Spesialisasi = "Ortopedi" Then
                With oSheet
                    .Cells(j, 13) = Trim(rsx![KhususBD] + Cell13)
                    .Cells(j, 14) = Trim(rsx![KhususGD] + Cell14)
                    .Cells(j, 15) = Trim(rsx![BesarBD] + Cell15)
                    .Cells(j, 16) = Trim(rsx![BesarGD] + Cell16)
                    .Cells(j, 17) = Trim(rsx![SedangBD] + Cell17)
                    .Cells(j, 18) = Trim(rsx![SedangGD] + Cell18)
                    .Cells(j, 19) = Trim(rsx![KecilBD] + Cell19)
                    .Cells(j, 20) = Trim(rsx![KecilGD] + Cell20)
                    .Cells(j, 21) = Trim(rsx![KecilPoli] + Cell21)
                End With
            ElseIf rsx!Spesialisasi = "Paru-Paru" Then
                With oSheet
                    .Cells(j, 13) = Trim(rsx![KhususBD] + Cell13)
                    .Cells(j, 14) = Trim(rsx![KhususGD] + Cell14)
                    .Cells(j, 15) = Trim(rsx![BesarBD] + Cell15)
                    .Cells(j, 16) = Trim(rsx![BesarGD] + Cell16)
                    .Cells(j, 17) = Trim(rsx![SedangBD] + Cell17)
                    .Cells(j, 18) = Trim(rsx![SedangGD] + Cell18)
                    .Cells(j, 19) = Trim(rsx![KecilBD] + Cell19)
                    .Cells(j, 20) = Trim(rsx![KecilGD] + Cell20)
                    .Cells(j, 21) = Trim(rsx![KecilPoli] + Cell21)
                End With
            ElseIf rsx!Spesialisasi = "Lain-lain" Then
                With oSheet
                    .Cells(j, 13) = Trim(rsx![KhususBD] + Cell13)
                    .Cells(j, 14) = Trim(rsx![KhususGD] + Cell14)
                    .Cells(j, 15) = Trim(rsx![BesarBD] + Cell15)
                    .Cells(j, 16) = Trim(rsx![BesarGD] + Cell16)
                    .Cells(j, 17) = Trim(rsx![SedangBD] + Cell17)
                    .Cells(j, 18) = Trim(rsx![SedangGD] + Cell18)
                    .Cells(j, 19) = Trim(rsx![KecilBD] + Cell19)
                    .Cells(j, 20) = Trim(rsx![KecilGD] + Cell20)
                    .Cells(j, 21) = Trim(rsx![KecilPoli] + Cell21)
                End With
            End If
            rsx.MoveNext
        Wend
    End If

    Set rsx = Nothing
    strSQL = "Select * from RL1_4a where Tglmasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'or tglmasuk is null"
    Call msubRecFO(rsx, strSQL)

    If rsx.RecordCount > 0 Then
        rsx.MoveFirst

        While Not rsx.EOF

            If rsx!TindakanMedis = "Persalinan Normal" Then
                j = 11
            ElseIf rsx!TindakanMedis = "Perd Sbl Persalinan" Then
                j = 13
            ElseIf rsx!TindakanMedis = "Perd Sdh Persalinan" Then
                j = 14
            ElseIf rsx!TindakanMedis = "Pre Eclampsi" Then
                j = 15
            ElseIf rsx!TindakanMedis = "Eclampsi" Then
                j = 16
            ElseIf rsx!TindakanMedis = "Infeksi" Then
                j = 17
            ElseIf rsx!TindakanMedis = "Lain-lain" Then
                j = 18
            ElseIf rsx!TindakanMedis = "Sectio Caesaria" Then
                j = 19
            ElseIf rsx!TindakanMedis = "Abortus" Then
                j = 20

            End If

            If oSheet.Cells(j, 12).value = "" Then Cell12 = 0 Else Cell12 = oSheet.Cells(j, 12).value
            If oSheet.Cells(j, 13).value = "" Then Cell13 = 0 Else Cell13 = oSheet.Cells(j, 13).value
            If oSheet.Cells(j, 14).value = "" Then Cell14 = 0 Else Cell14 = oSheet.Cells(j, 14).value
            If oSheet.Cells(j, 15).value = "" Then Cell15 = 0 Else Cell15 = oSheet.Cells(j, 15).value
            If oSheet.Cells(j, 16).value = "" Then Cell16 = 0 Else Cell16 = oSheet.Cells(j, 16).value
            If oSheet.Cells(j, 17).value = "" Then Cell17 = 0 Else Cell17 = oSheet.Cells(j, 17).value
            If oSheet.Cells(j, 18).value = "" Then Cell18 = 0 Else Cell18 = oSheet.Cells(j, 18).value

            If rsx!TindakanMedis = "Persalinan Normal" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    .Cells(j, 14) = rsx!JmlRujukan + Cell14
                    .Cells(j, 15) = rsx!MatiRujukan + Cell15
                    .Cells(j, 17) = rsx!MatiNonRujukan + Cell17
                    .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell18)
                End With
            ElseIf rsx!TindakanMedis = "Perd Sbl Persalinan" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    .Cells(j, 14) = rsx!JmlRujukan + Cell14
                    .Cells(j, 15) = rsx!MatiRujukan + Cell15
                    .Cells(j, 17) = rsx!MatiNonRujukan + Cell17
                    .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell18)
                End With
            ElseIf rsx!TindakanMedis = "Perd Sdh Persalinan" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    .Cells(j, 14) = rsx!JmlRujukan + Cell14
                    .Cells(j, 15) = rsx!MatiRujukan + Cell15
                    .Cells(j, 17) = rsx!MatiNonRujukan + Cell17
                    .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell18)
                End With
            ElseIf rsx!TindakanMedis = "Pre Eclampsi" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    .Cells(j, 14) = rsx!JmlRujukan + Cell14
                    .Cells(j, 15) = rsx!MatiRujukan + Cell15
                    .Cells(j, 17) = rsx!MatiNonRujukan + Cell17
                    .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell18)
                End With
            ElseIf rsx!TindakanMedis = "Eclampsi" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    .Cells(j, 14) = rsx!JmlRujukan + Cell14
                    .Cells(j, 15) = rsx!MatiRujukan + Cell15
                    .Cells(j, 17) = rsx!MatiNonRujukan + Cell17
                    .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell18)
                End With
            ElseIf rsx!TindakanMedis = "Infeksi" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    .Cells(j, 14) = rsx!JmlRujukan + Cell14
                    .Cells(j, 15) = rsx!MatiRujukan + Cell15
                    .Cells(j, 17) = rsx!MatiNonRujukan + Cell17
                    .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell18)
                End With
            ElseIf rsx!TindakanMedis = "Lain-lain" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    .Cells(j, 14) = rsx!JmlRujukan + Cell14
                    .Cells(j, 15) = rsx!MatiRujukan + Cell15
                    .Cells(j, 17) = rsx!MatiNonRujukan + Cell17
                    .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell18)
                End With
            ElseIf rsx!TindakanMedis = "Sectio Caesaria" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    .Cells(j, 14) = rsx!JmlRujukan + Cell14
                    .Cells(j, 15) = rsx!MatiRujukan + Cell15
                    .Cells(j, 17) = rsx!MatiNonRujukan + Cell17
                    .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell18)
                End With
            ElseIf rsx!TindakanMedis = "Abortus" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    .Cells(j, 14) = rsx!JmlRujukan + Cell14
                    .Cells(j, 15) = rsx!MatiRujukan + Cell15
                    .Cells(j, 17) = rsx!MatiNonRujukan + Cell17
                    .Cells(j, 18) = Trim(rsx![RujukAtas] + Cell18)
                End With

            End If
            rsx.MoveNext
        Wend
    End If

    Set rsx = Nothing
    strSQL = "Select * from RL1_4b where Tglmasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'or tglmasuk is null"
    Call msubRecFO(rsx, strSQL)

    If rsx.RecordCount > 0 Then
        rsx.MoveFirst

        While Not rsx.EOF

            If rsx!KeadaanLahirBayi = "Lahir Mati" Then
                j = 22
            End If

            If oSheet.Cells(j, 12).value = "" Then Cell12 = 0 Else Cell12 = oSheet.Cells(j, 12).value
            If oSheet.Cells(j, 13).value = "" Then Cell13 = 0 Else Cell13 = oSheet.Cells(j, 13).value
            If oSheet.Cells(j, 14).value = "" Then Cell14 = 0 Else Cell14 = oSheet.Cells(j, 14).value

            If rsx!KeadaanLahirBayi = "Lahir Mati" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    If rsx!StatusRujukan = "Rujukan" Then
                        .Cells(j, 14) = 1 + Cell14
                    End If

                End With

            End If
            rsx.MoveNext
        Wend
    End If

    Set rsx = Nothing
    strSQL = "Select * from RL1_4c where Tglmeninggal between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "' or tglmeninggal is null"
    Call msubRecFO(rsx, strSQL)

    If rsx.RecordCount > 0 Then
        rsx.MoveFirst

        While Not rsx.EOF

            If rsx!PenyebabKematian = "Asphyxia" Then
                j = 25
            ElseIf rsx!PenyebabKematian = "TraumaKelahiran" Then
                j = 26
            ElseIf rsx!PenyebabKematian = "BBLR" Then
                j = 27
            ElseIf rsx!PenyebabKematian = "Tetanus Neonatarum" Then
                j = 28
            ElseIf rsx!PenyebabKematian = "Kelainan Cogenital" Then
                j = 29
            ElseIf rsx!PenyebabKematian = "ISPA" Then
                j = 30
            ElseIf rsx!PenyebabKematian = "Diare" Then
                j = 31
            ElseIf rsx!PenyebabKematian = "Lain-lain" Then
                j = 32
            End If

            If oSheet.Cells(j, 12).value = "" Then Cell12 = 0 Else Cell12 = oSheet.Cells(j, 12).value
            If oSheet.Cells(j, 13).value = "" Then Cell13 = 0 Else Cell13 = oSheet.Cells(j, 13).value
            If oSheet.Cells(j, 14).value = "" Then Cell14 = 0 Else Cell14 = oSheet.Cells(j, 14).value

            If rsx!PenyebabKematian = "Asphyxia" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    If rsx!StatusRujukan = "Rujukan" Then
                        .Cells(j, 14) = 1 + Cell14
                    End If

                End With
            ElseIf rsx!PenyebabKematian = "TraumaKelahiran" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    If rsx!StatusRujukan = "Rujukan" Then
                        .Cells(j, 14) = 1 + Cell14
                    End If

                End With
            ElseIf rsx!PenyebabKematian = "BBLR" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    If rsx!StatusRujukan = "Rujukan" Then
                        .Cells(j, 14) = 1 + Cell14
                    End If

                End With
            ElseIf rsx!PenyebabKematian = "Tetanus Neonatarum" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    If rsx!StatusRujukan = "Rujukan" Then
                        .Cells(j, 14) = 1 + Cell14
                    End If

                End With
            ElseIf rsx!PenyebabKematian = "Kelainan Cogenital" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    If rsx!StatusRujukan = "Rujukan" Then
                        .Cells(j, 14) = 1 + Cell14
                    End If

                End With
            ElseIf rsx!PenyebabKematian = "ISPA" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    If rsx!StatusRujukan = "Rujukan" Then
                        .Cells(j, 14) = 1 + Cell14
                    End If

                End With
            ElseIf rsx!PenyebabKematian = "Diare" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    If rsx!StatusRujukan = "Rujukan" Then
                        .Cells(j, 14) = 1 + Cell14
                    End If

                End With
            ElseIf rsx!PenyebabKematian = "Lain-lain" Then
                With oSheet
                    .Cells(j, 12) = Trim(rsx![>2500] + Cell12)
                    .Cells(j, 13) = Trim(rsx![<2500] + Cell13)
                    If rsx!StatusRujukan = "Rujukan" Then
                        .Cells(j, 14) = 1 + Cell14
                    End If

                End With
            End If
            rsx.MoveNext
        Wend
    End If

    Set rsx = Nothing
    strSQL = "Select * from RL1_4d where Tglmasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "' or tglmasuk is null"
    Call msubRecFO(rsx, strSQL)

    If rsx.RecordCount > 0 Then
        rsx.MoveFirst

        While Not rsx.EOF

            If rsx!NamaImunisasi = "TT1" Then
                j = 33
            ElseIf rsx!NamaImunisasi = "TT2" Then
                j = 34
            End If

            If oSheet.Cells(j, 14).value = "" Then Cell14 = 0 Else Cell14 = oSheet.Cells(j, 14).value
            If oSheet.Cells(j, 16).value = "" Then Cell16 = 0 Else Cell16 = oSheet.Cells(j, 16).value

            If rsx!NamaImunisasi = "TT1" Then
                With oSheet

                    If rsx!StatusRujukan = "Rujukan" Then
                        .Cells(j, 14) = 1 + Cell14
                    ElseIf rsx!StatusRujukan = "Non Rujukan" Then
                        .Cells(j, 16) = 1 + Cell16
                    End If

                End With
            ElseIf rsx!NamaImunisasi = "TT2" Then
                With oSheet

                    If rsx!StatusRujukan = "Rujukan" Then
                        .Cells(j, 14) = 1 + Cell14
                    ElseIf rsx!StatusRujukan = "Non Rujukan" Then
                        .Cells(j, 16) = 1 + Cell16
                    End If

                End With

            End If
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
