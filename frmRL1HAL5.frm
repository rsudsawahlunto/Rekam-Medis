VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRL1HAL5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL1 Halaman 5"
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
   Icon            =   "frmRL1HAL5.frx":0000
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
         Format          =   115146755
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
         Format          =   115146755
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
         Format          =   115146755
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
      Picture         =   "frmRL1HAL5.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2955
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRL1HAL5.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL1HAL5.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmRL1HAL5"
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
Dim Cell3 As String
Dim Cell4 As String
Dim Cell5 As String
Dim Cell6 As String
Dim Cell7 As String
Dim Cell8 As String
Dim Cell11 As String
Dim Cell12 As String
Dim Cell13 As String
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
    Set oWB = oXL.Workbooks.Open(App.Path & "\RL1 Hal5.xls")
    Set oSheet = oWB.ActiveSheet

    Set rsb = Nothing
    strSQL = "select * from profilrs"
    Call msubRecFO(rsb, strSQL)

    Set oResizeRange = oSheet.Range("d1", "d2")
    oResizeRange.value = Trim(rsb!KdRs)

    Set rsx = Nothing
    strSQL = "Select * from RL1_22 where Tglperiksa between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'or tglperiksa is null"
    Call msubRecFO(rsx, strSQL)

    If rsx.RecordCount > 0 Then
        rsx.MoveFirst

        While Not rsx.EOF
            If rsx!MetodologiBayiTabung = "Konvensional" Then
                j = 46
            ElseIf rsx!MetodologiBayiTabung = "ICSI" Then
                j = 47
            ElseIf rsx!MetodologiBayiTabung = "MESA" Then
                j = 48
            ElseIf rsx!MetodologiBayiTabung = "TESA" Then
                j = 48
            ElseIf rsx!MetodologiBayiTabung = "TESE" Then
                j = 48
            ElseIf rsx!MetodologiBayiTabung = "Embrio Beku" Then
                j = 49
            ElseIf rsx!MetodologiBayiTabung = "Lainnya" Then
                j = 50
            End If

            Cell11 = oSheet.Cells(j, 11).value
            If oSheet.Cells(j, 12).value = "" Then Cell12 = 0 Else Cell12 = oSheet.Cells(j, 12).value

            If rsx!MetodologiBayiTabung = "Konvensional" Then
                With oSheet
                    .Cells(j, 11) = Trim(rsx![SiklusPengobatanBT] + Cell11)
                    .Cells(j, 12) = (rsx![Kehamilan+] + Cell12)
                    .Cells(j, 13) = (rsx![Persentase])
                End With
            ElseIf rsx!MetodologiBayiTabung = "ICSI" Then
                With oSheet
                    .Cells(j, 11) = Trim(rsx![SiklusPengobatanBT] + Cell11)
                    .Cells(j, 12) = Trim(rsx![Kehamilan+] + Cell12)
                    .Cells(j, 13) = (rsx![Persentase])
                End With
            ElseIf rsx!MetodologiBayiTabung = "MESA" Then
                With oSheet
                    .Cells(j, 11) = Trim(rsx![SiklusPengobatanBT] + Cell11)
                    .Cells(j, 12) = Trim(rsx![Kehamilan+] + Cell12)
                    .Cells(j, 13) = (rsx![Persentase])
                End With
            ElseIf rsx!MetodologiBayiTabung = "TESA" Then
                With oSheet
                    .Cells(j, 11) = Trim(rsx![SiklusPengobatanBT] + Cell11)
                    .Cells(j, 12) = Trim(rsx![Kehamilan+] + Cell12)
                    .Cells(j, 13) = (rsx![Persentase])
                End With
            ElseIf rsx!MetodologiBayiTabung = "TESE" Then
                With oSheet
                    .Cells(j, 11) = Trim(rsx![SiklusPengobatanBT] + Cell11)
                    .Cells(j, 12) = Trim(rsx![Kehamilan+] + Cell12)
                    .Cells(j, 13) = (rsx![Persentase])
                End With
            ElseIf rsx!MetodologiBayiTabung = "Embrio Beku" Then
                With oSheet
                    .Cells(j, 11) = Trim(rsx![SiklusPengobatanBT] + Cell11)
                    .Cells(j, 12) = Trim(rsx![Kehamilan+] + Cell12)
                    .Cells(j, 13) = (rsx![Persentase])
                End With
            ElseIf rsx!MetodologiBayiTabung = "Lainnya" Then
                With oSheet
                    .Cells(j, 11) = Trim(rsx![SiklusPengobatanBT] + Cell11)
                    .Cells(j, 12) = Trim(rsx![Kehamilan+] + Cell12)
                    .Cells(j, 13) = (rsx![Persentase])
                End With
            End If
            rsx.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select distinct * from RL1_19 where Tglmulai between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'or tglmulai is null"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF
            If dbRst!JenisDiklat = "Teknis" Then
                j = 7
            ElseIf dbRst!JenisDiklat = "Teknis Fungsional" Then
                j = 11
            End If

            If oSheet.Cells(j, 3).value = "" Then Cell3 = 0 Else Cell3 = oSheet.Cells(j, 3).value
            If oSheet.Cells(j, 4).value = "" Then Cell4 = 0 Else Cell4 = oSheet.Cells(j, 4).value
            If oSheet.Cells(j, 5).value = "" Then Cell5 = 0 Else Cell5 = oSheet.Cells(j, 5).value
            If oSheet.Cells(j, 6).value = "" Then Cell6 = 0 Else Cell6 = oSheet.Cells(j, 6).value
            If oSheet.Cells(j, 7).value = "" Then Cell7 = 0 Else Cell7 = oSheet.Cells(j, 7).value
            If oSheet.Cells(j, 8).value = "" Then Cell8 = 0 Else Cell8 = oSheet.Cells(j, 8).value

            If dbRst!JenisDiklat = "Teknis" Then
                With oSheet
                    If dbRst!AsalPeserta = "Rumah Sakit Sendiri" Then
                        If dbRst!JenisTenaga = "Dokter" Then
                            .Cells(j, 3) = 1 + Cell3
                        ElseIf dbRst!JenisTenaga = "NonKesehatan" Then
                            .Cells(j, 5) = 1 + Cell5
                        ElseIf dbRst!JenisTenaga = "KesehatanLain" Then
                            .Cells(j, 4) = 1 + Cell4
                        End If
                    ElseIf dbRst!AsalPeserta = "Instansi/ Rumah Sakit Lain" Then
                        If dbRst!JenisTenaga = "Dokter" Then
                            .Cells(j, 6) = 1 + Cell6
                        ElseIf dbRst!JenisTenaga = "NonKesehatan" Then
                            .Cells(j, 8) = 1 + Cell8
                        ElseIf dbRst!JenisTenaga = "KesehatanLain" Then
                            .Cells(j, 7) = 1 + Cell7
                        End If
                    End If
                End With
            ElseIf dbRst!JenisDiklat = "Teknis Fungsional" Then
                With oSheet
                    If dbRst!AsalPeserta = "Rumah Sakit Sendiri" Then
                        If dbRst!JenisTenaga = "Dokter" Then
                            .Cells(j, 3) = 1 + Cell3
                        ElseIf dbRst!JenisTenaga = "NonKesehatan" Then
                            .Cells(j, 5) = 1 + Cell5
                        ElseIf dbRst!JenisTenaga = "KesehatanLain" Then
                            .Cells(j, 4) = 1 + Cell4
                        End If
                    ElseIf dbRst!AsalPeserta = "Instansi/ Rumah Sakit Lain" Then
                        If dbRst!JenisTenaga = "Dokter" Then
                            .Cells(j, 6) = 1 + Cell6
                        ElseIf dbRst!JenisTenaga = "NonKesehatan" Then
                            .Cells(j, 8) = 1 + Cell8
                        ElseIf dbRst!JenisTenaga = "KesehatanLain" Then
                            .Cells(j, 7) = 1 + Cell7
                        End If
                    End If
                End With
            End If
            dbRst.MoveNext
        Wend
    End If

    Set dbRst = Nothing
    strSQL = "Select distinct * from RL1_20 where Tglmulaiperiksa between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'or tglmulaiperiksa is null"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF
            If dbRst!KdTindakanMedis = "1901" Then
                j = 20
            ElseIf dbRst!KdTindakanMedis = "1902" Then
                j = 21
            ElseIf dbRst!KdTindakanMedis = "1903" Then
                j = 22

            ElseIf dbRst!KdTindakanMedis = "1904)" Then
                j = 29
            ElseIf dbRst!KdTindakanMedis = "1905" Then
                j = 30

            ElseIf dbRst!KdTindakanMedis = "1906" Then
                j = 37
            ElseIf dbRst!KdTindakanMedis = "1907" Then
                j = 38

            ElseIf dbRst!KdTindakanMedis = "1304" Then
                j = 43
            ElseIf dbRst!KdTindakanMedis = "1908" Then
                j = 44
            ElseIf dbRst!KdTindakanMedis = "1909" Then
                j = 45
            ElseIf dbRst!KdTindakanMedis = "1910" Then
                j = 46
            ElseIf dbRst!KdTindakanMedis = "1911" Then
                j = 47
            ElseIf dbRst!KdTindakanMedis = "1912" Then
                j = 48
            ElseIf dbRst!KdTindakanMedis = "1913" Then
                j = 49
            ElseIf dbRst!KdTindakanMedis = "1914" Then
                j = 50
            ElseIf dbRst!KdTindakanMedis = "1915" Then
                j = 51
            ElseIf dbRst!KdTindakanMedis = "1916" Then
                j = 52
            ElseIf dbRst!KdTindakanMedis = "1917" Then
                j = 53
            ElseIf dbRst!KdTindakanMedis = "1918" Then
                j = 54
            ElseIf dbRst!KdTindakanMedis = "1919" Then
                j = 55
            ElseIf dbRst!KdTindakanMedis = "1920" Then
                j = 56
            End If

            If oSheet.Cells(j, 3).value = "" Then Cell3 = 0 Else Cell3 = oSheet.Cells(j, 3).value
            If oSheet.Cells(j, 4).value = "" Then Cell4 = 0 Else Cell4 = oSheet.Cells(j, 4).value
            If oSheet.Cells(j, 5).value = "" Then Cell5 = 0 Else Cell5 = oSheet.Cells(j, 5).value
            If oSheet.Cells(j, 6).value = "" Then Cell6 = 0 Else Cell6 = oSheet.Cells(j, 6).value

            If dbRst!KdTindakanMedis = "1901" Then
                With oSheet
                    If dbRst!KualitasHasil = "BAIK" Then

                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 6) = dbRst!Jml + Cell6

                    ElseIf dbRst!KualitasHasil = "SEDANG" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 6) = dbRst!Jml + Cell6
                    ElseIf dbRst!KualitasHasil = "BURUK" Then
                        .Cells(j, 5) = 1 + Cell5
                        .Cells(j, 6) = dbRst!Jml + Cell6
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1902" Then
                With oSheet
                    If dbRst!KualitasHasil = "BAIK" Then

                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 6) = dbRst!Jml + Cell6

                    ElseIf dbRst!KualitasHasil = "SEDANG" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 6) = dbRst!Jml + Cell6
                    ElseIf dbRst!KualitasHasil = "BURUK" Then
                        .Cells(j, 5) = 1 + Cell5
                        .Cells(j, 6) = dbRst!Jml + Cell6
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1903" Then
                With oSheet
                    If dbRst!KualitasHasil = "BAIK" Then

                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 6) = dbRst!Jml + Cell6

                    ElseIf dbRst!KualitasHasil = "SEDANG" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 6) = dbRst!Jml + Cell6
                    ElseIf dbRst!KualitasHasil = "BURUK" Then
                        .Cells(j, 5) = 1 + Cell5
                        .Cells(j, 6) = dbRst!Jml + Cell6
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1904" Then
                With oSheet
                    If dbRst!KualitasHasil = "TURUN OBAT +" Then

                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 6) = dbRst!Jml + Cell6

                    ElseIf dbRst!KualitasHasil = "TURUN OBAT -" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 6) = dbRst!Jml + Cell6
                    ElseIf dbRst!KualitasHasil = "TETAP/MENINGKAT" Then
                        .Cells(j, 5) = 1 + Cell5
                        .Cells(j, 6) = dbRst!Jml + Cell6
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1905" Then
                With oSheet
                    If dbRst!KualitasHasil = "TURUN OBAT +" Then

                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 6) = dbRst!Jml + Cell6

                    ElseIf dbRst!KualitasHasil = "TURUN OBAT -" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 6) = dbRst!Jml + Cell6
                    ElseIf dbRst!KualitasHasil = "TETAP/MENINGKAT" Then
                        .Cells(j, 5) = 1 + Cell5
                        .Cells(j, 6) = dbRst!Jml + Cell6
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1906" Then
                With oSheet
                    If dbRst!KualitasHasil = "PERBAIKAN ANATROMIK" Then

                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 5) = dbRst!Jml + Cell5

                    ElseIf dbRst!KualitasHasil = "PERBAIKAN VISUS" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1907" Then
                With oSheet
                    If dbRst!KualitasHasil = "PERBAIKAN ANATROMIK" Then

                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 5) = dbRst!Jml + Cell5

                    ElseIf dbRst!KualitasHasil = "PERBAIKAN VISUS" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1304" Then
                With oSheet
                    If dbRst!KualitasHasil = "PERBAIKAN -" Then

                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 5) = dbRst!Jml + Cell5

                    ElseIf dbRst!KualitasHasil = "PERBAIKAN +" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1908" Then
                With oSheet
                    If dbRst!KualitasHasil = "PERBAIKAN -" Then

                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 5) = dbRst!Jml + Cell5

                    ElseIf dbRst!KualitasHasil = "PERBAIKAN +" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1909" Then
                With oSheet
                    If dbRst!KualitasHasil = "PERBAIKAN -" Then

                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 5) = dbRst!Jml + Cell5

                    ElseIf dbRst!KualitasHasil = "PERBAIKAN +" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1910" Then
                With oSheet
                    If dbRst!KualitasHasil = "PERBAIKAN -" Then

                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 5) = dbRst!Jml + Cell5

                    ElseIf dbRst!KualitasHasil = "PERBAIKAN +" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1911" Then
                With oSheet
                    If dbRst!KualitasHasil = "PERBAIKAN -" Then
                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    ElseIf dbRst!KualitasHasil = "PERBAIKAN +" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1912" Then
                With oSheet
                    If dbRst!KualitasHasil = "PERBAIKAN -" Then
                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    ElseIf dbRst!KualitasHasil = "PERBAIKAN +" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1913" Then
                With oSheet
                    If dbRst!KualitasHasil = "PERBAIKAN -" Then
                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    ElseIf dbRst!KualitasHasil = "PERBAIKAN +" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1914" Then
                With oSheet
                    If dbRst!KualitasHasil = "PERBAIKAN -" Then
                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    ElseIf dbRst!KualitasHasil = "PERBAIKAN +" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1915" Then
                With oSheet
                    If dbRst!KualitasHasil = "PERBAIKAN -" Then
                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    ElseIf dbRst!KualitasHasil = "PERBAIKAN +" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1916" Then
                With oSheet
                    If dbRst!KualitasHasil = "PERBAIKAN -" Then
                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    ElseIf dbRst!KualitasHasil = "PERBAIKAN +" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1917" Then
                With oSheet
                    If dbRst!KualitasHasil = "PERBAIKAN -" Then
                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    ElseIf dbRst!KualitasHasil = "PERBAIKAN +" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1918" Then
                With oSheet
                    If dbRst!KualitasHasil = "PERBAIKAN -" Then
                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    ElseIf dbRst!KualitasHasil = "PERBAIKAN +" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1919" Then
                With oSheet
                    If dbRst!KualitasHasil = "PERBAIKAN -" Then
                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    ElseIf dbRst!KualitasHasil = "PERBAIKAN +" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    End If
                End With
            ElseIf dbRst!KdTindakanMedis = "1920" Then
                With oSheet
                    If dbRst!KualitasHasil = "PERBAIKAN -" Then
                        .Cells(j, 3) = 1 + Cell3
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    ElseIf dbRst!KualitasHasil = "PERBAIKAN +" Then
                        .Cells(j, 4) = 1 + Cell4
                        .Cells(j, 5) = dbRst!Jml + Cell5
                    End If
                End With
            End If
            dbRst.MoveNext
        Wend
    End If
    Set rsx = Nothing
    strSQL = "Select * from RL1_21 where Tglmulaiperiksa between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'or tglmulaiperiksa is null"
    Call msubRecFO(rsx, strSQL)

    If rsx.RecordCount > 0 Then
        rsx.MoveFirst

        While Not rsx.EOF
            If rsx!KdNapza = "901001" Then
                j = 21
            ElseIf rsx!KdNapza = "901002" Then
                j = 22
            ElseIf rsx!KdNapza = "901003" Then
                j = 23
            ElseIf rsx!KdNapza = "901004" Then
                j = 24
            ElseIf rsx!KdNapza = "901005" Then
                j = 25
            ElseIf rsx!KdNapza = "901006" Then
                j = 26
            ElseIf rsx!KdNapza = "901007" Then
                j = 27

            ElseIf rsx!KdNapza = "901008" Then
                j = 30
            ElseIf rsx!KdNapza = "901009" Then
                j = 31
            ElseIf rsx!KdNapza = "901010" Then
                j = 32

            ElseIf rsx!KdNapza = "901011" Then
                j = 34
            ElseIf rsx!KdNapza = "901012" Then
                j = 35
            ElseIf rsx!KdNapza = "901013" Then
                j = 36

            ElseIf rsx!KdNapza = "901014" Then
                j = 37
            ElseIf rsx!KdNapza = "901015" Then
                j = 38

            ElseIf rsx!KdNapza = "901016" Then
                j = 40
            ElseIf rsx!KdNapza = "901017" Then
                j = 41
            End If

            If oSheet.Cells(j, 11).value = "" Then Cell11 = 0 Else Cell11 = oSheet.Cells(j, 11).value
            If oSheet.Cells(j, 12).value = "" Then Cell12 = 0 Else Cell12 = oSheet.Cells(j, 12).value
            If oSheet.Cells(j, 13).value = "" Then Cell13 = 0 Else Cell13 = oSheet.Cells(j, 13).value

            If rsx!KdNapza = "901001" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
                End With
            ElseIf rsx!KdNapza = "901002" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
                End With
            ElseIf rsx!KdNapza = "901003" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
                End With
            ElseIf rsx!KdNapza = "901004" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
                End With
            ElseIf rsx!KdNapza = "901005" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
                End With
            ElseIf rsx!KdNapza = "901006" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
                End With
            ElseIf rsx!KdNapza = "901007" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
                End With
            ElseIf rsx!KdNapza = "901008" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
                End With
            ElseIf rsx!KdNapza = "901009" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
                End With
            ElseIf rsx!KdNapza = "901010" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
                End With
            ElseIf rsx!KdNapza = "901011" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
                End With
            ElseIf rsx!KdNapza = "901012" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
                End With
            ElseIf rsx!KdNapza = "901013" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
                End With
            ElseIf rsx!KdNapza = "901014" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
                End With
            ElseIf rsx!KdNapza = "901015" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
                End With
            ElseIf rsx!KdNapza = "901016" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
                End With
            ElseIf rsx!KdNapza = "901017" Then
                With oSheet
                    .Cells(j, 11) = (rsx![Kuratif] + Cell11)
                    .Cells(j, 12) = (rsx![Rehabilitatif] + Cell12)
                    .Cells(j, 13) = (rsx![AfterCare] + Cell13)
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
