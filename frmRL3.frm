VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmRL3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3 Data Dasar Rumah Sakit"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   Icon            =   "frmRL3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5295
   Begin VB.OptionButton Option2 
      Caption         =   "Hal.2"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Hal. 1"
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   1560
      Value           =   -1  'True
      Width           =   1455
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
      Top             =   2400
      Width           =   7605
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1905
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   3120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
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
   Begin VB.Frame Frame2 
      Caption         =   "RL 3 Halaman"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   7575
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   4680
      Picture         =   "frmRL3.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2955
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRL3.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL3.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmRL3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project/reference/microsoft excel 12.0 object library
'Selalu gunakan format file excel 2003  .xls sebagai standar agar pengguna excel 2003 atau diatasnya dpt menggunakan report laporannya
'Catatan: Format excel 2000 tidak dpt mengoperasikan beberapa fungsi yg ada pada excell 2003 atau diatasnya

Option Explicit

'Special Buat Excel
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel.Range
Dim oResizeRange As Excel.Range
Dim j As String
'Special Buat Excel

Dim Cell12 As String
Dim Cell15 As String
Dim Cell18 As String
Dim Cell21 As String
Dim Cell24 As String

Private Sub cmdCetak_Click()
    On Error GoTo hell

    If Option1.value = True Then

        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\Data Dasar Rumah Sakit RL3.xls")
        Set oSheet = oWB.ActiveSheet

        Set rs = Nothing

        strSQL = "select * from RL3_Judul"
        Call msubRecFO(rs, strSQL)

        With oSheet
            .Cells(6, 7) = Trim(IIf(IsNull(rs!KdRs.value), "", (rs!KdRs.value)))
            .Cells(7, 7) = Trim(IIf(IsNull(rs!NamaRS.value), "", (rs!NamaRS.value)))
            .Cells(9, 7) = Trim(IIf(IsNull(rs!JenisProfile.value), "", (rs!JenisProfile.value)))
            .Cells(10, 7) = Trim(IIf(IsNull(rs!KelasRS.value), "", (rs!KelasRS.value)))
            .Cells(11, 7) = Trim(IIf(IsNull(rs!Direktur.value), "", (rs!Direktur.value)))
            .Cells(12, 7) = Trim(IIf(IsNull(rs!alamat.value), "", (rs!alamat.value)))
            .Cells(13, 7) = Trim(IIf(IsNull(rs!KotaKodyaKab.value), "", (rs!KotaKodyaKab.value)))
            .Cells(13, 15) = Trim(IIf(IsNull(rs!KodePos.value), "", (rs!KodePos.value)))
            .Cells(14, 7) = Trim(IIf(IsNull(rs!Telepon.value), "", (rs!Telepon.value)))
            .Cells(14, 13) = Trim(IIf(IsNull(rs!Faks.value), "", (rs!Faks.value)))
            .Cells(16, 7) = Trim(IIf(IsNull(rs!NoSuratIjinLast.value), "", (rs!NoSuratIjinLast.value)))
            .Cells(17, 7) = Trim(IIf(IsNull(rs!TglSuratIjinLast.value), "", (rs!TglSuratIjinLast.value)))
            .Cells(18, 7) = Trim(IIf(IsNull(rs!SignatureByLast.value), "", (rs!SignatureByLast.value)))

            If rs!StatusSuratIjin.value = "Sementara" Then
                oSheet.Cells(19, 7) = "V"
                oSheet.Cells(19, 12) = ""
                oSheet.Cells(19, 15) = ""
            End If
            If rs!StatusSuratIjin.value = "Tetap" Then
                oSheet.Cells(19, 7) = ""
                oSheet.Cells(19, 12) = "V"
                oSheet.Cells(19, 15) = ""
            End If
            If rs!StatusSuratIjin.value = "Perpanjangan" Then
                oSheet.Cells(19, 7) = ""
                oSheet.Cells(19, 12) = ""
                oSheet.Cells(19, 15) = "V"
            End If

            .Cells(20, 12) = Trim(IIf(IsNull(rs!MasaBerlakuIjin.value), "", (rs!MasaBerlakuIjin.value)))
            .Cells(7, 25) = Trim(IIf(IsNull(rs!PemilikProfile.value), "", (rs!PemilikProfile.value)))

            If rs!TahapanAkreditasi.value = "Pentahapan I" Then
                oSheet.Cells(15, 22) = "V"
                oSheet.Cells(15, 25) = ""
                oSheet.Cells(15, 29) = ""
            End If
            If rs!TahapanAkreditasi.value = "Pentahapan II" Then
                oSheet.Cells(15, 22) = ""
                oSheet.Cells(15, 25) = "V"
                oSheet.Cells(15, 29) = ""
            End If
            If rs!TahapanAkreditasi.value = "Pentahapan III" Then
                oSheet.Cells(15, 22) = ""
                oSheet.Cells(15, 25) = ""
                oSheet.Cells(15, 29) = "V"
            End If

            If rs!statusAkreditasi.value = "Penuh" Then
                oSheet.Cells(18, 22) = "V"
                oSheet.Cells(19, 22) = ""
                oSheet.Cells(18, 25) = ""
                oSheet.Cells(19, 25) = ""
            End If
            If rs!statusAkreditasi.value = "Gagal" Then
                oSheet.Cells(18, 22) = ""
                oSheet.Cells(19, 22) = "V"
                oSheet.Cells(18, 25) = ""
                oSheet.Cells(19, 25) = ""
            End If
            If rs!statusAkreditasi.value = "Bersyarat" Then
                oSheet.Cells(18, 22) = ""
                oSheet.Cells(19, 22) = ""
                oSheet.Cells(18, 25) = "V"
                oSheet.Cells(19, 25) = ""
            End If
            If rs!statusAkreditasi.value = "Belum" Then
                oSheet.Cells(18, 22) = ""
                oSheet.Cells(19, 22) = ""
                oSheet.Cells(18, 25) = ""
                oSheet.Cells(19, 25) = "V"
            End If
        End With

        Set rsx = Nothing

        strSQL = "select * from RL3_RI"
        Call msubRecFO(rsx, strSQL)

        If rsx.RecordCount > 0 Then
            rsx.MoveFirst

            While Not rsx.EOF

                If rsx!kdsubinstalasi = "001" Then
                    j = 26
                ElseIf rsx!kdsubinstalasi = "002" Then
                    j = 27
                ElseIf rsx!kdsubinstalasi = "003" Then
                    j = 28
                ElseIf rsx!kdsubinstalasi = "004" Then
                    j = 29
                ElseIf rsx!kdsubinstalasi = "005" Then
                    j = 30
                ElseIf rsx!kdsubinstalasi = "006" Then
                    j = 31
                ElseIf rsx!kdsubinstalasi = "007" Then
                    j = 32
                ElseIf rsx!kdsubinstalasi = "008" Then
                    j = 33
                ElseIf rsx!kdsubinstalasi = "009" Then
                    j = 34
                ElseIf rsx!kdsubinstalasi = "010" Then
                    j = 35
                ElseIf rsx!kdsubinstalasi = "011" Then
                    j = 36
                ElseIf rsx!kdsubinstalasi = "012" Then
                    j = 37
                ElseIf rsx!kdsubinstalasi = "013" Then
                    j = 38
                ElseIf rsx!kdsubinstalasi = "014" Then
                    j = 39
                ElseIf rsx!kdsubinstalasi = "015" Then
                    j = 40
                ElseIf rsx!kdsubinstalasi = "016" Then
                    j = 41
                ElseIf rsx!kdsubinstalasi = "017" Then
                    j = 42
                ElseIf rsx!kdsubinstalasi = "018" Then
                    j = 43
                ElseIf rsx!kdsubinstalasi = "019" Then
                    j = 44
                ElseIf rsx!kdsubinstalasi = "020" Then
                    j = 45
                ElseIf rsx!kdsubinstalasi = "021" Then
                    j = 46
                ElseIf rsx!kdsubinstalasi = "022" Then
                    j = 47
                ElseIf rsx!kdsubinstalasi = "023" Then
                    j = 48
                ElseIf rsx!kdsubinstalasi = "024" Then
                    j = 49
                ElseIf rsx!kdsubinstalasi = "025" Then
                    j = 50
                ElseIf rsx!kdsubinstalasi = "026" Then
                    j = 51
                ElseIf rsx!kdsubinstalasi = "027" Then
                    j = 52
                ElseIf rsx!kdsubinstalasi = "028" Then
                    j = 54
                End If

                Cell12 = oSheet.Cells(j, 12).value
                Cell15 = oSheet.Cells(j, 15).value
                Cell18 = oSheet.Cells(j, 18).value
                Cell21 = oSheet.Cells(j, 21).value
                Cell24 = oSheet.Cells(j, 24).value

                If rsx!kdsubinstalasi = "001" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "001" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "001" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "001" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "001" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "002" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "002" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "002" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "002" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "002" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "003" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "003" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "003" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "003" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "003" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "004" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "004" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "004" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "004" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "004" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "005" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "005" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "005" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "005" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "005" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "006" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "006" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "006" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "006" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "006" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "007" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "007" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "007" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "007" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "007" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "008" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "008" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "008" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "008" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "008" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "009" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "009" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "009" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "009" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "009" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "010" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "010" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "010" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "010" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "010" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "011" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "011" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "011" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "011" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "011" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "012" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "012" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "012" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "012" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "012" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "013" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "013" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "013" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "013" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "013" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "014" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "014" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "014" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "014" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "014" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "015" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "015" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "015" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "015" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "015" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "016" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "016" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "016" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "016" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "016" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "017" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "017" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "017" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "017" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "017" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "018" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "018" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "018" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "018" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "018" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "019" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "019" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "019" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "019" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "019" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "020" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "020" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "020" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "020" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "020" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "021" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "021" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "021" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "021" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "021" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "022" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "022" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "022" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "022" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "022" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "023" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "023" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "023" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "023" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "023" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "024" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "024" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "024" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "024" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "024" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "025" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "025" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "025" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "025" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "025" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "026" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "026" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "026" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "026" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "026" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "027" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "027" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "027" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "027" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "027" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk

                ElseIf rsx!kdsubinstalasi = "028" And rsx!Kelas = "Kelas Utama" Then
                    Call Setcellku
                ElseIf rsx!kdsubinstalasi = "028" And rsx!Kelas = "Kelas I" Then
                    Call setcellki
                ElseIf rsx!kdsubinstalasi = "028" And rsx!Kelas = "Kelas II" Then
                    Call setcellkii
                ElseIf rsx!kdsubinstalasi = "028" And rsx!Kelas = "Kelas III" Then
                    Call setcellkiii
                ElseIf rsx!kdsubinstalasi = "028" And rsx!Kelas = "Tanpa Kelas" Then
                    Call setcelltk
                End If
                rsx.MoveNext
            Wend
        End If

    ElseIf Option2.value = True Then

        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\Data Dasar Rumah Sakit RL3.2.xls")
        Set oSheet = oWB.ActiveSheet

        Set rs = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='01'"
        Call msubRecFO(rs, strSQL)

        While Not rs.EOF
            With oSheet
                If rs!kdpelayananprofile = "01A" Then
                    .Cells(9, 5) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "01B" Then
                    .Cells(9, 8) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "01C" Then
                    .Cells(9, 11) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "01D" Then
                    .Cells(9, 14) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "01E" Then
                    .Cells(9, 17) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "01F" Then
                    .Cells(11, 5) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "01G" Then
                    .Cells(11, 8) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "01H" Then
                    .Cells(11, 11) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "01I" Then
                    .Cells(11, 14) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "01J" Then
                    .Cells(11, 17) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "01K" Then
                    .Cells(13, 5) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "01L" Then
                    .Cells(13, 8) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                End If
                rs.MoveNext
            End With
        Wend

        Set rsx = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='02'"
        Call msubRecFO(rsx, strSQL)
        While Not rsx.EOF
            With oSheet
                If rsx!kdpelayananprofile = "02A" Then
                    .Cells(15, 5) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                ElseIf rsx!kdpelayananprofile = "02B" Then
                    .Cells(15, 8) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                ElseIf rsx!kdpelayananprofile = "02C" Then
                    .Cells(15, 11) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                ElseIf rsx!kdpelayananprofile = "02D" Then
                    .Cells(15, 14) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                ElseIf rsx!kdpelayananprofile = "02E" Then
                    .Cells(15, 17) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                ElseIf rsx!kdpelayananprofile = "02F" Then
                    .Cells(17, 5) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                ElseIf rsx!kdpelayananprofile = "02G" Then
                    .Cells(17, 8) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                ElseIf rsx!kdpelayananprofile = "02H" Then
                    .Cells(17, 11) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                ElseIf rsx!kdpelayananprofile = "02I" Then
                    .Cells(17, 14) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                End If
                rsx.MoveNext
            End With
        Wend

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='03'"
        Call msubRecFO(rsy, strSQL)
        While Not rsy.EOF
            With oSheet
                If rsy!kdpelayananprofile = "03A" Then
                    .Cells(19, 5) = Trim(IIf(IsNull(rsy!qtybukadlmminggu.value), "", (rsy!qtybukadlmminggu.value)))
                ElseIf rsy!kdpelayananprofile = "03B" Then
                    .Cells(19, 8) = Trim(IIf(IsNull(rsy!qtybukadlmminggu.value), "", (rsy!qtybukadlmminggu.value)))
                ElseIf rsy!kdpelayananprofile = "03C" Then
                    .Cells(19, 11) = Trim(IIf(IsNull(rsy!qtybukadlmminggu.value), "", (rsy!qtybukadlmminggu.value)))
                ElseIf rsy!kdpelayananprofile = "03D" Then
                    .Cells(19, 14) = Trim(IIf(IsNull(rsy!qtybukadlmminggu.value), "", (rsy!qtybukadlmminggu.value)))
                ElseIf rsy!kdpelayananprofile = "03E" Then
                    .Cells(19, 17) = Trim(IIf(IsNull(rsy!qtybukadlmminggu.value), "", (rsy!qtybukadlmminggu.value)))
                ElseIf rsy!kdpelayananprofile = "03F" Then
                    .Cells(21, 5) = Trim(IIf(IsNull(rsy!qtybukadlmminggu.value), "", (rsy!qtybukadlmminggu.value)))
                ElseIf rsy!kdpelayananprofile = "03G" Then
                    .Cells(21, 8) = Trim(IIf(IsNull(rsy!qtybukadlmminggu.value), "", (rsy!qtybukadlmminggu.value)))
                ElseIf rsy!kdpelayananprofile = "03H" Then
                    .Cells(21, 11) = Trim(IIf(IsNull(rsy!qtybukadlmminggu.value), "", (rsy!qtybukadlmminggu.value)))
                ElseIf rsy!kdpelayananprofile = "03I" Then
                    .Cells(21, 14) = Trim(IIf(IsNull(rsy!qtybukadlmminggu.value), "", (rsy!qtybukadlmminggu.value)))
                ElseIf rsy!kdpelayananprofile = "03J" Then
                    .Cells(21, 17) = Trim(IIf(IsNull(rsy!qtybukadlmminggu.value), "", (rsy!qtybukadlmminggu.value)))
                ElseIf rsy!kdpelayananprofile = "03K" Then
                    .Cells(23, 5) = Trim(IIf(IsNull(rsy!qtybukadlmminggu.value), "", (rsy!qtybukadlmminggu.value)))
                ElseIf rsy!kdpelayananprofile = "03L" Then
                    .Cells(23, 8) = Trim(IIf(IsNull(rsy!qtybukadlmminggu.value), "", (rsy!qtybukadlmminggu.value)))
                ElseIf rsy!kdpelayananprofile = "03M" Then
                    .Cells(23, 11) = Trim(IIf(IsNull(rsy!qtybukadlmminggu.value), "", (rsy!qtybukadlmminggu.value)))
                ElseIf rsy!kdpelayananprofile = "03N" Then
                    .Cells(23, 14) = Trim(IIf(IsNull(rsy!qtybukadlmminggu.value), "", (rsy!qtybukadlmminggu.value)))
                ElseIf rsy!kdpelayananprofile = "03O" Then
                    .Cells(23, 17) = Trim(IIf(IsNull(rsy!qtybukadlmminggu.value), "", (rsy!qtybukadlmminggu.value)))
                End If
                rsy.MoveNext
            End With
        Wend

        Set rs1 = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='04'"
        Call msubRecFO(rs1, strSQL)
        While Not rs1.EOF
            With oSheet
                If rs1!kdpelayananprofile = "04A" Then
                    .Cells(25, 5) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "04B" Then
                    .Cells(25, 8) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "04C" Then
                    .Cells(25, 11) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "04D" Then
                    .Cells(25, 14) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "04E" Then
                    .Cells(25, 17) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "04F" Then
                    .Cells(27, 5) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                End If
                rs1.MoveNext
            End With
        Wend

        Set rs2 = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='05'"
        Call msubRecFO(rs2, strSQL)
        oSheet.Cells(28, 2) = Trim(IIf(IsNull(rs2!qtybukadlmminggu.value), 0, (rs2!qtybukadlmminggu.value)))

        Set rs1 = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='06'"
        Call msubRecFO(rs1, strSQL)
        While Not rs1.EOF
            With oSheet
                If rs1!kdpelayananprofile = "06A" Then
                    .Cells(30, 5) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "06B" Then
                    .Cells(30, 8) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "06C" Then
                    .Cells(30, 11) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "06D" Then
                    .Cells(30, 14) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "06E" Then
                    .Cells(30, 17) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                End If
                rs1.MoveNext
            End With
        Wend

        Set rsx = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='07'"
        Call msubRecFO(rsx, strSQL)
        While Not rsx.EOF
            With oSheet
                If rsx!kdpelayananprofile = "07A" Then
                    .Cells(32, 5) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                ElseIf rsx!kdpelayananprofile = "07B" Then
                    .Cells(32, 8) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                ElseIf rsx!kdpelayananprofile = "07C" Then
                    .Cells(32, 11) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                ElseIf rsx!kdpelayananprofile = "07D" Then
                    .Cells(32, 14) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                ElseIf rsx!kdpelayananprofile = "07E" Then
                    .Cells(32, 17) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                ElseIf rsx!kdpelayananprofile = "07F" Then
                    .Cells(34, 5) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                ElseIf rsx!kdpelayananprofile = "07G" Then
                    .Cells(34, 8) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                ElseIf rsx!kdpelayananprofile = "07H" Then
                    .Cells(34, 11) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                ElseIf rsx!kdpelayananprofile = "07I" Then
                    .Cells(34, 14) = Trim(IIf(IsNull(rsx!qtybukadlmminggu.value), "", (rsx!qtybukadlmminggu.value)))
                End If
                rsx.MoveNext
            End With
        Wend

        Set rs1 = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='08'"
        Call msubRecFO(rs1, strSQL)
        While Not rs1.EOF
            With oSheet
                If rs1!kdpelayananprofile = "08A" Then
                    .Cells(36, 5) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "08B" Then
                    .Cells(36, 8) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "08C" Then
                    .Cells(36, 11) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                End If
                rs1.MoveNext
            End With
        Wend

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='09'"
        Call msubRecFO(rs1, strSQL)
        While Not rs1.EOF
            With oSheet
                If rs1!kdpelayananprofile = "09A" Then
                    .Cells(38, 5) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "09B" Then
                    .Cells(38, 8) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "09C" Then
                    .Cells(38, 11) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "09D" Then
                    .Cells(38, 14) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "09E" Then
                    .Cells(38, 17) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "09F" Then
                    .Cells(40, 5) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "09G" Then
                    .Cells(40, 8) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                End If
                rs1.MoveNext
            End With
        Wend

        Set rs = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='10'"
        Call msubRecFO(rs, strSQL)
        While Not rs.EOF
            With oSheet
                If rs!kdpelayananprofile = "10A" Then
                    .Cells(42, 5) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "10B" Then
                    .Cells(42, 8) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "10C" Then
                    .Cells(42, 11) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "10D" Then
                    .Cells(42, 14) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "10E" Then
                    .Cells(42, 17) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "10F" Then
                    .Cells(44, 5) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "10G" Then
                    .Cells(44, 8) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "10H" Then
                    .Cells(44, 11) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "10I" Then
                    .Cells(44, 14) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "10J" Then
                    .Cells(44, 17) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                End If
                rs.MoveNext
            End With
        Wend

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='11'"
        Call msubRecFO(rs1, strSQL)
        While Not rs1.EOF
            With oSheet
                If rs1!kdpelayananprofile = "11A" Then
                    .Cells(46, 5) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "11B" Then
                    .Cells(46, 8) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "11C" Then
                    .Cells(46, 11) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "11D" Then
                    .Cells(46, 14) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "11E" Then
                    .Cells(46, 17) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "11F" Then
                    .Cells(48, 5) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "11G" Then
                    .Cells(48, 8) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                End If
                rs1.MoveNext
            End With
        Wend

        Set rs1 = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='12'"
        Call msubRecFO(rs1, strSQL)
        While Not rs1.EOF
            With oSheet
                If rs1!kdpelayananprofile = "12A" Then
                    .Cells(50, 5) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "12B" Then
                    .Cells(50, 8) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "12C" Then
                    .Cells(50, 11) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "12D" Then
                    .Cells(50, 14) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "12E" Then
                    .Cells(50, 17) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "12F" Then
                    .Cells(52, 5) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                End If
                rs1.MoveNext
            End With
        Wend

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='13'"
        Call msubRecFO(rs, strSQL)
        While Not rs.EOF
            With oSheet
                If rs!kdpelayananprofile = "13A" Then
                    .Cells(54, 5) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "13B" Then
                    .Cells(54, 8) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "13C" Then
                    .Cells(54, 11) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "13D" Then
                    .Cells(54, 14) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "13E" Then
                    .Cells(54, 17) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "13F" Then
                    .Cells(56, 5) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "13G" Then
                    .Cells(56, 8) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "13H" Then
                    .Cells(56, 11) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "13I" Then
                    .Cells(56, 14) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                ElseIf rs!kdpelayananprofile = "13J" Then
                    .Cells(56, 17) = Trim(IIf(IsNull(rs!qtybukadlmminggu.value), "", (rs!qtybukadlmminggu.value)))
                End If
                rs.MoveNext
            End With
        Wend

        Set rs1 = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='14'"
        Call msubRecFO(rs1, strSQL)
        While Not rs1.EOF
            With oSheet
                If rs1!kdpelayananprofile = "14A" Then
                    .Cells(58, 5) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "14B" Then
                    .Cells(58, 8) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "14C" Then
                    .Cells(58, 11) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                End If
                rs1.MoveNext
            End With
        Wend

        Set rs2 = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='15'"
        Call msubRecFO(rs2, strSQL)
        oSheet.Cells(60, 2) = Trim(IIf(IsNull(rs2!qtybukadlmminggu.value), 0, (rs2!qtybukadlmminggu.value)))

        Set rs1 = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='16'"
        Call msubRecFO(rs1, strSQL)
        While Not rs1.EOF
            With oSheet
                If rs1!kdpelayananprofile = "16A" Then
                    .Cells(62, 5) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "16B" Then
                    .Cells(62, 8) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "16C" Then
                    .Cells(62, 11) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "16D" Then
                    .Cells(62, 14) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "16E" Then
                    .Cells(62, 17) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                End If
                rs1.MoveNext
            End With
        Wend

        Set rs2 = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='17'"
        Call msubRecFO(rs2, strSQL)
        oSheet.Cells(64, 2) = Trim(IIf(IsNull(rs2!qtybukadlmminggu.value), 0, (rs2!qtybukadlmminggu.value)))

        Set rs1 = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='18'"
        Call msubRecFO(rs1, strSQL)
        While Not rs1.EOF
            With oSheet
                If rs1!kdpelayananprofile = "18A" Then
                    .Cells(66, 5) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "18B" Then
                    .Cells(66, 8) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                End If
                rs1.MoveNext
            End With
        Wend

        Set rs1 = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='19'"
        Call msubRecFO(rs1, strSQL)
        While Not rs1.EOF
            With oSheet
                If rs1!kdpelayananprofile = "19A" Then
                    .Cells(68, 5) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                ElseIf rs1!kdpelayananprofile = "19B" Then
                    .Cells(68, 8) = Trim(IIf(IsNull(rs1!qtybukadlmminggu.value), "", (rs1!qtybukadlmminggu.value)))
                End If
                rs1.MoveNext
            End With
        Wend

        Set rs2 = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='20'"
        Call msubRecFO(rs2, strSQL)
        oSheet.Cells(70, 2) = Trim(IIf(IsNull(rs2!qtybukadlmminggu.value), 0, (rs2!qtybukadlmminggu.value)))

        Set rs2 = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='29'"
        Call msubRecFO(rs2, strSQL)
        oSheet.Cells(71, 2) = Trim(IIf(IsNull(rs2!qtybukadlmminggu.value), 0, (rs2!qtybukadlmminggu.value)))

        Set rs2 = Nothing

        strSQL = "select * from RL3_RJ where kdjenispelayananprofile='30'"
        Call msubRecFO(rs2, strSQL)
        oSheet.Cells(72, 2) = Trim(IIf(IsNull(rs2!qtybukadlmminggu.value), 0, (rs2!qtybukadlmminggu.value)))

        Exit Sub
    End If

hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
End Sub

Private Sub Setcellku()
    With oSheet
        .Cells(j, 12) = Trim(IIf(IsNull(rsx!jmlbed), 0, (rsx!jmlbed + Cell12)))
    End With
End Sub

Private Sub setcellki()
    With oSheet
        .Cells(j, 15) = Trim(IIf(IsNull(rsx!jmlbed), 0, (rsx!jmlbed + Cell15)))
    End With
End Sub

Private Sub setcellkii()
    With oSheet
        .Cells(j, 18) = Trim(IIf(IsNull(rsx!jmlbed), 0, (rsx!jmlbed + Cell18)))
    End With
End Sub

Private Sub setcellkiii()
    With oSheet
        .Cells(j, 21) = Trim(IIf(IsNull(rsx!jmlbed), 0, (rsx!jmlbed + Cell21)))
    End With
End Sub

Private Sub setcelltk()
    With oSheet
        .Cells(j, 24) = Trim(IIf(IsNull(rsx!jmlbed), 0, (rsx!jmlbed + Cell24)))
    End With
End Sub
