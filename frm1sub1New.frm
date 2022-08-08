VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frm1sub1New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL1.1 Data Dasar Rumah Sakit"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
   Icon            =   "frm1sub1New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5250
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
      Top             =   1320
      Width           =   5205
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   3120
      Picture         =   "frm1sub1New.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2115
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frm1sub1New.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm1sub1New.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frm1sub1New"
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
Dim i As Integer
'Special Buat Excel

Dim Cell40 As String
Dim Cell41 As String
Dim Cell42 As String
Dim Cell43 As String
Dim Cell44 As String
Dim Cell45 As String
Dim Cell46 As String
Dim Cell47 As String
Dim Cell48 As String
Dim Cell49 As String
Dim Cell50 As String
Dim Cell51 As String
Dim Cell52 As String
Dim Cell53 As String
Dim Cell54 As String
Dim Cell55 As String
Dim Cell56 As String
Dim Cell57 As String
Dim Cell58 As String
Dim Cell59 As String
Dim Cell60 As String
Dim Cell61 As String
Dim Cell62 As String
Dim Cell63 As String
Dim Cell64 As String


Private Sub cmdCetak_Click()
    On Error GoTo hell

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.Path & "\Formulir RL 1.1.xlsx")
    Set oSheet = oWB.ActiveSheet

    Set rs = Nothing

    strSQL = "select * from V_RL1_1New"
    Call msubRecFO(rs, strSQL)

    With oSheet
        .Cells(7, 4) = Format(Now, "yyyy")
        .Cells(10, 8) = Trim(IIf(IsNull(rs!KdRs.value), "", (rs!KdRs.value)))
      '  .Cells(11, 8) = Trim(IIf(IsNull(rs!TglSuratIjinLast.value), "", (rs!TglSuratIjinLast.value)))
        .Cells(11, 8) = Trim(IIf(IsNull(rs!TglRegistrasi.value), "", (rs!TglRegistrasi.value)))
        
        .Cells(12, 8) = Trim(IIf(IsNull(rs!NamaRS.value), "", (rs!NamaRS.value)))
        .Cells(13, 8) = Trim(IIf(IsNull(rs!JenisProfile.value), "", (rs!JenisProfile.value)))
        .Cells(14, 8) = Trim(IIf(IsNull(rs!KelasRS.value), "", (rs!KelasRS.value)))
        .Cells(15, 8) = Trim(IIf(IsNull(rs!Direktur.value), "", (rs!Direktur.value)))
        .Cells(16, 8) = Trim(IIf(IsNull(rs!PemilikProfile.value), "", (rs!PemilikProfile.value)))
        .Cells(17, 8) = Trim(IIf(IsNull(rs!Alamat.value), "", (rs!Alamat.value)))
        .Cells(18, 8) = Trim(IIf(IsNull(rs!KotaKodyaKab.value), "", (rs!KotaKodyaKab.value)))
        .Cells(19, 8) = Trim(IIf(IsNull(rs!KodePos.value), "", (rs!KodePos.value)))
        .Cells(20, 8) = Trim(IIf(IsNull(rs!Telepon.value), "", (rs!Telepon.value)))
        .Cells(21, 8) = Trim(IIf(IsNull(rs!Faks.value), "", (rs!Faks.value)))
        .Cells(22, 8) = Trim(IIf(IsNull(rs!Email.value), "", (rs!Email.value)))
        .Cells(23, 8) = Trim(IIf(IsNull(rs!Telepon.value), "", (rs!Telepon.value)))
        .Cells(24, 8) = Trim(IIf(IsNull(rs!Website.value), "", (rs!Website.value)))
        .Cells(26, 8) = Trim(IIf(IsNull(rs!LuasTanah.value), "", (rs!LuasTanah.value)))
        .Cells(27, 8) = Trim(IIf(IsNull(rs!LuasBangunan.value), "", (rs!LuasBangunan.value)))

        .Cells(29, 8) = Trim(IIf(IsNull(rs!NoSuratIjinLast.value), "", (rs!NoSuratIjinLast.value)))
        .Cells(30, 8) = Trim(IIf(IsNull(rs!TglSuratIjinLast.value), "", (rs!TglSuratIjinLast.value)))
        .Cells(31, 8) = Trim(IIf(IsNull(rs!SignatureByLast.value), "", (rs!SignatureByLast.value)))
        .Cells(32, 8) = Trim(IIf(IsNull(rs!StatusSuratIjin.value), "", (rs!StatusSuratIjin.value)))
        .Cells(33, 8) = Trim(IIf(IsNull(rs!MasaBerlakuIjin.value), "", (rs!MasaBerlakuIjin.value)))
        .Cells(34, 8) = "-"
        .Cells(36, 8) = Trim(IIf(IsNull(rs!TahapanAkreditasi.value), "", (rs!TahapanAkreditasi.value)))
        .Cells(37, 8) = Trim(IIf(IsNull(rs!statusAkreditasi.value), "", (rs!statusAkreditasi.value)))
        .Cells(38, 8) = Trim(IIf(IsNull(rs!TglAkreditasi.value), "", (rs!TglAkreditasi.value)))
    End With

    Set rs1 = Nothing

    strSQL1 = "SELECT Kelas, SUM(JmlBed) AS JmlBed From V_JmlBedPerRuangan GROUP BY Kelas, TglAwalSK, TglAkhirSK Having (TglAwalSK <= GETDATE()) And (TglAkhirSK >= GETDATE())ORDER BY Kelas"
    Call msubRecFO(rs1, strSQL1)

    With oSheet
        For i = 1 To rs1.RecordCount
            If rs1!Kelas = "Kelas (VVIP)" Then
                Call SetcellforVVIP
            ElseIf rs1!Kelas = "MASTER (VIP)" Then
                Call SetcellforVIP
            ElseIf rs1!Kelas = "SUITE (I)" Then
                Call SetcellforI
            ElseIf rs1!Kelas = "DELUXE (II)" Then
                Call SetcellforII
            ElseIf rs1!Kelas = "STANDARD (III)" Then
                Call SetcellforIII
            End If
            rs1.MoveNext
        Next i
    End With

    strSQL2 = "select KdJenisPegawai,KdJabatan,NamaJabatan, Bagian,sum(Jumlah) as Jumlah From V_JumlahKaryawanBerdasarkanJabatan Group by KdJenisPegawai,KdJabatan,NamaJabatan,Bagian order by KdJenisPegawai"
    Call msubRecFO(rs2, strSQL2)
    
    With oSheet
        
        For i = 1 To rs2.RecordCount
            If rs2!KdJenisPegawai = "001" And rs2!NamaJabatan = "Dokter Ahli Anak" Then
                Call SetcellforDokterSpesialisAnak
            ElseIf rs2!KdJenisPegawai = "001" And rs2!NamaJabatan = "Dokter Ahli Kebidanan" Then
                Call SetcellforDokterSpesialisKebidanan
            ElseIf rs2!KdJenisPegawai = "001" And rs2!NamaJabatan = "Dokter Ahli Penyakit Dalam" Or rs2!NamaJabatan = "Dokter Spesialis Dalam" Then
                Call SetcellforDokterSpesialisPenyakitDalam
            ElseIf rs2!KdJenisPegawai = "001" And rs2!NamaJabatan = "Dokter Ahli Bedah" Then
                Call SetcellforDokterSpesialisBedah
            ElseIf rs2!KdJenisPegawai = "001" And rs2!NamaJabatan = "Dokter Spesialis Radiologi" Then
                Call SetcellforDokterSpesialisRadiologi
            ElseIf rs2!KdJenisPegawai = "001" And rs2!NamaJabatan = "Dokter Spesialis Rehabilitasi Medik" Then
                Call SetcellforDokterSpesialisRehabilitasiMedik
            ElseIf rs2!KdJenisPegawai = "001" And rs2!NamaJabatan = "Dokter Spesialis Anestesi" Then
                Call SetcellforDokterSpesialisAnestesiologi
            ElseIf rs2!KdJenisPegawai = "001" And rs2!NamaJabatan = "Dokter Spesialis Jantung" Then
                Call SetcellforDokterSpesialisJantung
            ElseIf rs2!KdJenisPegawai = "001" And rs2!NamaJabatan = "Dokter Spesialis Mata" Then
                Call SetcellforDokterSpesialisMata
            ElseIf rs2!KdJenisPegawai = "001" And rs2!NamaJabatan = "Dokter Spesialis THT" Then
                Call SetcellforDokterSpesialisTHT
            ElseIf rs2!KdJenisPegawai = "001" And rs2!NamaJabatan = "Dokter Spesialist Jiwa" Then
                Call SetcellforDokterSpesialisJiwa
            ElseIf rs2!KdJenisPegawai = "001" And rs2!NamaJabatan = "Dokter Umum" Then
                Call SetcellforDokterUmum
            ElseIf rs2!KdJenisPegawai = "001" And rs2!NamaJabatan = "Dokter Gigi & Mulut" Then
                Call SetcellforDokterGigi
            ElseIf rs2!KdJenisPegawai = "001" And rs2!NamaJabatan = "Dokter Gigi Spesialis" Then
                Call SetcellforDokterGigiSpesialis
            ElseIf rs2!KdJenisPegawai = "002" Then
                Call SetcellforPerawat
            ElseIf rs2!KdJenisPegawai = "006" Then
                Call SetcellforBidan
            ElseIf rs2!KdJenisPegawai = "012" Then
                Call SetcellforFarmasi
            ElseIf rs2!Bagian = "Bagian Tenaga Medis 2" Then
                Call SetcellforTenagaKesehatanLainnya
            ElseIf rs2!Bagian = "Bagian Tenaga Non Kesehatan" Then
                
                 Call SetcellforTenagaNonKesehatan
            End If
            rs2.MoveNext
        Next i
    
    End With

    oXL.Visible = True
    
 Exit Sub

hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
End Sub

Private Sub SetcellforVVIP()
    With oSheet
        
         Cell40 = oSheet.Cells(40, 8).value
        .Cells(40, 8) = Trim(IIf(IsNull(rs1!jmlbed), 0, (rs1!jmlbed)) + Cell40)
    End With
End Sub

Private Sub SetcellforVIP()
    With oSheet
         Cell41 = oSheet.Cells(41, 8).value
        .Cells(41, 8) = Trim(IIf(IsNull(rs1!jmlbed), 0, (rs1!jmlbed)) + Cell41)
    End With
End Sub

Private Sub SetcellforI()
    With oSheet
         Cell42 = oSheet.Cells(42, 8).value
        .Cells(42, 8) = Trim(IIf(IsNull(rs1!jmlbed), 0, (rs1!jmlbed)) + Cell42)
    End With
End Sub

Private Sub SetcellforII()
    With oSheet
         Cell43 = oSheet.Cells(43, 8).value
        .Cells(43, 8) = Trim(IIf(IsNull(rs1!jmlbed), 0, (rs1!jmlbed)) + Cell43)
    End With
End Sub

Private Sub SetcellforIII()
    With oSheet
         Cell44 = oSheet.Cells(44, 8).value
        .Cells(44, 8) = Trim(IIf(IsNull(rs1!jmlbed), 0, (rs1!jmlbed)) + Cell44)
    End With
End Sub

Private Sub SetcellforDokterSpesialisAnak()
    With oSheet
         Cell46 = oSheet.Cells(46, 8).value
        .Cells(46, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell46)
    End With
End Sub

Private Sub SetcellforDokterSpesialisKebidanan()
    With oSheet
         Cell47 = oSheet.Cells(47, 8).value
        .Cells(47, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell47)
    End With
End Sub

Private Sub SetcellforDokterSpesialisPenyakitDalam()
    With oSheet
         Cell48 = oSheet.Cells(48, 8).value
        .Cells(48, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell48)
    End With
End Sub

Private Sub SetcellforDokterSpesialisBedah()
    With oSheet
         Cell49 = oSheet.Cells(49, 8).value
        .Cells(49, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell49)
    End With
End Sub

Private Sub SetcellforDokterSpesialisRadiologi()
    With oSheet
         Cell50 = oSheet.Cells(50, 8).value
        .Cells(50, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell50)
    End With
End Sub

Private Sub SetcellforDokterSpesialisRehabilitasiMedik()
    With oSheet
         Cell51 = oSheet.Cells(51, 8).value
        .Cells(51, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell51)
    End With
End Sub

Private Sub SetcellforDokterSpesialisAnestesiologi()
    With oSheet
         Cell52 = oSheet.Cells(52, 8).value
        .Cells(52, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell52)
    End With
End Sub

Private Sub SetcellforDokterSpesialisJantung()
    With oSheet
         Cell53 = oSheet.Cells(53, 8).value
        .Cells(53, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell53)
    End With
End Sub

Private Sub SetcellforDokterSpesialisMata()
    With oSheet
         Cell54 = oSheet.Cells(54, 8).value
        .Cells(54, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell54)
    End With
End Sub

Private Sub SetcellforDokterSpesialisTHT()
    With oSheet
        Cell55 = oSheet.Cells(55, 8).value
        .Cells(55, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell55)
    End With
End Sub

Private Sub SetcellforDokterSpesialisJiwa()
    With oSheet
         Cell56 = oSheet.Cells(56, 8).value
        .Cells(56, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell56)
    End With
End Sub

Private Sub SetcellforDokterUmum()
    With oSheet
         Cell57 = oSheet.Cells(57, 8).value
        .Cells(57, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell57)
    End With
End Sub

Private Sub SetcellforDokterGigi()
    With oSheet
         Cell58 = oSheet.Cells(58, 8).value
        .Cells(58, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell58)
    End With
End Sub

Private Sub SetcellforDokterGigiSpesialis()
    With oSheet
         Cell59 = oSheet.Cells(59, 8).value
        .Cells(59, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell59)
    End With
End Sub

Private Sub SetcellforPerawat()
    With oSheet
         Cell60 = oSheet.Cells(60, 8).value
        .Cells(60, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell60)
    End With
End Sub

Private Sub SetcellforBidan()
    With oSheet
         Cell61 = oSheet.Cells(61, 8).value
        .Cells(61, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell61)
    End With
End Sub

Private Sub SetcellforFarmasi()
    With oSheet
         Cell62 = oSheet.Cells(62, 8).value
        .Cells(62, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell62)
    End With
End Sub

Private Sub SetcellforTenagaKesehatanLainnya()
    With oSheet
         Cell63 = oSheet.Cells(63, 8).value
        .Cells(63, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell63)
    End With
End Sub

Private Sub SetcellforTenagaNonKesehatan()
    With oSheet
         Cell64 = oSheet.Cells(64, 8).value
        .Cells(64, 8) = Trim(IIf(IsNull(rs2!Jumlah), 0, (rs2!Jumlah)) + Cell64)
    End With
End Sub
