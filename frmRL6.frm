VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRL6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL6 Formulir Pelaporan Infeksi Nosokmial"
   ClientHeight    =   3090
   ClientLeft      =   7725
   ClientTop       =   4020
   ClientWidth     =   6360
   Icon            =   "frmRL6.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6360
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
      TabIndex        =   2
      Top             =   2280
      Width           =   6285
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   1905
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
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   120
         TabIndex        =   1
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
         CustomFormat    =   " MMMM yyyy"
         Format          =   125829123
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   6
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
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL6.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRL6.frx":2328
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   3360
      Picture         =   "frmRL6.frx":4CE9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2955
   End
End
Attribute VB_Name = "frmRL6"
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

Private Sub cmdCetak_Click()
    On Error GoTo hell

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    oXL.Visible = True
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.Path & "\RL 6.xls")
    Set oSheet = oWB.ActiveSheet

    Set rsb = Nothing
    strSQL = "select * from profilrs"
    Call msubRecFO(rsb, strSQL)

    Set oResizeRange = oSheet.Range("i6", "i7")
    oResizeRange.value = Trim(rsb!NamaRS)

    Set oResizeRange = oSheet.Range("u6", "u7")
    oResizeRange.value = Trim(rsb!KdRs)

    oSheet.Cells(4, 13).value = Format(frmRL6.dtpAwal.value, "mmmm")

    strSQL = "Select * from RL6 where bulan = '" & Format(dtpAwal.value, "MM") & "' and tahun= '" & Format(dtpAwal.value, "yyyy") & "'or bulan is null and tahun is null"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst

        While Not dbRst.EOF
            If dbRst![SpesialisasiRuangan] = "Bedah" Then
                j = 12
            ElseIf dbRst![SpesialisasiRuangan] = "Pnykt. Dalam" Then
                j = 13
            ElseIf dbRst![SpesialisasiRuangan] = "Ruang Anak" Then
                j = 14
            ElseIf dbRst![SpesialisasiRuangan] = "Kebidanan" Then
                j = 15
            ElseIf dbRst![SpesialisasiRuangan] = "Syaraf" Then
                j = 16
            ElseIf dbRst![SpesialisasiRuangan] = "Umum" Then
                j = 17
            ElseIf dbRst![SpesialisasiRuangan] = "ICU" Then
                j = 18
            ElseIf dbRst![SpesialisasiRuangan] = "NICU" Then
                j = 19
            ElseIf dbRst![SpesialisasiRuangan] = "PICU" Then
                j = 20
            ElseIf dbRst![SpesialisasiRuangan] = "Perinatologi" Then
                j = 21
            ElseIf dbRst![SpesialisasiRuangan] = "Lain-lain" Then
                j = 22
            End If

            Cell7 = oSheet.Cells(j, 7).value
            Cell8 = oSheet.Cells(j, 8).value
            Cell9 = oSheet.Cells(j, 9).value
            Cell10 = oSheet.Cells(j, 10).value
            Cell11 = oSheet.Cells(j, 11).value
            Cell12 = oSheet.Cells(j, 12).value
            Cell13 = oSheet.Cells(j, 13).value
            Cell14 = oSheet.Cells(j, 14).value
            Cell15 = oSheet.Cells(j, 15).value
            Cell16 = oSheet.Cells(j, 16).value
            Cell17 = oSheet.Cells(j, 17).value
            Cell18 = oSheet.Cells(j, 18).value
            Cell19 = oSheet.Cells(j, 19).value
            Cell20 = oSheet.Cells(j, 20).value
            Cell21 = oSheet.Cells(j, 21).value

            If dbRst![SpesialisasiRuangan] = "Bedah" Then
                Call setcell
            ElseIf dbRst![SpesialisasiRuangan] = "Pnykt. Dalam" Then
                Call setcell
            ElseIf dbRst![SpesialisasiRuangan] = "Ruang Anak" Then
                Call setcell
            ElseIf dbRst![SpesialisasiRuangan] = "Kebidanan" Then
                Call setcell
            ElseIf dbRst![SpesialisasiRuangan] = "Syaraf" Then
                Call setcell
            ElseIf dbRst![SpesialisasiRuangan] = "Umum" Then
                Call setcell
            ElseIf dbRst![SpesialisasiRuangan] = "ICU" Then
                Call setcell
            ElseIf dbRst![SpesialisasiRuangan] = "NICU" Then
                Call setcell
            ElseIf dbRst![SpesialisasiRuangan] = "PICU" Then
                Call setcell
            ElseIf dbRst![SpesialisasiRuangan] = "Perinatologi" Then
                Call setcell
            ElseIf dbRst![SpesialisasiRuangan] = "Lain-lain" Then
                Call setcell
            End If

            j = j + 1

            dbRst.MoveNext
        Wend
    End If

hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    With Me
        .dtpAwal.value = Now
    End With
End Sub

Private Sub setcell()
    With oSheet
        .Cells(j, 7) = Trim(dbRst![PasienKeluar] + Cell7)
        .Cells(j, 8) = Trim(dbRst![ISK_IN] + Cell8)
        .Cells(j, 9) = Trim(dbRst![ISK_Pasien] + Cell9)
        .Cells(j, 10) = Trim(dbRst![ILO_IN] + Cell10)
        .Cells(j, 11) = Trim(dbRst![ILO_Pasien] + Cell11)
        .Cells(j, 12) = Trim(dbRst![Pneumonia_IN] + Cell12)
        .Cells(j, 13) = Trim(dbRst![Pneumonia_Pasien] + Cell13)
        .Cells(j, 14) = Trim(dbRst![Sepsis_IN] + Cell14)
        .Cells(j, 15) = Trim(dbRst![Sepsis_Pasien] + Cell15)
        .Cells(j, 16) = Trim(dbRst![Dekubitus_IN] + Cell16)
        .Cells(j, 17) = Trim(dbRst![Dekubitus_Pasien] + Cell17)
        .Cells(j, 18) = Trim(dbRst![Phlebitis_IN] + Cell18)
        .Cells(j, 19) = Trim(dbRst![Phlebitis_Pasien] + Cell19)
        .Cells(j, 20) = Trim(dbRst![LainLain_IN] + Cell20)
        .Cells(j, 21) = Trim(dbRst![LainLain_Pasien] + Cell21)
    End With
End Sub
