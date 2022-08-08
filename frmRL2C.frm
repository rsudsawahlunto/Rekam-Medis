VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRL2C 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL2.C Data Status Imunisasi"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   5925
   Icon            =   "frmRL2C.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5925
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
         Format          =   115539971
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
      Picture         =   "frmRL2C.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRL2C.frx":2328
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   3360
      Picture         =   "frmRL2C.frx":4CE9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2955
   End
End
Attribute VB_Name = "frmRL2C"
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

Private Sub cmdCetak_Click()
    On Error GoTo hell

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    oXL.Visible = True
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2C.xls")
    Set oSheet = oWB.ActiveSheet

    Set rsb = Nothing
    strSQL = "select * from profilrs"
    Call msubRecFO(rsb, strSQL)

    Set oResizeRange = oSheet.Range("i6", "i7")
    oResizeRange.value = Trim(rsb!NamaRS)

    Set oResizeRange = oSheet.Range("u6", "u7")
    oResizeRange.value = Trim(rsb!KdRs)

    oSheet.Cells(4, 13).value = Format(frmRL2C.dtpAwal.value, "mmmm")
    oSheet.Cells(5, 13).value = Format(frmRL2C.dtpAwal.value, "yyyy")

    strSQL = "Select top 30 * from rl2c where bulan = '" & Format(dtpAwal.value, "MM") & "' and tahun= '" & Format(dtpAwal.value, "yyyy") & "'"
    Call msubRecFO(dbRst, strSQL)

    If dbRst.RecordCount > 0 Then
        dbRst.MoveFirst
        j = 11

        While Not dbRst.EOF
            With oSheet
                .Cells(j, 6) = Trim(IIf(IsNull(dbRst!NoCM.value), 0, (dbRst!NoCM.value)))
                If dbRst!JenisKelamin.value = "P" Then
                    oSheet.Cells(j, 8) = Trim(IIf(IsNull(dbRst!Umur.value), 0, (dbRst!Umur.value)))
                    oSheet.Cells(j, 7) = ""
                Else
                    oSheet.Cells(j, 8) = ""
                    oSheet.Cells(j, 7) = Trim(IIf(IsNull(dbRst!Umur.value), 0, (dbRst!Umur.value)))
                End If

                .Cells(j, 9) = Trim(IIf(IsNull(dbRst!Dipteri.value), 0, (dbRst!Dipteri.value)))
                .Cells(j, 10) = Trim(IIf(IsNull(dbRst!Petrtusis.value), 0, (dbRst!Petrtusis.value)))
                .Cells(j, 11) = Trim(IIf(IsNull(dbRst!Tetanus.value), 0, (dbRst!Tetanus.value)))
                .Cells(j, 12) = Trim(IIf(IsNull(dbRst![Tetanus Neonaturum].value), 0, (dbRst![Tetanus Neonaturum].value)))
                .Cells(j, 13) = Trim(IIf(IsNull(dbRst![TBC Paru].value), 0, (dbRst![TBC Paru].value)))
                .Cells(j, 14) = Trim(IIf(IsNull(dbRst!Campak.value), 0, (dbRst!Campak.value)))
                .Cells(j, 15) = Trim(IIf(IsNull(dbRst!Polio.value), 0, (dbRst!Polio.value)))
                .Cells(j, 16) = Trim(IIf(IsNull(dbRst!Hepatitis.value), 0, (dbRst!Hepatitis.value)))
                .Cells(j, 17) = Trim(IIf(IsNull(dbRst![0].value), 0, (dbRst![0].value)))
                .Cells(j, 18) = Trim(IIf(IsNull(dbRst![1].value), 0, (dbRst![1].value)))
                .Cells(j, 19) = Trim(IIf(IsNull(dbRst![2].value), 0, (dbRst![2].value)))
                .Cells(j, 20) = Trim(IIf(IsNull(dbRst!TK.value), 0, (dbRst!TK.value)))
                .Cells(j, 21) = Trim(IIf(IsNull(dbRst!Hidup.value), 0, (dbRst!Hidup.value)))
                .Cells(j, 22) = Trim(IIf(IsNull(dbRst!Mati.value), 0, (dbRst!Mati.value)))
            End With
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
