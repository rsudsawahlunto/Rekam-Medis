VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm3sub15New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.15 Cara Bayar"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
   Icon            =   "frm3sub15New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2925
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
      Top             =   1680
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
   Begin MSComCtl2.DTPicker dtpAwal 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2040
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
      Format          =   134348803
      UpDown          =   -1  'True
      CurrentDate     =   40544
   End
   Begin MSComCtl2.DTPicker dtpAkhir 
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1920
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
      Format          =   134348803
      UpDown          =   -1  'True
      CurrentDate     =   40544
   End
   Begin MSComCtl2.DTPicker dtptahun 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   1200
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
      Format          =   134348803
      UpDown          =   -1  'True
      CurrentDate     =   40544
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblPersen 
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   2595
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   3120
      Picture         =   "frm3sub15New2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2115
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frm3sub15New2.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm3sub15New2.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frm3sub15New2"
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
Dim xx As Integer
'Special Buat Excel

Dim Cell7 As String
Dim Cell8 As String
Dim Cell9 As String
Dim Cell10 As String
Dim Cell11 As String
Dim Cell12 As String

Private Sub cmdCetak_Click()
    On Error GoTo errLoad

    ProgressBar1.value = ProgressBar1.Min
    lblPersen.Caption = "0 %"
    Screen.MousePointer = vbHourglass

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")

    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.Path & "\RL 3.15_cara bayar.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    For xx = 2 To 10
        With oSheet
            .Cells(xx, 3) = rsb("KdRS").value
            .Cells(xx, 2) = rsb("KotaKodyaKab").value
            .Cells(xx, 4) = rsb("NamaRS").value
            .Cells(xx, 5) = Format(dtptahun.value, "YYYY")
        End With
    Next xx

    Set rs = Nothing
    strSQL = " select * from RL3_15New where Year(TglMasuk) = '" & dtptahun.Year & "' or tglmasuk is null"
    Call msubRecFO(rs, strSQL)

    ProgressBar1.Min = 0
    ProgressBar1.Max = rs.RecordCount
    ProgressBar1.value = 0

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        While Not rs.EOF
            If rs!NamaExternal = "Membayar" Then
                j = 2
            ElseIf rs!NamaExternal = "Keringanan" Then
                j = 6
            ElseIf rs!NamaExternal = "Askes" Then
                j = 4
            ElseIf rs!NamaExternal = "Asuransi Lain" Then
                j = 5
            ElseIf rs!NamaExternal = "Kartu Sehat" Then
                j = 8
            ElseIf rs!NamaExternal = "Keterangan Tidak Mampu" Then
                j = 9
            End If

            Cell7 = oSheet.Cells(j, 8).value
            Cell8 = oSheet.Cells(j, 9).value
            Cell9 = oSheet.Cells(j, 10).value
            Cell10 = oSheet.Cells(j, 11).value
            Cell11 = oSheet.Cells(j, 12).value
            Cell12 = oSheet.Cells(j, 13).value

            If rs!NamaExternal = "Membayar" Then
                With oSheet
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!jmlpasienkeluar.value), 0, (rs!jmlpasienkeluar.value)) + Cell7)
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!lamadirawat.value), 0, (rs!lamadirawat.value)) + Cell8)
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpasienrj.value), 0, (rs!jmlpasienrj.value)) + Cell9)
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlpasienlab.value), 0, (rs!jmlpasienlab.value)) + Cell10)
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmlpasienrad.value), 0, (rs!jmlpasienrad.value)) + Cell11)
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!jmllainnya.value), 0, (rs!jmllainnya.value)) + Cell12)
                End With

            ElseIf rs!NamaExternal = "Keringanan" Then
                With oSheet
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!jmlpasienkeluar.value), 0, (rs!jmlpasienkeluar.value)) + Cell7)
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!lamadirawat.value), 0, (rs!lamadirawat.value)) + Cell8)
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpasienrj.value), 0, (rs!jmlpasienrj.value)) + Cell9)
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlpasienlab.value), 0, (rs!jmlpasienlab.value)) + Cell10)
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmlpasienrad.value), 0, (rs!jmlpasienrad.value)) + Cell11)
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!jmllainnya.value), 0, (rs!jmllainnya.value)) + Cell12)
                End With

            ElseIf rs!NamaExternal = "Askes" Then
                With oSheet
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!jmlpasienkeluar.value), 0, (rs!jmlpasienkeluar.value)) + Cell7)
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!lamadirawat.value), 0, (rs!lamadirawat.value)) + Cell8)
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpasienrj.value), 0, (rs!jmlpasienrj.value)) + Cell9)
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlpasienlab.value), 0, (rs!jmlpasienlab.value)) + Cell10)
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmlpasienrad.value), 0, (rs!jmlpasienrad.value)) + Cell11)
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!jmllainnya.value), 0, (rs!jmllainnya.value)) + Cell12)
                End With

            ElseIf rs!NamaExternal = "Asuransi Lain" Then
                With oSheet
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!jmlpasienkeluar.value), 0, (rs!jmlpasienkeluar.value)) + Cell7)
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!lamadirawat.value), 0, (rs!lamadirawat.value)) + Cell8)
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpasienrj.value), 0, (rs!jmlpasienrj.value)) + Cell9)
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlpasienlab.value), 0, (rs!jmlpasienlab.value)) + Cell10)
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmlpasienrad.value), 0, (rs!jmlpasienrad.value)) + Cell11)
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!jmllainnya.value), 0, (rs!jmllainnya.value)) + Cell12)
                End With

            ElseIf rs!NamaExternal = "Kartu Sehat" Then
                With oSheet
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!jmlpasienkeluar.value), 0, (rs!jmlpasienkeluar.value)) + Cell7)
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!lamadirawat.value), 0, (rs!lamadirawat.value)) + Cell8)
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpasienrj.value), 0, (rs!jmlpasienrj.value)) + Cell9)
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlpasienlab.value), 0, (rs!jmlpasienlab.value)) + Cell10)
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmlpasienrad.value), 0, (rs!jmlpasienrad.value)) + Cell11)
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!jmllainnya.value), 0, (rs!jmllainnya.value)) + Cell12)
                End With

            ElseIf rs!NamaExternal = "Keterangan Tidak Mampu" Then
                With oSheet
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!jmlpasienkeluar.value), 0, (rs!jmlpasienkeluar.value)) + Cell7)
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!lamadirawat.value), 0, (rs!lamadirawat.value)) + Cell8)
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpasienrj.value), 0, (rs!jmlpasienrj.value)) + Cell9)
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlpasienlab.value), 0, (rs!jmlpasienlab.value)) + Cell10)
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmlpasienrad.value), 0, (rs!jmlpasienrad.value)) + Cell11)
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!jmllainnya.value), 0, (rs!jmllainnya.value)) + Cell12)
                End With

            End If
            rs.MoveNext

            ProgressBar1.value = Int(ProgressBar1.value) + 1
            lblPersen.Caption = Int(ProgressBar1.value * 100 / ProgressBar1.Max) & " %"
        Wend
    End If
    oXL.Visible = True
    Screen.MousePointer = vbDefault

    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)

    dtptahun.value = Now
    dtptahun.CustomFormat = "yyyyy"
End Sub
