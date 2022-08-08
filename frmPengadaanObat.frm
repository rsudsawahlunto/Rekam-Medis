VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPengadaanObat 
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
   Icon            =   "frmPengadaanObat.frx":0000
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
         Format          =   63700995
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
         Format          =   63700995
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
         Format          =   63700995
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
      Picture         =   "frmPengadaanObat.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2955
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPengadaanObat.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPengadaanObat.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmPengadaanObat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim Cell12 As String
Dim Cell13 As String


Private Sub Check1_Click()
If Check1.Value = 0 Then
   dtpAwal.Enabled = True
   dtpAkhir.Enabled = True
   dtptahun.Enabled = False
   Option1.Enabled = False
   Option2.Enabled = False
   Option3.Enabled = False
   Option4.Enabled = False
   dtpAwal.Value = Now
   dtpAkhir.Value = Now
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
   dtptahun.Value = Now
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
     oResizeRange.Value = Trim(rsb!KdRs)

strSQL = "Select * from V_PengadaanObat where TglTerima between '" & Format(dtpAwal.Value, "yyyy/MM/dd") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd") & "'"
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
    
Cell12 = oSheet.Cells(j, 13).Value
Cell13 = oSheet.Cells(j, 16).Value

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
Exit Sub
errLoad:
'    msubPesanError
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
        .dtpAwal.Value = Now
        .dtpAkhir.Value = Now
        .dtptahun.Value = Now
    End With
    Check1.Value = 1
    Option1.Value = 1

End Sub

Private Sub Option1_Click()
        awal = CStr(dtptahun.Year) + "/01/01"
        akhir = CStr(dtptahun.Year) + "/03/31"

        dtpAwal.Value = awal
        dtpAkhir.Value = akhir
End Sub

Private Sub Option2_Click()
        awal = CStr(dtptahun.Year) + "/04/01"
        akhir = CStr(dtptahun.Year) + "/06/30"

        dtpAwal.Value = awal
        dtpAkhir.Value = akhir

End Sub

Private Sub Option3_Click()
        awal = CStr(dtptahun.Year) + "/07/01"
        akhir = CStr(dtptahun.Year) + "/09/30"

        dtpAwal.Value = awal
        dtpAkhir.Value = akhir

End Sub

Private Sub Option4_Click()
        awal = CStr(dtptahun.Year) + "/10/01"
        akhir = CStr(dtptahun.Year) + "/12/31"

        dtpAwal.Value = awal
        dtpAkhir.Value = akhir

End Sub


