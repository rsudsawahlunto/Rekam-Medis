VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUreqKegiatanRS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL1 Halaman 1"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5925
   Icon            =   "frmUreqKegiatanRS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
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
      TabIndex        =   11
      Top             =   3960
      Width           =   6285
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   3480
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   480
         TabIndex        =   12
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
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   2295
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   120
         TabIndex        =   10
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
         Format          =   63111171
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
      TabIndex        =   7
      Top             =   3120
      Width           =   2295
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   375
         Left            =   120
         TabIndex        =   8
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
         Format          =   63111171
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Triwulan"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   5655
      Begin VB.OptionButton Option1 
         Caption         =   "Triwulan1"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Triwulan4"
         Height          =   495
         Left            =   4200
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Triwulan3"
         Height          =   495
         Left            =   2880
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Triwulan2"
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Triwulan"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtptahun 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
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
         Format          =   63111171
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   16
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
      TabIndex        =   14
      Top             =   1080
      Width           =   6255
      Begin VB.Label Label1 
         Caption         =   "s/d"
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   2280
         Width           =   375
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmUreqKegiatanRS.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmUreqKegiatanRS.frx":2328
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   3360
      Picture         =   "frmUreqKegiatanRS.frx":4CE9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2955
   End
End
Attribute VB_Name = "frmUreqKegiatanRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim awal As String
Dim akhir As String
'Untuk Pengganti Group Dijadikan Penginputan Di Cell yg sama
Dim Cell7 As String
Dim Cell8 As String
Dim Cell9 As String
Dim Cell10 As String
Dim Cell11 As String
Dim Cell13 As String
Dim Cell16 As String
Dim Cell17 As String
Dim Cell18 As String
Dim Cell19 As String
Dim Cell20 As String
'Untuk Pengganti Group Dijadikan Penginputan Di Cell yg sama
'Special Buat Excel
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel.Range
Dim oResizeRange As Excel.Range
Dim j As String
'Special Buat Excel

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
On Error GoTo hell

'Buka Excel
      Set oXL = CreateObject("Excel.Application")
      oXL.Visible = True
'Buat Buka Template
      Set oWB = oXL.Workbooks.Open(App.Path & "\RL1 Hal1.xls")
      Set oSheet = oWB.ActiveSheet
      
      
If Check1.Value = vbChecked And Option1.Value = True Then
oSheet.Cells(4, 13).Value = "I"
ElseIf Check1.Value = vbChecked And Option2.Value = True Then
oSheet.Cells(4, 13).Value = "II"
ElseIf Check1.Value = vbChecked And Option3.Value = True Then
oSheet.Cells(4, 13).Value = "III"
ElseIf Check1.Value = vbChecked And Option4.Value = True Then
oSheet.Cells(4, 13).Value = "IV"
ElseIf Check1.Value = vbUnchecked Then
oSheet.Cells(4, 13).Value = ""
End If

    Set rsb = Nothing
strSQL = "select * from profilrs"
   Call msubRecFO(rsb, strSQL)

 Set oResizeRange = oSheet.Range("g6", "g7")
     oResizeRange.Value = Trim(rsb!NamaRs)
     
 Set oResizeRange = oSheet.Range("t6", "t7")
     oResizeRange.Value = Trim(rsb!KdRs)
     



    Set rs = Nothing
                
strSQL = "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienAwal as [3] WHERE  (TglMasuk BETWEEN DateAdd(MONTH,-3,'" & Format(dtpAwal, "yyyy/MM/dd ") & "') AND DATEADD(MONTH,-3,'" & Format(dtpAkhir, "yyyy/MM/dd") & "')) or (tglmasuk IS NULL)  " & _
"Union " & _
"select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienMasuk as [4]        WHERE   TglMasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "'  AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' or (tglmasuk IS NULL)   " & _
"Union " & _
"select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglPulang,KdSubInstalasi from LaporanRL11_PasienKeluarHidup as [5] WHERE   TglPulang BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' or (TglPulang IS NULL)  " & _
"Union " & _
"select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglPulang,KdSubInstalasi from LaporanRL11_PasienKeluarMati6 as [6] where   TglPulang BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' or (TglPulang IS NULL)  " & _
"Union " & _
"select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglPulang,KdSubInstalasi from LaporanRL11_PasienKeluarMati7 as [7] where   TglPulang BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' or (TglPulang IS NULL)  " & _
"Union " & _
"select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglPulang,KdSubInstalasi from LaporanRL11_PasienKeluarMati8 as [8] where   TglPulang BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' or (TglPulang IS NULL)  " & _
"Union " & _
"select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglPulang,KdSubInstalasi from LaporanRL11_LamaDirawat as [9]       where   TglPulang BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' or (TglPulang IS NULL)  " & _
"Union " & _
"select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_JmlHariRawat as [11]      where   TglMasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "'  AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' or (TglMasuk BETWEEN DateAdd(MONTH,-3,'" & Format(dtpAwal, "yyyy/MM/dd ") & "') AND DATEADD(MONTH,-3,'" & Format(dtpAkhir, "yyyy/MM/dd") & "')) or (tglmasuk IS NULL)  " & _
"Union " & _
"select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],ISNULL(DATEDIFF(day,TglMasuk, { fn NOW() }), 0) As [12],[13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienKelas as [12] where TglMasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' and (KdKelas in('05','05')) or (TglMasuk BETWEEN DateAdd(MONTH,-3,'" & Format(dtpAwal, "yyyy/MM/dd ") & "') AND DATEADD(MONTH,-3,'" & Format(dtpAkhir, "yyyy/MM/dd") & "')) and (KdKelas in('05','05'))  " & _
"Union " & _
"select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],ISNULL(DATEDIFF(day,TglMasuk, { fn NOW() }), 0) As [13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienKelas as [13] where TglMasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' and (KdKelas in('03')) or (TglMasuk BETWEEN DateAdd(MONTH,-3,'" & Format(dtpAwal, "yyyy/MM/dd ") & "') AND DATEADD(MONTH,-3,'" & Format(dtpAkhir, "yyyy/MM/dd") & "')) and (KdKelas in('03'))  " & _
"Union " & _
"select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],ISNULL(DATEDIFF(day,TglMasuk, { fn NOW() }), 0) As [14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienKelas as [14] where TglMasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' and (KdKelas in('02')) or (TglMasuk BETWEEN DateAdd(MONTH,-3,'" & Format(dtpAwal, "yyyy/MM/dd ") & "') AND DATEADD(MONTH,-3,'" & Format(dtpAkhir, "yyyy/MM/dd") & "')) and (KdKelas in('02'))  " & _
"Union " & _
"select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],ISNULL(DATEDIFF(day,TglMasuk, { fn NOW() }), 0) As [15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienKelas as [15] where TglMasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' and (KdKelas in('01')) or (TglMasuk BETWEEN DateAdd(MONTH,-3,'" & Format(dtpAwal, "yyyy/MM/dd ") & "') AND DATEADD(MONTH,-3,'" & Format(dtpAkhir, "yyyy/MM/dd") & "')) and (KdKelas in('01'))  " & _
"Union " & _
"select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],ISNULL(DATEDIFF(day,TglMasuk, { fn NOW() }), 0) As [16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienKelas as [16] where TglMasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' and (KdKelas in('07')) or (TglMasuk BETWEEN DateAdd(MONTH,-3,'" & Format(dtpAwal, "yyyy/MM/dd ") & "') AND DATEADD(MONTH,-3,'" & Format(dtpAkhir, "yyyy/MM/dd") & "')) and (KdKelas in('07')) "
                
Call msubRecFO(rs, strSQL)

'Start Report From Excel
    If rs.RecordCount > 0 Then
       rs.MoveFirst
 While Not rs.EOF


If rs!kdsubinstalasi = "001" Then
j = 12
ElseIf rs!kdsubinstalasi = "002" Then
j = 13
ElseIf rs!kdsubinstalasi = "003" Then
j = 14
ElseIf rs!kdsubinstalasi = "004" Then
j = 15
ElseIf rs!kdsubinstalasi = "005" Then
j = 16
ElseIf rs!kdsubinstalasi = "006" Then
j = 17
ElseIf rs!kdsubinstalasi = "007" Then
j = 18
ElseIf rs!kdsubinstalasi = "008" Then
j = 19
ElseIf rs!kdsubinstalasi = "009" Then
j = 20
ElseIf rs!kdsubinstalasi = "010" Then
j = 21
ElseIf rs!kdsubinstalasi = "011" Then
j = 22
ElseIf rs!kdsubinstalasi = "012" Then
j = 23
ElseIf rs!kdsubinstalasi = "013" Then
j = 24
ElseIf rs!kdsubinstalasi = "014" Then
j = 25
ElseIf rs!kdsubinstalasi = "015" Then
j = 26
ElseIf rs!kdsubinstalasi = "016" Then
j = 27
ElseIf rs!kdsubinstalasi = "017" Then
j = 28
ElseIf rs!kdsubinstalasi = "018" Then
j = 29
ElseIf rs!kdsubinstalasi = "019" Then
j = 30
ElseIf rs!kdsubinstalasi = "020" Then
j = 31
ElseIf rs!kdsubinstalasi = "021" Then
j = 32
ElseIf rs!kdsubinstalasi = "022" Then
j = 33
ElseIf rs!kdsubinstalasi = "023" Then
j = 34
ElseIf rs!kdsubinstalasi = "024" Then
j = 35
ElseIf rs!kdsubinstalasi = "025" Then
j = 36
ElseIf rs!kdsubinstalasi = "026" Then
j = 37
ElseIf rs!kdsubinstalasi = "027" Then
j = 38
ElseIf rs!kdsubinstalasi = "028" Then
j = 40
End If

Cell7 = oSheet.Cells(j, 7).Value
Cell8 = oSheet.Cells(j, 8).Value
Cell9 = oSheet.Cells(j, 9).Value
Cell10 = oSheet.Cells(j, 10).Value
Cell11 = oSheet.Cells(j, 11).Value
Cell13 = oSheet.Cells(j, 13).Value
Cell16 = oSheet.Cells(j, 16).Value
Cell17 = oSheet.Cells(j, 17).Value
Cell18 = oSheet.Cells(j, 18).Value
Cell19 = oSheet.Cells(j, 19).Value
Cell20 = oSheet.Cells(j, 20).Value

If rs!kdsubinstalasi = "001" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "002" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "003" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "004" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "005" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "006" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "007" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "008" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "009" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "010" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "011" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "012" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "013" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "014" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "015" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "016" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "017" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "018" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "019" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "020" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "021" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "022" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "023" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "024" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "025" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "026" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

ElseIf rs!kdsubinstalasi = "027" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With


ElseIf rs!kdsubinstalasi = "028" Then

With oSheet
.Cells(j, 7) = Trim(rs![3] + Cell7)
.Cells(j, 8) = Trim(rs![4] + Cell8)
.Cells(j, 9) = Trim(rs![5] + Cell9)
.Cells(j, 10) = Trim(rs![6] + Cell10)
.Cells(j, 11) = Trim(rs![7] + Cell11)
.Cells(j, 13) = Trim(rs![9] + Cell13)
.Cells(j, 16) = Trim(rs![12] + Cell16)
.Cells(j, 17) = Trim(rs![13] + Cell17)
.Cells(j, 18) = Trim(rs![14] + Cell18)
.Cells(j, 19) = Trim(rs![15] + Cell19)
.Cells(j, 20) = Trim(rs![16] + Cell20)
End With

End If
rs.MoveNext
Wend
End If
   
hell:
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

'Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then Me.cmdCetak.SetFocus
'End Sub
'
'Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then cmdCetak.SetFocus
'End Sub



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


