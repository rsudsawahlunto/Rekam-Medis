VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm3sub01New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.1 Kegiatan Pelayanan Rawat Inap"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frm3sub01New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6135
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   6135
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   1320
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   720
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
         Format          =   133562371
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   720
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
         Format          =   133562371
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin VB.Label Label1 
         Caption         =   "s/d"
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   840
         Width           =   375
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm3sub01New.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
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
      Left            =   4560
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frm3sub01New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel.Range
Dim oResizeRange As Excel.Range
Dim i, j, k, l As Integer
Dim w, X, y, z As String
Dim Cell5 As String
Dim Cell6 As String
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

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpAwal.value = Format(Now, "dd/mm/yyyy")
    dtpAwal.value = Format("01/01")
    dtpAkhir.value = Format(Now, "dd/mm/yyyy")

    ProgressBar1.value = ProgressBar1.Min
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo error

    ProgressBar1.value = ProgressBar1.Min
    lblPersen.Caption = "0 %"
    Screen.MousePointer = vbHourglass

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.Path & "\Formulir RL 3.1.xlsx")
    Set oSheet = oWB.ActiveSheet

    Set rsx = Nothing

    strSQL = "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienAwal as [3] WHERE  (TglMasuk BETWEEN DateAdd(MONTH,-3,'" & Format(dtpAwal, "yyyy/MM/dd ") & "') AND DATEADD(MONTH,-3,'" & Format(dtpAkhir, "yyyy/MM/dd") & "')) or (tglmasuk IS NULL)  " & _
    "Union " & _
    "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienMasuk as [4] WHERE   TglMasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "'  AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' or (tglmasuk IS NULL)   " & _
    "Union " & _
    "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglPulang,KdSubInstalasi from LaporanRL11_PasienKeluarHidup as [5] WHERE   TglPulang BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' or (TglPulang IS NULL)  " & _
    "Union " & _
    "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglPulang,KdSubInstalasi from LaporanRL11_PasienKeluarMati6 as [6] where   TglPulang BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' or (TglPulang IS NULL)  " & _
    "Union " & _
    "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglPulang,KdSubInstalasi from LaporanRL11_PasienKeluarMati7 as [7] where   TglPulang BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' or (TglPulang IS NULL)  " & _
    "Union " & _
    "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglPulang,KdSubInstalasi from LaporanRL11_PasienKeluarMati8 as [8] where   TglPulang BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "' or (TglPulang IS NULL)  " & _
    "Union " & _
    "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL_PasienAkhirTahun as [9] where   TglMasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd") & "'" & _
    "Union " & _
    "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_JmlHariRawatNew as [10] where   TglMasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "'  AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' or (TglMasuk BETWEEN DateAdd(MONTH,-3,'" & Format(dtpAwal, "yyyy/MM/dd ") & "') AND DATEADD(MONTH,-3,'" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "')) or (tglmasuk IS NULL)  " & _
    "Union " & _
    "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],ISNULL(DATEDIFF(day,TglMasuk, { fn NOW() }), 0) As [12],[13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienKelas as [12] where TglMasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and (KdKelas in('05','05')) or (TglMasuk BETWEEN DateAdd(MONTH,-3,'" & Format(dtpAwal, "yyyy/MM/dd ") & "') AND DATEADD(MONTH,-3,'" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "')) and (KdKelas in('05','05'))  " & _
    "Union " & _
    "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],ISNULL(DATEDIFF(day,TglMasuk, { fn NOW() }), 0) As [13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienKelas as [13] where TglMasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and (KdKelas in('03')) or (TglMasuk BETWEEN DateAdd(MONTH,-3,'" & Format(dtpAwal, "yyyy/MM/dd 23:59:59") & "') AND DATEADD(MONTH,-3,'" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "')) and (KdKelas in('03'))  " & _
    "Union " & _
    "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],ISNULL(DATEDIFF(day,TglMasuk, { fn NOW() }), 0) As [14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienKelas as [14] where TglMasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and (KdKelas in('02')) or (TglMasuk BETWEEN DateAdd(MONTH,-3,'" & Format(dtpAwal, "yyyy/MM/dd ") & "') AND DATEADD(MONTH,-3,'" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "')) and (KdKelas in('02'))  " & _
    "Union " & _
    "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],ISNULL(DATEDIFF(day,TglMasuk, { fn NOW() }), 0) As [15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienKelas as [15] where TglMasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and (KdKelas in('01')) or (TglMasuk BETWEEN DateAdd(MONTH,-3,'" & Format(dtpAwal, "yyyy/MM/dd ") & "') AND DATEADD(MONTH,-3,'" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "')) and (KdKelas in('01'))  " & _
    "Union " & _
    "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],ISNULL(DATEDIFF(day,TglMasuk, { fn NOW() }), 0) As [16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienKelas as [16] where TglMasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd ") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and (KdKelas in('07')) or (TglMasuk BETWEEN DateAdd(MONTH,-3,'" & Format(dtpAwal, "yyyy/MM/dd ") & "') AND DATEADD(MONTH,-3,'" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "')) and (KdKelas in('07')) "

    Call msubRecFO(rs, strSQL)

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF

            'Psikologi belum ada dan juga Geriatri
            If rs!kdsubinstalasi = "001" Then
                j = 15
            ElseIf rs!kdsubinstalasi = "002" Then
                j = 19
            ElseIf rs!kdsubinstalasi = "003" Then
                j = 16
            ElseIf rs!kdsubinstalasi = "004" Then
                j = 17
            ElseIf rs!kdsubinstalasi = "005" Then
                j = 18
            ElseIf rs!kdsubinstalasi = "006" Then
                j = 21
            ElseIf rs!kdsubinstalasi = "007" Then
                j = 23
            ElseIf rs!kdsubinstalasi = "008" Then
                j = 24
            ElseIf rs!kdsubinstalasi = "009" Then
                j = 27
            ElseIf rs!kdsubinstalasi = "010" Then
                j = 28
            ElseIf rs!kdsubinstalasi = "011" Then
                j = 29
            ElseIf rs!kdsubinstalasi = "012" Then
                j = 42
            ElseIf rs!kdsubinstalasi = "013" Then
                j = 30
            ElseIf rs!kdsubinstalasi = "014" Then
                j = 33
            ElseIf rs!kdsubinstalasi = "015" Then
                j = 20
            ElseIf rs!kdsubinstalasi = "016" Then
                j = 31
            ElseIf rs!kdsubinstalasi = "017" Then
                j = 35
            ElseIf rs!kdsubinstalasi = "018" Then
                j = 41
            ElseIf rs!kdsubinstalasi = "019" Then
                j = 43
            ElseIf rs!kdsubinstalasi = "020" Then
                j = 36
            ElseIf rs!kdsubinstalasi = "021" Then
                j = 37
            ElseIf rs!kdsubinstalasi = "022" Then
                j = 22
            ElseIf rs!kdsubinstalasi = "023" Then
                j = 38
            ElseIf rs!kdsubinstalasi = "024" Then
                j = 39
            ElseIf rs!kdsubinstalasi = "025" Then
                j = 40
            ElseIf rs!kdsubinstalasi = "026" Then
                j = 26
            ElseIf rs!kdsubinstalasi = "027" Then
                j = 34
            ElseIf rs!kdsubinstalasi = "028" Then
                j = 45
            End If

            'Bagian ini yg belum'

            Cell5 = oSheet.Cells(j, 5).value
            Cell6 = oSheet.Cells(j, 6).value
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

            If rs!kdsubinstalasi = "001" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "002" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "003" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "004" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "005" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "006" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "007" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "008" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "009" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "010" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "011" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "012" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "013" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "014" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "015" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "016" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "017" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "018" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "019" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "020" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "021" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "022" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "023" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "024" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "025" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "026" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "027" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            ElseIf rs!kdsubinstalasi = "028" Then

                With oSheet
                    .Cells(j, 5) = Trim(rs![3] + Cell5)
                    .Cells(j, 6) = Trim(rs![4] + Cell6)
                    .Cells(j, 7) = Trim(rs![5] + Cell7)
                    .Cells(j, 8) = Trim(rs![6] + Cell8)
                    .Cells(j, 9) = Trim(rs![7] + Cell9)
                    .Cells(j, 10) = Trim(rs![8] + Cell10)
                    .Cells(j, 11) = Trim(rs![9] + Cell11)
                    .Cells(j, 12) = Trim(rs![10] + Cell12)
                    .Cells(j, 13) = Trim(rs![11] + Cell13)
                    .Cells(j, 14) = Trim(rs![12] + Cell14)
                    .Cells(j, 15) = Trim(rs![13] + Cell15)
                    .Cells(j, 16) = Trim(rs![14] + Cell16)
                    .Cells(j, 17) = Trim(rs![15] + Cell17)
                    .Cells(j, 18) = Trim(rs![16] + Cell18)
                End With

            End If

            rs.MoveNext

        Wend
    End If

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With oSheet
        .Cells(7, 4) = rsb("KdRS").value
        .Cells(8, 4) = rsb("NamaRS").value
        .Cells(9, 4) = Right(dtpAwal.value, 4)
    End With

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    MsgBox "Data Tidak Ada", vbInformation, "Validasi"
    Screen.MousePointer = vbDefault
End Sub

