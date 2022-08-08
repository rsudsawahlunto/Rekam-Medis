VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm3sub01New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.1 Kegiatan Pelayanan Rawat Inap"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frm3sub01New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6135
   Begin VB.Frame Frame1 
      Height          =   1815
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
      Begin MSComCtl2.DTPicker dtptahun 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   600
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
         Format          =   126550019
         UpDown          =   -1  'True
         CurrentDate     =   40544
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
      Left            =   0
      TabIndex        =   4
      Top             =   3000
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm3sub01New2.frx":0CCA
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
      Left            =   5380
      TabIndex        =   5
      Top             =   3120
      Width           =   735
   End
End
Attribute VB_Name = "frm3sub01New2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Special Buat Excel
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel.Range
Dim oResizeRange As Excel.Range
Dim i, j, k, l, xx As Integer
Dim w, X, Y, z As String
'Untuk Pengganti Group Dijadikan Penginputan Di Cell yg sama
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
'Untuk Pengganti Group Dijadikan Penginputan Di Cell yg sama

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)

    dtptahun.value = Now
    dtptahun.CustomFormat = "yyyyy"
    dtptahun.MaxDate = Now
End Sub

Private Sub cmdCetak_Click()
'    On Error GoTo error

    ProgressBar1.value = ProgressBar1.Min
    lblPersen.Caption = "0 %"
    Screen.MousePointer = vbHourglass

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.path & "\RL 3.1_Rawat inap.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    For xx = 2 To 31
        With oSheet
            .Cells(xx, 1) = rsb("KdRS").value
            .Cells(xx, 3) = rsb("KotaKodyaKab").value
            .Cells(xx, 4) = rsb("NamaRS").value
            .Cells(xx, 5) = Format(dtptahun.value, "YYYY")
        End With
    Next xx

'    strSQL = "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienAwal as [3] WHERE Year(TglMasuk) between '" _
'    & dtptahun.Year & "' and '" & dtptahun.Year & "'  or (tglmasuk IS NULL) Union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienMasuk as [4] WHERE Year(TglMasuk) between '" _
'    & dtptahun.Year & "' and '" & dtptahun.Year & "'  or (tglmasuk IS NULL) Union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglPulang,KdSubInstalasi from LaporanRL11_PasienKeluarHidup as [5] WHERE Year(TglPulang) between '" _
'    & dtptahun.Year & "' and '" & dtptahun.Year & "'  or (tglpulang IS NULL) union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglPulang,KdSubInstalasi from LaporanRL11_PasienKeluarMati6 as [6] where Year(TglPulang) between '" _
'    & dtptahun.Year & "' and '" & dtptahun.Year & "' or (TglPulang IS NULL) union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglPulang,KdSubInstalasi from LaporanRL11_PasienKeluarMati7 as [7] where Year(TglPulang) between '" _
'    & dtptahun.Year & "' and '" & dtptahun.Year & "' or (TglPulang IS NULL) union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglPulang,KdSubInstalasi from LaporanRL11_PasienKeluarMati8 as [8] where Year(TglPulang) between '" _
'    & dtptahun.Year & "' and '" & dtptahun.Year & "' or (TglPulang IS NULL) union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL_PasienAkhirTahun as [9] where Year(TglMasuk) between '" _
'    & dtptahun.Year & "' and '" & dtptahun.Year & "' union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_JmlHariRawat as [11] where Year(TglMasuk) between '" _
'    & dtptahun.Year & "' and '" & dtptahun.Year & "' or (TglMasuk IS NULL) union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],ISNULL(DATEDIFF(day,TglMasuk, { fn NOW() }), 0) As [12],[13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienKelas as [12] where Year(TglMasuk) between '" _
'    & dtptahun.Year & "' and '" & dtptahun.Year & "' and (KdKelas in('05','05')) union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],ISNULL(DATEDIFF(day,TglMasuk, { fn NOW() }), 0) As [13],[14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienKelas as [13] where Year(TglMasuk) between '" _
'    & dtptahun.Year & "' and '" & dtptahun.Year & "' and (KdKelas in('03')) union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],ISNULL(DATEDIFF(day,TglMasuk, { fn NOW() }), 0) As [14],[15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienKelas as [14] where Year(TglMasuk) between '" _
'    & dtptahun.Year & "' and '" & dtptahun.Year & "' and (KdKelas in('02')) union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],ISNULL(DATEDIFF(day,TglMasuk, { fn NOW() }), 0) As [15],[16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienKelas as [15] where Year(TglMasuk) between '" _
'    & dtptahun.Year & "' and '" & dtptahun.Year & "' and (KdKelas in('01')) union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],ISNULL(DATEDIFF(day,TglMasuk, { fn NOW() }), 0) As [16],TglMasuk,KdSubInstalasi from LaporanRL11_PasienKelas as [16] where Year(TglMasuk) between '" & dtptahun.Year & " ' and '" & dtptahun.Year & "' and (KdKelas in('07'))"

    strSQL = "select [2],sum([3]) AS [3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan from LaporanRL11_PasienAwal as [3] WHERE ((Year(TglMasuk) <> YEAR(TglPulang)) " _
    & "AND YEAR(TglMasuk)='" & dtptahun.Year - 1 & "') OR (YEAR(TglMasuk)='" & dtptahun.Year - 1 & "' AND TglPulang IS NULL) GROUP BY [2],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan Union " _
    & "select [2],[3],SUM([4]) AS [4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan from LaporanRL11_PasienMasuk as [4] WHERE Year(TglMasuk) between '" _
    & dtptahun.Year & "' and '" & dtptahun.Year & "' GROUP BY [2],[3],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan Union " _
    & "select [2],[3],[4],SUM([5]) AS [5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan from LaporanRL11_PasienKeluarHidup as [5] WHERE Year(TglPulang) between '" _
    & dtptahun.Year & "' and '" & dtptahun.Year & "' GROUP BY [2],[3],[4],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan union " _
    & "select [2],[3],[4],[5],SUM([6]) AS [6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan from LaporanRL11_PasienKeluarMati6 as [6] where Year(TglPulang) between '" _
    & dtptahun.Year & "' and '" & dtptahun.Year & "' GROUP BY [2],[3],[4],[5],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan union " _
    & "select [2],[3],[4],[5],[6],SUM([7]) AS [7],[8],[9],[10],[11],[12],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan from LaporanRL11_PasienKeluarMati7 as [7] where Year(TglPulang) between '" _
    & dtptahun.Year & "' and '" & dtptahun.Year & "' GROUP BY [2],[3],[4],[5],[6],[8],[9],[10],[11],[12],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan union " _
    & "select [2],[3],[4],[5],[6],[7],SUM([9]) AS [8],0 AS [9],[10],[11],[12],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan from LaporanRL11_LamaDirawat as [8] where Year(TglPulang) between '" _
    & dtptahun.Year & "' and '" & dtptahun.Year & "' GROUP BY [2],[3],[4],[5],[6],[7],[9],[10],[11],[12],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan union " _
    & "select [2],[3],[4],[5],[6],[7],[8],SUM([9]) AS [9],[10],[11],[12],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan from LaporanRL_PasienAkhirTahun as [9] " _
    & "GROUP BY [2],[3],[4],[5],[6],[7],[8],[10],[11],[12],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan union " _
    & "select [2],[3],[4],[5],[6],[7],[8],[9],SUM(dbo.FB_TakeHariRawat2(NoPakai,'" & dtptahun.Year & "')) AS [10],[11],[12],[13],[14],[15],[16],kdsubinstalasi,KdJenisPelayanan from V_HariRawatRL31 where year(TglKeluar)='" & dtptahun.Year & "' or YEAR(TglMasuk)='" & dtptahun.Year & "' GROUP BY [2],[3],[4],[5],[6],[7],[8],[9],[11],[12],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan union " _
    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],case when KdKelas='05' then SUM(dbo.FB_TakeHariRawat2(NoPakai,'" & dtptahun.Year & "')) else 0 end AS [11],[12],[13],[14],[15],[16],kdsubinstalasi,KdJenisPelayanan from V_HariRawatRL31 where year(TglKeluar)='" & dtptahun.Year & "' or YEAR(TglMasuk)='" & dtptahun.Year & "' GROUP BY [2],[3],[4],[5],[6],[7],[8],[9],[10],[12],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan,KdKelas union " _
    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],case when KdKelas='06' then SUM(dbo.FB_TakeHariRawat2(NoPakai,'" & dtptahun.Year & "')) else 0 end AS [12],[13],[14],[15],[16],kdsubinstalasi,KdJenisPelayanan from V_HariRawatRL31 where year(TglKeluar)='" & dtptahun.Year & "' or YEAR(TglMasuk)='" & dtptahun.Year & "' GROUP BY [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[13],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan,KdKelas union " _
    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],case when KdKelas='03' then SUM(dbo.FB_TakeHariRawat2(NoPakai,'" & dtptahun.Year & "')) else 0 end AS [13],[14],[15],[16],kdsubinstalasi,KdJenisPelayanan from V_HariRawatRL31 where year(TglKeluar)='" & dtptahun.Year & "' or YEAR(TglMasuk)='" & dtptahun.Year & "' GROUP BY [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[14],[15],[16],KdSubInstalasi,KdJenisPelayanan,KdKelas union " _
    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],case when KdKelas='02' then SUM(dbo.FB_TakeHariRawat2(NoPakai,'" & dtptahun.Year & "')) else 0 end AS [14],[15],[16],kdsubinstalasi,KdJenisPelayanan from V_HariRawatRL31 where year(TglKeluar)='" & dtptahun.Year & "' or YEAR(TglMasuk)='" & dtptahun.Year & "' GROUP BY [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[15],[16],KdSubInstalasi,KdJenisPelayanan,KdKelas union " _
    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],case when KdKelas='01' then SUM(dbo.FB_TakeHariRawat2(NoPakai,'" & dtptahun.Year & "')) else 0 end AS [15],[16],kdsubinstalasi,KdJenisPelayanan from V_HariRawatRL31 where year(TglKeluar)='" & dtptahun.Year & "' or YEAR(TglMasuk)='" & dtptahun.Year & "' GROUP BY [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[16],KdSubInstalasi,KdJenisPelayanan,KdKelas union select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],case when KdKelas='04' then SUM(dbo.FB_TakeHariRawat2(NoPakai,'" & dtptahun.Year & "')) else 0 end AS [16],kdsubinstalasi,KdJenisPelayanan from V_HariRawatRL31 where year(TglKeluar)='" & dtptahun.Year & "' or YEAR(TglMasuk)='" & dtptahun.Year & "' GROUP BY [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],KdSubInstalasi,KdJenisPelayanan,KdKelas"
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],case when YEAR(TglMasuk)<'" & dtptahun.Year & "' then HariRawatAwal else 0 end + case when YEAR(TglKeluar)>'" & dtptahun.Year & "' then HariRawatAkhir else 0 end + HariRawat as [10],[11],[12],[13],[14],[15],[16],kdsubinstalasi,KdJenisPelayanan from V_HariRawatRL31 where year(TglKeluar)='" & dtptahun.Year & "' or YEAR(TglMasuk)='" & dtptahun.Year & "' union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],case when KdKelas='05' then (case when YEAR(TglMasuk)<'" & dtptahun.Year & "' then HariRawatAwal else 0 end + case when YEAR(TglKeluar)>'" & dtptahun.Year & "' then HariRawatAkhir else 0 end + HariRawat) else 0 end as [11],[12],[13],[14],[15],[16],kdsubinstalasi,KdJenisPelayanan from V_HariRawatRL31 where year(TglKeluar)='" & dtptahun.Year & "' or YEAR(TglMasuk)='" & dtptahun.Year & "' union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],case when KdKelas='04' then (case when YEAR(TglMasuk)<'" & dtptahun.Year & "' then HariRawatAwal else 0 end + case when YEAR(TglKeluar)>'" & dtptahun.Year & "' then HariRawatAkhir else 0 end + HariRawat) else 0 end as [12],[13],[14],[15],[16],kdsubinstalasi,KdJenisPelayanan from V_HariRawatRL31 where year(TglKeluar)='" & dtptahun.Year & "' or YEAR(TglMasuk)='" & dtptahun.Year & "' union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],case when KdKelas='03' then (case when YEAR(TglMasuk)<'" & dtptahun.Year & "' then HariRawatAwal else 0 end + case when YEAR(TglKeluar)>'" & dtptahun.Year & "' then HariRawatAkhir else 0 end + HariRawat) else 0 end as [13],[14],[15],[16],kdsubinstalasi,KdJenisPelayanan from V_HariRawatRL31 where year(TglKeluar)='" & dtptahun.Year & "' or YEAR(TglMasuk)='" & dtptahun.Year & "' union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],case when KdKelas='02' then (case when YEAR(TglMasuk)<'" & dtptahun.Year & "' then HariRawatAwal else 0 end + case when YEAR(TglKeluar)>'" & dtptahun.Year & "' then HariRawatAkhir else 0 end + HariRawat) else 0 end as [14],[15],[16],kdsubinstalasi,KdJenisPelayanan from V_HariRawatRL31 where year(TglKeluar)='" & dtptahun.Year & "' or YEAR(TglMasuk)='" & dtptahun.Year & "' union " _
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],case when KdKelas='01' then (case when YEAR(TglMasuk)<'" & dtptahun.Year & "' then HariRawatAwal else 0 end + case when YEAR(TglKeluar)>'" & dtptahun.Year & "' then HariRawatAkhir else 0 end + HariRawat) else 0 end as [15],[16],kdsubinstalasi,KdJenisPelayanan from V_HariRawatRL31 where year(TglKeluar)='" & dtptahun.Year & "' or YEAR(TglMasuk)='" & dtptahun.Year & "'"
'    & "select [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],ISNULL(COUNT([2]), 0) As [16],KdSubInstalasi,KdJenisPelayanan from LaporanRL11_PasienKelas as [16] where Year(TglMasuk) between '" & dtptahun.Year & " ' and '" & dtptahun.Year & "' and (KdKelas in('07')) GROUP BY [2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],KdSubInstalasi,KdJenisPelayanan"

    Call msubRecFO(rs, strSQL)
    If rs.RecordCount = 0 Then
        MsgBox "Data tidak ada", vbOKOnly, "Warning"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    ProgressBar1.Min = 0
    ProgressBar1.Max = rs.RecordCount
    ProgressBar1.value = 0

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF

            'Psikologi belum ada dan juga Geriatri
'            If rs!kdsubinstalasi = "001" Then     'Penyakit Dalam
'                j = 2
'            ElseIf rs!kdsubinstalasi = "002" Then 'Penyakit Bedah
'                j = 6
'            ElseIf rs!kdsubinstalasi = "003" Then 'Penyakit Anak
'                j = 3
'            ElseIf rs!kdsubinstalasi = "004" Then 'Penyakit Obstetri
'                j = 4
'            ElseIf rs!kdsubinstalasi = "005" Then 'Penyakit Ginekologi
'                j = 5
'            ElseIf rs!kdsubinstalasi = "006" Then 'Penyakit Bedah Syaraf
'                j = 8
'            ElseIf rs!kdsubinstalasi = "007" Then 'Penyakit Saraf
'                j = 10
'            ElseIf rs!kdsubinstalasi = "008" Then 'Penyakit Jiwa
'                j = 11
'            ElseIf rs!kdsubinstalasi = "009" Then 'Penyakit THT
'                j = 14
'            ElseIf rs!kdsubinstalasi = "010" Then 'Penyakit Mata
'                j = 15
'            ElseIf rs!kdsubinstalasi = "011" Then 'Penyakit Kulit & Kelamin
'                j = 16
'            ElseIf rs!kdsubinstalasi = "012" Then 'Penyakit Gigi & Mulut
'                j = 29
'            ElseIf rs!kdsubinstalasi = "013" Then 'Penyakit Kardiologi
'                j = 17
'            ElseIf rs!kdsubinstalasi = "014" Then 'Penyakit Radioterapi
'                j = 20
'            ElseIf rs!kdsubinstalasi = "015" Then 'Penyakit Bedah Ortophedi
'                j = 7
'            ElseIf rs!kdsubinstalasi = "016" Then 'Penyakit Paru-Paru
'                j = 18
'            ElseIf rs!kdsubinstalasi = "017" Then 'Penyakit Kusta
'                j = 22
'            ElseIf rs!kdsubinstalasi = "018" Then 'Penyakit Umum
'                j = 28
'            ElseIf rs!kdsubinstalasi = "019" Then 'Penyakit Rawat Darurat
'                j = 30
'            ElseIf rs!kdsubinstalasi = "020" Then 'Penyakit Rehabilitasi Medik
'                j = 23
'            ElseIf rs!kdsubinstalasi = "021" Then 'Isolasi
'                j = 24
'            ElseIf rs!kdsubinstalasi = "022" Then 'Luka Bakar
'                j = 9
'            ElseIf rs!kdsubinstalasi = "023" Then 'ICU
'                j = 25
'            ElseIf rs!kdsubinstalasi = "024" Then 'ICCU
'                j = 26
'            ElseIf rs!kdsubinstalasi = "025" Then 'NICU/PICU
'                j = 27
'            ElseIf rs!kdsubinstalasi = "026" Then 'Napza
'                j = 13
'            ElseIf rs!kdsubinstalasi = "027" Then 'Kedokteran Nuklir
'                j = 21
'            ElseIf rs!kdsubinstalasi = "028" Then 'Perinatologi
'                j = 31
'            End If

            If rs!KdJenisPelayanan = "001" Then     'Penyakit Dalam
                j = 2
            ElseIf rs!KdJenisPelayanan = "002" Then 'Penyakit Anak
                j = 3
            ElseIf rs!KdJenisPelayanan = "003" Then 'Penyakit Anak
                j = 4
            ElseIf rs!KdJenisPelayanan = "004" Then 'Penyakit Obstetri
                j = 5
            ElseIf rs!KdJenisPelayanan = "005" Then 'Penyakit Ginekologi
                j = 6
            ElseIf rs!KdJenisPelayanan = "006" Then 'Penyakit Bedah Syaraf
                j = 7
            ElseIf rs!KdJenisPelayanan = "007" Then 'Penyakit Saraf
                j = 8
            ElseIf rs!KdJenisPelayanan = "008" Then 'Penyakit Jiwa
                j = 9
            ElseIf rs!KdJenisPelayanan = "009" Then 'Penyakit THT
                j = 10
            ElseIf rs!KdJenisPelayanan = "010" Then 'Penyakit Mata
                j = 11
            ElseIf rs!KdJenisPelayanan = "011" Then 'Penyakit Kulit & Kelamin
                j = 12
            ElseIf rs!KdJenisPelayanan = "012" Then 'Penyakit Gigi & Mulut
                j = 13
            ElseIf rs!KdJenisPelayanan = "013" Then 'Penyakit Kardiologi
                j = 14
            ElseIf rs!KdJenisPelayanan = "014" Then 'Penyakit Radioterapi
                j = 15
            ElseIf rs!KdJenisPelayanan = "015" Then 'Penyakit Bedah Ortophedi
                j = 16
            ElseIf rs!KdJenisPelayanan = "016" Then 'Penyakit Paru-Paru
                j = 17
            ElseIf rs!KdJenisPelayanan = "017" Then 'Penyakit Kusta
                j = 18
            ElseIf rs!KdJenisPelayanan = "018" Then 'Penyakit Umum
                j = 19
            ElseIf rs!KdJenisPelayanan = "019" Then 'Penyakit Rawat Darurat
                j = 20
            ElseIf rs!KdJenisPelayanan = "020" Then 'Penyakit Rehabilitasi Medik
                j = 21
            ElseIf rs!KdJenisPelayanan = "021" Then 'Isolasi
                j = 22
            ElseIf rs!KdJenisPelayanan = "022" Then 'Luka Bakar
                j = 23
            ElseIf rs!KdJenisPelayanan = "023" Then 'ICU
                j = 24
            ElseIf rs!KdJenisPelayanan = "024" Then 'ICCU
                j = 25
            ElseIf rs!KdJenisPelayanan = "025" Then 'NICU/PICU
                j = 26
            ElseIf rs!KdJenisPelayanan = "026" Then 'Napza
                j = 27
            ElseIf rs!KdJenisPelayanan = "027" Then 'Kedokteran Nuklir
                j = 28
            ElseIf rs!KdJenisPelayanan = "028" Then 'Perinatologi
                j = 29
            ElseIf rs!KdJenisPelayanan = "029" Then 'Perinatologi
                j = 30
            ElseIf rs!KdJenisPelayanan = "030" Then 'Perinatologi
                j = 31
            End If
            Cell5 = oSheet.Cells(j, 8).value
            Cell6 = oSheet.Cells(j, 9).value
            Cell7 = oSheet.Cells(j, 10).value
            Cell8 = oSheet.Cells(j, 11).value
            Cell9 = oSheet.Cells(j, 12).value
            Cell10 = oSheet.Cells(j, 13).value
            Cell11 = oSheet.Cells(j, 14).value
            Cell12 = oSheet.Cells(j, 15).value
            Cell13 = oSheet.Cells(j, 16).value
            Cell14 = oSheet.Cells(j, 17).value
            Cell15 = oSheet.Cells(j, 18).value
            Cell16 = oSheet.Cells(j, 19).value
            Cell17 = oSheet.Cells(j, 20).value
            Cell18 = oSheet.Cells(j, 21).value

            With oSheet
                .Cells(j, 8) = Trim(rs![3] + Cell5)
                .Cells(j, 9) = Trim(rs![4] + Cell6)
                .Cells(j, 10) = Trim(rs![5] + Cell7)
                .Cells(j, 11) = Trim(rs![6] + Cell8)
                .Cells(j, 12) = Trim(rs![7] + Cell9)
                .Cells(j, 13) = Trim(rs![8] + Cell10)
                .Cells(j, 14) = Trim(rs![9] + Cell11)
                .Cells(j, 15) = Trim(rs![10] + Cell12)
                .Cells(j, 16) = Trim(rs![11] + Cell13)
                .Cells(j, 17) = Trim(rs![12] + Cell14)
                .Cells(j, 18) = Trim(rs![13] + Cell15)
                .Cells(j, 19) = Trim(rs![14] + Cell16)
                .Cells(j, 20) = Trim(rs![15] + Cell17)
                .Cells(j, 21) = Trim(rs![16] + Cell18)
            End With

            rs.MoveNext

            ProgressBar1.value = ProgressBar1.value + 1
            lblPersen.Caption = Int(ProgressBar1.value * 100 / ProgressBar1.Max) & " %"

        Wend

    End If

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
''error:
''    MsgBox "Data Tidak Ada", vbInformation, "Validasi"
''    Screen.MousePointer = vbDefault
End Sub

