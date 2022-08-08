VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm3sub05New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.05 Kegiatan Perinatologi"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6405
   Icon            =   "frm3sub05New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6405
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   6375
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   1320
         Width           =   1905
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   1320
         Width           =   1935
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
         Format          =   127401987
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   17
      Scrolling       =   1
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
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
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm3sub05New2.frx":0CCA
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
      Left            =   5760
      TabIndex        =   4
      Top             =   3120
      Width           =   615
   End
End
Attribute VB_Name = "frm3sub05New2"
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
Dim i As Integer
Dim j As Integer
Dim xx As Integer

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)

    dtptahun.value = Now
    dtptahun.CustomFormat = "yyyyy"
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo error
    'here
    j = 0
    ProgressBar1.value = ProgressBar1.Min
    lblPersen.Caption = "0 %"
    Screen.MousePointer = vbHourglass

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.Path & "\RL 3.5_perinatologi.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    For xx = 2 To 16
        With oSheet
            .Cells(xx, 1) = rsb("KodeExternal").value
            .Cells(xx, 2) = rsb("KotaKodyaKab").value
            .Cells(xx, 3) = rsb("KdRS").value
            .Cells(xx, 4) = rsb("NamaRS").value
            .Cells(xx, 5) = Format(dtptahun.value, "YYYY")
        End With
    Next xx

    Set rsx = Nothing
    strSQL = "Select Judul, SUM (Jml1) as Jml1, SUM(Jml2) as Jml2, SUM(Jml3) as Jml3, sum(Jml4) as Jml4, sum(Jml5) as Jml5, sum(Jml6) as Jml6, sum(Jml7) as Jml7, sum(Jml8) as Jml8, sum(Jml9) as Jml9, sum(Jml10) as Jml10, sum(Jml11) as Jml11,sum(Jml12) as Jml12" & _
    " from RL3_05New where Year(TglLahir) = '" & dtptahun.Year & "' Group by Judul"
    Call msubRecFO(rsx, strSQL)

    rsx.MoveFirst

    Set rs1 = Nothing

    strSQL1 = "Select Judul, KdRujukanAsal, SUM (Jml1) as Jml1, SUM(Jml2) as Jml2, SUM(Jml3) as Jml3, sum(Jml4) as Jml4, sum(Jml5) as Jml5, sum(Jml6) as Jml6, sum(Jml7) as Jml7, sum(Jml8) as Jml8, sum(Jml9) as Jml9, sum(Jml10) as Jml10, sum(Jml11) as Jml11,sum(Jml12) as Jml12" & _
    " from RL3_05New where Year(TglLahir) = '" & dtptahun.Year & "' Group by Judul, KdRujukanAsal"
    Call msubRecFO(rs1, strSQL1)

    ProgressBar1.Min = 0
    ProgressBar1.Max = rsx.RecordCount
    ProgressBar1.value = 0

    rs1.MoveFirst

    For i = 1 To rsx.RecordCount
        j = j + 1

        With oSheet
            If rsx!Judul = "LahirHidup1" Then  'Bayi Lahir Hidup <= 2500 gram
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(3, 8) = Trim(IIf(IsNull(rsx!Jml1.value), "", (rsx!Jml1.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(3, 9) = Trim(IIf(IsNull(rsx!Jml1.value), "", (rsx!Jml1.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(3, 10) = Trim(IIf(IsNull(rsx!Jml1.value), "", (rsx!Jml1.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(3, 11) = Trim(IIf(IsNull(rsx!Jml1.value), "", (rsx!Jml1.value)))
                End If
                'Bayi Lahir Hidup <= 2500 gram
            ElseIf rsx!Judul = "LahirHidup2" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(4, 8) = Trim(IIf(IsNull(rsx!Jml2.value), "", (rsx!Jml2.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(4, 9) = Trim(IIf(IsNull(rsx!Jml2.value), "", (rsx!Jml2.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(4, 10) = Trim(IIf(IsNull(rsx!Jml2.value), "", (rsx!Jml2.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(4, 11) = Trim(IIf(IsNull(rsx!Jml2.value), "", (rsx!Jml2.value)))
                End If
                'Kelahiran Mati
            ElseIf rsx!Judul = "LahirMati1" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(6, 8) = Trim(IIf(IsNull(rsx!Jml3.value), "", (rsx!Jml3.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(6, 9) = Trim(IIf(IsNull(rsx!Jml3.value), "", (rsx!Jml3.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(6, 10) = Trim(IIf(IsNull(rsx!Jml3.value), "", (rsx!Jml3.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(6, 11) = Trim(IIf(IsNull(rsx!Jml3.value), "", (rsx!Jml3.value)))
                End If
                'Mati Neonatal < 7 hr
            ElseIf rsx!Judul = "LahirMati2" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(7, 8) = Trim(IIf(IsNull(rsx!Jml4.value), "", (rsx!Jml4.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(7, 9) = Trim(IIf(IsNull(rsx!Jml4.value), "", (rsx!Jml4.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(7, 10) = Trim(IIf(IsNull(rsx!Jml4.value), "", (rsx!Jml4.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(7, 11) = Trim(IIf(IsNull(rsx!Jml4.value), "", (rsx!Jml4.value)))
                End If
                'Sebab Kematian Asphyxia
            ElseIf rsx!Judul = "LahirMati3" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(9, 8) = Trim(IIf(IsNull(rsx!Jml5.value), "", (rsx!Jml5.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(9, 9) = Trim(IIf(IsNull(rsx!Jml5.value), "", (rsx!Jml5.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(9, 10) = Trim(IIf(IsNull(rsx!Jml5.value), "", (rsx!Jml5.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(9, 11) = Trim(IIf(IsNull(rsx!Jml5.value), "", (rsx!Jml5.value)))
                End If
                'Sebab Kematian Trauma Kelahiran
            ElseIf rsx!Judul = "LahirMati4" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(10, 8) = Trim(IIf(IsNull(rsx!Jml6.value), "", (rsx!Jml6.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(10, 9) = Trim(IIf(IsNull(rsx!Jml6.value), "", (rsx!Jml6.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(10, 10) = Trim(IIf(IsNull(rsx!Jml6.value), "", (rsx!Jml6.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(10, 11) = Trim(IIf(IsNull(rsx!Jml6.value), "", (rsx!Jml6.value)))
                End If
                'Sebab Kematian Trauma BBLR
            ElseIf rsx!Judul = "LahirMati5" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(11, 8) = Trim(IIf(IsNull(rsx!Jml7.value), "", (rsx!Jml7.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(11, 9) = Trim(IIf(IsNull(rsx!Jml7.value), "", (rsx!Jml7.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(11, 10) = Trim(IIf(IsNull(rsx!Jml7.value), "", (rsx!Jml7.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(11, 11) = Trim(IIf(IsNull(rsx!Jml7.value), "", (rsx!Jml7.value)))
                End If
                'Sebab Kematian Tetanus Neonatorum
            ElseIf rsx!Judul = "LahirMati6" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(12, 8) = Trim(IIf(IsNull(rsx!Jml8.value), "", (rsx!Jml8.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(12, 9) = Trim(IIf(IsNull(rsx!Jml8.value), "", (rsx!Jml8.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(12, 10) = Trim(IIf(IsNull(rsx!Jml8.value), "", (rsx!Jml8.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(12, 11) = Trim(IIf(IsNull(rsx!Jml8.value), "", (rsx!Jml8.value)))
                End If
                'Sebab Kematian Kelainan Congenital
            ElseIf rsx!Judul = "LahirMati7" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(13, 8) = Trim(IIf(IsNull(rsx!Jml9.value), "", (rsx!Jml9.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(13, 9) = Trim(IIf(IsNull(rsx!Jml9.value), "", (rsx!Jml9.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(13, 10) = Trim(IIf(IsNull(rsx!Jml9.value), "", (rsx!Jml9.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(13, 11) = Trim(IIf(IsNull(rsx!Jml9.value), "", (rsx!Jml9.value)))
                End If
                'Sebab Kematian Kelainan ISPA
            ElseIf rsx!Judul = "LahirMati8" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(14, 8) = Trim(IIf(IsNull(rsx!Jml10.value), "", (rsx!Jml10.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(14, 9) = Trim(IIf(IsNull(rsx!Jml10.value), "", (rsx!Jml10.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(14, 10) = Trim(IIf(IsNull(rsx!Jml10.value), "", (rsx!Jml10.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(14, 11) = Trim(IIf(IsNull(rsx!Jml10.value), "", (rsx!Jml10.value)))
                End If
                'Sebab Kematian Kelainan Diare
            ElseIf rsx!Judul = "LahirMati9" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(15, 8) = Trim(IIf(IsNull(rsx!Jml11.value), "", (rsx!Jml11.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(15, 9) = Trim(IIf(IsNull(rsx!Jml11.value), "", (rsx!Jml11.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(15, 10) = Trim(IIf(IsNull(rsx!Jml11.value), "", (rsx!Jml11.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(15, 11) = Trim(IIf(IsNull(rsx!Jml11.value), "", (rsx!Jml11.value)))
                End If
                'Sebab Kematian Lain - Lain
            ElseIf rsx!Judul = "LahirMati10" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(16, 8) = Trim(IIf(IsNull(rsx!Jml12.value), "", (rsx!Jml12.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(16, 9) = Trim(IIf(IsNull(rsx!Jml12.value), "", (rsx!Jml12.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(16, 10) = Trim(IIf(IsNull(rsx!Jml12.value), "", (rsx!Jml12.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(16, 11) = Trim(IIf(IsNull(rsx!Jml12.value), "", (rsx!Jml12.value)))
                End If
            End If
        End With
        rsx.MoveNext
        ProgressBar1.value = Int(ProgressBar1.value) + 1
        lblPersen.Caption = Int(ProgressBar1.value * 100 / ProgressBar1.Max) & " %"
    Next i

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    MsgBox "Data Tidak Ada", vbInformation, "Validasi"
    Screen.MousePointer = vbDefault
End Sub

