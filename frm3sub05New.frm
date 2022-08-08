VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm3sub05New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.05 Kegiatan Perinatologi"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6405
   Icon            =   "frm3sub05New.frx":0000
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
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   133103619
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   133103619
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin VB.Label Label1 
         Caption         =   "s/d"
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   4
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
      TabIndex        =   8
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
      Picture         =   "frm3sub05New.frx":0CCA
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
      TabIndex        =   5
      Top             =   3120
      Width           =   615
   End
End
Attribute VB_Name = "frm3sub05New"
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

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpAwal.value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.value = Now

    ProgressBar1.value = ProgressBar1.Min
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
    Set oWB = oXL.Workbooks.Open(App.Path & "\Formulir RL 3.5.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With oSheet
        .Cells(5, 4) = rsb("KdRS").value
        .Cells(6, 4) = rsb("NamaRS").value
        .Cells(7, 4) = Right(dtpAwal.value, 4)
    End With

    Set rsx = Nothing
    strSQL = "Select Judul, SUM (Jml1) as Jml1, SUM(Jml2) as Jml2, SUM(Jml3) as Jml3, sum(Jml4) as Jml4, sum(Jml5) as Jml5, sum(Jml6) as Jml6, sum(Jml7) as Jml7, sum(Jml8) as Jml8, sum(Jml9) as Jml9, sum(Jml10) as Jml10, sum(Jml11) as Jml11,sum(Jml12) as Jml12" & _
    " from RL3_05New where TglLahir between '" & Format(dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "' Group by Judul"
    Call msubRecFO(rsx, strSQL)

    rsx.MoveFirst

    Set rs1 = Nothing

    strSQL1 = "Select Judul, KdRujukanAsal, SUM (Jml1) as Jml1, SUM(Jml2) as Jml2, SUM(Jml3) as Jml3, sum(Jml4) as Jml4, sum(Jml5) as Jml5, sum(Jml6) as Jml6, sum(Jml7) as Jml7, sum(Jml8) as Jml8, sum(Jml9) as Jml9, sum(Jml10) as Jml10, sum(Jml11) as Jml11,sum(Jml12) as Jml12" & _
    " from RL3_05New where TglLahir between '" & Format(dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "' Group by Judul, KdRujukanAsal"

    Call msubRecFO(rs1, strSQL1)

    rs1.MoveFirst

    For i = 1 To rsx.RecordCount
        j = j + 1
        ProgressBar1.Max = rsx.RecordCount

        With oSheet
            If rsx!Judul = "LahirHidup1" Then  'Bayi Lahir Hidup <= 2500 gram
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(15, 5) = Trim(IIf(IsNull(rsx!Jml1.value), "", (rsx!Jml1.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(15, 6) = Trim(IIf(IsNull(rsx!Jml1.value), "", (rsx!Jml1.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(15, 7) = Trim(IIf(IsNull(rsx!Jml1.value), "", (rsx!Jml1.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(15, 8) = Trim(IIf(IsNull(rsx!Jml1.value), "", (rsx!Jml1.value)))
                End If
                'Bayi Lahir Hidup <= 2500 gram
            ElseIf rsx!Judul = "LahirHidup2" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(16, 5) = Trim(IIf(IsNull(rsx!Jml2.value), "", (rsx!Jml2.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(16, 6) = Trim(IIf(IsNull(rsx!Jml2.value), "", (rsx!Jml2.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(16, 7) = Trim(IIf(IsNull(rsx!Jml2.value), "", (rsx!Jml2.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(16, 8) = Trim(IIf(IsNull(rsx!Jml2.value), "", (rsx!Jml2.value)))
                End If
                'Kelahiran Mati
            ElseIf rsx!Judul = "LahirMati1" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(18, 5) = Trim(IIf(IsNull(rsx!Jml3.value), "", (rsx!Jml3.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(18, 6) = Trim(IIf(IsNull(rsx!Jml3.value), "", (rsx!Jml3.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(18, 7) = Trim(IIf(IsNull(rsx!Jml3.value), "", (rsx!Jml3.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(18, 8) = Trim(IIf(IsNull(rsx!Jml3.value), "", (rsx!Jml3.value)))
                End If
                'Mati Neonatal < 7 hr
            ElseIf rsx!Judul = "LahirMati2" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(19, 5) = Trim(IIf(IsNull(rsx!Jml4.value), "", (rsx!Jml4.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(19, 6) = Trim(IIf(IsNull(rsx!Jml4.value), "", (rsx!Jml4.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(19, 7) = Trim(IIf(IsNull(rsx!Jml4.value), "", (rsx!Jml4.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(19, 8) = Trim(IIf(IsNull(rsx!Jml4.value), "", (rsx!Jml4.value)))
                End If
                'Sebab Kematian Asphyxia
            ElseIf rsx!Judul = "LahirMati3" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(21, 5) = Trim(IIf(IsNull(rsx!Jml5.value), "", (rsx!Jml5.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(21, 6) = Trim(IIf(IsNull(rsx!Jml5.value), "", (rsx!Jml5.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(21, 7) = Trim(IIf(IsNull(rsx!Jml5.value), "", (rsx!Jml5.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(21, 8) = Trim(IIf(IsNull(rsx!Jml5.value), "", (rsx!Jml5.value)))
                End If
                'Sebab Kematian Trauma Kelahiran
            ElseIf rsx!Judul = "LahirMati4" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(22, 5) = Trim(IIf(IsNull(rsx!Jml6.value), "", (rsx!Jml6.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(22, 6) = Trim(IIf(IsNull(rsx!Jml6.value), "", (rsx!Jml6.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(22, 7) = Trim(IIf(IsNull(rsx!Jml6.value), "", (rsx!Jml6.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(22, 8) = Trim(IIf(IsNull(rsx!Jml6.value), "", (rsx!Jml6.value)))
                End If
                'Sebab Kematian Trauma BBLR
            ElseIf rsx!Judul = "LahirMati5" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(23, 5) = Trim(IIf(IsNull(rsx!Jml7.value), "", (rsx!Jml7.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(23, 6) = Trim(IIf(IsNull(rsx!Jml7.value), "", (rsx!Jml7.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(23, 7) = Trim(IIf(IsNull(rsx!Jml7.value), "", (rsx!Jml7.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(23, 8) = Trim(IIf(IsNull(rsx!Jml7.value), "", (rsx!Jml7.value)))
                End If
                'Sebab Kematian Tetanus Neonatorum
            ElseIf rsx!Judul = "LahirMati6" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(24, 5) = Trim(IIf(IsNull(rsx!Jml8.value), "", (rsx!Jml8.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(24, 6) = Trim(IIf(IsNull(rsx!Jml8.value), "", (rsx!Jml8.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(24, 7) = Trim(IIf(IsNull(rsx!Jml8.value), "", (rsx!Jml8.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(24, 8) = Trim(IIf(IsNull(rsx!Jml8.value), "", (rsx!Jml8.value)))
                End If
                'Sebab Kematian Kelainan Congenital
            ElseIf rsx!Judul = "LahirMati7" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(25, 5) = Trim(IIf(IsNull(rsx!Jml9.value), "", (rsx!Jml9.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(25, 6) = Trim(IIf(IsNull(rsx!Jml9.value), "", (rsx!Jml9.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(25, 7) = Trim(IIf(IsNull(rsx!Jml9.value), "", (rsx!Jml9.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(25, 8) = Trim(IIf(IsNull(rsx!Jml9.value), "", (rsx!Jml9.value)))
                End If
                'Sebab Kematian Kelainan ISPA
            ElseIf rsx!Judul = "LahirMati8" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(26, 5) = Trim(IIf(IsNull(rsx!Jml10.value), "", (rsx!Jml10.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(26, 6) = Trim(IIf(IsNull(rsx!Jml10.value), "", (rsx!Jml10.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(26, 7) = Trim(IIf(IsNull(rsx!Jml10.value), "", (rsx!Jml10.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(26, 8) = Trim(IIf(IsNull(rsx!Jml10.value), "", (rsx!Jml10.value)))
                End If
                'Sebab Kematian Kelainan Diare
            ElseIf rsx!Judul = "LahirMati9" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(27, 5) = Trim(IIf(IsNull(rsx!Jml11.value), "", (rsx!Jml11.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(27, 6) = Trim(IIf(IsNull(rsx!Jml11.value), "", (rsx!Jml11.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(27, 7) = Trim(IIf(IsNull(rsx!Jml11.value), "", (rsx!Jml11.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(27, 8) = Trim(IIf(IsNull(rsx!Jml11.value), "", (rsx!Jml11.value)))
                End If
                'Sebab Kematian Lain - Lain
            ElseIf rsx!Judul = "LahirMati10" Then
                If rs1!KdRujukanAsal = "03" Or rs1!KdRujukanAsal = "04" Then 'RS Pemerintah & RS Swasta
                    .Cells(28, 5) = Trim(IIf(IsNull(rsx!Jml12.value), "", (rsx!Jml12.value)))
                ElseIf rs1!KdRujukanAsal = "13" Then    'Bidan
                    .Cells(28, 6) = Trim(IIf(IsNull(rsx!Jml12.value), "", (rsx!Jml12.value)))
                ElseIf rs1!KdRujukanAsal = "02" Then  'Puskesmas
                    .Cells(28, 7) = Trim(IIf(IsNull(rsx!Jml12.value), "", (rsx!Jml12.value)))
                ElseIf rs1!KdRujukanAsal = "14" Then  'Faskes
                    .Cells(28, 8) = Trim(IIf(IsNull(rsx!Jml12.value), "", (rsx!Jml12.value)))
                End If

            End If

        End With

        rsx.MoveNext
        ProgressBar1.value = Int(ProgressBar1.value) + 1
        lblPersen.Caption = Int(ProgressBar1.value / rsx.RecordCount * 100) & " %"
    Next i

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    MsgBox "Data Tidak Ada", vbInformation, "Validasi"
    Screen.MousePointer = vbDefault
End Sub

