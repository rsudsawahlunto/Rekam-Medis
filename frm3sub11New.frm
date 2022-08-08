VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm3sub11New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.11 Kegiatan Kesehatan Jiwa"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6135
   Icon            =   "frm3sub11New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6135
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   6135
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
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   132907011
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
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
      TabIndex        =   5
      Top             =   2640
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   17
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
      Left            =   5520
      TabIndex        =   6
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm3sub11New.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frm3sub11New"
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
Dim j As Integer
Dim Cell1 As String

'Special Buat Excel
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

    ProgressBar1.value = ProgressBar1.Min
    lblPersen.Caption = "0 %"
    Screen.MousePointer = vbHourglass

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.Path & "\Formulir RL 3.11.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With oSheet
        .Cells(7, 4) = rsb("KdRS").value
        .Cells(8, 4) = rsb("NamaRS").value
        .Cells(9, 4) = Right(dtpAwal.value, 4)
    End With

    Set rsx = Nothing

    strSQL = "Select * from RL3_11New where TglMasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "'or tglmasuk is null"
    Call msubRecFO(rsx, strSQL)

    If rsx.RecordCount > 0 Then
        rsx.MoveFirst

        While Not rsx.EOF

            'here
            If rsx![JenisPelayanan] = "Psikotest" Then
                j = 13
            ElseIf rsx![JenisPelayanan] = "Konsultasi" Then
                j = 14
            ElseIf rsx![JenisPelayanan] = "Terapi Medikamentosa" Then
                j = 15
            ElseIf rsx![JenisPelayanan] = "Elektro Medik" Then
                j = 16
            ElseIf rsx![JenisPelayanan] = "Psikoterapi" Then
                j = 17
            ElseIf rsx![JenisPelayanan] = "Play Therapy" Then
                j = 18
            ElseIf rsx![JenisPelayanan] = "Rehabilitasi Medik Psikiatrik" Then
                j = 19
            End If

            Cell1 = oSheet.Cells(j, 8).value

            If rsx![JenisPelayanan] = "Psikotest" Then
                With oSheet
                    .Cells(j, 8) = Trim(rsx![JmlKunjungan] + Cell1)
                End With
            ElseIf rsx![JenisPelayanan] = "Konsultasi" Then
                With oSheet
                    .Cells(j, 8) = Trim(rsx![JmlKunjungan] + Cell1)
                End With
            ElseIf rsx![JenisPelayanan] = "Terapi Medikamentosa" Then
                With oSheet
                    .Cells(j, 8) = Trim(rsx![JmlKunjungan] + Cell1)
                End With
            ElseIf rsx![JenisPelayanan] = "Elektro Medik" Then
                With oSheet
                    .Cells(j, 8) = Trim(rsx![JmlKunjungan] + Cell1)
                End With
            ElseIf rsx![JenisPelayanan] = "Psikoterapi" Then
                With oSheet
                    .Cells(j, 8) = Trim(rsx![JmlKunjungan] + Cell1)
                End With
            ElseIf rsx![JenisPelayanan] = "Play Therapy" Then
                With oSheet
                    .Cells(j, 8) = Trim(rsx![JmlKunjungan] + Cell1)
                End With
            ElseIf rsx![JenisPelayanan] = "Rehabilitasi Medik Psikiatrik" Then
                With oSheet
                    .Cells(j, 8) = Trim(rsx![JmlKunjungan] + Cell1)
                End With
            End If

            rsx.MoveNext
        Wend
    End If

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    MsgBox "Data Tidak Ada", vbInformation, "Validasi"
    Screen.MousePointer = vbDefault
End Sub

