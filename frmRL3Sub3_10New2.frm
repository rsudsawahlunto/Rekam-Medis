VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRL3Sub3_10New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.10 Kegiatan Pelayanan Khusus"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3525
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
      Begin MSComCtl2.DTPicker dtptahun 
         Height          =   375
         Left            =   1920
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
         Format          =   154533891
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
      Top             =   3120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   17
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
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
      Left            =   5160
      TabIndex        =   5
      Top             =   3190
      Width           =   615
   End
End
Attribute VB_Name = "frmRL3Sub3_10New2"
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
Dim j, xx As Integer
Dim Cell1 As Integer

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

    ProgressBar1.value = ProgressBar1.Min
    lblPersen.Caption = "0 %"
    Screen.MousePointer = vbHourglass

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    '    oXL.Visible = True
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.path & "\RL 3.10_pelayanan khusus.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    For xx = 2 To 15
        With oSheet
            .Cells(xx, 3) = rsb("KdRS").value
            .Cells(xx, 2) = rsb("KotaKodyaKab").value
            .Cells(xx, 4) = rsb("NamaRS").value
            .Cells(xx, 5) = Format(dtptahun.value, "YYYY")
        End With
    Next xx

    Set rsx = Nothing

    strSQL = "Select distinct * from RL3_10New where Year(TglPelayanan) = '" & dtptahun.Year & "'"
    Call msubRecFO(rs, strSQL)

    ProgressBar1.Min = 0
    ProgressBar1.Max = rs.RecordCount
    ProgressBar1.value = 0

    If rs.RecordCount > 0 Then

        rs.MoveFirst

        While Not rs.EOF
'            If rs!JenisKegiatan = "Electro Kardiographi (EKG)" Then
            If rs!JenisKegiatan = "Elektro Kardiographi (EKG)" Then
                j = 2
            ElseIf rs!JenisKegiatan = "Endoskopi (semua bentuk)" Then
                j = 5
            ElseIf rs!JenisKegiatan = "Hemodialisa" Then
                j = 6
            ElseIf rs!JenisKegiatan = "Densometri Tulang" Then
                j = 7
            ElseIf rs!JenisKegiatan = "Pungsi" Then
                j = 8
            ElseIf rs!JenisKegiatan = "Spirometri" Then
                j = 9
            ElseIf rs!JenisKegiatan = "Tes Kulit/Alergi/Histamin" Then
                j = 10
            ElseIf rs!JenisKegiatan = "Topometri" Then
                j = 11
            ElseIf rs!JenisKegiatan = "Lain-lain" Then
                j = 15
            End If
            
            If IsNull(rs("JenisKegiatan").value) = True Then
              Cell1 = oSheet.Cells(2, 8).value
            Else
              Cell1 = oSheet.Cells(j, 8).value
            End If

       

            If rs!JenisKegiatan = "Elektro Kardiographi (EKG)" Then
                With oSheet
                    .Cells(j, 8) = Trim(rs!Jumlah + Cell1)
                End With
            ElseIf rs!JenisKegiatan = "Endoskopi (semua bentuk)" Then
                With oSheet
                    .Cells(j, 8) = Trim(rs!Jumlah + Cell1)
                End With
            ElseIf rs!JenisKegiatan = "Hemodialisa" Then
                With oSheet
                    .Cells(j, 8) = Trim(rs!Jumlah + Cell1)
                End With
            ElseIf rs!JenisKegiatan = "Densometri Tulang" Then
                With oSheet
                    .Cells(j, 8) = Trim(rs!Jumlah + Cell1)
                End With
            ElseIf rs!JenisKegiatan = "Pungsi" Then
                With oSheet
                    .Cells(j, 8) = Trim(rs!Jumlah + Cell1)
                End With
            ElseIf rs!JenisKegiatan = "Spirometri" Then
                With oSheet
                    .Cells(j, 8) = Trim(rs!Jumlah + Cell1)
                End With
            ElseIf rs!JenisKegiatan = "Tes Kulit/Alergi/Histamin" Then
                With oSheet
                    .Cells(j, 8) = Trim(rs!Jumlah + Cell1)
                End With
            ElseIf rs!JenisKegiatan = "Topometri" Then
                With oSheet
                    .Cells(j, 8) = Trim(rs!Jumlah + Cell1)
                End With
            ElseIf rs!JenisKegiatan = "Lain-lain" Then
                With oSheet
                    .Cells(j, 8) = Trim(rs!Jumlah + Cell1)
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
error:
    Call msubPesanError
    Screen.MousePointer = vbDefault
End Sub
