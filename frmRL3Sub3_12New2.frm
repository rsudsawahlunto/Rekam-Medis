VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRL3Sub3_12New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.12 Kegiatan Keluarga Berencana"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6135
   Icon            =   "frmRL3Sub3_12New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
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
         Format          =   126812163
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
      TabIndex        =   5
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL3Sub3_12New2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmRL3Sub3_12New2"
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
Dim w, x, y, z As String
Dim Cell3 As String
Dim Cell4 As String
Dim Cell5 As String
Dim Cell7 As String
Dim Cell8 As String
Dim Cell9 As String
'Special Buat Excel

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
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.Path & "\RL 3.12_keluarga berencana.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    For xx = 2 To 9
        With oSheet
            .Cells(xx, 3) = rsb("KdRS").value
            .Cells(xx, 2) = rsb("KotaKodyaKab").value
            .Cells(xx, 4) = rsb("NamaRS").value
            .Cells(xx, 5) = Format(dtptahun.value, "YYYY")
        End With
    Next xx

    Set rsx = Nothing

    strSQL = "Select * from RL3_12New where Year(TglPeriksa) = '" & dtptahun.Year & "'or TglPeriksa is null"
    Call msubRecFO(rsx, strSQL)

    ProgressBar1.Min = 0
    ProgressBar1.Max = rsx.RecordCount
    ProgressBar1.value = 0

    If rsx.RecordCount = 0 Then
        MsgBox "Data Tidak Ada", vbInformation, "Validasi"
        Exit Sub
    End If

    rsx.MoveFirst

    For i = 1 To rsx.RecordCount

        If rsx!kdjeniskontrasepsi = "01" Then
            j = 2
        ElseIf rsx!kdjeniskontrasepsi = "02" Then
            j = 3
        ElseIf rsx!kdjeniskontrasepsi = "03" Then
            j = 4
        ElseIf rsx!kdjeniskontrasepsi = "04" Then
            j = 5
        ElseIf rsx!kdjeniskontrasepsi = "05" Then
            j = 6
        ElseIf rsx!kdjeniskontrasepsi = "06" Then
            j = 7
        ElseIf rsx!kdjeniskontrasepsi = "07" Then
            j = 8
        ElseIf rsx!kdjeniskontrasepsi = "08" Then
            j = 9
        End If

        Cell3 = oSheet.Cells(j, 10).value
        Cell4 = oSheet.Cells(j, 11).value
        Cell5 = oSheet.Cells(j, 12).value
        Cell7 = oSheet.Cells(j, 17).value
        Cell8 = oSheet.Cells(j, 18).value
        Cell9 = oSheet.Cells(j, 19).value

        If rsx!kdjeniskontrasepsi = "01" Then
            With oSheet
                .Cells(j, 10) = Trim(rsx!bukanrujukan + Cell3)
                .Cells(j, 11) = Trim(rsx!rujukanri + Cell4)
                .Cells(j, 12) = Trim(rsx!rujukanrj + Cell5)
                .Cells(j, 17) = Trim(rsx!kunjunganulang + Cell7)
                .Cells(j, 18) = Trim(rsx!jmlefek + Cell8)
                .Cells(j, 19) = Trim(rsx!dirujukkeatas + Cell9)
            End With
        ElseIf rsx!kdjeniskontrasepsi = "02" Then
            With oSheet
                .Cells(j, 10) = Trim(rsx!bukanrujukan + Cell3)
                .Cells(j, 11) = Trim(rsx!rujukanri + Cell4)
                .Cells(j, 12) = Trim(rsx!rujukanrj + Cell5)
                .Cells(j, 17) = Trim(rsx!kunjunganulang + Cell7)
                .Cells(j, 18) = Trim(rsx!jmlefek + Cell8)
                .Cells(j, 19) = Trim(rsx!dirujukkeatas + Cell9)
            End With
        ElseIf rsx!kdjeniskontrasepsi = "03" Then
            With oSheet
                .Cells(j, 10) = Trim(rsx!bukanrujukan + Cell3)
                .Cells(j, 11) = Trim(rsx!rujukanri + Cell4)
                .Cells(j, 12) = Trim(rsx!rujukanrj + Cell5)
                .Cells(j, 17) = Trim(rsx!kunjunganulang + Cell7)
                .Cells(j, 18) = Trim(rsx!jmlefek + Cell8)
                .Cells(j, 19) = Trim(rsx!dirujukkeatas + Cell9)
            End With
        ElseIf rsx!kdjeniskontrasepsi = "04" Then
            With oSheet
                .Cells(j, 10) = Trim(rsx!bukanrujukan + Cell3)
                .Cells(j, 11) = Trim(rsx!rujukanri + Cell4)
                .Cells(j, 12) = Trim(rsx!rujukanrj + Cell5)
                .Cells(j, 17) = Trim(rsx!kunjunganulang + Cell7)
                .Cells(j, 18) = Trim(rsx!jmlefek + Cell8)
                .Cells(j, 19) = Trim(rsx!dirujukkeatas + Cell9)
            End With
        ElseIf rsx!kdjeniskontrasepsi = "05" Then
            With oSheet
                .Cells(j, 10) = Trim(rsx!bukanrujukan + Cell3)
                .Cells(j, 11) = Trim(rsx!rujukanri + Cell4)
                .Cells(j, 12) = Trim(rsx!rujukanrj + Cell5)
                .Cells(j, 17) = Trim(rsx!kunjunganulang + Cell7)
                .Cells(j, 18) = Trim(rsx!jmlefek + Cell8)
                .Cells(j, 19) = Trim(rsx!dirujukkeatas + Cell9)
            End With
        ElseIf rsx!kdjeniskontrasepsi = "06" Then
            With oSheet
                .Cells(j, 10) = Trim(rsx!bukanrujukan + Cell3)
                .Cells(j, 11) = Trim(rsx!rujukanri + Cell4)
                .Cells(j, 12) = Trim(rsx!rujukanrj + Cell5)
                .Cells(j, 17) = Trim(rsx!kunjunganulang + Cell7)
                .Cells(j, 18) = Trim(rsx!jmlefek + Cell8)
                .Cells(j, 19) = Trim(rsx!dirujukkeatas + Cell9)
            End With
        ElseIf rsx!kdjeniskontrasepsi = "07" Then
            With oSheet
                .Cells(j, 10) = Trim(rsx!bukanrujukan + Cell3)
                .Cells(j, 11) = Trim(rsx!rujukanri + Cell4)
                .Cells(j, 12) = Trim(rsx!rujukanrj + Cell5)
                .Cells(j, 17) = Trim(rsx!kunjunganulang + Cell7)
                .Cells(j, 18) = Trim(rsx!jmlefek + Cell8)
                .Cells(j, 19) = Trim(rsx!dirujukkeatas + Cell9)
            End With
        ElseIf rsx!kdjeniskontrasepsi = "08" Then
            With oSheet
                .Cells(j, 10) = Trim(rsx!bukanrujukan + Cell3)
                .Cells(j, 11) = Trim(rsx!rujukanri + Cell4)
                .Cells(j, 12) = Trim(rsx!rujukanrj + Cell5)
                .Cells(j, 17) = Trim(rsx!kunjunganulang + Cell7)
                .Cells(j, 18) = Trim(rsx!jmlefek + Cell8)
                .Cells(j, 19) = Trim(rsx!dirujukkeatas + Cell9)
            End With
        End If
        rsx.MoveNext

        ProgressBar1.value = Int(ProgressBar1.value) + 1
        lblPersen.Caption = Int(ProgressBar1.value * 100 / ProgressBar1.Max) & " %"
    Next i

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    Call msubPesanError
    Screen.MousePointer = vbDefault
End Sub
