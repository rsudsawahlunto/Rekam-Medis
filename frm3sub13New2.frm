VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm3sub13New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.13 Pengadaan Obat, Penulisan Dan Pelayanan Resep"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6405
   Icon            =   "frm3sub13New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6405
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2160
      Width           =   1905
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   6375
      Begin MSComCtl2.DTPicker dtptahun 
         Height          =   375
         Left            =   2160
         TabIndex        =   4
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
         Format          =   116523011
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   1
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
      Top             =   2760
      Width           =   5295
      _ExtentX        =   9340
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
      Left            =   5520
      TabIndex        =   6
      Top             =   2820
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm3sub13New2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frm3sub13New2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Special Buat Excel
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oXL2 As Excel.Application
Dim oWB2 As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oSheet2 As Excel.Worksheet
Dim oRng As Excel.Range
Dim oResizeRange As Excel.Range
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim xx As Integer
Dim Cell1 As String
Dim Cell2 As String
Dim Cell3 As String
Dim Cell4 As String
Dim Cell5 As String

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    j = 0
    k = 0
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)

    dtptahun.value = Now
    dtptahun.CustomFormat = "yyyyy"
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo error

    Screen.MousePointer = vbHourglass

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    Set oWB = oXL.Workbooks.Open(App.Path & "\RL 3.13_Obat Pengadaan.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    For xx = 2 To 4
        With oSheet
            .Cells(xx, 3) = rsb("KdRS").value
            .Cells(xx, 2) = rsb("KotaKodyaKab").value
            .Cells(xx, 4) = rsb("NamaRS").value
            .Cells(xx, 5) = Format(dtptahun.value, "YYYY")
        End With
    Next xx

    Set rsx = Nothing

    strSQL = "Select * from RL3_13New where Year(TglTerima) = '" & dtptahun.Year & "'"
    Call msubRecFO(rsx, strSQL)

    If rsx.RecordCount > 0 Then
        rsx.MoveFirst

        While Not rsx.EOF
            If rsx!KdKategoryBarang = "01" Then
                j = 2
            ElseIf rsx!KdKategoryBarang = "02" Then
                j = 3
            ElseIf rsx!KdKategoryBarang = "03" Then
                j = 4
            End If

            Cell1 = oSheet.Cells(j, 9).value
            Cell2 = oSheet.Cells(j, 10).value

            If rsx!KdKategoryBarang = "01" Then
                With oSheet
                    .Cells(j, 9) = Trim(rsx!jmlnonformularium + Cell1)
                    .Cells(j, 10) = Trim(rsx!jmlformularium + Cell2)
                End With
            ElseIf rsx!KdKategoryBarang = "02" Then
                With oSheet
                    .Cells(j, 9) = Trim(rsx!jmlnonformularium + Cell1)
                    .Cells(j, 10) = Trim(rsx!jmlformularium + Cell2)
                End With
            ElseIf rsx!KdKategoryBarang = "03" Then
                With oSheet
                    .Cells(j, 9) = Trim(rsx!jmlnonformularium + Cell1)
                    .Cells(j, 10) = Trim(rsx!jmlformularium + Cell2)
                End With
            End If

            rsx.MoveNext
        Wend
    End If

    'Buka Excel
    Set oXL2 = CreateObject("Excel.Application")
    Set oWB2 = oXL2.Workbooks.Open(App.Path & "\RL 3.13_Obat Pelayanan Resep.xlsx")
    Set oSheet2 = oWB2.ActiveSheet

    Set rs1 = Nothing

    strSQL1 = "Select * from RL3_13_2New where Year(TglStruk) = '" & dtptahun.Year & "'"
    Call msubRecFO(rs1, strSQL1)

    ProgressBar1.Min = 0
    ProgressBar1.Max = rs1.RecordCount
    ProgressBar1.value = 0

    If rs1.RecordCount > 0 Then
        rs1.MoveFirst

        While Not rs1.EOF
            If rs1!KdKategoryBarang = "01" Then
                k = 2
            ElseIf rs1!KdKategoryBarang = "02" Then
                k = 3
            ElseIf rs1!KdKategoryBarang = "03" Then
                k = 4
            End If

            Cell3 = oSheet2.Cells(k, 8).value
            Cell4 = oSheet2.Cells(k, 9).value
            Cell5 = oSheet2.Cells(k, 10).value

            If rs1!KdKategoryBarang = "01" Then
                With oSheet2
                    If rs1!NamaInstalasi = "Instalasi Rawat Jalan" Then
                        .Cells(k, 8) = Trim(rs1!JmlBarang + Cell3)
                    ElseIf rs1!NamaInstalasi = "Instalasi Rawat Inap" Then
                        .Cells(k, 9) = Trim(rs1!JmlBarang + Cell4)
                    ElseIf rs1!NamaInstalasi = "Instalasi Gawat Darurat" Then
                        .Cells(k, 10) = Trim(rs1!JmlBarang + Cell5)
                    End If
                End With
            ElseIf rs1!KdKategoryBarang = "02" Then
                With oSheet2
                    If rs1!NamaInstalasi = "Instalasi Rawat Jalan" Then
                        .Cells(k, 8) = Trim(rs1!JmlBarang + Cell3)
                    ElseIf rs1!NamaInstalasi = "Instalasi Rawat Inap" Then
                        .Cells(k, 9) = Trim(rs1!JmlBarang + Cell4)
                    ElseIf rs1!NamaInstalasi = "Instalasi Gawat Darurat" Then
                        .Cells(k, 10) = Trim(rs1!JmlBarang + Cell5)
                    End If
                End With

            ElseIf rs1!KdKategoryBarang = "03" Then
                With oSheet2
                    If rs1!NamaInstalasi = "Instalasi Rawat Jalan" Then
                        .Cells(k, 8) = Trim(rs1!JmlBarang + Cell3)
                    ElseIf rs1!NamaInstalasi = "Instalasi Rawat Inap" Then
                        .Cells(k, 9) = Trim(rs1!JmlBarang + Cell4)
                    ElseIf rs1!NamaInstalasi = "Instalasi Gawat Darurat" Then
                        .Cells(k, 10) = Trim(rs1!JmlBarang + Cell5)
                    End If
                End With
            End If

            rs1.MoveNext

            ProgressBar1.value = Int(ProgressBar1.value) + 1
            lblPersen.Caption = Int(ProgressBar1.value * 100 / ProgressBar1.Max) & " %"
        Wend
    End If

    oXL.Visible = True
    oXL2.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    MsgBox "Data Tidak Ada", vbInformation, "Validasi"
    Screen.MousePointer = vbDefault
End Sub
