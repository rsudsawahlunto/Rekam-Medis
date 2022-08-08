VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm3sub13New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.13 Pengadaan Obat, Penulisan Dan Pelayanan Resep"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6405
   Icon            =   "frm3sub13New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6405
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   2280
      Width           =   1905
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   6375
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   133300227
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   133300227
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin VB.Label Label1 
         Caption         =   "s/d"
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   600
         Width           =   375
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
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm3sub13New.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frm3sub13New"
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
Dim k As Integer
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
    dtpAwal.value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAwal.value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.value = Now
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo error

    Screen.MousePointer = vbHourglass

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    Set oWB = oXL.Workbooks.Open(App.Path & "\Formulir RL 3.13.xlsx")
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

    strSQL = "Select * from RL3_13New where TglTerima between '" & Format(dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "'"
    Call msubRecFO(rsx, strSQL)

    If rsx.RecordCount > 0 Then
        rsx.MoveFirst

        While Not rsx.EOF
            If rsx!KdKategoryBarang = "01" Then
                j = 16
            ElseIf rsx!KdKategoryBarang = "02" Then
                j = 17
            ElseIf rsx!KdKategoryBarang = "03" Then
                j = 18
            End If

            Cell1 = oSheet.Cells(j, 7).value
            Cell2 = oSheet.Cells(j, 9).value

            If rsx!KdKategoryBarang = "01" Then
                With oSheet
                    .Cells(j, 7) = Trim(rsx!jmlnonformularium + Cell1)
                    .Cells(j, 9) = Trim(rsx!jmlformularium + Cell2)
                End With
            ElseIf rsx!KdKategoryBarang = "02" Then
                With oSheet
                    .Cells(j, 7) = Trim(rsx!jmlnonformularium + Cell1)
                    .Cells(j, 9) = Trim(rsx!jmlformularium + Cell2)
                End With
            ElseIf rsx!KdKategoryBarang = "03" Then
                With oSheet
                    .Cells(j, 7) = Trim(rsx!jmlnonformularium + Cell1)
                    .Cells(j, 9) = Trim(rsx!jmlformularium + Cell2)
                End With
            End If

            rsx.MoveNext
        Wend

    End If

    Set rs1 = Nothing

    strSQL1 = "Select * from RL3_13_2New where TglStruk between '" & Format(dtpAwal.value, "yyyy/MM/dd") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd") & "'"
    Call msubRecFO(rs1, strSQL1)

    If rs1.RecordCount > 0 Then
        rs1.MoveFirst

        While Not rs1.EOF

            If rs1!KdKategoryBarang = "01" Then
                k = 24
            ElseIf rs1!KdKategoryBarang = "02" Then
                k = 25
            ElseIf rs1!KdKategoryBarang = "03" Then
                k = 26
            End If

            Cell3 = oSheet.Cells(k, 5).value
            Cell4 = oSheet.Cells(k, 7).value
            Cell5 = oSheet.Cells(k, 9).value

            If rs1!KdKategoryBarang = "01" Then
                With oSheet
                    If rs1!NamaInstalasi = "Instalasi Rawat Jalan" Then
                        .Cells(k, 5) = Trim(rs1!JmlBarang + Cell3)
                    ElseIf rs1!NamaInstalasi = "Instalasi Rawat Inap" Then
                        .Cells(k, 7) = Trim(rs1!JmlBarang + Cell4)
                    ElseIf rs1!NamaInstalasi = "Instalasi Gawat Darurat" Then
                        .Cells(k, 9) = Trim(rs1!JmlBarang + Cell5)
                    End If
                End With
            ElseIf rsx!KdKategoryBarang = "02" Then
                With oSheet
                    If rs1!NamaInstalasi = "Instalasi Rawat Jalan" Then
                        .Cells(k, 5) = Trim(rs1!JmlBarang + Cell3)
                    ElseIf rs1!NamaInstalasi = "Instalasi Rawat Inap" Then
                        .Cells(k, 7) = Trim(rs1!JmlBarang + Cell4)
                    ElseIf rs1!NamaInstalasi = "Instalasi Gawat Darurat" Then
                        .Cells(k, 9) = Trim(rs1!JmlBarang + Cell5)
                    End If
                End With

            ElseIf rsx!KdKategoryBarang = "03" Then
                With oSheet
                    If rs1!NamaInstalasi = "Instalasi Rawat Jalan" Then
                        .Cells(k, 5) = Trim(rs1!JmlBarang + Cell3)
                    ElseIf rs1!NamaInstalasi = "Instalasi Rawat Inap" Then
                        .Cells(k, 7) = Trim(rs1!JmlBarang + Cell4)
                    ElseIf rs1!NamaInstalasi = "Instalasi Gawat Darurat" Then
                        .Cells(k, 9) = Trim(rs1!JmlBarang + Cell5)
                    End If
                End With
            End If

            rs1.MoveNext
        Wend
    End If

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    MsgBox "Data Tidak Ada", vbInformation, "Validasi"
    Screen.MousePointer = vbDefault
End Sub

