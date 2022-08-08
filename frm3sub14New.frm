VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm3sub14New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL3.14 Kegiatan Rujukan"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6135
   Icon            =   "frm3sub14New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3075
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
         TabIndex        =   5
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   133169155
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   375
         Left            =   3240
         TabIndex        =   6
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   133169155
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
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm3sub14New.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frm3sub14New"
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

Dim Cell7 As String
Dim Cell8 As String
Dim Cell9 As String
Dim Cell10 As String
Dim Cell11 As String
Dim Cell12 As String
Dim Cell13 As String
Dim Cell14 As String
Dim Cell15 As String

'Special Buat Excel
Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpAwal.value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.value = Now
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo error

    Screen.MousePointer = vbHourglass

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.Path & "\Formulir RL 3.14.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With oSheet
        .Cells(7, 4) = rsb("KdRS").value
        .Cells(8, 4) = rsb("NamaRS").value
        .Cells(9, 4) = Right(dtpAwal.value, 4)
    End With

    Set rs = Nothing
    strSQL = "Select distinct * from RL3_14New where Tglmasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "'or tglmasuk is null"
    Call msubRecFO(rs, strSQL)

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        While Not rs.EOF
            If rs!kdsubinstalasi = "001" Then
                j = 14
            ElseIf rs!kdsubinstalasi = "002" Then
                j = 15
            ElseIf rs!kdsubinstalasi = "003" Then
                j = 16
            ElseIf rs!kdsubinstalasi = "005" Then
                j = 17
            ElseIf rs!kdsubinstalasi = "004" Then
                j = 18
            ElseIf rs!kdsubinstalasi = "007" Then
                j = 19
            ElseIf rs!kdsubinstalasi = "008" Then
                j = 20
            ElseIf rs!kdsubinstalasi = "009" Then
                j = 21
            ElseIf rs!kdsubinstalasi = "010" Then
                j = 22
            ElseIf rs!kdsubinstalasi = "011" Then
                j = 23
            ElseIf rs!kdsubinstalasi = "012" Then
                j = 24
            ElseIf rs!kdsubinstalasi = "014" Then
                j = 25
            ElseIf rs!kdsubinstalasi = "016" Then
                j = 26
            ElseIf rs!Spesialisasi = "Spesialisasi Lain" Then
                j = 27
            End If

            Cell7 = oSheet.Cells(j, 5).value
            Cell8 = oSheet.Cells(j, 6).value
            Cell9 = oSheet.Cells(j, 7).value
            Cell10 = oSheet.Cells(j, 8).value
            Cell11 = oSheet.Cells(j, 9).value
            Cell12 = oSheet.Cells(j, 10).value
            Cell13 = oSheet.Cells(j, 11).value
            Cell14 = oSheet.Cells(j, 12).value
            Cell15 = oSheet.Cells(j, 13).value

            If rs!kdsubinstalasi = "001" Then
                With oSheet
                    .Cells(j, 5) = rs!DariPuskesmas + Cell7
                    .Cells(j, 6) = rs!DariFasilitasLain + Cell8
                    .Cells(j, 7) = rs!DariRSLain + Cell9
                    .Cells(j, 8) = rs!DikembalikanPuskesmas + Cell10
                    .Cells(j, 9) = rs!DikembalikanFasilitasLain + Cell11
                    .Cells(j, 10) = rs!DikembalikanRSLain + Cell12
                    .Cells(j, 11) = rs!PasienRujukan + Cell13
                    .Cells(j, 12) = rs!DatangSendiri + Cell14
                    .Cells(j, 13) = rs!DiterimaKembali + Cell15
                End With
            ElseIf rs!kdsubinstalasi = "002" Then
                With oSheet
                    .Cells(j, 5) = rs!DariPuskesmas + Cell7
                    .Cells(j, 6) = rs!DariFasilitasLain + Cell8
                    .Cells(j, 7) = rs!DariRSLain + Cell9
                    .Cells(j, 8) = rs!DikembalikanPuskesmas + Cell10
                    .Cells(j, 9) = rs!DikembalikanFasilitasLain + Cell11
                    .Cells(j, 10) = rs!DikembalikanRSLain + Cell12
                    .Cells(j, 11) = rs!PasienRujukan + Cell13
                    .Cells(j, 12) = rs!DatangSendiri + Cell14
                    .Cells(j, 13) = rs!DiterimaKembali + Cell15
                End With
            ElseIf rs!kdsubinstalasi = "003" Then
                With oSheet
                    .Cells(j, 5) = rs!DariPuskesmas + Cell7
                    .Cells(j, 6) = rs!DariFasilitasLain + Cell8
                    .Cells(j, 7) = rs!DariRSLain + Cell9
                    .Cells(j, 8) = rs!DikembalikanPuskesmas + Cell10
                    .Cells(j, 9) = rs!DikembalikanFasilitasLain + Cell11
                    .Cells(j, 10) = rs!DikembalikanRSLain + Cell12
                    .Cells(j, 11) = rs!PasienRujukan + Cell13
                    .Cells(j, 12) = rs!DatangSendiri + Cell14
                    .Cells(j, 13) = rs!DiterimaKembali + Cell15
                End With
            ElseIf rs!kdsubinstalasi = "004" Then
                With oSheet
                    .Cells(j, 5) = rs!DariPuskesmas + Cell7
                    .Cells(j, 6) = rs!DariFasilitasLain + Cell8
                    .Cells(j, 7) = rs!DariRSLain + Cell9
                    .Cells(j, 8) = rs!DikembalikanPuskesmas + Cell10
                    .Cells(j, 9) = rs!DikembalikanFasilitasLain + Cell11
                    .Cells(j, 10) = rs!DikembalikanRSLain + Cell12
                    .Cells(j, 11) = rs!PasienRujukan + Cell13
                    .Cells(j, 12) = rs!DatangSendiri + Cell14
                    .Cells(j, 13) = rs!DiterimaKembali + Cell15
                End With
            ElseIf rs!kdsubinstalasi = "005" Then
                With oSheet
                    .Cells(j, 5) = rs!DariPuskesmas + Cell7
                    .Cells(j, 6) = rs!DariFasilitasLain + Cell8
                    .Cells(j, 7) = rs!DariRSLain + Cell9
                    .Cells(j, 8) = rs!DikembalikanPuskesmas + Cell10
                    .Cells(j, 9) = rs!DikembalikanFasilitasLain + Cell11
                    .Cells(j, 10) = rs!DikembalikanRSLain + Cell12
                    .Cells(j, 11) = rs!PasienRujukan + Cell13
                    .Cells(j, 12) = rs!DatangSendiri + Cell14
                    .Cells(j, 13) = rs!DiterimaKembali + Cell15
                End With
            ElseIf rs!kdsubinstalasi = "007" Then
                With oSheet
                    .Cells(j, 5) = rs!DariPuskesmas + Cell7
                    .Cells(j, 6) = rs!DariFasilitasLain + Cell8
                    .Cells(j, 7) = rs!DariRSLain + Cell9
                    .Cells(j, 8) = rs!DikembalikanPuskesmas + Cell10
                    .Cells(j, 9) = rs!DikembalikanFasilitasLain + Cell11
                    .Cells(j, 10) = rs!DikembalikanRSLain + Cell12
                    .Cells(j, 11) = rs!PasienRujukan + Cell13
                    .Cells(j, 12) = rs!DatangSendiri + Cell14
                    .Cells(j, 13) = rs!DiterimaKembali + Cell15
                End With
            ElseIf rs!kdsubinstalasi = "008" Then
                With oSheet
                    .Cells(j, 5) = rs!DariPuskesmas + Cell7
                    .Cells(j, 6) = rs!DariFasilitasLain + Cell8
                    .Cells(j, 7) = rs!DariRSLain + Cell9
                    .Cells(j, 8) = rs!DikembalikanPuskesmas + Cell10
                    .Cells(j, 9) = rs!DikembalikanFasilitasLain + Cell11
                    .Cells(j, 10) = rs!DikembalikanRSLain + Cell12
                    .Cells(j, 11) = rs!PasienRujukan + Cell13
                    .Cells(j, 12) = rs!DatangSendiri + Cell14
                    .Cells(j, 13) = rs!DiterimaKembali + Cell15
                End With
            ElseIf rs!kdsubinstalasi = "009" Then
                With oSheet
                    .Cells(j, 5) = rs!DariPuskesmas + Cell7
                    .Cells(j, 6) = rs!DariFasilitasLain + Cell8
                    .Cells(j, 7) = rs!DariRSLain + Cell9
                    .Cells(j, 8) = rs!DikembalikanPuskesmas + Cell10
                    .Cells(j, 9) = rs!DikembalikanFasilitasLain + Cell11
                    .Cells(j, 10) = rs!DikembalikanRSLain + Cell12
                    .Cells(j, 11) = rs!PasienRujukan + Cell13
                    .Cells(j, 12) = rs!DatangSendiri + Cell14
                    .Cells(j, 13) = rs!DiterimaKembali + Cell15
                End With
            ElseIf rs!kdsubinstalasi = "010" Then
                With oSheet
                    .Cells(j, 5) = rs!DariPuskesmas + Cell7
                    .Cells(j, 6) = rs!DariFasilitasLain + Cell8
                    .Cells(j, 7) = rs!DariRSLain + Cell9
                    .Cells(j, 8) = rs!DikembalikanPuskesmas + Cell10
                    .Cells(j, 9) = rs!DikembalikanFasilitasLain + Cell11
                    .Cells(j, 10) = rs!DikembalikanRSLain + Cell12
                    .Cells(j, 11) = rs!PasienRujukan + Cell13
                    .Cells(j, 12) = rs!DatangSendiri + Cell14
                    .Cells(j, 13) = rs!DiterimaKembali + Cell15
                End With
            ElseIf rs!kdsubinstalasi = "011" Then
                With oSheet
                    .Cells(j, 5) = rs!DariPuskesmas + Cell7
                    .Cells(j, 6) = rs!DariFasilitasLain + Cell8
                    .Cells(j, 7) = rs!DariRSLain + Cell9
                    .Cells(j, 8) = rs!DikembalikanPuskesmas + Cell10
                    .Cells(j, 9) = rs!DikembalikanFasilitasLain + Cell11
                    .Cells(j, 10) = rs!DikembalikanRSLain + Cell12
                    .Cells(j, 11) = rs!PasienRujukan + Cell13
                    .Cells(j, 12) = rs!DatangSendiri + Cell14
                    .Cells(j, 13) = rs!DiterimaKembali + Cell15
                End With
            ElseIf rs!kdsubinstalasi = "012" Then
                With oSheet
                    .Cells(j, 5) = rs!DariPuskesmas + Cell7
                    .Cells(j, 6) = rs!DariFasilitasLain + Cell8
                    .Cells(j, 7) = rs!DariRSLain + Cell9
                    .Cells(j, 8) = rs!DikembalikanPuskesmas + Cell10
                    .Cells(j, 9) = rs!DikembalikanFasilitasLain + Cell11
                    .Cells(j, 10) = rs!DikembalikanRSLain + Cell12
                    .Cells(j, 11) = rs!PasienRujukan + Cell13
                    .Cells(j, 12) = rs!DatangSendiri + Cell14
                    .Cells(j, 13) = rs!DiterimaKembali + Cell15
                End With
            ElseIf rs!kdsubinstalasi = "014" Then
                With oSheet
                    .Cells(j, 5) = rs!DariPuskesmas + Cell7
                    .Cells(j, 6) = rs!DariFasilitasLain + Cell8
                    .Cells(j, 7) = rs!DariRSLain + Cell9
                    .Cells(j, 8) = rs!DikembalikanPuskesmas + Cell10
                    .Cells(j, 9) = rs!DikembalikanFasilitasLain + Cell11
                    .Cells(j, 10) = rs!DikembalikanRSLain + Cell12
                    .Cells(j, 11) = rs!PasienRujukan + Cell13
                    .Cells(j, 12) = rs!DatangSendiri + Cell14
                    .Cells(j, 13) = rs!DiterimaKembali + Cell15
                End With
            ElseIf rs!kdsubinstalasi = "016" Then
                With oSheet
                    .Cells(j, 5) = rs!DariPuskesmas + Cell7
                    .Cells(j, 6) = rs!DariFasilitasLain + Cell8
                    .Cells(j, 7) = rs!DariRSLain + Cell9
                    .Cells(j, 8) = rs!DikembalikanPuskesmas + Cell10
                    .Cells(j, 9) = rs!DikembalikanFasilitasLain + Cell11
                    .Cells(j, 10) = rs!DikembalikanRSLain + Cell12
                    .Cells(j, 11) = rs!PasienRujukan + Cell13
                    .Cells(j, 12) = rs!DatangSendiri + Cell14
                    .Cells(j, 13) = rs!DiterimaKembali + Cell15
                End With
            ElseIf rs!Spesialisasi = "Spesialisasi Lain" Then
                With oSheet
                    .Cells(j, 5) = rs!DariPuskesmas + Cell7
                    .Cells(j, 6) = rs!DariFasilitasLain + Cell8
                    .Cells(j, 7) = rs!DariRSLain + Cell9
                    .Cells(j, 8) = rs!DikembalikanPuskesmas + Cell10
                    .Cells(j, 9) = rs!DikembalikanFasilitasLain + Cell11
                    .Cells(j, 10) = rs!DikembalikanRSLain + Cell12
                    .Cells(j, 11) = rs!PasienRujukan + Cell13
                    .Cells(j, 12) = rs!DatangSendiri + Cell14
                    .Cells(j, 13) = rs!DiterimaKembali + Cell15
                End With
            End If
            rs.MoveNext
        Wend
    End If

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
error:
    MsgBox "Data Tidak Ada", vbInformation, "Validasi"
    Screen.MousePointer = vbDefault
End Sub

