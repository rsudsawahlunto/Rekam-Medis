VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frm1sub3New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL1.3 Fasilitas Tempat Tidur Rawat Inap"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
   Icon            =   "frm1sub3New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5250
   Begin VB.Frame fraButton 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   5205
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1905
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   3120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   3120
      Picture         =   "frm1sub3New2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2115
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frm1sub3New2.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm1sub3New2.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frm1sub3New2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project/reference/microsoft excel 12.0 object library
'Selalu gunakan format file excel 2003  .xls sebagai standar agar pengguna excel 2003 atau diatasnya dpt menggunakan report laporannya
'Catatan: Format excel 2000 tidak dpt mengoperasikan beberapa fungsi yg ada pada excell 2003 atau diatasnya

Option Explicit

'Special Buat Excel
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel.Range
Dim oResizeRange As Excel.Range
Dim i As Integer
Dim j As String
'Special Buat Excel

Dim Cell12 As String
Dim Cell15 As String
Dim Cell18 As String
Dim Cell21 As String
Dim Cell24 As String
Dim xx As Integer

Private Sub cmdCetak_Click()
    On Error GoTo hell

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.Path & "\RL 1.3_Tempat Tidur.xlsx")
    Set oSheet = oWB.ActiveSheet

    Set rsb = Nothing
    strSQL = "select * from profilrs"
    Call msubRecFO(rsb, strSQL)

    For xx = 2 To 31
        With oSheet
            .Cells(xx, 2) = rsb("KdRS").value
            .Cells(xx, 4) = rsb("KotaKodyaKab").value
        End With
    Next xx

    Set rs = Nothing
    strSQL = "select KdSubInstalasi, Kelas, SUM(JmlBed) as JmlBed from RL1_03New Group by KdSubInstalasi, Kelas order by kdsubinstalasi"
    Call msubRecFO(rs, strSQL)

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        While Not rs.EOF
            If rs!kdsubinstalasi = "001" Then
                j = 2
            ElseIf rs!kdsubinstalasi = "003" Then
                j = 3
            ElseIf rs!kdsubinstalasi = "004" Then
                j = 4
            ElseIf rs!kdsubinstalasi = "005" Then
                j = 5
            ElseIf rs!kdsubinstalasi = "002" Then
                j = 6
            ElseIf rs!kdsubinstalasi = "015" Then
                j = 7
            ElseIf rs!kdsubinstalasi = "006" Then
                j = 8
            ElseIf rs!kdsubinstalasi = "022" Then
                j = 9
            ElseIf rs!kdsubinstalasi = "007" Then
                j = 10
            ElseIf rs!kdsubinstalasi = "008" Then
                j = 11
            ElseIf rs!kdsubinstalasi = "009" Then
                j = 14
            ElseIf rs!kdsubinstalasi = "010" Then
                j = 15
            ElseIf rs!kdsubinstalasi = "011" Then
                j = 16
            ElseIf rs!kdsubinstalasi = "013" Then
                j = 17
            ElseIf rs!kdsubinstalasi = "016" Then
                j = 18
            ElseIf rs!kdsubinstalasi = "" Then
                j = 19
            ElseIf rs!kdsubinstalasi = "014" Then
                j = 20
            ElseIf rs!kdsubinstalasi = "027" Then
                j = 21
            ElseIf rs!kdsubinstalasi = "017" Then
                j = 22
            ElseIf rs!kdsubinstalasi = "020" Then
                j = 23
            ElseIf rs!kdsubinstalasi = "021" Then
                j = 24
            ElseIf rs!kdsubinstalasi = "023" Then
                j = 25
            ElseIf rs!kdsubinstalasi = "024" Then
                j = 26
            ElseIf rs!kdsubinstalasi = "025" Then
                j = 27
            ElseIf rs!kdsubinstalasi = "018" Then
                j = 28
            ElseIf rs!kdsubinstalasi = "012" Then
                j = 29
            ElseIf rs!kdsubinstalasi = "019" Then
                j = 30
            ElseIf rs!kdsubinstalasi = "028" Then  'Perinatologi / Bayi  (Sengaja Ditutup Karena Lum Ada)
                j = 31
            End If

            If rs!Kelas = "Kelas (VVIP)" Then
                Call setcellVVIP
            ElseIf rs!Kelas = "MASTER (VIP)" Then
                Call setcellVIP
            ElseIf rs!Kelas = "SUITE (I)" Then
                Call setcellI
            ElseIf rs!Kelas = "DELUXE (II)" Then
                Call setcellII
            ElseIf rs!Kelas = "STANDARD (III)" Then
                Call setcellIII
            ElseIf rs!Kelas = "INTENSIF" Then
                Call setcellKelasKhusus
            ElseIf rs!Kelas = "Kelas (VVIP)" Then
                Call setcellVVIP
            ElseIf rs!Kelas = "MASTER (VIP)" Then
                Call setcellVIP
            ElseIf rs!Kelas = "SUITE (I)" Then
                Call setcellI
            ElseIf rs!Kelas = "DELUXE (II)" Then
                Call setcellII
            ElseIf rs!Kelas = "STANDARD (III)" Then
                Call setcellIII
            ElseIf rs!Kelas = "INTENSIF" Then
                Call setcellKelasKhusus
            End If

            rs.MoveNext
        Wend
        oXL.Visible = True
    End If

    Exit Sub
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
End Sub

Private Sub setcellVVIP()
    With oSheet
        .Cells(j, 7) = Trim(IIf(IsNull(rs!jmlbed), 0, (rs!jmlbed)))
    End With
End Sub

Private Sub setcellVIP()
    With oSheet
        .Cells(j, 8) = Trim(IIf(IsNull(rs!jmlbed), 0, (rs!jmlbed)))
    End With
End Sub

Private Sub setcellI()
    With oSheet
        .Cells(j, 9) = Trim(IIf(IsNull(rs!jmlbed), 0, (rs!jmlbed)))
    End With
End Sub

Private Sub setcellII()
    With oSheet
        .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlbed), 0, (rs!jmlbed)))
    End With
End Sub

Private Sub setcellIII()
    With oSheet
        .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlbed), 0, (rs!jmlbed)))
    End With
End Sub

Private Sub setcellKelasKhusus()
    With oSheet
        .Cells(j, 12) = Trim(IIf(IsNull(rs!jmlbed), 0, (rs!jmlbed)))
    End With
End Sub
