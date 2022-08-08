VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm5sub04New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Rekapitulasi 10 Besar Penyakit Rawat Jalan"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm5sub04New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6510
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   0
      TabIndex        =   5
      Top             =   930
      Width           =   6495
      Begin VB.Frame Frame1 
         Caption         =   "Periode"
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
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   5595
         Begin MSComCtl2.DTPicker DTPickerAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   0
            Top             =   240
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
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   116981763
            UpDown          =   -1  'True
            CurrentDate     =   37956
         End
         Begin MSComCtl2.DTPicker DTPickerAkhir 
            Height          =   375
            Left            =   3360
            TabIndex        =   1
            Top             =   240
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
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   116981763
            UpDown          =   -1  'True
            CurrentDate     =   37956
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3000
            TabIndex        =   7
            Top             =   315
            Width           =   255
         End
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   2520
      Width           =   6495
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   3240
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   4560
      Picture         =   "frm5sub04New.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm5sub04New.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frm5sub04New.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frm5sub04New"
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

Private Sub cmdCetak_Click()

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")

    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.Path & "\Formulir RL 5.4.xlsx")
    Set oSheet = oWB.ActiveSheet

    Set rsb = Nothing
    strSQL = "select * from profilrs"
    Call msubRecFO(rsb, strSQL)

    With oSheet
        .Cells(6, 4) = Trim(IIf(IsNull(rsb!KdRs), 0, (rsb!KdRs)))
        .Cells(7, 4) = Trim(IIf(IsNull(rsb!NamaRS), 0, (rsb!NamaRS)))

        .Cells(8, 4) = Format(Now, "MM")
        .Cells(9, 4) = Format(Now, "yyyy")
    End With

    Set rs = Nothing

    strSQL = "SELECT top 10 kdDiagnosa, Diagnosa, sum(JmlPasienOutPria) JmlPasienOutPria, sum(JmlPasienOutWanita) as JmlPasienOutWanita,  sum(JmlPasienOutPria+JmlPasienOutWanita) as Total,sum(jumlahpasien) as [JmlPasien]" & _
    " FROM V_RekapitulasiDiagnosaTopTen " & _
    " WHERE TglPeriksa BETWEEN " & _
    " '" & Format(DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND " & _
    " '" & Format(DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' AND kdinstalasi = 02  " & _
    " group by KdDiagnosa,Diagnosa order by Diagnosa asc"

    Call msubRecFO(rs, strSQL)

    If rs.RecordCount = 0 Then
        MsgBox "Data Tidak Ada", vbInformation, "Validasi"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        j = 14
        Call setcell
    End If

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub DTPickerAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DTPickerAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    With Me
        .DTPickerAwal.value = Now
        .DTPickerAkhir.value = Now
    End With
End Sub

Private Sub setcell()
    While Not rs.EOF
        With oSheet
            .Cells(j, 2) = Trim(IIf(IsNull(rs!kdDiagnosa.value), 0, (rs!kdDiagnosa.value)))
            .Cells(j, 5) = Trim(IIf(IsNull(rs!Diagnosa.value), 0, (rs!Diagnosa.value)))
            .Cells(j, 6) = Trim(IIf(IsNull(rs!jmlpasienoutpria.value), 0, (rs!jmlpasienoutpria.value)))
            .Cells(j, 7) = Trim(IIf(IsNull(rs!JmlPasienOutWanita.value), 0, (rs!JmlPasienOutWanita.value)))
            .Cells(j, 8) = Trim(IIf(IsNull(rs!total.value), 0, (rs!total.value)))
            .Cells(j, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
        End With
        j = j + 1
        rs.MoveNext
    Wend
End Sub

