VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm5sub04New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Rekapitulasi 10 Besar Penyakit Rawat Jalan"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm5sub04New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6690
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
      TabIndex        =   3
      Top             =   930
      Width           =   6615
      Begin VB.Frame Frame1 
         Caption         =   "Periode"
         Height          =   735
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   5655
         Begin MSComCtl2.DTPicker dtptahun 
            Height          =   375
            Left            =   3120
            TabIndex        =   8
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
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
            CustomFormat    =   "dd MMM yyyy"
            Format          =   140312579
            UpDown          =   -1  'True
            CurrentDate     =   40544
         End
         Begin MSComCtl2.DTPicker dtptahunawal 
            Height          =   375
            Left            =   360
            TabIndex        =   9
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
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
            CustomFormat    =   "dd MMM yyyy"
            Format          =   140312579
            UpDown          =   -1  'True
            CurrentDate     =   40544
         End
         Begin VB.Label Label2 
            Caption         =   "s/d"
            Height          =   255
            Left            =   2760
            TabIndex        =   10
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.TextBox txtJmlData 
         Height          =   315
         Left            =   3240
         TabIndex        =   6
         Text            =   "10"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah Data"
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
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
      TabIndex        =   2
      Top             =   2520
      Width           =   6615
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   3480
         TabIndex        =   1
         Top             =   240
         Width           =   1695
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   4800
      Picture         =   "frm5sub04New2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm5sub04New2.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frm5sub04New2.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frm5sub04New2"
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
Dim j As String
Dim xx As Integer

Private Sub cmdCetak_Click()

    If txtJmlData.Text = "" Or txtJmlData.Text = "0" Then
        MsgBox "Jumlah data tidak boleh kosong", vbOKOnly, "Peringatan"
        Exit Sub
    End If
    'Buka Excel
    Set oXL = CreateObject("Excel.Application")

    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.path & "\RL 5.4 10_Besar Penyakit Rawat Jalan.xlsx")
    Set oSheet = oWB.ActiveSheet

    Set rs = Nothing
    
  
    
'    strSQL = "SELECT top " & txtJmlData.Text & " kdDiagnosa, Diagnosa, sum(JmlPasienOutPria) JmlPasienOutPria, sum(JmlPasienOutWanita) as JmlPasienOutWanita,  sum(JmlPasienOutPria+JmlPasienOutWanita) as Total,sum(jumlahpasien) as [JmlPasien]" & _
'    " FROM V_RekapitulasiDiagnosaTopTen " & _
'    " WHERE TglPeriksa BETWEEN '" & Format(dtpTahunAwal.value, "YYYY-mm-dd 00:00:00") & "' AND '" & Format(dtptahun.value, "YYYY-mm-dd 23:59:59") & "'" & _
'    " AND kdinstalasi = 02  " & _
'    " group by KdDiagnosa,Diagnosa order by [JmlPasien] desc"

    strSQL = "SELECT top " & txtJmlData.Text & " kdDiagnosa, Diagnosa, sum(case when StatusKasus='Baru' then JmlPasienOutPria else 0 end) JmlPasienOutPria, sum(case when StatusKasus='Baru' then JmlPasienOutWanita else 0 end) as JmlPasienOutWanita, sum(case when StatusKasus='Baru' then JmlPasienOutPria else 0 end) + sum(case when StatusKasus='Baru' then JmlPasienOutWanita else 0 end) as JmlPasienOutAll, sum(JmlPasienOutPria+JmlPasienOutWanita) as Total,sum(jumlahpasien) as [JmlPasien]" & _
    " FROM V_RekapitulasiDiagnosaTopTen " & _
    " WHERE TglPulang BETWEEN '" & Format(dtpTahunAwal.value, "YYYY-mm-dd 00:00:00") & "' AND '" & Format(dtptahun.value, "YYYY-mm-dd 23:59:59") & "'" & _
    " group by KdDiagnosa,Diagnosa order by [JmlPasienOutAll] desc"
    
    Call msubRecFO(rs, strSQL)

    If rs.RecordCount = 0 Then
        MsgBox "Data Tidak Ada", vbInformation, "Validasi"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        j = 2
        Call setcell
    End If

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    For xx = 2 To rs.RecordCount + 1
        With oSheet
            .Cells(xx, 1) = rsb("KodeExternal").value
            .Cells(xx, 3) = rsb("KdRS").value
            .Cells(xx, 2) = rsb("KotaKodyaKab").value
            .Cells(xx, 4) = rsb("NamaRS").value
            If dtpTahunAwal.Month = dtptahun.Month Then
            .Cells(xx, 5) = Format(dtptahun.value, "MMMM")
            Else
            .Cells(xx, 5) = Format(dtpTahunAwal.value, "MMMM") & " s/d " & Format(dtptahun.value, "MMMM")
            End If
            If dtpTahunAwal.Year = dtptahun.Year Then
            .Cells(xx, 6) = Format(dtptahun.value, "YYYY")
            Else
            .Cells(xx, 6) = Format(dtpTahunAwal.value, "YYYY") & " s/d " & Format(dtptahun.value, "YYYY")
            End If
        End With
    Next xx

    oXL.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTahunAwal.value = Now
    dtpTahunAwal.Day = "01"
    dtptahun.value = Now
    dtpTahunAwal.CustomFormat = "dd MMM yyyyy"
    dtptahun.CustomFormat = "dd MMM yyyyy"
End Sub

Private Sub setcell()
    While Not rs.EOF
        With oSheet
            .Cells(j, 8) = Trim(IIf(IsNull(rs!kdDiagnosa.value), 0, (rs!kdDiagnosa.value)))
            .Cells(j, 9) = Trim(IIf(IsNull(rs!Diagnosa.value), 0, (rs!Diagnosa.value)))
            .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpasienoutpria.value), 0, (rs!jmlpasienoutpria.value)))
            .Cells(j, 11) = Trim(IIf(IsNull(rs!JmlPasienOutWanita.value), 0, (rs!JmlPasienOutWanita.value)))
            .Cells(j, 13) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
        End With
        j = j + 1
        rs.MoveNext
    Wend
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = 13 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

