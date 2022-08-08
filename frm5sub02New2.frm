VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm5sub02New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kunjungan Rawat Jalan"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm5sub02New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6795
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
      Width           =   6735
      Begin MSComCtl2.DTPicker dtptahun 
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   840
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
         CustomFormat    =   "MMM yyyy"
         Format          =   57212931
         UpDown          =   -1  'True
         CurrentDate     =   40544
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
      Width           =   6735
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   3360
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
      Left            =   4920
      Picture         =   "frm5sub02New2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm5sub02New2.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frm5sub02New2.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frm5sub02New2"
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

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.path & "\RL 5.2_Kunjungan Rawat Jalan.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    For xx = 2 To 31
        With oSheet
            .Cells(xx, 1) = rsb("KdRS").value
            .Cells(xx, 2) = rsb("NamaRS").value
            .Cells(xx, 3) = Format(dtptahun.value, "MMMM")
            .Cells(xx, 4) = Format(dtptahun.value, "YYYY")
            .Cells(xx, 5) = rsb("KotaKodyaKab").value
            .Cells(xx, 6) = rsb("KodeExternal").value
        End With
    Next xx

    Set rs = Nothing

    strSQL = "SELECT NamaExternal, sum(JmlLama) + sum(JmlBaru) + sum(JmlRujukan) as JmlPasien" & _
    " FROM RL5_2New " & _
    " WHERE Month(TglMasuk) = '" & dtptahun.Month & "' and Year(TglMasuk) = '" & dtptahun.Year & "' " & _
    " group by NamaExternal"

    Call msubRecFO(rs, strSQL)

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        j = 2
        Call setcell
    End If

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

    dtptahun.value = Now
    dtptahun.CustomFormat = "MMM yyyyy"
End Sub

Private Sub setcell()
    While Not rs.EOF

        With oSheet
            If rs!NamaExternal = "Dalam" Then
                .Cells(2, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Bedah" Then
                .Cells(3, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Kesehatan Anak" Then
                .Cells(4, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Keluarga Berencana" Then
                .Cells(8, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Bedah Syaraf" Then
                .Cells(9, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Saraf" Then
                .Cells(10, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Jiwa" Then
                .Cells(11, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Napza" Then
                .Cells(12, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "THT" Then
                .Cells(14, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Mata" Then
                .Cells(15, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Kulit & Kelamin" Then
                .Cells(16, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Gigi & Mulut" Then
                .Cells(17, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Kardiologi" Then
                .Cells(19, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Bedah Ortophedi" Then
                .Cells(21, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Paru-Paru" Then
                .Cells(22, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Kusta" Then
                .Cells(23, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Umum" Then
                .Cells(24, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Rehabilitasi Medik" Then
                .Cells(26, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Akupungtur Medik" Then
                .Cells(27, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Gizi" Then
                .Cells(28, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            ElseIf rs!NamaExternal = "Day Care" Then
                .Cells(29, 9) = Trim(IIf(IsNull(rs!JMlPasien.value), 0, (rs!JMlPasien.value)))
            End If
        End With

        rs.MoveNext
    Wend
End Sub

