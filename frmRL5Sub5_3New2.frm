VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRL5Sub5_3New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL5.3 Daftar 10 Besar Penyakit Rawat Inap"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6300
   Icon            =   "frmRL5Sub5_3New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6300
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   6255
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   1905
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   6255
      Begin VB.Frame Frame2 
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   4
         Top             =   120
         Width           =   5055
         Begin MSComCtl2.DTPicker dtptahun 
            Height          =   375
            Left            =   2760
            TabIndex        =   5
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
            CustomFormat    =   "dd MMM yyyy"
            Format          =   136708099
            UpDown          =   -1  'True
            CurrentDate     =   40544
         End
         Begin MSComCtl2.DTPicker dtptahunawal 
            Height          =   375
            Left            =   240
            TabIndex        =   6
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
            CustomFormat    =   "dd MMM yyyy"
            Format          =   136708099
            UpDown          =   -1  'True
            CurrentDate     =   40544
         End
         Begin VB.Label Label2 
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   7
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.TextBox txtJmlData 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         TabIndex        =   2
         Text            =   "10"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
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
      Picture         =   "frmRL5Sub5_3New2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmRL5Sub5_3New2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel.Range
Dim oResizeRange As Excel.Range
Dim xx As Integer

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

Private Sub cmdCetak_Click()
    On Error GoTo error

    Dim i, j, k As Integer
    
    If txtJmlData = "" Or txtJmlData = "0" Then
        MsgBox "Jumlah data tidak boleh kosong", vbOKOnly, "Peringatan"
        Exit Sub
    Exit Sub
    End If
    
    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.path & "\RL 5.3 10_Besar Penyakit Rawat Inap.xlsx")
    Set oSheet = oWB.ActiveSheet

    strSQL = "SELECT top " & txtJmlData & " Diagnosa, kdDiagnosa, sum(jumlahpasien) as [JmlPasien],SUM(JmlPasienOutMatiPria) as [Mati Pria]," & _
    " SUM(JmlPasienOutMatiWanita) as [Mati Wanita], SUM(JmlPasienOutHidupPria) as [KeluarHidupPria]," & _
    " SUM(JmlPasienOutHidupWanita) as [KeluarHidupWanita] " & _
    " FROM V_RekapitulasiDiagnosaTopTen2New " & _
    " WHERE TglPulang Between '" & Format(dtpTahunAwal.value, "yyyy-mm-dd 00:00:00") & "' AND '" & Format(dtptahun.value, "yyyy-mm-dd 23:59:59") & "' " & _
    " group by Diagnosa, KdDiagnosa order by [JmlPasien] desc"
    Call msubRecFO(rs, strSQL)

    If rs.RecordCount = 0 Then
        MsgBox "Data Tidak Ada", vbInformation, "Validasi"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        For j = 1 To rs.RecordCount
            k = j + 1
            With oSheet
                .Cells(k, 8) = rs("kdDiagnosa").value
                .Cells(k, 9) = rs("Diagnosa").value
                .Cells(k, 12) = rs("Mati Pria").value
                .Cells(k, 13) = rs("Mati Wanita").value
                .Cells(k, 10) = rs("KeluarHidupPria").value
                .Cells(k, 11) = rs("KeluarHidupWanita").value
            End With
            rs.MoveNext
        Next j
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
error:
    Call msubPesanError
    Screen.MousePointer = vbDefault
End Sub


Private Sub txtJmlData_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = 13 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
