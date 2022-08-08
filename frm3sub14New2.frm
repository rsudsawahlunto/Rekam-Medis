VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm3sub14New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL 3.14 Kegiatan Rujukan"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6135
   Icon            =   "frm3sub14New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6135
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   6135
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
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   1320
         Width           =   1905
      End
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
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   1320
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtptahun 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
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
         Format          =   104595459
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
      TabIndex        =   5
      Top             =   3000
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   4080
      Picture         =   "frm3sub14New2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2115
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
      TabIndex        =   6
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frm3sub14New2.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frm3sub14New2"
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
Dim xx As Integer

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
    
    dtptahun.value = Now
'    dtptahun.CustomFormat = "yyyyy"
    
End Sub

Private Sub cmdCetak_Click()
On Error GoTo error
Dim k As Integer
Dim i As Integer
    ProgressBar1.value = ProgressBar1.Min
    lblPersen.Caption = "0 %"
    Screen.MousePointer = vbHourglass
    
    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
'    oXL.Visible = True
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.Path & "\RL 3.14_rujukan.xls")
    Set oSheet = oWB.ActiveSheet
    
    '===============================================================================================================
  
    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    
    
'   For xx = 2 To 15
'
'      With oSheet
'
'            .Cells(xx, 1) = rsb("KodeExternal").value
'            .Cells(xx, 3) = rsb("KdRS").value
'            .Cells(xx, 2) = rsb("KotaKodyaKab").value
'            .Cells(xx, 4) = rsb("NamaRS").value
'            .Cells(xx, 5) = Format(dtptahun.value, "YYYY")
'      End With
'
'    Next xx
      



   '###################################################---splakuk revision on 2013-09-05


Set rs = Nothing
    strSQL = "Select distinct Kode,SMF from MasterRL314 order by Kode"
    Call msubRecFO(rs, strSQL)
    ProgressBar1.Max = rs.RecordCount
    k = 2
    For i = 1 To rs.RecordCount


            With oSheet

                .Cells(k, 1) = rsb("KodeExternal").value
            .Cells(k, 3) = rsb("KdRS").value
            .Cells(k, 2) = rsb("KotaKodyaKab").value
            .Cells(k, 4) = rsb("NamaRS").value
            .Cells(k, 5) = Format(dtptahun.value, "YYYY")
                
                strSQL1 = "Select isnull(sum(RujukanPuskesmas),0) as Puskesmas,isnull(sum(RujukanFaskesLain),0) as FaskesLain,isnull(sum(RujukanRS),0) as RujukanRS,isnull(sum(RujukanBidan),0) as RujukanBidan " & _
                "from V_RL314 where kode=" & rs(0).value & " and Year(TglMasuk)='" & Format(dtptahun, "yyyy") & "'"
             
                Call msubRecFO(rs1, strSQL1)
                .Cells(k, 8) = rs1(0).value
                .Cells(k, 9) = rs1(1).value
                .Cells(k, 10) = rs1(2).value
                
            End With

        k = k + 1
        rs.MoveNext
        ProgressBar1.value = ProgressBar1.value + 1
        lblPersen.Caption = Int(ProgressBar1.value * 100 / ProgressBar1.Max) & " %"

    Next i
   
'######################################################

    oXL.Visible = True
    Screen.MousePointer = vbDefault
Exit Sub
error:
    Call msubPesanError
'    MsgBox "Data Tidak Ada", vbInformation, "Validasi"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
oXL.Quit
End Sub
