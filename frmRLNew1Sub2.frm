VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRLNew1Sub2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL1.2 Indikator Pelayanan Rumah Sakit"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6135
   Icon            =   "frmRLNew1Sub2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6135
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
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
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   5895
      Begin MSComCtl2.DTPicker dtptahun 
         Height          =   375
         Left            =   2160
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
         CustomFormat    =   "yyyy"
         Format          =   138870787
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   6135
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   1320
         Width           =   1905
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRLNew1Sub2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmRLNew1Sub2"
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

Dim BOR As String
Dim LOS As String
Dim TOI As String
Dim BTO As String
Dim GDR As String
Dim NDR As String

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
'Pengimputan ke tabel Indikator Pelayanan RS dimulai dari bulan 1 Januari - bulan 31 Desember
    On Error GoTo error
    Dim intTgl As Integer
    Dim jj As Integer
    Dim kk As Integer
    Dim varBulan As Integer
    Dim varTahun As Integer
    

    Screen.MousePointer = vbHourglass

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    Set oWB = oXL.Workbooks.Open(App.Path & "\Formulir RL 1.2.xlsx")
    Set oSheet = oWB.ActiveSheet

    varTahun = Format(dtptahun.value, "yyyy")
   
   
  For varBulan = 1 To 12
  

    strSQL = "SELECT dbo.S_HitungHariDlmBln('" & varBulan & "','" & varTahun & "')  as varHari"
    Call msubRecFO(rs, strSQL)
    
    varHari = rs(0).value
    
      For intTgl = 1 To varHari
         dTglHitung = DateValue(str(intTgl) + "/" + str(varBulan) + "/" + str(varTahun))
         If sp_IndikatorPelayanan() = False Then Exit Sub
      Next intTgl

  Next varBulan

  Set rsx = Nothing

    strSQL = "Select avg (LOS) as LOS, avg(BOR) as BOR, avg(TOI) as TOI, avg(BTO) as BTO, avg(GDR) as GDR, avg(NDR) as NDR" & _
    " from V_IndikatorPelayananRSPerRuangan where Year(TglHitung) = '" & dtptahun.Year & "'"
    Call msubRecFO(rsx, strSQL)

'    LOS = FormatNumber(rsx(0).value, 2)
'    BOR = FormatNumber(rsx(1).value, 2)
'    TOI = FormatNumber(rsx(2).value, 2)
'    BTO = FormatNumber(rsx(3).value, 2)
'    GDR = FormatNumber(rsx(4).value, 2)
'    NDR = FormatNumber(rsx(5).value, 2)
    
    
    LOS = FormatNumber(rsx(0).value, 10)
    BOR = FormatNumber(rsx(1).value, 10)
    TOI = FormatNumber(rsx(2).value, 10)
    BTO = FormatNumber(rsx(3).value, 10)
    GDR = FormatNumber(rsx(4).value, 10)
    NDR = FormatNumber(rsx(5).value, 10)

    If rsx.RecordCount = 0 Then
        MsgBox "Data tidak ada", vbInformation, "Validasi"
        Exit Sub
    End If

    With oSheet
        .Cells(13, 1) = Format(Now, "yyyy")
        .Cells(13, 2) = Trim(IIf(IsNull(BOR), "", (BOR)))
        .Cells(13, 4) = Trim(IIf(IsNull(LOS), "", (LOS)))
        .Cells(13, 5) = Trim(IIf(IsNull(BTO), "", (BTO)))
        .Cells(13, 6) = Trim(IIf(IsNull(TOI), "", (TOI)))
        .Cells(13, 7) = Trim(IIf(IsNull(NDR), "", (NDR)))
        .Cells(13, 8) = Trim(IIf(IsNull(GDR), "", (GDR)))
    End With

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With oSheet
        .Cells(6, 3) = rsb("KdRS").value
        .Cells(7, 3) = rsb("NamaRS").value
        .Cells(8, 3) = Format(dtptahun.value, "YYYY")

    End With

    oXL.Visible = True
    Screen.MousePointer = vbDefault

    Exit Sub
error:
    Call msubPesanError
    Screen.MousePointer = vbDefault
End Sub

Private Function sp_IndikatorPelayanan() As Boolean
    sp_IndikatorPelayanan = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglHitung", adDate, adParamInput, , Format(dTglHitung, "yyyy/MM/dd 23:59:59"))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_IndikatorPelayananRS_New"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_IndikatorPelayanan = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

