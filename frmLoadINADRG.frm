VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmLoadINADRG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Load INA DRG"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15105
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLoadINADRG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   15105
   Begin VB.TextBox txtCariNoCM 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CheckBox chkCheckSemua 
      Caption         =   "Check Semua"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   7800
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton cmdConvExcel 
      Caption         =   "Convert Ke &Excel"
      Height          =   615
      Left            =   5280
      TabIndex        =   3
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CheckBox chkCheck 
      Height          =   210
      Left            =   240
      TabIndex        =   8
      Top             =   1360
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Simpan"
      Height          =   615
      Left            =   9120
      TabIndex        =   5
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton cmdBersih 
      Caption         =   "&Bersih"
      Height          =   615
      Left            =   7200
      TabIndex        =   4
      Top             =   7680
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog cdLoad 
      Left            =   240
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   615
      Left            =   13080
      TabIndex        =   7
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   615
      Left            =   11040
      TabIndex        =   6
      Top             =   7680
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid fgLoad 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   11245
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   10
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
      Left            =   13230
      Picture         =   "frmLoadINADRG.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Label Label3 
      Caption         =   "Cari NoCM"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   7560
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLoadINADRG.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmLoadINADRG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOpen As String
Dim strArray() As String
Dim m_RecLen As Long
Dim Bariske As Integer
Dim bolFile As Boolean

Public Function LoadText(FromFile As String) As String
    On Error GoTo handle
    bolFile = True
    'jika nama file tidak diisi
    If FromFile = "" Then
        bolFile = False
        Exit Function
    End If
    'check keberadaan file
    If FileExists(FromFile) = False Then
        MsgBox "File tidak ditemukan. Check File jika memang benar-benar ada.", vbCritical, "Error"
        bolFile = False
        Exit Function
    End If
    Dim sTemp As String
Close #1
Open FromFile For Input As #1
    sTemp = Input(LOF(1), 1)
Close #1
LoadText = sTemp
Exit Function
handle:
MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
End Function

Public Function FileExists(FileName As String) As Boolean
    'fungsi untuk mengecek keberadaan file
    On Error GoTo handle
    If FileLen(FileName) >= 0 Then: FileExists = True: Exit Function
handle:
    FileExists = False
End Function

Function HitJmlRec(strFile As String) As Integer
    HitJmlRec = 0
    For x = 1 To Len(strFile)
        If Asc(Mid(strFile, x, 1)) = 13 Then HitJmlRec = HitJmlRec + 1
    Next x
End Function

Private Sub GetBarisData(strFile As String)
    Dim m_RecBrs As Long
    Dim m_EOL As Long
    Dim m_JmlRec As Long
    Dim strTmp As String

    m_EOL = 1
    m_RecBrs = 1228
    Bariske = 0

    For g = 1 To Len(strFile)
        If g = m_RecBrs Then
            m_JmlRec = m_JmlRec + 1
        End If

        If (Asc(Mid(strFile, g, 1)) = 10) Or (Asc(Mid(strFile, g, 1)) = 13) Then
            GoTo simpantxt
        Else
            strTmp = strTmp + Mid(strFile, g, 1)
            GoTo nextFor
        End If
simpantxt:
        If strTmp = "" Then GoTo nextFor
        Bariske = Bariske + 1
        ReDim Preserve strArray(Bariske) 'strArray(m_JmlRec)
        strArray(Bariske) = strTmp 'tidak ada counter
        strTmp = ""
        m_RecBrs = m_RecBrs + 1228
nextFor:

    Next g
End Sub

Private Sub setgrid()
    Dim i As Integer
    With fgLoad
        .Rows = 2
        .Cols = 112

        'judul
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Hospital ID"
        .TextMatrix(0, 2) = "Hospital Type"
        .TextMatrix(0, 3) = "Medical Record Number"
        .TextMatrix(0, 4) = "Patient Class"
        .TextMatrix(0, 5) = "Payment"
        .TextMatrix(0, 6) = "Recid"
        .TextMatrix(0, 7) = "Filler1"
        .TextMatrix(0, 8) = "Patient Type Designation"
        .TextMatrix(0, 9) = "Admin Date"
        .TextMatrix(0, 10) = "Discharge Date"
        .TextMatrix(0, 11) = "ALOS"
        .TextMatrix(0, 12) = "Birth Date"
        .TextMatrix(0, 13) = "Age In Years"
        .TextMatrix(0, 14) = "Age In Days"
        .TextMatrix(0, 15) = "Filler2"
        .TextMatrix(0, 16) = "Sex"
        .TextMatrix(0, 17) = "Discharge Disposition/Status"
        .TextMatrix(0, 18) = "Birth Weight In Grams"

        .TextMatrix(0, 19) = "Diagnoses_1"
        .TextMatrix(0, 20) = "Diagnoses_2"
        .TextMatrix(0, 21) = "Diagnoses_3"
        .TextMatrix(0, 22) = "Diagnoses_4"
        .TextMatrix(0, 23) = "Diagnoses_5"
        .TextMatrix(0, 24) = "Diagnoses_6"
        .TextMatrix(0, 25) = "Diagnoses_7"
        .TextMatrix(0, 26) = "Diagnoses_8"
        .TextMatrix(0, 27) = "Diagnoses_9"
        .TextMatrix(0, 28) = "Diagnoses_10"
        .TextMatrix(0, 29) = "Diagnoses_11"
        .TextMatrix(0, 30) = "Diagnoses_12"
        .TextMatrix(0, 31) = "Diagnoses_13"
        .TextMatrix(0, 32) = "Diagnoses_14"
        .TextMatrix(0, 33) = "Diagnoses_15"
        .TextMatrix(0, 34) = "Diagnoses_16"
        .TextMatrix(0, 35) = "Diagnoses_17"
        .TextMatrix(0, 36) = "Diagnoses_18"
        .TextMatrix(0, 37) = "Diagnoses_19"
        .TextMatrix(0, 38) = "Diagnoses_20"
        .TextMatrix(0, 39) = "Diagnoses_21"
        .TextMatrix(0, 40) = "Diagnoses_22"
        .TextMatrix(0, 41) = "Diagnoses_23"
        .TextMatrix(0, 42) = "Diagnoses_24"
        .TextMatrix(0, 43) = "Diagnoses_25"
        .TextMatrix(0, 44) = "Diagnoses_26"
        .TextMatrix(0, 45) = "Diagnoses_27"
        .TextMatrix(0, 46) = "Diagnoses_28"
        .TextMatrix(0, 47) = "Diagnoses_29"
        .TextMatrix(0, 48) = "Diagnoses_30"

        .TextMatrix(0, 49) = "Procedure_1"
        .TextMatrix(0, 50) = "Procedure_2"
        .TextMatrix(0, 51) = "Procedure_3"
        .TextMatrix(0, 52) = "Procedure_4"
        .TextMatrix(0, 53) = "Procedure_5"
        .TextMatrix(0, 54) = "Procedure_6"
        .TextMatrix(0, 55) = "Procedure_7"
        .TextMatrix(0, 56) = "Procedure_8"
        .TextMatrix(0, 57) = "Procedure_9"
        .TextMatrix(0, 58) = "Procedure_10"
        .TextMatrix(0, 59) = "Procedure_11"
        .TextMatrix(0, 60) = "Procedure_12"
        .TextMatrix(0, 61) = "Procedure_13"
        .TextMatrix(0, 62) = "Procedure_14"
        .TextMatrix(0, 63) = "Procedure_15"
        .TextMatrix(0, 64) = "Procedure_16"
        .TextMatrix(0, 65) = "Procedure_17"
        .TextMatrix(0, 66) = "Procedure_18"
        .TextMatrix(0, 67) = "Procedure_19"
        .TextMatrix(0, 68) = "Procedure_20"
        .TextMatrix(0, 69) = "Procedure_21"
        .TextMatrix(0, 70) = "Procedure_22"
        .TextMatrix(0, 71) = "Procedure_23"
        .TextMatrix(0, 72) = "Procedure_24"
        .TextMatrix(0, 73) = "Procedure_25"
        .TextMatrix(0, 74) = "Procedure_26"
        .TextMatrix(0, 75) = "Procedure_27"
        .TextMatrix(0, 76) = "Procedure_28"
        .TextMatrix(0, 77) = "Procedure_29"
        .TextMatrix(0, 78) = "Procedure_30"

        .TextMatrix(0, 79) = "Grouper Type"
        .TextMatrix(0, 80) = "Patient Type Used"
        .TextMatrix(0, 81) = "DRG"
        .TextMatrix(0, 82) = "Grouper Status"

        .TextMatrix(0, 83) = "Diagnosis Validity Flag_1"
        .TextMatrix(0, 84) = "Diagnosis Validity Flag_2"
        .TextMatrix(0, 85) = "Diagnosis Validity Flag_3"
        .TextMatrix(0, 86) = "Diagnosis Validity Flag_4"
        .TextMatrix(0, 87) = "Diagnosis Validity Flag_5"

        .TextMatrix(0, 88) = "Procedure Validity Flag_1"
        .TextMatrix(0, 89) = "Procedure Validity Flag_2"
        .TextMatrix(0, 90) = "Procedure Validity Flag_3"
        .TextMatrix(0, 91) = "Procedure Validity Flag_4"
        .TextMatrix(0, 92) = "Procedure Validity Flag_5"

        .TextMatrix(0, 93) = "Procedure Class_1"
        .TextMatrix(0, 94) = "Procedure Class_2"
        .TextMatrix(0, 95) = "Procedure Class_3"
        .TextMatrix(0, 96) = "Procedure Class_4"
        .TextMatrix(0, 97) = "Procedure Class_5"

        .TextMatrix(0, 98) = "Birth Weight Used"
        .TextMatrix(0, 99) = "Birth Weight Source"
        .TextMatrix(0, 100) = "Medical Surgical Flag"

        .TextMatrix(0, 101) = "Procedure DRG_1"
        .TextMatrix(0, 102) = "Procedure DRG_2"
        .TextMatrix(0, 103) = "Procedure DRG_3"
        .TextMatrix(0, 104) = "Procedure DRG_4"
        .TextMatrix(0, 105) = "Procedure DRG_5"

        .TextMatrix(0, 106) = "Procedure Statistic_1"
        .TextMatrix(0, 107) = "Procedure Statistic_2"
        .TextMatrix(0, 108) = "Procedure Statistic_3"
        .TextMatrix(0, 109) = "Procedure Statistic_4"
        .TextMatrix(0, 110) = "Procedure Statistic_5"

        .ColWidth(0) = 340
        .ColWidth(1) = 900
        .ColWidth(2) = 1100
        .ColWidth(3) = 1800
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 600
        .ColWidth(7) = 1000
        .ColWidth(8) = 1870
        .ColWidth(9) = 1000
        .ColWidth(10) = 1200
        .ColWidth(11) = 600
        .ColWidth(12) = 1000
        .ColWidth(13) = 1000
        .ColWidth(14) = 1000
        .ColWidth(15) = 1000
        .ColWidth(16) = 600
        .ColWidth(17) = 2120
        .ColWidth(18) = 1800

        .ColWidth(19) = 1000
        .ColWidth(20) = 1000
        .ColWidth(21) = 1000
        .ColWidth(22) = 1000
        .ColWidth(23) = 1000
        .ColWidth(24) = 0
        .ColWidth(25) = 0
        .ColWidth(26) = 0
        .ColWidth(27) = 0
        .ColWidth(28) = 0
        .ColWidth(29) = 0
        .ColWidth(30) = 0
        .ColWidth(31) = 0
        .ColWidth(32) = 0
        .ColWidth(33) = 0
        .ColWidth(34) = 0
        .ColWidth(35) = 0
        .ColWidth(36) = 0
        .ColWidth(37) = 0
        .ColWidth(38) = 0
        .ColWidth(39) = 0
        .ColWidth(40) = 0
        .ColWidth(41) = 0
        .ColWidth(42) = 0
        .ColWidth(43) = 0
        .ColWidth(44) = 0
        .ColWidth(45) = 0
        .ColWidth(46) = 0
        .ColWidth(47) = 0
        .ColWidth(48) = 0

        .ColWidth(49) = 1000
        .ColWidth(50) = 1000
        .ColWidth(51) = 1000
        .ColWidth(52) = 1000
        .ColWidth(53) = 1000
        .ColWidth(54) = 0
        .ColWidth(55) = 0
        .ColWidth(56) = 0
        .ColWidth(57) = 0
        .ColWidth(58) = 0
        .ColWidth(59) = 0
        .ColWidth(60) = 0
        .ColWidth(61) = 0
        .ColWidth(62) = 0
        .ColWidth(63) = 0
        .ColWidth(64) = 0
        .ColWidth(65) = 0
        .ColWidth(66) = 0
        .ColWidth(67) = 0
        .ColWidth(68) = 0
        .ColWidth(69) = 0
        .ColWidth(70) = 0
        .ColWidth(71) = 0
        .ColWidth(72) = 0
        .ColWidth(73) = 0
        .ColWidth(74) = 0
        .ColWidth(75) = 0
        .ColWidth(76) = 0
        .ColWidth(77) = 0
        .ColWidth(78) = 0

        .ColWidth(79) = 1650
        .ColWidth(80) = 1400
        .ColWidth(81) = 900
        .ColWidth(82) = 1200

        .ColWidth(83) = 1850
        .ColWidth(84) = 1850
        .ColWidth(85) = 1850
        .ColWidth(86) = 1850
        .ColWidth(87) = 1850

        .ColWidth(88) = 1900
        .ColWidth(89) = 1900
        .ColWidth(90) = 1900
        .ColWidth(91) = 1900
        .ColWidth(92) = 1900

        .ColWidth(93) = 1450
        .ColWidth(94) = 1450
        .ColWidth(95) = 1450
        .ColWidth(96) = 1450
        .ColWidth(97) = 1450

        .ColWidth(98) = 1400
        .ColWidth(99) = 1500
        .ColWidth(100) = 1550

        .ColWidth(101) = 1400
        .ColWidth(102) = 1400
        .ColWidth(103) = 1400
        .ColWidth(104) = 1400
        .ColWidth(105) = 1400

        .ColWidth(106) = 1600
        .ColWidth(107) = 1600
        .ColWidth(108) = 1600
        .ColWidth(109) = 1600
        .ColWidth(110) = 1600

        .ColWidth(111) = 0

        .ColAlignment(0) = flexAlignCenterCenter
    End With
End Sub

Private Sub chkCheck_Click()
    On Error GoTo errLoad

    If chkCheck.value = vbChecked Then
        fgLoad.TextMatrix(fgLoad.Row, fgLoad.Col) = Chr$(187)
        fgLoad.TextMatrix(fgLoad.Row, 111) = 1
    Else
        fgLoad.TextMatrix(fgLoad.Row, fgLoad.Col) = ""
        fgLoad.TextMatrix(fgLoad.Row, 111) = 0
    End If

    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub chkCheck_LostFocus()
    chkCheck.Visible = False
End Sub

Private Sub chkCheckSemua_Click()
    On Error GoTo hell
    Dim i As Integer

    If chkCheckSemua.value = Checked Then
        For i = 1 To fgLoad.Rows - 1
            fgLoad.TextMatrix(i, 0) = Chr$(187)
            fgLoad.TextMatrix(i, 111) = 1
        Next i
    Else
        For i = 1 To fgLoad.Rows - 1
            fgLoad.TextMatrix(i, 0) = ""
            fgLoad.TextMatrix(i, 111) = 0
        Next i
    End If

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdConvExcel_Click()
    On Error GoTo hell
    Dim sPath As String
    Dim sData As String
    Dim sJudulxls As String
    Dim c As Integer

    cdLoad.FileName = ""
    cdLoad.ShowSave
    If cdLoad.FileName = "" Then Exit Sub
    sPath = cdLoad.FileName & ".txt"
    sJudulxls = "Kdrs;Klsrs;Norm;Klsrawat;Biaya;Jnsrawat;Tglmsk;Tglklr;Los;Tgllhr;UmurThn;UmurHari;JK;CaraPlg;Berat;Dutama;D1;D2;D3;D4;D5;D6;D7;D8;D9;D10;D11;D12;D13;D14;D15;D16;D17;D18;D19;D20;D21;D22;D23;D24;D25;D26;D27;D28;D29;P1;P2;P3;P4;P5;P6;P7;P8;P9;P10;P11;P12;P13;P14;P15;P16;P17;P18;P19;P20;P21;P22;P23;P24;P25;P26;P27;P28;P29;P30;Recid;Inadrg;Tarif;Deskripsi;ALOS;"
    sData = ""
    strSQL = ""
    strSQL = "Select * From V_LaporanINADRGJEP2009"
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        MsgBox "Data tidak ada, coba proses simpan lagi", vbExclamation, "Validasi"
        Exit Sub
    End If
    rs.MoveFirst
    With fgLoad
        For c = 1 To rs.RecordCount '.Rows - 1
            sData = sData & Trim(rs("KdRS").value) & ";" & Trim(rs("TipeRS").value) & ";" & Trim(rs("NoRM").value) & ";" & Val(rs("KelasPerawatan").value) & ";" & Val(rs("TarifRS").value) & ";" & Val(rs("JenisPerawatan").value) & ";" & Format(Trim(rs("TglMasuk").value), "dd/MM/yyyy") & ";" & Format(Trim(rs("TglKeluar").value), "dd/MM/yyyy") & ";" & Val(rs("LOS").value) & ";" & Format(Trim(rs("TglLahir").value), "dd/MM/yyyy") & ";" & Val(rs("umurThn").value) & ";" & Val(rs("UmurHari").value) & ";" & Val(rs("JK").value) & ";" & Val(rs("CaraPulang").value) & ";" & Val(rs("BeratLahir").value) & ";" & Trim(rs("ICD10_1").value) & ";" & _
            "" & IIf(Len(Trim(rs("ICD10_2").value)) = 0, Null, Trim(rs("ICD10_2").value)) & ";" & IIf(Len(Trim(rs("ICD10_3").value)) = 0, Null, Trim(rs("ICD10_3").value)) & ";" & IIf(Len(Trim(rs("ICD10_4").value)) = 0, Null, Trim(rs("ICD10_4").value)) & ";" & IIf(Len(Trim(rs("ICD10_5").value)) = 0, Null, Trim(rs("ICD10_5").value)) & ";" & IIf(Len(Trim(rs("ICD10_6").value)) = 0, Null, Trim(rs("ICD10_6").value)) & ";" & IIf(Len(Trim(rs("ICD10_7").value)) = 0, Null, Trim(rs("ICD10_7").value)) & ";" & IIf(Len(Trim(rs("ICD10_8").value)) = 0, Null, Trim(rs("ICD10_8").value)) & ";" & IIf(Len(Trim(rs("ICD10_9").value)) = 0, Null, Trim(rs("ICD10_9").value)) & ";" & IIf(Len(Trim(rs("ICD10_10").value)) = 0, Null, Trim(rs("ICD10_10").value)) & ";" & _
            "" & IIf(Len(Trim(rs("ICD10_11").value)) = 0, Null, Trim(rs("ICD10_11").value)) & ";" & IIf(Len(Trim(rs("ICD10_12").value)) = 0, Null, Trim(rs("ICD10_12").value)) & ";" & IIf(Len(Trim(rs("ICD10_13").value)) = 0, Null, Trim(rs("ICD10_13").value)) & ";" & IIf(Len(Trim(rs("ICD10_14").value)) = 0, Null, Trim(rs("ICD10_14").value)) & ";" & IIf(Len(Trim(rs("ICD10_15").value)) = 0, Null, Trim(rs("ICD10_15").value)) & ";" & IIf(Len(Trim(rs("ICD10_16").value)) = 0, Null, Trim(rs("ICD10_16").value)) & ";" & IIf(Len(Trim(rs("ICD10_17").value)) = 0, Null, Trim(rs("ICD10_17").value)) & ";" & IIf(Len(Trim(rs("ICD10_18").value)) = 0, Null, Trim(rs("ICD10_18").value)) & ";" & IIf(Len(Trim(rs("ICD10_19").value)) = 0, Null, Trim(rs("ICD10_19").value)) & ";" & _
            "" & IIf(Len(Trim(rs("ICD10_20").value)) = 0, Null, Trim(rs("ICD10_20").value)) & ";" & IIf(Len(Trim(rs("ICD10_21").value)) = 0, Null, Trim(rs("ICD10_21").value)) & ";" & IIf(Len(Trim(rs("ICD10_22").value)) = 0, Null, Trim(rs("ICD10_22").value)) & ";" & IIf(Len(Trim(rs("ICD10_23").value)) = 0, Null, Trim(rs("ICD10_23").value)) & ";" & IIf(Len(Trim(rs("ICD10_24").value)) = 0, Null, Trim(rs("ICD10_24").value)) & ";" & IIf(Len(Trim(rs("ICD10_25").value)) = 0, Null, Trim(rs("ICD10_25").value)) & ";" & IIf(Len(Trim(rs("ICD10_26").value)) = 0, Null, Trim(rs("ICD10_26").value)) & ";" & IIf(Len(Trim(rs("ICD10_27").value)) = 0, Null, Trim(rs("ICD10_27").value)) & ";" & IIf(Len(Trim(rs("ICD10_28").value)) = 0, Null, Trim(rs("ICD10_28").value)) & ";" & _
            "" & IIf(Len(Trim(rs("ICD10_29").value)) = 0, Null, Trim(rs("ICD10_29").value)) & ";" & IIf(Len(Trim(rs("ICD10_30").value)) = 0, Null, Trim(rs("ICD10_30").value)) & ";" & _
            "" & IIf(Len(Trim(rs("ICD9_1").value)) = 0, Null, Trim(rs("ICD9_1").value)) & ";" & IIf(Len(Trim(rs("ICD9_2").value)) = 0, Null, Trim(rs("ICD9_2").value)) & ";" & IIf(Len(Trim(rs("ICD9_3").value)) = 0, Null, Trim(rs("ICD9_3").value)) & ";" & IIf(Len(Trim(rs("ICD9_4").value)) = 0, Null, Trim(rs("ICD9_4").value)) & ";" & IIf(Len(Trim(rs("ICD9_5").value)) = 0, Null, Trim(rs("ICD9_5").value)) & ";" & IIf(Len(Trim(rs("ICD9_6").value)) = 0, Null, Trim(rs("ICD9_6").value)) & ";" & IIf(Len(Trim(rs("ICD9_7").value)) = 0, Null, Trim(rs("ICD9_7").value)) & ";" & IIf(Len(Trim(rs("ICD9_8").value)) = 0, Null, Trim(rs("ICD9_8").value)) & ";" & IIf(Len(Trim(rs("ICD9_9").value)) = 0, Null, Trim(rs("ICD9_9").value)) & ";" & _
            "" & IIf(Len(Trim(rs("ICD9_10").value)) = 0, Null, Trim(rs("ICD9_10").value)) & ";" & IIf(Len(Trim(rs("ICD9_11").value)) = 0, Null, Trim(rs("ICD9_11").value)) & ";" & IIf(Len(Trim(rs("ICD9_12").value)) = 0, Null, Trim(rs("ICD9_12").value)) & ";" & IIf(Len(Trim(rs("ICD9_13").value)) = 0, Null, Trim(rs("ICD9_13").value)) & ";" & IIf(Len(Trim(rs("ICD9_14").value)) = 0, Null, Trim(rs("ICD9_14").value)) & ";" & IIf(Len(Trim(rs("ICD9_15").value)) = 0, Null, Trim(rs("ICD9_15").value)) & ";" & IIf(Len(Trim(rs("ICD9_16").value)) = 0, Null, Trim(rs("ICD9_16").value)) & ";" & IIf(Len(Trim(rs("ICD9_17").value)) = 0, Null, Trim(rs("ICD9_17").value)) & ";" & IIf(Len(Trim(rs("ICD9_18").value)) = 0, Null, Trim(rs("ICD9_18").value)) & ";" & _
            "" & IIf(Len(Trim(rs("ICD9_19").value)) = 0, Null, Trim(rs("ICD9_19").value)) & ";" & IIf(Len(Trim(rs("ICD9_20").value)) = 0, Null, Trim(rs("ICD9_20").value)) & ";" & IIf(Len(Trim(rs("ICD9_21").value)) = 0, Null, Trim(rs("ICD9_21").value)) & ";" & IIf(Len(Trim(rs("ICD9_22").value)) = 0, Null, Trim(rs("ICD9_22").value)) & ";" & IIf(Len(Trim(rs("ICD9_23").value)) = 0, Null, Trim(rs("ICD9_23").value)) & ";" & IIf(Len(Trim(rs("ICD9_24").value)) = 0, Null, Trim(rs("ICD9_24").value)) & ";" & IIf(Len(Trim(rs("ICD9_25").value)) = 0, Null, Trim(rs("ICD9_25").value)) & ";" & IIf(Len(Trim(rs("ICD9_26").value)) = 0, Null, Trim(rs("ICD9_26").value)) & ";" & IIf(Len(Trim(rs("ICD9_27").value)) = 0, Null, Trim(rs("ICD9_27").value)) & ";" & _
            "" & IIf(Len(Trim(rs("ICD9_28").value)) = 0, Null, Trim(rs("ICD9_28").value)) & ";" & IIf(Len(Trim(rs("ICD9_29").value)) = 0, Null, Trim(rs("ICD9_29").value)) & ";" & IIf(Len(Trim(rs("ICD9_30").value)) = 0, Null, Trim(rs("ICD9_30").value)) & ";;" & Trim(rs("KdINADRG").value) & ";" & CDec(rs("TarifINADRG").value) & ";" & Trim(rs("DESKRIPSI").value) & ";" & CDec(rs("ALOS").value) & ";" & vbNewLine
            rs.MoveNext
        Next c
    End With
    sJudulxls = sJudulxls & vbNewLine & sData
    Call subSimpanData(sJudulxls, sPath)
    MsgBox "Data berhasil disimpan", vbInformation, "Berhasil.."
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subSimpanData(strData As String, NamaFile As String)
    On Error GoTo hell
    Open NamaFile For Append As #1
        Print #1, strData
    Close #1

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo hell
    Dim s As Integer
    Dim sICD10 As String
    Dim sICD9 As String
    Dim bolDataFalse As Boolean
    Dim varTglUpdate  As Date

    bolDataFalse = False
    varTglUpdate = Now

    With fgLoad
        For s = 1 To .Rows - 1
            If .TextMatrix(s, 111) = 1 Then
                strSQL = " update OOOINADRG set " & _
                " TglPos2 = '" & Format(varTglUpdate, "yyyy/MM/dd HH:mm:ss") & "', idPegawai2 = '" & strIDPegawai & "', kdKelasINADRG = " & Val(.TextMatrix(s, 4)) & ", JenisPerawatan = " & Val(.TextMatrix(s, 8)) & ", TglMasukINADRG = '" & Format(Trim(.TextMatrix(s, 9)), "yyyy/MM/dd") & "', TglPulangINADRG = '" & Format(Trim(.TextMatrix(s, 10)), "yyyy/MM/dd") & "', ALOS = " & Val(.TextMatrix(s, 11)) & ", kdStatusKeluarINADRG = " & Val(.TextMatrix(s, 17)) & "," & _
                " ICD10_1 = '" & IIf(Trim(.TextMatrix(s, 19)) = "", Null, Trim(.TextMatrix(s, 19))) & "', ICD10_2 = '" & IIf(Trim(.TextMatrix(s, 20)) = "", Null, Trim(.TextMatrix(s, 20))) & "', ICD10_3 = '" & IIf(Trim(.TextMatrix(s, 21)) = "", Null, Trim(.TextMatrix(s, 21))) & "', ICD10_4 = '" & IIf(Trim(.TextMatrix(s, 22)) = "", Null, Trim(.TextMatrix(s, 22))) & "', ICD10_5 = '" & IIf(Trim(.TextMatrix(s, 23)) = "", Null, Trim(.TextMatrix(s, 23))) & "', ICD10_6 = '" & IIf(Trim(.TextMatrix(s, 24)) = "", Null, Trim(.TextMatrix(s, 24))) & "', ICD10_7 = '" & IIf(Trim(.TextMatrix(s, 25)) = "", Null, Trim(.TextMatrix(s, 25))) & "', ICD10_8 = '" & IIf(Trim(.TextMatrix(s, 26)) = "", Null, Trim(.TextMatrix(s, 26))) & "'," & _
                " ICD10_9 = '" & IIf(Trim(.TextMatrix(s, 27)) = "", Null, Trim(.TextMatrix(s, 27))) & "', ICD10_10 = '" & IIf(Trim(.TextMatrix(s, 28)) = "", Null, Trim(.TextMatrix(s, 28))) & "', ICD10_11 = '" & IIf(Trim(.TextMatrix(s, 29)) = "", Null, Trim(.TextMatrix(s, 29))) & "', ICD10_12 = '" & IIf(Trim(.TextMatrix(s, 30)) = "", Null, Trim(.TextMatrix(s, 30))) & "', ICD10_13 = '" & IIf(Trim(.TextMatrix(s, 31)) = "", Null, Trim(.TextMatrix(s, 31))) & "', ICD10_14 = '" & IIf(Trim(.TextMatrix(s, 32)) = "", Null, Trim(.TextMatrix(s, 32))) & "', ICD10_15 = '" & IIf(Trim(.TextMatrix(s, 33)) = "", Null, Trim(.TextMatrix(s, 33))) & "', ICD10_16 = '" & IIf(Trim(.TextMatrix(s, 34)) = "", Null, Trim(.TextMatrix(s, 34))) & "'," & _
                " ICD10_17 = '" & IIf(Trim(.TextMatrix(s, 35)) = "", Null, Trim(.TextMatrix(s, 35))) & "', ICD10_18 = '" & IIf(Trim(.TextMatrix(s, 36)) = "", Null, Trim(.TextMatrix(s, 36))) & "', ICD10_19 = '" & IIf(Trim(.TextMatrix(s, 37)) = "", Null, Trim(.TextMatrix(s, 37))) & "', ICD10_20 = '" & IIf(Trim(.TextMatrix(s, 38)) = "", Null, Trim(.TextMatrix(s, 38))) & "', ICD10_21 = '" & IIf(Trim(.TextMatrix(s, 39)) = "", Null, Trim(.TextMatrix(s, 39))) & "', ICD10_22 = '" & IIf(Trim(.TextMatrix(s, 40)) = "", Null, Trim(.TextMatrix(s, 40))) & "', ICD10_23 = '" & IIf(Trim(.TextMatrix(s, 41)) = "", Null, Trim(.TextMatrix(s, 41))) & "', ICD10_24 = '" & IIf(Trim(.TextMatrix(s, 42)) = "", Null, Trim(.TextMatrix(s, 42))) & "'," & _
                " ICD10_25 = '" & IIf(Trim(.TextMatrix(s, 43)) = "", Null, Trim(.TextMatrix(s, 43))) & "', ICD10_26 = '" & IIf(Trim(.TextMatrix(s, 44)) = "", Null, Trim(.TextMatrix(s, 44))) & "', ICD10_27 = '" & IIf(Trim(.TextMatrix(s, 45)) = "", Null, Trim(.TextMatrix(s, 45))) & "', ICD10_28 = '" & IIf(Trim(.TextMatrix(s, 46)) = "", Null, Trim(.TextMatrix(s, 46))) & "', ICD10_29 = '" & IIf(Trim(.TextMatrix(s, 47)) = "", Null, Trim(.TextMatrix(s, 47))) & "', ICD10_30 = '" & IIf(Trim(.TextMatrix(s, 48)) = "", Null, Trim(.TextMatrix(s, 48))) & "'," & _
                " ICD9_1 = '" & IIf(Trim(.TextMatrix(s, 49)) = "", Null, Trim(.TextMatrix(s, 49))) & "', ICD9_2 = '" & IIf(Trim(.TextMatrix(s, 50)) = "", Null, Trim(.TextMatrix(s, 50))) & "', ICD9_3 = '" & IIf(Trim(.TextMatrix(s, 51)) = "", Null, Trim(.TextMatrix(s, 51))) & "', ICD9_4 = '" & IIf(Trim(.TextMatrix(s, 52)) = "", Null, Trim(.TextMatrix(s, 52))) & "', ICD9_5 = '" & IIf(Trim(.TextMatrix(s, 53)) = "", Null, Trim(.TextMatrix(s, 53))) & "', ICD9_6 = '" & IIf(Trim(.TextMatrix(s, 54)) = "", Null, Trim(.TextMatrix(s, 54))) & "', ICD9_7 = '" & IIf(Trim(.TextMatrix(s, 55)) = "", Null, Trim(.TextMatrix(s, 55))) & "', ICD9_8 = '" & IIf(Trim(.TextMatrix(s, 56)) = "", Null, Trim(.TextMatrix(s, 56))) & "'," & _
                " ICD9_9 = '" & IIf(Trim(.TextMatrix(s, 57)) = "", Null, Trim(.TextMatrix(s, 57))) & "', ICD9_10 = '" & IIf(Trim(.TextMatrix(s, 58)) = "", Null, Trim(.TextMatrix(s, 58))) & "', ICD9_11 = '" & IIf(Trim(.TextMatrix(s, 59)) = "", Null, Trim(.TextMatrix(s, 59))) & "', ICD9_12 = '" & IIf(Trim(.TextMatrix(s, 60)) = "", Null, Trim(.TextMatrix(s, 60))) & "', ICD9_13 = '" & IIf(Trim(.TextMatrix(s, 61)) = "", Null, Trim(.TextMatrix(s, 61))) & "', ICD9_14 = '" & IIf(Trim(.TextMatrix(s, 62)) = "", Null, Trim(.TextMatrix(s, 62))) & "', ICD9_15 = '" & IIf(Trim(.TextMatrix(s, 63)) = "", Null, Trim(.TextMatrix(s, 63))) & "', ICD9_16 = '" & IIf(Trim(.TextMatrix(s, 64)) = "", Null, Trim(.TextMatrix(s, 64))) & "'," & _
                " ICD9_17 = '" & IIf(Trim(.TextMatrix(s, 65)) = "", Null, Trim(.TextMatrix(s, 65))) & "', ICD9_18 = '" & IIf(Trim(.TextMatrix(s, 66)) = "", Null, Trim(.TextMatrix(s, 66))) & "', ICD9_19 = '" & IIf(Trim(.TextMatrix(s, 67)) = "", Null, Trim(.TextMatrix(s, 67))) & "', ICD9_20 = '" & IIf(Trim(.TextMatrix(s, 68)) = "", Null, Trim(.TextMatrix(s, 68))) & "', ICD9_21 = '" & IIf(Trim(.TextMatrix(s, 69)) = "", Null, Trim(.TextMatrix(s, 69))) & "', ICD9_22 = '" & IIf(Trim(.TextMatrix(s, 70)) = "", Null, Trim(.TextMatrix(s, 70))) & "', ICD9_23 = '" & IIf(Trim(.TextMatrix(s, 71)) = "", Null, Trim(.TextMatrix(s, 71))) & "', ICD9_24 = '" & IIf(Trim(.TextMatrix(s, 72)) = "", Null, Trim(.TextMatrix(s, 72))) & "'," & _
                " ICD9_25 = '" & IIf(Trim(.TextMatrix(s, 73)) = "", Null, Trim(.TextMatrix(s, 73))) & "', ICD9_26 = '" & IIf(Trim(.TextMatrix(s, 74)) = "", Null, Trim(.TextMatrix(s, 74))) & "', ICD9_27 = '" & IIf(Trim(.TextMatrix(s, 75)) = "", Null, Trim(.TextMatrix(s, 75))) & "', ICD9_28 = '" & IIf(Trim(.TextMatrix(s, 76)) = "", Null, Trim(.TextMatrix(s, 76))) & "', ICD9_29 = '" & IIf(Trim(.TextMatrix(s, 77)) = "", Null, Trim(.TextMatrix(s, 77))) & "', ICD9_30 = '" & IIf(Trim(.TextMatrix(s, 78)) = "", Null, Trim(.TextMatrix(s, 78))) & "'," & _
                " KdINADRG = '" & IIf(Trim(.TextMatrix(s, 81)) = "", Null, Trim(.TextMatrix(s, 81))) & "'" & _
                " where   NoCM = '" & Trim(.TextMatrix(s, 3)) & "' and year(TglMasuk) = " & Format(Trim(.TextMatrix(s, 9)), "yyyy") & " and month(TglMasuk) = " & Format(Trim(.TextMatrix(s, 9)), "MM") & " and day(TglMasuk) = " & Format(Trim(.TextMatrix(s, 9)), "dd") & " and " & _
                " year(TglPulang) = " & Format(Trim(.TextMatrix(s, 10)), "yyyy") & " and month(TglPulang) = " & Format(Trim(.TextMatrix(s, 10)), "MM") & " and day(TglPulang) = " & Format(Trim(.TextMatrix(s, 10)), "dd") & " and Biaya+BiayaGD = " & Val(.TextMatrix(s, 5)) & ""

                Set rs = Nothing
                Call msubRecFO(rs, strSQL)
            Else
                bolDataFalse = True
            End If
        Next s
    End With

    MsgBox "Data berhasil disimpan", vbInformation, "Sukses"
    cmdSave.Enabled = False
    cmdBersih.SetFocus
    If bolDataFalse = True Then
        If MsgBox("Apakah data yang tidak dipilih mau disimpan ?", vbQuestion + vbYesNo, "Validasi") = vbNo Then GoTo saveChecklist
        Call simpanDataFalse("0")
saveChecklist:
        If MsgBox("Apakah data yang dipilih mau disimpan ?", vbQuestion + vbYesNo, "Validasi") = vbNo Then Exit Sub
        Call simpanDataFalse("1")
    End If

    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub simpanDataFalse(f_strKdSimpan As String) 'jika 1 = yg dichecklist, 0 = tdk dichecklist
    On Error GoTo hell
    Dim sFile As String
    Dim x, y As Integer
awal:
    cdLoad.FileName = ""
    strTampung = ""
    sFile = ""
    cdLoad.ShowSave
    If cdLoad.FileName = "" Then Exit Sub
    strICD10 = ""
    strICD9 = ""

    sFile = cdLoad.FileName & ".txt"
    Screen.MousePointer = vbHourglass

    With fgLoad
        For i = 1 To fgLoad.Rows - 1
            If .TextMatrix(i, 111) = f_strKdSimpan Then
                Call ValidasiINADRG("angka", Len(.TextMatrix(i, 1)), 7, .TextMatrix(i, 1)) 'hospital ID
                Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 2)), 1, .TextMatrix(i, 2))  'hospital type
                Call ValidasiINADRG("angka", Len(.TextMatrix(i, 3)), 11, .TextMatrix(i, 3)) 'medical record number
                Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 4)), 1, .TextMatrix(i, 4)) 'patient class
                Call ValidasiINADRG("angka", Len(.TextMatrix(i, 5)), 9, Val(.TextMatrix(i, 5)))  'payment
                Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 6)), 10, .TextMatrix(i, 6)) 'recid
                Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 7)), 25, .TextMatrix(i, 7)) 'filler1
                Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 8)), 1, .TextMatrix(i, 8)) 'patient type designation
                Call ValidasiINADRG("angka", Len(Format(.TextMatrix(i, 9), "dd/MM/yyyy")), 10, Format(.TextMatrix(i, 9), "dd/MM/yyyy")) 'admin date
                Call ValidasiINADRG("angka", Len(Format(.TextMatrix(i, 10), "dd/MM/yyyy")), 10, Format(.TextMatrix(i, 10), "dd/MM/yyyy")) 'discharge date
                Call ValidasiINADRG("angka", Len(.TextMatrix(i, 11)), 4, .TextMatrix(i, 11)) 'ALOS
                Call ValidasiINADRG("angka", Len(Format(.TextMatrix(i, 12), "dd/MM/yyyy")), 10, Format(.TextMatrix(i, 12), "dd/MM/yyyy")) 'birth date
                Call ValidasiINADRG("angka", Len(.TextMatrix(i, 13)), 3, .TextMatrix(i, 13)) 'umur thn
                Call ValidasiINADRG("angka", Len(.TextMatrix(i, 14)), 3, .TextMatrix(i, 14)) 'umur hr
                Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 15)), 3, .TextMatrix(i, 15)) 'filler2
                Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 16)), 1, .TextMatrix(i, 16)) 'sex
                Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 17)), 2, .TextMatrix(i, 17)) 'discharge disposition/status
                Call ValidasiINADRG("angka", Len(.TextMatrix(i, 18)), 4, .TextMatrix(i, 18)) 'birth weight in gram
                'principal diagnosis icd10 utama & diagnoses icd10 sekunder
                x = 0
                For x = 19 To (19 + 30) - 1
                    Call ValidasiINADRG("huruf", Len(.TextMatrix(i, x)), 10, .TextMatrix(i, x))
                Next x

                'procedure icd9 utama & sekunder
                x = 0
                For x = 49 To (49 + 30) - 1
                    Call ValidasiINADRG("huruf", Len(.TextMatrix(i, x)), 10, .TextMatrix(i, x))
                Next x

                Call ValidasiINADRG("angka", Len(.TextMatrix(i, 79)), 17, .TextMatrix(i, 79)) 'grouper type
                Call ValidasiINADRG("angka", Len(.TextMatrix(i, 80)), 1, .TextMatrix(i, 80))  'patient type used
                Call ValidasiINADRG("angka", Len(.TextMatrix(i, 81)), 6, .TextMatrix(i, 81))  'DRG
                Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 82)), 2, .TextMatrix(i, 82))  'grouper status

                'diagnosis validity flag
                x = 0
                For x = 83 To (83 + 5) - 1
                    If .TextMatrix(i, x) = "" Then
                        Call ValidasiINADRG("huruf", 1, 1, " ")
                    Else
                        Call ValidasiINADRG("huruf", Len(.TextMatrix(i, x)), 1, .TextMatrix(i, x))
                    End If
                Next x

                x = 0
                For x = 1 To 25
                    Call ValidasiINADRG("huruf", 1, 1, " ")
                Next x

                'procedure validity flag
                x = 0
                For x = 88 To (88 + 5) - 1
                    If .TextMatrix(i, x) = "" Then
                        Call ValidasiINADRG("huruf", 1, 1, " ")
                    Else
                        Call ValidasiINADRG("huruf", Len(.TextMatrix(i, x)), 1, .TextMatrix(i, x))
                    End If
                Next x

                x = 0
                For x = 1 To 25
                    Call ValidasiINADRG("huruf", 1, 1, " ")
                Next x

                'Procedure Class
                x = 0
                For x = 93 To (93 + 5) - 1
                    If .TextMatrix(i, x) = "" Then
                        Call ValidasiINADRG("huruf", 1, 1, " ")
                    Else
                        Call ValidasiINADRG("huruf", Len(.TextMatrix(i, x)), 1, .TextMatrix(i, x))
                    End If
                Next x

                x = 0
                For x = 1 To 25
                    Call ValidasiINADRG("huruf", 1, 1, " ")
                Next x

                Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 98)), 4, .TextMatrix(i, 98)) 'Birth Weight Used
                Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 99)), 1, .TextMatrix(i, 99)) 'Birth Weight Source
                Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 100)), 1, .TextMatrix(i, 100)) 'Medical Surgical Flag

                'Procedure DRG
                x = 0
                For x = 101 To (101 + 5) - 1
                    If .TextMatrix(i, x) = "" Then
                        Call ValidasiINADRG("huruf", 6, 6, "      ")
                    ElseIf .TextMatrix(i, x) = "0" Then
                        Call ValidasiINADRG("huruf", 6, 6, "      ")
                    Else
                        Call ValidasiINADRG("huruf", Len(.TextMatrix(i, x)), 6, .TextMatrix(i, x))
                    End If
                Next x

                x = 0
                For x = 1 To 25
                    Call ValidasiINADRG("huruf", 6, 6, "      ")
                Next x

                'Procedure Statistic
                x = 0
                For x = 106 To (106 + 5) - 1
                    If .TextMatrix(i, x) = "" Then
                        Call ValidasiINADRG("huruf", 7, 7, " 0,0000")
                    Else
                        Call ValidasiINADRG("huruf", Len(" " & .TextMatrix(i, x)), 7, " " & .TextMatrix(i, x))
                    End If
                Next x

                x = 0
                For x = 1 To 25
                    Call ValidasiINADRG("huruf", 7, 7, " 0,0000")
                Next x

                If i <> .Rows - 1 Then
                    strTampung = strTampung & vbNewLine
                Else
                    strTampung = strTampung
                End If

                x = 0
                y = 0

            End If
        Next i
    End With
    Screen.MousePointer = vbDefault

    Call subSimpanData(strTampung, sFile)
    MsgBox "Data berhasil disimpan", vbInformation, "Berhasil.."
    Exit Sub
hell:
    Screen.MousePointer = vbDefault
    Call msubPesanError
    strTampung = ""
End Sub

Private Sub cmdLoad_Click()
    On Error GoTo hell
    cdLoad.FileName = ""
    cdLoad.ShowOpen
    strOpen = LoadText(cdLoad.FileName)
    If bolFile = False Then Exit Sub
    Call setgrid

    m_RecLen = HitJmlRec(strOpen)

    Call GetBarisData(strOpen)

    With fgLoad
        .Rows = m_RecLen + 1
        For i = 1 To m_RecLen
            .TextMatrix(i, 0) = Chr$(187)
            .TextMatrix(i, 1) = Trim(Mid(strArray(i), 1, 7))
            .TextMatrix(i, 2) = Trim(Mid(strArray(i), 8, 1))
            .TextMatrix(i, 3) = Trim(Mid(strArray(i), 14, 6))
            .TextMatrix(i, 4) = Trim(Mid(strArray(i), 20, 1))
            .TextMatrix(i, 5) = Val(Mid(strArray(i), 21, 10))
            .TextMatrix(i, 6) = IIf(Len(Trim(Mid(strArray(i), 31, 10))) = 0, "", Trim(Mid(strArray(i), 31, 10)))
            .TextMatrix(i, 7) = IIf(Len(Trim(Mid(strArray(i), 41, 24))) = 0, "", Trim(Mid(strArray(i), 41, 24)))
            .TextMatrix(i, 8) = Trim(Mid(strArray(i), 65, 1))
            .TextMatrix(i, 9) = Trim(Mid(strArray(i), 66, 10))
            .TextMatrix(i, 10) = Trim(Mid(strArray(i), 76, 10))
            .TextMatrix(i, 11) = Trim(Mid(strArray(i), 86, 4))
            .TextMatrix(i, 12) = Trim(Mid(strArray(i), 90, 10))
            .TextMatrix(i, 13) = Trim(Mid(strArray(i), 100, 3))
            .TextMatrix(i, 14) = Trim(Mid(strArray(i), 103, 3))
            .TextMatrix(i, 15) = IIf(Len(Trim(Mid(strArray(i), 106, 3))) = 0, "", Trim(Mid(strArray(i), 106, 3)))
            .TextMatrix(i, 16) = Trim(Mid(strArray(i), 109, 1))
            .TextMatrix(i, 17) = Trim(Mid(strArray(i), 110, 2))
            .TextMatrix(i, 18) = Trim(Mid(strArray(i), 112, 4))

            .TextMatrix(i, 19) = Trim(Mid(strArray(i), 116, 10)) '30x
            .TextMatrix(i, 20) = Trim(Mid(strArray(i), 126, 10)) '30x
            .TextMatrix(i, 21) = Trim(Mid(strArray(i), 136, 10)) '30x
            .TextMatrix(i, 22) = Trim(Mid(strArray(i), 146, 10)) '30x
            .TextMatrix(i, 23) = Trim(Mid(strArray(i), 156, 10)) '30x
            .TextMatrix(i, 24) = Trim(Mid(strArray(i), 166, 10)) '30x
            .TextMatrix(i, 25) = Trim(Mid(strArray(i), 176, 10)) '30x
            .TextMatrix(i, 26) = Trim(Mid(strArray(i), 186, 10)) '30x
            .TextMatrix(i, 27) = Trim(Mid(strArray(i), 196, 10)) '30x
            .TextMatrix(i, 28) = Trim(Mid(strArray(i), 206, 10)) '30x
            .TextMatrix(i, 29) = Trim(Mid(strArray(i), 216, 10)) '30x
            .TextMatrix(i, 30) = Trim(Mid(strArray(i), 226, 10)) '30x
            .TextMatrix(i, 31) = Trim(Mid(strArray(i), 236, 10)) '30x
            .TextMatrix(i, 32) = Trim(Mid(strArray(i), 246, 10)) '30x
            .TextMatrix(i, 33) = Trim(Mid(strArray(i), 256, 10)) '30x
            .TextMatrix(i, 34) = Trim(Mid(strArray(i), 266, 10)) '30x
            .TextMatrix(i, 35) = Trim(Mid(strArray(i), 276, 10)) '30x
            .TextMatrix(i, 36) = Trim(Mid(strArray(i), 286, 10)) '30x
            .TextMatrix(i, 37) = Trim(Mid(strArray(i), 296, 10)) '30x
            .TextMatrix(i, 38) = Trim(Mid(strArray(i), 306, 10)) '30x
            .TextMatrix(i, 39) = Trim(Mid(strArray(i), 316, 10)) '30x
            .TextMatrix(i, 40) = Trim(Mid(strArray(i), 326, 10)) '30x
            .TextMatrix(i, 41) = Trim(Mid(strArray(i), 336, 10)) '30x
            .TextMatrix(i, 42) = Trim(Mid(strArray(i), 346, 10)) '30x
            .TextMatrix(i, 43) = Trim(Mid(strArray(i), 356, 10)) '30x
            .TextMatrix(i, 44) = Trim(Mid(strArray(i), 366, 10)) '30x
            .TextMatrix(i, 45) = Trim(Mid(strArray(i), 376, 10)) '30x
            .TextMatrix(i, 46) = Trim(Mid(strArray(i), 386, 10)) '30x
            .TextMatrix(i, 47) = Trim(Mid(strArray(i), 396, 10)) '30x
            .TextMatrix(i, 48) = Trim(Mid(strArray(i), 406, 10)) '30x

            .TextMatrix(i, 49) = Trim(Mid(strArray(i), 416, 10)) '30x
            .TextMatrix(i, 50) = Trim(Mid(strArray(i), 426, 10)) '30x
            .TextMatrix(i, 51) = Trim(Mid(strArray(i), 436, 10)) '30x
            .TextMatrix(i, 52) = Trim(Mid(strArray(i), 446, 10)) '30x
            .TextMatrix(i, 53) = Trim(Mid(strArray(i), 456, 10)) '30x
            .TextMatrix(i, 54) = Trim(Mid(strArray(i), 466, 10)) '30x
            .TextMatrix(i, 55) = Trim(Mid(strArray(i), 476, 10)) '30x
            .TextMatrix(i, 56) = Trim(Mid(strArray(i), 486, 10)) '30x
            .TextMatrix(i, 57) = Trim(Mid(strArray(i), 496, 10)) '30x
            .TextMatrix(i, 58) = Trim(Mid(strArray(i), 506, 10)) '30x
            .TextMatrix(i, 59) = Trim(Mid(strArray(i), 516, 10)) '30x
            .TextMatrix(i, 60) = Trim(Mid(strArray(i), 526, 10)) '30x
            .TextMatrix(i, 61) = Trim(Mid(strArray(i), 536, 10)) '30x
            .TextMatrix(i, 62) = Trim(Mid(strArray(i), 546, 10)) '30x
            .TextMatrix(i, 63) = Trim(Mid(strArray(i), 556, 10)) '30x
            .TextMatrix(i, 64) = Trim(Mid(strArray(i), 566, 10)) '30x
            .TextMatrix(i, 65) = Trim(Mid(strArray(i), 576, 10)) '30x
            .TextMatrix(i, 66) = Trim(Mid(strArray(i), 586, 10)) '30x
            .TextMatrix(i, 67) = Trim(Mid(strArray(i), 596, 10)) '30x
            .TextMatrix(i, 68) = Trim(Mid(strArray(i), 606, 10)) '30x
            .TextMatrix(i, 69) = Trim(Mid(strArray(i), 616, 10)) '30x
            .TextMatrix(i, 70) = Trim(Mid(strArray(i), 626, 10)) '30x
            .TextMatrix(i, 71) = Trim(Mid(strArray(i), 636, 10)) '30x
            .TextMatrix(i, 72) = Trim(Mid(strArray(i), 646, 10)) '30x
            .TextMatrix(i, 73) = Trim(Mid(strArray(i), 656, 10)) '30x
            .TextMatrix(i, 74) = Trim(Mid(strArray(i), 666, 10)) '30x
            .TextMatrix(i, 75) = Trim(Mid(strArray(i), 676, 10)) '30x
            .TextMatrix(i, 76) = Trim(Mid(strArray(i), 686, 10)) '30x
            .TextMatrix(i, 77) = Trim(Mid(strArray(i), 696, 10)) '30x
            .TextMatrix(i, 78) = Trim(Mid(strArray(i), 706, 10)) '30x

            .TextMatrix(i, 79) = Trim(Mid(strArray(i), 716, 17))
            .TextMatrix(i, 80) = Trim(Mid(strArray(i), 733, 1))
            .TextMatrix(i, 81) = Trim(Mid(strArray(i), 734, 6))
            .TextMatrix(i, 82) = Trim(Mid(strArray(i), 740, 2))
            If .TextMatrix(i, 82) = "01" Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
            End If

            .TextMatrix(i, 83) = Trim(Mid(strArray(i), 742, 1)) '30x
            .TextMatrix(i, 84) = Trim(Mid(strArray(i), 743, 1)) '30x
            .TextMatrix(i, 85) = Trim(Mid(strArray(i), 744, 1)) '30x
            .TextMatrix(i, 86) = Trim(Mid(strArray(i), 745, 1)) '30x
            .TextMatrix(i, 87) = Trim(Mid(strArray(i), 746, 1)) '30x

            .TextMatrix(i, 88) = Trim(Mid(strArray(i), 772, 1)) '30x
            .TextMatrix(i, 89) = Trim(Mid(strArray(i), 773, 1)) '30x
            .TextMatrix(i, 90) = Trim(Mid(strArray(i), 774, 1)) '30x
            .TextMatrix(i, 91) = Trim(Mid(strArray(i), 775, 1)) '30x
            .TextMatrix(i, 92) = Trim(Mid(strArray(i), 776, 1)) '30x

            .TextMatrix(i, 93) = Trim(Mid(strArray(i), 802, 1)) '30x
            .TextMatrix(i, 94) = Trim(Mid(strArray(i), 803, 1)) '30x
            .TextMatrix(i, 95) = Trim(Mid(strArray(i), 804, 1)) '30x
            .TextMatrix(i, 96) = Trim(Mid(strArray(i), 805, 1)) '30x
            .TextMatrix(i, 97) = Trim(Mid(strArray(i), 806, 1)) '30x

            .TextMatrix(i, 98) = Trim(Mid(strArray(i), 832, 4))
            .TextMatrix(i, 99) = Trim(Mid(strArray(i), 836, 1))
            .TextMatrix(i, 100) = Trim(Mid(strArray(i), 837, 1))

            .TextMatrix(i, 101) = Val(Trim(Mid(strArray(i), 838, 6))) '30x
            .TextMatrix(i, 102) = Val(Trim(Mid(strArray(i), 844, 6))) '30x
            .TextMatrix(i, 103) = Val(Trim(Mid(strArray(i), 850, 6))) '30x
            .TextMatrix(i, 104) = Val(Trim(Mid(strArray(i), 856, 6))) '30x
            .TextMatrix(i, 105) = Val(Trim(Mid(strArray(i), 862, 6))) '30x

            .TextMatrix(i, 106) = Trim(Mid(strArray(i), 1018, 7)) '30x
            .TextMatrix(i, 107) = Trim(Mid(strArray(i), 1025, 7)) '30x
            .TextMatrix(i, 108) = Trim(Mid(strArray(i), 1032, 7)) '30x
            .TextMatrix(i, 109) = Trim(Mid(strArray(i), 1039, 7)) '30x
            .TextMatrix(i, 110) = Trim(Mid(strArray(i), 1046, 7)) '30x

            .TextMatrix(i, 111) = 1
        Next i
    End With
    cmdSave.Enabled = True
    cmdConvExcel.Enabled = True
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdBersih_Click()
    Set rs = Nothing
    Call msubRecFO(rs, "Select * from LoadHasilINADRG")
    If rs.EOF = True Then GoTo kosong
    If MsgBox("Yakin data temporary INADRG yang tersimpan didatabase akan dihapus?", vbExclamation + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    dbConn.Execute "Delete From LoadHasilINADRG"
kosong:
    fgLoad.clear
    Call setgrid
    cmdSave.Enabled = False
    cmdConvExcel.Enabled = False
End Sub

Private Sub fgLoad_Click()
    On Error GoTo hell

    If fgLoad.Rows = 1 Then Exit Sub
    If fgLoad.Col <> 0 Then Exit Sub

    chkCheck.Visible = True
    chkCheck.Top = fgLoad.RowPos(fgLoad.Row) + 1110 '1360
    Dim intChk As Integer
    intChk = ((fgLoad.ColPos(fgLoad.Col + 1) - fgLoad.ColPos(fgLoad.Col)) / 2)
    chkCheck.Left = fgLoad.ColPos(fgLoad.Col) + 30 + intChk
    chkCheck.SetFocus
    If fgLoad.Col = 0 Then
        If fgLoad.TextMatrix(fgLoad.Row, 0) <> "" Then
            chkCheck.value = 1
        Else
            chkCheck.value = 0
        End If
    End If

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call setgrid
    cmdSave.Enabled = False
    cmdConvExcel.Enabled = False
    Set rs = Nothing
    Call msubRecFO(rs, "Delete From LoadHasilINADRG")
    Set rs = Nothing
End Sub

Private Sub txtCariNoCM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With fgLoad
            .Row = 1
            .Col = 0

            For i = 1 To .Rows - 1
                If UCase(Left(txtCariNoCM.Text, Len(txtCariNoCM.Text))) = UCase(Left(fgLoad.TextMatrix(i, 3), Len(txtCariNoCM.Text))) Then Exit For
            Next i
            .TopRow = i: .Row = i: .Col = 3: .SetFocus
        End With
    End If
End Sub

