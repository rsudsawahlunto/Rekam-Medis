VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmConvertINADRG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Convert INA DRG"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConvertINADRG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   15090
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
      Left            =   6000
      TabIndex        =   7
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CheckBox chkCheck 
      Height          =   210
      Left            =   240
      TabIndex        =   16
      Top             =   2195
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.TextBox txtCariNoCM 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   7440
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   120
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Simpan"
   End
   Begin VB.TextBox txtPath 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   8040
      Width           =   14175
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "&Convert"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9480
      TabIndex        =   8
      Top             =   7200
      Width           =   2655
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   615
      Left            =   12360
      TabIndex        =   9
      Top             =   7200
      Width           =   2655
   End
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
      Left            =   10980
      TabIndex        =   11
      Top             =   960
      Width           =   4035
      Begin VB.CommandButton cmdcari 
         Caption         =   "&Cari"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
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
         Format          =   127074307
         UpDown          =   -1  'True
         CurrentDate     =   37956
      End
   End
   Begin MSDataListLib.DataCombo dcInstalasi 
      Height          =   330
      Left            =   120
      TabIndex        =   5
      Top             =   7440
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   5055
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   8916
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
   Begin MSDataListLib.DataCombo dcJenisPasien 
      Height          =   330
      Left            =   2640
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcPenjamin 
      Height          =   330
      Left            =   5160
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   19
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmConvertINADRG.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label Label6 
      Caption         =   "Penjamin"
      Height          =   255
      Left            =   5160
      TabIndex        =   18
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Jenis Pasien"
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Instalasi"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Cari NoCM"
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label lblJumData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data 0/0"
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   12
      Top             =   8130
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   13230
      Picture         =   "frmConvertINADRG.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmConvertINADRG.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmConvertINADRG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsVal As New ADODB.recordset
Dim strVal As String
Dim sfilter As String

Private Sub chkCheck_Click()
    On Error GoTo errLoad

    If chkCheck.value = vbChecked Then
        fgData.TextMatrix(fgData.Row, fgData.Col) = Chr$(187)
        fgData.TextMatrix(fgData.Row, 20) = 1
    Else
        fgData.TextMatrix(fgData.Row, fgData.Col) = ""
        fgData.TextMatrix(fgData.Row, 20) = 0
    End If

    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub chkCheck_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then fgData.SetFocus
End Sub

Private Sub chkCheck_LostFocus()
    chkCheck.Visible = False
End Sub

Private Sub chkCheckSemua_Click()
    On Error GoTo hell
    Dim i As Integer

    If chkCheckSemua.value = Checked Then
        For i = 1 To fgData.Rows - 1
            fgData.TextMatrix(i, 0) = Chr$(187)
            fgData.TextMatrix(i, 20) = 1
        Next i
    Else
        For i = 1 To fgData.Rows - 1
            fgData.TextMatrix(i, 0) = ""
            fgData.TextMatrix(i, 20) = 0
        Next i
    End If

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Function sp_AdminINADRGsp(adoCommand As ADODB.Command, varIdPegawai As String, varkdKelompokPasien As String, varIdPenjamin As String, varPilihInstalasi As String) As Boolean
    On Error GoTo errSimpan
    Set adoCommand = New ADODB.Command
    sp_AdminINADRGsp = False
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglPOS", adDate, adParamInput, , Format(Now, "yyyy/MM/dd 00:00:00"))
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, varIdPegawai)
        .Parameters.Append .CreateParameter("TglKeluarAwal", adDate, adParamInput, , Format(dtpAwal.value, "yyyy/MM/dd 00:00:00"))
        .Parameters.Append .CreateParameter("TglKeluarAkhir", adDate, adParamInput, , Format(dtpAwal.value, "yyyy/MM/dd 23:59:59"))
        .Parameters.Append .CreateParameter("kdKelompokPasien", adChar, adParamInput, 2, dcJenisPasien.BoundText)
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, dcPenjamin.BoundText)
        .Parameters.Append .CreateParameter("PilihInstalasi", adChar, adParamInput, 2, dcInstalasi.BoundText)

        .ActiveConnection = dbConn
        .CommandText = "AdminINADRGsp"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data", vbCritical, "Validasi"
        Else
            sp_AdminINADRGsp = True
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Function
errSimpan:
    Call deleteADOCommandParameters(adoCommand)
    Set adoCommand = Nothing
    Call msubPesanError
End Function

Private Sub cmdCari_Click()
    On Error GoTo hell
    Dim i As Integer
    Dim j As Integer

    If dcJenisPasien.MatchedWithList = False Then
        MsgBox "Jenis Pasien Harus diisi", vbOKOnly + vbCritical
        dcJenisPasien.SetFocus
        Exit Sub
    End If

    If dcPenjamin.MatchedWithList = False Then
        MsgBox "Penjamin Pasien Harus diisi", vbOKOnly + vbCritical
        dcPenjamin.SetFocus
        Exit Sub
    End If

    If sp_AdminINADRGsp(dbcmd, strIDPegawai, dcJenisPasien.BoundText, dcPenjamin.BoundText, dcInstalasi.BoundText) = False Then Exit Sub

    If dcInstalasi.BoundText = "02" Then
        strSQL = " select NoPendaftaran, NoPendaftaranGD, KdRS AS HospitalID, KelasRS AS HospitalType, NoCM AS MedicalRecordNumber," & _
        " KdKelas AS PatientClass, TotalBiaya AS Payment, Filler1," & _
        " KdInstalasi AS PatientTypeDesignation, '0' AS BirthWeightInGrams, TglMasuk AS AdminDate," & _
        " TglKeluar AS DischargeDate, LamaDirawat AS ALOS, TglLahir AS BirthDate, Thn AS AgeInYears," & _
        " Hari AS AgeInDays, Filler2, JK AS Sex, KdStatusKeluar AS [DischargeDisposition/Status]," & _
        " Diagnosa AS [PrincipalDiagnosis&Dagnoses], Prosedur as [Procedure]" & _
        " FROM    V_INADRGJEP2009 where kdInstalasi = '02' and TglKeluar BETWEEN '" & Format(dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAwal.value, "yyyy/MM/dd 23:59:59") & "'" & _
        " and NoCM like '" & txtCariNoCM.Text & "%' and kdKelompokPasien Like '" & dcJenisPasien.BoundText & "%' and IdPenjamin Like '" & dcPenjamin.BoundText & "%' "
    Else
        strSQL = " select NoPendaftaran, NoPendaftaranGD, KdRS AS HospitalID, KelasRS AS HospitalType, NoCM AS MedicalRecordNumber," & _
        " KdKelas AS PatientClass, TotalBiaya AS Payment, Filler1," & _
        " KdInstalasi AS PatientTypeDesignation, '0' AS BirthWeightInGrams, TglMasuk AS AdminDate," & _
        " TglKeluar AS DischargeDate, LamaDirawat AS ALOS, TglLahir AS BirthDate, Thn AS AgeInYears," & _
        " Hari AS AgeInDays, Filler2, JK AS Sex, KdStatusKeluar AS [DischargeDisposition/Status]," & _
        " Diagnosa AS [PrincipalDiagnosis&Dagnoses], Prosedur as [Procedure]" & _
        " FROM    V_INADRGJEP2009 where kdInstalasi = '03' and TglKeluar BETWEEN '" & Format(dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAwal.value, "yyyy/MM/dd 23:59:59") & "'" & _
        " and NoCM like '" & txtCariNoCM.Text & "%' and kdKelompokPasien Like '" & dcJenisPasien.BoundText & "%' and IdPenjamin Like '" & dcPenjamin.BoundText & "%' "
    End If

    Set rs = Nothing
    MousePointer = vbHourglass
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    MousePointer = vbDefault
    Call setgrid
    If rs.EOF Then Exit Sub

    lblJumData.Caption = "Data : " & rs.RecordCount
    cmdConvert.Enabled = True

    With fgData
        .Rows = rs.RecordCount + 1
        For i = 1 To rs.RecordCount
            .TextMatrix(i, 0) = "" 'Chr$(187)
            .TextMatrix(i, 1) = rs("HospitalID")
            If IsNull(rs("HospitalType")) Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 2) = ""
            Else
                .TextMatrix(i, 2) = rs("HospitalType")
            End If
            If IsNull(rs("MedicalRecordNumber")) Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 3) = ""
            Else
                .TextMatrix(i, 3) = rs("MedicalRecordNumber")
            End If
            If IsNull(rs("PatientClass")) Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 4) = ""
            Else
                .TextMatrix(i, 4) = rs("PatientClass")
            End If
            If IsNull(rs("Payment")) Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 5) = ""
            ElseIf rs("Payment") = 0 Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 5) = Val(rs("Payment"))
            Else
                .TextMatrix(i, 5) = Val(rs("Payment"))
            End If
            .TextMatrix(i, 6) = rs("Filler1")
            If IsNull(rs("PatientTypeDesignation")) Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 7) = ""
            Else
                .TextMatrix(i, 7) = rs("PatientTypeDesignation")
            End If
            If IsNull(rs("AdminDate")) Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 8) = ""
            Else
                .TextMatrix(i, 8) = Format(rs("AdminDate"), "dd/MM/yyyy HH:mm:ss")
            End If
            If IsNull(rs("DischargeDate")) Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 9) = ""
            Else
                .TextMatrix(i, 9) = Format(rs("DischargeDate"), "dd/MM/yyyy HH:mm:ss")
            End If
            If IsNull(rs("ALOS")) Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 10) = ""
            Else
                .TextMatrix(i, 10) = rs("ALOS")
            End If
            If IsNull(rs("BirthDate")) Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 11) = ""
            Else
                .TextMatrix(i, 11) = rs("BirthDate")
            End If
            If IsNull(rs("AgeInYears")) Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 12) = ""
            Else
                .TextMatrix(i, 12) = rs("AgeInYears")
            End If
            If IsNull(rs("AgeInDays")) Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 13) = ""
            Else
                .TextMatrix(i, 13) = rs("AgeInDays")
            End If
            .TextMatrix(i, 14) = rs("Filler2")
            If IsNull(rs("Sex")) Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 15) = ""
            Else
                .TextMatrix(i, 15) = rs("Sex")
            End If
            If IsNull(rs("DischargeDisposition/Status")) Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 16) = ""
            Else
                .TextMatrix(i, 16) = rs("DischargeDisposition/Status")
            End If
            If IsNull(rs("BirthWeightInGrams")) Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 17) = ""
            Else
                .TextMatrix(i, 17) = rs("BirthWeightInGrams")
            End If
            If IsNull(rs("PrincipalDiagnosis&Dagnoses")) Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 18) = ""
            Else
                .TextMatrix(i, 18) = rs("PrincipalDiagnosis&Dagnoses")
            End If
            If IsNull(rs("Procedure")) Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
                .TextMatrix(i, 19) = ""
            Else
                .TextMatrix(i, 19) = rs("Procedure")
            End If
            strVal = ""
            strVal = "Select KdDiagnosa From PeriksaDiagnosa Where (Nopendaftaran='" & rs("NoPendaftaran") & "') AND KdJenisDiagnosa='05'"
            Set rsVal = Nothing
            rsVal.Open strVal, dbConn, adOpenForwardOnly, adLockReadOnly
            If rsVal.EOF = True Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
            Else
                If rsVal.RecordCount > 1 Then
                    .Row = i
                    .Col = 1
                    .CellForeColor = vbRed
                End If
            End If

            strVal = ""
            strVal = "Select KdDiagnosaTindakan From DetailPeriksaDiagnosa Where (Nopendaftaran in ('" & rs("NoPendaftaran") & "','" & rs("NoPendaftaranGD") & "')) AND KdJenisDiagnosa='05'"
            Set rsVal = Nothing
            rsVal.Open strVal, dbConn, adOpenForwardOnly, adLockReadOnly
            If rsVal.EOF = True Then
                .Row = i
                .Col = 1
                .CellForeColor = vbRed
            Else
                If rsVal.RecordCount > 1 Then
                    .Row = i
                    .Col = 1
                    .CellForeColor = vbRed
                End If
            End If

            .TextMatrix(i, 20) = 0

            rs.MoveNext
        Next i
    End With
    Exit Sub
hell:
    MousePointer = vbDefault
    Call msubPesanError
End Sub

Private Sub setgrid()
    Dim i As Integer
    With fgData
        .clear
        .Rows = 2
        .Cols = 21

        'judul
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Hospital ID"
        .TextMatrix(0, 2) = "Hospital Type"
        .TextMatrix(0, 3) = "Medical Record Number"
        .TextMatrix(0, 4) = "Patient Class"
        .TextMatrix(0, 5) = "Payment"
        .TextMatrix(0, 6) = "Filler1"
        .TextMatrix(0, 7) = "Patient Type Designation"
        .TextMatrix(0, 8) = "Admin Date"
        .TextMatrix(0, 9) = "Discharge Date"
        .TextMatrix(0, 10) = "ALOS"
        .TextMatrix(0, 11) = "Birth Date"
        .TextMatrix(0, 12) = "Age In Years"
        .TextMatrix(0, 13) = "Age In Days"
        .TextMatrix(0, 14) = "Filler2"
        .TextMatrix(0, 15) = "Sex"
        .TextMatrix(0, 16) = "Discharge Disposition/Status"
        .TextMatrix(0, 17) = "Birth Weight In Grams"
        .TextMatrix(0, 18) = "Principal Diagnosis & Diagnoses"
        .TextMatrix(0, 19) = "Procedure"

        .ColWidth(0) = 340
        .ColWidth(1) = 900
        .ColWidth(2) = 1100
        .ColWidth(3) = 1800
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 600
        .ColWidth(7) = 1000
        .ColWidth(8) = 1650
        .ColWidth(9) = 1650
        .ColWidth(10) = 600
        .ColWidth(11) = 1000
        .ColWidth(12) = 1000
        .ColWidth(13) = 1000
        .ColWidth(14) = 600
        .ColWidth(15) = 400
        .ColWidth(16) = 2120
        .ColWidth(17) = 1800
        .ColWidth(18) = 4000
        .ColWidth(19) = 4000
        .ColWidth(20) = 0

        .ColAlignment(0) = flexAlignCenterCenter

    End With
End Sub

Private Sub cmdcari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkCheck.Visible = False
        Call chkCheck_Click
        fgData.SetFocus
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdConvert_Click()
    On Error GoTo hell
    Dim sFile As String
    Dim strJK As String
    Dim strICD10 As String
    Dim strICD9 As String
    Dim x, y, z, k As Integer
    Dim jmlJarak As Integer
awal:
    cdFile.FileName = ""
    strTampung = ""
    sFile = ""
    cdFile.ShowSave
    If cdFile.FileName = "" Then Exit Sub
    strICD10 = ""
    strICD9 = ""

    sFile = cdFile.FileName '& ".txt"
    txtPath.Text = sFile
    Screen.MousePointer = vbHourglass

    With fgData
        For i = 1 To fgData.Rows - 1
            If .TextMatrix(i, 20) = 1 Then
                Call ValidasiINADRG("angka", Len(.TextMatrix(i, 1)), 7, .TextMatrix(i, 1)) 'hospital ID
                Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 2)), 1, .TextMatrix(i, 2))  'hospital type
                Call ValidasiINADRG("angka", Len(.TextMatrix(i, 3)), 11, .TextMatrix(i, 3)) 'medical record number
                If .TextMatrix(i, 4) = "01" Then 'patient class
                    Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 4)), 1, "3")
                ElseIf .TextMatrix(i, 4) = "02" Then
                    Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 4)), 1, "2")
                Else
                    Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 4)), 1, "1")
                End If
                Call ValidasiINADRG("angka", Len(.TextMatrix(i, 5)), 9, Val(.TextMatrix(i, 5)))  'payment
                Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 6)), 35, .TextMatrix(i, 6)) 'filler1
                If .TextMatrix(i, 7) = "03" Then  'patient type designation
                    Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 7)), 1, "1") 'RI
                Else
                    Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 7)), 1, "2") 'RJ, IGD, dll
                End If
                Call ValidasiINADRG("angka", Len(Format(.TextMatrix(i, 8), "dd/MM/yyyy")), 10, Format(.TextMatrix(i, 8), "dd/MM/yyyy")) 'admin date
                Call ValidasiINADRG("angka", Len(Format(.TextMatrix(i, 9), "dd/MM/yyyy")), 10, Format(.TextMatrix(i, 9), "dd/MM/yyyy")) 'discharge date
                Call ValidasiINADRG("angka", Len(.TextMatrix(i, 10)), 4, .TextMatrix(i, 10)) 'ALOS
                Call ValidasiINADRG("angka", Len(Format(.TextMatrix(i, 11), "dd/MM/yyyy")), 10, Format(.TextMatrix(i, 11), "dd/MM/yyyy")) 'birth date
                'validasi umur jika thn ada, maka hr = 0 dan sebaliknya
                If .TextMatrix(i, 12) <> "0" Then
                    Call ValidasiINADRG("angka", Len(.TextMatrix(i, 12)), 3, .TextMatrix(i, 12))    'umur thn
                    Call ValidasiINADRG("angka", 1, 3, "0")                   'umur hr
                Else
                    Call ValidasiINADRG("angka", 1, 3, "0")                   'umur thn
                    Call ValidasiINADRG("angka", Len(.TextMatrix(i, 13)), 3, .TextMatrix(i, 13))    'umur hr
                End If

                Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 14)), 3, .TextMatrix(i, 14)) 'filler2
                If .TextMatrix(i, 15) = "L" Then    'sex
                    'strJK = "1"
                    Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 15)), 1, "1")
                ElseIf .TextMatrix(i, 15) = "P" Then
                    'strJK = "2"
                    Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 15)), 1, "2")
                Else
                    'strJK = "0"
                    Call ValidasiINADRG("huruf", Len(.TextMatrix(i, 15)), 1, "0")
                End If
                
                If .TextMatrix(i, 7) = "01" Then 'IGD
                    Select Case .TextMatrix(i, 16)
                        Case "01"
                            Call ValidasiINADRG("angka", 1, 2, "2")
                        Case "02"
                            Call ValidasiINADRG("angka", 1, 2, "1")
                        Case "03"
                            Call ValidasiINADRG("angka", 1, 2, "4")
                        Case "04"
                            Call ValidasiINADRG("angka", 1, 2, "4")
                        Case "05"
                            Call ValidasiINADRG("angka", 1, 2, "3")
                        Case "06"
                            Call ValidasiINADRG("angka", 1, 2, "2")
                    End Select
                ElseIf .TextMatrix(i, 7) = "03" Then 'RI
                    Select Case .TextMatrix(i, 16)
                        Case "01"
                            Call ValidasiINADRG("angka", 1, 2, "2")
                        Case "02"
                            Call ValidasiINADRG("angka", 1, 2, "1")
                        Case "03"
                            Call ValidasiINADRG("angka", 1, 2, "4")
                        Case "04"
                            Call ValidasiINADRG("angka", 1, 2, "4")
                        Case "05"
                            Call ValidasiINADRG("angka", 1, 2, "3")
                        Case "06"
                            Call ValidasiINADRG("angka", 1, 2, "2")
                    End Select
                Else 'RJ dan lain2
                    Call ValidasiINADRG("angka", 1, 2, "1")
                End If
                Call ValidasiINADRG("angka", Len(.TextMatrix(i, 17)), 4, .TextMatrix(i, 17)) 'birth weight in gram

                'principal diagnosis icd10 utama & diagnoses icd10 sekunder
                x = 0: y = 0: jmlJarak = 0
                For x = 1 To Len(.TextMatrix(i, 18))
                    If Mid(.TextMatrix(i, 18), x, 1) = ";" Then
                        jmlJarak = jmlJarak + 10
                        For y = Len(strICD10) To jmlJarak - 1
                            strICD10 = strICD10 & " "
                        Next y
                        x = x + 1
                    Else
                        If (Mid(CStr(.TextMatrix(i, 18)), x, 1) <> " ") Then
                            If (Mid(CStr(.TextMatrix(i, 18)), x, 1) <> ".") Then
                                strICD10 = strICD10 + Mid(.TextMatrix(i, 18), x, 1)
                            End If
                        End If
                    End If
                Next x
                Call ValidasiINADRG("huruf", Len(strICD10), 300, strICD10)

                'procedure icd9 utama & sekunder
                x = 0: y = 0: jmlJarak = 0
                For x = 1 To Len(.TextMatrix(i, 19))
                    If Mid(.TextMatrix(i, 19), x, 1) = ";" Then
                        jmlJarak = jmlJarak + 10
                        For y = Len(strICD9) To jmlJarak - 1
                            strICD9 = strICD9 & " "
                        Next y
                        x = x + 1
                    Else
                        If (Mid(CStr(.TextMatrix(i, 19)), x, 1) <> " ") Then
                            If (Mid(CStr(.TextMatrix(i, 19)), x, 1) <> ".") Then
                                strICD9 = strICD9 + Mid(.TextMatrix(i, 19), x, 1)
                            End If
                        End If
                    End If
                Next x
                Call ValidasiINADRG("huruf", Len(strICD9), 300, strICD9)

                If i <> .Rows - 1 Then
                    strTampung = strTampung & vbNewLine
                Else
                    strTampung = strTampung
                End If

                strICD10 = ""
                strICD9 = ""
                x = 0
                y = 0
                z = 0
                jmlJarak = 0

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

Private Sub dcInstalasi_Change()
    If dcInstalasi.BoundText = "03" Then
        sfilter = " AND KdInstalasi = '" & dcInstalasi.BoundText & "'"
    ElseIf dcInstalasi.BoundText = "02" Then
        sfilter = " AND KdInstalasi <> '03'"
    Else
        sfilter = ""
    End If
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcInstalasi.MatchedWithList = True Then dcJenisPasien.SetFocus
        strSQL = "Select KdInstalasi,NamaInstalasi From Instalasi Where KdInstalasi IN('02','03') and StatusEnabled='1' and (Namainstalasi LIKE '%" & dcInstalasi.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcInstalasi.Text = ""
            Exit Sub
        End If
        dcInstalasi.BoundText = rs(0).value
        dcInstalasi.Text = rs(1).value
    End If
End Sub

Private Sub dcJenisPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcJenisPasien.MatchedWithList = True Then dcPenjamin.SetFocus
        strSQL = "Select KdKelompokPasien,JenisPasien From KelompokPasien where kdKelompokPasien = '21' and StatusEnabled='1' and (JenisPasien LIKE '%" & dcJenisPasien.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcJenisPasien.Text = ""
            Exit Sub
        End If
        dcJenisPasien.BoundText = rs(0).value
        dcJenisPasien.Text = rs(1).value
    End If
End Sub

Private Sub dcJenisPasien_LostFocus()
    On Error GoTo hell

    Set dbRst = Nothing
    Call msubDcSource(dcPenjamin, dbRst, "select  distinct a.idpenjamin, b.namapenjamin from PenjaminKelompokPasien a " & _
    " inner join Penjamin b on a.idpenjamin = b.idpenjamin " & _
    " inner join KelompokPasien c on a.kdkelompokpasien = c.kdkelompokpasien " & _
    " where   a.kdkelompokpasien like '%" & dcJenisPasien.BoundText & "%' and b.StatusEnabled='1'" & _
    " order by b.namapenjamin ")

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcPenjamin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Set rs = Nothing
        rs.Open "Select IdPenjamin,NamaPenjamin From Penjamin Where NamaPenjamin Like '" & dcPenjamin.Text & "%' and StatusEnabled='1'", dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.EOF = True Then
            dcPenjamin.Text = ""
            Exit Sub
        End If
        Set dcPenjamin.RowSource = rs
        dcPenjamin.BoundText = rs(0).value
        dcPenjamin.Text = rs(1).value
        cmdCari.SetFocus
    End If
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub fgData_Click()
    On Error GoTo hell

    If fgData.Rows = 1 Then Exit Sub
    If fgData.Col <> 0 Then Exit Sub

    chkCheck.Visible = True
    chkCheck.Top = fgData.RowPos(fgData.Row) + 1955 '2195
    Dim intChk As Integer
    intChk = ((fgData.ColPos(fgData.Col + 1) - fgData.ColPos(fgData.Col)) / 2)
    chkCheck.Left = fgData.ColPos(fgData.Col) + 30 + intChk
    chkCheck.SetFocus
    If fgData.Col = 0 Then
        If fgData.TextMatrix(fgData.Row, 0) <> "" Then
            chkCheck.value = 1
        Else
            chkCheck.value = 0
        End If
    End If

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub fgData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then Call fgData_Click
End Sub

Private Sub Form_Load()
    On Error GoTo hell
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)

    dtpAwal.value = Format(Now, "dd MMMM yyyy 00:00:00")
    Call LoadDataCombo

    Call setgrid

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub LoadDataCombo()
    'instalasi
    Set rs = Nothing
    rs.Open "Select KdInstalasi,NamaInstalasi From Instalasi Where KdInstalasi IN('02','03') and StatusEnabled='1'", dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcInstalasi.RowSource = rs
    dcInstalasi.BoundColumn = rs(0).Name
    dcInstalasi.ListField = rs(1).Name
    dcInstalasi.BoundText = "02"

    'jenis pasien
    Set rs = Nothing
    rs.Open "Select KdKelompokPasien,JenisPasien From KelompokPasien where kdKelompokPasien = '21' and StatusEnabled='1'", dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcJenisPasien.RowSource = rs
    dcJenisPasien.BoundColumn = rs(0).Name
    dcJenisPasien.ListField = rs(1).Name

    'penjamin
    Set rs = Nothing
    rs.Open "Select a.IdPenjamin,b.NamaPenjamin From PenjaminKelompokPasien a inner join Penjamin b on a.idPenjamin=b.idPenjamin inner join kelompokpasien c on a.kdKelompokPasien=c.kdKelompokPasien where a.kdKelompokPasien = '21' and b.StatusEnabled='1'", dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcPenjamin.RowSource = rs
    dcPenjamin.BoundColumn = rs(0).Name
    dcPenjamin.ListField = rs(1).Name
End Sub

Private Sub txtCariNoCM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With fgData
            .Row = 1
            .Col = 0

            For i = 1 To .Rows - 1
                If UCase(Left(txtCariNoCM.Text, Len(txtCariNoCM.Text))) = UCase(Left(fgData.TextMatrix(i, 3), Len(txtCariNoCM.Text))) Then Exit For
            Next i
            .TopRow = i: .Row = i: .Col = 3: .SetFocus
        End With
    End If
End Sub

