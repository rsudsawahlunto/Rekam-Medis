VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCekKepesertaanBPJS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Cek Kepesertaan BPJS"
   ClientHeight    =   7050
   ClientLeft      =   6195
   ClientTop       =   4245
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCekKepesertaanBPJS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   10140
   Begin VB.Frame Frame1 
      Caption         =   "Jenis Identitas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3855
      Begin VB.OptionButton optBPJS 
         Caption         =   "No Kartu BPJS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "NIK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgPeserta 
      Height          =   4695
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8281
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin VB.TextBox txtNoBPJS 
      Height          =   495
      Left            =   6840
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdValidasi 
      Caption         =   "&Proses"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtNoKartu 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   1320
      Width           =   4215
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Silakan Pilih Jenis Identitas"
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8280
      Picture         =   "frmCekKepesertaanBPJS.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmCekKepesertaanBPJS.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "frmCekKepesertaanBPJS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnKartuAktif As Boolean
Dim StatusVclaim As String
'untuk cek validasi
Private Function funcCekValidasi() As Boolean
 If optBPJS.value = True Then
    If Trim(txtNoKartu.Text) <> "" Then
        If Len(txtNoKartu.Text) <> 13 Then
            MsgBox "NoBPJS kurang harus diisi (13 digit)", vbExclamation, "Validasi"
            funcCekValidasi = False
            txtNoKartu.SetFocus
            Exit Function
        End If
    End If
 Else
     If Trim(txtNoKartu.Text) <> "" Then
        If Len(txtNoKartu.Text) <> 16 Then
            MsgBox "NoKartu kurang harus diisi (16 digit)", vbExclamation, "Validasi"
            funcCekValidasi = False
            txtNoKartu.SetFocus
            Exit Function
        End If
    End If
 End If
    funcCekValidasi = True
End Function

Private Sub cmdValidasi_Click()
    txtNoBPJS.Text = ""
    blnKartuAktif = False
    If funcCekValidasi = False Then Exit Sub
If StatusVclaim = "Y" Then
        If optBPJS.value = True Then
            Call ValidateKartuPeserta("Kartu BPJS")
        Else
            Call ValidateKartuPeserta("NIK")
    '        Call ValidateKartuPeserta("Kartu BPJS")
        End If
Else
    If optBPJS.value = True Then
             Call ValidateKartuPesertaV21("Kartu BPJS")
         Else
             Call ValidateKartuPesertaV21("NIK")
         End If
End If

    If blnKartuAktif = True Then
        MsgBox "Pasien tersebut aktif", vbInformation, "Cek Kepesertaan BPJS"
    End If
End Sub
Private Sub ValidateKartuPesertaV21(strJenisID As String)
On Error GoTo hell
    If (Dir("C:\SDK\askes\result.tlb") <> "") Then
        Dim context As Bpjs_V21.context
        Set context = New context
        Dim result() As String
        strSQL = "Select Value From SettingGlobal where Prefix In('ConsumerIDV21','PasswordKeyV21','KodeRSV21')"
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = False Then
            context.ConsumerId = rs(0).value
            'context.ConsumerID = "1001"
            rs.MoveNext
            context.KodeRumahSakit = rs(0).value
            rs.MoveNext
            context.PasswordKey = rs(0).value
        End If
        
        
         strSQL = "SELECT Value FROM SettingGlobal where Prefix='UrlGenerateSEPV21'"
            Call msubRecFO(rs, strSQL)
            Dim url  As String
              If rs.EOF = False Then
                  url = rs.Fields(0)
                  context.SetEndpointAskesLocal (url)
                 'Exit Sub
              End If
              
        'versi lama
        'result = context.CekDataPasienByNoKartu(txtNoKartuPA.Text)
        'versi baru (biasanya dibelakangnya ada 20
'        txtNamaPA.Text = ""
        If strJenisID = "Kartu BPJS" And optBPJS.value = True Then
            result = context.CekDataPasienByNoKartuV21(txtNoKartu.Text)
        ElseIf strJenisID = "Kartu BPJS" And optBPJS.value = False Then
            result = context.CekDataPasienByNoKartuV21(txtNoBPJS.Text)
        ElseIf strJenisID = "NIK" Then
            result = context.CekDataPasienByNik21(txtNoKartu.Text)
        End If
        Dim vloop As Integer
        'vloop = 0
        'For i = LBound(result) To UBound(result) - 1
        '    Debug.Print result(0).value
        'Next i
    End If
        'Call fillGridWithRiwayatPasien(fgPeserta, result)
        Call fillGridWithRiwayatPasienByRow(fgPeserta, result)
        Dim I As Long
        For I = LBound(result) To UBound(result)
            Debug.Print (result(I))
            Dim arr() As String
            arr = Split(result(I), ":")
            Select Case arr(0)
                    Case "noKartu"
                        blnKartuAktif = True
                        txtNoBPJS.Text = arr(1)
                        noKartu = arr(1)
                        Debug.Print "NoKartu : " & arr(1)
            End Select
        Next I
        
        If txtNoBPJS.Text = "" Or UBound(result) = -1 Then
            blnKartuAktif = False
            MsgBox "Data Peserta Tidak Ditemukan,,,,!!!" & vbCrLf & Replace(result(0), "message:", ""), vbInformation, "Validasi"
'            Debug.Print txtNoKartuPA.Text
            Debug.Print result(0)
            'txtNoKartuPA.Text = ""
            'Set DgPasien2.DataSource = Nothing
            'Call SubloadPasienBPJS
            Exit Sub
        End If
        'Call SubloadPasienBPJS
        'DgPasien2.SetFocus
    'Else
    '    MsgBox "Sdk Bridging askes tidak di temukan"
   ' End If
Exit Sub
hell:
MsgBox "Koneksi Bridging Bermasalah"
End Sub

Private Sub ValidateKartuPeserta(strJenisID As String)
On Error GoTo hell
    If (Dir("C:\SDK\Vclaim\result.tlb") <> "") Then
        Dim context As ContextVclaim
        Set context = New ContextVclaim
        Dim result() As String
        strSQL = "Select Value From SettingGlobal where Prefix In('ConsumerID','PasswordKey','KodeRS')"
        Call msubRecFO(rs, strSQL)
        
        If rs.EOF = False Then
            context.ConsumerId = rs(0).value
            'context.ConsumerID = "1001"
            rs.MoveNext
'            context.KodeRumahSakit = rs(0).value
            rs.MoveNext
            context.PasswordKey = rs(0).value
        End If
        
        
         strSQL = "SELECT Value FROM SettingGlobal where Prefix='UrlGenerateSEP'"
            Call msubRecFO(rs, strSQL)
            Dim url  As String
              If rs.EOF = False Then
                  url = rs.Fields(0)
'                  context.SetEndpointAskesLocal (url)
                  context.url = url
                 'Exit Sub
              End If
              
        'versi lama
        'result = context.CekDataPasienByNoKartu(txtNoKartuPA.Text)
        'versi baru (biasanya dibelakangnya ada 20
'        txtNamaPA.Text = ""
        If strJenisID = "Kartu BPJS" And optBPJS.value = True Then
            result = context.CariPesertaByNoKartuBpjs(txtNoKartu.Text, Format(Now, "yyyy-mm-dd"))
        ElseIf strJenisID = "Kartu BPJS" And optBPJS.value = False Then
            result = context.CariPesertaByNoKartuBpjs(txtNoBPJS.Text, Format(Now, "yyyy-mm-dd"))
        ElseIf strJenisID = "NIK" Then
            result = context.CariPesertaByNik(txtNoKartu.Text, Format(Now, "yyyy-mm-dd"))
        End If
        Dim vloop As Integer
        'vloop = 0
        'For i = LBound(result) To UBound(result) - 1
        '    Debug.Print result(0).value
        'Next i
    End If
        'Call fillGridWithRiwayatPasien(fgPeserta, result)
        Call fillGridWithRiwayatPasienByRow(fgPeserta, result)
        Dim I As Long
        For I = LBound(result) To UBound(result)
            Debug.Print (result(I))
            Dim arr() As String
            arr = Split(result(I), ":")
            Select Case arr(0)
                    Case "noKartu"
                        blnKartuAktif = True
                        txtNoBPJS.Text = arr(1)
                        noKartu = arr(1)
                        Debug.Print "NoKartu : " & arr(1)
            End Select
        Next I
        
        If txtNoBPJS.Text = "" Or UBound(result) = -1 Then
            blnKartuAktif = False
            MsgBox "Data Peserta Tidak Ditemukan,,,,!!!" & vbCrLf & Replace(result(0), "message:", ""), vbInformation, "Validasi"
'            Debug.Print txtNoKartuPA.Text
            Debug.Print result(0)
            'txtNoKartuPA.Text = ""
            'Set DgPasien2.DataSource = Nothing
            'Call SubloadPasienBPJS
            Exit Sub
        End If
        'Call SubloadPasienBPJS
        'DgPasien2.SetFocus
    'Else
    '    MsgBox "Sdk Bridging askes tidak di temukan"
   ' End If
Exit Sub
hell:
MsgBox "Koneksi Bridging Bermasalah"
End Sub

Private Sub fgPeserta_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intCtrlShift As Integer
    intCtrlShift = vbCtrlMask + Shift
    Select Case KeyCode
        Case vbKeyC
            If intCtrlShift = 4 Then
                Clipboard.Clear
                Clipboard.SetText fgPeserta.Clip
            End If
    End Select

End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    txtNoBPJS.Text = ""
    blnKartuAktif = False
    optBPJS.value = True
    
        strSQL = "Select Value From SettingGlobal where Prefix ='StatusVclaim'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            StatusVclaim = rs(0).value
        End If
        
End Sub

Sub fillGridWithRiwayatPasien(vFG As MSFlexGrid, vResult() As String)
    With vFG
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .cols = 1
        .rows = 2
'        Call subSetGrid
        Dim row As Integer
        Dim col As Integer
        Dim rows As Integer
        Dim cols As Integer
        Dim I As Integer
        row = 1
        For I = 0 To UBound(vResult)
            Dim arrResult() As String
'            Debug.Assert i <> 9
            arrResult = Split(vResult(I), ":")
            col = isHeaderExist(vFG, arrResult(0))
            If col > -1 Then
                'col = isHeaderExist(vFG, arrResult(0))
                
                If .TextMatrix(row, col) <> "" Then 'KALO TEXTMATRIX TARGET SUDAH ADA ISI BERARTI KITA HARUS TAMBAH
                                                    'ROWS/PINDAH KE BARIS SELANJUTNYA
                                                    'diharapkan kolom pertama adalah kolom yang selalu memiliki nilai
                    .rows = .rows + 1
                    row = .rows - 1
                End If
                .TextMatrix(row, col) = arrResult(1)
            Else 'KALAU ADA KOLOM BARU
                col = .cols - 1
                .TextMatrix(0, col) = arrResult(0) 'BERI HEADER BARU
                .TextMatrix(row, col) = arrResult(1)
                .cols = .cols + 1
                
            End If
        Next I
        .cols = .cols - 1 'MENGHILANGKAN KOLOM YG KELEBIHAN
    End With
End Sub

Function isHeaderExist(vFG As MSFlexGrid, strHeader As String) As Integer
    isHeaderExist = -1
    With vFG
        Dim col As Integer
        For col = 0 To vFG.cols - 1
            If UCase(.TextMatrix(0, col)) = UCase(strHeader) Then
                isHeaderExist = col
                Exit Function
            End If
        Next col
    End With
End Function


Sub fillGridWithRiwayatPasienByRow(vFG As MSFlexGrid, vResult() As String)
    With vFG
        .Clear
        .Redraw = False
        .cols = 3
        .rows = UBound(vResult) + 2
        .FixedRows = 1
        .FixedCols = 1
        .ColWidth(0) = 300
        .RowHeight(0) = 300
          
        .ColWidth(1) = 2025
        .ColWidth(2) = 6555
'        Call subSetGrid
        Dim row As Integer
        Dim col As Integer
        Dim rows As Integer
        Dim cols As Integer
        Dim I As Integer
        row = 1
        For I = 0 To UBound(vResult)
            Dim arrResult() As String
'            Debug.Assert i <> 9
            arrResult = Split(vResult(I), ":")
            If vResult(I) = "=============" Then GoTo lewati
            If vResult(I) = ">>cob" Then GoTo lewati
            If vResult(I) = ">>hakKelas" Then GoTo lewati
            If vResult(I) = ">>informasi" Then GoTo lewati
            If vResult(I) = ">>jenisPeserta" Then GoTo lewati
            If vResult(I) = ">>mr" Then GoTo lewati
            If vResult(I) = ">>provUmum" Then GoTo lewati
            If vResult(I) = ">>statusPeserta" Then GoTo lewati
            If vResult(I) = ">>umur" Then GoTo lewati


            .TextMatrix(I + 1, 1) = arrResult(0)
            .TextMatrix(I + 1, 2) = arrResult(1)
lewati:
        Next I
        .ColAlignment(1) = flexAlignLeftTop
        .ColAlignment(2) = flexAlignLeftTop
        .Redraw = True
    End With
End Sub

Private Sub txtNoBPJS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdValidasi.SetFocus
End If
End Sub
