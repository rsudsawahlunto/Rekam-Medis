Attribute VB_Name = "modUtility"
Public Sub msubDcSource(dcName As Object, rsName As ADODB.recordset, strName As String)
    On Error GoTo errLoad

    Set rsName = New ADODB.recordset
    Set rsName = dbConn.Execute(strName)

    Set dcName.RowSource = rsName
    dcName.BoundColumn = rsName(0).Name
    dcName.ListField = rsName(1).Name

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Function Periksa(vCase As String, vObj As Object, vPesan As String) As Boolean
    Periksa = True

    Select Case LCase(vCase)
        Case "text"
            If Len(Trim(vObj.Text)) > 0 Then Exit Function
        Case "datacombo"
            If vObj.MatchedWithList = True Then Exit Function
        Case "nilai"
            If Val(vObj) > 0 Then Exit Function
        Case "combobox"
            If vObj.ListIndex >= 0 Then Exit Function
    End Select

    MsgBox vPesan, vbExclamation, "Validasi"
    vObj.SetFocus
    Periksa = False
End Function

Public Function CekItemGrid(objGridDropName As Object, objGridArrayName As Object, intColDrop As Integer, intColArray As Integer, objGridName As Object) As Boolean
    CekItemGrid = False

    For i = objGridArrayName.LowerBound(1) To objGridArrayName.UpperBound(1)
        If IsNull(objGridDropName.SelectedItem) Then
            objGridName.ReBind
            Exit Function
        End If
        If LCase(objGridDropName.Columns(intColDrop)) = LCase(objGridArrayName(i, intColArray)) Then
            objGridName.ReBind
            objGridName.SetFocus
            objGridName.row = i
            Exit Function
        End If
    Next

    CekItemGrid = True
End Function

Public Sub msubPesanError(Optional s_Sub As String)
    Dim sQError As String
    Dim sErrNumber As String
    Dim sErrDesc As String
    Dim sError As String, sErrorHasil As String
    Dim iLoop As Integer

    MsgBox "Ada kesalahan dalam loading data" & vbNewLine _
    & "Hubungi administrator dan laporkan pesan berikut" & vbNewLine _
    & Err.Number & " - " & Err.Description, vbCritical, "Validasi" & s_Sub

    sErrDesc = Err.Description
    iLoop = 1
    Do While iLoop <= Len(sErrDesc)
        If Mid(sErrDesc, iLoop, 1) = "'" Then
            sError = "-"
        Else
            sError = Mid(sErrDesc, iLoop, 1)
        End If

        sErrorHasil = sErrorHasil & sError
        iLoop = iLoop + 1
    Loop

    sErrDesc = ""
    sErrDesc = sErrorHasil

    On Error Resume Next

    sQError = "insert into ErrorMessage values('" & Format(Now, "yyyy/MM/dd hh:MM:ss") & "', '" & strIDPegawaiAktif & "','" & strKdAplikasi & "','" & sErrDesc & "','" & strNamaHostLocal & "','" & s_Sub & "', '' )"
    dbConn.Execute sQError
    
    Resume 0
End Sub

Public Sub msubRecFO(recordset As ADODB.recordset, Query As String)
    On Error GoTo errLoad

    Set recordset = New ADODB.recordset
    recordset.Open Query, dbConn, adOpenForwardOnly, adLockReadOnly

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Public Sub centerForm(ByRef oForm1 As Form, ByVal oForm2 As Form)
    oForm1.Left = (oForm2.Width - oForm1.Width) / 2
    oForm1.Top = (oForm2.Height - 1500 - oForm1.Height) / 2
End Sub

Public Sub openConnection()
    On Error Resume Next
    blnError = False
    With dbConn
        If .State = adStateOpen Then Exit Sub
        .CursorLocation = adUseClient

        If strSQLIdentifikasi = "1" Then
            .ConnectionString = "Provider=SQLNCLI10.1;Password=" & strPassword & ";DataTypeCompatibility=80;Persist Security Info=True;User ID=" & strUserName & ";Initial Catalog=" & strDatabaseName & ";Data Source=" & strServerName
            .Open
            If dbConn.State = adStateOpen Then Exit Sub

            .ConnectionString = "Provider=SQLNCLI10;Password=" & strPassword & ";DataTypeCompatibility=80;Persist Security Info=True;User ID=" & strUserName & ";Initial Catalog=" & strDatabaseName & ";Data Source=" & strServerName
            .Open
            If dbConn.State = adStateOpen Then Exit Sub

            On Error GoTo NoConn
            .ConnectionString = "Provider=SQLNCLI.1;Password=" & strPassword & ";DataTypeCompatibility=80;Persist Security Info=True;User ID=" & strUserName & ";Initial Catalog=" & strDatabaseName & ";Data Source=" & strServerName
            .Open
            If dbConn.State = adStateOpen Then Exit Sub
        Else
            .ConnectionString = "Provider=SQLNCLI10.1;Integrated Security=SSPI;DataTypeCompatibility=80;Persist Security Info=False;Initial Catalog=" & strDatabaseName & ";Data Source=" & strServerName
            .Open
            If dbConn.State = adStateOpen Then Exit Sub

            .ConnectionString = "Provider=SQLNCLI10;Integrated Security=SSPI;DataTypeCompatibility=80;Persist Security Info=False;Initial Catalog=" & strDatabaseName & ";Data Source=" & strServerName
            .Open
            If dbConn.State = adStateOpen Then Exit Sub

            On Error GoTo NoConn
            .ConnectionString = "Provider=SQLNCLI.1;Integrated Security=SSPI;DataTypeCompatibility=80;Persist Security Info=False;Initial Catalog=" & strDatabaseName & ";Data Source=" & strServerName
            .Open
            If dbConn.State = adStateOpen Then Exit Sub
        End If

        If dbConn.State = adStateOpen Then
        Else
            MsgBox "Koneksi ke database error, hubungi administrator !" & vbCrLf & Err.Description & " (" & Err.Number & ")"
        End If
    End With
    Exit Sub
NoConn:
    MsgBox "Koneksi ke database error, ganti nama Server dan nama Database", vbCritical, "Validasi"
    frmSetServer.Show
    blnError = True
    Unload frmLogin
End Sub

Public Sub hitungUmur(ByVal tgllahir As String)
    Dim thnl As Integer, thns As Integer, thn As Integer
    Dim blnl As Integer, blns As Integer, bln As Integer
    Dim haril As Integer, haris As Integer, hari As Integer

    '********************************************************************************
    'perhitungan tahun, bulan dan hari
    '********************************************************************************
    If tgllahir = "__/__/____" Then
        Exit Sub
    End If
    thnl = Year(tgllahir)
    thns = Year(Now)
    blnl = Month(tgllahir)
    blns = Month(Now)
    haril = Day(tgllahir)
    haris = Day(Now)

    thn = thns - thnl
    '?tgllahir = februari & kabisat (29 hari)
    If (blnl = 2) And (haril = 29) Then
        '?bulan sekarang = februari & kabisat (29 hari)
        If (blns = 2) And (haris = 29) Then
            hari = 0
            bln = 0
        Else
            'bulan sekarang = februari & hari ke 28
            If (blns = 2) And (haris = 28) Then
                hari = 28
                bln = 11
            Else
                'tgl lahir <> 28 atau 29
                haril = 1
                blnl = 3
                If (blnl > blns) Then
                    bln = (12 - blnl) + blns
                    thn = thn - 1
                Else
                    bln = blns - blnl
                End If
                If (haris = haril) Then
                    hari = 0
                Else
                    hari = haris - haril
                End If
            End If
        End If
    Else
        If (blnl > blns) Then
            bln = (12 - blnl) + blns
            thn = thn - 1
        Else
            If (blns = blnl) And (haril > haris) Then
                bln = 11
                thn = thn - 1
            Else
                bln = blns - blnl
            End If
        End If

        If (haris < haril) Then
            hari = (totalHari(blnl, thns) - haril) + haris
            If (bln <> 0) And (blns <> blnl) And (haril > haris) Then
                bln = bln - 1
            End If
        Else
            hari = haris - haril
        End If
    End If

    If (thn = 1) And (blns <> 0) And ((blns <> blnl) And (haril < haris)) Then
        thn = 0
    End If

    Umur.tahun = thn
    Umur.bulan = bln
    Umur.hari = hari
End Sub

Function totalHari(ByVal bln As Integer, ByVal tahun As Integer) As Integer
    'bulan yg 31 hari : januari(1), maret, mei, juli, agustus, oktober, desember
    If (bln = 1) Or (bln = 3) Or (bln = 5) Or (bln = 7) Or (bln = 8) Or (bln = 10) Or (bln = 12) Then
        TtlHari = 31
    Else
        'bulan = 2, 4, 6, 8, 10, 12
        If (bln = 2) Then
            '?tahun = tahun kabisat
            If (tahun Mod 4) = 0 Then
                TtlHari = 29
            Else
                TtlHari = 28
            End If
        Else
            TtlHari = 30
        End If
    End If
End Function

Public Sub deleteADOCommandParameters(ByRef adoCommand As ADODB.Command)
    Dim prmcounter As Integer
    With adoCommand
        For prmcounter = 0 To .Parameters.Count - 1
            .Parameters.Delete (0)
        Next
    End With
End Sub

Public Sub GetIdPegawai()
    openConnection

    petugas = frmLogin.txtUserID.Text

    Set dbcmd = New Command
    With dbcmd
        .ActiveConnection = dbConn
        .CommandText = "SELECT IdPegawai FROM Login WHERE(UserName = '" & petugas & "')"
        .CommandType = adCmdText
    End With

    Set dbRst = New recordset
    Set dbRst = dbcmd.Execute

    While Not dbRst.EOF
        noidpegawai = dbRst.Fields("idpegawai").value
        dbRst.MoveNext
    Wend
    dbRst.Close
    Set dbRst = Nothing
End Sub

Public Sub subEnterSetFocus(KeyAscii As Integer, objObject As Object)
    If KeyAscii = 13 Then objObject.SetFocus
End Sub

Sub RecOpen(vSql As String)
    On Error GoTo errLoad
    Set DbRec = New ADODB.recordset
    Set DbRec = dbConn.Execute(vSql)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub MasterDcSource(vSql As String, vObj As DataCombo, vBound As String, vAlias As String)
    RecOpen vSql
    Set vObj.RowSource = DbRec
    vObj.BoundColumn = vBound
    vObj.ListField = vAlias
End Sub

Function PeriksaHapus(vCase As String, vObj As Object, vSql As String) As Boolean
    PeriksaHapus = False

    Select Case LCase(vCase)
        Case "text"
            If Len(Trim(vObj.Text)) < 1 Then Exit Function
            RecOpen vSql
            If DbRec.RecordCount < 1 Then
                MsgBox "Maaf data tidak bisa dihapus karena sudah di pakai", vbInformation
                Exit Function
            End If

    End Select

    PeriksaHapus = True
End Function

Sub RefreshGrid(vCari As Object)
    Dim vTemp As String

    vTemp = vCari
    vCari = "!"
    vCari = vTemp
End Sub

'---------------------------------------------------------------------------------------
'for counting how old a person is (year old,month old, day old) based on their birthdate
'---------------------------------------------------------------------------------------
Public Sub subYearOldCount(dTglLahir As String)
    dTglLahir = CDate(dTglLahir)
    Dim intYeaNow As Integer
    Dim intMonNow As Integer
    Dim intDayNow As Integer
    Dim intYeaBirth As Integer
    Dim intMonBirth As Integer
    Dim intDayBirth As Integer
    Dim intDayInMonth As Integer
    intYeaNow = Year(Now)
    intMonNow = Month(Now)
    intDayNow = Day(Now)
    intYeaBirth = Year(dTglLahir)
    intMonBirth = Month(dTglLahir)
    intDayBirth = Day(dTglLahir)

    If intDayBirth > intDayNow Then
        intMonNow = intMonNow - 1
        If intMonNow = 0 Then intMonNow = 12: intYeaNow = intYeaNow - 1
        intDayNow = intDayNow + funcHitungHari(intMonNow, intYeaNow)
'        intDayNow = intDayNow + funcHitungHari(intMonNow, intYeaBirth)
        YOC_intDay = intDayNow - intDayBirth
    Else
        YOC_intDay = intDayNow - intDayBirth
    End If

    If intMonBirth > intMonNow Then
        intYeaNow = intYeaNow - 1
        intMonNow = intMonNow + 12
        YOC_intMonth = intMonNow - intMonBirth
    Else
        YOC_intMonth = intMonNow - intMonBirth
    End If
    YOC_intYear = intYeaNow - intYeaBirth
End Sub

'------------------------------------------------
'For counting how many days in the spesific month
'------------------------------------------------
Public Function funcHitungHari(intMonNow As Integer, intYeaNow As Integer) As Integer
    Select Case intMonNow
        Case 1
            funcHitungHari = 31
        Case 2
            If intYeaNow Mod 4 = 0 Then
                funcHitungHari = 29
            Else
                funcHitungHari = 28
            End If
        Case 3
            funcHitungHari = 31
        Case 4
            funcHitungHari = 30
        Case 5
            funcHitungHari = 31
        Case 6
            funcHitungHari = 30
        Case 7
            funcHitungHari = 31
        Case 8
            funcHitungHari = 31
        Case 9
            funcHitungHari = 30
        Case 10
            funcHitungHari = 31
        Case 11
            funcHitungHari = 30
        Case 12
            funcHitungHari = 31
    End Select
End Function

Public Function funcCekValidasiTgl(mstrChoose As String, mobjObject As Object) As String
    Select Case mstrChoose
        Case "TglLahir"
            On Error GoTo errTglLahir:
            If mobjObject.Text = "__/__/____" Then
                funcCekValidasiTgl = "ErrEmpty"
                Exit Function
            End If
            If Mid(mobjObject.Text, 4, 2) > 12 Then
                MsgBox "Bulan tanggal lahir tidak boleh melebihi bulan 12!", vbExclamation, "Validasi"
                mobjObject.SelStart = 3
                mobjObject.SelLength = 2
                mobjObject.SetFocus
                funcCekValidasiTgl = "ErrTahun"
                Exit Function
            End If
            If CInt(Mid(mobjObject.Text, 1, 2)) > funcHitungHari(CInt(Mid(mobjObject.Text, 4, 2)), CInt(Mid(mobjObject.Text, 7, 4))) Then
                MsgBox "Hari tanggal lahir yang dimasukkan lebih besar dari jumlah hari dari bulan tanggal lahir yang dimasukkan!", vbExclamation, "Validasi"
                mobjObject.SelStart = 0
                mobjObject.SelLength = 2
                mobjObject.SetFocus
                funcCekValidasiTgl = "ErrTahun"
                Exit Function
            End If
            If Year(mobjObject.Text) > Year(Now) Then
                MsgBox "Tahun tanggal lahir tidak boleh melebihi tahun sekarang!", vbExclamation, "Validasi"
                mobjObject.SelStart = 6
                mobjObject.SelLength = 4
                mobjObject.SetFocus
                funcCekValidasiTgl = "ErrTahun"
                Exit Function
            End If
            If Format(mobjObject.Text, "yyyy/MM/dd") > Format(Now, "yyyy/MM/dd") Then
                MsgBox "Tanggal lahir tidak boleh melebihi tanggal sekarang!", vbExclamation, "Validasi"
                mobjObject.SelStart = 6
                mobjObject.SelLength = 4
                mobjObject.SetFocus
                funcCekValidasiTgl = "ErrTahun"
                Exit Function
            End If
            funcCekValidasiTgl = "NoErr"
            Exit Function
errTglLahir:
            MsgBox "Format Data Tanggal Salah, Format yang Benar Adalah : dd/mm/yyyy", vbExclamation, "Validasi"
            mobjObject.SelStart = 0
            mobjObject.SelLength = Len(mobjObject.Text)
            mobjObject.SetFocus
            funcCekValidasiTgl = "ErrFormat"
            Exit Function
        Case "YgLain"
    End Select
    funcCekValidasiTgl = "NoErr"
End Function

Public Sub SetKeyPressToNumber(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = 13 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Public Sub SetKeyPressToChar(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = 13 Then Exit Sub
    If KeyAscii >= 48 And KeyAscii <= 57 Then KeyAscii = 0
End Sub

Public Function Crypt(strData As String)
    Dim i As Integer
    Dim lokasi As Integer
    Code = "1234567890"
    Crypt = ""
    For i% = 1 To Len(strData)
        lokasi% = (i% Mod Len(Code)) + 1
        Crypt = Crypt + Chr$(Asc(Mid$(strData, i%, 1)) Xor _
        Asc(Mid$(Code, lokasi%, 1)))
    Next i%
End Function

Public Sub settingreport(ByRef namareport As Report, ByVal NamaPrinter As String, ByVal namadriver As String, _
    ByVal ukurankertas As String, ByVal jenisduplexing As String, _
    ByVal OrientasKertas As String)

    With namareport
        If OrientasKertas = "" Then
            .PaperOrientation = crDefaultPaperOrientation
        Else
            .PaperOrientation = OrientasKertas
        End If
        If ukurankertas <> "" Then
            .PaperSize = ukurankertas
        Else
            .PaperSize = crDefaultPaperSize
        End If
        If jenisduplexing <> "" Then
            .PrinterDuplex = jenisduplexing
        Else
            .PrinterDuplex = crPRDPDefault
        End If
    End With
End Sub

Public Property Get OrienKertas() As String
    OrienKertas = tmpOrien
End Property

Public Property Let OrienKertas(ByVal vNewValue As String)
    tmpOrien = vNewValue
End Property

Function NumToText(dblValue As Double) As String
    Static ones(0 To 9) As String
    Static teens(0 To 9) As String
    Static tens(0 To 9) As String
    Static thousands(0 To 4) As String
    Dim i As Integer, nPosition As Integer
    Dim nDigit As Integer, bAllZeros As Integer
    Dim strResult As String, strTemp As String
    Dim tmpBuff As String
    'Maksimum 15 Digit
    'Untuk yg 7 digit atau lebih dengan angka 1 pada ribuan dan
    '2 digit sebelumnya merupakan angka nol, maka satu ribu diganti seribu
    'Untuk yang 4 Digit dan diawali angka 1, harus dikonversi
    'Satu+spasi diganti jadi Se
    ones(0) = "nol"
    ones(1) = "satu"
    ones(2) = "dua"
    ones(3) = "tiga"
    ones(4) = "empat"
    ones(5) = "lima"
    ones(6) = "enam"
    ones(7) = "tujuh"
    ones(8) = "delapan"
    ones(9) = "sembilan"

    teens(0) = "sepuluh"
    teens(1) = "sebelas"
    teens(2) = "duabelas"
    teens(3) = "tigabelas"
    teens(4) = "empatbelas"
    teens(5) = "limabelas"
    teens(6) = "enambelas"
    teens(7) = "tujuhbelas"
    teens(8) = "delapanbelas"
    teens(9) = "sembilanbelas"

    tens(0) = ""
    tens(1) = "sepuluh"
    tens(2) = "duapuluh"
    tens(3) = "tigapuluh"
    tens(4) = "empatpuluh"
    tens(5) = "limapuluh"
    tens(6) = "enampuluh"
    tens(7) = "tujuhpuluh"
    tens(8) = "delapanpuluh"
    tens(9) = "sembilanpuluh"

    thousands(0) = ""
    thousands(1) = "ribu"
    thousands(2) = "juta"
    thousands(3) = "miliar"
    thousands(4) = "triliun"

    'Errors Handler
    On Error GoTo NumToTextError
    'Bagian akhir
    strResult = "rupiah "
    'Konversi ke string
    Dim des, j, t1, t2
    Dim ada As Boolean
    For j = 1 To Len(totalbiaya)
        des = Mid(totalbiaya, j, 1)
        If des = "." Or des = "," Then
            ada = True
            t1 = Mid(totalbiaya, 1, j - 1)
            t2 = Mid(totalbiaya, j + 1)
            j = Len(totalbiaya)
        End If
    Next j
    If ada = True Then
        strTemp = CStr(Int(t1))
        ada = False
    Else
        strTemp = CStr(Int(dblValue))
    End If
    'Diulang sebanyak panjang teks
    For i = Len(strTemp) To 1 Step -1
        'Ambil nilai angka posisi ke-i
        nDigit = Val(Mid$(strTemp, i, 1))
        'Ambil posisi angka
        nPosition = (Len(strTemp) - i) + 1
        'Pilihan proses tergantung posisi satuan, puluhan, atau ratusan
        Select Case (nPosition Mod 3)
            Case 1  'Posisi satuan
                bAllZeros = False
                If i = 1 Then
                    tmpBuff = ones(nDigit) & " "
                ElseIf Mid$(strTemp, i - 1, 1) = "1" Then
                    tmpBuff = teens(nDigit) & " "
                    i = i - 1   'Skip posisi puluhan
                ElseIf nDigit > 0 Then
                    tmpBuff = ones(nDigit) & " "
                Else
                    'Jika angka Puluhan dan Ratusan juga
                    'angka nol, maka Jangan tampilkan ribuan
                    bAllZeros = True
                    If i > 1 Then
                        If Mid$(strTemp, i - 1, 1) <> "0" Then
                            bAllZeros = False
                        End If
                    End If
                    If i > 2 Then
                        If Mid$(strTemp, i - 2, 1) <> "0" Then
                            bAllZeros = False
                        End If
                    End If
                    tmpBuff = ""
                End If
                If bAllZeros = False And nPosition > 1 Then
                    tmpBuff = tmpBuff & thousands(nPosition / 3) & " "
                End If
                strResult = tmpBuff & strResult
            Case 2  'Posisi Puluhan
                If nDigit > 0 Then
                    strResult = tens(nDigit) & " " & strResult
                End If
            Case 0  'Posisi Ratusan
                If nDigit > 0 Then
                    If nDigit = 1 Then
                        strResult = "seratus " & strResult
                    Else
                        strResult = ones(nDigit) & " ratus " & strResult
                    End If
                End If
        End Select
    Next i
    'Konversi huruf pertama ke upper case
    If Len(strResult) > 0 Then
        strResult = UCase$(Left$(strResult, 1)) & Mid$(strResult, 2)
    End If

EndNumToText:
    'Mengembalikan Teks hasil
    NumToText = strResult
    Exit Function

NumToTextError:
    strResult = "#Error#"
    Resume EndNumToText
End Function

Public Sub msubRemoveItem(hgrid As Object, intRow As Integer)
    Dim i As Integer
    Dim j As Integer
    With hgrid
        Dim intRowNow As Integer
        intRowNow = intRow
        For i = 1 To .rows - 2
            If i = intRow Then
                For j = 0 To .cols - 1
                    .TextMatrix(intRow, j) = .TextMatrix(intRow + 1, j)
                Next j
                intRow = intRow + 1
            End If
        Next i
        .rows = .rows - 1
        .row = intRowNow
    End With
End Sub

Public Function ValidasiTanggal(f_TglAwal As Object, f_TglAkhir As Object) As Boolean
    Dim tanggalAwal As Integer
    Dim tanggalAkhir As Integer
    Dim bulanAwal As Integer
    Dim bulanAkhir As Integer
    Dim tahunAwal As Integer
    Dim tahunAkhir As Integer
    Dim pilihBulan As String

    tanggalAwal = CStr(Format(mdTglAwal, "dd"))
    tanggalAkhir = CStr(Format(mdTglAkhir, "dd"))
    bulanAwal = CStr(Format(mdTglAwal, "mm"))
    bulanAkhir = CStr(Format(mdTglAkhir, "mm"))
    tahunAwal = CStr(Format(mdTglAwal, "yyyy"))
    tahunAkhir = CStr(Format(mdTglAkhir, "yyyy"))
    'Tahun salah
    If tahunAkhir < tahunAwal Then
        MsgBox "Tahun Awal tidak boleh lebih besar dari tahun akhir", vbExclamation, "Validasi"
        ValidasiTanggal = False
        f_TglAwal.SetFocus
    End If
    'Tahun Benar Beda
    If tahunAkhir > tahunAwal Then
        ' Bulan Benar Sama
        If bulanAkhir = bulanAwal Then
            'Tanggal Salah
            If tanggalAkhir < tanggalAwal Then
                MsgBox "ada Kesalahan dalam Set tanggal", vbExclamation, "Validasi"
                ValidasiTanggal = False
                f_TglAwal.SetFocus
            Else
                'Tanggal Benar
                ValidasiTanggal = True
            End If
            ' Bulan Benar beda
        Else
            ValidasiTanggal = True
        End If
        'Tahun benar Sama
    ElseIf tahunAwal = tahunAkhir Then
        'Bulan Salah
        If bulanAkhir < bulanAwal Then
            MsgBox "Ada Kesalahan dalam Set Bulan", vbExclamation, "Validasi"
            ValidasiTanggal = False
            f_TglAwal.SetFocus
            'Bulan Benar Beda
        ElseIf bulanAkhir > bulanAwal Then
            ValidasiTanggal = True
            'Bulan Benar Sama
        ElseIf bulanAkhir = bulanAwal Then
            'Tanggal salah
            If tanggalAkhir < tanggalAwal Then
                MsgBox "ada Kesalahan dalam Set tanggal", vbExclamation, "Validasi"
                ValidasiTanggal = False
                f_TglAwal.SetFocus
            Else
                'tanggal Benar
                ValidasiTanggal = True
            End If
        End If
    End If
End Function

Public Function msubKonversiKomaTitik(f_Nilai As String)
    Dim tempKomaTitik  As String
    Dim j As Integer

    tempKomaTitik = f_Nilai
    For j = 1 To Len(tempKomaTitik)
        If Mid(tempKomaTitik, j, 1) = "," Then
            Mid(tempKomaTitik, j, 1) = "."
        End If
    Next j
    msubKonversiKomaTitik = tempKomaTitik
End Function

Public Function funcCekFormatTanggal(mstrChoose As String, mobjObject As Object) As String
    Select Case mstrChoose
        Case "TglLahir"
            On Error GoTo errTglLahir:
            If mobjObject.Text = "__/__/____" Then
                funcCekFormatTanggal = "ErrEmpty"
                Exit Function
            End If
            funcCekFormatTanggal = "NoErr"
            Exit Function
errTglLahir:
            MsgBox "Format Data Tanggal Salah, Format yang Benar Adalah : dd/mm/yyyy", vbCritical, "Informasi"
            mobjObject.SelStart = 0
            mobjObject.SelLength = Len(mobjObject.Text)
            mobjObject.SetFocus
            funcCekFormatTanggal = "ErrFormat"
            Exit Function
        Case "YgLain"
    End Select
    funcCekFormatTanggal = "NoErr"
End Function

Public Sub msubOpenRecFO(recordset As ADODB.recordset, strString As String, connection As ADODB.connection)
    On Error GoTo errLoad
    Set recordset = New ADODB.recordset
    recordset.Open strString, connection, adOpenForwardOnly, adLockReadOnly
    Exit Sub
errLoad:
    MsgBox "Ada kesalahan dalam loading recordset, laporkan kepada administrator pesan kesalahan berikut" & vbNewLine _
    & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
End Sub

Public Function funcRoundUp(strNumber As String) As Double
    If InStr(strNumber, ".") = 0 And InStr(strNumber, ",") = 0 Then
        funcRoundUp = CDbl(strNumber)
        Exit Function
    End If
    If InStr(strNumber, ".") <> 0 Then
        funcRoundUp = CDbl(Left(strNumber, InStr(strNumber, ".") - 2) & "0") + CDbl(CDbl(Mid(strNumber, (InStr(strNumber, ".") - 1), 1)) + 1)
        Exit Function
    End If
    If InStr(strNumber, ",") <> 0 Then
        funcRoundUp = CDbl(Left(strNumber, InStr(strNumber, ",") - 2) & "0") + CDbl(CDbl(Mid(strNumber, (InStr(strNumber, ",") - 1), 1)) + 1)
        Exit Function
    End If
End Function

Public Function funcRoundDown(strNumber As String) As Double
    If InStr(strNumber, ".") = 0 And InStr(strNumber, ",") = 0 Then
        funcRoundDown = CDbl(strNumber)
        Exit Function
    End If
    If InStr(strNumber, ".") <> 0 Then
        funcRoundDown = CDbl(Left(strNumber, InStr(strNumber, ".") - 1))
        Exit Function
    End If
    If InStr(strNumber, ",") <> 0 Then
        funcRoundDown = CDbl(Left(strNumber, InStr(strNumber, ",") - 1))
        Exit Function
    End If
End Function

Public Function funcRound(strNumber As String, dblPer As Double) As Double
    Dim dblNumber As Double
    Dim dblPerMinOne As Double
    Dim strNum As String
    Dim intLenNum As Integer
    Dim strZero As String
    Dim strNumLeft As String
    Dim dblNumLeft As Double
    Dim intNumSum As Integer
    dblNumber = funcRoundUp(strNumber)
    strNum = CStr(dblNumber)
    intLenNum = Len(CStr(dblPer - 1))
    If (CDbl(Right(strNum, intLenNum)) Mod dblPer) < dblPer And CDbl(Right(strNum, intLenNum)) <> 0 Then
        For i = 1 To intLenNum
            strZero = strZero & "0"
        Next i
        If (Len(strNum) - intLenNum) < 0 Then
            funcRound = dblPer
            Exit Function
        End If
        strNumLeft = Left(strNum, Len(strNum) - intLenNum) & strZero
        dblNumLeft = CDbl(strNumLeft)
        funcRound = (funcRoundUp(CDbl(strNum) / dblPer) * dblPer)
        Exit Function
    End If
    funcRound = dblNumber
End Function

Public Sub msubLoadDataArray(dTglAwal As Date, dTglAkhir As Date)
    Select Case mstrGrafik
        Case "JenisPasienPerJP"
            strSQL = "SELECT JenisPasien FROM V_RekapitulasiPasienBJenis GROUP BY JenisPasien"
            Call msubRecFO(rs, strSQL)
            mintJmlBarisGrafik = rs.RecordCount
            mintJmlKolomGrafik = 3
            ReDim arrGrafik(1 To mintJmlBarisGrafik, 1 To 4)
            i = 0
            While rs.EOF = False
                i = i + 1
                'isi kolom pertama dengan jenis pasien
                arrGrafik(i, 1) = rs(0).value
                rs.MoveNext
            Wend
            ReDim JnsKriteria(1 To 3) ' criteria
            JnsKriteria(1) = "Pasien Pria"
            JnsKriteria(2) = "Pasien Wanita"
            JnsKriteria(3) = "Total"
            ' isi array data untuk grafik
            strSQL = "SELECT JenisPasien,SUM(JmlPasienPria) AS JmlPasienPria," _
            & "SUM(JmlPasienWanita) AS JmlPasienWanita,SUM(Total) AS Total " _
            & "FROM V_RekapitulasiPasienBJenis " _
            & "WHERE dbo.S_AmbilTanggal(TglPendaftaran) BETWEEN '" _
            & Format(dTglAwal, "yyyy/MM/dd") _
            & "' AND '" & Format(dTglAkhir, "yyyy/MM/dd") _
            & "' AND (KdRuangan = '" & mstrKdRuangan & "') GROUP BY JenisPasien"
            Call msubRecFO(rs, strSQL)
            i = 0
            While rs.EOF = False
                i = i + 1
                'Cek apakah jenis kriteria ke - j sama dengan recordset _
                , bila tidak set jumlahnya menjadi 0
                If arrGrafik(i, 1) = rs("JenisPasien").value Then
                    arrGrafik(i, 2) = rs("JmlPasienPria").value
                    arrGrafik(i, 3) = rs("JmlPasienWanita").value
                    arrGrafik(i, 4) = rs("Total").value
                    rs.MoveNext
                Else
                    arrGrafik(i, 2) = 0
                    arrGrafik(i, 3) = 0
                    arrGrafik(i, 4) = 0
                End If
            Wend
        Case "StatusPasienPerSP"
            strSQL = "SELECT StatusPasien FROM V_RekapitulasiPasienBStatus GROUP BY StatusPasien"
            Call msubRecFO(rs, strSQL)
            mintJmlBarisGrafik = rs.RecordCount
            mintJmlKolomGrafik = 3
            ReDim arrGrafik(1 To mintJmlBarisGrafik, 1 To 4)
            i = 0
            While rs.EOF = False
                i = i + 1
                'isi kolom pertama dengan jenis pasien
                arrGrafik(i, 1) = rs(0).value
                rs.MoveNext
            Wend
            ReDim JnsKriteria(1 To 3) ' criteria
            JnsKriteria(1) = "Pasien Pria"
            JnsKriteria(2) = "Pasien Wanita"
            JnsKriteria(3) = "Total"
            ' isi array data untuk grafik
            strSQL = "SELECT StatusPasien,SUM(JmlPasienPria) AS JmlPasienPria," _
            & "SUM(JmlPasienWanita) AS JmlPasienWanita,SUM(Total) AS Total " _
            & "FROM V_RekapitulasiPasienBStatus " _
            & "WHERE dbo.S_AmbilTanggal(TglPendaftaran) BETWEEN '" _
            & Format(dTglAwal, "yyyy/MM/dd") _
            & "' AND '" & Format(dTglAkhir, "yyyy/MM/dd") _
            & "' AND (KdRuangan = '" & mstrKdRuangan & "') GROUP BY StatusPasien"
            Call msubRecFO(rs, strSQL)
            i = 0
            While rs.EOF = False
                i = i + 1
                'Cek apakah jenis kriteria ke - j sama dengan recordset _
                , bila tidak set jumlahnya menjadi 0
                If arrGrafik(i, 1) = rs("StatusPasien").value Then
                    arrGrafik(i, 2) = rs("JmlPasienPria").value
                    arrGrafik(i, 3) = rs("JmlPasienWanita").value
                    arrGrafik(i, 4) = rs("Total").value
                    rs.MoveNext
                Else
                    arrGrafik(i, 2) = 0
                    arrGrafik(i, 3) = 0
                    arrGrafik(i, 4) = 0
                End If
            Wend
    End Select
End Sub

Public Sub msubSetChart(oMSChart As MSChart, ChartType As ComboBox)
    ' Use the array MPGandMiles to create a three column chart. Set the
    '   ChartData property to the array, then set the title and column
    '   labels.
    With oMSChart
        ' Load chart data
        .ChartData = arrGrafik
        ' Set chart title
        Select Case mstrGrafik
            Case "JenisPasienPerRuangan"
                .Title = "Grafik Jenis Pasien per Ruangan"
            Case "JenisPasienPerJP"
                .Title = "Grafik Jenis Pasien"
            Case "StatusPasienPerRuangan"
                .Title = "Grafik Status Pasien per Ruangan"
            Case "StatusPasienPerSP"
                .Title = "Grafik Status Pasien"
        End Select
        ' Set number of colomns and their values
        If mstrGrafik = "JenisPasienPerJP" Or mstrGrafik = "StatusPasienPerSP" Then
            .ColumnCount = 3
            .ColumnLabelCount = 3
            .Column = 1
            .ColumnLabel = "Pasien Pria"
            .Column = 2
            .ColumnLabel = "Pasien Wanita"
            .Column = 3
            .ColumnLabel = "Total"
        Else
            .ColumnCount = mintJmlKolomGrafik
            .ColumnLabelCount = mintJmlKolomGrafik
            For i = 1 To mintJmlKolomGrafik
                .Plot.SeriesCollection(i).SeriesMarker.Show = True
                .Column = i
                .ColumnLabel = JnsKriteria(i)
            Next i
        End If
        .Refresh
        .Visible = True

        Select Case ChartType.ListIndex
            Case 0 To 9
                .ChartType = ChartType.ListIndex
            Case 10
                .ChartType = VtChChartType2dPie
            Case 11
                .ChartType = VtChChartType2dXY
        End Select

        With .Legend
            ' Make Legend Visible.
            .Location.Visible = True
            .Location.LocationType = VtChLocationTypeRight
            ' Set Legend properties.
            ' Right justify.
            .TextLayout.HorzAlignment = VtHorizontalAlignmentRight
            ' Use Yellow text.
            .VtFont.VtColor.Set 255, 255, 0
            .Backdrop.Fill.Style = VtFillStyleBrush
            .Backdrop.Fill.Brush.Style = VtBrushStyleSolid
            .Backdrop.Fill.Brush.FillColor.Set 255, 0, 255
        End With
    End With
End Sub

'untuk meload data dokter di grid
Public Sub msubLoadDokter(oForm As Form)
    '    On Error Resume Next
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & mstrFilterDokter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    mintJmlDokter = rs.RecordCount
    With oForm
        Set .dgDokter.DataSource = rs
        .dgDokter.Columns(0).Width = 1200
        .dgDokter.Columns(1).Width = 3000
        .dgDokter.Columns(2).Width = 400
        .dgDokter.Columns(3).Width = 3000
    End With
End Sub

Public Sub msubSetDeleteKeyComma(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 44 Then KeyAscii = 0
End Sub

Public Sub subSetSubTotalRow(vobjForm As Form, intRowNow As Integer, iColBegin As Integer, vbBackColor, vbForeColor)
    Dim i As Integer
    With vobjForm.fgData
        'tampilan Black & White
        For i = iColBegin To .cols - 1
            .row = intRowNow
            .col = i
            .CellBackColor = vbBackColor
            .CellForeColor = vbForeColor
            .CellFontBold = True
        Next
    End With
End Sub

Public Sub Animate(vForm As Form, vHeight As Integer, vTambah As Boolean)
    Dim i  As Integer

    If vTambah = True Then
        For i = 1 To vHeight
            DoEvents
            If vForm.Height >= vHeight Then Exit For
            vForm.Height = vForm.Height + 100
            Call centerForm(vForm, MDIUtama)
        Next
    Else
        For i = 1 To vHeight
            DoEvents
            If vForm.Height <= vHeight Then Exit For
            vForm.Height = vForm.Height - 100
            Call centerForm(vForm, MDIUtama)
        Next
    End If
End Sub

Public Function KeyHurufBesar(KeyAscii As Integer) As Integer
    KeyHurufBesar = Asc(UCase(Chr(KeyAscii)))
End Function

'seting visible menu
Public Sub SetVisibleMenu(strKdAplikasi As String, strNamaForm As String, strNamaMenu As Object, strStatus As String)
    strSQL = "SELECT * " & _
    " FROM StatusObject " & _
    " WHERE KdAplikasi= '" & strKdAplikasi & "' AND NamaForm= '" & strNamaForm & "' AND NamaObject= '" & strNamaMenu.Name & "' AND StatusEnable= '" & strStatus & "'"
    msubRecFO rs, strSQL
    If rs.RecordCount = 1 Then
        strNamaMenu.Visible = False
    Else
        strNamaMenu.Visible = True
    End If
End Sub

Public Function CROpenMSQLReport(strFile As String, conn As ADODB.connection, strTables As String, Optional blnErrorMessages As Boolean = False) As CRAXDRT.Report
    '====================================================================================================================
    ' Fungsi untuk membuka file report crystal (rpt) dengan koneksi database OLEDB
    '====================================================================================================================
    '
    ' strFile     : Nama file report
    ' strServer   : Nama server database
    ' strDatabase : Nama database
    ' strTables   : Daftar tabel yang dipisahkan dengan titik koma sesuai dengan urutannya pada crystal report designer
    '
    ' Return      : Instance obyek report jika sukses, nothing jika gagal
    '
    ' Note        : - Semua tabel/view yang ada pada report harus terletak pada satu database
    '               - Untuk melakukan query berkondisi, gunakan property RecordSelectionFormula setelah pemanggilan
    '                 fungsi ini
    '
    '====================================================================================================================

    Dim strREP As String, strSVR As String, strDBS As String, strTBL As String, strTmp As String
    Dim intCtr As Integer, intMax As Integer, intErr As Integer
    Dim strPAR() As String
    Dim rptOut As CRAXDRT.Report
    Dim crxApp As New CRAXDRT.Application

    'Init
    intErr = 0
    Set rptOut = Nothing
    strREP = Trim$(strFile)
    strSVR = Trim$(conn.Properties("Data Source").value)
    strDBS = Trim$(conn.Properties("Initial Catalog").value)
    strTBL = Trim$(strTables)
    'Cek
    If (strREP <> "") And (strSVR <> "") And (strDBS <> "") And (strTBL <> "") Then
        'Buka...
        Set rptOut = crxApp.OpenReport(strFile)
        If Err.Number = 0 Then
            'Parse tabel
            strPAR = Split(strTBL, ";")
            intMax = UBound(strPAR) + 1
            'Tentukan maksimum
            If intMax < rptOut.Database.Tables.Count Then
                intMax = rptOut.Database.Tables.Count
                ReDim Preserve strPAR(intMax - 1)
            Else
                intMax = rptOut.Database.Tables.Count
            End If
            intMax = intMax - 1
            'Loop tabel dan lakukan penambahan lokasi tabel ke report
            For intCtr = 0 To intMax
                'Handling nama tabel
                strTmp = Trim$(strPAR(intCtr))
                If strTmp = "" Then strTmp = rptOut.Database.Tables(intCtr + 1).Name
                'Tambahkan koneksi ke report
                Err.Clear
                rptOut.Database.AddOLEDBSource conn.ConnectionString, strTmp
                'OK?
                If Err.Number = 0 Then
                    With rptOut.Database.Tables(intCtr + 1)
                        .SetLogOnInfo strSVR, strDBS
                        .SetTableLocation strTmp, strDBS, ""
                    End With
                Else
                    intErr = intErr + 1
                End If
            Next
        Else
            If blnErrorMessages Then MsgBox "Tidak dapat membuka report pada lokasi:" & vbCrLf & strREP & vbCrLf & "File tidak ditemukan, atau file korup!", vbCritical
            Set rptOut = Nothing
        End If
    Else
        If blnErrorMessages Then MsgBox "Parameter yang diperlukan untuk membuka report tidak lengkap!", vbExclamation
    End If
    'Destroy...
    Set crxApp = Nothing
    'Err?
    If blnErrorMessages And (intErr > 0) Then MsgBox "Report dibuka dengan " & CStr(intErr) & " buah koneksi tabel mengalami error!", vbInformation
    'Return
    Set CROpenMSQLReport = rptOut
End Function

Public Sub GantiWarnaChk(ObjChk As Object, Objdc As Object)
    If ObjChk.value = 1 Then
        ObjChk.ForeColor = vbBlue
        ObjChk.Caption = Mid(ObjChk.Caption, 1, Len(ObjChk.Caption) - 4)
        Objdc.Enabled = True
    Else
        ObjChk.ForeColor = vbBlack
        Objdc.Enabled = flase
        ObjChk.Caption = ObjChk.Caption + " All"
    End If
End Sub

Public Sub ValidasiINADRG(Jenis As String, Panjang As Integer, PanjangField As Integer, strIsi As String)
    On Error GoTo hell
    Dim sTemp As String
    Dim s
    Dim X As Integer

    Select Case Jenis
        Case "angka"
            If Panjang <> PanjangField Then
                sTemp = ""
                For X = Panjang To PanjangField - 1
                    sTemp = sTemp & "0"
                Next X
                sTemp = sTemp & strIsi
            Else
                sTemp = strIsi
            End If

        Case "huruf"
            If Panjang <> PanjangField Then
                sTemp = ""
                For X = Panjang To PanjangField - 1
                    sTemp = sTemp & " "
                Next X
                sTemp = strIsi & sTemp
            Else
                sTemp = strIsi
            End If
    End Select

    strTampung = strTampung & sTemp

    Exit Sub
hell:
    Call msubPesanError
End Sub

Public Sub subSimpanData(strData As String, NamaFile As String)
    On Error GoTo hell
    'Open NamaFile For Append As #1
Close #1
Open NamaFile For Output As #1
    Print #1, strData
Close #1

Exit Sub
hell:
Call msubPesanError
End Sub
'Mengetahui jumlah hari dalam bulan
Public Function LastDayOfMonth(ByVal ValidDate As Date) As Byte
  Dim LastDay As Byte
  LastDay = DatePart("d", DateAdd("d", -1, DateAdd("m", 1, _
                     DateAdd("d", -DatePart("d", ValidDate) + 1, _
                     ValidDate))))
  LastDayOfMonth = LastDay
End Function
