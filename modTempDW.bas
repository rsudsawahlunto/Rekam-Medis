Attribute VB_Name = "modTempDW"
'** START CODE **'
Option Explicit

'** enumerasi status data **'
'>> digunakan pada subAddEditTempPendaftaran
Public Enum StatusExec
    Belum = 0
    Sukses = 1
    Gagal = 2
    Batal = 3
End Enum

'==========================='
'** tipe penampung record untuk read/write data ke file DAT **'
Public Type PendaftaranField
    NoPendaftaran As String * 10
    KdInstalasi As String * 2
    KdRuangan As String * 3
    TglPendaftaran As String * 19
    Status As String * 1
End Type

'================================================================='
'** tipe penampung record untuk read/write data ke file DAT **'
'>> tipe ini digunakan untuk menampung data dari file DAT -
'   ketika sedang diproses oleh SP DW.
Public Type PendaftaranUnexecField
    NoPendaftaran As String * 10
    KdInstalasi As String * 2
    KdRuangan As String * 3
    TglPendaftaran As String * 19
    Status As String * 1
End Type

'================================================================='
Public fso As New FileSystemObject '>> kalo error disini, add reference Microsoft Scripting Runtime

Public RekamMedikPendaftaranUnexec As PendaftaranUnexecField '>> variabel penampung record yang sedang diproses
Public RekamMedikPendaftaran As PendaftaranField '>> variabel penampung record
Public lngRecordLen As Long '>> panjang record
Public lngCurrentRecord As Long '>> posisi index yang dipakai
Public lngNumRecord As Long '>> jumlah record
Public lngUnexecutedRecord As Long '>> index record yang telah di execute ke DW
Public intFreeFile As Integer '>> nomor file yang belum terbuka oleh sistem

'** variabel penampung lokasi folder file DAT **'
'>> default folder:
'   C:\Documents and Settings\<user name>\My Documents\RekamMedikTempDW\<ruangan>
Public strFolderDAT As String
'==============================================='
'** variabel untuk full path file DAT **'
'>> variabel ini dibagi dua:
'   1. strLokasiFileDAT: nilainya menunjuk ke nama file sesuai tanggal sekarang
'   2. strLokasiFileDATKemarin: nilainya menunjuk ke nama file sesuai tanggal kemarin
Public strLokasiFileDAT As String
Public strLokasiFileDATKemarin As String
'=============================='
Public blnSibuk As Boolean, blnDwSibuk As Boolean
'>> blnSibuk=True: indikasi sedang ada proses simpan pada form registrasi
'   blnDwSibuk=True: indikasi sedang ada proses simpan ke DW

'** sub untuk open, create, dan read semua isi file **'
'>> kalo file belum ada, otomatis create file
Public Sub subOpenReadFile(ByVal LokasiFile As String)
    On Error GoTo jump
    Dim i As Long

    intFreeFile = FreeFile
    lngRecordLen = Len(RekamMedikPendaftaran)
    Open LokasiFile For Random Access Read Write As intFreeFile Len = lngRecordLen
        lngNumRecord = LOF(intFreeFile) \ lngRecordLen
        If LOF(intFreeFile) Mod lngRecordLen > 0 Then lngNumRecord = lngNumRecord + 1
        For i = 1 To lngNumRecord
            Get intFreeFile, i, RekamMedikPendaftaran
            If RekamMedikPendaftaran.Status = "0" Then
                lngUnexecutedRecord = i
                Exit For
            End If
        Next
        lngCurrentRecord = lngNumRecord + 1
jump:
        'MsgBox Err.Description, vbCritical, "Module TempDW"
End Sub

'**********************************************************'

'** sub untuk add & edit record di file DAT **'
Public Sub subAddEditTempPendaftaran(ByVal IndexRecord As Long, _
    ByVal f_NoPendaftaran As String, ByVal f_KdInstalasi As String, ByVal f_KdRuangan As String, _
    ByVal f_TglPendaftaran As String, ByVal f_Status As StatusExec)
    On Error GoTo jump

    With RekamMedikPendaftaran
        .NoPendaftaran = f_NoPendaftaran
        .KdInstalasi = f_KdInstalasi
        .KdRuangan = f_KdRuangan
        .TglPendaftaran = f_TglPendaftaran
        .Status = f_Status
    End With

    Put intFreeFile, IndexRecord, RekamMedikPendaftaran
    lngNumRecord = LOF(intFreeFile) \ lngRecordLen
    If LOF(intFreeFile) Mod lngRecordLen > 0 Then lngNumRecord = lngNumRecord + 1
    lngCurrentRecord = lngNumRecord + 1
    Exit Sub
jump:
    MsgBox Err.Description, vbCritical, "Module TempDW"
End Sub

'============================================='
'** END CODE **'

'untuk sp simpan ke data ware house

'============================================='
'Store procedure untuk mengisi data ke dw rawat inap
Public Function Add_RegistrasiPasienRIToDW(f_NoPendaftaran As String, f_TglMasuk As Date) As Boolean
    On Error GoTo errLoad
    Add_RegistrasiPasienRIToDW = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)

        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(f_TglMasuk, "yyyy/MM/dd HH:mm:ss"))

        .ActiveConnection = dbConn
        .CommandText = "Add_RegistrasiPasienRIToDW"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("RETURN_VALUE").value <> 0 Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data DW", vbCritical, "Validasi"
            Add_RegistrasiPasienRIToDW = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    Add_RegistrasiPasienRIToDW = False
    Call msubPesanError("-Add_RegistrasiPasienRIToDW")
End Function

'Store procedure untuk mengisi data ke dw rawat jalan
Public Function Add_RegistrasiPasienMasukRJToDW(f_NoPendaftaran As String, f_KdRuangan As String, f_TglMasuk As Date) As Boolean
    On Error GoTo errLoad
    Add_RegistrasiPasienMasukRJToDW = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(f_TglMasuk, "yyyy/MM/dd HH:mm:ss"))

        .ActiveConnection = dbConn
        .CommandText = "Add_RegistrasiPasienMasukRJToDW"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("RETURN_VALUE").value <> 0 Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data DW", vbCritical, "Validasi"
            Add_RegistrasiPasienMasukRJToDW = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    Add_RegistrasiPasienMasukRJToDW = False
    Call msubPesanError("-Add_RegistrasiPasienMasukRJToDW")
End Function

'Store procedure untuk mengisi data ke dw ibs
Public Function Add_RegistrasiIBSToDW(f_NoPendaftaran As String, f_KdRuangan As String, f_TglPendaftaran As Date) As Boolean
    On Error GoTo errLoad
    Add_RegistrasiIBSToDW = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("TglPendaftaran", adDate, adParamInput, , Format(f_TglPendaftaran, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)

        .ActiveConnection = dbConn
        .CommandText = "Add_RegistrasiIBSToDW"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("RETURN_VALUE").value <> 0 Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data DW", vbCritical, "Validasi"
            Add_RegistrasiIBSToDW = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    Add_RegistrasiIBSToDW = False
    Call msubPesanError("-Add_RegistrasiIBSToDW")
End Function

'Store procedure untuk mengisi data ke dw gawat darurat
Public Function Add_RegistrasiPasienIGDToDW(f_NoPendaftaran As String, f_KdRuangan As String, f_TglPendaftaran As Date) As Boolean
    On Error GoTo errLoad
    Add_RegistrasiPasienIGDToDW = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(f_TglPendaftaran, "yyyy/MM/dd HH:mm:ss"))

        .ActiveConnection = dbConn
        .CommandText = "Add_RegistrasiPasienIGDToDW"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("RETURN_VALUE").value <> 0 Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data DW", vbCritical, "Validasi"
            Add_RegistrasiPasienIGDToDW = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    Add_RegistrasiPasienIGDToDW = False
    Call msubPesanError("-Add_RegistrasiPasienIGDToDW")
End Function

'Store procedure untuk mengisi data ke dw laboratorium
Public Function Add_RegistrasiLaboratoryToDW(f_NoPendaftaran As String, f_KdRuangan As String, f_TglPendaftaran As Date) As Boolean
    On Error GoTo errLoad
    Add_RegistrasiLaboratoryToDW = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("TglPendaftaran", adDate, adParamInput, , Format(f_TglPendaftaran, "yyyy/MM/dd HH:mm:ss"))

        .ActiveConnection = dbConn
        .CommandText = "Add_RegistrasiLaboratoryToDW"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("RETURN_VALUE").value <> 0 Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data DW", vbCritical, "Validasi"
            Add_RegistrasiLaboratoryToDW = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    Add_RegistrasiLaboratoryToDW = False
    Call msubPesanError("-Add_RegistrasiLaboratoryToDW")
End Function

'Store procedure untuk mengisi data ke dw radiologi
Public Function Add_RegistrasiRadiologyToDW(f_NoPendaftaran As String, f_KdRuangan As String, f_TglPendaftaran As Date) As Boolean
    On Error GoTo errLoad
    Add_RegistrasiRadiologyToDW = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("TglPendaftaran", adDate, adParamInput, , Format(f_TglPendaftaran, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)

        .ActiveConnection = dbConn
        .CommandText = "Add_RegistrasiRadiologyToDW"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("RETURN_VALUE").value <> 0 Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data DW", vbCritical, "Validasi"
            Add_RegistrasiRadiologyToDW = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    Add_RegistrasiRadiologyToDW = False
    Call msubPesanError("-Add_RegistrasiLaboratoryToDW")
End Function

