VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmKirimTerimaDokumen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kirim Terima Dokumen Rekam Medis"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13845
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKirimTerimaDokumen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   13845
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   10320
      TabIndex        =   13
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   12240
      TabIndex        =   14
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Frame frTerimaDokumen 
      Caption         =   "Terima Dokumen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   27
      Top             =   3240
      Width           =   13815
      Begin VB.TextBox txtKeterangaTerima 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6000
         TabIndex        =   11
         Top             =   480
         Width           =   4815
      End
      Begin MSComCtl2.DTPicker dtpTglTerima 
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy HH:mm"
         Format          =   117899267
         UpDown          =   -1  'True
         CurrentDate     =   38212
      End
      Begin MSDataListLib.DataCombo dcUserTerima 
         Height          =   315
         Left            =   11040
         TabIndex        =   12
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
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
      Begin MSDataListLib.DataCombo dcRuanganPengirim 
         Height          =   315
         Left            =   2760
         TabIndex        =   10
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Terima"
         Height          =   210
         Left            =   480
         TabIndex        =   31
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Ruang Pengirim"
         Height          =   210
         Left            =   2760
         TabIndex        =   30
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan Terima"
         Height          =   210
         Left            =   6000
         TabIndex        =   29
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "User Terima"
         Height          =   210
         Left            =   11040
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frKirimDokumen 
      Caption         =   "Kirim Dokumen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   22
      Top             =   2160
      Width           =   13815
      Begin VB.TextBox txtKeteranganKirim 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6000
         TabIndex        =   7
         Top             =   480
         Width           =   4815
      End
      Begin MSComCtl2.DTPicker dtpTglKirim 
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy HH:mm"
         Format          =   117899267
         UpDown          =   -1  'True
         CurrentDate     =   38212
      End
      Begin MSDataListLib.DataCombo dcUserKirim 
         Height          =   315
         Left            =   11040
         TabIndex        =   8
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
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
      Begin MSDataListLib.DataCombo dcRuanganTujuan 
         Height          =   315
         Left            =   2760
         TabIndex        =   6
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Kirim"
         Height          =   210
         Left            =   480
         TabIndex        =   26
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Ruang Tujuan"
         Height          =   210
         Left            =   2760
         TabIndex        =   25
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan Kirim"
         Height          =   210
         Left            =   6000
         TabIndex        =   24
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "User Kirim"
         Height          =   210
         Left            =   11040
         TabIndex        =   23
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame frDataPasien 
      Caption         =   "Data Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   16
      Top             =   1080
      Width           =   11535
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   480
         TabIndex        =   0
         Text            =   "1234567890"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtNoCM 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Text            =   "123456"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2760
         TabIndex        =   2
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6840
         TabIndex        =   3
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtRuanganPelayanan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7200
         TabIndex        =   4
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   480
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "NoCM"
         Height          =   210
         Left            =   1920
         TabIndex        =   20
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   2760
         TabIndex        =   19
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "JK"
         Height          =   210
         Left            =   6840
         TabIndex        =   18
         Top             =   240
         Width           =   180
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan Pelayanan"
         Height          =   210
         Left            =   7200
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.ComboBox cbojnsPrinter 
      Height          =   330
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   32
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
      Left            =   12000
      Picture         =   "frmKirimTerimaDokumen.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmKirimTerimaDokumen.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmKirimTerimaDokumen.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmKirimTerimaDokumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Barcode39 As clsBarCode39
Dim subPrinterZebra As Printer
Dim X As String

Private Sub cmdSimpan_Click()
    If frKirimDokumen.Enabled = True Then
        If Periksa("datacombo", dcRuanganTujuan, "Isi Ruangan Tujuan !!") = False Then Exit Sub
    End If
    If frTerimaDokumen.Enabled = True Then
        If Periksa("datacombo", dcRuanganPengirim, "Isi Ruangan Pengirim!!") = False Then Exit Sub
    End If
    Call sp_KirimTerimaDokumenRekamMedis(dbcmd)
    MsgBox "Pemasukan data dokumen rekam medis pasien Berhasil", vbInformation, "Informasi"
    cmdSimpan.Enabled = False
    Call subClearData
    frmDaftarDokumenRekamMedisPasien.cmdCari_Click
End Sub

Private Sub cmdTutup_Click()
    Unload Me
    frmDaftarDokumenRekamMedisPasien.Enabled = True
End Sub

Private Sub dcRuanganPengirim_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcRuanganPengirim.MatchedWithList = True Then txtKeterangaTerima.SetFocus
        strSQL = "SELECT KdRuangan, NamaRuangan FROM  Ruangan  WHERE NamaRuangan LIKE '" & dcRuanganPengirim.Text & "%' and StatusEnabled='1' ORDER BY NamaRuangan"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcRuanganPengirim.Text = ""
            Exit Sub
        End If
        dcRuanganPengirim.BoundText = rs(0).value
        dcRuanganPengirim.Text = rs(1).value
    End If
End Sub

Private Sub dcRuanganPengirim_LostFocus()
If dcRuanganPengirim.MatchedWithList = False Then dcRuanganPengirim.Text = ""
End Sub

Private Sub dcRuanganTujuan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcRuanganTujuan.MatchedWithList = True Then txtKeteranganKirim.SetFocus
        strSQL = "SELECT KdRuangan, NamaRuangan FROM  Ruangan  WHERE NamaRuangan LIKE '" & dcRuanganPengirim.Text & "%' and StatusEnabled='1' ORDER BY NamaRuangan"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcRuanganTujuan.Text = ""
            Exit Sub
        End If
        dcRuanganTujuan.BoundText = rs(0).value
        dcRuanganTujuan.Text = rs(1).value
    End If
End Sub

Private Sub dcUserKirim_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcUserKirim.MatchedWithList = True Then cmdSimpan.SetFocus
        strSQL = "SELECT IdPegawai, NamaLengkap FROM DataPegawai where NamaLengkap LIKE '" & dcUserKirim.Text & "%' ORDER BY NamaLengkap"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcUserKirim.Text = ""
            Exit Sub
        End If
        dcUserKirim.BoundText = rs(0).value
        dcUserKirim.Text = rs(1).value
    End If
End Sub

Private Sub dcUserTerima_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcUserTerima.MatchedWithList = True Then cmdSimpan.SetFocus
        strSQL = "SELECT IdPegawai, NamaLengkap FROM DataPegawai where NamaLengkap LIKE '" & dcUserTerima.Text & "%' ORDER BY NamaLengkap"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcUserTerima.SetFocus
            Exit Sub
        End If
        dcUserTerima.BoundText = rs(0).value
        dcUserTerima.Text = rs(1).value
    End If
End Sub

Private Sub dtpTglKirim_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcRuanganTujuan.SetFocus
End Sub

Private Sub dtpTglTerima_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcRuanganPengirim.SetFocus
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim subCtrlKey As String
'
'    subCtrlKey = (Shift + vbCtrlMask)
'
'    Select Case KeyCode
'        Case vbKeyF1
'
'        Case vbKeyL
'            If subCtrlKey = 4 Then
'                Unload Me
'                frmRegistrasiAll.Show
'            End If
'
'        Case vbKeyF3
'            Unload Me
'            frmCariPasien.Show
'    End Select
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call subDcSource
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
'    If strPasien = "View" Then
'        If strRegistrasi = "RJ" Then
'        ElseIf strRegistrasi = "DaftarPasienRIRJIGD" Then
'            Call frmDaftarPasienRJRIIGD.cmdCari_Click
'        ElseIf strRegistrasi = "PasienLama" Then
'            Call frmRegistrasiAll.CariData
'        End If
'    End If
    frmDaftarDokumenRekamMedisPasien.Enabled = True
End Sub

Private Sub txtJK_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtKeteranganKirim_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeteranganKirim.SetFocus
End Sub

Private Sub txtKeterangaTerima_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangaTerima.SetFocus
End Sub

Private Sub txtNamaPasien_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtNamaPasien_LostFocus()
    txtNamaPasien = StrConv(txtNamaPasien, vbProperCase)
End Sub

Private Sub TxtNoCM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        strSQL = "SELECT NoCM, Title + ' ' + NamaLengkap AS NamaPasien FROM Pasien WHERE (NoCM = '" & txtNoCM.Text & "' )"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            MsgBox "No. CM tersebut sudah dipakai " & rs("NamaPasien").value & "", vbExclamation, "Validasi"
            Exit Sub
        End If
    End If
    If Not (KeyAscii >= 0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

'Private Sub sp_KirimTerimaDokumenRekamMedis(ByVal adoCommand As ADODB.Command)
'    Set dbcmd = New ADODB.Command
'    With adoCommand
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
'        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
'        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
'        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan) 'ruangan pengirim
'        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, dcRuanganTujuan.BoundText) 'ruangan penerima
'        .Parameters.Append .CreateParameter("TglKirim", adDate, adParamInput, , Format(dtpTglKirim.value, "yyyy/MM/dd HH:mm:ss"))
'
'        .Parameters.Append .CreateParameter("TglTerima", adDate, adParamInput, , IIf(frTerimaDokumen.Enabled = False, Null, Format(dtpTglTerima.value, "yyyy/MM/dd HH:mm:ss")))
'        .Parameters.Append .CreateParameter("IdUserKirim", adChar, adParamInput, 10, dcUserKirim.BoundText)
'        .Parameters.Append .CreateParameter("IdUserTerima", adChar, adParamInput, 10, IIf(frTerimaDokumen.Enabled = False, Null, dcUserTerima.BoundText))
'        .Parameters.Append .CreateParameter("KeteranganKirim", adVarChar, adParamInput, 200, txtKeteranganKirim.Text)
'        .Parameters.Append .CreateParameter("KeteranganTerima", adVarChar, adParamInput, 200, IIf(frTerimaDokumen.Enabled = False, Null, txtKeterangaTerima.Text))
'
'        .ActiveConnection = dbConn
'        .CommandText = "dbo.Add_KirimTerimaDokumenRekamMedisPasien"
'        .CommandType = adCmdStoredProc
'        .Execute
'
'        If Not (.Parameters("RETURN_VALUE").value = 0) Then
'            MsgBox "Ada kesalahan dalam pemasukan dokumen rekam medis Pasien", vbCritical, "Validasi"
'        Else
'            Call Add_HistoryLoginActivity("Add_KirimTerimaDokumenRekamMedisPasien")
'        End If
'        Call deleteADOCommandParameters(adoCommand)
'        Set adoCommand = Nothing
'
'    End With
'    Exit Sub
'End Sub

'Store procedure untuk mengisi identitas pasien
Private Sub sp_KirimTerimaDokumenRekamMedis(ByVal adoCommand As ADODB.Command)
    Set dbcmd = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        If frKirimDokumen.Enabled = False Then
         .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, dcRuanganPengirim.BoundText) 'ruangan pengirim
        Else
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan) 'ruangan pengirim
        End If
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, dcRuanganTujuan.BoundText) 'ruangan penerima
        .Parameters.Append .CreateParameter("TglKirim", adDate, adParamInput, , Format(dtpTglKirim.value, "yyyy/MM/dd HH:mm:ss"))

        .Parameters.Append .CreateParameter("TglTerima", adDate, adParamInput, , IIf(frTerimaDokumen.Enabled = False, Null, Format(dtpTglTerima.value, "yyyy/MM/dd HH:mm:ss")))
        .Parameters.Append .CreateParameter("IdUserKirim", adChar, adParamInput, 10, dcUserKirim.BoundText)
        .Parameters.Append .CreateParameter("IdUserTerima", adChar, adParamInput, 10, IIf(frTerimaDokumen.Enabled = False, Null, dcUserTerima.BoundText))
        .Parameters.Append .CreateParameter("KeteranganKirim", adVarChar, adParamInput, 200, txtKeteranganKirim.Text)
        .Parameters.Append .CreateParameter("KeteranganTerima", adVarChar, adParamInput, 200, IIf(frTerimaDokumen.Enabled = False, Null, txtKeterangaTerima.Text))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_KirimTerimaDokumenRekamMedisPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan dokumen rekam medis Pasien", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Add_KirimTerimaDokumenRekamMedisPasien")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'untuk membersihkan data pasien
Private Sub subClearData()
    On Error Resume Next
    dcRuanganTujuan.Text = ""
    txtKeteranganKirim.Text = ""
    dcRuanganTujuan.SetFocus

    dcRuanganPengirim.Text = ""
    txtKeterangaTerima.Text = ""

End Sub

Private Sub subDcSource()
    On Error GoTo errLoad

    strSQL = "SELECT IdPegawai, NamaLengkap FROM DataPegawai where IdPegawai LIKE '%" & dcUserKirim.Text & "%' ORDER BY NamaLengkap"
    Call msubDcSource(dcUserKirim, rs, strSQL)

    strSQL = "SELECT IdPegawai, NamaLengkap FROM DataPegawai where IdPegawai LIKE '%" & dcUserTerima.Text & "%' ORDER BY NamaLengkap"
    Call msubDcSource(dcUserTerima, rs, strSQL)

    strSQL = "SELECT KdRuangan, NamaRuangan FROM  Ruangan where StatusEnabled='1' ORDER BY NamaRuangan"
    Call msubDcSource(dcRuanganTujuan, rs, strSQL)

    strSQL = "SELECT KdRuangan, NamaRuangan FROM  Ruangan where StatusEnabled='1' ORDER BY NamaRuangan"
    Call msubDcSource(dcRuanganPengirim, rs, strSQL)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub
