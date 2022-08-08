VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEditRuangPelayanan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Edit Data Ruang Pelayanan"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8325
   Icon            =   "frmEditRuangPelayanan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8325
   Begin VB.TextBox txtKdKelas 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      Left            =   3960
      TabIndex        =   24
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtKdRuangan 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      Left            =   2400
      TabIndex        =   23
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Perubahan Data Ruangan Pelayanan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   18
      Top             =   3000
      Width           =   8295
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcKelas 
         Height          =   315
         Left            =   3480
         TabIndex        =   22
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcNoKamarRI 
         Height          =   315
         Left            =   5640
         TabIndex        =   27
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcNoBedRI 
         Height          =   315
         Left            =   7080
         TabIndex        =   28
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label9 
         Caption         =   "No. Bed"
         Height          =   255
         Left            =   7080
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "No. Kamar"
         Height          =   255
         Left            =   5640
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Kelas"
         Height          =   255
         Left            =   3480
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label txtRuangan1 
         Caption         =   "Ruangan"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   14
      Top             =   4200
      Width           =   8295
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   6645
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   5040
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   495
         Left            =   3480
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Sebelumnya"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   8295
      Begin VB.TextBox txtNoKamar 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         TabIndex        =   30
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtNoBed 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   330
         Left            =   5040
         TabIndex        =   29
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtKelas 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   330
         Left            =   5040
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtRuangan 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         TabIndex        =   10
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtNmPasien 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   330
         Left            =   5040
         TabIndex        =   9
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtNoCM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtNoPendaftaran 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpTglPendaftaran 
         Height          =   360
         Left            =   5040
         TabIndex        =   6
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
         Format          =   120913923
         UpDown          =   -1  'True
         CurrentDate     =   41148
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kamar"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   1485
         Width           =   750
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Bed"
         Height          =   195
         Left            =   3600
         TabIndex        =   31
         Top             =   1485
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kelas"
         Height          =   195
         Left            =   3600
         TabIndex        =   13
         Top             =   1125
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1125
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pasien"
         Height          =   195
         Left            =   3600
         TabIndex        =   8
         Top             =   780
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No CM"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   765
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Pendaftaran"
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Pendaftaran"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   400
         Width           =   1215
      End
   End
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6480
      Picture         =   "frmEditRuangPelayanan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmEditRuangPelayanan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmEditRuangPelayanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdBatal_Click()
    txtNoPendaftaran.Text = ""
    txtNoCM.Text = ""
    txtRuangan.Text = ""
    txtKelas.Text = ""
    txtNmPasien.Text = ""
    txtKdRuangan.Text = ""
    txtKdKelas.Text = ""
    dtpTglPendaftaran.value = Now
    txtNoKamar.Text = ""
    txtNoBed.Text = ""
    
    dcRuangan.Text = ""
    dcKelas.Text = ""
    dcNoKamarRI.Text = ""
    dcNoBedRI.Text = ""
'    txtNoPendaftaran.SetFocus
    Call subLoadDcSource
    
End Sub

Private Sub cmdSimpan_Click()
Dim k As Integer
Dim a As Integer
On Error GoTo errLoad
    'uncheck Relationship
    dbConn.Execute "alter table dbo.DetailBiayaPelayanan Nocheck Constraint FK_DetailBiayaPelayanan_BiayaPelayanan"
    dbConn.Execute "alter table dbo.TempHargaKomponen Nocheck Constraint FK_TempHargaKomponen_DetailBiayaPelayanan"
    
    If sp_RuangBiayaPelayanan(txtNoPendaftaran.Text, txtKdRuangan.Text, txtKdKelas.Text, txtNoKamar.Text, txtNoBed.Text, dcRuangan.BoundText, dcKelas.BoundText, dcNoKamarRI.BoundText, dcNoBedRI.BoundText) = False Then Exit Sub
    
    strSQL = "Select Distinct KdPelayananRS From BiayaPelayanan Where NoPendaftaran = '" & txtNoPendaftaran.Text & "'"
    Call msubRecFO(rs, strSQL)
    
    If rs.RecordCount > 1 Then
        For k = 1 To rs.RecordCount
            If sp_BiayaPelayananPasienRI(txtNoPendaftaran.Text, txtKdRuangan.Text, rs(0).value, txtKdKelas.Text, dcRuangan.BoundText, dcKelas.BoundText) = False Then Exit Sub
'            If sp_TempHargaKomponenBPelayananPasienRI(txtNoPendaftaran.Text, rs(0).Value, dcKelas.BoundText) = False Then Exit Sub
            rs.MoveNext
        Next k
    End If
    
    strSQLX = "Select Distinct KdPelayananRS From TempHargaKomponen Where NoPendaftaran = '" & txtNoPendaftaran.Text & "'"
    Call msubRecFO(rsx, strSQLX)
    
    If rsx.RecordCount > 1 Then
        For a = 1 To rsx.RecordCount
'            If sp_BiayaPelayananPasienRI(txtNoPendaftaran.Text, txtKdRuangan.Text, rs(0).Value, txtKdKelas.Text, dcRuangan.BoundText, dcKelas.BoundText) = False Then Exit Sub
            If sp_TempHargaKomponenBPelayananPasienRI(txtNoPendaftaran.Text, rsx(0).value, dcKelas.BoundText) = False Then Exit Sub
            rsx.MoveNext
        Next a
    End If
    
    'check Relationship
    dbConn.Execute "alter table dbo.DetailBiayaPelayanan check Constraint FK_DetailBiayaPelayanan_BiayaPelayanan"
    dbConn.Execute "alter table dbo.TempHargaKomponen check Constraint FK_TempHargaKomponen_DetailBiayaPelayanan"
    
    Call cmdBatal_Click
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcKelas_GotFocus()
On Error GoTo errLoad
Dim tempKode As String

    tempKode = dcKelas.BoundText
    strSQL = "SELECT distinct KdKelas, Kelas FROM V_KelasPelayanan WHERE KdRuangan = '" & dcRuangan.BoundText & "' "
    Call msubDcSource(dcKelas, rs, strSQL)

    dcKelas.BoundText = tempKode

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKelas_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    
    If KeyAscii = 13 Then
        If Len(Trim(dcKelas.Text)) = 0 Then dcNoKamarRI.SetFocus: Exit Sub
'        If dcKelas.MatchedWithList = True Then dcNoKamarRI.SetFocus: Exit Sub
        strSQL = "select KdKelas,Kelas from V_KelasPelayanan WHERE (Kelas LIKE '%" & dcKelas.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKelas.BoundText = rs(0).value
        dcKelas.Text = rs(1).value
        dcNoKamarRI.SetFocus
    End If
    
    Exit Sub
    
errLoad:
    Call msubPesanError

End Sub

Private Sub dcNoBedRI_GotFocus()
On Error GoTo errLoad
Dim tempKode As String

    tempKode = dcNoBedRI.BoundText
    strSQL = "SELECT distinct dbo.StatusBed.NoBed, dbo.StatusBed.NoBed AS Alias, dbo.StatusBed.StatusEnabled, dbo.NoKamar.StatusEnabled as Expr" & _
               " FROM dbo.NoKamar INNER JOIN dbo.StatusBed ON dbo.NoKamar.KdKamar = dbo.StatusBed.KdKamar" & _
               " WHERE (KdRuangan = '" & dcRuangan.BoundText & "') AND (KdKelas = '" & dcKelas.BoundText & "') AND (dbo.StatusBed.StatusBed = 'K') AND (dbo.NoKamar.KdKamar = '" & dcNoKamarRI.BoundText & "')"
   
    Call msubDcSource(dcNoBedRI, rs, strSQL)
    dcNoBedRI.BoundText = tempKode
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcNoBedRI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub dcNoKamarRI_GotFocus()
On Error GoTo errLoad
Dim tempKode As String
    
    tempKode = dcNoKamarRI.BoundText
    strSQL = "SELECT distinct dbo.NoKamar.KdKamar,dbo.NoKamar.NamaKamar AS Alias, dbo.NoKamar.StatusEnabled, dbo.StatusBed.StatusEnabled " & _
            " FROM dbo.NoKamar INNER JOIN dbo.StatusBed ON dbo.NoKamar.KdKamar = dbo.StatusBed.KdKamar " & _
            " WHERE (NamaKamar NOT LIKE '%BOX%') AND (KdRuangan = '" & dcRuangan.BoundText & "') AND (KdKelas = '" & dcKelas.BoundText & "') AND (dbo.StatusBed.StatusBed = 'K') and dbo.NoKamar.StatusEnabled='1' and dbo.StatusBed.StatusEnabled='1' "
    
    Call msubDcSource(dcNoKamarRI, rs, strSQL)
    dcNoKamarRI.BoundText = tempKode
    
Exit Sub
errLoad:
    Call msubPesanError

End Sub

Private Sub dcNoKamarRI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcNoBedRI.SetFocus
End Sub

Private Sub dcRuangan_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    
    If KeyAscii = 13 Then
        If Len(Trim(dcRuangan.Text)) = 0 Then dcKelas.SetFocus: Exit Sub
        If dcRuangan.MatchedWithList = True Then dcKelas.SetFocus: Exit Sub
        strSQL = "select * from Ruangan WHERE (NamaRuangan LIKE '%" & dcRuangan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcRuangan.BoundText = rs(0).value
        dcRuangan.Text = rs(1).value
    End If
    
    Exit Sub
    
errLoad:
    Call msubPesanError

End Sub

Private Sub Form_Load()
    openConnection
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call cmdBatal_Click
    Call subLoadDcSource
End Sub

Private Sub txtNoPendaftaran_Change()
strSQL = "Select NoCM,NamaPasien,TglMasuk,RuanganPerawatan,KdRuangan,Kelas,KdKelas From V_DaftarInfoPasienRIAll Where NoPendaftaran = '" & txtNoPendaftaran.Text & "' And StatusPulang='T'"
        Call msubRecFO(rs, strSQL)
If txtNoPendaftaran <> "" Then
        If rs.RecordCount = 0 Then
            MsgBox "Data Tidak Ada"
            Exit Sub
        Else
            txtNoCM.Text = rs("NoCM").value
            txtNmPasien.Text = rs("NamaPasien").value
            txtRuangan.Text = rs("RuanganPerawatan").value
            txtKelas.Text = rs("Kelas").value
            dtpTglPendaftaran.value = Format(rs("TglMasuk").value, "dd/mm/yyyy hh:mm:ss")
            txtKdRuangan.Text = rs("KdRuangan").value
            txtKdKelas.Text = rs("KdKelas").value
            
            strSQLX = "Select KdKamar, NoBed From PemakaianKamar Where NoPendaftaran = '" & txtNoPendaftaran.Text & "' "
            Call msubRecFO(rsx, strSQLX)
            txtNoKamar.Text = rsx("KdKamar").value
            txtNoBed.Text = rsx("NoBed").value
            dcRuangan.SetFocus
        End If
End If
End Sub

Sub subLoadDcSource()

On Error GoTo hell
    
    Call msubDcSource(dcRuangan, rs, "Select * from Ruangan where StatusEnabled='1' and kdInstalasi = '03'")
'    Call msubDcSource(dcKelas, rsx, "Select KdKelas, Kelas from V_KelasPelayanan Where StatusEnabled='1' and KdRuangan='" & dcRuangan.BoundText & "'")
    Exit Sub
    
hell:
    Call msubPesanError
    Set rs = Nothing
    
End Sub

Private Function sp_RuangBiayaPelayanan(f_NoPendaftaran, f_KdRuangan, f_KdKelas, f_KdKamar, f_NoBed, f_KdRuanganBaru, f_KdKelasBaru, f_KdKamarBaru, f_NoBedBaru) As Boolean
On Error GoTo errLoad
    
    sp_RuangBiayaPelayanan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, f_KdKelas)
        .Parameters.Append .CreateParameter("KdKamar", adChar, adParamInput, 4, f_KdKamar)
        .Parameters.Append .CreateParameter("NoBed", adChar, adParamInput, 2, f_NoBed)
        .Parameters.Append .CreateParameter("KdRuanganBaru", adChar, adParamInput, 3, f_KdRuanganBaru)
        .Parameters.Append .CreateParameter("KdKelasBaru", adChar, adParamInput, 2, f_KdKelasBaru)
        .Parameters.Append .CreateParameter("KdKamarBaru", adChar, adParamInput, 4, f_KdKamarBaru)
        .Parameters.Append .CreateParameter("NoBedBaru", adChar, adParamInput, 2, f_NoBedBaru)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_RuangBiayaPelayanan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_RuangBiayaPelayanan = False
        End If
    End With
    
Exit Function
errLoad:
    sp_RuangBiayaPelayanan = False
    Call msubPesanError("sp_RuangBiayaPelayanan")
End Function

Private Function sp_BiayaPelayananPasienRI(f_NoPendaftaran, f_KdRuangan, f_KdPelayananRS, f_KdKelas, f_KdRuanganBaru, f_KdKelasBaru) As Boolean
On Error GoTo errLoad
    
    sp_BiayaPelayananPasienRI = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 10, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, f_KdKelas)
        .Parameters.Append .CreateParameter("KdRuanganBaru", adChar, adParamInput, 3, f_KdRuanganBaru)
        .Parameters.Append .CreateParameter("KdKelasBaru", adChar, adParamInput, 2, f_KdKelasBaru)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_BiayaPelayananPasienRI"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_BiayaPelayananPasienRI = False
        End If
    End With
    
Exit Function
errLoad:
    sp_BiayaPelayananPasienRI = False
    Call msubPesanError("sp_BiayaPelayananPasienRI")
End Function

Private Function sp_TempHargaKomponenBPelayananPasienRI(f_NoPendaftaran, f_KdPelayananRS, f_KdKelasBaru) As Boolean
On Error GoTo errLoad
    
    sp_TempHargaKomponenBPelayananPasienRI = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 10, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("KdKelasBaru", adChar, adParamInput, 2, f_KdKelasBaru)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_TempHargaKomponenBPelayananPasienRI"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_TempHargaKomponenBPelayananPasienRI = False
        End If
    End With
    
Exit Function
errLoad:
    sp_TempHargaKomponenBPelayananPasienRI = False
    Call msubPesanError("sp_TempHargaKomponenBPelayananPasienRI")
End Function


