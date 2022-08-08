VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDaftarPasienAsuransi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Asuransi"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienAsuransi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   12795
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   7320
      Width           =   12735
      Begin VB.CommandButton cmdRegistrasi 
         Caption         =   "Registrasi Pasien"
         Height          =   495
         Left            =   7920
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdDaftarPasien 
         Caption         =   "Pasien Daftar"
         Height          =   495
         Left            =   7920
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   495
         Left            =   9600
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   3855
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         Height          =   495
         Left            =   11160
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Cari berdasarkan no. CM atau nama peserta"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   5775
      End
   End
   Begin VB.Frame fraDaftar 
      Caption         =   "Daftar Pasien Asuransi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   12735
      Begin VB.Frame Frame3 
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
         Left            =   6840
         TabIndex        =   9
         Top             =   240
         Width           =   5775
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   11
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd  MMMM, yyyy"
            Format          =   128778243
            UpDown          =   -1  'True
            CurrentDate     =   37967
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   12
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd  MMMM, yyyy"
            Format          =   128778243
            UpDown          =   -1  'True
            CurrentDate     =   37967
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   13
            Top             =   315
            Width           =   255
         End
      End
      Begin VB.CheckBox chkTerdaftar 
         Caption         =   "Pasien Terdaftar"
         Height          =   230
         Left            =   5040
         TabIndex        =   8
         Top             =   600
         Width           =   2100
      End
      Begin MSDataGridLib.DataGrid dgDaftarPasienAsuransi 
         Height          =   5055
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   8916
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label LblJumData 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1245
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
      Left            =   10920
      Picture         =   "frmDaftarPasienAsuransi.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPasienAsuransi.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "frmDaftarPasienAsuransi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkTerdaftar_Click()
    Call cmdCari_Click
    
    If chkTerdaftar.value = vbChecked Then
        cmdDaftarPasien.Visible = False
        cmdRegistrasi.Visible = True
        cmdHapus.Enabled = False
    Else
        cmdDaftarPasien.Visible = True
        cmdRegistrasi.Visible = False
        cmdHapus.Enabled = True
    End If
End Sub

Private Sub cmdCari_Click()
On Error GoTo errLoad

    lblJumData.Caption = "Data 0/0"
    
    If chkTerdaftar.value = vbUnchecked Then
        Set rs = Nothing
        strSQL = "Select TglDaftar,IdPeserta,NoCM,NamaPeserta,TglLahir,Alamat,NamaPenjamin,Institusiasal,NamaGolongan " & _
                 "From V_DaftarAsuransi WHERE NamaPeserta like '%" & txtParameter.Text & "%' " & _
                 "and TglDaftar BETWEEN '" & Format(dtpAwal.value, "yyyy-MM-dd 00:00:00") & "' AND '" & Format(dtpAkhir.value, "yyyy-MM-dd 23:59:59") & "' and NoCM is null"
    Else
        
        If txtParameter.Text = "" Then
        
            strSQL = "Select TglDaftar,IdPeserta,NoCM,NamaPeserta,TglLahir,alamat,NamaPenjamin,Institusiasal,NamaGolongan " & _
                 "From V_DaftarAsuransi WHERE TglDaftar BETWEEN '" & Format(dtpAwal.value, "yyyy-MM-dd 00:00:00") & "' AND '" & Format(dtpAkhir.value, "yyyy-MM-dd 23:59:59") & "' and NoCM is not null"
        Else
        
            strSQL = "Select TglDaftar,IdPeserta,NoCM,NamaPeserta,TglLahir,alamat,NamaPenjamin,Institusiasal,NamaGolongan " & _
                 "From V_DaftarAsuransi WHERE TglDaftar BETWEEN '" & Format(dtpAwal.value, "yyyy-MM-dd 00:00:00") & "' AND '" & Format(dtpAkhir.value, "yyyy-MM-dd 23:59:59") & "' " & _
                 "and  NoCM like '%" & txtParameter.Text & "%' OR NamaPeserta like '%" & txtParameter.Text & "%' and NoCM is not null"
        
        End If
        
    End If

    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    
    Set dgDaftarPasienAsuransi.DataSource = rs
    Call SetGridAntrianPasien
    lblJumData.Caption = "Data 0 / " & dgDaftarPasienAsuransi.ApproxCount
'    If dgDaftarPasienAsuransi.ApproxCount > 0 Then
'        dgDaftarPasienAsuransi.SetFocus
'    End If
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdDaftarPasien_Click()
On Error Resume Next
    If dgDaftarPasienAsuransi.ApproxCount = 0 Then Exit Sub
    With frmPasienBaru
        strPasien = "Asuransi"
        
        .txtNoCM.Text = ""
        .dcPropinsi.Text = ""
        .dcKecamatan.Text = ""
        .dcKelurahan.Text = "" '
        .dcKota.Text = ""
        .txtAlamat.Text = ""
        .txtTelepon.Text = ""
        .txtKodePos.Text = ""
        .txtTahun.Text = ""
        .txtBulan.Text = ""
        .txtHari.Text = ""
        .txtIbuKandung.Text = ""
        .txtNamaPanggilan.Text = ""
        .txtNoKK.Text = ""
        .txtKepalaKeluarga.Text = ""
        .txtNoIdentitas.Text = ""
        .meRTRW.Text = "__/__"
        .Show
        
        
        
        .txtNamaPasien.Text = dgDaftarPasienAsuransi.Columns("Nama Peserta")
        .meTglLahir.Text = dgDaftarPasienAsuransi.Columns("Tgl Lahir")
        .txtFormPengirim.Text = Me.Name
        
    End With
'    frmPasienBaru.meTglLahir.Text = "__/__/____"
    
Exit Sub
'errload:
'    Call msubPesanError
End Sub

Private Sub cmdHapus_Click()

    If dgDaftarPasienAsuransi.ApproxCount = 0 Then Exit Sub
    
    If dgDaftarPasienAsuransi.Columns(1).value = "" Then Exit Sub
    If MsgBox("Anda yakin akan menghapus data pasien asuransi", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    
    If (dgDaftarPasienAsuransi.Columns(2)) <> "" Then
        strSQL = "delete from datapesertaAsuransi where nocm='" & dgDaftarPasienAsuransi.Columns(2).value & "'"
        Call msubRecFO(rs, strSQL)
    Else
        strSQL = "delete DataPesertaAsuransi where IDPeserta='" & dgDaftarPasienAsuransi.Columns(1).value & "'"
        dbConn.Execute strSQL
    End If
'    Call msubRecFO(rs, strSQL)
'    If dgDaftarPasienAsuransi.Columns(2).value Is Not Null Then
'        strSQL = "delete from datapesertaAsuransi where nocm='" & dgDaftarPasienAsuransi.Columns(2).value & "'"
'        Call msubRecFO(rs, strSQL)
'    End If
    Call cmdCari_Click
    
End Sub

Private Sub cmdRegistrasi_Click()
On Error GoTo errLoad

    If dgDaftarPasienAsuransi.ApproxCount = 0 Then Exit Sub
    If MsgBox("Pasien mau ke RAWAT JALAN atau PENUNJANG ?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
        strPasien = "Lama"
        mstrNoCM = dgDaftarPasienAsuransi.Columns(2).value
        With frmRegistrasiRJPenunjang
            .Show
            .txtAlamatRI.Text = ""
            .txtNoCM.Text = mstrNoCM
            .CariData
            
            strSQL = "Select Top(1) * from V_DataPesertaAsuransi where NoCM = '" & mstrNoCM & "'"
            Call msubRecFO(rs1, strSQL)
            If rs1.EOF = False Then
             frmRegistrasiRJPenunjang.tempKelompokPasien = rs1("KdKelompokPasien").value
            End If
        End With
    Else
        strPasien = "Lama"
        mstrNoCM = dgDaftarPasienAsuransi.Columns(2).value
        With frmRegistrasiAll
            .Show
            .txtAlamatRI.Text = ""
            .txtNoCM.Text = mstrNoCM
            .CariData
             strSQL = "Select Top(1) * from V_DataPesertaAsuransi where NoCM = '" & mstrNoCM & "'"
            Call msubRecFO(rs1, strSQL)
            If rs1.EOF = False Then
             frmRegistrasiAll.tempKelompokPasien = rs1("KdKelompokPasien").value
            End If
        End With
    End If
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgDaftarPasienAsuransi_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDaftarPasienAsuransi
    WheelHook.WheelHook dgDaftarPasienAsuransi
End Sub

Private Sub dgDaftarPasienAsuransi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    lblJumData.Caption = dgDaftarPasienAsuransi.Bookmark & " / " & dgDaftarPasienAsuransi.ApproxCount & " Data"
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    
    dtpAwal.value = Now
    dtpAkhir.value = Now
    
    Call cmdCari_Click
End Sub
Sub SetGridAntrianPasien()
On Error GoTo errLoad
    With dgDaftarPasienAsuransi
        .Columns(0).Caption = "TglDaftar"
        .Columns(0).Width = 1200
        .Columns(1).Caption = "Id Peserta"
        .Columns(1).Width = 1800
        .Columns(2).Caption = "No CM"
        .Columns(2).Width = 2800
        .Columns(3).Caption = "Nama Peserta"
        .Columns(3).Width = 2800
        .Columns(4).Caption = "Tgl Lahir"
        .Columns(4).Width = 2200
        .Columns(5).Caption = "Alamat"
        .Columns(5).Width = 1700
        .Columns(6).Caption = "Nama Penjamin"
        .Columns(6).Width = 1800
        .Columns(7).Caption = "Institusi Asal"
        .Columns(7).Width = 1800
        .Columns(8).Caption = "Nama Golongan"
        .Columns(8).Width = 1800
    End With
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtParameter_Change()
    Call cmdCari_Click
End Sub

Private Sub txtParameter_GotFocus()
    txtParameter.Text = ""
    ForeColor = &H80000012
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

