VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmUpdateReservasi 
   BorderStyle     =   0  'None
   Caption         =   "Update No Rekam Medis Pasien Reservasi"
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1815
   ScaleWidth      =   12915
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Umur"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   9720
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      Begin VB.TextBox txtTahun 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         MaxLength       =   6
         TabIndex        =   6
         Top             =   330
         Width           =   375
      End
      Begin VB.TextBox txtBulan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   5
         Top             =   330
         Width           =   375
      End
      Begin VB.TextBox txtHari 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   4
         Top             =   330
         Width           =   375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "thn"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   720
         TabIndex        =   9
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "bln"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   270
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "hr"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   195
      End
   End
   Begin VB.TextBox TxtNoReservasi 
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Text            =   "No Reservasi"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   11520
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   10320
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Frame Data 
      Caption         =   "Update Data Pasien"
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   12615
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   13
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox TxtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         MaxLength       =   6
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox cbJenisKelamin 
         Height          =   315
         ItemData        =   "FrmUpdateReservasi.frx":0000
         Left            =   6120
         List            =   "FrmUpdateReservasi.frx":000A
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin MSMask.MaskEdBox meTglLahir 
         Height          =   390
         Left            =   7800
         TabIndex        =   14
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   688
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         HideSelection   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mm-yy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   645
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6120
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Lahir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7800
         TabIndex        =   15
         Top             =   360
         Width           =   1230
      End
   End
End
Attribute VB_Name = "FrmUpdateReservasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdSimpan_Click()

    If funcCekValidasi = False Then Exit Sub
    If MsgBox("Anda yakin akan UPDATE No. RM -> Reservasi pasien", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    
    
'    Call sp_ReservasiPasien(dbcmd)
    
    strSQL = "Update ReservasiPasien set NoCM = '" & txtNoCM.Text & "', NamaLengkap = '" & txtNamaPasien.Text & "', TglLahir = '" & Format(meTglLahir, "yyyy/MM/dd HH:mm:ss") & "'  where Noreservasi='" & Trim(txtNoReservasi.Text) & "'"
    Call msubRecFO(rs, strSQL)
    
    MsgBox "Update Data Sukses !", vbOKOnly, "Update Sukses"
    Call Add_HistoryLoginActivity("Update_NoRM_Reservasi")
    Call CmdTutup_Click

End Sub

Private Sub CmdTutup_Click()
    Unload FrmUpdateReservasi
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload FrmUpdateReservasi
End Sub


Private Sub TxtNoCM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        blnSibuk = True
        Call CariData
        cmdSimpan.SetFocus
    Else
'        cmdSimpan.SetFocus
    End If

    'If Not (KeyAscii >= 0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = 13 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    If KeyAscii = Asc(",") Then Exit Sub
    If KeyAscii = Asc(".") Then Exit Sub
End Sub

Private Function funcCekValidasi() As Boolean
'Validasi Untuk digit No CM & Telepon (Cyber 13 April 2014)
    If Len(txtNoCM.Text) < 6 Then
        MsgBox "No CM harus 6 Digit", vbExclamation, "Validasi"
        funcCekValidasi = False
        txtNoCM.SetFocus
        Exit Function
    End If
'    If txtTelepon.Text = "" Then
'        MsgBox "No Telepon harus diisi", vbExclamation, "Validasi"
'        funcCekValidasi = False
'        txtTelepon.SetFocus
'        Exit Function
'    End If
    
'----------------------Cyber-------------------

funcCekValidasi = True
End Function

Public Sub CariData()


    strSQL = "Select * from v_CariPasien WHERE [No. CM]='" & txtNoCM.Text & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        mstrNoCM = ""
'        cmdSimpan.Enabled = False
        Exit Sub
    End If
    
    mstrNoCM = txtNoCM.Text
    txtNamaPasien.Text = rs.Fields("Nama Lengkap").value
    If rs.Fields("JK").value = "P" Then
        cbJenisKelamin.Text = "Perempuan"
    ElseIf rs.Fields("JK").value = "L" Then
        cbJenisKelamin.Text = "Laki-laki"
    End If
    meTglLahir.Text = rs.Fields("tgllahir").value
    txtTahun.Text = rs.Fields("UmurTahun").value
    txtBulan.Text = rs.Fields("UmurBulan").value
    txtHari.Text = rs.Fields("UmurHari").value
    Set rs = Nothing
End Sub

Private Sub sp_ReservasiPasien(ByVal adoCommand As ADODB.Command)
 Set adoCommand = New ADODB.Command
    
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoReservasi", adInteger, adParamInput, 0, Trim(txtNoReservasi.Text))
        .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 6, Trim(txtNoCM.Text))
        .Parameters.Append .CreateParameter("NamaLengkap", adVarChar, adParamInput, 50, txtNamaPasien.Text)
        If cbJenisKelamin.Text = "Laki-laki" Then
            .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 2, "01")
        Else
            .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 2, "02")
        End If
        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(meTglLahir, "yyyy/MM/dd HH:mm:ss"))
'        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "U")

        .ActiveConnection = dbConn
        .CommandText = "Update_ReservasiPasien"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data", vbCritical, "Validasi"
        Else
            MsgBox "Update Data Sukses !", vbOKOnly, "Update Sukses"
            Call CmdTutup_Click
            
'            TxtNoReservasi.Text = .Parameters("OutputNoReservasi").value
'            TxtNoAntrian.Text = .Parameters("OutputNoAntrian").value
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
'    MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub TxtNoCM_LostFocus()
 Call CariData
End Sub
