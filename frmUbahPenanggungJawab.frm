VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmUbahPenanggungJawab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Ubah Penanggung Jawab Pasien"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUbahPenanggungJawab.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   12285
   Begin VB.TextBox txtNoBKM 
      Height          =   375
      Left            =   2640
      TabIndex        =   38
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame5 
      Caption         =   "Data Penanggungjawab Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   25
      Top             =   2160
      Width           =   12255
      Begin VB.CheckBox chkDiriSendiri 
         Caption         =   "&Diri Sendiri"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   10560
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtTlpRI 
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
         Left            =   6960
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtAlamatRI 
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
         Left            =   8160
         MaxLength       =   50
         TabIndex        =   7
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtNamaRI 
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
         Left            =   120
         MaxLength       =   20
         TabIndex        =   4
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtKodePos 
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
         Left            =   5280
         MaxLength       =   5
         TabIndex        =   13
         Top             =   2040
         Width           =   1455
      End
      Begin MSMask.MaskEdBox meRTRWPJ 
         Height          =   390
         Left            =   4200
         TabIndex        =   12
         Top             =   2040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo dcKotaPJ 
         Height          =   390
         Left            =   3960
         TabIndex        =   9
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcKecamatanPJ 
         Height          =   390
         Left            =   8040
         TabIndex        =   10
         Top             =   1320
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcKelurahanPJ 
         Height          =   390
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcPropinsiPJ 
         Height          =   390
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcHubungan 
         Height          =   390
         Left            =   3000
         TabIndex        =   5
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcPekerjaanPJ 
         Height          =   390
         Left            =   5520
         TabIndex        =   6
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Pekerjaan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5520
         TabIndex        =   36
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hubungan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   3000
         TabIndex        =   35
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Telephone"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   6960
         TabIndex        =   34
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Alamat Lengkap"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   8160
         TabIndex        =   33
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Nama Lengkap"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Kode Pos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5280
         TabIndex        =   31
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "RT/RW"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4200
         TabIndex        =   30
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Kelurahan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Kecamatan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8040
         TabIndex        =   28
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Propinsi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Kota/Kabupaten"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   26
         Top             =   1080
         Width           =   1350
      End
   End
   Begin VB.TextBox txtNoPakai 
      Height          =   495
      Left            =   480
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   20
      Top             =   4800
      Width           =   12255
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         TabIndex        =   39
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8670
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10365
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1095
      Left            =   0
      TabIndex        =   17
      Top             =   960
      Width           =   12255
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   7200
         MaxLength       =   9
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         MaxLength       =   12
         TabIndex        =   0
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7200
         TabIndex        =   21
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2280
         TabIndex        =   18
         Top             =   360
         Width           =   1350
      End
   End
   Begin VB.TextBox txtNoPendaftaran 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   0
      MaxLength       =   10
      TabIndex        =   22
      Top             =   1200
      Width           =   1695
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   37
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
      Left            =   10440
      Picture         =   "frmUbahPenanggungJawab.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmUbahPenanggungJawab.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmUbahPenanggungJawab.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "No. Pendaftaran"
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
      Index           =   0
      Left            =   0
      TabIndex        =   23
      Top             =   960
      Width           =   1605
   End
End
Attribute VB_Name = "frmUbahPenanggungJawab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim j As Integer

Dim varPropinsi As String
Dim varKota As String
Dim varKecamatan As String
Dim varKelurahan As String

Private Sub chkDiriSendiri_Click()
    On Error GoTo errLoad
    If chkDiriSendiri.value = vbChecked Then
        strSQL = "SELECT NamaLengkap, Alamat, Telepon,Propinsi,Kota,Kecamatan,Kelurahan,RTRW,Kodepos FROM Pasien WHERE NocM='" & txtnocm.Text & "'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount <> 0 Then
            txtNamaRI.Text = rs("NamaLengkap").value
            txtAlamatRI.Text = IIf(IsNull(rs("Alamat").value), "-", rs("Alamat").value)
            txtTlpRI.Text = IIf(IsNull(rs("Telepon")), "-", rs("Telepon").value)
            dcPropinsiPJ.Text = IIf(IsNull(rs("Propinsi")), "-", rs("Propinsi"))
            dcKotaPJ.Text = IIf(IsNull(rs("Kota")), "-", rs("Kota"))
            dcKecamatanPJ.Text = IIf(IsNull(rs("Kecamatan")), "-", rs("Kecamatan"))
            dcKelurahanPJ.Text = IIf(IsNull(rs("Kelurahan")), "-", rs("Kelurahan"))

            'load Pekerjaan Pasien
            strSQL = "SELECT Pekerjaan FROM detailPasien WHERE NocM='" & txtnocm.Text & "'"
            Call msubRecFO(rs, strSQL)
            dcPekerjaanPJ.Text = IIf(rs.RecordCount = 0, "-", rs("Pekerjaan"))

        Else
            txtNamaRI.Text = ""
            txtAlamatRI.Text = ""
            txtTlpRI.Text = ""
            dcPropinsiPJ.Text = ""
            dcKotaPJ.Text = ""
            dcKecamatanPJ.Text = ""
            dcKelurahanPJ.Text = ""
        End If
        
        
            txtNamaRI.Enabled = False
            txtAlamatRI.Enabled = False
            txtTlpRI.Enabled = False
            dcPropinsiPJ.Enabled = False
            dcKotaPJ.Enabled = False
            dcKecamatanPJ.Enabled = False
            dcKelurahanPJ.Enabled = False
            dcPekerjaanPJ.Enabled = False
            dcHubungan.Enabled = False
            meRTRWPJ.Enabled = False
            txtKodePos.Enabled = False
            
    Else
        txtNamaRI.Text = ""
        txtAlamatRI.Text = ""
        txtTlpRI.Text = ""
        dcPropinsiPJ.Text = ""
        dcKotaPJ.Text = ""
        dcKecamatanPJ.Text = ""
        dcKelurahanPJ.Text = ""
        
        
        txtNamaRI.Enabled = True
        txtAlamatRI.Enabled = True
        txtTlpRI.Enabled = True
        dcPropinsiPJ.Enabled = True
        dcKotaPJ.Enabled = True
        dcKecamatanPJ.Enabled = True
        dcKelurahanPJ.Enabled = True
        dcPekerjaanPJ.Enabled = True
        dcHubungan.Enabled = True
        meRTRWPJ.Enabled = True
        txtKodePos.Enabled = True
        
    End If
    dcHubungan.BoundText = ""
    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub chkDiriSendiri_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If chkDiriSendiri.value = vbChecked Then
            cmdSimpan.SetFocus
        Else
            txtNamaRI.SetFocus
        End If
    End If
End Sub

Private Sub cmdBatal_Click()
    txtNamaRI.Text = ""
    txtAlamatRI.Text = ""
    dcHubungan.BoundText = ""
    dcPekerjaanPJ.Text = ""
    dcPropinsiPJ.Text = ""
    dcKotaPJ.Text = ""
    dcKecamatanPJ.Text = ""
    dcKelurahanPJ.Text = ""
    meRTRWPJ.Text = "__/__"
    txtKodePos.Text = ""
    txtTlpRI.Text = ""
    txtNamaRI.Enabled = True
    txtNamaRI.SetFocus
    chkDiriSendiri.value = vbUnchecked
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    If txtNamaRI.Text = "" Then
        MsgBox "Nama Penanggung Jawab? !!", vbExclamation, "Validasi"
        Exit Sub
    End If
    
    If txtAlamatRI.Text = "" Then
        MsgBox "Alamat Penanggung Jawab? !!", vbExclamation, "Validasi"
        Exit Sub
    End If
    
  If chkDiriSendiri.value = vbChecked Then
        
    Else
        If Periksa("datacombo", dcHubungan, "Hubungan Kosong") = False Then Exit Sub
        
        If Periksa("datacombo", dcPekerjaanPJ, "Pekerjaan Kosong") = False Then Exit Sub
        
        If dcKecamatanPJ.Text <> "" Then
            If Periksa("datacombo", dcKecamatanPJ, "Kecamatan Tidak Terdaftar") = False Then Exit Sub
        End If
    
        If dcKelurahanPJ.Text <> "" Then
            If Periksa("datacombo", dcKelurahanPJ, "Kelurahan Tidak Terdaftar") = False Then Exit Sub
        End If
    
        If dcKotaPJ.Text <> "" Then
            If Periksa("datacombo", dcKotaPJ, "Kota Tidak Terdaftar") = False Then Exit Sub
        End If
        
        If dcPropinsiPJ.Text <> "" Then
            If Periksa("datacombo", dcPropinsiPJ, "Provinsi Tidak Terdaftar") = False Then Exit Sub
        End If
    End If

    
    Call Update_PenanggungJawabPasien(dbcmd)
    cmdSimpan.Enabled = False
    cmdTutup.SetFocus
    Exit Sub
errLoad:
    Call msubPesanError
    cmdSimpan.Enabled = True
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcHubungan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcPekerjaanPJ.SetFocus
End Sub

Private Sub dcHubungan_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    If KeyAscii = 13 Then
        If Len(Trim(dcHubungan.Text)) = 0 Then cmdSimpan.SetFocus
        If dcHubungan.MatchedWithList = True Then dcPekerjaanPJ.SetFocus
        strSQL = "SELECT Hubungan, NamaHubungan FROM HubunganKeluarga WHERE (NamaHubungan LIKE '%" & dcHubungan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcHubungan.Text = ""
            Exit Sub
        End If
        dcHubungan.BoundText = rs(0).value
        dcHubungan.Text = rs(1).value
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKecamatanPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcKelurahanPJ.SetFocus
End Sub

'Private Sub dcKecamatanPJ_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If dcKecamatanPJ.MatchedWithList = True Then dcKelurahanPJ.SetFocus
'        strSQL = "SELECT DISTINCT NamaKecamatan FROM V_Wilayah where  expr1=1 and Expr2='1'and statusenabled=1 and (NamaKecamatan LIKE '%" & dcKecamatanPJ.Text & "%')"
'        Call msubRecFO(rs, strSQL)
'        If rs.EOF = True Then
'            dcKecamatanPJ.Text = ""
'            dcKelurahanPJ.SetFocus
'            Exit Sub
'        End If
'        dcKecamatanPJ.BoundText = rs(0).value
'    End If
'End Sub
Private Sub dcKecamatanPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        j = 3
        dcKelurahanPJ.Enabled = True
        Call subLoadDataWilayah("kecamatan")
        If dcKelurahanPJ.Enabled = True Then
            dcKelurahanPJ.SetFocus
        Else

        End If
    End If
'    If KeyAscii = 13 Then
'        j = 3
'        dcKelurahanPJ.Enabled = True
'        Call subLoadDataWilayah("kecamatan")
'        If dcKelurahanPJ.Enabled = True Then
'            dcKelurahanPJ.SetFocus
'        Else
'
'        End If
'    End If
End Sub
Private Sub dcKecamatanPJ_LostFocus()
    dcKecamatanPJ = Trim(StrConv(dcKecamatanPJ, vbProperCase))

End Sub
Private Sub dcKelurahanPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then meRTRWPJ.SetFocus
End Sub

'Private Sub dcKelurahanPJ_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If dcKelurahanPJ.MatchedWithList = True Then meRTRWPJ.SetFocus
'        strSQL = "SELECT DISTINCT NamaKelurahan, NamaKelurahan AS alias FROM V_Wilayah where Expr3='1' and (NamaKelurahan LIKE '%" & dcKelurahanPJ.Text & "%')"
'        Call msubRecFO(rs, strSQL)
'        If rs.EOF = True Then
'            dcKelurahanPJ.Text = ""
'            Exit Sub
'        End If
'        dcKelurahanPJ.BoundText = rs(0).value
'        dcKelurahanPJ.Text = rs(1).value
'    End If
'End Sub
Private Sub dcKelurahanPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        j = 4
        Call subLoadDataWilayah("desa")
      '  txtKodePos.SetFocus
    End If
End Sub

Private Sub dcKelurahanPJ_LostFocus()
    dcKelurahanPJ = Trim(StrConv(dcKelurahanPJ, vbProperCase))
End Sub

Private Sub dcKotaPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcKecamatanPJ.SetFocus
End Sub

'Private Sub dcKotaPJ_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If dcKotaPJ.MatchedWithList = True Then dcKecamatanPJ.SetFocus
'        strSQL = "SELECT DISTINCT NamaPropinsi, NamaKotaKabupaten, NamaKotaKabupaten AS alias FROM V_Wilayah where Expr1='1' and statusenabled='1' and (NamaKotaKabupaten LIKE '%" & dcKotaPJ.Text & "%')"
'        Call msubRecFO(rs, strSQL)
'        If rs.EOF = True Then
'            dcKotaPJ.Text = ""
'            Exit Sub
'        End If
'        dcKotaPJ.BoundText = rs(0).value
'        dcKotaPJ.Text = rs(1).value
'    End If
'End Sub

Private Sub dcKotaPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        j = 2
        dcKecamatanPJ.Enabled = True
        dcKelurahanPJ.Enabled = True
        Call subLoadDataWilayah("kota")
        If dcKecamatanPJ.Enabled = True Then
            dcKecamatanPJ.SetFocus
        End If
    End If
End Sub

Private Sub dcKotaPJ_LostFocus()
    dcKotaPJ = Trim(StrConv(dcKotaPJ, vbProperCase))
End Sub


Private Sub dcPekerjaanPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtAlamatRI.SetFocus
End Sub

Private Sub dcPekerjaanPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcPekerjaanPJ.MatchedWithList = True Then txtAlamatRI.SetFocus
        strSQL = "SELECT DISTINCT kdpekerjaan, Pekerjaan, Pekerjaan AS alias FROM Pekerjaan WHERE (Pekerjaan LIKE '%" & dcPekerjaanPJ.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcPekerjaanPJ.Text = ""
            Exit Sub
        End If
        dcPekerjaanPJ.BoundText = rs(0).value
        dcPekerjaanPJ.Text = rs(1).value
    End If
End Sub

Private Sub dcPropinsiPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcKotaPJ.SetFocus
End Sub

'Private Sub dcPropinsiPJ_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If dcPropinsiPJ.MatchedWithList = True Then dcKotaPJ.SetFocus
'        strSQL = "SELECT DISTINCT NamaPropinsi, NamaPropinsi AS alias FROM V_Wilayah where StatusEnabled='1' and (NamaPropinsi LIKE '%" & dcPropinsiPJ.Text & "%')"
'        Call msubRecFO(rs, strSQL)
'        If rs.EOF = True Then
'            dcPropinsiPJ.Text = ""
'            Exit Sub
'        End If
'        dcPropinsiPJ.BoundText = rs(0).value
'        dcPropinsiPJ.Text = rs(1).value
'    End If
'End Sub

Private Sub dcPropinsiPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        j = 1
        dcKotaPJ.Enabled = True
        dcKecamatanPJ.Enabled = True
        dcKelurahanPJ.Enabled = True
        Call subLoadDataWilayah("propinsi")
        If dcKotaPJ.Enabled = True Then
            dcKotaPJ.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    If mblnCariPasien = True Then frmCariPasien.Enabled = False
    subDcSource "Propinsi"
    Call subDcSource2
    
End Sub

Private Sub hgPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub meRTRWPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKodePos.SetFocus
End Sub

Private Sub txtAlamatRI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcPropinsiPJ.SetFocus
End Sub

Private Sub txtKodePos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtTlpRI.SetFocus
End Sub

Private Sub txtKodePos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTlpRI.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtNamaRI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcHubungan.SetFocus
End Sub

Private Sub txtNoPendaftaran_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'Store procedure untuk mengisi registrasi pasien RI
Private Sub Update_PenanggungJawabPasien(ByVal adoCommand As ADODB.Command)
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtnocm.Text)
        .Parameters.Append .CreateParameter("NamaPJ", adVarChar, adParamInput, 20, txtNamaRI.Text)
        .Parameters.Append .CreateParameter("PekerjaanPJ", adVarChar, adParamInput, 30, dcPekerjaanPJ.Text)
        .Parameters.Append .CreateParameter("Hubungan", adChar, adParamInput, 2, IIf(dcHubungan.BoundText = "", Null, dcHubungan.BoundText))
        .Parameters.Append .CreateParameter("AlamatPJ", adVarChar, adParamInput, 50, IIf(txtAlamatRI.Text = "", Null, txtAlamatRI.Text))
        .Parameters.Append .CreateParameter("PropinsiPJ", adVarChar, adParamInput, 25, IIf(dcPropinsiPJ.Text = "", Null, dcPropinsiPJ.Text))
        .Parameters.Append .CreateParameter("KotaPJ", adVarChar, adParamInput, 25, IIf(dcKotaPJ.Text = "", Null, dcKotaPJ.Text))
        .Parameters.Append .CreateParameter("KecamatanPJ", adVarChar, adParamInput, 25, IIf(dcKecamatanPJ.Text = "", Null, dcKecamatanPJ.Text))
        .Parameters.Append .CreateParameter("KelurahanPJ", adVarChar, adParamInput, 25, IIf(dcKelurahanPJ.Text = "", Null, dcKelurahanPJ.Text))
        .Parameters.Append .CreateParameter("RTRWPJ", adVarChar, adParamInput, 25, IIf(meRTRWPJ.Text = "", Null, meRTRWPJ.Text))
        .Parameters.Append .CreateParameter("KodePosPJ", adVarChar, adParamInput, 25, IIf(meRTRWPJ.Text = "", Null, txtKodePos.Text))
        .Parameters.Append .CreateParameter("TeleponPJ", adVarChar, adParamInput, 20, IIf(Len(Trim(txtTlpRI.Text)) = 0, Null, Trim(txtTlpRI.Text)))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_PenanggungJawabPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Error", vbCritical, "Validasi"
        Else
            MsgBox "Data Penanggung Jawab Berhasil diubah", vbInformation, "Informasi"
            Call Add_HistoryLoginActivity("Update_PenanggungJawabPasien")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

Private Sub subDcSource2()
    On Error GoTo errLoad
    Call msubDcSource(dcHubungan, rs, "SELECT Hubungan, NamaHubungan FROM HubunganKeluarga")

'    strSQL = "SELECT DISTINCT NamaPropinsi, NamaPropinsi AS alias FROM V_Wilayah where StatusEnabled='1'"
'    Call msubDcSource(dcPropinsiPJ, rs, strSQL)
'
'    strSQL = "SELECT DISTINCT NamaKotaKabupaten, NamaKotaKabupaten AS alias FROM V_Wilayah where Expr1='1' and statusenabled=1"
'    Call msubDcSource(dcKotaPJ, rs, strSQL)
'
'    strSQL = "SELECT DISTINCT NamaKecamatan, NamaKecamatan AS alias FROM V_Wilayah where expr1=1 and Expr2='1'and statusenabled=1"
'    Call msubDcSource(dcKecamatanPJ, rs, strSQL)
'
'    strSQL = "SELECT DISTINCT NamaKelurahan, NamaKelurahan AS alias FROM V_Wilayah where expr1=1 and Expr2='1'and statusenabled=1 and Expr3='1'"
'    Call msubDcSource(dcKelurahanPJ, rs, strSQL)

    strSQL = "SELECT DISTINCT Pekerjaan,Pekerjaan AS alias FROM Pekerjaan"
    Call msubDcSource(dcPekerjaanPJ, rs, strSQL)

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subDcSource(varstrPilihan As String, Optional varStrSQL As String)
    Select Case varstrPilihan

        Case "Propinsi"
            strSQL = "SELECT DISTINCT KdPropinsi, NamaPropinsi AS alias FROM V_Wilayah where StatusEnabled=1 order by NamaPropinsi"
            Set rsPropinsi = Nothing
            rsPropinsi.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dcPropinsiPJ.RowSource = rsPropinsi
            dcPropinsiPJ.BoundColumn = rsPropinsi(0).Name
            dcPropinsiPJ.ListField = rsPropinsi(1).Name
        Case "Kota"
            strSQL = "SELECT DISTINCT KdKotaKabupaten, NamaKotaKabupaten AS alias FROM V_Wilayah " & varStrSQL & ""
            Set rsKota = Nothing
            rsKota.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dcKotaPJ.RowSource = rsKota
            dcKotaPJ.BoundColumn = rsKota(0).Name
            dcKotaPJ.ListField = rsKota(1).Name
        Case "Kecamatan"
            strSQL = "SELECT DISTINCT KdKecamatan, NamaKecamatan AS alias FROM V_Wilayah " & varStrSQL & ""
            Set rsKecamatan = Nothing
            rsKecamatan.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dcKecamatanPJ.RowSource = rsKecamatan
            dcKecamatanPJ.BoundColumn = rsKecamatan(0).Name
            dcKecamatanPJ.ListField = rsKecamatan(1).Name
        Case "Kelurahan"
            strSQL = "SELECT DISTINCT KdKelurahan, NamaKelurahan AS alias FROM V_Wilayah " & varStrSQL & ""
            Set rsKelurahan = Nothing
            rsKelurahan.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dcKelurahanPJ.RowSource = rsKelurahan
            dcKelurahanPJ.BoundColumn = rsKelurahan(0).Name
            dcKelurahanPJ.ListField = rsKelurahan(1).Name
    End Select

    Exit Sub
End Sub

Private Sub txtTlpRI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSimpan.SetFocus
    Call SetKeyPressToNumber(KeyCode)
End Sub

'untuk mengganti nocm on change
Public Sub CariData()
    On Error GoTo hell
    'cek pasien igd
    strSQL = "SELECT NoCM FROM V_DaftarPasienIGDAktif WHERE (NoCM = '" & txtnocm.Text & "')"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        MsgBox "Pasien tersebut belum keluar dari IGD", vbInformation, "Informasi"
        mstrNoCM = ""
        cmdSimpan.Enabled = False
        Exit Sub
    End If

    'cek pasien ri
    strSQL = "SELECT dbo.RegistrasiRI.NoCM, dbo.Ruangan.NamaRuangan FROM dbo.RegistrasiRI INNER JOIN dbo.Ruangan ON dbo.RegistrasiRI.KdRuangan = dbo.Ruangan.KdRuangan WHERE (NoCM = '" & txtnocm.Text & "') AND StatusPulang = 'T'"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        MsgBox "Pasien tersebut belum keluar dari Rawat Inap," & vbNewLine & "Ruangan " & rs("NamaRuangan") & " ", vbInformation, "Informasi"
        mstrNoCM = ""
        cmdSimpan.Enabled = False
        Exit Sub
    End If

    strSQL = "Select * from v_CariPasien WHERE [No. CM]='" & txtnocm.Text & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        mstrNoCM = ""
        cmdSimpan.Enabled = False
        Exit Sub
    End If

    mstrNoCM = txtnocm.Text
    txtNamaPasien.Text = rs.Fields("Nama Lengkap").value
    Set rs = Nothing
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub txtTlpRI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub


Private Sub subLoadDataWilayah(strPencarian As String)
    'On Error GoTo errLoad
    On Error Resume Next
    Dim strTempSql As String

    Select Case strPencarian
        Case "propinsi"
            If Len(Trim(dcPropinsiPJ.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaPropinsi LIKE '%" & dcPropinsiPJ.Text & "%')and statusenabled=1"

        Case "kota"
            If Len(Trim(dcKotaPJ.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaPropinsi LIKE '%" & dcPropinsiPJ.Text & "%') and (NamaKotaKabupaten LIKE '%" & dcKotaPJ.Text & "%')"

        Case "kecamatan"
            If Len(Trim(dcKecamatanPJ.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaPropinsi LIKE '%" & dcPropinsiPJ.Text & "%') and (NamaKotaKabupaten LIKE '%" & dcKotaPJ.Text & "%') and (NamaKecamatan LIKE '%" & dcKecamatanPJ.Text & "%')"
        Case "desa"
            If Len(Trim(dcKelurahanPJ.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaPropinsi LIKE '%" & dcPropinsiPJ.Text & "%') and (NamaKotaKabupaten LIKE '%" & dcKotaPJ.Text & "%') and (NamaKecamatan LIKE '%" & dcKecamatanPJ.Text & "%') and (NamaKelurahan LIKE '%" & dcKelurahanPJ.Text & "%')"

        Case "kodepos"
            If Len(Trim(txtKodePos.Text)) = 0 Then Exit Sub
            strTempSql = " WHERE (NamaPropinsi LIKE '%" & dcPropinsiPJ.Text & "%') and (NamaKotaKabupaten LIKE '%" & dcKotaPJ.Text & "%') and (NamaKecamatan LIKE '%" & dcKecamatanPJ.Text & "%') and (NamaKelurahan LIKE '%" & dcKelurahanPJ.Text & "%') and (KodePos LIKE '%" & txtKodePos.Text & "%')"

    End Select

    strSQL = "SELECT DISTINCT ISNULL(NamaPropinsi, '') AS NamaPropinsi, ISNULL(NamaKotaKabupaten, '') AS NamaKotaKabupaten, ISNULL(NamaKecamatan, '')  AS NamaKecamatan, ISNULL(NamaKelurahan, '') AS NamaKelurahan, ISNULL(KodePos, '') AS KodePos" & _
    " FROM V_Wilayah" & _
    " " & strTempSql

    Call msubRecFO(rs, strSQL)
    If rs.EOF Then
        MsgBox "Data Wilayah Tidak Sesuai, Harap Cek Data Wilayah", vbInformation, "Validasi"

        dcPropinsiPJ.BoundText = ""
        dcKotaPJ.BoundText = ""
        dcKecamatanPJ.BoundText = ""
        dcKelurahanPJ.BoundText = ""
        txtKodePos.Text = ""

    ElseIf j = 1 Then
        If rs(1).value = "" Then
            MsgBox "Data Kota/Kabupaten Belum Ada", vbInformation, "Validasi"
            dcKotaPJ.Enabled = False
            dcKecamatanPJ.Enabled = False
            dcKelurahanPJ.Enabled = False
        Else

        End If

    ElseIf j = 2 Then
        If rs(2).value = "" Then
            MsgBox "Data Kecamatan Belum Ada", vbInformation, "Validasi"
            dcKecamatanPJ.Enabled = False
            dcKelurahanPJ.Enabled = False
        Else

        End If

    ElseIf j = 3 Then
        If rs(3).value = "" Then
            MsgBox "Data Kelurahan Belum Ada", vbInformation, "Validasi"
            dcKelurahanPJ.Enabled = False
        Else

        End If

    Else
        dcPropinsiPJ.Text = rs("NamaPropinsi")
        dcKotaPJ.Text = rs("NamaKotaKabupaten")
        dcKecamatanPJ.Text = rs("NamaKecamatan")
        dcKelurahanPJ.Text = rs("NamaKelurahan")
        txtKodePos.Text = rs("KodePos")
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub
Private Sub dcPropinsipj_Click(Area As Integer)
    dcKotaPJ.Text = ""
    dcKecamatanPJ.Text = ""
    dcKelurahanPJ.Text = ""
    txtKodePos = ""
    CekPilihanWilayah "dcPropinsiPJ", "Click"
End Sub

Private Sub dcKotaPJ_Click(Area As Integer)
    dcKecamatanPJ.Text = ""
    dcKelurahanPJ.Text = ""
    txtKodePos = ""
    CekPilihanWilayah "dcKotaPJ", "Click"
End Sub

Private Sub dcKecamatanPJ_Click(Area As Integer)
    dcKelurahanPJ.Text = ""
    txtKodePos = ""
    CekPilihanWilayah "dcKecamatanPJ", "Click"
End Sub

Private Sub dcKelurahanPJ_Click(Area As Integer)
    txtKodePos = ""
    CekPilihanWilayah "dcKelurahanPJ", "Click"
End Sub

Private Sub CekPilihanWilayah(strItem As String, Optional strEvent As String)
    Dim X As Integer
    Dim Y

    X = 0
    Select Case strItem
        Case "dcPropinsiPJ"
            Set dcKotaPJ.RowSource = Nothing
            Set dcKecamatanPJ.RowSource = Nothing
            Set dcKelurahanPJ.RowSource = Nothing
            dcKotaPJ.Text = ""
            dcKecamatanPJ.Text = ""
            dcKelurahanPJ.Text = ""
            txtKodePos = ""
            Select Case strEvent
                Case "Click"
                    subDcSource "Kota", " where kdPropinsi = '" & dcPropinsiPJ.BoundText & "' order by NamaKotaKabupaten"
                Case "KeyPress"
                    If dcPropinsiPJ.MatchedWithList = False Then
                        MsgBox "Pilih Propinsi"
                        X = 1
                        GoTo kosong
                        dcPropinsiPJ.SetFocus
                    Else
                        subDcSource "Kota", " where kdPropinsi = '" & dcPropinsiPJ.BoundText & "' order by NamaKotaKabupaten"
                        dcKotaPJ.SetFocus
                    End If
                Case "LostFocus"
                    If dcPropinsiPJ.MatchedWithList = False Then
                        MsgBox "Pilih Propinsi"
                        X = 1
                        GoTo kosong
                        dcPropinsiPJ.SetFocus
                    Else
                        subDcSource "Kota", " where kdPropinsi = '" & dcPropinsiPJ.BoundText & "' order by NamaKotaKabupaten"
                        dcKotaPJ.SetFocus
                    End If
            End Select
        Case "dcKotaPJ"
            Set dcKecamatanPJ.RowSource = Nothing
            Set dcKelurahanPJ.RowSource = Nothing
            dcKecamatanPJ.Text = ""
            dcKelurahanPJ.Text = ""
            txtKodePos = ""
            If dcPropinsiPJ.MatchedWithList = True Then
                Select Case strEvent
                    Case "Click"
                        If dcKotaPJ.Text = "" Then Exit Sub
                        subDcSource "Kecamatan", " where kdKotaKabupaten = '" & dcKotaPJ.BoundText & "' order by NamaKecamatan"
                    Case "KeyPress"
                        If dcKotaPJ.MatchedWithList = False Then
                           MsgBox "Pilih Kota"
                            X = 2
                            GoTo kosong
                            dcKotaPJ.SetFocus
                        Else
                            subDcSource "Kecamatan", " where kdKotaKabupaten = '" & dcKotaPJ.BoundText & "' order by NamaKecamatan"
                            dcKecamatanPJ.SetFocus
                        End If
                    Case "LostFocus"
                        If dcKotaPJ.MatchedWithList = False Then
                            MsgBox "Pilih Kota"
                            X = 2
                            GoTo kosong
                            dcKotaPJ.SetFocus
                        Else
                            subDcSource "Kecamatan", " where kdKotaKabupaten = '" & dcKotaPJ.BoundText & "' order by NamaKecamatan"
                            dcKecamatanPJ.SetFocus
                        End If
                End Select
            End If
        Case "dcKecamatanPJ"
            Set dcKelurahanPJ.RowSource = Nothing
            dcKelurahanPJ.Text = ""
            txtKodePos = ""
            If dcKotaPJ.MatchedWithList = True Then
                Select Case strEvent
                    Case "Click"
                        If dcKecamatanPJ.Text = "" Then Exit Sub
                        subDcSource "Kelurahan", " where kdkecamatan = '" & dcKecamatanPJ.BoundText & "' order by NamaKelurahan"
                    Case "KeyPress"
                        If dcKecamatanPJ.MatchedWithList = False Then
                            MsgBox "Pilih Kecamatan"
                            X = 3
                            GoTo kosong
                            dcKecamatanPJ.SetFocus
                        Else
                            subDcSource "Kelurahan", " where kdkecamatan = '" & dcKecamatanPJ.BoundText & "' order by NamaKelurahan"
                            dcKelurahanPJ.SetFocus
                        End If
                    Case "LostFocus"
                        If dcKecamatanPJ.MatchedWithList = False Then
                            MsgBox "Pilih Kecamatan"
                            X = 3
                            GoTo kosong
                            dcKecamatanPJ.SetFocus
                        Else
                            subDcSource "Kelurahan", " where kdkecamatan = '" & dcKecamatanPJ.BoundText & "' order by NamaKelurahan"
                            dcKelurahanPJ.SetFocus
                        End If
                End Select
            End If
        Case "dcKelurahanPJ"
            txtKodePos = ""
            If dcKecamatanPJ.MatchedWithList = True Then
                Select Case strEvent
                    Case "KeyPress"
                        If dcKelurahanPJ.MatchedWithList = False Then
                            MsgBox "Pilih Desa/Kelurahan"
                            X = 4
                            GoTo kosong
                            dcKelurahanPJ.Text = ""
                            dcKelurahanPJ.SetFocus
                        Else
                            txtKodePos.SetFocus
                        End If
                    Case "LostFocus"
                        If dcKelurahanPJ.MatchedWithList = False Then
                            MsgBox "Pilih Desa/Kelurahan"
                            X = 4
                            GoTo kosong
                            dcKelurahanPJ.SetFocus
                        End If
                End Select
            End If
    End Select

    Exit Sub

kosong:
    Y = MsgBox("Mulai lagi dari awal", vbYesNo, "Wilayah") ' vbYesNoCancel
    Select Case Y
        Case vbYes
            dcPropinsiPJ.Text = ""
            dcKotaPJ.Text = ""
            dcKecamatanPJ.Text = ""
            dcKelurahanPJ.Text = ""
            dcPropinsiPJ.SetFocus
        Case vbNo
            Exit Sub

    End Select
End Sub
