VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9f.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmDaftarTagihanMaskin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Daftar Tagihan Pasien MASKIN"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmDaftarTagihanMaskin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   8715
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   6840
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   2160
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   8715
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
         Left            =   3480
         TabIndex        =   3
         Top             =   150
         Width           =   5055
         Begin MSComCtl2.DTPicker DTPickerAwal 
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy "
            Format          =   55902211
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker DTPickerAkhir 
            Height          =   375
            Left            =   2760
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   55902211
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   2400
            TabIndex        =   4
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Instalasi Pemeriksaan"
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
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
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6840
      Picture         =   "FrmDaftarTagihanMaskin.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "FrmDaftarTagihanMaskin.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "FrmDaftarTagihanMaskin.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "FrmDaftarTagihanMaskin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'by splakuk 2008-Nop
Option Explicit

Private Sub cmdCetak_Click()
On Error GoTo errLoad

If dcInstalasi.BoundText = "" Then MsgBox "Pilih Instalasi", vbExclamation, "Validasi": Exit Sub
    cmdCetak.Enabled = False

    strSQL = "SELECT * FROM V_DaftarTagihanMaskin WHERE (TglPeriksa BETWEEN '" _
            & Format(DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
            & Format(DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "' AND KdInstalasi = '" & dcInstalasi.BoundText & "') order by TglPeriksa"

    Set rsx = Nothing
    
    rsx.Open strSQL, dbConn, adOpenForwardOnly, adLockBatchOptimistic
    If rsx.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbExclamation, "Informasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    
    frmCetakDaftarTagihanMASKIN.Show
    cmdCetak.Enabled = True
Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub DTPickerAkhir_Change()
    DTPickerAkhir.MaxDate = Now
End Sub

Private Sub DTPickerAwal_Change()
    DTPickerAwal.MaxDate = Now
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    With Me
        .DTPickerAwal.Value = Format(Now, "dd MMM yyyy 00:00:00")
        .DTPickerAkhir.Value = Now
    End With
    strSQL = "SELECT KdInstalasi, NamaInstalasi FROM Instalasi WHERE KdInstalasi NOT IN ('05','07','13','14','15','17','18','19','20','21','23')"
    Call msubDcSource(dcInstalasi, dbRst, strSQL)
Exit Sub
errLoad:
    Call msubPesanError
End Sub

