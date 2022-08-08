VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmMutasiBarangNM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Mutasi Barang Non Medis"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMutasiBarangNM.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   10485
   Begin MSDataGridLib.DataGrid dgNamaPenerima 
      Height          =   2535
      Left            =   3000
      TabIndex        =   23
      Top             =   4920
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   16
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
         Size            =   9
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
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Pengiriman"
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
      TabIndex        =   24
      Top             =   960
      Width           =   10455
      Begin VB.TextBox txtregister 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   34
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtNoKirim 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         MaxLength       =   15
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtNamaPenerima 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7680
         TabIndex        =   3
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtKdUserPenerima 
         Height          =   315
         Left            =   9240
         TabIndex        =   25
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpTglKirim 
         Height          =   330
         Left            =   3840
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
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
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   120061955
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSDataListLib.DataCombo dcRuanganPenerima 
         Height          =   330
         Left            =   5280
         TabIndex        =   2
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Register"
         Height          =   210
         Index           =   5
         Left            =   1680
         TabIndex        =   35
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Kirim"
         Height          =   210
         Index           =   10
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Kirim"
         Height          =   210
         Index           =   9
         Left            =   3840
         TabIndex        =   28
         Top             =   240
         Width           =   750
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan Penerima"
         Height          =   210
         Index           =   11
         Left            =   5280
         TabIndex        =   27
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Penerima"
         Height          =   210
         Index           =   8
         Left            =   7680
         TabIndex        =   26
         Top             =   240
         Width           =   1260
      End
   End
   Begin MSDataGridLib.DataGrid dgCariBarang 
      Height          =   2535
      Left            =   840
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   15
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
         Size            =   9
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
         AllowRowSizing  =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   0
      TabIndex        =   17
      Top             =   4440
      Width           =   10455
      Begin VB.TextBox txtCariBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   9
         Top             =   3360
         Width           =   3240
      End
      Begin MSDataGridLib.DataGrid dgMutasiBarang 
         Height          =   3015
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   5318
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   15
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
            Size            =   9
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
            AllowRowSizing  =   0   'False
            Locked          =   -1  'True
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Cari Barang"
         Height          =   210
         Index           =   6
         Left            =   255
         TabIndex        =   19
         Top             =   3480
         Width           =   900
      End
      Begin VB.Label lblJmlData 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Jumlah Barang"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   8985
         TabIndex        =   18
         Top             =   4260
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   8640
      TabIndex        =   13
      Top             =   8400
      Width           =   1575
   End
   Begin VB.Frame fraBarang 
      Height          =   2535
      Left            =   0
      TabIndex        =   14
      Top             =   1920
      Width           =   10455
      Begin VB.TextBox txtketerangan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1920
         MaxLength       =   25
         TabIndex        =   38
         Top             =   2040
         Width           =   5520
      End
      Begin VB.TextBox txtmutasi 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1920
         MaxLength       =   25
         TabIndex        =   36
         Top             =   1680
         Width           =   5520
      End
      Begin VB.TextBox txtStok 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1920
         MaxLength       =   25
         TabIndex        =   21
         Top             =   1320
         Width           =   1320
      End
      Begin VB.TextBox txtKdBarang 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   20
         Text            =   "txtkdbarang"
         Top             =   0
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.TextBox txtNamaBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   4
         Top             =   600
         Width           =   5520
      End
      Begin MSDataListLib.DataCombo dcAsalBarang 
         Height          =   330
         Left            =   1920
         TabIndex        =   6
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
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
      Begin VB.TextBox txtQtyBarang 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4800
         MaxLength       =   25
         TabIndex        =   7
         Top             =   960
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo dcKondisiBarang 
         Height          =   330
         Left            =   4800
         TabIndex        =   31
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
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
      Begin MSDataListLib.DataCombo dcverifikator 
         Height          =   330
         Left            =   1920
         TabIndex        =   41
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Verifikator"
         Height          =   210
         Index           =   13
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan Lainnya"
         Height          =   210
         Index           =   12
         Left            =   240
         TabIndex        =   39
         Top             =   2040
         Width           =   1605
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Alasan Mutasi"
         Height          =   210
         Index           =   7
         Left            =   240
         TabIndex        =   37
         Top             =   1680
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
         Height          =   210
         Index           =   3
         Left            =   4080
         TabIndex        =   33
         Top             =   960
         Width           =   300
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Kondisi"
         Height          =   210
         Index           =   4
         Left            =   3960
         TabIndex        =   32
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Stok Ruangan"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   1140
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nama Barang"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Asal Barang"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   930
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   30
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
      Left            =   8640
      Picture         =   "frmMutasiBarangNM.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMutasiBarangNM.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMutasiBarangNM.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmMutasiBarangNM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tempbolTampil As Boolean
Dim tempbolEdit As Boolean

Private Sub cmdBatal_Click()
On Error GoTo Errload

    Call subKosong
    Call subLoadDcSource
    Call subLoadGridSource
    tempbolEdit = False
    dtpTglKirim.SetFocus

Exit Sub
Errload:
End Sub

Private Sub cmdHapus_Click()
On Error GoTo Errload
    If (txtNoKirim.Text = "") Then Exit Sub
    If txtKdBarang.Text = "" Then
        MsgBox "Nama barang kosong", vbExclamation, "Validasi": txtNamaBarang.SetFocus: Exit Sub
    End If
    
    If MsgBox("Anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub

    'If sp_StockBarang(CDbl(txtStok.Text) + CDbl(txtQtyBarang.Text)) = False Then Exit Sub
    'If sp_StockBarang(CDbl(txtStok.Text)) = False Then Exit Sub
    'dbConn.Execute "DELETE MutasiBarangNonMedis WHERE NoKirim = '" & txtNoKirim.Text & "' AND KdBarang ='" & txtKdBarang.Text & "' AND KdAsal='" & dcAsalBarang.BoundText & "' AND NoRegister='" & txtregister.Text & "' "
    If sp_MutasiBarangNonMedis("D") = False Then Exit Sub
    Call cmdBatal_Click
    
Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo Errload
    If Periksa("datacombo", dcRuanganPenerima, "Ruangan Tujuan kosong") = False Then Exit Sub
    If Periksa("text", txtNamaPenerima, "Nama Penerima kosong") = False Then Exit Sub
    If txtKdUserPenerima.Text = "" Then
        txtNamaPenerima.SetFocus
        MsgBox "Nama Penerima Salah", vbInformation
        Exit Sub
    End If
    
    If Periksa("text", txtNamaBarang, "Barang kosong") = False Then Exit Sub
    If Periksa("datacombo", dcAsalBarang, "Asal barang kosong") = False Then Exit Sub
    If Periksa("datacombo", dcKondisiBarang, "Kondisi barang kosong") = False Then Exit Sub
    If Periksa("datacombo", dcverifikator, "Nama Verifikator kosong") = False Then Exit Sub
    If txtQtyBarang = "" Or txtQtyBarang < 1 Then
        MsgBox "Jumlah Barang Tidak Boleh kurang Dari 0 atau Kosong", vbCritical, "validasi"
        txtQtyBarang.SetFocus
        Exit Sub
    End If
    If CDbl(txtQtyBarang.Text) > CDbl(txtStok.Text) Then
        MsgBox "Jumlah Barang tidak boleh melebihi Stok Ruangan", vbCritical, "Validasi"
        Exit Sub
    End If
    
    If sp_StrukKirim() = False Then Exit Sub
    If tempbolEdit = True Then
        If sp_StockBarang(CDbl(txtStok.Text + CDbl(dgMutasiBarang.Columns("QtyBarang")))) = False Then Exit Sub
    End If
    
    Call msubRecFO(rs, "select dbo.FB_TakeStokBrgNonMedis('" & mstrKdRuangan & "', '" & txtKdBarang & "','" & dcAsalBarang.BoundText & "') as stok")
    If rs.EOF = False Then txtStok.Text = rs(0).Value Else txtStok.Text = 0
        
    'If sp_StockBarang(CDbl(txtStok.Text - txtQtyBarang.Text)) = False Then Exit Sub
    If sp_StockBarang(CDbl(txtStok.Text)) = False Then Exit Sub
    If sp_MutasiBarangNonMedis("A") = False Then Exit Sub
    
    Call cmdBatal_Click

Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcAsalBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtQtyBarang.SetFocus
End Sub

Private Sub dcKondisiBarang_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtmutasi.SetFocus
End Sub

Private Sub dcRuanganPenerima_KeyPress(KeyAscii As Integer)
On Error GoTo Errload
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToChar(KeyAscii)

    If KeyAscii = 13 Then
        'If Len(Trim(dcRuanganPenerima.Text)) = 0 Then txtNamaPenerima.SetFocus: Exit Sub
'        If Periksa("datacombo", dcRuanganPenerima, "Ruangan Penerima Salah") = False Then
'            MsgBox "Ruangan Penerima Salah", vbInformation
'
'            Exit Sub
'        End If
        If dcRuanganPenerima.MatchedWithList = True Then txtNamaPenerima.SetFocus
        Call msubRecFO(dbRst, "SELECT KdRuangan, NamaRuangan FROM Ruangan WHERE NamaRuangan LIKE '%" & dcRuanganPenerima.Text & "%' and StatusEnabled='1'")
        If dbRst.EOF = True Then dcRuanganPenerima.BoundText = "": Exit Sub
        dcRuanganPenerima.BoundText = dbRst(0).Value
        dcRuanganPenerima.Text = dbRst(1).Value
'        If txtNamaPenerima.Enabled = False Then Exit Sub
'        txtNamaPenerima.SetFocus
         End If
    
Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dcRuanganPenerima_LostFocus()
    If dcRuanganPenerima.Text = "" Then Exit Sub
    If dcRuanganPenerima.MatchedWithList = False Then dcRuanganPenerima.Text = ""

'       If Periksa("datacombo", dcRuanganPenerima, "Ruangan Penerima Salah") = False Then
'            MsgBox "Ruangan Penerima Salah", vbInformation
'
'            Exit Sub
'        End If

'    If dcRuanganPenerima.MatchedWithList = True Then txtNamaPenerima.SetFocus: Exit Sub
'    Call msubRecFO(dbRst, "SELECT KdRuangan, NamaRuangan FROM Ruangan WHERE NamaRuangan LIKE '%" & dcRuanganPenerima.Text & "%' and StatusEnabled='1'")
'    If dbRst.EOF = True Then dcRuanganPenerima.BoundText = "": Exit Sub
'    dcRuanganPenerima.BoundText = dbRst(0).Value
'    dcRuanganPenerima.Text = dbRst(1).Value
    
End Sub

Private Sub dcverifikator_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then
      If dcverifikator.MatchedWithList = True Then txtNamaBarang.SetFocus
'       Call msubDcSource(dcverifikator, rs, "SELECT idpegawai, namalengkap FROM datapegawai where namalengkap like '%" & dcverifikator.Text & "%'")
'        If rs.EOF = True Then dcverifikator = "": Exit Sub
'        dcverifikator.BoundText = rs(0).Value
'        dcverifikator.Text = rs(1).Value
    End If

End Sub

'Private Sub dcverifikator_LostFocus()
'    Call msubDcSource(dcverifikator, rs, "SELECT idpegawai, namalengkap FROM datapegawai where namalengkap like '%" & dcverifikator.Text & "%'")
'    If rs.EOF = True Then dcverifikator = "": Exit Sub
'    dcverifikator.BoundText = rs(0).Value
'    dcverifikator.Text = rs(1).Value
'End Sub

Private Sub dgCariBarang_Click()
WheelHook.WheelUnHook
        Set MyProperty = dgCariBarang
        WheelHook.WheelHook dgCariBarang
End Sub

Private Sub dgCariBarang_DblClick()
On Error GoTo Errload

    With dgCariBarang
    
    
    
        If .ApproxCount = 0 Then Exit Sub
        
        If .Columns("NoRegisterAsset") = "0" Then
        
        If MsgBox("Nomor Register barang 0 / belum register, lanjutkan", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then .Visible = False: txtNamaBarang.SetFocus: Exit Sub
        
        End If
        txtStok.Text = .Columns("jmlStok")
        txtKdBarang.Text = .Columns("KdBarang")
      
        dcAsalBarang.BoundText = .Columns("KdAsal")
        txtregister.Text = .Columns("NoRegisterAsset")
        txtNamaBarang.Text = .Columns("Nama Barang")
        .Visible = False
    End With
        
'    Call msubRecFO(rs, "select dbo.FB_TakeStokBrgNonMedis('" & mstrKdRuangan & "', '" & txtKdBarang & "','" & dcAsalBarang.BoundText & "') as stok")
'    If rs.EOF = False Then txtStok.Text = rs(0).Value Else txtStok.Text = 0
    
    dcAsalBarang.SetFocus

Exit Sub
Errload:
End Sub

Private Sub dgCariBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgCariBarang_DblClick
End Sub

Private Sub dgMutasiBarang_Click()
WheelHook.WheelUnHook
        Set MyProperty = dgMutasiBarang
        WheelHook.WheelHook dgMutasiBarang
End Sub

Private Sub dgMutasiBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaBarang.SetFocus
End Sub

Private Sub dgMutasiBarang_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Errload

    With dgMutasiBarang
        If .ApproxCount = 0 Then Exit Sub
        tempbolEdit = True
        txtKdBarang.Text = .Columns("KdBarang")
        txtNamaBarang.Text = .Columns("Nama Barang")
        dcAsalBarang.BoundText = .Columns("KdAsal")
        txtQtyBarang.Text = .Columns("QtyBarang")
        txtNoKirim.Text = .Columns("NoKirim")
        dtpTglKirim.Value = .Columns("TglKirim")
        dcRuanganPenerima.BoundText = .Columns("KdRuanganTujuan")
'        txtStok.Text
        dcKondisiBarang.Text = .Columns("Kondisi")
        txtNamaPenerima.Text = .Columns("UserPenerima")
        txtKdUserPenerima.Text = .Columns("IdUserPenerima")
        dcKondisiBarang.BoundText = .Columns("Kdkondisi")
        txtmutasi.Text = .Columns("AlasanMutasi")
        txtKeterangan.Text = .Columns("KeteranganLainnya")
        dcverifikator.BoundText = .Columns("idVerifikator")
         dgCariBarang.Visible = False
         dgNamaPenerima.Visible = False
        If .Columns("NoRegister") = "" Then
           txtregister.Text = ""
        Else
           txtregister.Text = .Columns("NoRegister")
        End If
'        txtregister.Text = .Columns("NoRegister")
        dgCariBarang.Visible = False
        dgNamaPenerima.Visible = False
    End With
    Call msubRecFO(rs, "select dbo.FB_TakeStokBrgNonMedis('" & mstrKdRuangan & "', '" & txtKdBarang & "','" & dcAsalBarang.BoundText & "') as stok")
    If rs.EOF = False Then txtStok.Text = rs(0).Value Else txtStok.Text = CDbl(txtQtyBarang.Text)
    
   
    lblJmlData.Caption = dgMutasiBarang.Bookmark & " / " & dgMutasiBarang.ApproxCount & " Data"

Exit Sub
Errload:
End Sub

Private Sub dgNamaPenerima_Click()
WheelHook.WheelUnHook
        Set MyProperty = dgNamaPenerima
        WheelHook.WheelHook dgNamaPenerima
End Sub

Private Sub dgNamaPenerima_DblClick()
On Error GoTo Errload
    If dgNamaPenerima.ApproxCount = 0 Then Exit Sub
    txtKdUserPenerima.Text = dgNamaPenerima.Columns("IdPegawai").Value
    txtNamaPenerima.Text = dgNamaPenerima.Columns("Nama Pemeriksa").Value
    substrKdPegawai = dgNamaPenerima.Columns("IdPegawai").Value
    dgNamaPenerima.Visible = False
    dcverifikator.SetFocus
Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dgNamaPenerima_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgNamaPenerima_DblClick
End Sub

Private Sub dtpTglKirim_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcRuanganPenerima.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo Errload
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call subKosong
    Call subLoadDcSource
    Call subLoadGridSource
    
   
Exit Sub
Errload:
End Sub

Private Sub txtCariBarang_Change()
On Error GoTo Errload

    Call subLoadGridSource

Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub txtLokasi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then dgMutasiBarang.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtmutasi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub txtqtyBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcKondisiBarang.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtQtyBarang_LostFocus()
    txtQtyBarang.Text = IIf(Val(txtQtyBarang) = 0, 0, Format(txtQtyBarang, "#,###"))
End Sub

Private Sub txtLokasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtNamaBarang_Change()
On Error GoTo Errload

    If tempbolTampil = True Then Exit Sub
    Call subCariBarang

Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub txtNamaBarang_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyDown Then If dgCariBarang.Visible = True Then dgCariBarang.SetFocus
End Sub

Private Sub txtNamaBarang_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then If dgCariBarang.Visible = True Then dgCariBarang.SetFocus Else dcAsalBarang.SetFocus
    If KeyAscii = 27 Then If dgCariBarang.Visible = True Then dgCariBarang.Visible = False
End Sub

Private Sub subKosong()
    txtNoKirim.Text = ""
    txtregister.Text = ""
    txtKdUserPenerima.Text = ""
    dtpTglKirim.Value = Now
    dcRuanganPenerima.Text = ""
    txtNamaPenerima.Text = ""
    dgNamaPenerima.Visible = False
    
    txtKdBarang.Text = ""
    txtNamaBarang.Text = ""
    txtCariBarang.Text = ""
    dcAsalBarang.BoundText = ""
    dcKondisiBarang.BoundText = ""
    txtQtyBarang.Text = 0
    txtmutasi.Text = ""
    txtKeterangan.Text = ""
    dcverifikator.BoundText = ""
    txtStok.Text = 0
    dgCariBarang.Visible = False
End Sub

Private Sub subLoadDcSource()
On Error GoTo Errload

    Call msubDcSource(dcAsalBarang, rs, "SELECT KdAsal, NamaAsal FROM AsalBarang where StatusEnabled='1' ORDER BY NamaAsal")
    Call msubDcSource(dcRuanganPenerima, rs, "SELECT KdRuangan, NamaRuangan FROM Ruangan where StatusEnabled='1' ORDER BY NamaRuangan")
    Call msubDcSource(dcKondisiBarang, rs, "SELECT KdKondisi, Kondisi FROM KondisiBarang where StatusEnabled='1' ORDER BY Kondisi")
    Call msubDcSource(dcverifikator, rs, "SELECT idpegawai, namalengkap FROM datapegawai order by namalengkap")
    
Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub subCariBarang()
On Error GoTo Errload

'    strsql = "SELECT  [Nama Barang], Satuan, DetailJenisBarang AS [Jenis Barang], KdBarang, KdAsal,AsalBarang, NoRegisterAsset,jmlStok FROM V_CariBarangNonMedisx" & _
'        " WHERE kdruangan='" & mstrKdRuangan & "' AND [Nama Barang] LIKE '%" & txtNamaBarang.Text & "%' AND NoRegisterAsset <> '' AND NoRegisterAsset <> '0' AND NoRegisterAsset <> '000000' AND NoRegisterAsset <> '0000000' " & _
'        " ORDER BY [Nama Barang]"

'    strsql = "SELECT  [Nama Barang], Satuan, DetailJenisBarang AS [Jenis Barang], KdBarang, KdAsal,AsalBarang, NoRegisterAsset,jmlStok FROM V_CariBarangNonMedis" & _
'        " WHERE ((NoRegisterAsset <> '0000000' AND NoRegisterAsset <> '0' AND NoRegisterAsset <>'000000') OR KdJenisAset is null) And kdruangan='" & mstrKdRuangan & "' AND [Nama Barang] LIKE '%" & txtNamaBarang.Text & "%' " & _
'        " ORDER BY [Nama Barang]"
    strsql = "SELECT  [Nama Barang], Satuan, DetailJenisBarang AS [Jenis Barang], KdBarang, KdAsal,AsalBarang, NoRegisterAsset,jmlStok FROM V_CariBarangNonMedis" & _
        " WHERE kdruangan='" & mstrKdRuangan & "' AND [Nama Barang] LIKE '%" & txtNamaBarang.Text & "%' " & _
        " ORDER BY [Nama Barang]"
                
    Call msubRecFO(rs, strsql)
    Set dgCariBarang.DataSource = rs
    With dgCariBarang
        .Columns("Nama Barang").Width = 2900
        .Columns("Satuan").Width = 1000
        .Columns("Jenis Barang").Width = 1440
        .Columns("KdBarang").Width = 0
        .Columns("AsalBarang").Width = 800
        .Columns("jmlStok").Width = 800
        .Columns("KdAsal").Width = 0
        .Columns("NoRegisterAsset").Width = 1900
        
        .Height = 2390
        .Top = 2900
        .Left = 1560
        .Visible = True
    End With

Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()
On Error GoTo Errload
Dim i As Integer

    tempbolTampil = True
    strsql = "SELECT * " & _
        " FROM V_MutasiBarangNonMedis " & _
        " WHERE kdruangan='" & mstrKdRuangan & "' AND [Nama Barang] LIKE '%" & txtCariBarang & "%'"
    Call msubRecFO(rs, strsql)
    Set dgMutasiBarang.DataSource = rs
    With dgMutasiBarang
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i

        .Columns("Nama Barang").Width = 2200
        .Columns("Asal").Width = 1000

        .Columns("QtyBarang").Width = 1500
        .Columns("RuanganTujuan").Width = 2000
        .Columns("Kondisi").Width = 1500
        
        .Columns("AlasanMutasi").Width = 2000
        .Columns("KeteranganLainnya").Width = 2000
        .Columns("NamaVerifikator").Width = 1500
        
        .Columns("NoRegister").Width = 2200
    End With
    lblJmlData.Caption = 0 & " / " & dgMutasiBarang.ApproxCount & " Data"
    tempbolTampil = False
    
Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Function sp_StrukKirim() As Boolean
On Error GoTo Errload
    sp_StrukKirim = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoKirim", adChar, adParamInput, 10, txtNoKirim.Text)
        .Parameters.Append .CreateParameter("TglKirim", adDate, adParamInput, , Format(dtpTglKirim.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, IIf(substrNoOrder = "", Null, substrNoOrder))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, dcRuanganPenerima.BoundText)
        .Parameters.Append .CreateParameter("IdUserPenerima", adChar, adParamInput, 10, txtKdUserPenerima.Text)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputNoKirim", adChar, adParamOutput, 10, Null)
    
        .ActiveConnection = dbConn
        .CommandText = "Add_StrukKirim"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data struk kirim antar ruangan", vbCritical, "Validasi"
            sp_StrukKirim = False
        Else
            txtNoKirim.Text = .Parameters("OutputNoKirim").Value
        End If
    End With
Exit Function
Errload:
    Call msubPesanError
    sp_StrukKirim = False
End Function

Private Function sp_MutasiBarangNonMedis(f_Status) As Boolean
On Error GoTo Errload

    sp_MutasiBarangNonMedis = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoKirim", adChar, adParamInput, 10, txtNoKirim.Text)
        .Parameters.Append .CreateParameter("NoRegister", adChar, adParamInput, 15, txtregister.Text)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, txtKdBarang.Text)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, dcAsalBarang.BoundText)
        .Parameters.Append .CreateParameter("QtyBarang", adInteger, adParamInput, , txtQtyBarang)
        .Parameters.Append .CreateParameter("KdKondisi", adChar, adParamInput, 2, dcKondisiBarang.BoundText)
        .Parameters.Append .CreateParameter("AlasanMutasi", adVarChar, adParamInput, 100, txtmutasi.Text)
        .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 100, txtKeterangan.Text)
        .Parameters.Append .CreateParameter("idVerifikator", adChar, adParamInput, 10, dcverifikator.BoundText)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
    
        .ActiveConnection = dbConn
        .CommandText = "AUD_MutasiBarangNonMedis"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_MutasiBarangNonMedis = False
        End If
    End With

Exit Function
Errload:
    Call msubPesanError
End Function

Private Function sp_StockBarang(f_JmlStok As Double) As Boolean
On Error GoTo Errload

    sp_StockBarang = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, txtKdBarang.Text)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, dcAsalBarang.BoundText)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("JmlMinimum", adDouble, adParamInput, , 1)
        .Parameters.Append .CreateParameter("JmlStok", adDouble, adParamInput, , f_JmlStok)
        .Parameters.Append .CreateParameter("Lokasi", adVarChar, adParamInput, 12, Null)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")
        .Parameters.Append .CreateParameter("NoRegisterAsset", adVarChar, adParamInput, 15, txtregister.Text)
    
        .ActiveConnection = dbConn
        .CommandText = "AUD_StokBarangNonMedis"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_StockBarang = False
        End If
    End With

Exit Function
Errload:
    Call msubPesanError
End Function

Private Sub txtNamaPenerima_Change()
On Error GoTo Errload
Dim i As Integer

    strsql = " SELECT [Nama Pemeriksa], JK, [Jenis Pemeriksa], IdPegawai " & _
        " From V_DaftarPemeriksaPasien" & _
        " where [Nama Pemeriksa] like '" & txtNamaPenerima.Text & "%' " & _
        " ORDER BY [Nama Pemeriksa], [Jenis Pemeriksa]"
    Call msubRecFO(dbRst, strsql)
    
    Set dgNamaPenerima.DataSource = dbRst
    With dgNamaPenerima
        .Columns("Nama Pemeriksa").Width = 2000
        .Columns("JK").Width = 360
        .Columns("Jenis Pemeriksa").Width = 1500
        .Columns("IdPegawai").Width = 0
        .Columns("JK").Alignment = dbgCenter
        
        .Top = 1800
        .Left = 5880
    End With
    If dgNamaPenerima.Visible = True Then Exit Sub
    dgNamaPenerima.Visible = True

Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub txtNamaPenerima_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then If dgNamaPenerima.Visible = True Then dgNamaPenerima.SetFocus Else dcverifikator.SetFocus
    If KeyAscii = 27 Then dgNamaPenerima.Visible = False
End Sub


