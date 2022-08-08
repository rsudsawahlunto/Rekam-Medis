VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash32_11_2_202_228.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInfoPesanPelayananTMOA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Informasi Pesan Pelayanan Ruangan"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInfoPesanPelayananTMOA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   13875
   Begin VB.Frame Frame3 
      Caption         =   "Informasi Pesan Pelayanan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   13815
      Begin VB.Frame Frame5 
         Caption         =   "Pelayanan"
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
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   4335
         Begin VB.OptionButton optTM 
            Caption         =   "Tindakan Medis"
            Height          =   375
            Left            =   480
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optOA 
            Caption         =   "Obat Alkes"
            Height          =   375
            Left            =   2400
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status"
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
         Left            =   4560
         TabIndex        =   10
         Top             =   360
         Width           =   3255
         Begin VB.OptionButton optSudah 
            Caption         =   "Sudah"
            Height          =   375
            Left            =   2040
            TabIndex        =   12
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optBelum 
            Caption         =   "Belum"
            Height          =   375
            Left            =   480
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
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
         Left            =   7920
         TabIndex        =   5
         Top             =   360
         Width           =   5775
         Begin VB.CommandButton cmdTampilkan 
            Caption         =   "&Cari"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpTglAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   7
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   16515075
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpTglAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   8
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   16515075
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   9
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgInfoPesanPelayanan 
         Height          =   5535
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   9763
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
               LCID            =   1033
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
               LCID            =   1033
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
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   8040
      Width           =   13815
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   520
         Left            =   10560
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   520
         Left            =   12240
         TabIndex        =   2
         Top             =   240
         Width           =   1455
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
      Left            =   12000
      Picture         =   "frmInfoPesanPelayananTMOA.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmInfoPesanPelayananTMOA.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmInfoPesanPelayananTMOA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTampilkan_Click()
    Call subLoadData
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgInfoPesanPelayanan_Click()
WheelHook.WheelUnHook
        Set MyProperty = dgInfoPesanPelayanan
        WheelHook.WheelHook dgInfoPesanPelayanan
End Sub

Private Sub dtpTglAkhir_Change()
    dtpTglAkhir.MaxDate = Now
End Sub

Private Sub dtpTglAwal_Change()
    dtpTglAwal.MaxDate = Now
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    optBelum.Value = True
    optTM.Value = True
    dtpTglAkhir.Value = Now
    dtpTglAwal.Value = Now
    Call subLoadData
    
     
End Sub

Sub subLoadData()
On Error GoTo hell
Dim i As Integer

    If optTM.Value = True Then
        If optBelum.Value = True Then
            strSQL = "Select NamaPelayanan,JmlPelayanan,StatusCito,NoPendaftaran,NoCM,NamaPasien,RuanganTujuan,DokterOrder,UserOrder from V_DaftarDetailOrderTM where TglOrder Between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "' and KdRuangan='" & mstrKdRuangan & "' and NoRiwayat is null"
        ElseIf optSudah.Value = True Then
            strSQL = "Select NamaPelayanan,JmlPelayanan,StatusCito,NoPendaftaran,NoCM,NamaPasien,RuanganTujuan,DokterOrder,UserOrder from V_DaftarDetailOrderTM where TglOrder Between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "' and KdRuangan='" & mstrKdRuangan & "' and NoRiwayat is not null"
        End If
    ElseIf optOA.Value = True Then
        If optBelum.Value = True Then
            strSQL = "Select JenisBarang,NamaBarang,NamaAsal,JmlBarang,HargaFIFO,Satuan,NoPendaftaran,NoCM,NamaPasien,RuanganTujuan,DokterOrder,UserOrder from V_DaftarDetailOrderOA where TglOrder Between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "' and KdRuangan='" & mstrKdRuangan & "' and NoRiwayat is null"
        ElseIf optSudah.Value = True Then
            strSQL = "Select JenisBarang,NamaBarang,NamaAsal,JmlBarang,HargaFIFO,Satuan,NoPendaftaran,NoCM,NamaPasien,RuanganTujuan,DokterOrder,UserOrder from V_DaftarDetailOrderOA where TglOrder Between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "' and KdRuangan='" & mstrKdRuangan & "' and NoRiwayat is not null"
        End If
    End If
    
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    
    Set dgInfoPesanPelayanan.DataSource = rs
    
    With dgInfoPesanPelayanan
'        For i = 0 To .Columns.Count - 1
'            .Columns(i).Width = 0
'        Next i
        
        If optTM.Value = True Then
            .Columns("NamaPelayanan").Width = 2000
            .Columns("JmlPelayanan").Width = 800
            .Columns("JmlPelayanan").Caption = "Jumlah"
''            .Columns("BiayaSatuan").Width = 1500
''            .Columns("BiayaSatuan").Alignment = dbgRight
            .Columns("StatusCito").Width = 800
            .Columns("NoPendaftaran").Width = 1350
            .Columns("NoCM").Width = 1000
            .Columns("NamaPasien").Width = 2500
            .Columns("RuanganTujuan").Width = 2200
            .Columns("DokterOrder").Width = 2000
            .Columns("UserOrder").Width = 2000
        ElseIf optOA.Value = True Then
            .Columns("JenisBarang").Width = 1500
            .Columns("NamaBarang").Width = 2000
            .Columns("NamaAsal").Width = 1500
            .Columns("JmlBarang").Width = 800
            .Columns("JmlBarang").Caption = "Jumlah"
            .Columns("HargaFIFO").Width = 1500
            .Columns("HargaFIFO").Alignment = dbgRight
            .Columns("Satuan").Width = 1000
            .Columns("NoPendaftaran").Width = 1350
            .Columns("NoCM").Width = 1000
            .Columns("NamaPasien").Width = 2500
            .Columns("RuanganTujuan").Width = 2200
            .Columns("DokterOrder").Width = 2000
            .Columns("UserOrder").Width = 2000
        End If
    End With
            
Exit Sub
hell:
    Call msubPesanError
End Sub
