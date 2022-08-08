VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLapRKonPsnPlng_Stat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kunjungan Pasien Berdasarkan Status & Jenis Pasien"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLapRKonPsnPlng_Stat.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7425
   Begin VB.Frame fraPeriode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   7335
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Group By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3960
         TabIndex        =   9
         Top             =   360
         Width           =   3135
         Begin VB.OptionButton optGroupBy 
            Caption         =   "Hari"
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
            Left            =   960
            TabIndex        =   1
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optGroupBy 
            Caption         =   "Bulan"
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
            Index           =   1
            Left            =   1920
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   390
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
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
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   60620803
         UpDown          =   -1  'True
         CurrentDate     =   38209
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   495
         Left            =   3960
         TabIndex        =   4
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   60620803
         UpDown          =   -1  'True
         CurrentDate     =   38209
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   11
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Instalasi Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1665
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   3000
      Width           =   7335
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   5490
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Spreadsheet"
         Height          =   495
         Left            =   3825
         TabIndex        =   5
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   -2760
      Picture         =   "frmLapRKonPsnPlng_Stat.frx":08CA
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "frmLapRKonPsnPlng_Stat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub optGroupBy_Click(Index As Integer)
If optGroupBy(0).Value = True Then
   dtpAwal.CustomFormat = "dd MMMM yyyy 00:00:00"
   dtpAkhir.CustomFormat = "dd MMMM yyyy 23:59:59"
ElseIf optGroupBy(1).Value = True Then
  dtpAwal.CustomFormat = "MMMM yyyy 00:00:00"
  dtpAkhir.CustomFormat = "MMMM yyyy 23:59:59"
End If
End Sub
Private Sub cmdCetak_Click()
    If Periksa("datacombo", dcInstalasi, "Data instalasi kosong") = False Then Exit Sub
    cmdCetak.Enabled = False
    mstrInstalasi = dcInstalasi.BoundText

'**********************************************
' memfilter laprekap SR oleh option box
'**********************************************
    If optGroupBy(0).Value = True Then
    '**Jika Perhari >>
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
        If strCetak = "laprekapkpps" Then
        strSQL = "SELECT * FROM   V_RekapitulasiKunjunganPasienBKondisipulangdanStatus " _
        & "WHERE (TglKeluar BETWEEN '" _
        & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') and KdInstalasi='" & dcInstalasi.BoundText & "' "
        mstrCetak2 = "laprekapkppshr"
        End If
    '**Jika B Bulan >>
    ElseIf optGroupBy(1).Value = True Then
         If strCetak = "laprekapkpps" Then
        strSQL = "SELECT {fn MONTHNAME (TglKeluar) } AS TglKeluar, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM  V_RekapitulasiKunjunganPasienBKondisipulangdanStatus " _
        & "WHERE (TglKeluar BETWEEN '" _
        & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') and KdInstalasi='" & dcInstalasi.BoundText & "' "
        mstrCetak2 = "laprekapkppsbln"
        End If
    End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    msubRecFO rs, strSQL
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbExclamation, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If

    Set frmCtkLapRekap_Viewer = Nothing
'    frmCtkLapRekap_Viewer.Show
    If mstrCetak2 = "laprekapkppshr" Then
        frmCtkLapRekap_Viewer.Caption = "Medifirst2000 - Laporan Rekapitulasi Pasien Per Status & Rujukan (hari)"
    ElseIf mstrCetak2 = "laprekapkppsbln" Then
        frmCtkLapRekap_Viewer.Caption = "Medifirst2000 - Laporan Rekapitulasi Pasien Per Status & Rujukan (bulan)"
    Else
        
    End If
    cmdCetak.Enabled = True
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then optGroupBy(0).SetFocus
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    With Me
        .dtpAwal.Value = Now
        .dtpAkhir.Value = Now
    End With
    Call subDcSource
End Sub

Private Sub optGroupBy_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub


'user defined procedure(s) & function(s)>>
Private Sub subDcSource()
    If strCetak = "laprekapkpps" Then
        strSQL = "SELECT KdInstalasi, NamaInstalasi " & _
            " From instalasi" & _
            " WHERE (KdInstalasi IN ('01', '02', '03', '04', '06', '08', '09', '10', '16'))"
        Call msubDcSource(dcInstalasi, rs, strSQL)
        If rs.EOF = False Then dcInstalasi.BoundText = rs(0)
    Else
    End If
End Sub





