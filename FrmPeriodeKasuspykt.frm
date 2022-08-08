VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmPeriodeKasuspykt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rekapitulasi Kasus Penyakit IGD"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5295
   Begin VB.Frame Frame1 
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
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   5295
      Begin MSComCtl2.DTPicker DTPickerAwal 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   64487427
         CurrentDate     =   38177
      End
      Begin MSComCtl2.DTPicker DTPickerAkhir 
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   64487427
         CurrentDate     =   38177
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Awal"
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
         TabIndex        =   8
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal AKhir"
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
         Index           =   0
         Left            =   2880
         TabIndex        =   7
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
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
         Left            =   2520
         TabIndex        =   6
         Top             =   675
         Width           =   270
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   2490
      Width           =   5295
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   0
      Picture         =   "FrmPeriodeKasuspykt.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "FrmPeriodeKasuspykt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCetak_Click()
    strSQL = "SELECT * FROM V_RekapitulasiPasienBKasusPenyakit " _
        & "WHERE (TglPendaftaran BETWEEN '" _
        & Format(DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "') ORDER BY Ruangan,JenisPasien"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    cetak = "kasuspykt"
    FrmViewerLaporan.Show
    FrmViewerLaporan.Caption = "Rekapitulasi Kasus Penyakit IGD"
    Unload Me
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    With Me
        .DTPickerAwal.Value = Now
        .DTPickerAkhir.Value = Now
    End With
End Sub
