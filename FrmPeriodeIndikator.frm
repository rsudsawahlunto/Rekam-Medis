VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPeriodeIndikator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst22000 - Indikator Pelayanan Rumah Sakit"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   Icon            =   "FrmPeriodeIndikator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7650
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   7665
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Spreadsheet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5550
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
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
      Top             =   960
      Width           =   7635
      Begin VB.ComboBox cboKriteria 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "FrmPeriodeIndikator.frx":08CA
         Left            =   5400
         List            =   "FrmPeriodeIndikator.frx":08D7
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   1965
      End
      Begin MSComCtl2.DTPicker DTPickerAwal 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   50855939
         UpDown          =   -1  'True
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   50855939
         UpDown          =   -1  'True
         CurrentDate     =   38177
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Kriteria"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   5400
         TabIndex        =   10
         Top             =   360
         Width           =   555
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
         Caption         =   "Tanggal Akhir"
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
         Width           =   1185
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
         Index           =   0
         Left            =   2520
         TabIndex        =   6
         Top             =   675
         Width           =   270
      End
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   -2520
      Picture         =   "FrmPeriodeIndikator.frx":08FA
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "FrmPeriodeIndikator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCetak_Click()
If cboKriteria.Text = "Semua" Then
    strSQL = "SELECT AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR " _
        & "FROM RekapitulasiIndikatorPelayananRS " _
        & "WHERE TglHitung BETWEEN '" _
        & Format(FrmPeriodeIndikator.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(FrmPeriodeIndikator.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    frmUtilitasRS2.Show
'    frmUtilitasRS2.Caption = "Indikator Pelayanan RS"
    Unload Me
ElseIf (cboKriteria.Text = "Per Ruangan") Then
    strSQL = "SELECT NamaRuangan AS Ruangan,AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR " _
            & "FROM dbo.v_S_RekapIndikatorPlyn " _
            & "WHERE TglHitung BETWEEN '" & Format(FrmPeriodeIndikator.DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(FrmPeriodeIndikator.DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " _
            & "GROUP BY NamaRuangan"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    frmUtilitasRS.Show
'    frmUtilitasRS.Caption = "Indikator Pelayanan RS"
    Unload Me
ElseIf (cboKriteria.Text = "Per Kelas") Then
    strSQL = "SELECT DeskKelas AS Kelas,AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR " _
            & "FROM dbo.v_S_RekapIndikatorPlyn " _
            & "WHERE TglHitung BETWEEN '" & Format(FrmPeriodeIndikator.DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(FrmPeriodeIndikator.DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " _
            & "GROUP BY DeskKelas"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    frmUtilitasRS.Show
'    frmUtilitasRS.Caption = "Indikator Pelayanan RS"
    Unload Me
End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    With Me
        .DTPickerAwal.Value = Now
        .DTPickerAkhir.Value = Now
    End With
End Sub
