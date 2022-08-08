VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPeriodeKunjungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Periode Kunjungan Pasien"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPeriodeKunjungan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5280
   Begin VB.Frame Frame3 
      Caption         =   "Instalasi Pelayanan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Width           =   5295
      Begin VB.ComboBox CboInstalasi 
         Height          =   330
         Left            =   240
         TabIndex        =   10
         Top             =   570
         Width           =   4815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Instalasi Pelayanan"
         Height          =   210
         Index           =   1
         Left            =   270
         TabIndex        =   11
         Top             =   300
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   3720
      Width           =   5295
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   2880
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   600
         TabIndex        =   7
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
      TabIndex        =   0
      Top             =   1320
      Width           =   5295
      Begin MSComCtl2.DTPicker DTPickerAwal 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   22609923
         CurrentDate     =   38177
      End
      Begin MSComCtl2.DTPicker DTPickerAkhir 
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   22609923
         CurrentDate     =   38177
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
         TabIndex        =   5
         Top             =   675
         Width           =   270
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
         TabIndex        =   2
         Top             =   360
         Width           =   1200
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
         TabIndex        =   1
         Top             =   360
         Width           =   1170
      End
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   0
      Picture         =   "frmPeriodeKunjungan.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmPeriodeKunjungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCetak_Click()
'Select Case cetak
'Case "Morbid"
kdinstalasi = Left(frmPeriodeKunjungan.CboInstalasi, 2)
If Left(frmPeriodeKunjungan.CboInstalasi, 2) = "01" Then
    strSQL = "SELECT NoDTD,NoDTerperinci,NamaDTD,SUM(Kel_Umur1) AS Kel_Umur1,SUM(Kel_Umur2) AS Kel_Umur2,SUM(Kel_Umur3) AS Kel_Umur3,SUM(Kel_Umur4) AS Kel_Umur4,SUM(Kel_Umur5) AS Kel_Umur5,SUM(Kel_Umur6) AS Kel_Umur6,SUM(Kel_Umur7) AS Kel_Umur7,SUM(Kel_Umur8) AS Kel_Umur8,SUM(Kel_L) AS Kel_L,SUM(Kel_P) AS Kel_P,SUM(Kel_Kunj) AS Kel_Kunj " _
        & "FROM v_S_RekapMorbidRJ " _
        & "WHERE TglPeriksa BETWEEN '" & Format(frmPeriodeKunjungan.DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmPeriodeKunjungan.DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " _
         & "and kdinstalasi = '01' " _
        & "GROUP BY NoDTD,NoDTerperinci,NamaDTD"
        
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    frmMorbiditasRJ.Show
    frmMorbiditasRJ.Caption = "Medifirst2000 - Data Surveilans Morbiditas Pasien Gawat Darurat"
Else
If Left(frmPeriodeKunjungan.CboInstalasi, 2) = "02" Then
    strSQL = "SELECT NoDTD,NoDTerperinci,NamaDTD,SUM(Kel_Umur1) AS Kel_Umur1,SUM(Kel_Umur2) AS Kel_Umur2,SUM(Kel_Umur3) AS Kel_Umur3,SUM(Kel_Umur4) AS Kel_Umur4,SUM(Kel_Umur5) AS Kel_Umur5,SUM(Kel_Umur6) AS Kel_Umur6,SUM(Kel_Umur7) AS Kel_Umur7,SUM(Kel_Umur8) AS Kel_Umur8,SUM(Kel_L) AS Kel_L,SUM(Kel_P) AS Kel_P,SUM(Kel_Kunj) AS Kel_Kunj " _
        & "FROM v_S_RekapMorbidRJ " _
        & "WHERE TglPeriksa BETWEEN '" & Format(frmPeriodeKunjungan.DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmPeriodeKunjungan.DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " _
         & "and kdinstalasi = '02' " _
        & "GROUP BY NoDTD,NoDTerperinci,NamaDTD"
        
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    frmMorbiditasRJ.Show
    frmMorbiditasRJ.Caption = "Medifirst2000 - Data Surveilans Morbiditas Pasien Rawat Jalan"
Else
If Left(frmPeriodeKunjungan.CboInstalasi, 2) = "03" Then
    strSQL = "SELECT NoDTD,NoDTerperinci,NamaDTD,SUM(Kel_Umur1) AS Kel_Umur1,SUM(Kel_Umur2) AS Kel_Umur2,SUM(Kel_Umur3) AS Kel_Umur3,SUM(Kel_Umur4) AS Kel_Umur4,SUM(Kel_Umur5) AS Kel_Umur5,SUM(Kel_Umur6) AS Kel_Umur6,SUM(Kel_Umur7) AS Kel_Umur7,SUM(Kel_Umur8) AS Kel_Umur8,SUM(Kel_L) AS Kel_L,SUM(Kel_P) AS Kel_P,SUM(Kel_H) AS Kel_H,SUM(Kel_M) AS Kel_M " _
        & "FROM v_S_RekapMorbidRI " _
        & "WHERE TglPeriksa BETWEEN '" & Format(frmPeriodeKunjungan.DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmPeriodeKunjungan.DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " _
        & "and kdinstalasi = '03' " _
        & "GROUP BY NoDTD,NoDTerperinci,NamaDTD"

    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    ctk = "RI"
    frmMorbiditasRI.Show
'    frmMorbiditasRI.Caption = "Medifirst2000 - Data Surveilans Morbiditas Pasien Rawat Inap"
End If
End If
End If
'Unload Me
'End Select
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
    
    Set dbRec = New ADODB.recordset
    dbRec.Open " SELECT     KdInstalasi, NamaInstalasi " _
             & " FROM         Instalasi where kdinstalasi = '01' or kdinstalasi = '02' or kdinstalasi ='03'", dbConn, adOpenDynamic, adLockOptimistic

    While dbRec.EOF = False
        CboInstalasi.AddItem dbRec.Fields(0).Value & " - " & dbRec.Fields(1).Value
        dbRec.MoveNext
    Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPeriodeKunjungan = Nothing
End Sub
