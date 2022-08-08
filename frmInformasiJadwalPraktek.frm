VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInformasiJadwalPraktek 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informasi Jadwal Praktek Dokter"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInformasiJadwalPraktek.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   13095
   Begin VB.CommandButton cmdTutup 
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   11160
      TabIndex        =   3
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   13095
      Begin VB.Frame Frame2 
         Caption         =   "Tanggal"
         Height          =   855
         Left            =   5040
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   7935
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   4920
            TabIndex        =   11
            Top             =   360
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMMM yyyy HH:mm"
            Format          =   118358019
            UpDown          =   -1  'True
            CurrentDate     =   41505
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   1680
            TabIndex        =   10
            Top             =   360
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMMM yyyy HH:mm"
            Format          =   118358019
            UpDown          =   -1  'True
            CurrentDate     =   41505
         End
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4560
            TabIndex        =   14
            Top             =   480
            Width           =   255
         End
      End
      Begin VB.ComboBox cboHari 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "frmInformasiJadwalPraktek.frx":0CCA
         Left            =   120
         List            =   "frmInformasiJadwalPraktek.frx":0CE3
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid fgJadwalDokter 
         Height          =   4215
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   7435
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcStatus 
         Height          =   315
         Left            =   5400
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   315
         Left            =   5040
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Status Kehadiran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Ruangan"
         Height          =   255
         Left            =   5040
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Hari"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1815
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
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmInformasiJadwalPraktek.frx":0D19
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmInformasiJadwalPraktek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboHari_Change()
    Call subSetGrid
    Call Isi
    fgJadwalDokter.Refresh
End Sub
Private Sub cboHari_GotFocus()
    Call subSetGrid
    Call Isi
    fgJadwalDokter.Refresh
End Sub

Private Sub cboHari_KeyPress(KeyAscii As Integer)
    Call subSetGrid
    Call Isi
    fgJadwalDokter.Refresh
End Sub

Private Sub cboHari_LostFocus()
    Call subSetGrid
    Call Isi
End Sub

Private Sub cmdCari_Click()
    Call subSetGrid
    Call Isi
    fgJadwalDokter.Refresh
End Sub

Private Sub cmdTutup_Click()
Unload Me
End Sub


Private Sub Form_Load()
On Error GoTo gabril
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)

    dtpAkhir.value = Now
    
    Call subSetGrid
    Call Isi
    fgJadwalDokter.Refresh
    
Exit Sub
gabril:
    Call msubPesanError
End Sub

Public Sub subSetGrid()
On Error GoTo errLoad
    With fgJadwalDokter
        .Clear
        .Rows = 2
        .Cols = 9
        
        .MergeCells = flexMergeFree
        
        .RowHeight(0) = 500
        
        
        
        
        .TextMatrix(0, 0) = "Kode Praktek"
        .TextMatrix(0, 1) = "Hari"
        .TextMatrix(0, 2) = "Tanggal"
        .TextMatrix(0, 3) = "Nama Dokter"
        .TextMatrix(0, 4) = "Nama Ruangan"
        .TextMatrix(0, 5) = "Jam Mulai"
        .TextMatrix(0, 6) = "Jam Selesai"
        .TextMatrix(0, 7) = "Status Hadir"
        .TextMatrix(0, 8) = "Keterangan"
'        .TextMatrix(0, 9) = "Keterangan"
        
    
        .ColWidth(0) = 0
        .ColWidth(1) = 1300
        .ColWidth(2) = 0
        .ColWidth(3) = 3400
        .ColWidth(4) = 1800
        .ColWidth(5) = 1500
        .ColWidth(6) = 1500
        .ColWidth(7) = 1500
        .ColWidth(8) = 1800
'        .ColWidth(9) = 1700
        
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
    
        .MergeCol(1) = True
        .MergeCol(2) = True
        
        

    End With

Exit Sub
errLoad:
    Call msubPesanError
End Sub
Public Sub Isi()
On Error GoTo gabril
Set rs = Nothing
strSQL = ""
strSQL = "select KdPraktek,Hari,Tgl,NamaLengkap,NamaRuangan,JamMulai,JamSelesai,StatusHadir,Keterangan from V_JadwalPraktekDokter " & _
         "where Hari like '%" & cboHari.Text & "%' order by Hari"
Call msubRecFO(rs, strSQL)
If rs.RecordCount <> 0 Then
    fgJadwalDokter.Rows = rs.RecordCount + 1
     For i = 1 To rs.RecordCount
        With fgJadwalDokter
            .TextMatrix(i, 0) = IIf(IsNull(rs.Fields(0).value), "-", rs.Fields(0))  '
            .TextMatrix(i, 1) = IIf(IsNull(rs.Fields(1).value), "-", rs.Fields(1))  '
            .TextMatrix(i, 2) = IIf(IsNull(rs.Fields(2).value), "-", rs.Fields(2))  '
            .TextMatrix(i, 3) = IIf(IsNull(rs.Fields(3).value), "-", rs.Fields(3))  '
            .TextMatrix(i, 4) = IIf(IsNull(rs.Fields(4).value), "-", rs.Fields(4))  '
            .Row = i
            .Col = 3
            .CellFontBold = True
            .TextMatrix(i, 5) = IIf(IsNull(rs.Fields(5).value), "-", rs.Fields(5))  '
            .TextMatrix(i, 6) = IIf(IsNull(rs.Fields(6).value), "-", rs.Fields(6))  '
            .TextMatrix(i, 7) = IIf(IsNull(rs.Fields(7).value), "-", rs.Fields(7))  '
            .TextMatrix(i, 8) = IIf(IsNull(rs.Fields(8).value), "-", rs.Fields(8))  '
'            .TextMatrix(i, 9) = IIf(IsNull(rs.Fields(9).value), "-", rs.Fields(8))  '
        End With
        rs.MoveNext
     Next i
End If
Exit Sub
gabril:
    Call msubPesanError
End Sub

