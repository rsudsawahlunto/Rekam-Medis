VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9f.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPelayananResep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - PenulisanDanPelayananResep"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPelayananResep.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   8625
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
      TabIndex        =   3
      Top             =   1080
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
         Left            =   1920
         TabIndex        =   4
         Top             =   150
         Width           =   5055
         Begin MSComCtl2.DTPicker DTPickerAwal 
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy "
            Format          =   58785795
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker DTPickerAkhir 
            Height          =   375
            Left            =   2760
            TabIndex        =   6
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   58785795
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   2400
            TabIndex        =   7
            Top             =   315
            Width           =   255
         End
      End
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   2280
      Width           =   1665
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   2
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
      Picture         =   "frmPelayananResep.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPelayananResep.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPelayananResep.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "frmPelayananResep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCetak_Click()
On Error GoTo errLoad
'
'Set rs = Nothing
'strSQL1 = "select Jumlah = count(a.kdBarang)" & _
'         " from Masterbarang a" & _
'         " Inner join KategoryBarang b on a.kdKategoryBarang = b.kdKategoryBarang" & _
'         " where a.kdKategoryBarang = '01' and a.StatusAktif = 'Y'"
'Call msubRecFO(rs, strSQL1)
'iJml1 = rs.Fields("Jumlah")
'
''--------------------------------------------------------------------------------
'Set rs = Nothing
'strSQL2 = "select Jumlah = count(a.kdBarang)" & _
'          " from Masterbarang a" & _
'          " Inner join KategoryBarang b on a.kdKategoryBarang = b.kdKategoryBarang" & _
'          " inner join StatusBarang c on a.kdStatusBarang = c.kdStatusBarang" & _
'          " where a.kdKategoryBarang = '02' and a.kdStatusBarang = '01' and a.StatusAktif = 'Y'"
'
'Call msubRecFO(rs, strSQL2)
'iJml2 = rs.Fields("Jumlah")
'
''----------------------------------------------------------------------------------
'Set rs = Nothing
'strSQL3 = "select Jumlah = count(a.kdBarang)" & _
'          " from Masterbarang a" & _
'          " left outer join KategoryBarang b on a.kdKategoryBarang = b.kdKategoryBarang" & _
'          " left outer join StatusBarang c on a.kdStatusBarang = c.kdStatusBarang" & _
'          " where a.kdKategoryBarang = '02' and a.StatusAktif = 'Y'"
'
'Call msubRecFO(rs, strSQL3)
'iJml3 = rs.Fields("Jumlah")
'
''---------------------------------------------------------------------------------------
'Set rs = Nothing
'strSQL4 = "select  Jumlah = sum(a.JmlStok)" & _
'          " from StokRuangan a" & _
'          " Inner Join(select a.kdBarang from Masterbarang a" & _
'          " Inner join KategoryBarang b on a.kdKategoryBarang = b.kdKategoryBarang" & _
'          " where a.kdKategoryBarang = '01' and a.StatusAktif = 'Y') b on a.kdBarang = b.kdBarang"
'
'Call msubRecFO(rs, strSQL4)
'cJml4 = rs.Fields("Jumlah")
'
''--------------------------------------------------------------------------------------
'Set rs = Nothing
'strSQL5 = "select  Jumlah = sum(a.JmlStok)" & _
'          " from StokRuangan a" & _
'          " Inner Join(select a.kdBarang from Masterbarang a" & _
'          " Inner join KategoryBarang b on a.kdKategoryBarang = b.kdKategoryBarang" & _
'          " inner join StatusBarang c on a.kdStatusBarang = c.kdStatusBarang" & _
'          " where a.kdKategoryBarang = '02' and a.kdStatusBarang = '01' and a.StatusAktif = 'Y') b on a.kdBarang = b.kdBarang"
'
'Call msubRecFO(rs, strSQL5)
'cJml5 = rs.Fields("Jumlah")
'
''---------------------------------------------------------------------------------------
'Set rs = Nothing
'strSQL6 = "select  Jumlah = sum(a.JmlStok)" & _
'          " from StokRuangan a" & _
'          " Inner Join(select a.kdBarang from Masterbarang a" & _
'          " left outer join KategoryBarang b on a.kdKategoryBarang = b.kdKategoryBarang" & _
'          " left outer join StatusBarang c on a.kdStatusBarang = c.kdStatusBarang" & _
'          " where a.kdKategoryBarang = '02' and a.StatusAktif = 'Y') b on a.kdBarang = b.kdBarang"
'
'Call msubRecFO(rs, strSQL6)
'cJml6 = rs.Fields("Jumlah")
'
''-------------------------------------------------------------------------------------------------
'Set rs = Nothing
'strSQL7 = "select  Jumlah = sum(a.JmlStok)" & _
'          " from StokRuangan a" & _
'          " Inner Join(select a.kdBarang from Masterbarang a" & _
'          " Inner join KategoryBarang b on a.kdKategoryBarang = b.kdKategoryBarang" & _
'          " left outer join StatusBarang c on a.kdStatusBarang = c.kdStatusBarang" & _
'          " where a.kdKategoryBarang = '01' and a.kdStatusBarang = '01' and a.StatusAktif = 'Y') b on a.kdBarang = b.kdBarang"
'
' Call msubRecFO(rs, strSQL7)
' cJml7 = rs.Fields("Jumlah")
'
''-------------------------------------------------------------------------------------
'Set rs = Nothing
'strSQL8 = "select  Jumlah = sum(a.JmlStok)" & _
'          " from StokRuangan a" & _
'          " Inner Join(select a.kdBarang from Masterbarang a" & _
'          " Inner join KategoryBarang b on a.kdKategoryBarang = b.kdKategoryBarang" & _
'          " inner join StatusBarang c on a.kdStatusBarang = c.kdStatusBarang" & _
'          " where a.kdKategoryBarang = '02' and a.kdStatusBarang = '01' and a.StatusAktif = 'Y') b on a.kdBarang = b.kdBarang"
'
'Call msubRecFO(rs, strSQL8)
'cJml8 = rs.Fields("Jumlah")
''---------------------------------------------------------------------------------------
'
''---------------------------------------------------------------------------------------
'
'Set rs = Nothing
'strSQL9 = "select  Jumlah = sum(a.JmlStok)" & _
'          " from StokRuangan a " & _
'          " Inner Join(select a.kdBarang from Masterbarang a" & _
'          " left outer join KategoryBarang b on a.kdKategoryBarang = b.kdKategoryBarang" & _
'          " left outer join StatusBarang c on a.kdStatusBarang = c.kdStatusBarang" & _
'          " where a.kdKategoryBarang = '02' and a.kdStatusBarang = '01' and a.StatusAktif = 'Y') b on a.kdBarang = b.kdBarang"
'
'Call msubRecFO(rs, strSQL9)
'cJml9 = rs.Fields("Jumlah")
''-------------------------------------------------------------------------------------------
' Call msubRecFO(dbRst, strSQL1)
' Call msubRecFO(dbRst, strSQL2)
' Call msubRecFO(dbRst, strSQL3)
' Call msubRecFO(dbRst, strSQL4)
' Call msubRecFO(dbRst, strSQL5)
' Call msubRecFO(dbRst, strSQL6)
' Call msubRecFO(dbRst, strSQL7)
' Call msubRecFO(dbRst, strSQL8)
' Call msubRecFO(dbRst, strSQL9)


 frm_cetak_PelayananResepRS.Show
Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub
'
'Private Sub DTPickerAkhir_Change()
'    DTPickerAkhir.MaxDate = Now
'End Sub

'Private Sub DTPickerAwal_Change()
'    DTPickerAwal.MaxDate = Now
'End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    With Me
        .DTPickerAwal.Value = Format(Now, "dd MMM yyyy 00:00:00")
        .DTPickerAkhir.Value = Now
    End With
   ' strSQL = "SELECT KdInstalasi, NamaInstalasi FROM Instalasi WHERE KdInstalasi NOT IN ('05','07','13','14','15','17','18','19','20','21','23')"
   ' Call msubDcSource(dcInstalasi, dbRst, strSQL)
Exit Sub
errLoad:
    Call msubPesanError
End Sub


