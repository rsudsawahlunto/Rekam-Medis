VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMasterLaporan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Laporan"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMasterLaporan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9090
   Begin MSDataListLib.DataCombo dcJenisPemeriksaan 
      Height          =   330
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   9015
      Begin MSComctlLib.ListView lvwPelayanan 
         Height          =   4455
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nama Diagnosa"
            Object.Width           =   13229
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Jenis Pemeriksaan"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frmMasterLaporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Set rs = Nothing
    Call msubDcSource(dcJenisPemeriksaan, rs, "SELECT DISTINCT NomorUrut,JenisPemeriksaan FROM MasterRekapRadiologi ORDER BY NomorUrut")
End Sub
