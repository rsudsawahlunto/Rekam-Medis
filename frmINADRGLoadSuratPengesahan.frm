VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmINADRGLoadSuratPengesahan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Surat Pengesahan"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   Icon            =   "frmINADRGLoadSuratPengesahan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6015
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   6015
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
         Left            =   4080
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
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
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   6015
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
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   5775
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
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
            CustomFormat    =   "ddMMMMyyyy"
            Format          =   126812163
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
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
            CustomFormat    =   "ddMMMMyyyy"
            Format          =   126812163
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   1680
            TabIndex        =   6
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   8
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmINADRGLoadSuratPengesahan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   4080
      Picture         =   "frmINADRGLoadSuratPengesahan.frx":368B
      Stretch         =   -1  'True
      Top             =   15
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmINADRGLoadSuratPengesahan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmINADRGLoadSuratPengesahan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCetak_Click()
    frmINADRGSuratPengesahan.Show
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpAwal.value = Format(Now, "dd MMMM yyyy 00:00:00")
    dtpAkhir.value = Format(Now, "dd MMMM yyyy 23:59:59")
End Sub
