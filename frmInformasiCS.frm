VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmInformasiCS 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Informasi Current System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.Frame Frame3 
         Caption         =   "Informasi Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   7455
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2175
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   3836
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Caption         =   "4000"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3000
            TabIndex        =   11
            Top             =   3120
            Width           =   840
         End
         Begin VB.Label lblJmlPerempuan 
            Alignment       =   1  'Right Justify
            Caption         =   "4000"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3000
            TabIndex        =   10
            Top             =   2880
            Width           =   840
         End
         Begin VB.Label lblJmlLaki 
            Alignment       =   1  'Right Justify
            Caption         =   "4000"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3000
            TabIndex        =   9
            Top             =   2640
            Width           =   840
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   2760
            TabIndex        =   8
            Top             =   3120
            Width           =   90
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   2760
            TabIndex        =   7
            Top             =   2880
            Width           =   90
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   2760
            TabIndex        =   6
            Top             =   2640
            Width           =   90
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Jumlah Total"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   240
            TabIndex        =   5
            Top             =   3120
            Width           =   1590
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Jumlah Perempuan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   240
            TabIndex        =   4
            Top             =   2880
            Width           =   2340
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Jumlah Laki-Laki"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   240
            TabIndex        =   3
            Top             =   2640
            Width           =   2070
         End
      End
      Begin VB.Label lblTgl 
         AutoSize        =   -1  'True
         Caption         =   "Selasa, 31 November 2004 23:45"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   4005
      End
   End
End
Attribute VB_Name = "frmInformasiCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Call centerForm(Me, MDIUtama)
End Sub
