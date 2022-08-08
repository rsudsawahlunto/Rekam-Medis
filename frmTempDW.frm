VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTempDW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Monitoring Data"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12315
   Icon            =   "frmTempDW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   12315
   Begin VB.Frame frInfoData 
      Caption         =   "Info Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   7800
      Width           =   9975
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Data Gagal Diproses"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5880
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblGagal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000000 data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5880
         TabIndex        =   9
         Top             =   480
         Width           =   1665
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Data Sudah Diproses"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   19
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Data Belum Diproses"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   18
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   750
      End
      Begin VB.Label lblTotData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000000 data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1665
      End
      Begin VB.Label lblSudah 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000000 data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3960
         TabIndex        =   8
         Top             =   480
         Width           =   1665
      End
      Begin VB.Label lblBelum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000000 data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   480
         Width           =   1665
      End
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   10200
      TabIndex        =   10
      Top             =   8040
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   12135
      Begin VB.Frame Frame2 
         Height          =   6495
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   3255
         Begin VB.Timer tmrAutoRefresh 
            Enabled         =   0   'False
            Interval        =   3000
            Left            =   120
            Top             =   6000
         End
         Begin VB.CheckBox chkAutoRefresh 
            Appearance      =   0  'Flat
            Caption         =   "Auto Refresh"
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
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1403
         End
         Begin VB.ComboBox cmbStatus 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmTempDW.frx":0CCA
            Left            =   120
            List            =   "frmTempDW.frx":0CCC
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   1200
            Width           =   3015
         End
         Begin VB.CommandButton cmdBuka 
            Caption         =   "&Buka File"
            Height          =   495
            Left            =   1560
            TabIndex        =   4
            Top             =   2640
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dtpTglFileDAT 
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   3015
            _ExtentX        =   5318
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
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   126418947
            UpDown          =   -1  'True
            CurrentDate     =   39932
         End
         Begin VB.Frame frAutoRefresh 
            Height          =   855
            Left            =   120
            TabIndex        =   21
            Top             =   1680
            Width           =   3015
            Begin MSComCtl2.DTPicker dtpAutoRefresh 
               Height          =   375
               Left            =   240
               TabIndex        =   3
               Top             =   360
               Width           =   1095
               _ExtentX        =   1931
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
               CustomFormat    =   "mm:ss"
               Format          =   126418947
               UpDown          =   -1  'True
               CurrentDate     =   39932.0000347222
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "mm:ss"
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
               Left            =   1440
               TabIndex        =   22
               Top             =   480
               Width           =   585
            End
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Filter Status Proses Data"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   2130
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tanggal File Temporary"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   2010
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridTemp 
         Height          =   6375
         Left            =   3480
         TabIndex        =   5
         Top             =   240
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   11245
         _Version        =   393216
         Cols            =   5
         HighLight       =   2
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
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
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   11
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
      Left            =   10440
      Picture         =   "frmTempDW.frx":0CCE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTempDW.frx":1A56
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmTempDW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type LoadPendaftaranField
    NoPendaftaran As String * 10
    KdRuangan As String * 3
    KdInstalasi As String * 2
    TglPendaftaran As String * 19
    Status As String * 1
End Type

Private RekamMedikPendaftaranLoad As LoadPendaftaranField
Private intFreeFileLoad As Integer
Private lngRecordLenLoad As Long
Private lngNumRecordLoad As Long

Private Sub subSetCmbStatus()
    With Me.cmbStatus
        .AddItem "Belum Diproses", 0
        .AddItem "Sudah Diproses", 1
        .AddItem "Gagal Diproses", 2
        .AddItem "Tampilkan Semua Status", 3
        .ListIndex = 3
    End With
End Sub

Private Sub subSetGrid()
    With Me.MSFlexGridTemp
        .clear
        .Rows = 2
        .Cols = 6
        .FixedRows = 1
        .FixedCols = 1

        .ColWidth(0) = 800
        .ColWidth(1) = 2000 'lngLebar - 700
        .ColWidth(2) = 1400 'lngLebar
        .ColWidth(3) = 1200 'lngLebar - 1200
        .ColWidth(4) = 1200 'lngLebar
        .ColWidth(5) = 1800 'lngLebar
        .ColAlignment(0) = 3
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 3
        .ColAlignment(4) = 3
        .TextMatrix(0, 0) = "No. Data"
        .TextMatrix(0, 1) = "Tanggal Pendaftaran"
        .TextMatrix(0, 2) = "No. Pendaftaran"
        .TextMatrix(0, 3) = "Kode Instalasi"
        .TextMatrix(0, 4) = "Kode Ruangan"
        .TextMatrix(0, 5) = "Status"
    End With
End Sub

Public Sub subLoadTempDataToFlexGrid(ByVal NamaFile As String)
    Dim i As Long, r As Long
    Dim strLokasiFile As String
    Dim lngTotalData As Long, lngProcData As Long
    Dim lngUnprocData As Long, lngFailData As Long

    Call subSetGrid

    strLokasiFile = strFolderDAT & "\" & NamaFile
    If Not fso.FileExists(strLokasiFile) Then
        MsgBox "Tidak ada file temporary untuk tanggal " & Me.dtpTglFileDAT.value, vbExclamation, "Konfirmasi"
        Me.dtpTglFileDAT.value = Now
        Exit Sub
    End If
    intFreeFileLoad = FreeFile
    lngRecordLenLoad = Len(RekamMedikPendaftaranLoad)
    Open strLokasiFile For Random Access Read Write As intFreeFileLoad Len = lngRecordLenLoad
        lngNumRecordLoad = LOF(intFreeFileLoad) \ lngRecordLenLoad
        If LOF(intFreeFileLoad) Mod lngRecordLenLoad > 0 Then lngNumRecordLoad = lngNumRecordLoad + 1
        For i = 1 To lngNumRecordLoad
            Get intFreeFileLoad, i, RekamMedikPendaftaranLoad
            Select Case Me.cmbStatus.ListIndex
                Case 0
                    If RekamMedikPendaftaranLoad.Status = Me.cmbStatus.ListIndex Then
                        r = r + 1
                    Else
                        GoTo jump
                    End If
                Case 1
                    If RekamMedikPendaftaranLoad.Status = Me.cmbStatus.ListIndex Then
                        r = r + 1
                    Else
                        GoTo jump
                    End If
                Case 2
                    If RekamMedikPendaftaranLoad.Status = Me.cmbStatus.ListIndex Then
                        r = r + 1
                    Else
                        GoTo jump
                    End If
                Case 3
                    r = r + 1
            End Select
            With Me.MSFlexGridTemp
                If r > 1 Then .Rows = .Rows + 1
                .TextMatrix(r, 0) = i
                .TextMatrix(r, 1) = RekamMedikPendaftaranLoad.TglPendaftaran
                .TextMatrix(r, 2) = RekamMedikPendaftaranLoad.NoPendaftaran
                .TextMatrix(r, 3) = RekamMedikPendaftaranLoad.KdInstalasi
                .TextMatrix(r, 4) = RekamMedikPendaftaranLoad.KdRuangan
                .TextMatrix(r, 5) = IIf(RekamMedikPendaftaranLoad.Status = "0", "Belum Diproses", "Sudah Diproses")
            End With

jump:
            Select Case RekamMedikPendaftaranLoad.Status
                Case 0
                    lngUnprocData = lngUnprocData + 1
                Case 1
                    lngProcData = lngProcData + 1
                Case 2
                    lngFailData = lngFailData + 1
            End Select
        Next
        lngTotalData = lngNumRecordLoad

        Me.lblTotData.Caption = lngTotalData & " data"
        Me.lblBelum.Caption = lngUnprocData & " data"
        Me.lblSudah.Caption = lngProcData & " data"
        Me.lblGagal.Caption = lngFailData & " data"
    Close intFreeFileLoad
End Sub

Public Function funcBuatNamaFile() As String
    Dim strTglFile As String

    strTglFile = Format(Me.dtpTglFileDAT.value, "yyyyMMdd")
    funcBuatNamaFile = "tempDW" & strTglFile & ".dat"
End Function

Private Sub subSetIntervalAutoRefresh()
    Dim intMiliDetik As Integer

    With Me.tmrAutoRefresh
        intMiliDetik = ((Me.dtpAutoRefresh.Minute * 60) + Me.dtpAutoRefresh.Second) * 1000
        .Interval = intMiliDetik
        .Enabled = True
    End With
End Sub

Private Sub chkAutoRefresh_Click()
    If Me.chkAutoRefresh.value = 1 Then
        Me.frAutoRefresh.Enabled = True
        Me.chkAutoRefresh.Font.Bold = True
        Call subSetIntervalAutoRefresh
    Else
        Me.frAutoRefresh.Enabled = False
        Me.chkAutoRefresh.Font.Bold = False
        Me.tmrAutoRefresh.Enabled = False
    End If
End Sub

Private Sub chkAutoRefresh_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.chkAutoRefresh.value = 1 Then
            Me.frAutoRefresh.Enabled = True
            Me.chkAutoRefresh.Font.Bold = True
            Call subSetIntervalAutoRefresh
        Else
            Me.frAutoRefresh.Enabled = False
            Me.chkAutoRefresh.Font.Bold = False
            Me.tmrAutoRefresh.Enabled = False
        End If
    End If
End Sub

Private Sub cmbStatus_Change()
    Call subLoadTempDataToFlexGrid(funcBuatNamaFile)
End Sub

Private Sub cmbStatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkAutoRefresh.SetFocus
End Sub

Private Sub cmdBuka_Click()
    Call subLoadTempDataToFlexGrid(funcBuatNamaFile)
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpAutoRefresh_Change()
    Call subSetIntervalAutoRefresh
End Sub

Private Sub dtpAutoRefresh_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdBuka.SetFocus
End Sub

Private Sub dtpTglFileDAT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmbStatus.SetFocus
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)

    Me.dtpTglFileDAT.value = Now
    Call subSetCmbStatus
    Call subLoadTempDataToFlexGrid(funcBuatNamaFile)
    If Me.chkAutoRefresh.value = 1 Then Call subSetIntervalAutoRefresh
End Sub

Private Sub tmrAutoRefresh_Timer()
    Call subLoadTempDataToFlexGrid(funcBuatNamaFile)
End Sub
