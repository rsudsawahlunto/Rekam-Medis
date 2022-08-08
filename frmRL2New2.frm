VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRL2New2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL2 Ketenagaan"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6135
   Icon            =   "frmRL2New2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6135
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   6135
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   1905
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPickerAwal 
         Height          =   375
         Left            =   5160
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   136970243
         UpDown          =   -1  'True
         CurrentDate     =   38212
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblPersen 
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   2355
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL2New2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmRL2New2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'made by mario 12/04/12
Option Explicit

'Special Buat Excel
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel.Range
Dim oResizeRange As Excel.Range
Dim i, j, k, l, xx As Integer
Dim w, X, y, z As String
'Special Buat Excel

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    DTPickerAwal.value = Format(Now, "dd/mm/yyyy")

    ProgressBar1.value = ProgressBar1.Min
    ProgressBar1.Max = 120
End Sub

' tenaga medis
Private Sub cmdCetak_Click()
    On Error GoTo error

    ProgressBar1.value = ProgressBar1.Min
    lblPersen.Caption = "0 %"
    Screen.MousePointer = vbHourglass

    'Buka Excel
    Set oXL = CreateObject("Excel.Application")
    'Buat Buka Template
    Set oWB = oXL.Workbooks.Open(App.Path & "\RL 2_Ketenagaan.xlsx")
    Set oSheet = oWB.ActiveSheet

    i = 0
    k = 0
    For i = 1 To 2
        If i = 1 Then
            j = 8
            w = "L"
        Else
            j = 9
            w = "P"
        End If

        For k = 1 To 40
            If k = 1 Then
                l = 3
                X = "0034"
            ElseIf k = 2 Then
                l = 4
                X = "0035"
            ElseIf k = 3 Then
                l = 5
                X = "0036"
            ElseIf k = 4 Then
                l = 6
                X = "0037"
            ElseIf k = 5 Then
                l = 7
                X = "0038"
            ElseIf k = 6 Then
                l = 8
                X = "0039"
            ElseIf k = 7 Then
                l = 9
                X = "0040"
            ElseIf k = 8 Then
                l = 10
                X = "0205"
            ElseIf k = 9 Then
                l = 11
                X = "0206"
            ElseIf k = 10 Then
                l = 12
                X = "0041"
                '==========================
            ElseIf k = 11 Then
                l = 13
                X = "0042"
            ElseIf k = 12 Then
                l = 14
                X = "0043"
            ElseIf k = 13 Then
                l = 15
                X = "0045"
            ElseIf k = 14 Then
                l = 16
                X = "0046"
            ElseIf k = 15 Then
                l = 17
                X = "0049"
            ElseIf k = 16 Then
                l = 18
                X = "0053"
            ElseIf k = 17 Then
                l = 19
                X = "0054"
            ElseIf k = 18 Then
                l = 20
                X = "0055"
            ElseIf k = 19 Then
                l = 21
                X = "0056"
            ElseIf k = 20 Then
                l = 22
                X = "0207"
                '==========================
            ElseIf k = 21 Then
                l = 23
                X = "0060"
            ElseIf k = 22 Then
                l = 24
                X = "0061"
            ElseIf k = 23 Then
                l = 25
                X = "0064"
            ElseIf k = 24 Then
                l = 26
                X = "0065"
            ElseIf k = 25 Then
                l = 27
                X = "0067"
            ElseIf k = 26 Then
                l = 28
                X = "0069"
            ElseIf k = 27 Then
                l = 29
                X = "0071"
            ElseIf k = 28 Then
                l = 30
                X = "0208"
            ElseIf k = 29 Then
                l = 31
                X = "0074"
            ElseIf k = 30 Then
                l = 32
                X = "0076"
                '==========================
            ElseIf k = 31 Then
                l = 33
                X = "0078"
            ElseIf k = 32 Then
                l = 34
                X = "0209"
            ElseIf k = 33 Then
                l = 35
                X = "0081"
            ElseIf k = 34 Then
                l = 36
                X = "0083"
            End If

            strSQL = "SELECT sum (Jml) as jml From RL2New " & _
            " WHERE JenisKelamin = '" & w & "' and KdKualifikasiJurusan = '" & X & "' "
            Set rsb = Nothing
            rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            With oSheet
                If rsb("jml").value <> "" Then
                    .Cells(l, j) = rsb("jml").value
                Else
                    .Cells(l, j) = "0"
                End If
            End With
        Next k
    Next i
    ProgressBar1.value = Int(ProgressBar1.value) + 10
    lblPersen.Caption = Int(ProgressBar1.value / 120 * 100) & " %"

    Call tenaga_keperawatan
    Call kefarmasian
    Call kesehatan_masyarakat
    Call gizi
    Call keterapian_fisik
    Call keteknisian_medis

    Call doktoral
    Call pasca_sarjana
    Call sarjana
    Call sarjana_muda
    Call smuSederajatdanDibawahnya
    Call KebutuhandanKekuranganPegawai

    strSQL = "SELECT * From ProfilRS"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    For xx = 2 To 178
        With oSheet
            .Cells(xx, 1) = rsb("KdRS").value
            .Cells(xx, 2) = rsb("KotaKodyaKab").value
            .Cells(xx, 4) = rsb("NamaRS").value
            .Cells(xx, 5) = Right(DTPickerAwal.value, 4)
        End With
    Next xx

    oXL.Visible = True
    Screen.MousePointer = vbDefault

    Exit Sub
error:
    Call msubPesanError
    Screen.MousePointer = vbDefault
End Sub

Sub tenaga_keperawatan()
    i = 0
    k = 0
    For i = 1 To 2
        If i = 1 Then
            j = 8
            w = "L"
        Else
            j = 9
            w = "P"
        End If

        For k = 1 To 12
            If k = 1 Then
                l = 41
                X = "0005"
            ElseIf k = 2 Then
                l = 42
                X = "0006"
            ElseIf k = 3 Then
                l = 43
                X = "0007"
            ElseIf k = 4 Then
                l = 44
                X = "0146"
            ElseIf k = 5 Then
                l = 45
                X = "0210"
            ElseIf k = 6 Then
                l = 46
                X = "0211"
            ElseIf k = 7 Then
                l = 47
                X = "0212"
            ElseIf k = 8 Then
                l = 48
                X = "0187"
            ElseIf k = 9 Then
                l = 49
                X = "0188"
            ElseIf k = 10 Then
                l = 50
                X = "0189"
            ElseIf k = 11 Then
                l = 51
                X = "0147"
            ElseIf k = 12 Then
                l = 52
                X = "0178"
            End If

            strSQL = "SELECT sum (Jml) as jml From RL4New " & _
            " WHERE JenisKelamin = '" & w & "' and KdKualifikasiJurusan = '" & X & "' "
            Set rsb = Nothing
            rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            With oSheet
                If rsb("jml").value <> "" Then
                    .Cells(l, j) = rsb("jml").value
                Else
                    .Cells(l, j) = "0"
                End If
            End With
        Next k
    Next i
    ProgressBar1.value = Int(ProgressBar1.value) + 10
    lblPersen.Caption = Int(ProgressBar1.value / 120 * 100) & " %"
End Sub

Sub kefarmasian()
    i = 0
    k = 0
    For i = 1 To 2
        If i = 1 Then
            j = 8
            w = "L"
        Else
            j = 9
            w = "P"
        End If

        For k = 1 To 10
            If k = 1 Then
                l = 54
                X = "0008"
            ElseIf k = 2 Then
                l = 55
                X = "0009"
            ElseIf k = 3 Then
                l = 56
                X = "0153"
            ElseIf k = 4 Then
                l = 57
                X = "0154"
            ElseIf k = 5 Then
                l = 58
                X = "0155"
            ElseIf k = 6 Then
                l = 59
                X = "0190"
            ElseIf k = 7 Then
                l = 60
                X = "0156"
            ElseIf k = 8 Then
                l = 61
                X = "0157"
            ElseIf k = 9 Then
                l = 62
                X = "0158"
            ElseIf k = 10 Then
                l = 63
                X = "0159"
            End If

            strSQL = "SELECT sum (Jml) as Jml From RL4New " & _
            " WHERE JenisKelamin = '" & w & "' and KdKualifikasiJurusan = '" & X & "' "

            Set rsb = Nothing
            rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            With oSheet
                If rsb("jml").value <> "" Then
                    .Cells(l, j) = rsb("jml").value
                Else
                    .Cells(l, j) = "0"
                End If
            End With
        Next k
    Next i
    ProgressBar1.value = Int(ProgressBar1.value) + 10
    lblPersen.Caption = Int(ProgressBar1.value / 120 * 100) & " %"
End Sub

Sub kesehatan_masyarakat()
    i = 0
    k = 0
    For i = 1 To 2
        If i = 1 Then
            j = 8
            w = "L"
        Else
            j = 9
            w = "P"
        End If

        For k = 1 To 13
            If k = 1 Then
                l = 65
                X = "0016"
            ElseIf k = 2 Then
                l = 66
                X = "0017"
            ElseIf k = 3 Then
                l = 67
                X = "0191"
            ElseIf k = 4 Then
                l = 68
                X = "0018"
            ElseIf k = 5 Then
                l = 69
                X = "0019"
            ElseIf k = 6 Then
                l = 70
                X = "0192"
            ElseIf k = 7 Then
                l = 71
                X = "0193"
            ElseIf k = 8 Then
                l = 72
                X = "0020"
            ElseIf k = 9 Then
                l = 73
                X = "0194"
            ElseIf k = 10 Then
                l = 74
                X = "0021"
            ElseIf k = 11 Then
                l = 75
                X = "0022"
            ElseIf k = 12 Then
                l = 76
                X = "0023"
            ElseIf k = 13 Then
                l = 77
                X = "0180"
            End If

            strSQL = "SELECT sum (Jml) as jml From RL4New " & _
            " WHERE JenisKelamin = '" & w & "' and KdKualifikasiJurusan = '" & X & "' "
            Set rsb = Nothing
            rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            With oSheet
                If rsb("jml").value <> "" Then
                    .Cells(l, j) = rsb("jml").value
                Else
                    .Cells(l, j) = "0"
                End If
            End With
        Next k
    Next i
    ProgressBar1.value = Int(ProgressBar1.value) + 10
    lblPersen.Caption = Int(ProgressBar1.value / 120 * 100) & " %"
End Sub

Sub gizi()
    i = 0
    k = 0
    For i = 1 To 2
        If i = 1 Then
            j = 8
            w = "L"
        Else
            j = 9
            w = "P"
        End If

        For k = 1 To 7
            If k = 1 Then
                l = 79
                X = "0024"
            ElseIf k = 2 Then
                l = 80
                X = "0025"
            ElseIf k = 3 Then
                l = 81
                X = "0026"
            ElseIf k = 4 Then
                l = 82
                X = "0027"
            ElseIf k = 5 Then
                l = 83
                X = "0029"
            ElseIf k = 6 Then
                l = 84
                X = "0030"
            ElseIf k = 7 Then
                l = 85
                X = "0181"
            End If

            strSQL = "SELECT sum (Jml) as jml From RL4New " & _
            " WHERE JenisKelamin = '" & w & "' and KdKualifikasiJurusan = '" & X & "' "

            Set rsb = Nothing
            rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            With oSheet
                If rsb("jml").value <> "" Then
                    .Cells(l, j) = rsb("jml").value
                Else
                    .Cells(l, j) = "0"
                End If
            End With
        Next k
    Next i
    ProgressBar1.value = Int(ProgressBar1.value) + 10
    lblPersen.Caption = Int(ProgressBar1.value / 120 * 100) & " %"
End Sub

Sub keterapian_fisik()
    i = 0
    k = 0
    For i = 1 To 2
        If i = 1 Then
            j = 8
            w = "L"
        Else
            j = 9
            w = "P"
        End If

        For k = 1 To 7
            If k = 1 Then
                l = 87
                X = "0195"
            ElseIf k = 2 Then
                l = 88
                X = "0031"
            ElseIf k = 3 Then
                l = 89
                X = "0032"
            ElseIf k = 4 Then
                l = 90
                X = "0033"
            ElseIf k = 5 Then
                l = 91
                X = "0196"
            ElseIf k = 6 Then
                l = 92
                X = "0197"
            ElseIf k = 7 Then
                l = 93
                X = "0182"
            End If

            strSQL = "SELECT sum (Jml) as jml From RL4New " & _
            " WHERE JenisKelamin = '" & w & "' and KdKualifikasiJurusan = '" & X & "' "

            Set rsb = Nothing
            rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            With oSheet
                If rsb("jml").value <> "" Then
                    .Cells(l, j) = rsb("jml").value
                Else
                    .Cells(l, j) = "0"
                End If
            End With
        Next k
    Next i
    ProgressBar1.value = Int(ProgressBar1.value) + 10
    lblPersen.Caption = Int(ProgressBar1.value / 120 * 100) & " %"
End Sub

Sub keteknisian_medis()
    i = 0
    k = 0
    For i = 1 To 2
        If i = 1 Then
            j = 8
            w = "L"
        Else
            j = 9
            w = "P"
        End If

        For k = 1 To 23
            If k = 1 Then
                l = 95
                X = "0095"
            ElseIf k = 2 Then
                l = 96
                X = "0096"
            ElseIf k = 3 Then
                l = 97
                X = "0098"
            ElseIf k = 4 Then
                l = 98
                X = "0100"
            ElseIf k = 5 Then
                l = 99
                X = "0198"
            ElseIf k = 6 Then
                l = 100
                X = "0102"
            ElseIf k = 7 Then
                l = 101
                X = "0103"
            ElseIf k = 8 Then
                l = 102
                X = "0105"
            ElseIf k = 9 Then
                l = 103
                X = "0107"
            ElseIf k = 10 Then
                l = 104
                X = "0109"
            ElseIf k = 11 Then
                l = 105
                X = "0110"
            ElseIf k = 12 Then
                l = 106
                X = "0112"
            ElseIf k = 13 Then
                l = 107
                X = "0113"
            ElseIf k = 14 Then
                l = 108
                X = "0114"
            ElseIf k = 15 Then
                l = 109
                X = "0116"
            ElseIf k = 16 Then
                l = 110
                X = "0183"
            ElseIf k = 17 Then
                l = 111
                X = "0199"
            ElseIf k = 18 Then
                l = 112
                X = "0200"
            ElseIf k = 19 Then
                l = 113
                X = "0201"
            ElseIf k = 20 Then
                l = 114
                X = "0184"
            ElseIf k = 21 Then
                l = 115
                X = "0202"
            ElseIf k = 22 Then
                l = 116
                X = "0203"
            ElseIf k = 23 Then
                l = 117
                X = "0182"
            End If

            strSQL = "SELECT sum (Jml) as jml From RL4New " & _
            " WHERE JenisKelamin = '" & w & "' and KdKualifikasiJurusan = '" & X & "' "

            Set rsb = Nothing
            rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            With oSheet
                If rsb("jml").value <> "" Then
                    .Cells(l, j) = rsb("jml").value
                Else
                    .Cells(l, j) = "0"
                End If
            End With
        Next k
    Next i
    ProgressBar1.value = Int(ProgressBar1.value) + 10
    lblPersen.Caption = Int(ProgressBar1.value / 120 * 100) & " %"
End Sub

Sub doktoral()
    i = 0
    k = 0
    For i = 1 To 2
        If i = 1 Then
            j = 8
            w = "L"
        Else
            j = 9
            w = "P"
        End If

        For k = 1 To 11
            If k = 1 Then
                l = 120
                X = "0115"
            ElseIf k = 2 Then
                l = 121
                X = "0117"
            ElseIf k = 3 Then
                l = 122
                X = "0119"
            ElseIf k = 4 Then
                l = 123
                X = "0120"
            ElseIf k = 5 Then
                l = 124
                X = "0121"
            ElseIf k = 6 Then
                l = 125
                X = "0122"
            ElseIf k = 7 Then
                l = 126
                X = "0123"
            ElseIf k = 8 Then
                l = 127
                X = "0124"
            ElseIf k = 9 Then
                l = 128
                X = "0125"
            ElseIf k = 10 Then
                l = 129
                X = "0126"
            ElseIf k = 11 Then
                l = 130
                X = "0127"
            End If

            strSQL = "SELECT sum (Jml) as jml From RL4New " & _
            " WHERE JenisKelamin = '" & w & "' and KdKualifikasiJurusan = '" & X & "' "

            Set rsb = Nothing
            rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            With oSheet
                If rsb("jml").value <> "" Then
                    .Cells(l, j) = rsb("jml").value
                Else
                    .Cells(l, j) = "0"
                End If
            End With
        Next k
    Next i
    ProgressBar1.value = Int(ProgressBar1.value) + 10
    lblPersen.Caption = Int(ProgressBar1.value / 120 * 100) & " %"
End Sub

Sub pasca_sarjana()
    i = 0
    k = 0
    For i = 1 To 2
        If i = 1 Then
            j = 8
            w = "L"
        Else
            j = 9
            w = "P"
        End If

        For k = 1 To 12
            If k = 1 Then
                l = 133
                X = "0128"
            ElseIf k = 2 Then
                l = 134
                X = "0129"
            ElseIf k = 3 Then
                l = 135
                X = "0131"
            ElseIf k = 4 Then
                l = 136
                X = "0132"
            ElseIf k = 5 Then
                l = 137
                X = "0133"
            ElseIf k = 6 Then
                l = 138
                X = "0134"
            ElseIf k = 7 Then
                l = 139
                X = "0135"
            ElseIf k = 8 Then
                l = 140
                X = "0136"
            ElseIf k = 9 Then
                l = 141
                X = "0137"
            ElseIf k = 10 Then
                l = 142
                X = "0138"
            ElseIf k = 11 Then
                l = 143
                X = "0139"
            ElseIf k = 12 Then
                l = 144
                X = "0140"
            End If

            strSQL = "SELECT sum (Jml) as jml From RL4New " & _
            " WHERE JenisKelamin = '" & w & "' and KdKualifikasiJurusan = '" & X & "' "

            Set rsb = Nothing
            rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            With oSheet
                If rsb("jml").value <> "" Then
                    .Cells(l, j) = rsb("jml").value
                Else
                    .Cells(l, j) = "0"
                End If
            End With
        Next k
    Next i
    ProgressBar1.value = Int(ProgressBar1.value) + 10
    lblPersen.Caption = Int(ProgressBar1.value / 120 * 100) & " %"
End Sub

Sub sarjana()
    i = 0
    k = 0
    For i = 1 To 2
        If i = 1 Then
            j = 8
            w = "L"
        Else
            j = 9
            w = "P"
        End If

        For k = 1 To 11
            If k = 1 Then
                l = 147
                X = "0062"
            ElseIf k = 2 Then
                l = 148
                X = "0063"
            ElseIf k = 3 Then
                l = 149
                X = "0068"
            ElseIf k = 4 Then
                l = 150
                X = "0070"
            ElseIf k = 5 Then
                l = 151
                X = "0072"
            ElseIf k = 6 Then
                l = 152
                X = "0073"
            ElseIf k = 7 Then
                l = 153
                X = "0075"
            ElseIf k = 8 Then
                l = 154
                X = "0077"
            ElseIf k = 9 Then
                l = 155
                X = "0079"
            ElseIf k = 10 Then
                l = 156
                X = "0080"
            ElseIf k = 11 Then
                l = 157
                X = "0082"
            End If

            strSQL = "SELECT sum (Jml) as Jml From RL4New " & _
            " WHERE JenisKelamin = '" & w & "' and KdKualifikasiJurusan = '" & X & "' "

            Set rsb = Nothing
            rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            With oSheet
                If rsb("Jml").value <> "" Then
                    .Cells(l, j) = rsb("Jml").value
                Else
                    .Cells(l, j) = "0"
                End If
            End With
        Next k
    Next i
    ProgressBar1.value = Int(ProgressBar1.value) + 10
    lblPersen.Caption = Int(ProgressBar1.value / 120 * 100) & " %"
End Sub

Sub sarjana_muda()
    i = 0
    k = 0
    For i = 1 To 2
        If i = 1 Then
            j = 8
            w = "L"
        Else
            j = 9
            w = "P"
        End If

        For k = 1 To 11
            If k = 1 Then
                l = 159
                X = "0087"
            ElseIf k = 2 Then
                l = 160
                X = "0088"
            ElseIf k = 3 Then
                l = 161
                X = "0093"
            ElseIf k = 4 Then
                l = 162
                X = "0097"
            ElseIf k = 5 Then
                l = 163
                X = "0094"
            ElseIf k = 6 Then
                l = 164
                X = "0204"
            ElseIf k = 7 Then
                l = 165
                X = "0101"
            ElseIf k = 8 Then
                l = 166
                X = "0104"
            ElseIf k = 9 Then
                l = 167
                X = "0106"
            ElseIf k = 10 Then
                l = 168
                X = "0108"
            ElseIf k = 11 Then
                l = 169
                X = "0111"
            End If

            strSQL = "SELECT sum (Jml) as jml From RL4New " & _
            " WHERE JenisKelamin = '" & w & "' and KdKualifikasiJurusan = '" & X & "' "

            Set rsb = Nothing
            rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            With oSheet
                If rsb("jml").value <> "" Then
                    .Cells(l, j) = rsb("jml").value
                Else
                    .Cells(l, j) = "0"
                End If
            End With
        Next k
    Next i
    ProgressBar1.value = Int(ProgressBar1.value) + 10
    lblPersen.Caption = Int(ProgressBar1.value / 120 * 100) & " %"
End Sub

Sub smuSederajatdanDibawahnya()
    i = 0
    k = 0
    For i = 1 To 2
        If i = 1 Then
            j = 8
            w = "L"
        Else
            j = 9
            w = "P"
        End If

        For k = 1 To 8
            If k = 1 Then
                l = 171
                X = "0013"
            ElseIf k = 2 Then
                l = 172
                X = "0047"
            ElseIf k = 3 Then
                l = 173
                X = "0048"
            ElseIf k = 4 Then
                l = 174
                X = "0050"
            ElseIf k = 5 Then
                l = 175
                X = "0051"
            ElseIf k = 6 Then
                l = 176
                X = "0057"
            ElseIf k = 7 Then
                l = 177
                X = "0059"
            ElseIf k = 8 Then
                l = 178
                X = "0052"
            End If

            strSQL = "SELECT sum (Jml) as jml From RL4New " & _
            " WHERE JenisKelamin = '" & w & "' and KdKualifikasiJurusan = '" & X & "' "

            Set rsb = Nothing
            rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            With oSheet
                If rsb("jml").value <> "" Then
                    .Cells(l, j) = rsb("jml").value
                Else
                    .Cells(l, j) = "0"
                End If
            End With
        Next k
    Next i
    ProgressBar1.value = Int(ProgressBar1.value) + 10
    lblPersen.Caption = Int(ProgressBar1.value / 120 * 100) & " %"
End Sub

Sub KebutuhandanKekuranganPegawai()
    
        For k = 1 To 159  ' Didapat dari penjumlahan semua nilai k (for k = 1 to n) semua bidang kualifikasijurusan
        'Tenaga Medis
            If k = 1 Then
                l = 3
                X = "0034"
            ElseIf k = 2 Then
                l = 4
                X = "0035"
            ElseIf k = 3 Then
                l = 5
                X = "0036"
            ElseIf k = 4 Then
                l = 6
                X = "0037"
            ElseIf k = 5 Then
                l = 7
                X = "0038"
            ElseIf k = 6 Then
                l = 8
                X = "0039"
            ElseIf k = 7 Then
                l = 9
                X = "0040"
            ElseIf k = 8 Then
                l = 10
                X = "0205"
            ElseIf k = 9 Then
                l = 11
                X = "0206"
            ElseIf k = 10 Then
                l = 12
                X = "0041"
            ElseIf k = 11 Then
                l = 13
                X = "0042"
            ElseIf k = 12 Then
                l = 14
                X = "0043"
            ElseIf k = 13 Then
                l = 15
                X = "0045"
            ElseIf k = 14 Then
                l = 16
                X = "0046"
            ElseIf k = 15 Then
                l = 17
                X = "0049"
            ElseIf k = 16 Then
                l = 18
                X = "0053"
            ElseIf k = 17 Then
                l = 19
                X = "0054"
            ElseIf k = 18 Then
                l = 20
                X = "0055"
            ElseIf k = 19 Then
                l = 21
                X = "0056"
            ElseIf k = 20 Then
                l = 22
                X = "0207"
            ElseIf k = 21 Then
                l = 23
                X = "0060"
            ElseIf k = 22 Then
                l = 24
                X = "0061"
            ElseIf k = 23 Then
                l = 25
                X = "0064"
            ElseIf k = 24 Then
                l = 26
                X = "0065"
            ElseIf k = 25 Then
                l = 27
                X = "0067"
            ElseIf k = 26 Then
                l = 28
                X = "0069"
            ElseIf k = 27 Then
                l = 29
                X = "0071"
            ElseIf k = 28 Then
                l = 30
                X = "0208"
            ElseIf k = 29 Then
                l = 31
                X = "0074"
            ElseIf k = 30 Then
                l = 32
                X = "0076"
            ElseIf k = 31 Then
                l = 33
                X = "0078"
            ElseIf k = 32 Then
                l = 34
                X = "0209"
            ElseIf k = 33 Then
                l = 35
                X = "0081"
            ElseIf k = 34 Then
                l = 36
                X = "0083"
        '=========================== Tenaga Keperawatan
            ElseIf k = 35 Then
                l = 41
                X = "0005"
            ElseIf k = 36 Then
                l = 42
                X = "0006"
            ElseIf k = 37 Then
                l = 43
                X = "0007"
            ElseIf k = 38 Then
                l = 44
                X = "0146"
            ElseIf k = 39 Then
                l = 45
                X = "0210"
            ElseIf k = 40 Then
                l = 46
                X = "0211"
            ElseIf k = 41 Then
                l = 47
                X = "0212"
            ElseIf k = 42 Then
                l = 48
                X = "0187"
            ElseIf k = 43 Then
                l = 49
                X = "0188"
            ElseIf k = 44 Then
                l = 50
                X = "0189"
            ElseIf k = 45 Then
                l = 51
                X = "0147"
            ElseIf k = 46 Then
                l = 52
                X = "0178"
            
  '======================================= Kefarmasian
            ElseIf k = 47 Then
                l = 54
                X = "0008"
            ElseIf k = 48 Then
                l = 55
                X = "0009"
            ElseIf k = 49 Then
                l = 56
                X = "0153"
            ElseIf k = 50 Then
                l = 57
                X = "0154"
            ElseIf k = 51 Then
                l = 58
                X = "0155"
            ElseIf k = 52 Then
                l = 59
                X = "0190"
            ElseIf k = 53 Then
                l = 60
                X = "0156"
            ElseIf k = 54 Then
                l = 61
                X = "0157"
            ElseIf k = 55 Then
                l = 62
                X = "0158"
            ElseIf k = 56 Then
                l = 63
                X = "0159"
           '============================================kesehatan_masyarakat
            ElseIf k = 57 Then
                l = 65
                X = "0016"
            ElseIf k = 58 Then
                l = 66
                X = "0017"
            ElseIf k = 59 Then
                l = 67
                X = "0191"
            ElseIf k = 60 Then
                l = 68
                X = "0018"
            ElseIf k = 61 Then
                l = 69
                X = "0019"
            ElseIf k = 62 Then
                l = 70
                X = "0192"
            ElseIf k = 63 Then
                l = 71
                X = "0193"
            ElseIf k = 64 Then
                l = 72
                X = "0020"
            ElseIf k = 65 Then
                l = 73
                X = "0194"
            ElseIf k = 66 Then
                l = 74
                X = "0021"
            ElseIf k = 67 Then
                l = 75
                X = "0022"
            ElseIf k = 68 Then
                l = 76
                X = "0023"
            ElseIf k = 69 Then
                l = 77
                X = "0180"
          '==============================Gizi
            ElseIf k = 70 Then
                l = 79
                X = "0024"
            ElseIf k = 71 Then
                l = 80
                X = "0025"
            ElseIf k = 72 Then
                l = 81
                X = "0026"
            ElseIf k = 73 Then
                l = 82
                X = "0027"
            ElseIf k = 74 Then
                l = 83
                X = "0029"
            ElseIf k = 75 Then
                l = 84
                X = "0030"
            ElseIf k = 76 Then
                l = 85
                X = "0181"
            '==============================keterapian_fisik
            ElseIf k = 77 Then
                l = 87
                X = "0195"
            ElseIf k = 78 Then
                l = 88
                X = "0031"
            ElseIf k = 79 Then
                l = 89
                X = "0032"
            ElseIf k = 80 Then
                l = 90
                X = "0033"
            ElseIf k = 81 Then
                l = 91
                X = "0196"
            ElseIf k = 82 Then
                l = 92
                X = "0197"
            ElseIf k = 83 Then
                l = 93
                X = "0182"
             '==============================keteknisian_medis
            ElseIf k = 84 Then
                l = 95
                X = "0095"
            ElseIf k = 85 Then
                l = 96
                X = "0096"
            ElseIf k = 86 Then
                l = 97
                X = "0098"
            ElseIf k = 87 Then
                l = 98
                X = "0100"
            ElseIf k = 88 Then
                l = 99
                X = "0198"
            ElseIf k = 89 Then
                l = 100
                X = "0102"
            ElseIf k = 90 Then
                l = 101
                X = "0103"
            ElseIf k = 91 Then
                l = 102
                X = "0105"
            ElseIf k = 92 Then
                l = 103
                X = "0107"
            ElseIf k = 93 Then
                l = 104
                X = "0109"
            ElseIf k = 94 Then
                l = 105
                X = "0110"
            ElseIf k = 95 Then
                l = 106
                X = "0112"
            ElseIf k = 96 Then
                l = 107
                X = "0113"
            ElseIf k = 97 Then
                l = 108
                X = "0114"
            ElseIf k = 98 Then
                l = 109
                X = "0116"
            ElseIf k = 99 Then
                l = 110
                X = "0183"
            ElseIf k = 100 Then
                l = 111
                X = "0199"
            ElseIf k = 101 Then
                l = 112
                X = "0200"
            ElseIf k = 102 Then
                l = 113
                X = "0201"
            ElseIf k = 103 Then
                l = 114
                X = "0184"
            ElseIf k = 104 Then
                l = 115
                X = "0202"
            ElseIf k = 105 Then
                l = 116
                X = "0203"
            ElseIf k = 106 Then
                l = 117
                X = "0182"
            '===============================================doktoral
            ElseIf k = 107 Then
                l = 120
                X = "0115"
            ElseIf k = 108 Then
                l = 121
                X = "0117"
            ElseIf k = 109 Then
                l = 122
                X = "0119"
            ElseIf k = 110 Then
                l = 123
                X = "0120"
            ElseIf k = 111 Then
                l = 124
                X = "0121"
            ElseIf k = 112 Then
                l = 125
                X = "0122"
            ElseIf k = 113 Then
                l = 126
                X = "0123"
            ElseIf k = 114 Then
                l = 127
                X = "0124"
            ElseIf k = 115 Then
                l = 128
                X = "0125"
            ElseIf k = 116 Then
                l = 129
                X = "0126"
            ElseIf k = 117 Then
                l = 130
                X = "0127"
            '==========================pasca sarjana
            ElseIf k = 118 Then
                l = 133
                X = "0128"
            ElseIf k = 119 Then
                l = 134
                X = "0129"
            ElseIf k = 120 Then
                l = 135
                X = "0131"
            ElseIf k = 121 Then
                l = 136
                X = "0132"
            ElseIf k = 122 Then
                l = 137
                X = "0133"
            ElseIf k = 123 Then
                l = 138
                X = "0134"
            ElseIf k = 124 Then
                l = 139
                X = "0135"
            ElseIf k = 125 Then
                l = 140
                X = "0136"
            ElseIf k = 126 Then
                l = 141
                X = "0137"
            ElseIf k = 127 Then
                l = 142
                X = "0138"
            ElseIf k = 128 Then
                l = 143
                X = "0139"
            ElseIf k = 129 Then
                l = 144
                X = "0140"
            '===========================sarjana
            ElseIf k = 130 Then
                l = 147
                X = "0062"
            ElseIf k = 131 Then
                l = 148
                X = "0063"
            ElseIf k = 132 Then
                l = 149
                X = "0068"
            ElseIf k = 133 Then
                l = 150
                X = "0070"
            ElseIf k = 134 Then
                l = 151
                X = "0072"
            ElseIf k = 135 Then
                l = 152
                X = "0073"
            ElseIf k = 136 Then
                l = 153
                X = "0075"
            ElseIf k = 137 Then
                l = 154
                X = "0077"
            ElseIf k = 138 Then
                l = 155
                X = "0079"
            ElseIf k = 139 Then
                l = 156
                X = "0080"
            ElseIf k = 140 Then
                l = 157
                X = "0082"
            '=============================sarjana muda
            ElseIf k = 141 Then
                l = 159
                X = "0087"
            ElseIf k = 142 Then
                l = 160
                X = "0088"
            ElseIf k = 143 Then
                l = 161
                X = "0093"
            ElseIf k = 144 Then
                l = 162
                X = "0097"
            ElseIf k = 145 Then
                l = 163
                X = "0094"
            ElseIf k = 146 Then
                l = 164
                X = "0204"
            ElseIf k = 147 Then
                l = 165
                X = "0101"
            ElseIf k = 148 Then
                l = 166
                X = "0104"
            ElseIf k = 149 Then
                l = 167
                X = "0106"
            ElseIf k = 150 Then
                l = 168
                X = "0108"
            ElseIf k = 151 Then
                l = 169
                X = "0111"
           '==========================smuSederajatdanDibawahnya
           ElseIf k = 152 Then
                l = 171
                X = "0013"
            ElseIf k = 153 Then
                l = 172
                X = "0047"
            ElseIf k = 154 Then
                l = 173
                X = "0048"
            ElseIf k = 155 Then
                l = 174
                X = "0050"
            ElseIf k = 156 Then
                l = 175
                X = "0051"
            ElseIf k = 157 Then
                l = 176
                X = "0057"
            ElseIf k = 158 Then
                l = 177
                X = "0059"
            ElseIf k = 159 Then
                l = 178
                X = "0052"
            End If
            
            strSQL = "select * from RL2_KebutuhandanKekuranganPegawai where KdKualifikasiJurusan = '" & X & "' "
            Set rsb = Nothing
            rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            With oSheet
                If rsb.RecordCount <> 0 Then
                    .Cells(l, 10) = rsb("KebutuhanPegawaiPria").value
                    .Cells(l, 11) = rsb("KebutuhanPegawaiWanita").value
                    .Cells(l, 12) = rsb("KekuranganPegawaiPria").value
                    .Cells(l, 13) = rsb("KekuranganPegawaiWanita").value
                Else
                    .Cells(l, 10) = "0"
                    .Cells(l, 11) = "0"
                    .Cells(l, 12) = "0"
                    .Cells(l, 13) = "0"
                End If
            End With
        Next k

End Sub


