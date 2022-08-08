VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDaftarDokumenRekamMedisPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Dokumen Rekam Medis Pasien"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarDokumenRekamMedisPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   14895
   Begin VB.Frame Frame2 
      Height          =   840
      Left            =   0
      TabIndex        =   17
      Top             =   7320
      Width           =   4575
      Begin VB.Label Label2 
         Caption         =   "F1 - Cetak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fraCari 
      Caption         =   "Cari Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   4560
      TabIndex        =   15
      Top             =   7320
      Width           =   10335
      Begin VB.CommandButton cmdKirim 
         Caption         =   "&Kirim Dokumen"
         Height          =   495
         Left            =   5520
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdTerima 
         Caption         =   "&Terima Dokumen"
         Height          =   495
         Left            =   7080
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   600
         TabIndex        =   9
         Top             =   480
         Width           =   3495
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   8640
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdKirimTerima 
         Caption         =   "&Kirim dan Terima"
         Height          =   495
         Left            =   8490
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien /  No. CM/ Ruangan"
         Height          =   210
         Left            =   600
         TabIndex        =   16
         Top             =   240
         Width           =   2700
      End
   End
   Begin VB.Frame fraDaftar 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   0
      TabIndex        =   12
      Top             =   1080
      Width           =   14895
      Begin VB.CheckBox chkPilihSemua 
         Caption         =   "Semua"
         Height          =   220
         Left            =   120
         TabIndex        =   28
         Top             =   5930
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkCheck 
         BackColor       =   &H0000FF00&
         Caption         =   "Check1"
         Height          =   250
         Left            =   240
         TabIndex        =   25
         Top             =   1250
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSFlexGridLib.MSFlexGrid fgDaftarDokumenRekamMedis 
         Height          =   4935
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   8705
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   8454016
         Appearance      =   0
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
         Height          =   735
         Left            =   4320
         TabIndex        =   13
         Top             =   150
         Width           =   10455
         Begin VB.OptionButton optTglTerima 
            Caption         =   "TglTerima"
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
            Left            =   3240
            TabIndex        =   5
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optTglKirim 
            Caption         =   "TglKirim"
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
            Left            =   2280
            TabIndex        =   4
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optTglPulang 
            Caption         =   "TglPulang"
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
            Left            =   1200
            TabIndex        =   3
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optTglMasuk 
            Caption         =   "TglMasuk"
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
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   8
            Top             =   240
            Width           =   735
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   4440
            TabIndex        =   6
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   108003331
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   7080
            TabIndex        =   7
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   108003331
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   6720
            TabIndex        =   14
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataListLib.DataCombo dcRuangPelayanan 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcRuanganPengirim 
         Height          =   330
         Left            =   2040
         TabIndex        =   1
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan Pengirim"
         Height          =   210
         Left            =   2040
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan Pelayanan"
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   21
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
   Begin MSComCtl2.DTPicker dtpTglKirim 
      Height          =   375
      Left            =   2520
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy HH:mm"
      Format          =   129105923
      UpDown          =   -1  'True
      CurrentDate     =   38212
   End
   Begin MSComCtl2.DTPicker dtpTglTerima 
      Height          =   375
      Left            =   4680
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy HH:mm"
      Format          =   129105923
      UpDown          =   -1  'True
      CurrentDate     =   38212
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   13080
      Picture         =   "frmDaftarDokumenRekamMedisPasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarDokumenRekamMedisPasien.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarDokumenRekamMedisPasien.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "frmDaftarDokumenRekamMedisPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim refreshOtomatis As Boolean
Dim dTglMasuk As Date

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    Call msubDcSource(dcRuangPelayanan, rs, "SELECT distinct KdRuanganPelayanan, RuanganPelayanan FROM  V_DokumenRekamMedisPasien ") 'StatusEnabled='1'")
    Call msubDcSource(dcRuanganPengirim, rs, "SELECT distinct KdRuanganPengirim, RuanganPengirim FROM V_DokumenRekamMedisPasien ") 'StatusEnabled='1'")

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub chkCheck_Click()
On Error GoTo errLoad

    If chkCheck.value = vbChecked Then
        fgDaftarDokumenRekamMedis.TextMatrix(fgDaftarDokumenRekamMedis.Row, 0) = Chr$(187)
        fgDaftarDokumenRekamMedis.TextMatrix(fgDaftarDokumenRekamMedis.Row, 15) = 1
    Else
        fgDaftarDokumenRekamMedis.TextMatrix(fgDaftarDokumenRekamMedis.Row, 0) = ""
        fgDaftarDokumenRekamMedis.TextMatrix(fgDaftarDokumenRekamMedis.Row, 15) = 0
    End If
Exit Sub
errLoad:
msubPesanError

End Sub

Private Sub chkCheck_GotFocus()
    If fgDaftarDokumenRekamMedis.TextMatrix(fgDaftarDokumenRekamMedis.Row, 0) = Chr$(187) Then
        chkCheck.value = 1
    Else
        chkCheck.value = 0
    End If
End Sub

Private Sub chkCheck_LostFocus()
    chkCheck.Visible = False
End Sub

Private Sub chkPilihSemua_Click()
On Error Resume Next
Dim i As Integer
    strSQL = "SELECT distinct top 100 NoPendaftaran, NoCM, NamaPasien, JK,RuanganPelayanan, TglMasuk, TglPulang, TglKirim, TglTerima, RuanganPengirim, KeteranganKirim, UserPengirim,RuanganTujuan ,KeteranganTerima , UserPenerima,KdRuanganTujuan,KdRuanganPelayanan " & _
             " FROM V_DokumenRekamMedisPasien " & _
             " WHERE NoCM like '%" & txtParameter.Text & "%' or NamaPasien like '%" & txtParameter.Text & "%' or RuanganPelayanan like '%" & txtParameter.Text & "%' and  RuanganPelayanan Like '%" & dcRuangPelayanan.Text & "%' " & _
             " And TglMasuk between '" & Format(dtpAwal.value, "yyyy/mm/dd 00:00:00") & "' and '" & Format(dtpAkhir.value, "yyyy/mm/dd 23:59:00") & "' "
    Set rsCek = Nothing
Call msubRecFO(rsCek, strSQL)
If rsCek.EOF = False Then
    For i = 1 To rsCek.RecordCount
    With fgDaftarDokumenRekamMedis
        If chkPilihSemua.value = vbUnchecked Then
            .TextMatrix(i, 0) = ""
        Else
            .TextMatrix(i, 0) = Chr$(187)
        End If
    End With
        rsCek.MoveNext
    Next i
End If
Exit Sub
'errload:
'    Call msubPesanError
End Sub

Public Sub cmdCari_Click()
    On Error GoTo errLoad
    mstrFilterData = ""
    MousePointer = vbHourglass
    If Len(Trim(dcRuangPelayanan.Text)) <> 0 Then
        mstrFilterData = "AND KdRuanganPelayanan ='" & dcRuangPelayanan.BoundText & "' "
    End If
    If Len(Trim(dcRuanganPengirim.Text)) <> 0 Then
        mstrFilterData = mstrFilterData & "AND KdRuanganPengirim ='" & dcRuanganPengirim.BoundText & "' "
    End If
    If optTglMasuk.value = True Then
        mstrFilter = "AND TglMasuk between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "'and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'  "
    ElseIf optTglPulang.value = True Then
        mstrFilter = "AND TglPulang between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "'and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'  "
    ElseIf optTglKirim.value = True Then
        mstrFilter = "AND TglKirim between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "'and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'  "
    ElseIf optTglTerima.value = True Then
        mstrFilter = "AND TglTerima between '" & Format(dtpAwal.value, "yyyy/MM/dd HH:mm:00") & "'and '" & Format(dtpAkhir.value, "yyyy/MM/dd HH:mm:59") & "'  "
    End If

    Call subSetGrid
    Call Isi
    refreshOtomatis = False
    MousePointer = vbDefault
    Exit Sub
errLoad:
    MousePointer = vbDefault
End Sub
Public Sub subSetGrid()
On Error GoTo gabril
    With fgDaftarDokumenRekamMedis
       ' .Clear
       If (refreshOtomatis = True) Then
       Else
           .Rows = 2
       End If
        
        .Cols = 17
        
'        .MergeCells = flexMergeFree
        
        .RowHeight(0) = 500
        
        
        
        
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "No Pendaftaran"
        .TextMatrix(0, 2) = "No.CM"
        .TextMatrix(0, 3) = "Nama Pasien"
        .TextMatrix(0, 4) = "JK"
        .TextMatrix(0, 5) = "Ruangan Pelayanan"
        .TextMatrix(0, 6) = "Tgl Masuk"
        .TextMatrix(0, 7) = "Tgl Pulang"
        .TextMatrix(0, 8) = "Tgl Kirim"
        .TextMatrix(0, 9) = "Tgl Terima"
        .TextMatrix(0, 10) = "Ruangan Pengirim"
        .TextMatrix(0, 11) = "Keterangan Kirim"
        .TextMatrix(0, 12) = "User Pengirim"
        .TextMatrix(0, 13) = "Ruangan Tujuan"
        .TextMatrix(0, 14) = "Keterangan Penerima"
        .TextMatrix(0, 15) = "User Penerima"
        .TextMatrix(0, 16) = "Status"

        
    
        .ColWidth(0) = 400
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 2500
        .ColWidth(4) = 1800
        .ColWidth(5) = 2000
        .ColWidth(6) = 2000
        .ColWidth(7) = 2000
        .ColWidth(8) = 2000
        .ColWidth(9) = 2000
        .ColWidth(10) = 1700
        .ColWidth(11) = 0
        .ColWidth(12) = 1700
        .ColWidth(13) = 1700
        .ColWidth(14) = 0
        .ColWidth(15) = 1700
        .ColWidth(16) = 4000
        

    End With

Exit Sub
gabril:
    Call msubPesanError
End Sub
Private Sub Isi()
On Error GoTo gabril
Dim i As Integer
Dim j As Integer

Set rs = Nothing
strSQL = ""
    strSQL = "SELECT distinct top 100 NoPendaftaran, NoCM, NamaPasien, JK,RuanganPelayanan, TglMasuk, TglPulang, TglKirim, TglTerima, RuanganPengirim, KeteranganKirim, UserPengirim,RuanganTujuan ,KeteranganTerima , UserPenerima,KdRuanganPelayanan " & _
    " FROM V_DokumenRekamMedisPasien " & _
    " WHERE (NamaPasien like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR RuanganPelayanan like '%" & txtParameter.Text & "%' ) " & _
    " " & mstrFilter & "" & _
    " " & mstrFilterData & " "
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
If rs.RecordCount <> 0 Then
    fgDaftarDokumenRekamMedis.Rows = rs.RecordCount + 1
     For i = 1 To rs.RecordCount
        With fgDaftarDokumenRekamMedis
            .TextMatrix(i, 0) = Chr$(187)
            .TextMatrix(i, 1) = IIf(IsNull(rs.Fields(0).value), "-", rs.Fields(0))  '
            .TextMatrix(i, 2) = IIf(IsNull(rs.Fields(1).value), "-", rs.Fields(1))  '
            .TextMatrix(i, 3) = IIf(IsNull(rs.Fields(2).value), "-", rs.Fields(2))  '
            .TextMatrix(i, 4) = IIf(IsNull(rs.Fields(3).value), "-", rs.Fields(3))  '
            .TextMatrix(i, 5) = IIf(IsNull(rs.Fields(4).value), "-", rs.Fields(4))  '
            .TextMatrix(i, 6) = IIf(IsNull(rs.Fields(5).value), "-", rs.Fields(5))  '
            .TextMatrix(i, 7) = IIf(IsNull(rs.Fields(6).value), "-", rs.Fields(6))  '
            .TextMatrix(i, 8) = IIf(IsNull(rs.Fields(7).value), "-", rs.Fields(7))  '
            .TextMatrix(i, 9) = IIf(IsNull(rs.Fields(8).value), "-", rs.Fields(8))  '
            .TextMatrix(i, 10) = IIf(IsNull(rs.Fields(9).value), "-", rs.Fields(9))  '
            .TextMatrix(i, 11) = IIf(IsNull(rs.Fields(10).value), "-", rs.Fields(10))  '
            .TextMatrix(i, 12) = IIf(IsNull(rs.Fields(11).value), "-", rs.Fields(11))  '
            .TextMatrix(i, 13) = IIf(IsNull(rs.Fields(12).value), "-", rs.Fields(12))  '
            .TextMatrix(i, 14) = IIf(IsNull(rs.Fields(13).value), "-", rs.Fields(13))  '
            .TextMatrix(i, 15) = IIf(IsNull(rs.Fields(14).value), "-", rs.Fields(14))  '
            
            If .TextMatrix(i, 8) <> "-" Then
                .Row = i
                .Col = 0
                .CellBackColor = &HFFFFFF
            End If
            
            If .TextMatrix(i, 8) <> "-" Then
                .Row = i
                .Col = 1
                .CellBackColor = &HFFFFFF
            End If

            If .TextMatrix(i, 8) <> "-" Then
                .Row = i
                .Col = 2
                .CellBackColor = &HFFFFFF
            End If

            If .TextMatrix(i, 8) <> "-" Then
                .Row = i
                .Col = 3
                .CellBackColor = &HFFFFFF
            End If

            If .TextMatrix(i, 8) <> "-" Then
                .Row = i
                .Col = 4
                .CellBackColor = &HFFFFFF
            End If

            If .TextMatrix(i, 8) <> "-" Then
                .Row = i
                .Col = 5
                .CellBackColor = &HFFFFFF
            End If

            If .TextMatrix(i, 8) <> "-" Then
                .Row = i
                .Col = 6
                .CellBackColor = &HFFFFFF
            End If

            If .TextMatrix(i, 8) <> "-" Then
                .Row = i
                .Col = 7
                .CellBackColor = &HFFFFFF
            End If

            If .TextMatrix(i, 8) <> "-" Then
                .Row = i
                .Col = 8
                .CellBackColor = &HFFFFFF
            End If

            If .TextMatrix(i, 8) <> "-" Then
                .Row = i
                .Col = 9
                .CellBackColor = &HFFFFFF
            End If

            If .TextMatrix(i, 8) <> "-" Then
                .Row = i
                .Col = 10
                .CellBackColor = &HFFFFFF
            End If

            If .TextMatrix(i, 8) <> "-" Then
                .Row = i
                .Col = 11
                .CellBackColor = &HFFFFFF
            End If

            If .TextMatrix(i, 8) <> "-" Then
                .Row = i
                .Col = 12
                .CellBackColor = &HFFFFFF
            End If

            If .TextMatrix(i, 8) <> "-" Then
                .Row = i
                .Col = 13
                .CellBackColor = &HFFFFFF
            End If

            If .TextMatrix(i, 8) <> "-" Then
                .Row = i
                .Col = 14
                .CellBackColor = &HFFFFFF
            End If

            If .TextMatrix(i, 8) <> "-" Then
                .Row = i
                .Col = 15
                .CellBackColor = &HFFFFFF
            End If


        End With
        rs.MoveNext
     Next i
End If
Exit Sub
gabril:
    Call msubPesanError

End Sub

Private Sub cmdKirim_Click()
On Error GoTo errLoad
Dim i As Integer

Set rsDokumen = Nothing
strSQL = ""

'    strSQL = "SELECT distinct top 100 NoPendaftaran, NoCM, NamaPasien, JK,RuanganPelayanan, TglMasuk, TglPulang, TglKirim, TglTerima, RuanganPengirim, KeteranganKirim, UserPengirim,RuanganTujuan ,KeteranganTerima , UserPenerima,KdRuanganTujuan " & _
             "FROM V_DokumenRekamMedisPasien " & _
             "WHERE TglMasuk between '" & Format(dtpAwal.value, "yyyy/mm/dd 00:00:00") & "' and '" & Format(dtpAkhir.value, "yyyy/mm/dd 23:59:00") & "' " & _
             "AND (NoCM like '%" & txtParameter.Text & "%' or NamaPasien like '%" & txtParameter.Text & "%' or RuanganPelayanan like '%" & txtParameter.Text & "%' and RuanganPelayanan Like '%" & dcRuangPelayanan.Text & "%') "
strSQL = "SELECT distinct top 100 NoPendaftaran, NoCM, NamaPasien, JK,RuanganPelayanan, TglMasuk, TglPulang, TglKirim, TglTerima, RuanganPengirim, KeteranganKirim, UserPengirim,RuanganTujuan ,KeteranganTerima , UserPenerima,KdRuanganPelayanan " & _
    " FROM V_DokumenRekamMedisPasien " & _
    " WHERE (NamaPasien like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR RuanganPelayanan like '%" & txtParameter.Text & "%' ) " & _
    " " & mstrFilter & "" & _
    " " & mstrFilterData & " "
    
Set rsDokumen = Nothing
Call msubRecFO(rsDokumen, strSQL)

If rsDokumen.EOF = True Then
    fgDaftarDokumenRekamMedis.TextMatrix(i, 16) = "Data tidak ada"
Else
    For i = 1 To rsDokumen.RecordCount
        Set rs = Nothing
        strSQL = "SELECT distinct top 100 NoPendaftaran, NoCM, NamaPasien, JK,RuanganPelayanan, TglMasuk, TglPulang, TglKirim, TglTerima, RuanganPengirim, KeteranganKirim, UserPengirim,RuanganTujuan ,KeteranganTerima , UserPenerima,KdRuanganTujuan,KdRuanganPelayanan " & _
                 " FROM V_DokumenRekamMedisPasien " & _
                 " WHERE NoPendaftaran = '" & fgDaftarDokumenRekamMedis.TextMatrix(i, 1) & "' and TglMasuk between '" & Format(dtpAwal.value, "yyyy/mm/dd 00:00:00") & "' and '" & Format(dtpAkhir.value, "yyyy/mm/dd 23:59:00") & "'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
                If rs.Fields("TglKirim") <> "Null" Then
                    mstrTglKirim = rs.Fields("TglKirim")
                End If
                If fgDaftarDokumenRekamMedis.TextMatrix(i, 0) = Chr$(187) And fgDaftarDokumenRekamMedis.TextMatrix(i, 8) = "-" Then
                    If rs.Fields("RuanganPengirim") <> mstrNamaRuangan Or IsNull(rs.Fields("RuanganPengirim")) Then
                        Call sp_KirimTerimaDokumenRekamMedis(dbcmd)
                        fgDaftarDokumenRekamMedis.TextMatrix(i, 16) = "Dokumen Terkirim"
                    Else
                        fgDaftarDokumenRekamMedis.TextMatrix(i, 16) = "Dokumen tidak bisa dikirim"
                    End If
                Else
                    fgDaftarDokumenRekamMedis.TextMatrix(i, 16) = "Dokumen Sudah di kirim Sebelumnya"
                End If
        rs.MoveNext
        End If
    Next i
Call Add_HistoryLoginActivity("Add_KirimTerimaDokumenRekamMedisPasien")
refreshOtomatis = True

Call cmdCari_Click

End If
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdKirimTerima_Click()
Dim i As Integer

Set rsDokumen = Nothing
strSQL = ""

    strSQL = "SELECT distinct top 100 NoPendaftaran, NoCM, NamaPasien, JK,RuanganPelayanan, TglMasuk, TglPulang, TglKirim, TglTerima, RuanganPengirim, KeteranganKirim, UserPengirim,RuanganTujuan ,KeteranganTerima , UserPenerima,KdRuanganTujuan " & _
    " FROM V_DokumenRekamMedisPasien " & _
    " WHERE NoPendaftaran = '" & fgDaftarDokumenRekamMedis.TextMatrix(fgDaftarDokumenRekamMedis.Row, 0) & "' and (NamaPasien like '%" & fgDaftarDokumenRekamMedis.TextMatrix(fgDaftarDokumenRekamMedis.Row, 2) & "%' And NoCM like '%" & fgDaftarDokumenRekamMedis.TextMatrix(fgDaftarDokumenRekamMedis.Row, 1) & "%' And RuanganPelayanan like '%" & fgDaftarDokumenRekamMedis.TextMatrix(fgDaftarDokumenRekamMedis.Row, 4) & "%' ) " & _
    " " & mstrFilter & "" & _
    " " & mstrFilterData & " "
    Set rsDokumen = Nothing
Call msubRecFO(rsDokumen, strSQL)

         If rsDokumen.EOF = False Then
            rsDokumen.MoveLast
            If rsDokumen.Fields("tglKirim") <> fgDaftarDokumenRekamMedis.TextMatrix(fgDaftarDokumenRekamMedis.Row, 7) Then
                fgDaftarDokumenRekamMedis.TextMatrix(fgDaftarDokumenRekamMedis.Row, 16) = "Dukumen Sudah pernah Dikirim dari ruangan Ini"
                Exit Sub
            End If
            
            If IsNull(rsDokumen.Fields("tglterima")) Then
            
                    Call sp_KirimTerimaDokumenRekamMedis(dbcmd)

            Else
                 If rsDokumen.Fields("KdruanganTujuan") = mstrKdRuangan Then
                 
                     Call sp_KirimTerimaDokumenRekamMedis(dbcmd)
                Else
                   fgDaftarDokumenRekamMedis.TextMatrix(fgDaftarDokumenRekamMedis.Row, 16) = "Dukumen tidak dapat Di kirim Atau di terima di ruangan Ini"
                   Exit Sub
                End If
            End If
        End If
End Sub
Private Sub sp_KirimTerimaDokumenRekamMedis(ByVal adoCommand As ADODB.Command)
    Set dbcmd = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, rs.Fields("NoPendaftaran"))
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, rs.Fields("NoCM"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan) 'ruangan pengirim
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, rs.Fields("KdRuanganPelayanan")) 'ruangan penerima
        If IsNull(rs.Fields("TglKirim")) Then
            .Parameters.Append .CreateParameter("TglKirim", adDate, adParamInput, , Format(dtpTglKirim.value, "yyyy/MM/dd HH:mm:ss"))
        Else
            .Parameters.Append .CreateParameter("TglKirim", adDate, adParamInput, , mstrTglKirim)
        End If
        .Parameters.Append .CreateParameter("TglTerima", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("IdUserKirim", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("IdUserTerima", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("KeteranganKirim", adVarChar, adParamInput, 200, Null)
        .Parameters.Append .CreateParameter("KeteranganTerima", adVarChar, adParamInput, 200, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_KirimTerimaDokumenRekamMedisPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan dokumen rekam medis Pasien", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

Private Sub cmdTerima_Click()
On Error GoTo errLoad
Dim i As Integer

Set rsDokumen = Nothing
strSQL = ""

'    strSQL = "SELECT distinct top 100 NoPendaftaran, NoCM, NamaPasien, JK,RuanganPelayanan, TglMasuk, TglPulang, TglKirim, TglTerima, RuanganPengirim, KeteranganKirim, UserPengirim,RuanganTujuan ,KeteranganTerima , UserPenerima,KdRuanganTujuan,KdRuanganPelayanan " & _
'             " FROM V_DokumenRekamMedisPasien " & _
'             " WHERE NoCM like '%" & txtParameter.Text & "%' and NamaPasien like '%" & txtParameter.Text & "%' and RuanganPelayanan like '%" & txtParameter.Text & "%' And RuanganPelayanan Like '%" & dcRuangPelayanan.Text & "%' " & _
'             " And TglMasuk between '" & Format(dtpAwal.value, "yyyy/mm/dd 00:00:00") & "' and '" & Format(dtpAkhir.value, "yyyy/mm/dd 23:59:00") & "' "
    strSQL = "SELECT distinct top 100 NoPendaftaran, NoCM, NamaPasien, JK,RuanganPelayanan, TglMasuk, TglPulang, TglKirim, TglTerima, RuanganPengirim, KeteranganKirim, UserPengirim,RuanganTujuan ,KeteranganTerima , UserPenerima,KdRuanganPelayanan " & _
              " FROM V_DokumenRekamMedisPasien " & _
              " WHERE (NamaPasien like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%' OR RuanganPelayanan like '%" & txtParameter.Text & "%' ) " & _
              " " & mstrFilter & "" & _
              " " & mstrFilterData & " "
    Set rsDokumen = Nothing
Call msubRecFO(rsDokumen, strSQL)

If rsDokumen.EOF = True Then
    fgDaftarDokumenRekamMedis.TextMatrix(i, 16) = "Data tidak ada"
Else
    For i = 1 To rsDokumen.RecordCount
        Set rs = Nothing
        strSQL = "SELECT distinct top 100 NoPendaftaran, NoCM, NamaPasien, JK,RuanganPelayanan, TglMasuk, TglPulang, TglKirim, TglTerima, RuanganPengirim, KeteranganKirim, UserPengirim,RuanganTujuan ,KeteranganTerima , UserPenerima,KdRuanganTujuan,KdRuanganPelayanan " & _
                 " FROM V_DokumenRekamMedisPasien " & _
                 " WHERE NoPendaftaran = '" & fgDaftarDokumenRekamMedis.TextMatrix(i, 1) & "' and TglMasuk between '" & Format(dtpAwal.value, "yyyy/mm/dd 00:00:00") & "' and '" & Format(dtpAkhir.value, "yyyy/mm/dd 23:59:00") & "'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            If fgDaftarDokumenRekamMedis.TextMatrix(i, 0) = Chr$(187) And fgDaftarDokumenRekamMedis.TextMatrix(i, 8) <> "-" Then
                If fgDaftarDokumenRekamMedis.TextMatrix(i, 13) = mstrNamaRuangan Then
                    dbConn.Execute "Update DokumenRekamMedisPasien Set TglTerima = '" & Format(dtpTglTerima.value, "yyyy/mm/dd HH:mm:ss") & "' Where NoPendaftaran = '" & fgDaftarDokumenRekamMedis.TextMatrix(i, 1) & "'"
                        If fgDaftarDokumenRekamMedis.TextMatrix(fgDaftarDokumenRekamMedis.Row, 8) = "-" Then
                            fgDaftarDokumenRekamMedis.TextMatrix(i, 16) = "Dokumen belum dikirim"
                        Else
                            fgDaftarDokumenRekamMedis.TextMatrix(i, 16) = "Dokumen Berhasil Diterima"
                        End If
                Else
                    fgDaftarDokumenRekamMedis.TextMatrix(i, 16) = "Dokumen tidak bisa diterima di ruangan ini"
                End If
            End If
        rs.MoveNext
        End If
    Next i
    
Call Add_HistoryLoginActivity("Add_KirimTerimaDokumenRekamMedisPasien")
refreshOtomatis = True
Call cmdCari_Click

End If
Exit Sub
errLoad:
    Call msubPesanError



End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcRuanganPengirim_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcRuanganPengirim.MatchedWithList = True Then optTglMasuk.SetFocus
        strSQL = "SELECT distinct KdRuanganPengirim, RuanganPengirim FROM V_DokumenRekamMedisPasien  where (RuanganPengirim LIKE '%" & dcRuanganPengirim.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcRuanganPengirim.Text = ""
            optTglMasuk.SetFocus
            Exit Sub
        End If
        dcRuanganPengirim.BoundText = rs(0).value
        dcRuanganPengirim.Text = rs(1).value
    End If
End Sub

Private Sub dcRuangPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcRuangPelayanan.MatchedWithList = True Then dcRuanganPengirim.SetFocus
        strSQL = "SELECT distinct KdRuanganPelayanan, RuanganPelayanan FROM  V_DokumenRekamMedisPasien  where (RuanganPelayanan LIKE '%" & dcRuangPelayanan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcRuangPelayanan.Text = ""
            dcRuanganPengirim.SetFocus
            Exit Sub
        End If
        dcRuangPelayanan.BoundText = rs(0).value
        dcRuangPelayanan.Text = rs(1).value
    End If
End Sub



Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub fgDaftarDokumenRekamMedis_Click()
On Error GoTo hell
    If fgDaftarDokumenRekamMedis.Rows = 1 Then Exit Sub
    If fgDaftarDokumenRekamMedis.Col <> 0 Then Exit Sub
        chkCheck.Visible = True
        chkCheck.Top = fgDaftarDokumenRekamMedis.RowPos(fgDaftarDokumenRekamMedis.Row) + 975
        Dim intChk As Integer
        intChk = ((fgDaftarDokumenRekamMedis.ColPos(fgDaftarDokumenRekamMedis.Col + 1) - fgDaftarDokumenRekamMedis.ColPos(fgDaftarDokumenRekamMedis.Col)) / 2)
        chkCheck.Left = fgDaftarDokumenRekamMedis.ColPos(fgDaftarDokumenRekamMedis.Col) + intChk - 20 ' - 250  '+ intChk
        chkCheck.SetFocus
    If fgDaftarDokumenRekamMedis.Col <> 0 Then
        If fgDaftarDokumenRekamMedis.TextMatrix(fgDaftarDokumenRekamMedis.Row, 0) <> "" Then
            chkCheck.value = 1
        Else
            chkCheck.value = 0
        End If
    End If
    Exit Sub
hell:
    Call msubPesanError

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    mdTglAwal = dtpAwal.value
    mdTglAkhir = dtpAkhir.value

    If KeyCode = vbKeyF1 Then
        Call cmdCari_Click
        frmCetakDaftarDokumenRekamMedisPasien.Show
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call subLoadDcSource
    dtpAwal.value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.value = Now
    dtpTglKirim = Now
    dtpTglTerima = Now
    optTglMasuk.value = True
    Call cmdCari_Click
    mblnForm = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mblnForm = False
End Sub

Private Sub txtParameter_Change()
    Call cmdCari_Click
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call cmdCari_Click
'            cmdKirimTerima.SetFocus
    End If
End Sub

