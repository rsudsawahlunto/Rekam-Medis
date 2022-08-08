VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNilaiPersediaanNM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Nilai Persediaan"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNilaiPersediaanNM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   14835
   Begin MSDataListLib.DataCombo dcJenisBarang 
      Height          =   390
      Left            =   12840
      TabIndex        =   15
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   688
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Appearance      =   0
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox ChkPerjenis 
      Caption         =   "Perjenis Barang"
      ForeColor       =   &H80000006&
      Height          =   255
      Left            =   12840
      TabIndex        =   14
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtTotalStokReal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8880
      TabIndex        =   10
      Text            =   "999.999"
      Top             =   7920
      Width           =   1095
   End
   Begin VB.TextBox txtTotalReal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10080
      TabIndex        =   9
      Text            =   "999.999.999.999,99"
      Top             =   7920
      Width           =   2655
   End
   Begin MSDataListLib.DataCombo dcNoClosing 
      Height          =   5430
      Left            =   12840
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   9578
      _Version        =   393216
      Style           =   1
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtcariNamaBarang 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   7920
      Width           =   3015
   End
   Begin VB.TextBox txtCariJenisBarang 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   7920
      Width           =   3015
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   735
      Left            =   12855
      TabIndex        =   3
      Top             =   7680
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   10821
      _Version        =   393216
      FixedCols       =   0
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   8460
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   26114
            Text            =   "Cetak Nilai Persediaan (F1)"
            TextSave        =   "Cetak Nilai Persediaan (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   13044
            Text            =   "Ctrl C : Copy Stok System To Stok Real"
            TextSave        =   "Ctrl C : Copy Stok System To Stok Real"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   16
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
      Left            =   12960
      Picture         =   "frmNilaiPersediaanNM.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmNilaiPersediaanNM.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmNilaiPersediaanNM.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nilai Persediaan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   10080
      TabIndex        =   12
      Top             =   7680
      Width           =   1410
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Stok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   8880
      TabIndex        =   11
      Top             =   7680
      Width           =   960
   End
   Begin VB.Label lblJumlahData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data 0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Closing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   12840
      TabIndex        =   7
      Top             =   1800
      Width           =   1950
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Nama Barang"
      Height          =   210
      Index           =   13
      Left            =   3240
      TabIndex        =   5
      Top             =   7680
      Width           =   1410
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Jenis Barang"
      Height          =   210
      Index           =   12
      Left            =   120
      TabIndex        =   4
      Top             =   7680
      Width           =   1350
   End
End
Attribute VB_Name = "frmNilaiPersediaanNM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrJmlStokReal() As Long
Dim arrTotal() As Currency
Dim i As Integer

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    strSQL = "SELECT DISTINCT  TglClosing, TglClosing AS Alias FROM V_DataStokBarangNonMedisRekapx WHERE KdRuangan = '" & mstrKdRuangan & "' ORDER BY TglClosing"
    Call msubDcSource(dcNoClosing, rs, strSQL)

    strSQL = "SELECT KdDetailJenisBarang, DetailJenisBarang FROM V_S_DetailJenisBarangNonMedis Order By DetailJenisBarang"
    Call msubDcSource(dcJenisBarang, rs, strSQL)

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()
    On Error GoTo errLoad
    Dim i As Integer

    If ChkPerjenis.value = 0 Then
'        strsql = "SELECT * FROM V_DataStokBarangNonMedisRekapx " & _
'        " WHERE ((NoRegisterAsset <> '0000000' AND NoRegisterAsset <> '0' AND NoRegisterAsset <>'000000') OR KdJenisAset is null) and KdRuangan = '" & mstrKdRuangan & "' AND (TglClosing = '" & Format(dcNoClosing.BoundText, "yyyy/MM/dd hh:mm:ss") & "') AND StokReal<> 0" & _
'        " ORDER By JenisBarang, NamaBarang"
        strSQL = "SELECT * FROM V_DataStokBarangNonMedisRekapx " & _
        " WHERE KdRuangan = '" & mstrKdRuangan & "' AND (TglClosing = '" & Format(dcNoClosing.BoundText, "yyyy/MM/dd hh:mm:ss") & "') AND StokReal<> 0" & _
        " ORDER By JenisBarang, NamaBarang"
    Else
        If Periksa("datacombo", dcJenisBarang, "Jenis barang kosong") = False Then Exit Sub
'        strsql = "SELECT * FROM V_DataStokBarangNonMedisRekapx " & _
'        " WHERE ((NoRegisterAsset <> '0000000' AND NoRegisterAsset <> '0' AND NoRegisterAsset <>'000000') OR KdJenisAset is null) and JenisBarang like '" & dcJenisBarang.Text & "%' and KdRuangan = '" & mstrKdRuangan & "' AND (TglClosing = '" & Format(dcNoClosing.BoundText, "yyyy/MM/dd HH:mm:ss") & "')" & _
'        " ORDER By JenisBarang, NamaBarang"
        strSQL = "SELECT * FROM V_DataStokBarangNonMedisRekapx " & _
        " WHERE JenisBarang like '" & dcJenisBarang.Text & "%' and KdRuangan = '" & mstrKdRuangan & "' AND (TglClosing = '" & Format(dcNoClosing.BoundText, "yyyy/MM/dd HH:mm:ss") & "')" & _
        " ORDER By JenisBarang, NamaBarang"
    End If

    Call msubRecFO(rs, strSQL)
    Call subSetGrid

    If rs.EOF = True Then Exit Sub
    
    lblJumlahData.Caption = "Data 0/" & Format(rs.RecordCount, "#,###")
    MousePointer = vbHourglass
    txtTotalStokReal.Text = 0
    txtTotalReal.Text = 0
    For i = 1 To rs.RecordCount
        With fgData
            .TextMatrix(i, 0) = rs("JenisBarang")
            .TextMatrix(i, 1) = rs("NamaBarang")
            .TextMatrix(i, 2) = ""
            .TextMatrix(i, 3) = rs("AsalBarang")
            .TextMatrix(i, 4) = rs("StokSystem")
            .TextMatrix(i, 5) = rs("StokReal")
            txtTotalStokReal.Text = CCur(txtTotalStokReal.Text) + rs("StokReal")

            Set dbRst = Nothing
'            strsql = "SELECT Discount FROM HargaNettoBarangNonMedis WHERE KdBarang='" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "'"
'            Call msubRecFO(dbRst, strsql)

            .TextMatrix(i, 6) = ""
            .TextMatrix(i, 7) = Format(rs("HargaNetto"), "#,###.00")
            .TextMatrix(i, 8) = 0
            .TextMatrix(i, 9) = Format(rs("Discount"), "#,###")
            .TextMatrix(i, 10) = Format((rs("StokReal") * rs("HargaNetto")), "#,###.00")
            txtTotalReal.Text = CCur(txtTotalReal.Text) + (rs("StokReal") * rs("HargaNetto"))
            .TextMatrix(i, 11) = rs("KdBarang")
            .TextMatrix(i, 12) = rs("KdAsal")

            Set dbRst = Nothing
            strSQL = "SELECT Lokasi FROM StokBarangNonMedis WHERE KdBarang='" & rs("KdBarang") & "' AND KdAsal='" & rs("KdAsal") & "' AND KdRuangan='" & mstrKdRuangan & "'"
            Call msubRecFO(dbRst, strSQL)
            If dbRst.EOF = True Then
            .TextMatrix(i, 13) = ""
            Else
            .TextMatrix(i, 13) = IIf(IsNull(dbRst("Lokasi")), "", dbRst("Lokasi"))
            End If
            .TextMatrix(i, 14) = IIf(IsNull(rs("NoFIFO")), "", rs("NoFIFO"))
            .TextMatrix(i, 15) = IIf(IsNull(rs("NoRegisterAsset")), "", rs("NoRegisterAsset"))
        End With
        rs.MoveNext
        fgData.Rows = fgData.Rows + 1
    Next i
    txtTotalStokReal.Text = Format(txtTotalStokReal.Text, "#,##0")
    txtTotalReal.Text = Format(txtTotalReal.Text, "#,###.00")

    MousePointer = vbDefault
    fgData.Rows = fgData.Rows - 1

    Exit Sub
errLoad:
    MousePointer = vbDefault
    Call msubPesanError
End Sub

Private Sub subSetGrid()
    Dim i As Integer
    With fgData
        .Cols = 16
        .Rows = 2

        .RowHeight(0) = 500
        For i = 0 To .Cols - 1
            .Col = i
            .Row = 0
            .CellBackColor = Me.BackColor: .CellAlignment = flexAlignCenterCenter
        Next i

        .TextMatrix(0, 0) = "Jenis Barang"
        .TextMatrix(0, 1) = "Nama Barang"
        .TextMatrix(0, 2) = "Kekuatan"
        .TextMatrix(0, 3) = "Asal Barang"
        .TextMatrix(0, 4) = "Stok System"
        .TextMatrix(0, 5) = "Stok Real"
        .TextMatrix(0, 6) = "Tgl Kadaluarsa"
        .TextMatrix(0, 7) = "Harga Netto"
        .TextMatrix(0, 8) = "Harga Netto 2"
        .TextMatrix(0, 9) = "Disc (%)"
        .TextMatrix(0, 10) = "Total"

        .TextMatrix(0, 11) = "KdBarang"
        .TextMatrix(0, 12) = "KdAsal"
        .TextMatrix(0, 13) = "Lokasi"
        .TextMatrix(0, 14) = "NoTerima"
        .TextMatrix(0, 15) = "NoRegisterAsset"
        
        .ColWidth(0) = 1500
        .ColWidth(1) = 3000
        .ColWidth(2) = 0
        .ColWidth(3) = 1100
        .ColWidth(4) = 900
        .ColWidth(5) = 900
        .ColWidth(6) = 0
        .ColWidth(7) = 1400
        .ColWidth(8) = 0
        .ColWidth(9) = 0 '1000
        .ColWidth(10) = 1450
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0 ' 1450
        .ColWidth(14) = 1200
        .ColWidth(15) = 1750
        
    End With
End Sub

Private Sub subSetGridNM()
    Dim i As Integer
    With fgData
        .Cols = 17
        .Rows = 2

        .RowHeight(0) = 500
        For i = 0 To .Cols - 1
            .Col = i
            .Row = 0
            .CellBackColor = Me.BackColor: .CellAlignment = flexAlignCenterCenter
        Next i

        .TextMatrix(0, 0) = "Jenis Barang"
        .TextMatrix(0, 1) = "Nama Barang"
        .TextMatrix(0, 2) = "Asal Barang"
        .TextMatrix(0, 3) = "Merk"
        .TextMatrix(0, 4) = "Type"
        .TextMatrix(0, 5) = "Bahan"
        .TextMatrix(0, 6) = "Stok System"
        .TextMatrix(0, 7) = "Stok Real"
        .TextMatrix(0, 8) = "Harga Netto"
        .TextMatrix(0, 9) = "Disc (%)"
        .TextMatrix(0, 10) = "Total"
        .TextMatrix(0, 11) = "KdBarang"
        .TextMatrix(0, 12) = "KdAsal"
        .TextMatrix(0, 13) = "KdMerk"
        .TextMatrix(0, 14) = "KdType"
        .TextMatrix(0, 15) = "KdBahanBarang"
        .TextMatrix(0, 16) = "Lokasi"

        .ColWidth(0) = 1500
        .ColWidth(1) = 2500
        .ColWidth(2) = 1000
        .ColWidth(3) = 1100
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100
        .ColWidth(6) = 900
        .ColWidth(7) = 900
        .ColWidth(8) = 1100
        .ColWidth(9) = 0 '1000
        .ColWidth(10) = 1300
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        .ColWidth(16) = 0
    End With
End Sub

Private Sub ChkPerjenis_Click()
    If ChkPerjenis.value = 1 Then
        ChkPerjenis.ForeColor = &HFF0000
        dcJenisBarang.Enabled = True
    Else
        ChkPerjenis.ForeColor = &H80000006
        dcJenisBarang.Enabled = False
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcJenisBarang_Change()
    On Error GoTo errLoad

    If dcJenisBarang.BoundText = "" Then Exit Sub
    If Periksa("datacombo", dcNoClosing, "Pilih Nomor Closing") = False Then Exit Sub
    Call subLoadGridSource

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcNoClosing_Click(Area As Integer)
    On Error GoTo errLoad

    If Periksa("datacombo", dcNoClosing, "Pilih Nomor Closing") = False Then Exit Sub
    Call subLoadGridSource

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub fgData_RowColChange()
    On Error GoTo errLoad

    lblJumlahData.Caption = "Data " & Format(fgData.Row, "#,###") & "/" & Format(fgData.Rows - 1, "#,###")

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        mdtglclosing = Format(dcNoClosing.BoundText, "yyyy/MM/dd hh:mm:ss")
        frmCetakNilaiPersediaanNM.Show
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)

    Call subLoadDcSource

    Call subSetGrid
    Call subLoadGridSource

    txtCariJenisBarang.Text = ""
    txtcariNamaBarang.Text = ""
End Sub

Private Sub txtCariJenisBarang_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    If KeyAscii = 13 Then
        With fgData
            .Row = 1
            .Col = 0

            For i = 1 To .Rows - 1
                If UCase(Left(txtCariJenisBarang.Text, Len(txtCariJenisBarang.Text))) = UCase(Left(fgData.TextMatrix(i, 0), Len(txtCariJenisBarang.Text))) Then Exit For
            Next i
            .Row = i:  .SetFocus
            SendKeys ("{DOWN}"): SendKeys ("{UP}")
        End With
    End If
    Exit Sub
errLoad:
End Sub

Private Sub txtcariNamaBarang_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        With fgData
            .Row = 1
            .Col = 0

            For i = 1 To .Rows - 2
                If UCase(Left(txtcariNamaBarang.Text, Len(txtcariNamaBarang.Text))) = UCase(Left(fgData.TextMatrix(i, 1), Len(txtcariNamaBarang.Text))) Then Exit For
            Next i
            .TopRow = i: .Row = i: .Col = 1: .SetFocus
        End With
    End If
End Sub



