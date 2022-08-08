VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPemesananBarang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pemesanan Barang"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPemesananBarang.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8925
   Begin VB.TextBox txtNamaFormPengirim 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   5760
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid dgObatAlkes 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
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
            LCID            =   1057
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
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame0 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   14
      Top             =   1920
      Width           =   8895
      Begin VB.TextBox txtNamaBarang 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtSatuan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtIsi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtKdBarang 
         Height          =   315
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtStock 
         Alignment       =   2  'Center
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   420
         Left            =   7815
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdtambah 
         Caption         =   "&Tambah"
         Height          =   420
         Left            =   6840
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtJumlah 
         Alignment       =   2  'Center
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
         Left            =   4920
         MaxLength       =   4
         TabIndex        =   6
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   2655
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4683
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
         Appearance      =   0
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
         Height          =   210
         Index           =   4
         Left            =   5760
         TabIndex        =   24
         Top             =   240
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   3960
         TabIndex        =   18
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   5160
         TabIndex        =   15
         Top             =   120
         Width           =   300
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Order"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   20
      Top             =   960
      Width           =   8895
      Begin MSComCtl2.DTPicker dtpTglOrder 
         Height          =   330
         Left            =   2040
         TabIndex        =   1
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
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
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   118816771
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.TextBox txtNoOrder 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   120
         MaxLength       =   15
         TabIndex        =   0
         Top             =   480
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo dcRuanganTujuan 
         Height          =   330
         Left            =   6240
         TabIndex        =   3
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo dcStatusBarang 
         Height          =   330
         Left            =   4080
         TabIndex        =   2
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Barang"
         Height          =   210
         Index           =   2
         Left            =   4080
         TabIndex        =   26
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Order"
         Height          =   210
         Index           =   1
         Left            =   2040
         TabIndex        =   23
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Order"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   750
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan Tujuan Pesanan"
         Height          =   210
         Index           =   10
         Left            =   6240
         TabIndex        =   21
         Top             =   240
         Width           =   2070
      End
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   5775
      TabIndex        =   12
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   7350
      TabIndex        =   13
      Top             =   5760
      Width           =   1455
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   27
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
      Left            =   7080
      Picture         =   "frmPemesananBarang.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPemesananBarang.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPemesananBarang.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmPemesananBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strKdAsal As String
Dim i As Integer

Private Sub cmdBatal_Click()
    Call subKosong
    Call subLoadDcSource
    Call subSetGrid
    dtpTglOrder.SetFocus
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo Errload
    
    If fgData.TextMatrix(1, 0) = "" Then MsgBox "Data barang harus diisi", vbExclamation, "Validasi": Exit Sub

    If sp_StrukOrder() = False Then Exit Sub
    For i = 1 To fgData.Rows - 2
        With fgData
            If dcStatusBarang.BoundText = "02" Then
                If sp_DetailOrderRuangan(.TextMatrix(i, 4), .TextMatrix(i, 2), .TextMatrix(i, 6), "A") = False Then Exit Sub
            ElseIf dcStatusBarang.BoundText = "01" Then
                If sp_DetailOrderRuanganNonMedis(.TextMatrix(i, 4), .TextMatrix(i, 2), "A") = False Then Exit Sub
            End If
        End With
    Next i
    MsgBox "No Order : " & txtNoOrder.Text, vbInformation, "Informasi"

    If dcStatusBarang.BoundText = "02" Then
        Call Add_HistoryLoginActivity("Add_StrukOrder+Add_DetailOrderRuangan")
    ElseIf dcStatusBarang.BoundText = "01" Then
        Call Add_HistoryLoginActivity("Add_StrukOrder+Add_DetailOrderRuanganNonMedis")
    End If
    Call cmdBatal_Click

    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub cmdTambah_Click()
    On Error GoTo Errload
    Dim i As Integer

    If Periksa("text", txtNamaBarang, "Nama barang kosong") = False Then Exit Sub
    If Periksa("nilai", txtJumlah, "Jumlah barang kosong") = False Then Exit Sub
'cek max stok
If dcStatusBarang.BoundText = "02" Then
    If bolStatusFIFO = False Then
        strSQL = "select TOP 1 * from stokruangan where KdRuangan = '" & mstrKdRuangan & "' and  KdBarang = '" & txtKdBarang.Text & "' and kdAsal = '" & strKdAsal & "'"
    Else
        strSQL = "select TOP 1 * from stokruanganFIFO where KdRuangan = '" & mstrKdRuangan & "' and  KdBarang = '" & txtKdBarang.Text & "' and kdAsal = '" & strKdAsal & "' " 'and noTerima = '" & strNoTerima & "'"
    End If
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
           If Not IsNull(rs("JmlMax")) Then
           If rs("JmlMax") <> "0" Then
               If Val(txtJumlah.Text) > Val(rs("JmlMax")) Then
               
                   MsgBox " Stok melebihi Jumlah Max Pemesanan", vbInformation
                   Exit Sub
               End If
            End If
           End If
        End If
End If
'end cek
    With fgData
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 4) = txtKdBarang.Text Then
                MsgBox txtNamaBarang.Text & " sudah diinput", vbExclamation, "Validasi"
                txtNamaBarang.SetFocus
                Exit Sub
            End If
        Next i
    End With

    With fgData
        .TextMatrix(.Rows - 1, 0) = txtNamaBarang.Text
        .TextMatrix(.Rows - 1, 1) = txtStock.Text
        .TextMatrix(.Rows - 1, 2) = CDbl(txtJumlah.Text)
        .TextMatrix(.Rows - 1, 3) = txtSatuan.Text
        .TextMatrix(.Rows - 1, 4) = txtKdBarang.Text
        .TextMatrix(.Rows - 1, 5) = txtSatuan.Text
        .TextMatrix(.Rows - 1, 6) = strKdAsal
        
        .Rows = .Rows + 1
    End With

    txtKdBarang.Text = ""
    txtSatuan.Text = ""

    txtNamaBarang.Text = ""
    txtStock.Text = 0
    txtJumlah.Text = 0
    strKdAsal = ""
    strNoTerima = ""
    dgObatAlkes.Visible = False
    txtNamaBarang.SetFocus
    Exit Sub
Errload:
    Call msubPesanError
    'Resume 0
End Sub

Private Sub Cmdtambah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaBarang.SetFocus
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo Errload
    Dim i As Integer
    With fgData
        If .Row = .Rows Then Exit Sub
        If .Row = 0 Then Exit Sub

        If .Rows = 2 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
            Next i
            Exit Sub
        Else
            .RemoveItem .Row
        End If
    End With

    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dcRuanganTujuan_GotFocus()
    On Error GoTo Errload
    Dim tempKode As String

    tempKode = dcRuanganTujuan.BoundText
    Call msubDcSource(dcRuanganTujuan, rs, "SELECT KdRuangan, NamaRuangan FROM V_StrukOrderRuanganTujuan WHERE KdKelompokBarang='" & dcStatusBarang.BoundText & "' AND StatusEnabled='1' and KdRuangan<>'" & mstrKdRuangan & "' ORDER BY NamaRuangan")
    dcRuanganTujuan.BoundText = tempKode
    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dcStatusBarang_Change()
    dcRuanganTujuan.BoundText = ""
    fgData.Clear
End Sub

Private Sub dcStatusBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcRuanganTujuan.SetFocus
    On Error GoTo Errload

    If KeyAscii = 13 Then
        If Len(Trim(dcStatusBarang.Text)) = 0 Then dcRuanganTujuan.SetFocus: Exit Sub
        If dcStatusBarang.MatchedWithList = True Then dcRuanganTujuan.SetFocus
        Call msubRecFO(dbRst, "SELECT KdKelompokBarang, KelompokBarang FROM KelompokBarang WHERE (KelompokBarang = '" & dcStatusBarang.BoundText & "') and StatusEnabled=1")
        If dbRst.EOF = True Then Exit Sub
        dcStatusBarang.BoundText = dbRst(0).value
        dcStatusBarang.Text = dbRst(1).value
    End If
    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dcRuanganTujuan_KeyPress(KeyAscii As Integer)
    On Error GoTo Errload

    If KeyAscii = 13 Then
        If Len(Trim(dcRuanganTujuan.Text)) = 0 Then txtNamaBarang.SetFocus: Exit Sub
        If dcRuanganTujuan.MatchedWithList = True Then txtNamaBarang.SetFocus
        Call msubRecFO(dbRst, "SELECT KdRuangan, NamaRuangan FROM V_StrukOrderRuanganTujuan WHERE KdKelompokBarang='" & dcStatusBarang.BoundText & "' AND (NamaRuangan LIKE '%" & dcRuanganTujuan.BoundText & "%') and StatusEnabled='1' and KdRuangan<>'" & mstrKdRuangan & "' ORDER BY NamaRuangan")
        If dbRst.EOF = True Then Exit Sub
        dcRuanganTujuan.BoundText = dbRst(0).value
        dcRuanganTujuan.Text = dbRst(1).value
    End If
    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dgObatAlkes_DblClick()
    On Error GoTo Errload
    
    
    If dcStatusBarang.BoundText = "02" Then
'        txtStock.Text = IIf(dgObatAlkes.Columns("JmlStok").Value = "", "0", dgObatAlkes.Columns("JmlStok").Value)
       
        'While (rs.EOF)
           
        'End
        'MsgBox dgObatAlkes.Row
        'dgObatAlkes.
        txtKdBarang.Text = dgObatAlkes.Columns("KdBarang").value
        txtSatuan.Text = dgObatAlkes.Columns("Satuan").value
        
        strKdAsal = IIf(dgObatAlkes.Columns("kdAsal").value = "", "", dgObatAlkes.Columns("kdAsal").value)
         'chandra 03 03 2014
        Dim jumlahStokOnHand As Long
        jumlahStokOnHand = 0
        If (bolStatusFIFO = True) Then
            strSQL = "select JmlStokOnHand from stokRuanganFIFO where kdBarang='" & dgObatAlkes.Columns("KdBarang").value & "' and kdRUangan='" & dcRuanganTujuan.BoundText & "' and kdasal='" & dgObatAlkes.Columns("kdAsal") & "' "
        ElseIf (bolStatusFIFO = False) Then
            strSQL = "select JmlStokOnHand from stokRuangan where kdBarang='" & dgObatAlkes.Columns("KdBarang").value & "' and kdRUangan='" & dcRuanganTujuan.BoundText & "' and kdasal='" & dgObatAlkes.Columns("kdAsal") & "'"
        End If
        
        
        Call msubRecFO(rs, strSQL)
        For i = 1 To rs.RecordCount
         jumlahStokOnHand = jumlahStokOnHand + rs("JmlStokOnHand").value
         rs.MoveNext
         Next i
         txtStock.Text = IIf(dgObatAlkes.Columns("JmlStokOnHand").value = "", "0", jumlahStokOnHand)
         txtNamaBarang.Text = dgObatAlkes.Columns("Nama Barang").value
        'strNoTerima = dgObatAlkes.Columns("NoTerima").Value
    Else
        If (dgObatAlkes.Columns("JmlStok").value <> "" And Val(dgObatAlkes.Columns("JmlStok").value) <= 0) Then
            MsgBox "Barang Tidak Mencukupi untuk pemensanan"
            dgObatAlkes.SetFocus
            Exit Sub
        End If
        txtStock.Text = IIf(dgObatAlkes.Columns("JmlStok").value = "", "0", dgObatAlkes.Columns("JmlStok").value)
        txtKdBarang.Text = dgObatAlkes.Columns("KdBarang")
        txtSatuan.Text = dgObatAlkes.Columns("SatuanJmlB")
        txtNamaBarang.Text = dgObatAlkes.Columns("NamaBarang")
    End If
    dgObatAlkes.Visible = False
    txtJumlah.Text = 1
    txtJumlah.SetFocus
    txtJumlah.SelStart = 0
    txtJumlah.SelLength = Len(txtJumlah)
    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dgObatAlkes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgObatAlkes_DblClick
End Sub

Private Sub dtpTglOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcStatusBarang.SetFocus
End Sub

Private Sub fgData_DblClick()
    Call fgData_KeyDown(13, 0)
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)

    Select Case KeyCode
        Case 13
            If fgData.TextMatrix(fgData.Row, 2) = "" Then Exit Sub
            Call SubLoadText
            txtIsi.Text = Trim(fgData.TextMatrix(fgData.Row, fgData.Col))
            txtIsi.SelStart = 0
            txtIsi.SelLength = Len(txtIsi.Text)

        Case vbKeyDelete
            If fgData.Row = fgData.Rows - 1 Then Exit Sub
            fgData.RemoveItem fgData.Row

    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 27 Then dgObatAlkes.Visible = False
End Sub

Private Sub Form_Load()
    On Error GoTo Errload
    Call PlayFlashMovie(Me)

    Call centerForm(Me, MDIUtama)
    dtpTglOrder.value = Now
    mstrKdRuangan = mstrKdRuanganLogin
    Call subSetGrid
    Call subLoadDcSource

    dgObatAlkes.Top = 2880
    dgObatAlkes.Left = 120
    dgObatAlkes.Visible = False

    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If KeyAscii = 13 Then
        Select Case fgData.Col
            Case 4
                If Val(txtIsi.Text) = 0 Then txtIsi.Text = 0
            Case 5
                If Val(txtIsi.Text) = 0 Then
                    txtIsi.Text = 0
                ElseIf Val(txtIsi.Text) > 99.99 Then
                    txtIsi.Text = 99.99
                End If
        End Select

        fgData.TextMatrix(fgData.Row, fgData.Col) = txtIsi.Text
        txtIsi.Visible = False

        If fgData.RowPos(fgData.Row) >= fgData.Height - 360 Then
            fgData.SetFocus
            SendKeys "{DOWN}"
            Exit Sub
        End If
        fgData.SetFocus
    ElseIf KeyAscii = 27 Then
        txtIsi.Visible = False
        fgData.SetFocus
    End If
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = Asc(".")) Then KeyAscii = 0
End Sub

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

Private Sub txtJumlah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdtambah.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = 44) Then KeyAscii = 0
End Sub

Private Sub txtJumlah_LostFocus()
    If Trim(txtJumlah.Text) = "," Then txtJumlah.Text = 0
End Sub

Private Sub txtNamaBarang_Change()
    On Error GoTo Errload
    Dim i As Integer
    If dcStatusBarang.BoundText = "02" Then
        strSQL = "EXECUTE CariBarang_V '" & txtNamaBarang.Text & "', '" & dcRuanganTujuan.BoundText & "'"
         
        Set dbRst = Nothing
        
        Call msubRecFO(dbRst, strSQL)
        If dbRst.EOF Then Exit Sub
        Set dgObatAlkes.DataSource = dbRst
        With dgObatAlkes
            For i = 0 To .Columns.Count - 1
                .Columns(i).Width = 0
            Next i
            .Columns("Nama Barang").Width = 2800
            .Columns("Jenis Barang").Width = 1100
            .Columns("Satuan").Width = 0
            .Columns("Kekuatan").Width = 0 '800
            .Columns("KdBarang").Width = 0
            .Columns("AsalBarang").Width = 1000
            .Columns("kdAsal").Width = 0
            .Columns("JmlStokOnHand").Width = 1500
            .Width = .Columns("Nama Barang").Width + .Columns("Jenis Barang").Width + .Columns("Kekuatan").Width + .Columns("AsalBarang").Width + .Columns("JmlStok").Width + 2000
        End With
    
        

    ElseIf dcStatusBarang.BoundText = "01" Then
        
        strSQL = "SELECT distinct JenisBarang, NamaBarang, SatuanJmlB, SUM(JmlStok) AS JmlStok, KdBarang," & _
        " KdSatuanJmlB, KdRuangan " & _
        " FROM V_AmbilStockBarangNonMedis " & _
        " WHERE NamaBarang like '%" & txtNamaBarang.Text & "%' AND KdRuangan = '" & dcRuanganTujuan.BoundText & "' " & _
        " GROUP BY JenisBarang, NamaBarang, SatuanJmlB, KdBarang, KdSatuanJmlB, KdRuangan" & _
        " ORDER BY NamaBarang "

        Call msubRecFO(dbRst, strSQL)

        Set dgObatAlkes.DataSource = dbRst
        With dgObatAlkes
            For i = 0 To .Columns.Count - 1
                .Columns(i).Width = 0
            Next i
            .Columns("JenisBarang").Width = 1500
            .Columns("NamaBarang").Width = 3000
            .Columns("SatuanJmlB").Width = 1500
            .Columns("KdBarang").Width = 1500
        End With

    End If
    dgObatAlkes.Visible = True

    Exit Sub
Errload:
    Call msubPesanError
    'Resume 0
End Sub

Private Sub txtNamaBarang_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            If dgObatAlkes.Visible = True Then
                dgObatAlkes.SetFocus
            Else
                txtJumlah.SetFocus
            End If
        Case 27
            dgObatAlkes.Visible = False
    End Select
End Sub

Private Sub subKosong()
    txtNoOrder.Text = ""
    dtpTglOrder.value = Now
    dcRuanganTujuan.BoundText = ""
    txtNamaBarang.Text = ""
    txtStock.Text = 0
    txtJumlah.Text = 0
    dgObatAlkes.Visible = False
End Sub

Private Sub subSetGrid()
    On Error GoTo Errload
    With fgData
        .Clear
        .Rows = 2
        .Cols = 8

        .RowHeight(0) = 500

        .TextMatrix(0, 0) = "Nama Barang"
        .TextMatrix(0, 1) = "Stock"
        .TextMatrix(0, 2) = "Jumlah"
        .TextMatrix(0, 3) = "Satuan"
        .TextMatrix(0, 4) = "KdBarang"
        .TextMatrix(0, 5) = "KdSatuan"
        .TextMatrix(0, 6) = "KdAsal"
        .TextMatrix(0, 7) = "NoTerima"
        
        .ColWidth(0) = 5000
        .ColWidth(1) = 1300
        .ColWidth(2) = 1000
        .ColWidth(3) = 1100
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignCenterCenter
    End With

    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub subLoadDcSource()
    On Error GoTo Errload

    Call msubDcSource(dcStatusBarang, rs, "SELECT KdKelompokBarang, KelompokBarang FROM KelompokBarang Where KdKelompokBarang='01' AND StatusEnabled=1 ORDER BY KelompokBarang")
    If rs.EOF = False Then dcStatusBarang.BoundText = rs(0).value
    Call msubDcSource(dcRuanganTujuan, rs, "SELECT KdRuangan, NamaRuangan FROM V_StrukOrderRuanganTujuan WHERE KdKelompokBarang='" & dcStatusBarang.BoundText & "' AND StatusEnabled=1 and KdRuangan<>'" & mstrKdRuangan & "' ORDER BY NamaRuangan")
    If rs.EOF = False Then dcRuanganTujuan.BoundText = rs(0).value

    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub SubLoadText()
    Dim i As Integer
    txtIsi.Left = fgData.Left
    Select Case fgData.Col
        Case 4, 5
            txtIsi.MaxLength = 5
        Case Else
            Exit Sub
    End Select

    For i = 0 To fgData.Col - 1
        txtIsi.Left = txtIsi.Left + fgData.ColWidth(i)
    Next i
    txtIsi.Visible = True
    txtIsi.Top = fgData.Top - 7

    For i = 0 To fgData.Row - 1
        txtIsi.Top = txtIsi.Top + fgData.RowHeight(i)
    Next i

    If fgData.TopRow > 1 Then
        txtIsi.Top = txtIsi.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If

    txtIsi.Width = fgData.ColWidth(fgData.Col)
    txtIsi.Height = fgData.RowHeight(fgData.Row)

    txtIsi.Visible = True
    txtIsi.SelStart = Len(txtIsi.Text)
    txtIsi.SetFocus
End Sub

Private Function sp_StrukOrder() As Boolean
    On Error GoTo Errload
    sp_StrukOrder = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder.Text)
        .Parameters.Append .CreateParameter("TglOrder", adDate, adParamInput, , Format(dtpTglOrder.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, dcRuanganTujuan.BoundText)
        .Parameters.Append .CreateParameter("KdSupplier", adChar, adParamInput, 4, Null)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputNoOrder", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_StrukOrder"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data struk order", vbCritical, "Validasi"
            sp_StrukOrder = False
        Else
            txtNoOrder.Text = .Parameters("OutputNoOrder").value
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Exit Function
Errload:
    Call msubPesanError(" sp_StrukOrder")
    sp_StrukOrder = False
End Function

Private Function sp_DetailOrderRuangan(f_KdBarang As String, f_JumlahBarang As Double, f_KdAsal As String, f_Status As String) As Boolean
    On Error GoTo Errload
    sp_DetailOrderRuangan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder.Text)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("JmlOrder", adDouble, adParamInput, , CDbl(f_JumlahBarang))
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("NoKonfirmasi", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("JmlBarangKonfirmasi", adInteger, adParamInput, , Null)
        .Parameters.Append .CreateParameter("NamaKonfirmasi", adVarChar, adParamInput, 150, Null)
        .Parameters.Append .CreateParameter("JmlAkomodir", adDouble, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DetailOrderRuangan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data detail pemesanan", vbCritical, "Validasi"
            sp_DetailOrderRuangan = False
        ElseIf .Parameters("JmlAkomodir").value < f_JumlahBarang Then
            MsgBox "Untuk " & fgData.TextMatrix(i, 0) & " hanya dapat diakomodir sejumlah " & .Parameters("JmlAkomodir").value, vbInformation, "Informasi"
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Exit Function
Errload:
    Call msubPesanError(" sp_DetailOrderRuangan")
End Function

Private Function sp_DetailOrderRuanganNonMedis(f_KdBarang As String, f_JumlahBarang As Double, f_Status As String) As Boolean
    On Error GoTo Errload
    sp_DetailOrderRuanganNonMedis = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder.Text)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("JmlOrder", adDouble, adParamInput, , CDbl(f_JumlahBarang))
        .Parameters.Append .CreateParameter("NoKonfirmasi", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("JmlBarangKonfirmasi", adDouble, adParamInput, , Null)
        .Parameters.Append .CreateParameter("NamaKonfirmasi", adVarChar, adParamInput, 150, Null)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DetailOrderRuanganNonMedis"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data detail pemesanan", vbCritical, "Validasi"
            sp_DetailOrderRuanganNonMedis = False
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Exit Function
Errload:
    Call msubPesanError(" sp_DetailOrderRuanganNonMedis")
End Function

Public Function subLoadDataOrder() As Boolean
    On Error GoTo Errload
    Dim i As Integer
    Dim j As Integer
    Dim tempDiscount As String

    txtNamaBarang.Text = "": txtSatuan.Text = ""
    dgObatAlkes.Visible = False
    Call subSetGrid

    strSQL = "SELECT * FROM V_StrukOrderCetakMedis WHERE NoOrder = '" & txtNoOrder.Text & "' AND KdRuangan = '" & mstrKdRuangan & "'"
    Call msubRecFO(rs, strSQL)

    If rs.EOF = True Then
        dtpTglOrder.value = Now
        dcRuanganTujuan.BoundText = ""
        subLoadDataOrder = False
        Exit Function
    End If

    subLoadDataOrder = True
    dtpTglOrder.value = rs("TglOrder").value
    dcRuanganTujuan.BoundText = rs("KdSupplier").value
    With fgData
        For i = 1 To rs.RecordCount
            .TextMatrix(i, 0) = rs("Nama Barang").value
            .TextMatrix(i, 1) = rs("AsalBarang").value
            .TextMatrix(i, 2) = rs("Satuan").value
            .TextMatrix(i, 3) = rs("JmlStok").value
            .TextMatrix(i, 4) = rs("JmlOrder").value
            .Col = 4: .Row = i: .CellForeColor = vbBlue: .CellFontBold = True

            For j = 1 To Len(rs("Discount").value)
                tempDiscount = Mid(rs("Discount").value, j, 1)
                If tempDiscount = "," Then tempDiscount = "."
                .TextMatrix(i, 5) = .TextMatrix(i, 5) & tempDiscount
            Next j

            .TextMatrix(i, 6) = rs("KdBarang").value
            .TextMatrix(i, 7) = rs("KdAsal").value
            .TextMatrix(i, 8) = rs("KdSatuanJmlB").value
            rs.MoveNext
            .Rows = .Rows + 1
        Next i
        .Row = 1
    End With

    Exit Function
Errload:
    Call msubPesanError
End Function


