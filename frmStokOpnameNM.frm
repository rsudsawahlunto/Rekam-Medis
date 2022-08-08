VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStokOpnameNM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Stok Opname Barang Non Medis"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStokOpnameNM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   15045
   Begin VB.CheckBox chkSetStokReal 
      Caption         =   "Set Stok Real = 0"
      Height          =   375
      Left            =   11280
      TabIndex        =   21
      Top             =   7200
      Width           =   1815
   End
   Begin VB.TextBox txtCariNama 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1200
      TabIndex        =   8
      Top             =   7200
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox txtIsi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   7680
      Visible         =   0   'False
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtNoClosing 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   960
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtKeterangan 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6240
      TabIndex        =   9
      Top             =   7200
      Width           =   4935
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   11280
      TabIndex        =   10
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   13095
      TabIndex        =   11
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   14775
      Begin VB.Frame Frame1 
         Caption         =   "Group by"
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
         Left            =   3720
         TabIndex        =   14
         Top             =   120
         Width           =   10935
         Begin VB.TextBox txtKriteriaJenis 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   7800
            TabIndex        =   5
            Top             =   285
            Width           =   1815
         End
         Begin VB.OptionButton optLokasi 
            Caption         =   "Lokasi Barang"
            Height          =   495
            Left            =   -240
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.OptionButton optAsal 
            Caption         =   "Asal Barang"
            Height          =   495
            Left            =   3360
            TabIndex        =   2
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optJenis 
            Caption         =   "Jenis Barang"
            Height          =   495
            Left            =   1800
            TabIndex        =   1
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   390
            Left            =   9840
            TabIndex        =   6
            Top             =   290
            Width           =   855
         End
         Begin MSDataListLib.DataCombo dcCariData 
            Height          =   390
            Left            =   4800
            TabIndex        =   4
            Top             =   285
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   688
            _Version        =   393216
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cari Barang"
            Height          =   210
            Index           =   2
            Left            =   6840
            TabIndex        =   23
            Top             =   360
            Width           =   900
         End
      End
      Begin MSComCtl2.DTPicker dtpTglPenutupan 
         Height          =   450
         Left            =   960
         TabIndex        =   0
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   794
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   419758083
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Penutupan"
         Height          =   210
         Index           =   1
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   1275
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   4935
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   8705
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   8250
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   13229
            Text            =   "F1 : Cetak"
            TextSave        =   "F1 : Cetak"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   26485
            Text            =   "Ctrl + C : Copy Stok System To Stok Real"
            TextSave        =   "Ctrl + C : Copy Stok System To Stok Real"
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
      TabIndex        =   22
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
      Left            =   13200
      Picture         =   "frmStokOpnameNM.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmStokOpnameNM.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmStokOpnameNM.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
   Begin VB.Label lblJmlData 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data 0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   13920
      TabIndex        =   20
      Top             =   7200
      Width           =   915
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      Height          =   210
      Index           =   0
      Left            =   5160
      TabIndex        =   16
      Top             =   7320
      Width           =   945
   End
End
Attribute VB_Name = "frmStokOpnameNM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrJmlStokReal() As Long
Dim subCopy As Boolean
Dim mebIsi As Object
Dim i As Integer

Private Sub subLoadText()
    txtIsi.Left = fgData.Left
    Select Case fgData.Col
        Case 4
            txtIsi.MaxLength = 6
        Case Else
            Exit Sub
    End Select

    If fgData.Col = 6 Then
        With mebIsi
            For i = 0 To fgData.Col - 1
                .Left = .Left + fgData.ColWidth(i)
            Next i
            .Visible = True
            .Top = fgData.Top - 7

            If fgData.TopRow > 1 Then
                .Top = .Top + fgData.RowHeight(0)
                For i = fgData.TopRow To fgData.Row - 1
                    .Top = .Top + fgData.RowHeight(i)
                Next i
            Else
                For i = 0 To fgData.Row - 1
                    .Top = .Top + fgData.RowHeight(i)
                Next i
            End If

            .Width = fgData.ColWidth(fgData.Col)
            .Height = fgData.RowHeight(fgData.Row)

            .Visible = True
            .SelStart = Len(mebIsi.Text)
            .SetFocus
            .Text = IIf(fgData.TextMatrix(fgData.Row, fgData.Col) = "0", "__/__/____", fgData.TextMatrix(fgData.Row, fgData.Col))
            .SelStart = 0
            .SelLength = Len(mebIsi.Text)
        End With
    Else
        With txtIsi
            For i = 0 To fgData.Col - 1
                .Left = .Left + fgData.ColWidth(i)
            Next i
            .Visible = True
            .Top = fgData.Top - 7

            If fgData.TopRow > 1 Then
                .Top = .Top + fgData.RowHeight(0)
                For i = fgData.TopRow To fgData.Row - 1
                    .Top = .Top + fgData.RowHeight(i)
                Next i
            Else
                For i = 0 To fgData.Row - 1
                    .Top = .Top + fgData.RowHeight(i)
                Next i
            End If

            .Width = fgData.ColWidth(fgData.Col)
            .Height = fgData.RowHeight(fgData.Row)

            .Visible = True
            .SelStart = Len(txtIsi.Text)
            .SetFocus
            .Text = Trim(fgData.TextMatrix(fgData.Row, fgData.Col))
            .SelStart = 0
            .SelLength = Len(txtIsi.Text)
        End With
    End If
End Sub

Private Sub chkSetStokReal_Click()
    On Error GoTo Errload

    If chkSetStokReal.Value = vbChecked Then
        MousePointer = vbHourglass
        ReDim Preserve arrJmlStokReal(fgData.Rows - 1)

        For i = 1 To fgData.Rows - 1
            With fgData
                .TextMatrix(i, 4) = 0
            End With
        Next i
        MousePointer = vbDefault
    Else
        MousePointer = vbHourglass
        For i = 1 To fgData.Rows - 1
            With fgData
                .TextMatrix(i, 4) = .TextMatrix(i, 3)  'arrJmlStokReal(i)
            End With
        Next i
        MousePointer = vbDefault
    End If

    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub cmdCari_Click()
    Dim str As String
    On Error GoTo errTampilkan

    MousePointer = vbHourglass
    fgData.Visible = False

    If optJenis.Value = True Then
        mstrFilter = " AND (KdDetailJenisBarang = '" & dcCariData.BoundText & "')"
    ElseIf optAsal.Value = True Then
        mstrFilter = " AND (KdAsal = '" & dcCariData.BoundText & "')"
    ElseIf optLokasi.Value = True Then
        mstrFilter = " AND (Lokasi = '" & dcCariData.BoundText & "')"
    End If

    If dcCariData.BoundText = "" Then mstrFilter = ""

    strsql = "SELECT * FROM V_DataStokBarangNonMedisNonRekap WHERE (KdRuangan = '" & mstrKdRuangan & "') " & mstrFilter & " AND NamaBarang Like '" & txtKriteriaJenis.Text & "%' ORDER BY JenisBarang, NamaBarang"
    '((NoRegisterAsset <> '0000000' AND NoRegisterAsset <> '0' AND NoRegisterAsset <>'000000') OR KdJenisAset is null) AND
    
'    strsql = "SELECT * FROM V_DataStokBarangNonMedisRekap WHERE (KdRuangan = '" & mstrKdRuangan & "') " & mstrFilter & " AND NamaBarang Like '" & txtKriteriaJenis.Text & "%' ORDER BY JenisBarang, NamaBarang"
    Call msubRecFO(rs, strsql)

    Call subSetGrid
    If IsNull(rs(0)) Then Exit Sub
    For i = 1 To rs.RecordCount
        fgData.Rows = fgData.Rows + 1
        fgData.TextMatrix(i, 0) = IIf(IsNull(rs("JenisBarang").Value), "", rs("JenisBarang").Value)
        fgData.TextMatrix(i, 1) = IIf(IsNull(rs("NamaBarang").Value), "", rs("NamaBarang").Value)
        fgData.TextMatrix(i, 2) = IIf(IsNull(rs("AsalBarang").Value), "", rs("AsalBarang").Value)
        fgData.TextMatrix(i, 3) = IIf(IsNull(rs("StokSystem").Value), "", Format(rs("StokSystem").Value, "#,##0"))
        fgData.TextMatrix(i, 4) = IIf(IsNull(rs("StokSystem").Value), "", Format(rs("StokSystem").Value, "#,##0"))
        fgData.TextMatrix(i, 5) = IIf(IsNull(rs("HargaNetto").Value), "", Format(rs("HargaNetto").Value, "##,###,##0"))
        fgData.TextMatrix(i, 6) = IIf(IsNull(rs("Discount").Value), "", Format(rs("Discount").Value, "##,###,##0"))
        fgData.TextMatrix(i, 7) = IIf(IsNull(rs("Ruangan").Value), "", rs("Ruangan").Value)
        fgData.TextMatrix(i, 8) = IIf(IsNull(rs("KdBarang").Value), "", rs("KdBarang").Value)
        fgData.TextMatrix(i, 9) = IIf(IsNull(rs("KdAsal").Value), "", rs("KdAsal").Value)
        fgData.TextMatrix(i, 10) = IIf(IsNull(rs("KdRuangan").Value), "", rs("KdRuangan").Value)
        fgData.TextMatrix(i, 11) = IIf(IsNull(rs("NoRegisterAsset").Value), "", rs("NoRegisterAsset").Value)
'        fgData.TextMatrix(i, 12) = IIf(IsNull(rs("TglInputStock").Value), "", rs("TglInputStock").Value) ' add colum TglinputStock
        rs.MoveNext
    Next i
    MousePointer = vbDefault
    fgData.Visible = True
    If fgData.Rows < 1 Then dcCariData.SetFocus Else fgData.SetFocus: fgData.Col = 5
    lblJmlData.Caption = "Data 0 / " & fgData.Rows - 2
    Exit Sub
errTampilkan:
    MousePointer = vbDefault
    fgData.Visible = True
    Call msubPesanError
'    Resume 0
End Sub

Private Sub cmdSimpan_Click()
    Dim i As Integer
    Dim mstrValue As String

    If fgData.TextMatrix(1, 8) = "" Then Exit Sub
    If MsgBox("Apakah Anda yakin akan mengupdate Stok Barang Non Medis ?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub

    cmdSimpan.Enabled = False
    cmdTutup.Enabled = False
    ProgressBar.Visible = True
    ProgressBar.Min = 0
    ProgressBar.Max = fgData.Rows - 2
    ProgressBar.Value = 0

    If sp_Closing = False Then Exit Sub
    For i = 1 To fgData.Rows - 2
        mstrValue = i
        ProgressBar.Value = i
        With fgData
            If sp_DataStokBarangNonMedis(.TextMatrix(i, 8), .TextMatrix(i, 9), .TextMatrix(i, 3), .TextMatrix(i, 4), IIf(Len(.TextMatrix(i, 5)) = 0, 0, .TextMatrix(i, 5)), IIf(Len(.TextMatrix(i, 6)) = 0, 0, .TextMatrix(i, 6)), .TextMatrix(i, 11), Format(Now, "yyyy/MM/dd HH:mm:ss")) = False Then Exit Sub
            
        End With
    Next i
    MsgBox "Stok Opname sukses", vbInformation, "Informasi"

    Call Add_HistoryLoginActivity("Add_Closing+AU_DataStokBarangNonMedis")
    ProgressBar.Visible = False
    cmdSimpan.Enabled = False
    cmdTutup.Enabled = True

    Exit Sub
Errload:
    ProgressBar.Visible = False
    cmdSimpan.Enabled = True
    cmdTutup.Enabled = True
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcCariData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdCari_Click
        If fgData.Rows < 1 Then dcCariData.SetFocus Else fgData.SetFocus
    End If
End Sub

Private Sub dtpTglPenutupan_Change()
    txtKeterangan.Text = "STOK OPNAME BULAN " & UCase(MonthName(dtpTglPenutupan.Month))
End Sub

Private Sub dtpTglPenutupan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then optJenis.SetFocus
End Sub

Private Sub fgData_DblClick()
    If fgData.Row = 0 Or fgData.Row = fgData.Rows - 1 Then Exit Sub
    If fgData.TextMatrix(fgData.Row, 1) = "" Then Exit Sub
    Call fgData_KeyDown(13, 0)
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)

    Select Case KeyCode
        Case 13
            If fgData.Row = 0 Or fgData.Row = fgData.Rows - 1 Then Exit Sub
            If fgData.TextMatrix(fgData.Row, 1) = "" Then Exit Sub
            Call subLoadText
            txtIsi.Text = Trim(fgData.TextMatrix(fgData.Row, fgData.Col))
            txtIsi.SelStart = 0
            txtIsi.SelLength = Len(txtIsi.Text)

        Case vbKeyC
            If strCtrlKey = 4 Then
                If fgData.Row = 0 Then Exit Sub
                For i = 1 To fgData.Rows - 2
                    fgData.TextMatrix(i, 4) = fgData.TextMatrix(i, 3)
                Next i
            End If
    End Select
End Sub

Private Sub fgData_RowColChange()
    On Error Resume Next
    lblJmlData.Caption = "Data " & fgData.Row & " / " & fgData.Rows - 2
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglPenutupan.Value = Now
    txtKeterangan.Text = "STOK OPNAME BULAN " & UCase(MonthName(dtpTglPenutupan.Month))
    optJenis.Value = True
    Call subSetGrid
    subCopy = False
End Sub

Private Sub optAsal_Click()
    dcCariData.BoundText = ""
    Call msubDcSource(dcCariData, rs, "SELECT KdAsal, NamaAsal FROM AsalBarang where KdInstalasi='05' and StatusEnabled='1' ORDER BY NamaAsal")
End Sub

Private Sub optAsal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcCariData.SetFocus
End Sub

Private Sub optJenis_Click()
    dcCariData.BoundText = ""
    Call msubDcSource(dcCariData, rs, "SELECT KdDetailJenisBarang, DetailJenisBarang FROM V_DetailJenisBrgPerKelompokBrg WHERE KdKelompokBarang = '" & mstrKdKelompokBarang & "' and StatusEnabled='1' Order By DetailJenisBarang")
End Sub

Private Sub optJenis_GotFocus()
    Call optJenis_Click
End Sub

Private Sub optJenis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcCariData.SetFocus
End Sub

Private Sub optLokasi_Click()
    dcCariData.BoundText = ""
    Call msubDcSource(dcCariData, rs, "SELECT DISTINCT Lokasi FROM StokRuangan WHERE KdRuangan = '" & mstrKdRuangan & "' ORDER BY Lokasi")
End Sub

Private Sub optLokasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcCariData.SetFocus
End Sub

Private Sub txtCariNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With fgData
            .Row = 1
            .Col = 0

            For i = 1 To .Rows - 2
                If UCase(Left(txtCariNama.Text, Len(txtCariNama.Text))) = UCase(Left(fgData.TextMatrix(i, 1), Len(txtCariNama.Text))) Then Exit For
            Next i
            .TopRow = i: .Row = i: .Col = 5: .SetFocus
        End With
    End If
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If KeyAscii = 13 Then
        If Val(txtIsi.Text) = 0 Then txtIsi.Text = 0

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
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Function sp_Closing() As Boolean
    On Error GoTo Errload

    sp_Closing = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoClosing", adChar, adParamInput, 10, IIf(Len(Trim(txtNoClosing.Text)) = 0, Null, Trim(txtNoClosing.Text)))
        .Parameters.Append .CreateParameter("TglClosing", adDate, adParamInput, , Format(dtpTglPenutupan.Value, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(Len(Trim(txtKeterangan.Text)) = 0, Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputNoClosing", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "Add_Closing"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data closing", vbCritical, "Validasi"
            sp_Closing = False
        Else
            txtNoClosing.Text = .Parameters("OutputNoClosing").Value
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
Errload:
    sp_Closing = False
    Call msubPesanError
    cmdSimpan.Enabled = True
    cmdTutup.Enabled = True
End Function

Private Function sp_DataStokBarangNonMedis(f_KdBarang As String, f_KdAsal As String, _
    f_JmlStokSystem As Double, f_JmlStokReal As Double, f_HargaNetto As Double, f_Discount As Double, f_NoRegisterAsset As String, f_tglInputStok As Date) As Boolean
    On Error GoTo Errload

    sp_DataStokBarangNonMedis = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoClosing", adChar, adParamInput, 10, Trim(txtNoClosing.Text))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("JmlStokSystem", adDouble, adParamInput, , CDbl(f_JmlStokSystem))
        .Parameters.Append .CreateParameter("JmlStokReal", adDouble, adParamInput, , CDbl(f_JmlStokReal))
        .Parameters.Append .CreateParameter("HargaNetto", adCurrency, adParamInput, , f_HargaNetto)
        .Parameters.Append .CreateParameter("Discount", adCurrency, adParamInput, , f_Discount)
        .Parameters.Append .CreateParameter("NoRegisterAsset", adVarChar, adParamInput, 15, f_NoRegisterAsset)
'        .Parameters.Append .CreateParameter("TglInputStock", adDate, adParamInput, , IIf(f_tglInputStok = "__/__/____ __:__", Null, Format(f_tglInputStok, "yyyy/MM/dd HH:mm")))
        .Parameters.Append .CreateParameter("TglInputStock", adDate, adParamInput, , f_tglInputStok)
        .ActiveConnection = dbConn
        .CommandText = "AU_DataStokBarangNonMedisNew"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data Stok Opname Barang Non Medis.", vbCritical, "Validasi"
            sp_DataStokBarangNonMedis = False
            cmdSimpan.Enabled = True
            cmdTutup.Enabled = True
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
Errload:
    Call msubPesanError
    cmdSimpan.Enabled = True
    cmdTutup.Enabled = True
    Resume 0
End Function

Private Sub subSetGrid()
    With fgData
        .Clear
        .Cols = 13
        .Rows = 2

        .RowHeight(0) = 400

        .MergeCells = flexMergeFree
'        .MergeCol(0) = True
'        .MergeCol(2) = True

'        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter

        .ColWidth(0) = 1300
        .ColWidth(1) = 3000
        .ColWidth(2) = 1500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1200
        .ColWidth(6) = 1000
        .ColWidth(7) = 1400
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 1750

        .TextMatrix(0, 0) = "JenisBarang"
        .TextMatrix(0, 1) = "NamaBarang"
        .TextMatrix(0, 2) = "AsalBarang"
        .TextMatrix(0, 3) = "StokSystem"
        .TextMatrix(0, 4) = "StokReal"
        .TextMatrix(0, 5) = "HargaNetto"
        .TextMatrix(0, 6) = "Discount"
        .TextMatrix(0, 7) = "Ruangan"
        .TextMatrix(0, 8) = "KdBarang"
        .TextMatrix(0, 9) = "KdAsal"
        .TextMatrix(0, 10) = "KdRuangan"
        .TextMatrix(0, 11) = "NoRegisterAsset"
        .TextMatrix(0, 12) = "TglInputStock"
    End With
End Sub

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKriteriaJenis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdCari_Click
End Sub



