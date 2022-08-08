VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmPeriodeLaporanTopTen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPeriodeLaporanTopTen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   9990
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6675
      Left            =   0
      TabIndex        =   16
      Top             =   930
      Width           =   9975
      Begin VB.Frame Frame4 
         Caption         =   "Kriteria"
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
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   9680
         Begin VB.OptionButton opt_pnama 
            Caption         =   "Pembayaran"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1920
            TabIndex        =   10
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton opt_jmlPasien 
            Caption         =   "Jumlah Pasien"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame frInstalasi 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4200
         TabIndex        =   20
         Top             =   1200
         Width           =   5595
         Begin VB.CheckBox chRuangPoli 
            Caption         =   "Ruang / Poli"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   1575
         End
         Begin MSDataListLib.DataCombo dcRuangPoli 
            Height          =   360
            Left            =   1800
            TabIndex        =   8
            Top             =   315
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            Style           =   2
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.CheckBox chkInstalasi 
         Caption         =   "Instalasi Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox chkGroup 
         Caption         =   "Jenis Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2415
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
         Left            =   4200
         TabIndex        =   17
         Top             =   315
         Width           =   5595
         Begin VB.CommandButton cmdcari 
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
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPickerAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
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
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   127401987
            UpDown          =   -1  'True
            CurrentDate     =   37956
         End
         Begin MSComCtl2.DTPicker DTPickerAkhir 
            Height          =   375
            Left            =   3360
            TabIndex        =   1
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
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
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   127401987
            UpDown          =   -1  'True
            CurrentDate     =   37956
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3000
            TabIndex        =   18
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Top             =   1515
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid fgdata 
         Height          =   3375
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   9630
         _ExtentX        =   16986
         _ExtentY        =   5953
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSDataListLib.DataCombo dcJenisPasien 
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Top             =   555
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
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
      Height          =   855
      Left            =   0
      TabIndex        =   15
      Top             =   7680
      Width           =   9975
      Begin VB.CommandButton cmdgrafik 
         Caption         =   "&Grafik"
         Height          =   495
         Left            =   6270
         TabIndex        =   13
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Spredsheet"
         Height          =   495
         Left            =   4440
         TabIndex        =   12
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   8100
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   19
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
      Left            =   8160
      Picture         =   "FrmPeriodeLaporanTopTen.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "FrmPeriodeLaporanTopTen.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "FrmPeriodeLaporanTopTen.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "FrmPeriodeLaporanTopTen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iRowNow As Integer
Dim rstopten As New ADODB.recordset
Dim iRowNow2 As Integer

Private Sub chkGroup_Click()
    If chkGroup.value = vbChecked Then
        dcJenisPasien.Enabled = True
        Call msubDcSource(dcJenisPasien, rs, "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien")
        dcJenisPasien.Text = rs(1).value
    Else
        dcJenisPasien.Enabled = False
        dcJenisPasien.Text = ""
    End If
End Sub

Private Sub chkGroup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkGroup.value = vbChecked Then
            dcJenisPasien.Enabled = True
            Call msubDcSource(dcJenisPasien, rs, "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien")
            dcJenisPasien.Text = rs(1).value
            dcJenisPasien.SetFocus
        Else
            dcJenisPasien.Enabled = False
            dcJenisPasien.Text = ""
            chkInstalasi.SetFocus
        End If
    End If
End Sub

Private Sub chkInstalasi_Click()
    If chkInstalasi.value = vbChecked Then
        dcInstalasi.Enabled = True
        Call msubDcSource(dcInstalasi, rs, "SELECT KdInstalasi, NamaInstalasi FROM Instalasi WHERE (KdInstalasi IN ('01', '02', '03', '06', '08'))")
        dcInstalasi.Text = rs(1).value
    Else
        dcInstalasi.Enabled = False
        dcInstalasi.Text = ""
    End If
End Sub

Private Sub chkInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkInstalasi.value = vbChecked Then
            dcInstalasi.Enabled = True
            Call msubDcSource(dcInstalasi, rs, "SELECT KdInstalasi, NamaInstalasi FROM Instalasi WHERE (KdInstalasi IN ('01', '02', '03', '06', '08'))")
            dcInstalasi.Text = rs(1).value
            dcInstalasi.SetFocus
        Else
            dcInstalasi.Enabled = False
            dcInstalasi.Text = ""
            DTPickerAwal.SetFocus
        End If

    End If
End Sub

Private Sub chRuangPoli_Click()
    If chRuangPoli.value = vbChecked Then
        dcRuangPoli.Enabled = True
        Call msubDcSource(dcRuangPoli, rs, "SELECT KdRuangan, NamaRuangan FROM Ruangan WHERE (KdInstalasi = '" & dcInstalasi.BoundText & "')")
        dcRuangPoli.Text = rs(1).value
    Else
        dcRuangPoli.Enabled = False
        dcRuangPoli.Text = ""
    End If
End Sub

Private Sub chRuangPoli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chRuangPoli.value = vbChecked Then
            dcRuangPoli.Enabled = True
            Call msubDcSource(dcRuangPoli, rs, "SELECT KdRuangan, NamaRuangan FROM Ruangan WHERE (KdInstalasi = '" & dcInstalasi.BoundText & "')")
            dcRuangPoli.Text = rs(1).value
            dcRuangPoli.SetFocus
        Else
            dcRuangPoli.Enabled = False
            dcRuangPoli.Text = ""
            opt_jmlPasien.SetFocus
        End If
    End If
End Sub

Private Sub cmdCari_Click()
    On Error GoTo errLoad

    Dim intJmlRow As Integer
    Dim intJmlPria As Integer
    Dim intJmlWanita As Integer
    Dim intJmlTotal As Integer

    Call subSetGrid
    'u/ mempercepat
    fgData.Visible = False: MousePointer = vbHourglass

    If chkInstalasi.value = vbChecked Then
        mstrFilter = " = " + "'" + dcInstalasi.BoundText + "'"
    Else
        mstrFilter = "IN ('01', '02', '03', '06', '08')"
    End If

    'Hitung jumlah row dari data yang hendak ditampilkan
    strSQL = "SELECT COUNT(TglPeriksa) AS JmlRow " & _
    " FROM V_RekapitulasiDiagnosaTopTen " & _
    " WHERE TglPeriksa BETWEEN " & _
    " '" & Format(DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND " & _
    " '" & Format(DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' " & _
    " AND kdinstalasi " & mstrFilter & " AND KdRuangan LIKE '%" & dcRuangPoli.BoundText & "%' AND " & _
    " JenisPasien LIKE '%" & dcJenisPasien & "%'"

    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    'jika tidak ada data
    If rs(0).value = 0 Then
        fgData.Visible = True: MousePointer = vbNormal
        txtJmlPria = "0": txtJmlTotal = "0": txtJmlWanita = "0"
        If dcInstalasi.Enabled = True Then dcInstalasi.SetFocus
        Exit Sub
    End If

    intJmlRow = rs("JmlRow").value

    'tampilan grid
    strSQL = "SELECT Instalasi,COUNT(instalasi) AS Jmldiagnosa, " & _
    " SUM(jumlahpasien)as jumlah From V_RekapitulasiDiagnosaTopTen " & _
    " WHERE TglPeriksa BETWEEN " & _
    " '" & Format(DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND " & _
    " '" & Format(DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "' " & _
    " AND kdinstalasi " & mstrFilter & " AND KdRuangan LIKE '%" & dcRuangPoli.BoundText & "%' AND JenisPasien LIKE '%" & dcJenisPasien & "%'" & _
    " GROUP BY instalasi"

    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    'Tambahkan jumlah row dengan jumlah subtotal
    intJmlRow = intJmlRow + rs.RecordCount

    'u/ menampilkan yang di group by
    Dim rstopten As New ADODB.recordset
    With fgData
        'jml baris akhir
        .Rows = intJmlRow + 2
        While rs.EOF = False
            strSQL = "SELECT instalasi,diagnosa,sum(jumlahpasien) as TjumlahPasien " & _
            " From V_RekapitulasiDiagnosaTopTen " & _
            " WHERE (TglPeriksa BETWEEN " & _
            " '" & Format(DTPickerAwal.value, "yyyy/MM/dd 00:00:00") & "' AND " & _
            " '" & Format(DTPickerAkhir.value, "yyyy/MM/dd 23:59:59") & "') " & _
            " AND instalasi ='" & rs("instalasi").value & "' " & _
            " AND kdinstalasi " & mstrFilter & " AND KdRuangan LIKE '%" & dcRuangPoli.BoundText & "%' AND JenisPasien LIKE '%" & dcJenisPasien & "%'" & _
            " GROUP BY instalasi, diagnosa" & _
            " ORDER BY diagnosa"

            Set rstopten = Nothing
            rstopten.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            While rstopten.EOF = False
                'baris u/ sub total
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = rstopten("Instalasi").value
                .TextMatrix(iRowNow, 2) = rstopten("Diagnosa").value
                .TextMatrix(iRowNow, 3) = rstopten("TJumlahPasien").value
                rstopten.MoveNext
            Wend

            iRowNow = iRowNow + 1
            'isi sub total
            .TextMatrix(iRowNow, 1) = .TextMatrix(iRowNow - 1, 1)
            .TextMatrix(iRowNow, 2) = "Sub Total"
            .TextMatrix(iRowNow, 3) = IIf(rs("jumlah").value = 0, 0, Format(rs("jumlah").value, "#,###"))

            Call subSetSubTotalRow(Me, iRowNow, 2, vbBlackness, vbWhite)

            'disimpan u/ jml total
            intJmlTotal = intJmlTotal + rs("jumlah").value

            rs.MoveNext
        Wend
        'banyak baris berdasarkan irownow
        .Rows = iRowNow + 2

        iRowNow = iRowNow + 1
        .TextMatrix(iRowNow, 1) = "Total"
        .TextMatrix(iRowNow, 3) = IIf(intJmlTotal = 0, 0, Format(intJmlTotal, "#,###"))

        Call subSetSubTotalRow(Me, iRowNow, 1, vbBlue, vbWhite)

        .Col = 1
        For i = 1 To .Rows - 1
            .Row = i
            .CellFontBold = True
        Next

        .Visible = True: MousePointer = vbNormal
    End With

    Exit Sub
errLoad:
    Call msubPesanError
    fgData.Visible = True
End Sub

Private Sub cmdCetak_Click()
    cmdCetak.Enabled = False
    strSQL = "SELECT * FROM V_RekapitulasiDiagnosaTopTen " _
    & "WHERE (TglPeriksa BETWEEN '" _
    & Format(DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
    & Format(DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "') and kdinstalasi " & mstrFilter & " AND KdRuangan LIKE '%" & dcRuangPoli.BoundText & "%' AND JenisPasien LIKE '%" & dcJenisPasien & "%' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    cetak = "RekapTopten"
    FrmViewerLaporan.Show
    FrmViewerLaporan.Caption = "Medifirst2000 - Rekapitulasi 10 Besar Penyakit"
    cmdCetak.Enabled = True

End Sub

Private Sub cmdgrafik_Click()
    cmdCetak.Enabled = False

    strSQL = "SELECT * FROM V_RekapitulasiDiagnosaTopTen " _
    & "WHERE (TglPeriksa BETWEEN '" _
    & Format(DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
    & Format(DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "') and kdinstalasi " & mstrFilter & " AND KdRuangan LIKE '%" & dcRuangPoli.BoundText & "%' AND JenisPasien LIKE '%" & dcJenisPasien & "%' ORDER BY instalasi, diagnosa"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If

    cetak = "RekapToptenGrafik"
    FrmViewerLaporan.Show
    FrmViewerLaporan.Caption = "Medifirst2000 - Grafik Rekapitulasi 10 Besar Penyakit"
    cmdCetak.Enabled = True
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcInstalasix()
    chRuangPoli.value = vbUnchecked
    If chkInstalasi.value = vbChecked Then
        Call msubDcSource(dcRuangPoli, rs, "SELECT KdRuangan, NamaRuangan FROM Ruangan WHERE (KdInstalasi = '" & dcInstalasi.BoundText & "')")
        If rs.RecordCount > 0 Then
            frInstalasi.Enabled = True
        Else
            frInstalasi.Enabled = False
        End If
    Else
        frInstalasi.Enabled = False
    End If
    frInstalasi.Caption = dcInstalasi
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcInstalasi.MatchedWithList = True Then DTPickerAwal.SetFocus
        strSQL = "SELECT KdInstalasi, NamaInstalasi FROM Instalasi WHERE (KdInstalasi IN ('01', '02', '03', '06', '08')) and (NamaInstalasi LIKE '%" & dcInstalasi.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcInstalasi.Text = ""
            Exit Sub
        End If
        dcInstalasi.BoundText = rs(0).value
        dcInstalasi.Text = rs(1).value
        Call dcInstalasix
    End If
End Sub

Private Sub dcJenisPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcJenisPasien.MatchedWithList = True Then chkInstalasi.SetFocus
        strSQL = "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien WHERE (JenisPasien LIKE '%" & dcJenisPasien.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcJenisPasien.Text = ""
            Exit Sub
        End If
        dcJenisPasien.BoundText = rs(0).value
        dcJenisPasien.Text = rs(1).value
    End If
End Sub

Private Sub dcRuangPoli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcRuangPoli.MatchedWithList = True Then opt_jmlPasien.SetFocus
        strSQL = "SELECT KdRuangan, NamaRuangan FROM Ruangan WHERE (KdInstalasi = '" & dcInstalasi.BoundText & "') and (NamaRuangan LIKE '%" & dcRuangPoli.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcRuangPoli.Text = ""
            Exit Sub
        End If
        dcRuangPoli.BoundText = rs(0).value
        dcRuangPoli.Text = rs(1).value
    End If
End Sub

Private Sub DTPickerAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdcari.SetFocus
End Sub

Private Sub DTPickerAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DTPickerAkhir.SetFocus
End Sub

Private Sub Form_Load()

    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    txtJmlPria = "0": txtJmlTotal = "0": txtJmlWanita = "0"

    With Me
        .DTPickerAwal.value = Now
        .DTPickerAkhir.value = Now
    End With

    Call subSetGrid
End Sub

Private Sub subSetGrid()
    With fgData
        .Visible = False
        .clear
        .Cols = 4
        .Rows = 2
        .Row = 0

        For i = 1 To .Cols - 1
            .Col = i
            .CellFontBold = True
            .RowHeight(0) = 300
            .CellAlignment = flexAlignCenterCenter
        Next

        .MergeCells = 1
        .MergeCol(1) = True

        .TextMatrix(0, 1) = "Nama Sub Instalasi"
        .TextMatrix(0, 2) = "Diagnosa Penyakit"
        .TextMatrix(0, 3) = "Jumlah Pasien"

        .ColWidth(0) = 500
        .ColWidth(1) = 2850
        .ColWidth(2) = 4000
        .ColWidth(3) = 2000

        .Visible = True
        iRowNow = 0
    End With
End Sub

