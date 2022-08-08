VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9b.ocx"
Begin VB.Form FrmPeriodeLaporanKematian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPeriodeLaporanKematian.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   9930
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
      TabIndex        =   20
      Top             =   7560
      Width           =   9855
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   7980
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Spredsheet"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4320
         TabIndex        =   22
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdgrafik 
         Caption         =   "&Grafik"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6150
         TabIndex        =   21
         Top             =   240
         Width           =   1665
      End
   End
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
      TabIndex        =   0
      Top             =   930
      Width           =   9855
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
         TabIndex        =   12
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
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPickerAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   14
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
            Format          =   58785795
            UpDown          =   -1  'True
            CurrentDate     =   37956
         End
         Begin MSComCtl2.DTPicker DTPickerAkhir 
            Height          =   375
            Left            =   3360
            TabIndex        =   15
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
            Format          =   58785795
            UpDown          =   -1  'True
            CurrentDate     =   37956
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3000
            TabIndex        =   16
            Top             =   315
            Width           =   255
         End
      End
      Begin VB.CheckBox chkGroup 
         Caption         =   "Jenis Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2415
      End
      Begin VB.CheckBox chkInstalasi 
         Caption         =   "Instalasi Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   2175
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
         TabIndex        =   7
         Top             =   1200
         Width           =   5595
         Begin VB.CheckBox chRuangPoli 
            Caption         =   "Ruang / Poli"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1575
         End
         Begin MSDataListLib.DataCombo dcRuangPoli 
            Height          =   360
            Left            =   1800
            TabIndex        =   9
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
      Begin VB.Frame Frame4 
         Caption         =   "Kriteria / Urutan"
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
         TabIndex        =   1
         Top             =   2040
         Width           =   9680
         Begin VB.OptionButton opt_jmlPasien 
            Caption         =   "Jumlah Pasien"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2040
            TabIndex        =   5
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton opt_pnama 
            Caption         =   "Jumlah Biaya"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3840
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.OptionButton optKodeDiagnosa 
            Caption         =   "Diagnosa"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtJmlData 
            Height          =   375
            Left            =   8760
            TabIndex        =   2
            Text            =   "0"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Jumlah Data"
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
            Left            =   7440
            TabIndex        =   6
            Top             =   360
            Width           =   1065
         End
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   360
         Left            =   120
         TabIndex        =   17
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
      Begin MSDataListLib.DataCombo dcJenisPasien 
         Height          =   360
         Left            =   120
         TabIndex        =   18
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
      Begin MSDataGridLib.DataGrid fgData 
         Height          =   3375
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   24
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
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "FrmPeriodeLaporanKematian.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "FrmPeriodeLaporanKematian.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8040
      Picture         =   "FrmPeriodeLaporanKematian.frx":4CE9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "FrmPeriodeLaporanKematian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iRowNow As Integer
Dim rstopten As New ADODB.recordset
Dim iRowNow2 As Integer

Private Sub chkGroup_Click()
    If chkGroup.Value = vbChecked Then
        dcJenisPasien.Enabled = True
        Call msubDcSource(dcJenisPasien, rs, "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien order by JenisPasien")
        dcJenisPasien.Text = rs(1).Value
    Else
        dcJenisPasien.Enabled = False
        dcJenisPasien.Text = ""
    End If
End Sub

Private Sub chkInstalasi_Click()
    If chkInstalasi.Value = vbChecked Then
        dcInstalasi.Enabled = True
        Call msubDcSource(dcInstalasi, rs, "SELECT KdInstalasi, NamaInstalasi FROM Instalasi WHERE (KdInstalasi IN ('01', '02', '03', '06', '08'))")
        dcInstalasi.Text = rs(1).Value
    Else
        dcInstalasi.Enabled = False
        dcInstalasi.Text = ""
    End If
End Sub

Private Sub chRuangPoli_Click()
    If chRuangPoli.Value = vbChecked Then
        dcRuangPoli.Enabled = True
        Call msubDcSource(dcRuangPoli, rs, "SELECT KdRuangan, NamaRuangan FROM Ruangan WHERE (KdInstalasi = '" & dcInstalasi.BoundText & "')")
        dcRuangPoli.Text = rs(1).Value
    Else
        dcRuangPoli.Enabled = False
        dcRuangPoli.Text = ""
    End If
End Sub

Private Sub Cek(SekalianCetak As Boolean, Biasa As Boolean)
    If Val(txtJmlData) = 0 Then
        MsgBox "Jumlah data harus diisi.", vbOKOnly + vbExclamation, "Informasi"
        txtJmlData.SetFocus
        cmdCetak.Enabled = False
        cmdgrafik.Enabled = False
        Exit Sub
    End If
    
    If chkInstalasi.Value = vbChecked Then
        mstrFilter = " = " + "'" + dcInstalasi.BoundText + "'"
    Else
        mstrFilter = "IN ('01', '02', '03', '06', '08')"
    End If
    
    
    If optKodeDiagnosa.Value = True Then
        strSQL = "SELECT top " & Val(txtJmlData) & " Diagnosa, sum(jumlahpasien) as [JmlPasien]" & _
            " FROM V_RekapitulasiDiagnosaKematian" & _
            " WHERE TglPeriksa BETWEEN " & _
            " '" & Format(DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
            " '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & _
            " AND kdinstalasi " & mstrFilter & " AND KdRuangan LIKE '%" & dcRuangPoli.BoundText & "%' AND " & _
            " JenisPasien LIKE '%" & dcJenisPasien & "%' group by Diagnosa order by Diagnosa asc"
    Else
        strSQL = "SELECT top " & Val(txtJmlData) & " Diagnosa, sum(jumlahpasien) as [JmlPasien]" & _
            " FROM V_RekapitulasiDiagnosaKematian " & _
            " WHERE TglPeriksa BETWEEN " & _
            " '" & Format(DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
            " '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & _
            " AND kdinstalasi " & mstrFilter & " AND KdRuangan LIKE '%" & dcRuangPoli.BoundText & "%' AND " & _
            " JenisPasien LIKE '%" & dcJenisPasien & "%' group by Diagnosa order by [JmlPasien] desc"
    End If
    
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    
    Set fgData.DataSource = rs
    subSetGrid
    
    'jika tidak ada data cari
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = False
        cmdgrafik.Enabled = False
        If dcInstalasi.Enabled = True Then dcInstalasi.SetFocus
        Exit Sub
    Else
        If SekalianCetak = True Then
            If Biasa = True Then
                cetak = "RekapTopten"
                FrmViewerLaporan10.Show
                FrmViewerLaporan10.Caption = "Medifirst2000 - Rekapitulasi Kematian"
            ElseIf Biasa = False Then
                cetak = "RekapToptenGrafik"
                FrmViewerLaporan10.Show
                FrmViewerLaporan10.Caption = "Medifirst2000 - Grafik Rekapitulasi Kematian"
            End If
        End If
        cmdCetak.Enabled = True
        cmdgrafik.Enabled = True
    End If
End Sub

Private Sub cmdCari_Click()
On Error GoTo errLoad

    Cek False, False

    Exit Sub
errLoad:
    Call msubPesanError
    fgData.Visible = True
End Sub

Private Sub cmdCetak_Click()
    cmdCetak.Enabled = False
    
    Cek True, True
        
End Sub

Private Sub cmdgrafik_Click()
    cmdgrafik.Enabled = False
    
    Cek True, False

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcInstalasi_Change()
    chRuangPoli.Value = vbUnchecked
    If chkInstalasi.Value = vbChecked Then
        Call msubDcSource(dcRuangPoli, rs, "SELECT KdRuangan, NamaRuangan FROM Ruangan WHERE (KdInstalasi = '" & dcInstalasi.BoundText & "')")
        If rs.RecordCount > 0 Then
            frInstalasi.Enabled = True
            'dcRuangPoli.Text = rs(1).Value
        Else
            frInstalasi.Enabled = False
        End If
    Else
        frInstalasi.Enabled = False
    End If
    frInstalasi.Caption = dcInstalasi
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then DTPickerAwal.SetFocus
End Sub

Private Sub DTPickerAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdcari.SetFocus
End Sub

Private Sub DTPickerAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then DTPickerAkhir.SetFocus
End Sub

'Private Sub fgData_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
'    MsgBox fgData.Columns(0).Width
'End Sub

Private Sub Form_Load()
    
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    
    With Me
        .DTPickerAwal.Value = Now
        .DTPickerAkhir.Value = Now
    End With
    
    optKodeDiagnosa.Value = True

    Call subSetGrid
End Sub

Private Sub subSetGrid()
    With fgData
        .Columns(0).Width = 7980
        .Columns(0).Alignment = dbgCenter
    End With
End Sub

Private Sub txtJmlData_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13
        If Val(txtJmlData) > 0 Then
            cmdcari.SetFocus
        Else
            MsgBox "Jumlah Data harus diisi.", vbOKOnly + vbExclamation, "Informasi"
        End If
    Case Else
        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
            Beep
            KeyAscii = 0
        End If
    End Select
End Sub


