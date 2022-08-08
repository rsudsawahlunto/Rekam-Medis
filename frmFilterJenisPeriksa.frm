VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmFilterJenisPeriksa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kunjungan Pasien"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFilterJenisPeriksa.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   8580
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
      TabIndex        =   14
      Top             =   960
      Width           =   8565
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   16
         Top             =   120
         Width           =   3645
         Begin VB.OptionButton optGroupBy 
            Caption         =   "Jenis Pasien"
            Height          =   210
            Index           =   4
            Left            =   1800
            TabIndex        =   1
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton optGroupBy 
            Caption         =   "Instalasi Asal"
            Height          =   210
            Index           =   3
            Left            =   240
            TabIndex        =   0
            Top             =   240
            Width           =   1815
         End
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   360
         Left            =   4440
         TabIndex        =   2
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   635
         _Version        =   393216
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   4440
         TabIndex        =   15
         Top             =   120
         Width           =   1755
      End
   End
   Begin VB.Frame fraPeriode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   0
      TabIndex        =   11
      Top             =   1800
      Width           =   8565
      Begin VB.Frame Frame1 
         Caption         =   "Group"
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton optGroupBy 
            Caption         =   "Total"
            Height          =   210
            Index           =   5
            Left            =   2520
            TabIndex        =   18
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optGroupBy 
            Caption         =   "Bulan"
            Height          =   210
            Index           =   1
            Left            =   810
            TabIndex        =   4
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optGroupBy 
            Caption         =   "Hari"
            Height          =   210
            Index           =   0
            Left            =   100
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optGroupBy 
            Caption         =   "Tahun"
            Height          =   210
            Index           =   2
            Left            =   1630
            TabIndex        =   5
            Top             =   360
            Width           =   855
         End
      End
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
         Left            =   3600
         TabIndex        =   12
         Top             =   240
         Width           =   4815
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   120
            TabIndex        =   6
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
            OLEDropMode     =   1
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   62521347
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   2640
            TabIndex        =   7
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
            Format          =   62521347
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2250
            TabIndex        =   13
            Top             =   315
            Width           =   255
         End
      End
   End
   Begin VB.Frame fraButton 
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
      Left            =   0
      TabIndex        =   10
      Top             =   3000
      Width           =   8565
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Spreadsheet"
         Height          =   375
         Left            =   4200
         TabIndex        =   8
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   6360
         TabIndex        =   9
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
      Left            =   6720
      Picture         =   "frmFilterJenisPeriksa.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmFilterJenisPeriksa.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmFilterJenisPeriksa.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmFilterJenisPeriksa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrInstalasi2 As String
Dim strcetak4 As String

Private Sub cmdCetak_Click()
On Error GoTo hell

Dim mdtBulan As Integer
Dim MdtTahun As Integer
If Periksa("datacombo", dcInstalasi, "Data instalasi kosong") = False Then Exit Sub
'-----------------------------------------------------------------------
    If (optGroupBy(0).Value = True And optGroupBy(3).Value = True) Or (optGroupBy(5).Value = True And optGroupBy(3).Value = True) Then
        mdTglAwal = dtpAwal.Value
        mdTglAkhir = dtpAkhir.Value
        mstrKdInstalasi = dcInstalasi.BoundText
        mstrInstalasi2 = dcInstalasi.Text
        Select Case strCetak
             Case "LapKunjunganJenisPeriksa"
                  strcetak4 = IIf(optGroupBy(5).Value = True, "Total", strCetak)
                  strCetak2 = "LapKunjunganJenisPeriksahari"
                  strCetak3 = "JenisPeriksaBInstalasiAsal"
                  strSQL = "Select * from V_DataKunjunganPasienBJenisPeriksa " & _
                          "WHERE (KdInstalasi= '" & mstrKdInstalasi & "' And TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPelayanan asc"
   
             Case "lapkunjunganjenistindakan"
                 strcetak4 = IIf(optGroupBy(5).Value = True, "Total", strCetak)
                 strCetak2 = "LapKunjunganJenisTindakanHari"
                 strCetak3 = "LapKunjunganJenisTindakanBinstalasiAsal"
                 strSQL = "Select * from V_DataKunjunganPasienBJenisPelayanan " & _
                          "WHERE (KdInstalasi= '" & mstrKdInstalasi & "' And TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPelayanan asc"

        End Select
'-------------------------------------------------------------------------
    ElseIf (optGroupBy(0).Value = True And optGroupBy(4).Value = True) Or (optGroupBy(5).Value = True And optGroupBy(4).Value = True) Then
        mdTglAwal = dtpAwal.Value
        mdTglAkhir = dtpAkhir.Value
        mstrKdInstalasi = dcInstalasi.BoundText
        mstrInstalasi2 = dcInstalasi.Text
            Select Case strCetak
            Case "LapKunjunganJenisPeriksa"
                  strcetak4 = IIf(optGroupBy(5).Value = True, "Total", strCetak)
                  strCetak2 = "LapKunjunganJenisPeriksahari"
                  strCetak3 = "JenisPeriksaBJenispasien"
                  strSQL = "Select * from V_DataKunjunganPasienBJenisPeriksa " & _
                          "WHERE (KdInstalasi= '" & mstrKdInstalasi & "' and  TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPelayanan asc"
   
            Case "lapkunjunganjenistindakan"
                  strcetak4 = IIf(optGroupBy(5).Value = True, "Total", strCetak)
                  strCetak2 = "LapKunjunganJenisTindakanHari"
                  strCetak3 = "LapKunjunganJenisTindakanBJenisPasienHari"
                  strSQL = "Select * from V_DataKunjunganPasienBJenisPelayanan " & _
                          "WHERE (KdInstalasi= '" & mstrKdInstalasi & "' And TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPelayanan asc"

   
                        
            End Select
'-------------------------------------------------------------------------
    ElseIf optGroupBy(1).Value = True And optGroupBy(4).Value = True Then
        mdTglAwal = CDate(Format(dtpAwal.Value, "yyyy-mm ") & "-01 00:00:00") 'TglAwal
        mdtBulan = CStr(Format(dtpAkhir.Value, "mm"))
        MdtTahun = CStr(Format(dtpAkhir.Value, "yyyy"))
        mdTglAkhir = CDate(Format(dtpAkhir.Value, "yyyy-mm") & "-" & funcHitungHari(mdtBulan, MdtTahun) & " 23:59:59")
        
        mstrKdInstalasi = dcInstalasi.BoundText
        mstrInstalasi2 = dcInstalasi.Text
        Select Case strCetak
        Case "LapKunjunganJenisPeriksa"
              strCetak2 = "LapKunjunganJenisPeriksabulan"
              strCetak3 = "LapKunjunganJenisPeriksaBulanJenisPasien"
              strSQL = "SELECT dbo.FB_TakeBlnThn(TglPelayanan)  AS TglPelayanan, RuanganPelayanan, JenisPeriksa, Jenispasien, JK, JmlPelayanan  FROM   V_DataKunjunganPasienBJenisPeriksa " _
                        & "WHERE (TglPelayanan BETWEEN '" _
                        & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
                        & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " & _
                        " And KdInstalasi= '" & mstrKdInstalasi & "'"
                        
        Case "lapkunjunganjenistindakan"
              strCetak2 = "LapKunjunganJenisTindakanBulan"
              strCetak3 = "LapKunjunganJenisTindakanBJenisPasienBulan"
              strSQL = "SELECT dbo.FB_TakeBlnThn(TglPelayanan)  AS TglPelayanan, RuanganPelayanan, JenisPelayanan, JenisPasien, JK, JmlPelayanan  FROM   V_DataKunjunganPasienBJenisPelayanan " _
                        & "WHERE (TglPelayanan BETWEEN '" _
                        & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "' AND '" _
                        & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " & _
                        "And KdInstalasi= '" & mstrKdInstalasi & "'"
        End Select
'-------------------------------------------------------------------------
    ElseIf optGroupBy(1).Value = True And optGroupBy(3).Value = True Then
        mdTglAwal = CDate(Format(dtpAwal.Value, "yyyy-mm ") & "-01 00:00:00") 'TglAwal
        mdtBulan = CStr(Format(dtpAkhir.Value, "mm"))
        MdtTahun = CStr(Format(dtpAkhir.Value, "yyyy"))
        mdTglAkhir = CDate(Format(dtpAkhir.Value, "yyyy-mm") & "-" & funcHitungHari(mdtBulan, MdtTahun) & " 23:59:59")
        mstrKdInstalasi = dcInstalasi.BoundText
        mstrInstalasi2 = dcInstalasi.Text
        Select Case strCetak
         Case "LapKunjunganJenisPeriksa"
              strCetak2 = "LapKunjunganJenisPeriksabulan"
              strCetak3 = "LapKunjunganJenisPeriksaBulanInstalasiAsal"
              strSQL = "SELECT dbo.FB_TakeBlnThn(TglPelayanan) AS TglPelayanan, RuanganPelayanan, JenisPeriksa, InstalasiAsal, JK, JmlPelayanan  FROM   V_DataKunjunganPasienBJenisPeriksa " _
                        & "WHERE (TglPelayanan BETWEEN '" _
                        & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
                        & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " & _
                        " And KdInstalasi= '" & mstrKdInstalasi & "'"
                        
         Case "lapkunjunganjenistindakan"
              strCetak2 = "LapKunjunganJenisTindakanBulan"
              strCetak3 = "LapKunjunganJenisTindakanBinstalasiAsalBulan"
              strSQL = "SELECT dbo.FB_TakeBlnThn(TglPelayanan) AS TglPelayanan, RuanganPelayanan, JenisPelayanan, InstalasiAsal, JK, JmlPelayanan  FROM   V_DataKunjunganPasienBJenisPelayanan " _
                        & "WHERE (TglPelayanan BETWEEN '" _
                        & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
                        & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " & _
                        " And KdInstalasi= '" & mstrKdInstalasi & "'"
         End Select
         
'-------------------------------------------------------------------------
        
    ElseIf optGroupBy(2).Value = True And optGroupBy(3).Value = True Then
        mdTglAwal = CDate("01-01-" & Format(dtpAwal.Value, "yyyy HH:MM:SS")) 'TglAwal
        mdTglAkhir = CDate("31-12-" & Format(dtpAkhir.Value, "yyyy HH:MM:SS")) 'TglAkhir
        mstrKdInstalasi = dcInstalasi.BoundText
        mstrInstalasi2 = dcInstalasi.Text
         Select Case strCetak
         Case "LapKunjunganJenisPeriksa"
              strCetak2 = "LapKunjunganJenisPeriksaTahun"
              strCetak3 = "LapKunjunganJenisPeriksaInstalasiAsalTahun"
              strSQL = "Select * from V_DataKunjunganPasienBJenisPeriksa " & _
                          "WHERE (TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                          " And KdInstalasi= '" & mstrKdInstalasi & "'"
                          
         Case "lapkunjunganjenistindakan"
              strCetak2 = "LapKunjunganJenisTindakantahun"
              strCetak3 = "LapKunjunganJenisTindakanBinstalasiAsaltahun"
              strSQL = "Select * from V_DataKunjunganPasienBJenisPelayanan " & _
                          "WHERE (TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                          " And KdInstalasi= '" & mstrKdInstalasi & "'"
        End Select
'-------------------------------------------------------------------------
        ElseIf optGroupBy(2).Value = True And optGroupBy(4).Value = True Then
        mdTglAwal = CDate("01-01-" & Format(dtpAwal.Value, "yyyy HH:MM:SS")) 'TglAwal
        mdTglAkhir = CDate("31-12-" & Format(dtpAkhir.Value, "yyyy HH:MM:SS")) 'TglAkhir
        mstrKdInstalasi = dcInstalasi.BoundText
        mstrInstalasi2 = dcInstalasi.Text
         Select Case strCetak
         Case "LapKunjunganJenisPeriksa"
              strCetak2 = "LapKunjunganJenisPeriksaTahun"
              strCetak3 = "LapKunjunganJenisPeriksaJenisPasienTahun"
              strSQL = "Select * from V_DataKunjunganPasienBJenisPeriksa " & _
                          "WHERE (TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                          " And KdInstalasi= '" & mstrKdInstalasi & "'"
                          
         Case "lapkunjunganjenistindakan"
              strCetak2 = "LapKunjunganJenisTindakantahun"
              strCetak3 = "LapKunjunganJenisTindakanBJenisPtahun"
              strSQL = "Select * from V_DataKunjunganPasienBJenisPelayanan " & _
                          "WHERE (TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                          " And KdInstalasi= '" & mstrKdInstalasi & "'"
        End Select
'--------------------------------------------------------------------------
   

        End If
    If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then MsgBox "Tidak Ada Data", vbExclamation, "Validasi": Exit Sub
    frmcetakviewer.Show
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then optGroupBy(0).SetFocus
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    With Me
        .dtpAwal.Value = Now
        .dtpAkhir.Value = Now
    End With
    Call subDcSource
    Call cekOpt
End Sub

Private Sub cekOpt()
   If optGroupBy(0).Value = True Then
      Call optGroupBy_Click(0)
   ElseIf optGroupBy(1).Value = True Then
      Call optGroupBy_Click(1)
   ElseIf optGroupBy(2).Value = True Then
     Call optGroupBy_Click(2)
   End If
End Sub

Private Sub optGroupBy_Click(Index As Integer)
  Select Case Index
   Case 0
        dtpAwal.CustomFormat = "dd MMMM yyyyy"
        dtpAkhir.CustomFormat = "dd MMMM yyyyy"
        optGroupBy(1).Value = False
        optGroupBy(2).Value = False
        
   Case 1
        dtpAkhir.CustomFormat = "MMMM yyyyy"
        dtpAwal.CustomFormat = "MMMM yyyyy"
        optGroupBy(0).Value = False
        optGroupBy(2).Value = False
        
   Case 2
        dtpAkhir.CustomFormat = "yyyyy"
        dtpAwal.CustomFormat = "yyyyy"
        optGroupBy(0).Value = False
        optGroupBy(1).Value = False
    Case 5
        dtpAwal.CustomFormat = "dd MMMM yyyyy"
        dtpAkhir.CustomFormat = "dd MMMM yyyyy"
        optGroupBy(1).Value = False
        optGroupBy(2).Value = False
   End Select
End Sub

Private Sub optGroupBy_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub subDcSource()
On Error GoTo hell
  ' If strCetak = "LapKunjunganJenisStatus" Or strCetak = "LapKunjunganSt_PnyktPsn" Or strCetak = "LapKunjunganRujukanBStatus" Then
        Select Case strCetak
        Case "lapkunjunganjenistindakan"
        strSQL = "SELECT KdInstalasi, NamaInstalasi " & _
            " From instalasi" & _
            " WHERE (KdInstalasi IN ('01', '02', '03', '04', '06', '08', '09', '10','13', '16'))"
        Case "LapKunjunganJenisPeriksa"
        strSQL = "SELECT KdInstalasi, NamaInstalasi " & _
            " From instalasi WHERE (KdInstalasi IN ('09', '10'))"
        End Select
        Call msubDcSource(dcInstalasi, rs, strSQL)
  '  Else
        
   ' End If
Exit Sub
hell:
    Call msubPesanError
End Sub

