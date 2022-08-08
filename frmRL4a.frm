VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmRL4a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL4 A Jumlah Tenaga Kesehatan Menurut Jenis"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   Icon            =   "frmRL4a.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7560
   Begin VB.OptionButton Option4 
      Caption         =   "Hal. 4"
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Hal. 3"
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Hal. 2"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Hal. 1"
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   1560
      Value           =   -1  'True
      Width           =   1455
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
      TabIndex        =   0
      Top             =   2400
      Width           =   7605
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   1905
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   4200
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
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
   Begin VB.Frame Frame2 
      Caption         =   "RL 4 A Halaman"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   7575
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   4680
      Picture         =   "frmRL4a.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2955
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRL4a.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL4a.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmRL4a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project/reference/microsoft excel 12.0 object library
'Selalu gunakan format file excel 2003  .xls sebagai standar agar pengguna excel 2003 atau diatasnya dpt menggunakan report laporannya
'Catatan: Format excel 2000 tidak dpt mengoperasikan beberapa fungsi yg ada pada excell 2003 atau diatasnya

Option Explicit

'Special Buat Excel
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel.Range
Dim oResizeRange As Excel.Range
Dim j As String
'Special Buat Excel

Private Sub cmdCetak_Click()
    On Error GoTo hell

    If Option1.value = True Then

        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL4 Hal1.xls")
        Set oSheet = oWB.ActiveSheet

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("g5")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("y7", "y8")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing

        strSQL = "SELECT  KdKualifikasiJurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak," & _
        "jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart," & _
        "jmlhonorer " & _
        "FROM RL4 WHERE  (KdkualifikasiJurusan IN ('0034', '0035','0036','0037','0038','0039','0040','0041','0042','0043','0045','0046','0049','0053','0054','0055','0056','0058','0060','0061','0064','0065','0067','0069','0071','0074','0076','0078','0081','0083')) " & _
        "GROUP BY  KdKualifikasiJurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak,jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart,jmlhonorer" & _
        " order by KdKualifikasiJurusan"

        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 14

            While Not rs.EOF
                With oSheet
                    .Cells(j, 7) = Trim(IIf(IsNull(rs!jmldpkfull.value), 0, (rs!jmldpkfull.value)))
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!jmldpbfull.value), 0, (rs!jmldpbfull.value)))
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!jmldaerahfull.value), 0, (rs!jmldaerahfull.value)))
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpnkfull.value), 0, (rs!jmlpnkfull.value)))
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlabrifull.value), 0, (rs!jmlabrifull.value)))
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmldeplainfull.value), 0, (rs!jmldeplainfull.value)))
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!jmlpttfull.value), 0, (rs!jmlpttfull.value)))
                    .Cells(j, 14) = Trim(IIf(IsNull(rs!jmlswastafull.value), 0, (rs!jmlswastafull.value)))
                    .Cells(j, 15) = Trim(IIf(IsNull(rs!jmlkontrak.value), 0, (rs!jmlkontrak.value)))
                    .Cells(j, 17) = Trim(IIf(IsNull(rs!jmldpkpart.value), 0, (rs!jmldpkpart.value)))
                    .Cells(j, 18) = Trim(IIf(IsNull(rs!jmldpbpart.value), 0, (rs!jmldpbpart.value)))
                    .Cells(j, 19) = Trim(IIf(IsNull(rs!jmldaerahpart.value), 0, (rs!jmldaerahpart.value)))
                    .Cells(j, 20) = Trim(IIf(IsNull(rs!jmlpnkpart.value), 0, (rs!jmlpnkpart.value)))
                    .Cells(j, 21) = Trim(IIf(IsNull(rs!jmlabripart.value), 0, (rs!jmlabripart.value)))
                    .Cells(j, 22) = Trim(IIf(IsNull(rs!jmldeplainpart.value), 0, (rs!jmldeplainpart.value)))
                    .Cells(j, 23) = Trim(IIf(IsNull(rs!jmlpttpart.value), 0, (rs!jmlpttpart.value)))
                    .Cells(j, 24) = Trim(IIf(IsNull(rs!jmlswastapart.value), 0, (rs!jmlswastapart.value)))
                    .Cells(j, 26) = Trim(IIf(IsNull(rs!jmlhonorer.value), 0, (rs!jmlhonorer.value)))
                End With
                j = j + 1
                rs.MoveNext
            Wend
        End If
        j = j + 1
        Set rsb = Nothing
        strSQL = "SELECT  KdKualifikasiJurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak," & _
        "jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart," & _
        "jmlhonorer " & _
        "FROM RL4 WHERE  (KdkualifikasiJurusan IN ('0086','0090','0092')) " & _
        "GROUP BY  KdKualifikasiJurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak,jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart,jmlhonorer" & _
        " order by KdKualifikasiJurusan"
        Call msubRecFO(rsb, strSQL)

        If rsb.RecordCount > 0 Then
            rsb.MoveFirst
            j = 45

            While Not rsb.EOF
                With oSheet
                    .Cells(j, 7) = Trim(IIf(IsNull(rsb!jmldpkfull.value), 0, (rsb!jmldpkfull.value)))
                    .Cells(j, 8) = Trim(IIf(IsNull(rsb!jmldpbfull.value), 0, (rsb!jmldpbfull.value)))
                    .Cells(j, 9) = Trim(IIf(IsNull(rsb!jmldaerahfull.value), 0, (rsb!jmldaerahfull.value)))
                    .Cells(j, 10) = Trim(IIf(IsNull(rsb!jmlpnkfull.value), 0, (rsb!jmlpnkfull.value)))
                    .Cells(j, 11) = Trim(IIf(IsNull(rsb!jmlabrifull.value), 0, (rsb!jmlabrifull.value)))
                    .Cells(j, 12) = Trim(IIf(IsNull(rsb!jmldeplainfull.value), 0, (rsb!jmldeplainfull.value)))
                    .Cells(j, 13) = Trim(IIf(IsNull(rsb!jmlpttfull.value), 0, (rsb!jmlpttfull.value)))
                    .Cells(j, 14) = Trim(IIf(IsNull(rsb!jmlswastafull.value), 0, (rsb!jmlswastafull.value)))
                    .Cells(j, 15) = Trim(IIf(IsNull(rsb!jmlkontrak.value), 0, (rsb!jmlkontrak.value)))
                    .Cells(j, 17) = Trim(IIf(IsNull(rsb!jmldpkpart.value), 0, (rsb!jmldpkpart.value)))
                    .Cells(j, 18) = Trim(IIf(IsNull(rsb!jmldpbpart.value), 0, (rsb!jmldpbpart.value)))
                    .Cells(j, 19) = Trim(IIf(IsNull(rsb!jmldaerahpart.value), 0, (rsb!jmldaerahpart.value)))
                    .Cells(j, 20) = Trim(IIf(IsNull(rsb!jmlpnkpart.value), 0, (rsb!jmlpnkpart.value)))
                    .Cells(j, 21) = Trim(IIf(IsNull(rsb!jmlabripart.value), 0, (rsb!jmlabripart.value)))
                    .Cells(j, 22) = Trim(IIf(IsNull(rsb!jmldeplainpart.value), 0, (rsb!jmldeplainpart.value)))
                    .Cells(j, 23) = Trim(IIf(IsNull(rsb!jmlpttpart.value), 0, (rsb!jmlpttpart.value)))
                    .Cells(j, 24) = Trim(IIf(IsNull(rsb!jmlswastapart.value), 0, (rsb!jmlswastapart.value)))
                    .Cells(j, 26) = Trim(IIf(IsNull(rsb!jmlhonorer.value), 0, (rsb!jmlhonorer.value)))
                End With
                j = j + 1
                rsb.MoveNext
            Wend
        End If

    ElseIf Option2.value = True Then

        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL4 Hal2.xls")
        Set oSheet = oWB.ActiveSheet

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("g5")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("y7", "y8")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing

        strSQL = "SELECT  KdKualifikasiJurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak," & _
        "jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart," & _
        "jmlhonorer " & _
        "FROM RL4 WHERE  (KdkualifikasiJurusan IN ('0005', '0006','0007','0144','0145','0146','0147','0148','0149','0150','0174','0175','0176','0177','0178'))  " & _
        "GROUP BY  KdKualifikasiJurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak,jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart,jmlhonorer" & _
        " order by KdKualifikasiJurusan"

        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 14

            While Not rs.EOF
                With oSheet
                    .Cells(j, 7) = Trim(IIf(IsNull(rs!jmldpkfull.value), 0, (rs!jmldpkfull.value)))
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!jmldpbfull.value), 0, (rs!jmldpbfull.value)))
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!jmldaerahfull.value), 0, (rs!jmldaerahfull.value)))
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpnkfull.value), 0, (rs!jmlpnkfull.value)))
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlabrifull.value), 0, (rs!jmlabrifull.value)))
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmldeplainfull.value), 0, (rs!jmldeplainfull.value)))
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!jmlpttfull.value), 0, (rs!jmlpttfull.value)))
                    .Cells(j, 14) = Trim(IIf(IsNull(rs!jmlswastafull.value), 0, (rs!jmlswastafull.value)))
                    .Cells(j, 15) = Trim(IIf(IsNull(rs!jmlkontrak.value), 0, (rs!jmlkontrak.value)))
                    .Cells(j, 17) = Trim(IIf(IsNull(rs!jmldpkpart.value), 0, (rs!jmldpkpart.value)))
                    .Cells(j, 18) = Trim(IIf(IsNull(rs!jmldpbpart.value), 0, (rs!jmldpbpart.value)))
                    .Cells(j, 19) = Trim(IIf(IsNull(rs!jmldaerahpart.value), 0, (rs!jmldaerahpart.value)))
                    .Cells(j, 20) = Trim(IIf(IsNull(rs!jmlpnkpart.value), 0, (rs!jmlpnkpart.value)))
                    .Cells(j, 21) = Trim(IIf(IsNull(rs!jmlabripart.value), 0, (rs!jmlabripart.value)))
                    .Cells(j, 22) = Trim(IIf(IsNull(rs!jmldeplainpart.value), 0, (rs!jmldeplainpart.value)))
                    .Cells(j, 23) = Trim(IIf(IsNull(rs!jmlpttpart.value), 0, (rs!jmlpttpart.value)))
                    .Cells(j, 24) = Trim(IIf(IsNull(rs!jmlswastapart.value), 0, (rs!jmlswastapart.value)))
                    .Cells(j, 26) = Trim(IIf(IsNull(rs!jmlhonorer.value), 0, (rs!jmlhonorer.value)))
                End With
                j = j + 1
                rs.MoveNext
            Wend
        End If

        j = j + 4
        Set rsb = Nothing
        strSQL = "SELECT  KdKualifikasiJurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak," & _
        "jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart," & _
        "jmlhonorer " & _
        "FROM RL4 WHERE  (KdkualifikasiJurusan IN ('0008', '0009','0153','0154','0155','0156','0157','0158','0159')) " & _
        "GROUP BY  KdKualifikasiJurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak,jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart,jmlhonorer" & _
        " order by KdKualifikasiJurusan"
        Call msubRecFO(rsb, strSQL)

        If rsb.RecordCount > 0 Then
            rsb.MoveFirst
            j = 33

            While Not rsb.EOF
                With oSheet
                    .Cells(j, 7) = Trim(IIf(IsNull(rsb!jmldpkfull.value), 0, (rsb!jmldpkfull.value)))
                    .Cells(j, 8) = Trim(IIf(IsNull(rsb!jmldpbfull.value), 0, (rsb!jmldpbfull.value)))
                    .Cells(j, 9) = Trim(IIf(IsNull(rsb!jmldaerahfull.value), 0, (rsb!jmldaerahfull.value)))
                    .Cells(j, 10) = Trim(IIf(IsNull(rsb!jmlpnkfull.value), 0, (rsb!jmlpnkfull.value)))
                    .Cells(j, 11) = Trim(IIf(IsNull(rsb!jmlabrifull.value), 0, (rsb!jmlabrifull.value)))
                    .Cells(j, 12) = Trim(IIf(IsNull(rsb!jmldeplainfull.value), 0, (rsb!jmldeplainfull.value)))
                    .Cells(j, 13) = Trim(IIf(IsNull(rsb!jmlpttfull.value), 0, (rsb!jmlpttfull.value)))
                    .Cells(j, 14) = Trim(IIf(IsNull(rsb!jmlswastafull.value), 0, (rsb!jmlswastafull.value)))
                    .Cells(j, 15) = Trim(IIf(IsNull(rsb!jmlkontrak.value), 0, (rsb!jmlkontrak.value)))
                    .Cells(j, 17) = Trim(IIf(IsNull(rsb!jmldpkpart.value), 0, (rsb!jmldpkpart.value)))
                    .Cells(j, 18) = Trim(IIf(IsNull(rsb!jmldpbpart.value), 0, (rsb!jmldpbpart.value)))
                    .Cells(j, 19) = Trim(IIf(IsNull(rsb!jmldaerahpart.value), 0, (rsb!jmldaerahpart.value)))
                    .Cells(j, 20) = Trim(IIf(IsNull(rsb!jmlpnkpart.value), 0, (rsb!jmlpnkpart.value)))
                    .Cells(j, 21) = Trim(IIf(IsNull(rsb!jmlabripart.value), 0, (rsb!jmlabripart.value)))
                    .Cells(j, 22) = Trim(IIf(IsNull(rsb!jmldeplainpart.value), 0, (rsb!jmldeplainpart.value)))
                    .Cells(j, 23) = Trim(IIf(IsNull(rsb!jmlpttpart.value), 0, (rsb!jmlpttpart.value)))
                    .Cells(j, 24) = Trim(IIf(IsNull(rsb!jmlswastapart.value), 0, (rsb!jmlswastapart.value)))
                    .Cells(j, 26) = Trim(IIf(IsNull(rsb!jmlhonorer.value), 0, (rsb!jmlhonorer.value)))
                End With
                j = j + 1
                rsb.MoveNext
            Wend
        End If

    ElseIf Option3.value = True Then

        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL4 Hal3.xls")
        Set oSheet = oWB.ActiveSheet

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("g5")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("y7", "y8")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing

        strSQL = "SELECT  KdKualifikasiJurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak," & _
        "jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart," & _
        "jmlhonorer " & _
        "FROM RL4 WHERE  (KdkualifikasiJurusan IN ('0016', '0017','0018','0019','0020','0021','0022','0023','0180')) " & _
        "GROUP BY  KdKualifikasiJurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak,jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart,jmlhonorer" & _
        " order by KdKualifikasiJurusan"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 14

            While Not rs.EOF
                With oSheet
                    .Cells(j, 7) = Trim(IIf(IsNull(rs!jmldpkfull.value), 0, (rs!jmldpkfull.value)))
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!jmldpbfull.value), 0, (rs!jmldpbfull.value)))
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!jmldaerahfull.value), 0, (rs!jmldaerahfull.value)))
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpnkfull.value), 0, (rs!jmlpnkfull.value)))
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlabrifull.value), 0, (rs!jmlabrifull.value)))
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmldeplainfull.value), 0, (rs!jmldeplainfull.value)))
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!jmlpttfull.value), 0, (rs!jmlpttfull.value)))
                    .Cells(j, 14) = Trim(IIf(IsNull(rs!jmlswastafull.value), 0, (rs!jmlswastafull.value)))
                    .Cells(j, 15) = Trim(IIf(IsNull(rs!jmlkontrak.value), 0, (rs!jmlkontrak.value)))
                    .Cells(j, 17) = Trim(IIf(IsNull(rs!jmldpkpart.value), 0, (rs!jmldpkpart.value)))
                    .Cells(j, 18) = Trim(IIf(IsNull(rs!jmldpbpart.value), 0, (rs!jmldpbpart.value)))
                    .Cells(j, 19) = Trim(IIf(IsNull(rs!jmldaerahpart.value), 0, (rs!jmldaerahpart.value)))
                    .Cells(j, 20) = Trim(IIf(IsNull(rs!jmlpnkpart.value), 0, (rs!jmlpnkpart.value)))
                    .Cells(j, 21) = Trim(IIf(IsNull(rs!jmlabripart.value), 0, (rs!jmlabripart.value)))
                    .Cells(j, 22) = Trim(IIf(IsNull(rs!jmldeplainpart.value), 0, (rs!jmldeplainpart.value)))
                    .Cells(j, 23) = Trim(IIf(IsNull(rs!jmlpttpart.value), 0, (rs!jmlpttpart.value)))
                    .Cells(j, 24) = Trim(IIf(IsNull(rs!jmlswastapart.value), 0, (rs!jmlswastapart.value)))
                    .Cells(j, 26) = Trim(IIf(IsNull(rs!jmlhonorer.value), 0, (rs!jmlhonorer.value)))
                End With
                j = j + 1
                rs.MoveNext
            Wend
        End If

        j = j + 4
        Set rsb = Nothing
        strSQL = "SELECT  KdKualifikasiJurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak," & _
        "jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart," & _
        "jmlhonorer " & _
        "FROM RL4 WHERE  (KdkualifikasiJurusan IN ('0024', '0025','0026','0027','0029','0030','0181')) " & _
        "GROUP BY  KdKualifikasiJurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak,jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart,jmlhonorer" & _
        " order by KdKualifikasiJurusan"
        Call msubRecFO(rsb, strSQL)

        If rsb.RecordCount > 0 Then
            rsb.MoveFirst
            j = 27

            While Not rsb.EOF
                With oSheet
                    .Cells(j, 7) = Trim(IIf(IsNull(rsb!jmldpkfull.value), 0, (rsb!jmldpkfull.value)))
                    .Cells(j, 8) = Trim(IIf(IsNull(rsb!jmldpbfull.value), 0, (rsb!jmldpbfull.value)))
                    .Cells(j, 9) = Trim(IIf(IsNull(rsb!jmldaerahfull.value), 0, (rsb!jmldaerahfull.value)))
                    .Cells(j, 10) = Trim(IIf(IsNull(rsb!jmlpnkfull.value), 0, (rsb!jmlpnkfull.value)))
                    .Cells(j, 11) = Trim(IIf(IsNull(rsb!jmlabrifull.value), 0, (rsb!jmlabrifull.value)))
                    .Cells(j, 12) = Trim(IIf(IsNull(rsb!jmldeplainfull.value), 0, (rsb!jmldeplainfull.value)))
                    .Cells(j, 13) = Trim(IIf(IsNull(rsb!jmlpttfull.value), 0, (rsb!jmlpttfull.value)))
                    .Cells(j, 14) = Trim(IIf(IsNull(rsb!jmlswastafull.value), 0, (rsb!jmlswastafull.value)))
                    .Cells(j, 15) = Trim(IIf(IsNull(rsb!jmlkontrak.value), 0, (rsb!jmlkontrak.value)))
                    .Cells(j, 17) = Trim(IIf(IsNull(rsb!jmldpkpart.value), 0, (rsb!jmldpkpart.value)))
                    .Cells(j, 18) = Trim(IIf(IsNull(rsb!jmldpbpart.value), 0, (rsb!jmldpbpart.value)))
                    .Cells(j, 19) = Trim(IIf(IsNull(rsb!jmldaerahpart.value), 0, (rsb!jmldaerahpart.value)))
                    .Cells(j, 20) = Trim(IIf(IsNull(rsb!jmlpnkpart.value), 0, (rsb!jmlpnkpart.value)))
                    .Cells(j, 21) = Trim(IIf(IsNull(rsb!jmlabripart.value), 0, (rsb!jmlabripart.value)))
                    .Cells(j, 22) = Trim(IIf(IsNull(rsb!jmldeplainpart.value), 0, (rsb!jmldeplainpart.value)))
                    .Cells(j, 23) = Trim(IIf(IsNull(rsb!jmlpttpart.value), 0, (rsb!jmlpttpart.value)))
                    .Cells(j, 24) = Trim(IIf(IsNull(rsb!jmlswastapart.value), 0, (rsb!jmlswastapart.value)))
                    .Cells(j, 26) = Trim(IIf(IsNull(rsb!jmlhonorer.value), 0, (rsb!jmlhonorer.value)))
                End With
                j = j + 1
                rsb.MoveNext
            Wend
        End If

        j = j + 4
        Set rsx = Nothing
        strSQL = "SELECT  KdKualifikasiJurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak," & _
        "jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart," & _
        "jmlhonorer " & _
        "FROM RL4 WHERE  (KdkualifikasiJurusan IN ('0031', '0032','0033','0182')) " & _
        "GROUP BY  KdKualifikasiJurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak,jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart,jmlhonorer" & _
        " order by KdKualifikasiJurusan"
        Call msubRecFO(rsx, strSQL)

        If rsx.RecordCount > 0 Then
            rsx.MoveFirst
            j = 38

            While Not rsx.EOF
                With oSheet
                    .Cells(j, 7) = Trim(IIf(IsNull(rsx!jmldpkfull.value), 0, (rsx!jmldpkfull.value)))
                    .Cells(j, 8) = Trim(IIf(IsNull(rsx!jmldpbfull.value), 0, (rsx!jmldpbfull.value)))
                    .Cells(j, 9) = Trim(IIf(IsNull(rsx!jmldaerahfull.value), 0, (rsx!jmldaerahfull.value)))
                    .Cells(j, 10) = Trim(IIf(IsNull(rsx!jmlpnkfull.value), 0, (rsx!jmlpnkfull.value)))
                    .Cells(j, 11) = Trim(IIf(IsNull(rsx!jmlabrifull.value), 0, (rsx!jmlabrifull.value)))
                    .Cells(j, 12) = Trim(IIf(IsNull(rsx!jmldeplainfull.value), 0, (rsx!jmldeplainfull.value)))
                    .Cells(j, 13) = Trim(IIf(IsNull(rsx!jmlpttfull.value), 0, (rsx!jmlpttfull.value)))
                    .Cells(j, 14) = Trim(IIf(IsNull(rsx!jmlswastafull.value), 0, (rsx!jmlswastafull.value)))
                    .Cells(j, 15) = Trim(IIf(IsNull(rsx!jmlkontrak.value), 0, (rsx!jmlkontrak.value)))
                    .Cells(j, 17) = Trim(IIf(IsNull(rsx!jmldpkpart.value), 0, (rsx!jmldpkpart.value)))
                    .Cells(j, 18) = Trim(IIf(IsNull(rsx!jmldpbpart.value), 0, (rsx!jmldpbpart.value)))
                    .Cells(j, 19) = Trim(IIf(IsNull(rsx!jmldaerahpart.value), 0, (rsx!jmldaerahpart.value)))
                    .Cells(j, 20) = Trim(IIf(IsNull(rsx!jmlpnkpart.value), 0, (rsx!jmlpnkpart.value)))
                    .Cells(j, 21) = Trim(IIf(IsNull(rsx!jmlabripart.value), 0, (rsx!jmlabripart.value)))
                    .Cells(j, 22) = Trim(IIf(IsNull(rsx!jmldeplainpart.value), 0, (rsx!jmldeplainpart.value)))
                    .Cells(j, 23) = Trim(IIf(IsNull(rsx!jmlpttpart.value), 0, (rsx!jmlpttpart.value)))
                    .Cells(j, 24) = Trim(IIf(IsNull(rsx!jmlswastapart.value), 0, (rsx!jmlswastapart.value)))
                    .Cells(j, 26) = Trim(IIf(IsNull(rsx!jmlhonorer.value), 0, (rsx!jmlhonorer.value)))
                End With
                j = j + 1
                rsx.MoveNext
            Wend
        End If

    ElseIf Option4.value = True Then

        'Buka Excel
        Set oXL = CreateObject("Excel.Application")
        oXL.Visible = True
        'Buat Buka Template
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL4 Hal4.xls")
        Set oSheet = oWB.ActiveSheet

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("g5")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("y7", "y8")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing

        strSQL = "SELECT  KdKualifikasiJurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak," & _
        "jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart," & _
        "jmlhonorer " & _
        "FROM RL4 WHERE  (KdkualifikasiJurusan IN ('0095', '0096','0098','0100','102','103','0105','0107','0109','0110','0112','0113','0114','0116','0183','0184','0110','0185','0184','0102','0103')) " & _
        "GROUP BY  KdKualifikasiJurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak,jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart,jmlhonorer" & _
        " order by KdKualifikasiJurusan"
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount > 0 Then
            rs.MoveFirst
            j = 14

            While Not rs.EOF
                With oSheet
                    .Cells(j, 7) = Trim(IIf(IsNull(rs!jmldpkfull.value), 0, (rs!jmldpkfull.value)))
                    .Cells(j, 8) = Trim(IIf(IsNull(rs!jmldpbfull.value), 0, (rs!jmldpbfull.value)))
                    .Cells(j, 9) = Trim(IIf(IsNull(rs!jmldaerahfull.value), 0, (rs!jmldaerahfull.value)))
                    .Cells(j, 10) = Trim(IIf(IsNull(rs!jmlpnkfull.value), 0, (rs!jmlpnkfull.value)))
                    .Cells(j, 11) = Trim(IIf(IsNull(rs!jmlabrifull.value), 0, (rs!jmlabrifull.value)))
                    .Cells(j, 12) = Trim(IIf(IsNull(rs!jmldeplainfull.value), 0, (rs!jmldeplainfull.value)))
                    .Cells(j, 13) = Trim(IIf(IsNull(rs!jmlpttfull.value), 0, (rs!jmlpttfull.value)))
                    .Cells(j, 14) = Trim(IIf(IsNull(rs!jmlswastafull.value), 0, (rs!jmlswastafull.value)))
                    .Cells(j, 15) = Trim(IIf(IsNull(rs!jmlkontrak.value), 0, (rs!jmlkontrak.value)))
                    .Cells(j, 17) = Trim(IIf(IsNull(rs!jmldpkpart.value), 0, (rs!jmldpkpart.value)))
                    .Cells(j, 18) = Trim(IIf(IsNull(rs!jmldpbpart.value), 0, (rs!jmldpbpart.value)))
                    .Cells(j, 19) = Trim(IIf(IsNull(rs!jmldaerahpart.value), 0, (rs!jmldaerahpart.value)))
                    .Cells(j, 20) = Trim(IIf(IsNull(rs!jmlpnkpart.value), 0, (rs!jmlpnkpart.value)))
                    .Cells(j, 21) = Trim(IIf(IsNull(rs!jmlabripart.value), 0, (rs!jmlabripart.value)))
                    .Cells(j, 22) = Trim(IIf(IsNull(rs!jmldeplainpart.value), 0, (rs!jmldeplainpart.value)))
                    .Cells(j, 23) = Trim(IIf(IsNull(rs!jmlpttpart.value), 0, (rs!jmlpttpart.value)))
                    .Cells(j, 24) = Trim(IIf(IsNull(rs!jmlswastapart.value), 0, (rs!jmlswastapart.value)))
                    .Cells(j, 26) = Trim(IIf(IsNull(rs!jmlhonorer.value), 0, (rs!jmlhonorer.value)))
                End With
                j = j + 1
                rs.MoveNext
            Wend
        End If
    End If
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
End Sub

