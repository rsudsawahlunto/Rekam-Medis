VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmRL4b 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - RL4 B Jumlah Tenaga Non Kesehatan Menurut Jenis"
   ClientHeight    =   3210
   ClientLeft      =   6960
   ClientTop       =   3645
   ClientWidth     =   7560
   Icon            =   "frmRL4b.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7560
   Begin VB.OptionButton Option3 
      Caption         =   "Hal. 7"
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Hal.6"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Hal. 5"
      Height          =   495
      Left            =   1440
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
      Caption         =   "RL 4 B Halaman"
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
      Picture         =   "frmRL4b.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2955
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRL4b.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRL4b.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmRL4b"
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
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL4 Hal5.xls")
        Set oSheet = oWB.ActiveSheet

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("g5")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("y7", "y8")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing

        strSQL = "SELECT  kdkualifikasijurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak," & _
        "(sum(jmldpkfull)+(jmldpbfull)+(jmldaerahfull)+(jmlpnkfull)+(jmlabrifull)+(jmldeplainfull)+(jmlpttfull)+(jmlswastafull)+(jmlkontrak)) as subtotal1, " & _
        "jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart," & _
        "(sum(jmldpkpart)+(jmldpbpart)+(jmldaerahpart)+(jmlpnkpart)+(jmlabripart)+(jmldeplainpart)+(jmlpttpart)+(jmlswastapart))as subtotal2," & _
        "jmlhonorer " & _
        "FROM RL4 WHERE  (KdkualifikasiJurusan IN ('0115', '0117','0118','0119','0120','0121','0122','0123','0124','0125','0126','0127')) " & _
        "GROUP BY  kdkualifikasijurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak,jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart,jmlhonorer" & _
        " order by kdkualifikasijurusan"
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
        strSQL = "SELECT  kdkualifikasijurusan, kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak," & _
        "(sum(jmldpkfull)+(jmldpbfull)+(jmldaerahfull)+(jmlpnkfull)+(jmlabrifull)+(jmldeplainfull)+(jmlpttfull)+(jmlswastafull)+(jmlkontrak)) as subtotal1, " & _
        "jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart," & _
        "(sum(jmldpkpart)+(jmldpbpart)+(jmldaerahpart)+(jmlpnkpart)+(jmlabripart)+(jmldeplainpart)+(jmlpttpart)+(jmlswastapart))as subtotal2," & _
        "jmlhonorer " & _
        "FROM RL4 WHERE  (KdkualifikasiJurusan IN ('0128', '0129','0130','0131','0132','0133','0134','0135','0136','0137','0138','0139','0140')) " & _
        "GROUP BY  kdkualifikasijurusan, kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak,jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart,jmlhonorer" & _
        " order by kdkualifikasijurusan"
        Call msubRecFO(rsb, strSQL)

        If rsb.RecordCount > 0 Then
            rsb.MoveFirst
            j = 30

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
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL4 Hal6.xls")
        Set oSheet = oWB.ActiveSheet

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("g5")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("y7", "y8")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing

        strSQL = "SELECT  kdkualifikasijurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak," & _
        "(sum(jmldpkfull)+(jmldpbfull)+(jmldaerahfull)+(jmlpnkfull)+(jmlabrifull)+(jmldeplainfull)+(jmlpttfull)+(jmlswastafull)+(jmlkontrak)) as subtotal1, " & _
        "jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart," & _
        "(sum(jmldpkpart)+(jmldpbpart)+(jmldaerahpart)+(jmlpnkpart)+(jmlabripart)+(jmldeplainpart)+(jmlpttpart)+(jmlswastapart))as subtotal2," & _
        "jmlhonorer " & _
        "FROM RL4 WHERE  (KdkualifikasiJurusan IN ('0062', '0063','0066','0068','0070','0072','0075','0077','0079','0080','0082')) " & _
        "GROUP BY  kdkualifikasijurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak,jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart,jmlhonorer" & _
        " order by kdkualifikasijurusan"
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
        strSQL = "SELECT  kdkualifikasijurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak," & _
        "(sum(jmldpkfull)+(jmldpbfull)+(jmldaerahfull)+(jmlpnkfull)+(jmlabrifull)+(jmldeplainfull)+(jmlpttfull)+(jmlswastafull)+(jmlkontrak)) as subtotal1, " & _
        "jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart," & _
        "(sum(jmldpkpart)+(jmldpbpart)+(jmldaerahpart)+(jmlpnkpart)+(jmlabripart)+(jmldeplainpart)+(jmlpttpart)+(jmlswastapart))as subtotal2," & _
        "jmlhonorer " & _
        "FROM RL4 WHERE  (KdkualifikasiJurusan IN ('0087', '0088','0091','0093','0094','0186','0097','0101','0104','0106','0108','0111')) " & _
        "GROUP BY  kdkualifikasijurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak,jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart,jmlhonorer" & _
        " order by kdkualifikasijurusan"
        Call msubRecFO(rsb, strSQL)

        If rsb.RecordCount > 0 Then
            rsb.MoveFirst
            j = 30

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
        Set oWB = oXL.Workbooks.Open(App.Path & "\RL4 Hal7.xls")
        Set oSheet = oWB.ActiveSheet

        Set rsb = Nothing
        strSQL = "select * from profilrs"
        Call msubRecFO(rsb, strSQL)

        Set oResizeRange = oSheet.Range("g5")
        oResizeRange.value = Trim(rsb!NamaRS)

        Set oResizeRange = oSheet.Range("y7", "y8")
        oResizeRange.value = Trim(rsb!KdRs)

        Set rs = Nothing

        strSQL = "SELECT  kdkualifikasijurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak," & _
        "(sum(jmldpkfull)+(jmldpbfull)+(jmldaerahfull)+(jmlpnkfull)+(jmlabrifull)+(jmldeplainfull)+(jmlpttfull)+(jmlswastafull)+(jmlkontrak)) as subtotal1, " & _
        "jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart," & _
        "(sum(jmldpkpart)+(jmldpbpart)+(jmldaerahpart)+(jmlpnkpart)+(jmlabripart)+(jmldeplainpart)+(jmlpttpart)+(jmlswastapart))as subtotal2," & _
        "jmlhonorer " & _
        "FROM RL4 WHERE  (KdkualifikasiJurusan IN ('0013', '0047','0048','0051','0050','0052')) " & _
        "GROUP BY  kdkualifikasijurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak,jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart,jmlhonorer" & _
        " order by kdkualifikasijurusan"
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
        strSQL = "SELECT  kdkualifikasijurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak," & _
        "(sum(jmldpkfull)+(jmldpbfull)+(jmldaerahfull)+(jmlpnkfull)+(jmlabrifull)+(jmldeplainfull)+(jmlpttfull)+(jmlswastafull)+(jmlkontrak)) as subtotal1, " & _
        "jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart," & _
        "(sum(jmldpkpart)+(jmldpbpart)+(jmldaerahpart)+(jmlpnkpart)+(jmlabripart)+(jmldeplainpart)+(jmlpttpart)+(jmlswastapart))as subtotal2," & _
        "jmlhonorer " & _
        "FROM RL4 WHERE  (KdkualifikasiJurusan IN ('0057', '0059')) " & _
        "GROUP BY  kdkualifikasijurusan,kualifikasijurusan,jmldpkfull,jmldpbfull,jmldaerahfull,jmlpnkfull,jmlabrifull,jmldeplainfull,jmlpttfull,jmlswastafull,jmlkontrak,jmldpkpart,jmldpbpart,jmldaerahpart,jmlpnkpart,jmlabripart,jmldeplainpart,jmlpttpart,jmlswastapart,jmlhonorer" & _
        " order by kdkualifikasijurusan"
        Call msubRecFO(rsb, strSQL)

        If rsb.RecordCount > 0 Then
            rsb.MoveFirst
            j = 24

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

