VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMasterDiagnosa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Diagnosa"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMasterDiagnosa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   9105
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   0
      TabIndex        =   11
      Top             =   2280
      Width           =   9015
      Begin VB.CheckBox chkPilihSemua 
         Caption         =   "Pilih Semua"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   4800
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwDiagnosa 
         Height          =   4455
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nama Diagnosa"
            Object.Width           =   13229
         EndProperty
      End
      Begin VB.Label lblJumData 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data 0/0"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   8160
         TabIndex        =   14
         Top             =   4800
         Width           =   720
      End
   End
   Begin VB.TextBox txtParameterDiagnosa 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1440
      MaxLength       =   250
      TabIndex        =   10
      Top             =   7200
      Width           =   5295
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   9015
      Begin VB.TextBox txtDiagnosa 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3360
         TabIndex        =   6
         Top             =   480
         Width           =   5415
      End
      Begin MSDataListLib.DataCombo dcSubInstalasi 
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "SMF (Kasus Penyakit)"
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Diagnosa"
         Height          =   210
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   7680
      Width           =   9015
      Begin VB.CommandButton Tutup1 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   7440
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdubah 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   6000
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdbatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   4550
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
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
      Left            =   7200
      Picture         =   "frmMasterDiagnosa.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterDiagnosa.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmMasterDiagnosa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ValidasiError As String
Dim itemAll As ListItem
Dim rska As New ADODB.recordset
Dim strKdSubIns As String
Dim msg As VbMsgBoxResult
Dim i As Integer

Private Sub chkPilihSemua_Click()
    On Error GoTo errLoad
    Dim i As Integer

    If chkPilihSemua.value = vbChecked Then
        For i = 1 To lvwDiagnosa.ListItems.Count
            lvwDiagnosa.ListItems(i).Checked = True
        Next i
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdBatal_Click()
    Call blankfield
End Sub

'Private Sub cmdSimpanDiagnosa_Click()
'    On Error GoTo errLoad
'    Dim retVal As VbMsgBoxResult
'    Dim j As Integer
'
'    For i = 1 To lvwDiagnosa2.ListItems.Count
'        If lvwDiagnosa2.ListItems(i).Checked = True Then
'            dbConn.Execute "Update Diagnosa set NoDTD = '" & txtNoDtd.Text & "' where KdDiagnosa = '" & lvwDiagnosa2.ListItems(i).Key & "'"
'        End If
'    Next i
'    MsgBox "Update Data Sukses", vbOKOnly, "Informasi"
'    lvwDiagnosa2.ListItems.clear
'    Call loadGridDiagnosa
'    Exit Sub
'errLoad:
'    msubPesanError
'End Sub



Private Sub cmdUbah_Click()
    Dim retVal As VbMsgBoxResult
    
        If Periksa("datacombo", dcSubInstalasi, "SMF masih kosong") = False Then Exit Sub
    
        msg = MsgBox("Apakah anda yakin menyimpan data?", vbYesNo, "Informasi")
        If msg = vbNo Then Exit Sub
        If dcSubInstalasi.Text = "" Then
            MsgBox "Sub Instalasi Belum Diisi"
            dcSubInstalasi.SetFocus
            Exit Sub
        End If
        
'        If txtDiagnosa.Text = "" Then
'            MsgBox "Nama Diagnosa Belum Diisi"
'            txtDiagnosa.SetFocus
'            Exit Sub
'        End If
        Me.MousePointer = vbHourglass
        
        For i = 1 To lvwDiagnosa.ListItems.Count
            Me.Caption = i & "/" & lvwDiagnosa.ListItems.Count
            dbConn.Execute "DELETE FROM DiagnosaRuangan WHERE KdDiagnosa='" & lvwDiagnosa.ListItems(i).Key & "' AND KdSubInstalasi='" & strKdSubIns & "'"
            If lvwDiagnosa.ListItems(i).Checked = True Then
                dbConn.Execute "INSERT INTO DiagnosaRuangan(KdDiagnosa,KdSubInstalasi) " & "VALUES ('" & lvwDiagnosa.ListItems(i).Key & "','" & strKdSubIns & "')"
            End If
        Next i
        Me.MousePointer = vbDefault
        Me.Caption = "Medifirst2000 - Data Diagnosa"
        MsgBox "Penyimpanan data berhasil"
        dcSubInstalasi.SetFocus
    Call loadGrid
    Call blankfield

End Sub

Sub blankfield()
        dcSubInstalasi.Text = ""
        txtDiagnosa.Text = ""
        chkPilihSemua.value = 0
End Sub

Private Sub dcSubInstalasi_Change()
    On Error GoTo errLoad

    strKdSubIns = dcSubInstalasi.BoundText
    txtDiagnosa.Text = ""
    If dcSubInstalasi.BoundText = "" Then
        lvwDiagnosa.ListItems.clear
    Else
        Call loadGrid
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcSubInstalasi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    openConnection
    Call centerForm(Me, MDIUtama)
    Call blankfield
    Call loadGrid
    Me.MousePointer = vbDefault
End Sub

'Sub loadGridDiagnosa()
'    On Error GoTo errLoad
'
'    strSQL = "SELECT top 100 KdDiagnosa,NamaDiagnosa FROM Diagnosa WHERE KdDiagnosa LIKE '%" & txtSearchDiagnosa & "%' OR NamaDiagnosa like '%" & txtSearchDiagnosa & "%' order by NamaDiagnosa"
'
'    Set rs = Nothing
'    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
'
'    lvwDiagnosa2.ListItems.clear
'    Do While rs.EOF = False
'        Set itemAll = lvwDiagnosa2.ListItems.Add(, rs(0).value, rs(1).value)
'        rs.MoveNext
'    Loop
'
'    Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub

Sub loadGrid()
    On Error GoTo errLoad

        strSQL = "SELECT KdSubInstalasi,NamaSubInstalasi FROM SubInstalasi where StatusEnabled='1' "
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        Set dcSubInstalasi.RowSource = rs
        dcSubInstalasi.ListField = rs(1).Name
        dcSubInstalasi.BoundColumn = rs(0).Name

        dcSubInstalasi.SetFocus
    Exit Sub
errLoad:
End Sub

Private Sub lvwDiagnosa_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If dcSubInstalasi.Text = "" Then
        lvwDiagnosa.ListItems(Item.Key).Checked = False
        Exit Sub
    End If
End Sub

Private Sub Tutup1_Click()
    Unload Me
End Sub

Private Sub subClearData()
    dcSubInstalasi.Text = ""
    txtDiagnosa.Text = ""
    lvwDiagnosa.ListItems.clear
End Sub

Private Sub txtDiagnosa_Change()
    On Error Resume Next
    Dim jmlData As Integer

    jmlData = 0

'    Me.MousePointer = vbHourglass

    strSQL = "SELECT top 100 KdDiagnosa,NamaDiagnosa FROM Diagnosa WHERE NamaDiagnosa LIKE '%" & txtDiagnosa.Text & "%' " 'ORDER BY NamaDiagnosa ASC"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    lvwDiagnosa.ListItems.clear
    Do While rs.EOF = False
        Set itemAll = lvwDiagnosa.ListItems.Add(, rs(0).value, rs(1).value)
        rs.MoveNext
    Loop

    strSQL = "SELECT  top 100 DiagnosaRuangan.KdDiagnosa, Diagnosa.NamaDiagnosa AS NamaDiagnosa, DiagnosaRuangan.KdSubInstalasi FROM Diagnosa INNER JOIN" _
    & " DiagnosaRuangan ON Diagnosa.KdDiagnosa = DiagnosaRuangan.KdDiagnosa" _
    & " WHERE KdSubInstalasi='" & strKdSubIns & "' AND NamaDiagnosa LIKE '" & txtDiagnosa.Text & "%' "
    Set rska = Nothing
    rska.Open strSQL, dbConn, adOpenForwardOnly
    If rska.EOF = True Then Exit Sub
    For i = 1 To lvwDiagnosa.ListItems.Count
        Me.Caption = jmlData & "/" & i & "/" & lvwDiagnosa.ListItems.Count
        If rska.EOF = True Then
            GoTo d
        Else
            If lvwDiagnosa.ListItems(i).Key = rska(0).value Then
                lvwDiagnosa.ListItems(i).Checked = True
                lvwDiagnosa.ListItems(i).ForeColor = vbBlue
                jmlData = jmlData + 1
                rska.MoveNext
            Else
                GoTo d
            End If

        End If
d:
    Next i
    lblJumData.Caption = "Data " & jmlData & "/" & lvwDiagnosa.ListItems.Count
    Me.MousePointer = vbDefault
    Me.Caption = "Medifirst2000 - Data Diagnosa"
    Exit Sub
errLoad:
    Call msubPesanError
    Me.MousePointer = vbDefault
End Sub

Private Sub txtDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
