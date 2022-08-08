VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakCatatanMedis0 
   Caption         =   "Medifirst2000 - Cetak Catatan Medis"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   Icon            =   "frmCetakCatatanMedis0.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   0   'False
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakCatatanMedis0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crCetakCatatanMedis0

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    Set frmCetakCatatanMedis = Nothing
    Set Report = New crCetakCatatanMedis0
    adocomd.ActiveConnection = dbConn
    strSQL = " select  A.NoCM, A.Title, A.NamaLengkap, isnull(B.NamaLengkap, 'fulan') as Bin, A.JenisKelamin, A.TempatLahir, " _
    & " A.TglLahir, dbo.S_HitungUmur(A.TglLahir, getdate()) as Umur, isnull(C.NamaLengkap, '-') as Suami, " _
    & " D.Pendidikan, D.Agama, D.Pekerjaan, " _
    & " (rtrim(ltrim(A.Alamat)) + ' RT/RW ' + rtrim(ltrim(A.RTRW)) + ' ' + rtrim(ltrim(A.Kelurahan)) + ' ' + rtrim(ltrim(A.Kecamatan))) as Alamat " _
    & " from    pasien A " _
    & " left outer join (select NoCM, NamaLengkap, Hubungan from KeluargaPasien " _
    & " where Hubungan like '%ayah%' or Hubungan like '%bapa%') B on A.NoCM = B.NoCM " _
    & " left outer join (select NoCM, NamaLengkap, Hubungan from KeluargaPasien " _
    & " where Hubungan like '%suami%' or Hubungan like '%istri%') C on A.NoCM = C.NoCM " _
    & " left outer join DetailPasien D on A.NoCM = D.NoCM " _
    & " WHERE A.NoCM ='" & mstrNoCM & "'"
    Call msubRecFO(rs, strSQL)

    With Report
        .txtNoCM.SetText mstrNoCM
        .txtNoCM1.SetText mstrNoCM
        .txtNoCM2.SetText mstrNoCM
        mstrNoCM = ""
        If IsNull(rs("NamaLengkap").value) = False Then .txtNama.SetText rs("Title").value & ". " & rs("NamaLengkap").value
        If IsNull(rs("TglLahir").value) = False Then .txtTglLahir.SetText Format(rs("TglLahir").value, "dd mmmm yyyy")
        If IsNull(rs("JenisKelamin").value) = False Then
            If rs("JenisKelamin").value = "P" Then
                .txtL.Font.Strikethrough = True
            ElseIf rs("JenisKelamin").value = "L" Then
                .txtP.Font.Strikethrough = True
            End If
        End If
        If IsNull(rs("TempatLahir").value) = False Then .txtKeluarga.SetText rs("TempatLahir").value
        If IsNull(rs("Umur").value) = False Then .txtUmur.SetText rs("Umur").value
        If IsNull(rs("Bin").value) = False Then .txtBin.SetText rs("Bin").value
        If IsNull(rs("Pekerjaan").value) = False Then .txtPekerjaan.SetText rs("Pekerjaan").value
        If IsNull(rs("Alamat").value) = False Then .txtAlamat.SetText rs("Alamat").value
        If IsNull(rs("Agama").value) = False Then .txtAgama.SetText rs("Agama").value
        If IsNull(rs("Pendidikan").value) = False Then .txtPendidikan.SetText rs("Pendidikan").value
        If IsNull(rs("Suami").value) = False Then .txtSuami.SetText rs("Suami").value
        .PrintOut False
    End With

    Screen.MousePointer = vbHourglass

    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakCatatanMedis = Nothing
End Sub

