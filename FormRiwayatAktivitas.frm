VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormRiwayatAktivitas 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Riwayat Aktivitas"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11670
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormRiwayatAktivitas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdodcUtama 
      Height          =   330
      Left            =   3120
      Top             =   4560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Dacara_dcButton.dcButton cmBersihkanRiwayat 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "Bersihkan "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FormRiwayatAktivitas.frx":1085C
      PicHot          =   "FormRiwayatAktivitas.frx":10CAE
      PicNormal       =   "FormRiwayatAktivitas.frx":11100
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   7435
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin Dacara_dcButton.dcButton cmSetting 
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Setting"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   5
      PicDown         =   "FormRiwayatAktivitas.frx":11552
      PicHot          =   "FormRiwayatAktivitas.frx":21DBE
      PicNormal       =   "FormRiwayatAktivitas.frx":3262A
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
   End
   Begin Dacara_dcButton.dcButton cmTutup 
      Height          =   495
      Left            =   9840
      TabIndex        =   3
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Tutup"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   5
      PicDown         =   "FormRiwayatAktivitas.frx":42E96
      PicHot          =   "FormRiwayatAktivitas.frx":432E8
      PicNormal       =   "FormRiwayatAktivitas.frx":4373A
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
   End
   Begin XPEngine.XPControl XP_Engine 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
End
Attribute VB_Name = "FormRiwayatAktivitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
If FormPengaturan.cmbBahasa.ListIndex = 0 Then
    With LV
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Tanggal", 1400
        .ColumnHeaders.Add , , "Waktu", 1400, vbCenter
        .ColumnHeaders.Add , , "Aktifitas", 6500
        .ColumnHeaders.Add , , "Nama Komputer", 2000, vbCenter
        .View = lvwReport
        .Sorted = True
    End With
ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
    With LV
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Date", 1400
        .ColumnHeaders.Add , , "Time", 1400, vbCenter
        .ColumnHeaders.Add , , "Activity", 6500
        .ColumnHeaders.Add , , "Computer Name", 2000, vbCenter
        .View = lvwReport
        .Sorted = True
    End With
End If
    MasukkanDataKeDalamTabel
    If AdodcUtama.Recordset.RecordCount = 0 Then
        Me.cmBersihkanRiwayat.Enabled = False
    Else
        Me.cmBersihkanRiwayat.Enabled = True
    End If
    'PENGATURAN UNTUK ALWAYS ON TOP
    If FormPengaturan.cekAlwaysOnTop.Value = Checked Then
        SetOnTop (Me.hwnd)
    ElseIf FormPengaturan.cekAlwaysOnTop.Value = Unchecked Then
        NotOnTop (Me.hwnd)
    End If
    For Each Objek In Me
        If TypeName(Objek) = "Label" Or TypeName(Objek) = "dcButton" Or TypeName(Objek) = "AeroCheckBox" Or TypeName(Objek) = "TextBox" Or TypeName(Objek) = "ComboBox" Then
            With Objek
                .Font.Name = "Agency FB"
                .Font.Size = 11
            End With
        End If
        If TypeName(Objek) = "XPText" Then Objek.Font.Name = "Agency FB"
    Next
    XP_Engine.StartEngine
End Sub

Sub MasukkanDataKeDalamTabel()
        If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
            CN_FormUtama.CursorLocation = adUseClient
            CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
        With AdodcUtama
            .ConnectionString = CN_FormUtama.ConnectionString
            .RecordSource = "Select * From tbRiwayat"
            .Refresh
        End With
        LV.ListItems.Clear
        Do Until AdodcUtama.Recordset.EOF
        Set LI = LV.ListItems.Add(, , AdodcUtama.Recordset.Fields(0).Value & " - " & AdodcUtama.Recordset.Fields(1).Value & " - " & AdodcUtama.Recordset.Fields(2).Value)
            LI.SubItems(1) = AdodcUtama.Recordset.Fields(3).Value & ":" & AdodcUtama.Recordset.Fields(4).Value & ":" & AdodcUtama.Recordset.Fields(5).Value
            LI.SubItems(2) = AdodcUtama.Recordset.Fields(6).Value
            LI.SubItems(3) = AdodcUtama.Recordset.Fields(7).Value
            AdodcUtama.Recordset.MoveNext
        Loop
        AdodcUtama.Refresh
End Sub


Private Sub cmBersihkanRiwayat_Click()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Pesan = MsgBox("Apakah Anda yakin ingin membersihkan data riwayat aktivitas?" & vbCrLf & _
                    "(Ingat, perintah ini tidak dapat dibatalkan!)", vbQuestion + vbYesNo, "Konfirmasi?")
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Pesan = MsgBox("Are you sure want to clear the activity history data?" & vbCrLf & _
                    "(Remember, this command can't be undone!)", vbQuestion + vbYesNo, "Confirmation?")
    End If
        If Pesan = vbYes Then
            On Error Resume Next
            If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
                CN_FormUtama.CursorLocation = adUseClient
                CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
            With AdodcUtama
                .ConnectionString = CN_FormUtama.ConnectionString
                .RecordSource = "Delete From tbRiwayat"
                .Refresh
            End With
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                MsgBox "Data riwayat berhasil dibersihkan!", vbInformation + vbOKOnly, "Berhasil"
            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                MsgBox "History Data is cleared!", vbInformation + vbOKOnly, "Success"
            End If
            AturKontrol
            MasukkanDataKeDalamTabel
        End If
End Sub

Private Sub cmSetting_Click()
    FormSettingRiwayatLebihLanjut.Show vbModal, Me
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    AturKontrol
    PENGATURAN_WARNA
    PENGATURAN_BAHASA
    PENGATURAN_FORM
End Sub

Private Sub Form_Load()
    AturKontrol
    PENGATURAN_WARNA
    PENGATURAN_BAHASA
    PENGATURAN_FORM
End Sub

Sub PENGATURAN_WARNA()
    'PENGATURAN WARNA UNTUK FORM INI
    For Each Objek In Me
        Select Case FormPengaturan.cmbWarnaTampilan.ListIndex
        Case Is = 0 'Ungu Natural
            Me.BackColor = UnguNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguNatural
        Case Is = 1 'Merah
            Me.BackColor = Merah
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Merah
        Case Is = 2 'Pink
            Me.BackColor = Pink
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Pink
        Case Is = 3 'HijauMuda
            Me.BackColor = HijauMuda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = HijauMuda
        Case Is = 4 'Hitam
            Me.BackColor = Hitam
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Hitam
        Case Is = 5 'Silver
            Me.BackColor = Silver
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Silver
        Case Is = 6 'SilverNatural
            Me.BackColor = SilverNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = SilverNatural
        Case Is = 7 'Orange
            Me.BackColor = Orange
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Orange
        Case Is = 8 'UnguJanda
            Me.BackColor = UnguJanda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguJanda
        End Select
    Next
    'PENGATURAN THEMA UNTUK FORM INI
    For Each Objek In Me
        If TypeName(Objek) = "dcButton" Then
            Select Case FormPengaturan.cmbTemaTampilan.ListIndex
            Case Is = 0 'RST_Office 2003
                Objek.ButtonStyle = 3
                Objek.BackColor = &HC0C0C0
            Case Is = 1 'RST_Office XP
                Objek.ButtonStyle = 4
            Case Is = 2 'RST_Opera Browser
                Objek.ButtonStyle = 5
            Case Is = 3 'RST_Classic
                Objek.ButtonStyle = 6
            Case Is = 4 'RST_XP Blue
                Objek.ButtonStyle = 7
            Case Is = 5 'RST_XP Olive Green
                Objek.ButtonStyle = 8
            Case Is = 6 'RST_XP Silver
                Objek.ButtonStyle = 9
            Case Is = 7 'RST_XP Toolbar
                Objek.ButtonStyle = 10
            Case Is = 8 'RST_Yahoo
                Objek.ButtonStyle = 11
                Objek.BackColor = &H12BCFF
            Case Is = 9 'RST_Mac
                Objek.ButtonStyle = 1
                Objek.BackColor = &HFF9B48
            Case Is = 10 'RST_MacOSX
                Objek.ButtonStyle = 2
            End Select
        End If
    Next
End Sub

Sub PENGATURAN_BAHASA()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Me.Caption = "Riwayat Aktivitas"
        cmBersihkanRiwayat.Caption = "Bersihkan"
        cmSetting.Caption = "&Pengaturan"
        cmTutup.Caption = "&Tutup"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Me.Caption = "History Activity"
        cmBersihkanRiwayat.Caption = "Clean Up"
        cmSetting.Caption = "&Setting"
        cmTutup.Caption = "&Close"
    End If
End Sub

Sub PENGATURAN_FORM()
    If FormPengaturan.cekGarisGrid.Value = Checked Then
        LV.GridLines = True
    ElseIf FormPengaturan.cekGarisGrid.Value = Unchecked Then
        LV.GridLines = False
    End If
    
End Sub
