VERSION 5.00
Object = "{02353968-C1C9-4E0A-88D3-18759BDC60FE}#1.0#0"; "AeroSuite.ocx"
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormExportDataToExcel 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormExportDataToExcel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pilihan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   1800
      TabIndex        =   8
      Top             =   840
      Width           =   4695
      Begin AeroSuite.AeroCheckBox cekBukaFile 
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   529
         Align           =   0
         Caption         =   "Buka File setelah export"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         BackColor       =   14737632
         ForeColor       =   0
         MousePointer    =   0
         MouseIcon       =   "FormExportDataToExcel.frx":0442
         Value           =   0
      End
      Begin AeroSuite.AeroCheckBox cekBukaFolder 
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   529
         Align           =   0
         Caption         =   "Buka Folder setelah export"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         BackColor       =   14737632
         ForeColor       =   0
         MousePointer    =   0
         MouseIcon       =   "FormExportDataToExcel.frx":045E
         Value           =   0
      End
   End
   Begin XPControls.XPText TextLokasi 
      Height          =   345
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   609
      Text            =   "XPText1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Dacara_dcButton.dcButton cmBrowse 
      Height          =   345
      Left            =   5280
      TabIndex        =   1
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "Jelajahi..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FormExportDataToExcel.frx":047A
      PicHot          =   "FormExportDataToExcel.frx":D539
      PicNormal       =   "FormExportDataToExcel.frx":1A5F8
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin XPControls.XPText textNamaFile 
      Height          =   345
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   609
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Dacara_dcButton.dcButton cmExport 
      Height          =   345
      Left            =   5280
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Export..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   5
      PicDown         =   "FormExportDataToExcel.frx":276B7
      PicHot          =   "FormExportDataToExcel.frx":29E69
      PicNormal       =   "FormExportDataToExcel.frx":2C61B
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmBatal 
      Height          =   345
      Left            =   3960
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Batal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FormExportDataToExcel.frx":2EDCD
      PicHot          =   "FormExportDataToExcel.frx":2F21F
      PicNormal       =   "FormExportDataToExcel.frx":2F671
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin XPEngine.XPControl XP_Engine 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama File"
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lokasi/Path"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   945
   End
End
Attribute VB_Name = "FormExportDataToExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
With TextLokasi
    .Locked = True
    .ForeColor = Silver
End With
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        TextLokasi.Text = "Klik tombol 'Jelajahi...'"
        Label2.Caption = "Nama File"
        Label1.Caption = "Lokasi/Path"
        cekBukaFile.Caption = "Buka File setelah Ekspor"
        cekBukaFolder.Caption = "Buka Folder setelah Ekspor"
        cmBrowse.Caption = "Jelajahi..."
        cmBatal.Caption = "&Batal"
        Frame1.Caption = "Pilihan"
    Else
        TextLokasi.Text = "Click 'Browse...'"
        Label2.Caption = "File Name"
        Label1.Caption = "Location/Path"
        cekBukaFile.Caption = "Open File after Export"
        cekBukaFolder.Caption = "Open Folder after Export"
        cmBrowse.Caption = "Browse..."
        cmBatal.Caption = "&Cancel"
        Frame1.Caption = "Options"
    End If
    DisableCloseBtn Me
    'PENGATURAN UNTUK ALWAYS ON TOP
    If FormPengaturan.cekAlwaysOnTop.Value = Checked Then
        SetOnTop (Me.hwnd)
    ElseIf FormPengaturan.cekAlwaysOnTop.Value = Unchecked Then
        NotOnTop (Me.hwnd)
    End If
    For Each Objek In Me
        If TypeName(Objek) = "Label" Or TypeName(Objek) = "dcButton" Or TypeName(Objek) = "AeroCheckBox" Then
            With Objek
                .Font.Name = "Agency FB"
                .Font.Size = 11
            End With
        End If
        If TypeName(Objek) = "XPText" Then Objek.Font.Name = "Agency FB"
    Next
    XP_Engine.StartEngine
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmBrowse_Click()
    BrowseForFolder (Me.hwnd)
    TextLokasi.ToolTipText = TextLokasi.Text
End Sub

'FUNCTION YANG DIPAKAI UNTUK MEMBUAT BROWSE FOR FOLDER
Public Function BrowseForFolder(hwnd As Long, Optional Title As String = "Browse For Folder") As String
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        szTitle = "Pilih folder untuk menyimpan file"
    Else
        szTitle = "Select a folder to save the file"
    End If
    With tBrowseInfo
        .hWndOwner = FormExportDataToExcel.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        BrowseForFolder = sBuffer
            With TextLokasi
                .Text = sBuffer
                .ForeColor = vbBlack
            End With
    End If
End Function

Private Sub cmExport_Click()
    If textNamaFile.Text = "" Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Silahkan isi Nama File!", vbExclamation + vbOKOnly, ""
            textNamaFile.SetFocus
        Else
            MsgBox "Please write the File Name!", vbExclamation + vbOKOnly, ""
            textNamaFile.SetFocus
        End If
    ElseIf TextLokasi.Text = "Klik tombol 'Jelajahi...'" Or TextLokasi.Text = "Click 'Browse...'" Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Silahkan klik 'Jelajahi' untuk menentukan lokasi file!", vbExclamation + vbOKOnly, ""
            cmBrowse.SetFocus
        Else
            MsgBox "Please click 'Browse' for the location of the file!", vbExclamation + vbOKOnly, ""
            cmBrowse.SetFocus
        End If
    Else
        If Me.Caption = "Export Data >> Excel 2003" Then
            X = TextLokasi.Text & "\" & textNamaFile.Text & ".xls"
        ElseIf Me.Caption = "Export Data >> Word 2003" Then
            X = TextLokasi.Text & "\" & textNamaFile.Text & ".doc"
        End If
        Open X For Output As #1
        'Menyimpan semua data
        With FormBuatAkunBaru
            Print #1, "==============" & .Frame1.Caption & "=============="
            Print #1, .Label24; .Caption & " : " & .textNamaPengguna.Text
            Print #1, .Label26.Caption & " : " & .textPasswordBaru.Text
            Print #1, .Label29.Caption & " : " & .textKonfirmasiPassword.Text
            Print #1, ""
            Print #1, "==============" & .Frame2.Caption & "=============="
            Print #1, .Label1.Caption & " : " & .textNamaAsli.Text
            Print #1, .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text
            Print #1, .Label7.Caption & " : " & .cmbJenisKelamin.Text
            Print #1, .Label10.Caption & " : " & .cmbAgama.Text
            Print #1, .Label9.Caption & " : " & .textHobby.Text
            Print #1, .Label12.Caption & " : " & .textAlamat.Text
            Print #1, .Label15.Caption & " : " & .textNomorTelepon.Text
            Print #1, .Label17.Caption & " : " & .textAlamatEmail.Text
            Print #1, .Label19.Caption & " : " & .cmbAlamatWebsite.Text & .textAlamatWebsite.Text
            Print #1, .Label20.Caption & " : " & .textStatusAktivitas.Text
            Print #1, .Label22.Caption & " : " & .textStatusHubungan.Text
            Print #1, ""
            Print #1, "==============" & .Frame3.Caption & "=============="
            Print #1, .Label35.Caption & " : " & .cmbPertanyaanRahasia.Text
            Print #1, .Label33.Caption & " : " & .textJawaban.Text
        End With
       'Menutup file
        Close #1
        If cekBukaFile.Value = Checked Then OpenLocation X, SHOWNORMAL
        If cekBukaFolder.Value = Checked Then OpenDirectory (TextLokasi.Text)
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Data berhasil di ekspor ke : " & vbCrLf & _
                    X, vbInformation + vbOKOnly, ""
        Else
            MsgBox "Data successfully exported to : " & vbCrLf & _
                    X, vbInformation + vbOKOnly, ""
        End If
    End If
        
End Sub

Private Sub Form_Load()
    AturKontrol
    PENGATURAN_WARNA
End Sub

Sub PENGATURAN_WARNA()
    'PENGATURAN WARNA UNTUK FORM INI
    For Each Objek In Me
        Select Case FormPengaturan.cmbWarnaTampilan.ListIndex
        Case Is = 0 'Ungu Natural
            Me.BackColor = UnguNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguNatural
            If TypeName(Objek) = "Frame" Then Objek.BackColor = UnguNatural
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = UnguNatural
        Case Is = 1 'Merah
            Me.BackColor = Merah
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Merah
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Merah
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = Merah
        Case Is = 2 'Pink
            Me.BackColor = Pink
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Pink
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Pink
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = Pink
        Case Is = 3 'HijauMuda
            Me.BackColor = HijauMuda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = HijauMuda
            If TypeName(Objek) = "Frame" Then Objek.BackColor = HijauMuda
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = HijauMuda
        Case Is = 4 'Hitam
            Me.BackColor = Hitam
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Hitam
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Hitam
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = Hitam
        Case Is = 5 'Silver
            Me.BackColor = Silver
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Silver
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Silver
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = Silver
        Case Is = 6 'SilverNatural
            Me.BackColor = SilverNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = SilverNatural
            If TypeName(Objek) = "Frame" Then Objek.BackColor = SilverNatural
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = SilverNatural
        Case Is = 7 'Orange
            Me.BackColor = Orange
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Orange
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Orange
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = Orange
        Case Is = 8 'UnguJanda
            Me.BackColor = UnguJanda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguJanda
            If TypeName(Objek) = "Frame" Then Objek.BackColor = UnguJanda
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = UnguJanda
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
