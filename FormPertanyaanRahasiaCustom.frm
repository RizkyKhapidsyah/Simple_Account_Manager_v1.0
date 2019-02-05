VERSION 5.00
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormPertanyaanRahasiaCustom 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormPertanyaanRahasiaCustom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox textPertanyaanRahasia 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "FormPertanyaanRahasiaCustom.frx":0442
      Top             =   480
      Width           =   5295
   End
   Begin Dacara_dcButton.dcButton cmSimpan 
      Height          =   345
      Left            =   4320
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Simpan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   5
      PicDown         =   "FormPertanyaanRahasiaCustom.frx":0448
      PicHot          =   "FormPertanyaanRahasiaCustom.frx":079A
      PicNormal       =   "FormPertanyaanRahasiaCustom.frx":0AEC
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmBatal 
      Height          =   345
      Left            =   3120
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Batal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FormPertanyaanRahasiaCustom.frx":0E3E
      PicHot          =   "FormPertanyaanRahasiaCustom.frx":1290
      PicNormal       =   "FormPertanyaanRahasiaCustom.frx":16E2
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
   Begin VB.Label LabelJumlahHuruf 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "FormPertanyaanRahasiaCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Label1.Caption = "Isi pertanyaan sesuai dengan keinginan Anda : (max = 254 karakter)"
        cmBatal.Caption = "&Batal"
        cmSimpan.Caption = "&Simpan"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Label1.Caption = "Please fill in the questions as you want here : (max = 254 character)"
        cmBatal.Caption = "&Cancel"
        cmSimpan.Caption = "&Save"
    End If
    With textPertanyaanRahasia
        .Text = ""
        .MaxLength = 254
    End With
    LabelJumlahHuruf.Caption = "0"
    DisableCloseBtn Me
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

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmSimpan_Click()
If textPertanyaanRahasia.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Pertanyaan Rahasia yang Anda inginkan.", vbExclamation + vbOKOnly, "Tidak boleh kosong"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please fill in your type question.", vbExclamation + vbOKOnly, "Don't empty"
    End If
    textPertanyaanRahasia.SetFocus
Else
    Select Case FormPengaturan.cmbBahasa.ListIndex
    Case Is = 0
        Pesan = MsgBox("Anda yakin dengan Isian Pertanyaan Rahasia Anda?", vbQuestion + vbYesNo, "Konfirmasi")
    Case Is = 1
        Pesan = MsgBox("Are You sure with this Security Question?", vbQuestion + vbYesNo, "Confirmation")
    End Select
        If Pesan = vbYes Then
            With FormBuatAkunBaru
                .AdodcPertanyaanRahasia.Recordset.AddNew
                .AdodcPertanyaanRahasia.Recordset.Fields(0).Value = textPertanyaanRahasia.Text
                .AdodcPertanyaanRahasia.Recordset.Update
                .AdodcPertanyaanRahasia.Refresh
                .IsiCMBPertanyaanRahasia
                .cmbPertanyaanRahasia.Text = textPertanyaanRahasia.Text
                Unload Me
            End With
        ElseIf Pesan = vbNo Then
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                MsgBox "Saran : Harap mengisi pertanyaan yang mudah untuk diingat.", vbExclamation + vbOKOnly, "Saran"
            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                MsgBox "Suggestion : Please fill out a questionnaire that you can remember", vbExclamation + vbOKOnly, "Suggestion"
            End If
            textPertanyaanRahasia.SetFocus
        End If
End If
End Sub

Private Sub Form_Load()
    AturKontrol
    PENGATURAN_WARNA
End Sub

Private Sub textPertanyaanRahasia_Change()
    LabelJumlahHuruf.Caption = Len(textPertanyaanRahasia.Text)
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
