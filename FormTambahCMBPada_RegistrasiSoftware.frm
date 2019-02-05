VERSION 5.00
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormTambahCMBPada_RegistrasiSoftware 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "--------------"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormTambahCMBPada_RegistrasiSoftware.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin Dacara_dcButton.dcButton cmOK 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&OK/Simpan"
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
      PicDown         =   "FormTambahCMBPada_RegistrasiSoftware.frx":0442
      PicHot          =   "FormTambahCMBPada_RegistrasiSoftware.frx":0794
      PicNormal       =   "FormTambahCMBPada_RegistrasiSoftware.frx":0AE6
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin XPControls.XPText textTambah 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Dacara_dcButton.dcButton cmBatal 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      PicDown         =   "FormTambahCMBPada_RegistrasiSoftware.frx":0E38
      PicHot          =   "FormTambahCMBPada_RegistrasiSoftware.frx":128A
      PicNormal       =   "FormTambahCMBPada_RegistrasiSoftware.frx":16DC
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
   Begin VB.Label LabelPenanda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kategori Baru"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1125
   End
End
Attribute VB_Name = "FormTambahCMBPada_RegistrasiSoftware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmOK_Click()
Select Case LabelPenanda.Caption
Case "1"
    If textTambah.Text = "" Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Silahkan isi Nama Kategori baru yang akan ditambahkan!", vbExclamation + vbOKOnly, ""
            textTambah.SetFocus
        Else
            MsgBox "Please write the name of New Category!", vbExclamation + vbOKOnly, ""
            textTambah.SetFocus
        End If
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            Pesan = MsgBox("Anda yakin ingin menambahkan '" & textTambah.Text & "' sebagai kategori baru?", vbQuestion + vbYesNo, "Konfirmasi?")
        Else
            Pesan = MsgBox("Are you sure to add '" & textTambah.Text & "' as new categori?", vbQuestion + vbYesNo, "Confirmation?")
        End If
            If Pesan = vbYes Then
                With Form_REGISTRASI_SOFTWARE
                    .AdodcKategori.Recordset.AddNew
                    .AdodcKategori.Recordset.Fields(0).Value = textTambah.Text
                    .AdodcKategori.Recordset.Update
                    .AdodcKategori.Refresh
                    .IsiCMBKategori
                    .cmbKategori.Text = textTambah.Text
                End With
                Unload Me
            End If
    End If
Case "2"
    If textTambah.Text = "" Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Silahkan isi Nama Lisensi baru yang akan ditambahkan!", vbExclamation + vbOKOnly, ""
            textTambah.SetFocus
        Else
            MsgBox "Please write the name of New License!", vbExclamation + vbOKOnly, ""
            textTambah.SetFocus
        End If
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            Pesan = MsgBox("Anda yakin ingin menambahkan '" & textTambah.Text & "' sebagai lisensi baru?", vbQuestion + vbYesNo, "Konfirmasi?")
        Else
            Pesan = MsgBox("Are you sure to add '" & textTambah.Text & "' as new license?", vbQuestion + vbYesNo, "Confirmation?")
        End If
            If Pesan = vbYes Then
                With Form_REGISTRASI_SOFTWARE
                    .AdodcJenisLisensi.Recordset.AddNew
                    .AdodcJenisLisensi.Recordset.Fields(0).Value = textTambah.Text
                    .AdodcJenisLisensi.Recordset.Update
                    .AdodcJenisLisensi.Refresh
                    .IsiCMBJenisLisensi
                    .cmbJenisLisensi.Text = textTambah.Text
                End With
                Unload Me
            End If
    End If
End Select
End Sub

Private Sub Form_Load()
    DisableCloseBtn Me
    PENGATURAN_WARNA
    PENGATURAN_BAHASA
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
        Label1.Caption = "Kategori Baru : "
        cmBatal.Caption = "&Batal"
        cmOK.Caption = "&OK/Simpan"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Label1.Caption = "New Category : "
        cmBatal.Caption = "&Cancel"
        cmOK.Caption = "&OK/Save"
    End If
End Sub

