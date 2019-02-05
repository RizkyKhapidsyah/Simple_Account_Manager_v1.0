VERSION 5.00
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormKalender 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Kelender"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   855
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pra Lihat :"
         Height          =   270
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   825
      End
   End
   Begin Dacara_dcButton.dcButton cmOK 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&OK"
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
      PicDown         =   "FormKalender.frx":0000
      PicHot          =   "FormKalender.frx":031A
      PicNormal       =   "FormKalender.frx":0634
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
   End
   Begin XPControls.XPText textTanggal 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Text            =   "99 - 99 - 9999"
      Alignment       =   2
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
   Begin MSACAL.Calendar Kalender 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _Version        =   524288
      _ExtentX        =   8705
      _ExtentY        =   4895
      _StockProps     =   1
      BackColor       =   14737632
      Year            =   2012
      Month           =   10
      Day             =   27
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Dacara_dcButton.dcButton cmBatal 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
      PicDown         =   "FormKalender.frx":094E
      PicHot          =   "FormKalender.frx":0DA0
      PicNormal       =   "FormKalender.frx":11F2
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   3495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   5175
   End
   Begin XPEngine.XPControl XP_Engine 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
End
Attribute VB_Name = "FormKalender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    textTanggal.Locked = True
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        cmBatal.Caption = "&Batal"
        Label1.Caption = "Pra Lihat :"
    Else
        cmBatal.Caption = "&Cancel"
        Label1.Caption = "Preview :"
    End If
    Kalender.Today
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

Private Sub cmOK_Click()
If FORM_UTAMA.cmJejaringSosial.FontBold = True Then
    If textTanggal.Text = "0 - 0 - 0" Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Tanggal Salah!", vbCritical + vbOKOnly, ""
        Else
            MsgBox "Wrong Time!", vbCritical + vbOKOnly, ""
        End If
    Else
        Form_JEJARING_SOSIAL.textJejaringSosial(6).Text = textTanggal.Text
        Unload Me
        With Form_JEJARING_SOSIAL.textJejaringSosial(6)
            .ForeColor = Hitam
            .SetFocus
        End With
    End If
ElseIf FORM_UTAMA.cmElectronicMail.FontBold = True Then
    If textTanggal.Text = "0 - 0 - 0" Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Tanggal Salah!", vbCritical + vbOKOnly, ""
        Else
            MsgBox "Wrong Time!", vbCritical + vbOKOnly, ""
        End If
    Else
        Form_ELECTRONIC_MAIL.textTanggal.Text = textTanggal.Text
        Unload Me
        With Form_ELECTRONIC_MAIL.textTanggal
            .ForeColor = Hitam
            .SetFocus
        End With
    End If
ElseIf FORM_UTAMA.cmForumInternet.FontBold = True Then
    If textTanggal.Text = "0 - 0 - 0" Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Tanggal Salah!", vbCritical + vbOKOnly, ""
        Else
            MsgBox "Wrong Time!", vbCritical + vbOKOnly, ""
        End If
    Else
        Form_FORUM_INTERNET.textTanggal.Text = textTanggal.Text
        Unload Me
        With Form_FORUM_INTERNET.textTanggal
            .ForeColor = Hitam
            .SetFocus
        End With
    End If
ElseIf FORM_UTAMA.cmFTP.FontBold = True Then
    If textTanggal.Text = "0 - 0 - 0" Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Tanggal Salah!", vbCritical + vbOKOnly, ""
        Else
            MsgBox "Wrong Time!", vbCritical + vbOKOnly, ""
        End If
    Else
        Form_FTP.textTanggal.Text = textTanggal.Text
        Unload Me
        With Form_FTP.textTanggal
            .ForeColor = Hitam
            .SetFocus
        End With
    End If
ElseIf FORM_UTAMA.cmBlogging.FontBold = True Then
    If textTanggal.Text = "0 - 0 - 0" Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Tanggal Salah!", vbCritical + vbOKOnly, ""
        Else
            MsgBox "Wrong Time!", vbCritical + vbOKOnly, ""
        End If
    Else
        Form_BLOGGING_WEBSITE.textTanggal.Text = textTanggal.Text
        Unload Me
        With Form_BLOGGING_WEBSITE.textTanggal
            .ForeColor = Hitam
            .SetFocus
        End With
    End If
ElseIf FORM_UTAMA.cmUlangTahun.FontBold = True Then
    If textTanggal.Text = "0 - 0 - 0" Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Tanggal Salah!", vbCritical + vbOKOnly, ""
        Else
            MsgBox "Wrong Time!", vbCritical + vbOKOnly, ""
        End If
    Else
        Form_ULANG_TAHUN.textTTL.Text = textTanggal.Text
        Unload Me
        With Form_ULANG_TAHUN.textTTL
            .ForeColor = Hitam
            .SetFocus
        End With
    End If
ElseIf FORM_UTAMA.cmAgenda.FontBold = True Then
    If textTanggal.Text = "0 - 0 - 0" Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Tanggal Salah!", vbCritical + vbOKOnly, ""
        Else
            MsgBox "Wrong Time!", vbCritical + vbOKOnly, ""
        End If
    Else
        Form_AGENDA.textTanggal.Text = textTanggal.Text
        Unload Me
        With Form_AGENDA.textTanggal
            .ForeColor = Hitam
            .SetFocus
        End With
    End If
ElseIf FORM_UTAMA.cmIdentitasPribadi.FontBold = True Then
    If textTanggal.Text = "0 - 0 - 0" Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Tanggal Salah!", vbCritical + vbOKOnly, ""
        Else
            MsgBox "Wrong Time!", vbCritical + vbOKOnly, ""
        End If
    Else
        With Form_IDENTITAS_PRIBADI
            .cmbTahun.Text = Kalender.Year
            Select Case Kalender.Month
            Case Is = 1, 2, 3, 4, 5, 6, 7, 8, 9
                .cmbBulan.Text = "0" & Kalender.Month
            Case Is = 10, 11, 12
                .cmbBulan.Text = Kalender.Month
            End Select
            .cmbTanggal.Text = Kalender.Day
        End With
            Unload Me
            With Form_IDENTITAS_PRIBADI.textTempat
                .ForeColor = Hitam
                .SetFocus
            End With
    End If
End If
End Sub


Private Sub Form_Load()
    AturKontrol
    PENGATURAN_WARNA
    PENGATURAN_BAHASA
End Sub

Private Sub Kalender_Click()
    textTanggal.Text = Kalender.Day & " - " & Kalender.Month & " - " & Kalender.Year
End Sub

Private Sub Kalender_NewMonth()
    textTanggal.Text = Kalender.Day & " - " & Kalender.Month & " - " & Kalender.Year
End Sub

Private Sub Kalender_NewYear()
    textTanggal.Text = Kalender.Day & " - " & Kalender.Month & " - " & Kalender.Year
End Sub

Sub PENGATURAN_WARNA()
    'PENGATURAN WARNA UNTUK FORM INI
    For Each Objek In Me
        Select Case FormPengaturan.cmbWarnaTampilan.ListIndex
        Case Is = 0 'Ungu Natural
            Me.BackColor = UnguNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguNatural
            If TypeName(Objek) = "Calendar" Then Objek.BackColor = UnguNatural
        Case Is = 1 'Merah
            Me.BackColor = Merah
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Merah
            If TypeName(Objek) = "Calendar" Then Objek.BackColor = Merah
        Case Is = 2 'Pink
            Me.BackColor = Pink
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Pink
            If TypeName(Objek) = "Calendar" Then Objek.BackColor = Pink
        Case Is = 3 'HijauMuda
            Me.BackColor = HijauMuda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = HijauMuda
            If TypeName(Objek) = "Calendar" Then Objek.BackColor = HijauMuda
        Case Is = 4 'Hitam
            Me.BackColor = Hitam
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Hitam
            If TypeName(Objek) = "Calendar" Then Objek.BackColor = Hitam
        Case Is = 5 'Silver
            Me.BackColor = Silver
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Silver
            If TypeName(Objek) = "Calendar" Then Objek.BackColor = Silver
        Case Is = 6 'SilverNatural
            Me.BackColor = SilverNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = SilverNatural
            If TypeName(Objek) = "Calendar" Then Objek.BackColor = SilverNatural
        Case Is = 7 'Orange
            Me.BackColor = Orange
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Orange
            If TypeName(Objek) = "Calendar" Then Objek.BackColor = UnguNatural
        Case Is = 8 'UnguJanda
            Me.BackColor = UnguJanda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguJanda
            If TypeName(Objek) = "Calendar" Then Objek.BackColor = UnguNatural
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
        cmBatal.Caption = "&Batal"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        cmBatal.Caption = "&Cancel"
    End If
End Sub
