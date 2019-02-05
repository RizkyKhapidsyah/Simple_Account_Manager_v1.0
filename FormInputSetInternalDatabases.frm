VERSION 5.00
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormInputSetInternalDatabases 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "======="
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormInputSetInternalDatabases.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPControls.XPText textInputSetInternalDatabases 
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   714
      Text            =   "XPText1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Dacara_dcButton.dcButton cmOK 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   5
      PicDown         =   "FormInputSetInternalDatabases.frx":27A2
      PicHot          =   "FormInputSetInternalDatabases.frx":2ABC
      PicNormal       =   "FormInputSetInternalDatabases.frx":2DD6
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmBatal 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Batal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FormInputSetInternalDatabases.frx":30F0
      PicHot          =   "FormInputSetInternalDatabases.frx":3542
      PicNormal       =   "FormInputSetInternalDatabases.frx":3994
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
      Caption         =   "---"
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "FormInputSetInternalDatabases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmBatal_Click()
If textInputSetInternalDatabases.Text = "" Then
    Unload Me
Else
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Pesan = MsgBox("Apakah Anda yakin ingin membatalkan?", vbQuestion + vbYesNo, "")
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Pesan = MsgBox("Are you sure to cancel?", vbQuestion + vbYesNo, "")
    End If
    If Pesan = vbYes Then Unload Me
End If
End Sub

Private Sub cmOK_Click()
If textInputSetInternalDatabases.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Maaf, input tidak boleh kosong!", vbExclamation + vbOKOnly, ""
    Else
        MsgBox "Sorry, input don't empty!", vbExclamation + vbOKOnly, ""
    End If
    textInputSetInternalDatabases.SetFocus
Else
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        If LabelPenanda.Caption = "Baru" Then
            Pesan = MsgBox("Apakah Anda yakin ingin menambahkan entry '" & textInputSetInternalDatabases.Text & "' ke dalam " & UCase(FormInternalDatabases.cmbNamaTabel.Text) & " ?", vbQuestion + vbYesNo, "Konfirmasi")
        ElseIf LabelPenanda.Caption = "Edit" Then
            If FormInternalDatabases.cmbNamaTabel.ListIndex = 0 Then
                Pesan = MsgBox("Apakah Anda yakin ingin merubah nama entry '" & FormInternalDatabases.AdodcUtama.Recordset.Fields(1).Value & "' menjadi '" & textInputSetInternalDatabases.Text & "' ?", vbQuestion + vbYesNo, "Konfirmasi")
            Else
                Pesan = MsgBox("Apakah Anda yakin ingin merubah nama entry '" & FormInternalDatabases.AdodcUtama.Recordset.Fields(0).Value & "' menjadi '" & textInputSetInternalDatabases.Text & "' ?", vbQuestion + vbYesNo, "Konfirmasi")
            End If
        End If
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        If LabelPenanda.Caption = "Baru" Then
            Pesan = MsgBox("Are you sure to add this entry '" & textInputSetInternalDatabases.Text & "' into " & UCase(FormInternalDatabases.cmbNamaTabel.Text) & " ?", vbQuestion + vbYesNo, "Confirmation")
        ElseIf LabelPenanda.Caption = "Edit" Then
            If FormInternalDatabases.cmbNamaTabel.ListIndex = 0 Then
                Pesan = MsgBox("Are you sure to add this entry '" & FormInternalDatabases.AdodcUtama.Recordset.Fields(1).Value & "' into '" & textInputSetInternalDatabases.Text & "' ?", vbQuestion + vbYesNo, "Confirmation")
            Else
                Pesan = MsgBox("Are you sure to replace this entry '" & FormInternalDatabases.AdodcUtama.Recordset.Fields(0).Value & "' into '" & textInputSetInternalDatabases.Text & "' ?", vbQuestion + vbYesNo, "Confirmation")
            End If
        End If
    End If
    If Pesan = vbYes Then
        Select Case LabelPenanda.Caption
        Case "Baru"
            With FormInternalDatabases
                If .cmbNamaTabel.ListIndex = 0 Then
                        .AdodcUtama.Recordset.AddNew
                        .AdodcUtama.Recordset.Fields(1).Value = textInputSetInternalDatabases.Text
                        .AdodcUtama.Recordset.Update
                        .AdodcUtama.Refresh
                Else
                        .AdodcUtama.Recordset.AddNew
                        .AdodcUtama.Recordset.Fields(0).Value = textInputSetInternalDatabases.Text
                        .AdodcUtama.Recordset.Update
                        .AdodcUtama.Refresh
                End If
                .AdodcUtama.Refresh
            End With
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                cmBatal.Caption = "Tutup"
            Else
                cmBatal.Caption = "Close"
            End If
            With textInputSetInternalDatabases
                .Text = ""
                .SetFocus
            End With
        Case "Edit"
            With FormInternalDatabases
                If .cmbNamaTabel.ListIndex = 0 Then
                        .AdodcUtama.Recordset.Delete
                        .AdodcUtama.Recordset.AddNew
                        .AdodcUtama.Recordset.Fields(1).Value = textInputSetInternalDatabases.Text
                        .AdodcUtama.Recordset.Update
                        .AdodcUtama.Refresh
                Else
                        .AdodcUtama.Recordset.Delete
                        .AdodcUtama.Recordset.AddNew
                        .AdodcUtama.Recordset.Fields(0).Value = textInputSetInternalDatabases.Text
                        .AdodcUtama.Recordset.Update
                        .AdodcUtama.Refresh
                End If
                .AdodcUtama.Refresh
            End With
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                cmBatal.Caption = "Tutup"
            Else
                cmBatal.Caption = "Close"
            End If
            With textInputSetInternalDatabases
                .Text = ""
                .SetFocus
            End With
        End Select
    End If
End If
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
        With Me
            .cmBatal.Caption = "&Batal"
            .Label1.Caption = "Mohon masukkan Agama Anda : "
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With Me
            .cmBatal.Caption = "&Cancel"
            .Label1.Caption = "Please input your religion : "
        End With
    End If
End Sub

