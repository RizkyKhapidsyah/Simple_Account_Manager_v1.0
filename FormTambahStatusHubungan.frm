VERSION 5.00
Object = "{02353968-C1C9-4E0A-88D3-18759BDC60FE}#1.0#0"; "AeroSuite.ocx"
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormTambahStatusHubungan 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tambah Status Hubungan"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3705
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormTambahStatusHubungan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Dacara_dcButton.dcButton cmOK 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "OK/Simpan"
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
      PicDown         =   "FormTambahStatusHubungan.frx":0442
      PicHot          =   "FormTambahStatusHubungan.frx":0794
      PicNormal       =   "FormTambahStatusHubungan.frx":0AE6
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin AeroSuite.AeroTextBox TextStatusHubungan 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      BackColor       =   16777215
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "AeroTextBox1"
   End
   Begin XPEngine.XPControl XP_Engine 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tambah Status Hubungan : "
      BeginProperty Font 
         Name            =   "Nina"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2085
   End
End
Attribute VB_Name = "FormTambahStatusHubungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmOK_Click()
On Error GoTo BinasakanlahError
If TextStatusHubungan.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Maaf, input tidak boleh kosong!", vbExclamation + vbOKOnly, ""
    Else
        MsgBox "Sorry, value don't empty!", vbExclamation + vbOKOnly, ""
    End If
Else
    With Form_IDENTITAS_PRIBADI
        .AdodcStatusHubungan.Recordset.AddNew
        .AdodcStatusHubungan.Recordset.Fields(0).Value = TextStatusHubungan.Text
        .AdodcStatusHubungan.Recordset.Update
        .AdodcStatusHubungan.Refresh
            .cmbStatusHubungan.Clear
            Do Until .AdodcStatusHubungan.Recordset.EOF
                .cmbStatusHubungan.AddItem .AdodcStatusHubungan.Recordset.Fields(0).Value
                .AdodcStatusHubungan.Recordset.MoveNext
            Loop
            .cmbStatusHubungan.Text = TextStatusHubungan.Text
    End With
    Unload Me
End If
Exit Sub
BinasakanlahError:
    HancurkanError
    Form_IDENTITAS_PRIBADI.AdodcStatusHubungan.Refresh
End Sub

Private Sub Form_Load()
    NyambunggUtama
    PENGATURAN_WARNA
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
        Me.Caption = "Tambah Status Hubungan :"
        Label1.Caption = "Tambah Status Hubungan : "
        cmOK.Caption = "&OK/Simpan"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Me.Caption = "Add New Relationship :"
        Label1.Caption = "Add New Relationship : "
        cmOK.Caption = "&OK/Save"
    End If
End Sub
