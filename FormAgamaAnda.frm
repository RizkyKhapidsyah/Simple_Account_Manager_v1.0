VERSION 5.00
Object = "{02353968-C1C9-4E0A-88D3-18759BDC60FE}#1.0#0"; "AeroSuite.ocx"
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormAgamaAnda 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agama Lain"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2520
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormAgamaAnda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Dacara_dcButton.dcButton cmOK 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "OK/Simpan"
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
      PicDown         =   "FormAgamaAnda.frx":0442
      PicHot          =   "FormAgamaAnda.frx":0794
      PicNormal       =   "FormAgamaAnda.frx":0AE6
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin AeroSuite.AeroTextBox textAgama 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BackColor       =   16777215
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Caption         =   "Tambahkan Agama Anda :"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2160
   End
End
Attribute VB_Name = "FormAgamaAnda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmOK_Click()
If textAgama.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi nama Agama Anda.", vbExclamation + vbOKOnly, ""
    Else
        MsgBox "Please put in your religi's name", vbExclamation + vbOKOnly, ""
    End If
    textAgama.SetFocus
Else
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Pesan = MsgBox("Apakah Anda yakin?", vbQuestion + vbYesNo, "")
    Else
        Pesan = MsgBox("Are you sure?", vbQuestion + vbYesNo, "")
    End If
    Select Case Pesan
    Case Is = vbYes
        With FormBuatAkunBaru
            .AdodcAgama.Recordset.AddNew
            .AdodcAgama.Recordset.Fields(1).Value = textAgama.Text
            .AdodcAgama.Recordset.Update
            .AdodcAgama.Refresh
            .cmbAgama.Clear
            Do Until .AdodcAgama.Recordset.EOF
                .cmbAgama.AddItem .AdodcAgama.Recordset.Fields(1).Value
                .AdodcAgama.Recordset.MoveNext
            Loop
            .cmbAgama.Text = textAgama.Text
        End With
        Unload Me
    End Select
End If
End Sub

Private Sub Form_Load()
With FormAgamaAnda
    If FormPengaturan.cmbBahasa.ListIndex = 1 Then
        .Caption = "Others Religi"
        Label1.Caption = "Add Your Religi : "
        cmOK.Caption = "&OK/Save"
    Else
        .Caption = "Agama Lain"
        Label1.Caption = "Tambahkan Agama Anda : "
        cmOK.Caption = "&OK/Simpan"
    End If
    .textAgama.Text = ""
End With
PENGATURAN_WARNA
PENGATURAN_BAHASA
    'PENGATURAN UNTUK ALWAYS ON TOP
    If FormPengaturan.cekAlwaysOnTop.Value = Checked Then
        SetOnTop (Me.hwnd)
    ElseIf FormPengaturan.cekAlwaysOnTop.Value = Unchecked Then
        NotOnTop (Me.hwnd)
    End If
    For Each Objek In Me
        If TypeName(Objek) = "Label" Or TypeName(Objek) = "dcButton" Then
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
        Me.Label1.Caption = "Tambahkan Agama Anda :"
        Me.cmOK.Caption = "OK/Simpan"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Me.Label1.Caption = "Add Your Religion :"
        Me.cmOK.Caption = "OK/Save"
    End If
End Sub
